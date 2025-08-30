from astrbot import Plugin, on_request, on_command
from astrbot.adapters.onebot.v11 import Bot, Event, GroupRequestEvent
from astrbot.permission import SUPERUSER
import json
import time
import requests
import re
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

# 初始化插件
plugin = Plugin("group_manager", "群管理插件（带卡密验证功能）")

# 自定义配置
ADMIN_QQS = [1537008949, 1579648302]  # 管理员QQ
API_TIMEOUT = 10  # API请求超时时间（秒）
API_RETRIES = 3  # API请求重试次数
API_BASE_URL = "https://qun.yz01.baby/api/"  # API基础URL

# 创建支持重试的Session
api_session = requests.Session()
retries = Retry(total=API_RETRIES, backoff_factor=1)
api_session.mount('https://', HTTPAdapter(max_retries=retries))

# 加群请求处理
@on_request(plugin=plugin)
async def handle_join_group(bot: Bot, event: GroupRequestEvent):
    # 只处理加群请求
    if event.sub_type != "add":
        return
    
    group_id = str(event.group_id)
    user_qq = str(event.user_id)
    comment = event.comment or ""
    
    # 从备注中提取卡密
    def extract_key(comment):
        match = re.search(r'[A-Za-z0-9]{12}', comment)
        return match.group(0) if match else ''
    
    key = extract_key(comment)
    plugin.logger.info(f"收到加群请求 - 群号: {group_id}, 用户: {user_qq}, 备注: {comment}, 提取卡密: {key}")
    
    # 调用API验证卡密
    api_url = f"https://qun.yz01.baby/api/check_key.php?group_id={group_id}&key={key}"
    try:
        response = requests.get(api_url)
        result = response.json()
        plugin.logger.info(f"API响应: {result}")
        
        if result.get('status') == 'success' and result.get('usable') == 1:
            # 卡密有效，通过申请
            plugin.logger.info(f"卡密验证通过 - 群号: {group_id}, 用户: {user_qq}, 卡密: {key}")
            
            await bot.set_group_add_request(
                flag=event.flag,
                sub_type='add',
                approve=True
            )
            
            # 标记卡密为已使用
            mark_url = f"https://qun.yz01.baby/api/mark_key.php?group_id={group_id}&key={key}&used_by={user_qq}"
            mark_response = requests.get(mark_url)
            plugin.logger.info(f"标记卡密API响应: {mark_response.text}")
            
        else:
            # 卡密无效，拒绝加群
            error_msg = result.get('message', '卡密错误')
            plugin.logger.warning(f"卡密验证失败 - 群号: {group_id}, 用户: {user_qq}, 卡密: {key}, 原因: {error_msg}")
            
            reason = '卡密已使用' if error_msg == '卡密已使用' else '卡密错误'
                
            await bot.set_group_add_request(
                flag=event.flag,
                sub_type='add',
                approve=False,
                reason=reason
            )
            
    except Exception as e:
        plugin.logger.error(f"验证卡密API请求异常: {e}")
        # 默认拒绝加群
        await bot.set_group_add_request(
            flag=event.flag,
            sub_type='add',
            approve=False,
            reason='卡密验证系统暂时不可用，请稍后再试'
        )

# 获取单个群成员
@on_command("获取群成员", aliases={"获取群员QQ"}, plugin=plugin)
async def get_group_members(bot: Bot, event: Event):
    # 权限检查
    user_id = event.get_user_id()
    if int(user_id) not in ADMIN_QQS and not await SUPERUSER(bot, event):
        await bot.send(event, "你没有权限执行此操作")
        return
    
    # 提取群号
    cmd_text = event.get_plaintext().strip()
    match = re.search(r'(\d+)', cmd_text)
    if not match:
        await bot.send(event, "请指定群号，格式：获取群成员 123456789")
        return
    
    group_id = match.group(1)
    await bot.send(event, f"开始获取群 {group_id} 的成员信息...")
    
    # 获取成员
    members, error = await fetch_group_members(bot, group_id)
    if error:
        await bot.send(event, f"获取失败：{error}")
        return
    
    # 推送成员数据到数据库
    try:
        data = {
            "bot_qq": bot.self_id,
            "members": members
        }
        response = api_session.post(
            f"{API_BASE_URL}push_group_members.php",
            json=data,
            timeout=API_TIMEOUT
        )
        response.raise_for_status()
        result = response.json()
        
        if result.get("status") == "success":
            await bot.send(event, f"成功记录群 {group_id} 的 {len(members)} 名成员")
        else:
            await bot.send(event, f"记录失败：{result.get('message', '未知错误')}")
    except Exception as e:
        await bot.send(event, f"推送数据失败：{str(e)}")

# 获取全部群成员
@on_command("获取所有群成员", aliases={"全量更新群成员"}, plugin=plugin)
async def get_all_group_members(bot: Bot, event: Event):
    # 权限检查
    user_id = event.get_user_id()
    if int(user_id) not in ADMIN_QQS and not await SUPERUSER(bot, event):
        await bot.send(event, "你没有权限执行此操作")
        return
    
    try:
        # 获取机器人已加入的所有群
        group_list = await bot.get_group_list()
        if not group_list:
            await bot.send(event, "机器人未加入任何群")
            return
        
        total_groups = len(group_list)
        success_count = 0
        failed_groups = []
        await bot.send(event, f"发现 {total_groups} 个群，开始批量获取成员信息...")
        
        # 逐个处理群
        for i, group in enumerate(group_list, 1):
            group_id = group['group_id']
            group_name = group.get('group_name', f"群{group_id}")
            
            # 发送进度提示
            await bot.send(event, f"正在处理 {group_name}（{i}/{total_groups}）")
            
            # 获取成员
            members, error = await fetch_group_members(bot, group_id)
            if error:
                failed_groups.append(f"{group_name}：{error}")
                continue
            
            # 推送数据
            try:
                response = api_session.post(
                    f"{API_BASE_URL}push_group_members.php",
                    json={"bot_qq": bot.self_id, "members": members},
                    timeout=API_TIMEOUT
                )
                response.raise_for_status()
                success_count += 1
            except Exception as e:
                failed_groups.append(f"{group_name}：推送失败 - {str(e)}")
            
            # 避免高频请求触发风控
            time.sleep(1)
        
        # 生成结果报告
        report = f"批量处理完成！\n成功：{success_count} 个群\n失败：{len(failed_groups)} 个群"
        if failed_groups:
            report += "\n失败详情：\n" + "\n".join(failed_groups)
        await bot.send(event, report)
        
    except Exception as e:
        await bot.send(event, f"执行失败：{str(e)}")

# 辅助函数：分页获取单个群成员
async def fetch_group_members(bot: Bot, group_id: str):
    try:
        all_members = []
        next_token = None
        
        # 分页获取成员（处理大群）
        while True:
            params = {"group_id": group_id}
            if next_token:
                params["next_token"] = next_token
            
            # 调用API获取成员列表
            members = await bot.get_group_member_list(** params)
            
            # 处理不同适配器的返回格式
            if isinstance(members, list):
                all_members.extend(members)
                next_token = None  # 列表格式无分页
            elif isinstance(members, dict):
                if 'data' in members:
                    all_members.extend(members['data'])
                next_token = members.get('next_token')  # 字典格式带分页标记
            else:
                break
            
            # 没有下一页则退出循环
            if not next_token:
                break
            
            # 分页间隔，避免触发风控
            time.sleep(0.5)
        
        # 整理成员数据（只保留需要的字段）
        return [
            {
                "group_id": group_id,
                "user_id": m["user_id"],
                "nickname": m.get("nickname", ""),
                "card": m.get("card", "")
            } 
            for m in all_members
        ], None
    
    except Exception as e:
        return [], str(e)
