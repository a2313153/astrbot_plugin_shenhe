from astrbot.api.all import *
from astrbot.api.event.filter import command, permission_type, event_message_type, EventMessageType, PermissionType
from astrbot.api.star import StarTools
from astrbot.api import logger
import json
import os

# 自定义配置
ADMIN_QQS = [1537008949, 1579648302]  # 替换为实际管理员QQ
API_TIMEOUT = 10  # API请求超时时间（秒）
API_RETRIES = 3  # API请求重试次数
API_BASE_URL = "https://qun.yz01.baby/api/"  # API基础URL

# 创建支持重试的Session（复用请求连接，提高稳定性）
api_session = requests.Session()
retries = Retry(total=API_RETRIES, backoff_factor=1)
api_session.mount('https://', HTTPAdapter(max_retries=retries))

# 加群请求处理
joingroup = on_request()

@joingroup.handle()
async def handle_join_request(bot: Bot, event: GroupRequestEvent):
    # 只处理加群请求
    if event.request_type != "group" or event.sub_type != "add":
        return
    
    # 安全获取参数
    group_id = str(event.group_id) if event.group_id else ""
    user_id = str(event.user_id) if event.user_id else ""
    comment = event.comment or ""
    
    if not group_id or not user_id:
        print("群号或用户ID为空，跳过处理")
        return
    
    # 提取卡密
    def extract_key(text):
        if not text:
            return ""
        match = re.search(r'[A-Za-z0-9]{12}', text)
        return match.group(0) if match else ''
    
    key = extract_key(comment)
    print(f"加群请求 - 群: {group_id}, 用户: {user_id}, 卡密: {key}")
    
    # 验证卡密
    try:
        api_url = f"https://qun.yz01.baby/api/check_key.php?group_id={group_id}&key={key}"
        response = requests.get(api_url, timeout=API_TIMEOUT)
        result = response.json()
        
        if result.get('status') == 'success' and result.get('usable') == 1:
            # 卡密有效
            print(f"卡密验证通过 - 群: {group_id}, 用户: {user_id}")
            
            await bot.set_group_add_request(
                flag=event.flag,
                sub_type="add",
                approve=True
            )
            
            # 标记卡密为已使用
            mark_url = f"https://qun.yz01.baby/api/mark_key.php?group_id={group_id}&key={key}&used_by={user_id}"
            requests.get(mark_url, timeout=API_TIMEOUT)
            
        else:
            # 卡密无效
            error_msg = result.get('message', '卡密错误')
            print(f"卡密验证失败 - 原因: {error_msg}")
            
            reason = '卡密已使用' if error_msg == '卡密已使用' else '卡密错误'
            await bot.set_group_add_request(
                flag=event.flag,
                sub_type="add",
                approve=False,
                reason=reason
            )
            
    except Exception as e:
        print(f"验证卡密时出错: {e}")
        await bot.set_group_add_request(
            flag=event.flag,
            sub_type="add",
            approve=False,
            reason='系统错误，请稍后再试'
        )

# 获取群成员命令
get_group_members = on_command("获取群成员", aliases={"获取群员QQ"})

@get_group_members.handle()
async def handle_get_members(bot: Bot, event: Event):
    # 权限检查
    user_id = event.get_user_id()
    if int(user_id) not in ADMIN_QQS and not await SUPERUSER(bot, event):
        await get_group_members.finish("权限不足")
    
    # 提取群号
    args = event.get_plaintext().strip().split()
    if len(args) < 2:
        await get_group_members.finish("用法: 获取群成员 <群号>")
    
    group_id = args[1]
    if not group_id.isdigit():
        await get_group_members.finish("群号必须为数字")
    
    await get_group_members.send(f"开始获取群 {group_id} 成员信息...")
    
    # 获取群成员
    try:
        members = await bot.get_group_member_list(group_id=int(group_id))
        
        # 格式化成员数据
        formatted_members = []
        for member in members:
            formatted_members.append({
                "group_id": group_id,
                "user_id": member.get("user_id", ""),
                "nickname": member.get("nickname", "") or "",
                "card": member.get("card", "") or ""
            })
        
        # 推送数据
        data = {
            "bot_qq": bot.self_id,
            "members": formatted_members
        }
        
        response = api_session.post(
            f"{API_BASE_URL}push_group_members.php",
            json=data,
            timeout=API_TIMEOUT
        )
        response.raise_for_status()
        
        result = response.json()
        if result.get("status") == "success":
            await get_group_members.finish(f"成功记录 {len(formatted_members)} 名成员")
        else:
            await get_group_members.finish(f"记录失败: {result.get('message', '未知错误')}")
            
    except Exception as e:
        await get_group_members.finish(f"获取成员失败: {str(e)}")

# 获取所有群成员命令
get_all_group_members = on_command("获取所有群成员", aliases={"全量更新群成员"})

@get_all_group_members.handle()
async def handle_get_all_members(bot: Bot, event: Event):
    # 权限检查
    user_id = event.get_user_id()
    if int(user_id) not in ADMIN_QQS and not await SUPERUSER(bot, event):
        await get_all_group_members.finish("权限不足")
    
    try:
        # 获取群列表
        groups = await bot.get_group_list()
        if not groups:
            await get_all_group_members.finish("未加入任何群组")
        
        total = len(groups)
        success = 0
        failed = []
        
        await get_all_group_members.send(f"发现 {total} 个群，开始处理...")
        
        # 处理每个群
        for i, group in enumerate(groups, 1):
            group_id = str(group.get("group_id", ""))
            group_name = group.get("group_name", f"群{group_id}")
            
            if not group_id:
                failed.append(f"未知群: 缺少群ID")
                continue
            
            await get_all_group_members.send(f"处理中 ({i}/{total}): {group_name}")
            
            try:
                # 获取成员
                members = await bot.get_group_member_list(group_id=int(group_id))
                
                # 格式化成员数据
                formatted_members = []
                for member in members:
                    formatted_members.append({
                        "group_id": group_id,
                        "user_id": member.get("user_id", ""),
                        "nickname": member.get("nickname", "") or "",
                        "card": member.get("card", "") or ""
                    })
                
                # 推送数据
                data = {
                    "bot_qq": bot.self_id,
                    "members": formatted_members
                }
                
                response = api_session.post(
                    f"{API_BASE_URL}push_group_members.php",
                    json=data,
                    timeout=API_TIMEOUT
                )
                response.raise_for_status()
                
                success += 1
                
            except Exception as e:
                failed.append(f"{group_name}: {str(e)}")
            
            # 避免频率限制
            time.sleep(1)
        
        # 生成报告
        report = f"处理完成! 成功: {success}, 失败: {len(failed)}"
        if failed:
            report += "\n失败详情:\n" + "\n".join(failed[:3])
            if len(failed) > 3:
                report += f"\n...还有 {len(failed)-3} 个失败项"
        
        await get_all_group_members.finish(report)
        
    except Exception as e:
        await get_all_group_members.finish(f"执行失败: {str(e)}")
