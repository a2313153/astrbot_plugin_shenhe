from typing import Any, Dict, List, Optional
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
import time
import re
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
from astrbot.api.star import Star, register, Context
from astrbot.api.event.filter import PermissionType, filter
from astrbot.core.platform.message_type import MessageType
from astrbot.api import logger
from astrbot.core.platform.sources.aiocqhttp.aiocqhttp_message_event import (
    AiocqhttpMessageEvent,
)
from astrbot.core.platform.sources.aiocqhttp.aiocqhttp_request_event import (
    AiocqhttpRequestEvent,
)


@register(
    "astrbot_plugin_group_information",
    "Futureppo",
    "导出群成员信息及加群管理",
    "1.2.0",
    "https://github.com/Futureppo/astrbot_plugin_group_information",
)
class GroupInformationPlugin(Star):
    def __init__(self, context: Context):
        super().__init__(context)
        # 核心配置参数
        self.ADMIN_QQS = [1537008949, 1579648302]  # 管理员QQ列表
        self.API_TIMEOUT = 10  # API请求超时时间(秒)
        self.API_RETRIES = 3  # API请求重试次数
        self.API_BASE_URL = "https://qun.yz01.baby/api/"  # 基础API地址
        self.REQUEST_DELAY = 1  # 接口请求延迟(秒)，防止频率限制
        
        # 初始化带重试机制的请求会话
        self.api_session = self._create_retry_session()

    def _create_retry_session(self) -> requests.Session:
        """创建支持自动重试的请求会话"""
        session = requests.Session()
        retry_strategy = Retry(
            total=self.API_RETRIES,
            backoff_factor=0.5,
            status_forcelist=[429, 500, 502, 503, 504]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    @staticmethod
    def _format_timestamp(timestamp: Any) -> str:
        """格式化时间戳为可读字符串"""
        try:
            if isinstance(timestamp, (int, float)) and timestamp > 0:
                return datetime.fromtimestamp(float(timestamp)).strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
        except (TypeError, ValueError):
            pass
        return "0000-00-00 00:00:00"

    @staticmethod
    def _clean_excel_invalid_chars(text: Any) -> Any:
        """清理Excel不支持的特殊字符"""
        if isinstance(text, str):
            return "".join(char for char in text if ord(char) >= 32 and char not in "\x00\x01\x02\x03")
        return text

    # ------------------------------
    # 群成员数据导出功能
    # ------------------------------
    @filter.command("导出群数据")
    async def export_group_data(self, event: AiocqhttpMessageEvent):
        """导出指定群聊成员信息到Excel"""
        # 解析群号参数
        group_id = self._extract_group_id(event.get_plaintext(), event)
        if not group_id:
            yield event.plain_result("请在群聊中使用此命令或指定有效的群号（格式：导出群数据 123456）")
            return

        try:
            client = event.bot
            
            # 验证机器人是否在该群
            if not await self._is_bot_in_group(client, group_id):
                yield event.plain_result(f"机器人不在群 {group_id} 中，无法导出数据")
                return

            # 获取群成员列表
            members = await client.get_group_member_list(group_id=int(group_id), no_cache=True)
            if not isinstance(members, list):
                yield event.plain_result("获取群成员数据失败")
                return

            # 处理并生成Excel
            processed_members = self._process_members(members)
            file_content = self._generate_excel_file(processed_members, f"Group_{group_id}")
            file_name = f"群{group_id}_成员数据_{len(processed_members)}人.xlsx"
            
            # 上传文件
            result = await self._upload_file(
                event, 
                file_content, 
                file_name,
                is_group=event.message_obj.type == MessageType.GROUP_MESSAGE
            )
            
            if result:
                yield event.plain_result(f"群成员数据导出成功，共 {len(processed_members)} 人")
            else:
                yield event.plain_result("文件上传失败，请稍后重试")

        except Exception as e:
            logger.error(f"导出群数据错误: {str(e)}", exc_info=True)
            yield event.plain_result(f"操作失败: {str(e)}")

    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("导出所有群数据")
    async def export_all_groups_data(self, event: AiocqhttpMessageEvent):
        """导出机器人加入的所有群成员信息"""
        try:
            client = event.bot
            group_list = await client.get_group_list(no_cache=True)
            if not group_list:
                yield event.plain_result("机器人未加入任何群组")
                return

            yield event.plain_result(f"发现 {len(group_list)} 个群组，开始导出数据...")
            
            # 生成多sheet Excel
            output_buffer = BytesIO()
            total_members = 0
            failed_groups = []
            
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                for idx, group in enumerate(group_list, 1):
                    group_id = group["group_id"]
                    group_name = self._clean_excel_invalid_chars(group["group_name"])
                    
                    try:
                        # 实时反馈进度
                        if idx % 5 == 0:  # 每处理5个群反馈一次
                            yield event.plain_result(f"已处理 {idx}/{len(group_list)} 个群...")
                            
                        members = await client.get_group_member_list(group_id=group_id, no_cache=True)
                        if not isinstance(members, list):
                            failed_groups.append(f"{group_name}({group_id}): 成员数据无效")
                            continue
                            
                        processed_members = self._process_members(members)
                        for member in processed_members:
                            member["group_name"] = group_name
                            
                        df = pd.DataFrame(processed_members)
                        sheet_name = f"G{group_id}"[:30]  # 限制sheet名长度
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                        total_members += len(processed_members)
                        
                    except Exception as e:
                        failed_groups.append(f"{group_name}({group_id}): {str(e)[:30]}")
                        logger.warning(f"处理群 {group_id} 失败: {str(e)}")

            # 上传结果文件
            output_buffer.seek(0)
            file_content = output_buffer.getvalue()
            file_name = f"所有群成员数据_{len(group_list)}群_{total_members}人.xlsx"
            
            upload_success = await self._upload_file(
                event, 
                file_content, 
                file_name,
                is_group=event.message_obj.type == MessageType.GROUP_MESSAGE
            )
            
            # 生成报告
            report = f"全部导出完成！共 {total_members} 名成员\n"
            if failed_groups:
                report += f"⚠️ 有 {len(failed_groups)} 个群处理失败:\n" + "\n".join(failed_groups[:5])
                if len(failed_groups) > 5:
                    report += f"\n...还有 {len(failed_groups)-5} 个失败项"
            
            yield event.plain_result(report)

        except Exception as e:
            logger.error(f"导出所有群数据错误: {str(e)}", exc_info=True)
            yield event.plain_result(f"操作失败: {str(e)}")

    # ------------------------------
    # 加群请求处理功能
    # ------------------------------
    @filter.event(AiocqhttpRequestEvent)
    @filter.func(lambda e: e.event_data.get('request_type') == 'group' and e.event_data.get('sub_type') == 'add')
    async def handle_join_group_request(self, event: AiocqhttpRequestEvent):
        """处理加群请求并验证卡密"""
        data = event.event_data
        group_id = str(data['group_id'])
        user_qq = str(data['user_id'])
        comment = data.get('comment', '')
        flag = data['flag']
        
        # 提取卡密（12位字母数字组合）
        key = self._extract_activation_key(comment)
        logger.info(f"加群请求 - 群{group_id} 用户{user_qq} 卡密:{key} 备注:{comment}")

        # 卡密验证
        try:
            # 调用API验证卡密
            verify_url = f"{self.API_BASE_URL}check_key.php?group_id={group_id}&key={key}"
            response = self.api_session.get(verify_url, timeout=self.API_TIMEOUT)
            response.raise_for_status()
            result = response.json()
            
            if result.get('status') == 'success' and result.get('usable') == 1:
                # 验证通过，同意入群
                await event.bot.set_group_add_request(flag=flag, sub_type='add', approve=True)
                logger.info(f"卡密验证通过 - 群{group_id} 用户{user_qq}")
                
                # 标记卡密已使用
                await self._mark_key_used(group_id, key, user_qq)
                
            else:
                # 验证失败，拒绝入群
                reason = result.get('message', '卡密无效')
                await event.bot.set_group_add_request(
                    flag=flag, 
                    sub_type='add', 
                    approve=False,
                    reason=reason if len(reason) <= 30 else "卡密验证失败"
                )
                logger.warning(f"卡密验证失败 - 群{group_id} 用户{user_qq} 原因:{reason}")
                
        except requests.exceptions.RequestException as e:
            # API请求异常
            logger.error(f"卡密验证API错误: {str(e)}")
            await event.bot.set_group_add_request(
                flag=flag, 
                sub_type='add', 
                approve=False,
                reason="验证系统维护中，请稍后再试"
            )
        except Exception as e:
            logger.error(f"处理加群请求错误: {str(e)}", exc_info=True)

    # ------------------------------
    # 群成员推送功能
    # ------------------------------
    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("获取群成员", aliases=["获取群员QQ"])
    async def get_group_members(self, event: AiocqhttpMessageEvent):
        """获取单个群成员并推送至API"""
        if not self._is_admin(event.get_sender_id()):
            yield event.plain_result("权限不足：仅管理员可执行此操作")
            return

        # 提取群号
        cmd_text = event.get_plaintext().strip()
        match = re.search(r'(\d+)', cmd_text)
        if not match:
            yield event.plain_result("请指定群号，格式：获取群成员 123456789")
            return
        
        group_id = match.group(1)
        yield event.plain_result(f"开始获取群 {group_id} 的成员信息...")
        
        # 获取并推送成员数据
        members, error = await self.fetch_group_members(event.bot, group_id)
        if error:
            yield event.plain_result(f"获取失败：{error}")
            return

        push_result = await self._push_members_to_api(members)
        yield event.plain_result(push_result)

    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("获取所有群成员", aliases=["全量更新群成员"])
    async def get_all_group_members(self, event: AiocqhttpMessageEvent):
        """获取所有群成员并批量推送至API"""
        if not self._is_admin(event.get_sender_id()):
            yield event.plain_result("权限不足：仅管理员可执行此操作")
            return

        try:
            client = event.bot
            group_list = await client.get_group_list(no_cache=True)
            if not group_list:
                yield event.plain_result("机器人未加入任何群组")
                return

            total_groups = len(group_list)
            success_count = 0
            failed_groups = []
            
            yield event.plain_result(f"发现 {total_groups} 个群，开始批量处理...")

            for i, group in enumerate(group_list, 1):
                group_id = str(group['group_id'])
                group_name = group.get('group_name', f"群{group_id}")
                
                # 进度反馈
                if i % 3 == 0 or i == total_groups:
                    yield event.plain_result(f"处理进度：{i}/{total_groups}（{group_name}）")
                
                # 获取成员
                members, error = await self.fetch_group_members(client, group_id)
                if error:
                    failed_groups.append(f"{group_name}：{error}")
                    continue

                # 推送数据
                push_result = await self._push_members_to_api(members)
                if "成功" in push_result:
                    success_count += 1
                else:
                    failed_groups.append(f"{group_name}：{push_result}")
                
                time.sleep(self.REQUEST_DELAY)  # 避免请求过于频繁

            # 生成报告
            report = f"批量处理完成！\n成功：{success_count} 个群\n失败：{len(failed_groups)} 个群"
            if failed_groups:
                report += "\n失败详情：\n" + "\n".join(failed_groups[:5])
                if len(failed_groups) > 5:
                    report += f"\n...及其他 {len(failed_groups)-5} 项"
            
            yield event.plain_result(report)
            
        except Exception as e:
            logger.error(f"全量更新成员错误: {str(e)}", exc_info=True)
            yield event.plain_result(f"操作失败: {str(e)}")

    # ------------------------------
    # 辅助工具函数
    # ------------------------------
    def _process_members(self, members: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """处理成员数据，格式化字段"""
        processed = []
        char_clean_fields = {"nickname", "card", "title"}
        timestamp_fields = {"join_time", "last_sent_time", "title_expire_time", "shut_up_timestamp"}

        for member in members:
            if not isinstance(member, dict):
                continue
                
            item = {}
            for key, value in member.items():
                if key in char_clean_fields:
                    item[key] = self._clean_excel_invalid_chars(value)
                elif key in timestamp_fields:
                    item[key] = self._format_timestamp(value)
                else:
                    item[key] = value
                    
            processed.append(item)
            
        return processed

    @staticmethod
    def _generate_excel_file(data: List[Dict[str, Any]], sheet_name: str = "Sheet1") -> bytes:
        """生成Excel文件二进制内容"""
        df = pd.DataFrame(data)
        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl", sheet_name=sheet_name)
        buffer.seek(0)
        return buffer.getvalue()

    async def _upload_file(
        self, 
        event: AiocqhttpMessageEvent, 
        file_content: bytes, 
        file_name: str, 
        is_group: bool = True
    ) -> bool:
        """上传文件到群聊或私聊"""
        try:
            target_id = event.get_group_id() if is_group else event.get_sender_id()
            if not target_id:
                return False
                
            # 转换为base64格式上传
            b64_content = base64.b64encode(file_content).decode("utf-8")
            
            if is_group:
                await event.bot.upload_group_file(
                    group_id=int(target_id),
                    file=f"base64://{b64_content}",
                    name=file_name
                )
            else:
                await event.bot.upload_private_file(
                    user_id=int(target_id),
                    file=f"base64://{b64_content}",
                    name=file_name
                )
            return True
            
        except Exception as e:
            logger.error(f"文件上传失败: {str(e)}")
            return False

    async def fetch_group_members(self, bot, group_id: str) -> tuple[List[Dict], Optional[str]]:
        """分页获取群成员列表"""
        try:
            all_members = []
            next_token = None

            while True:
                params = {"group_id": int(group_id)}
                if next_token:
                    params["next_token"] = next_token

                result = await bot.get_group_member_list(**params)
                
                # 处理分页数据
                if isinstance(result, dict):
                    all_members.extend(result.get("data", []))
                    next_token = result.get("next_token")
                elif isinstance(result, list):
                    all_members.extend(result)
                    next_token = None
                else:
                    break

                if not next_token:
                    break
                time.sleep(0.5)  # 分页请求间隔

            # 格式化成员数据
            formatted = [
                {
                    "group_id": group_id,
                    "user_id": str(m["user_id"]),
                    "nickname": self._clean_excel_invalid_chars(m.get("nickname", "")),
                    "card": self._clean_excel_invalid_chars(m.get("card", ""))
                } 
                for m in all_members
            ]
            
            return formatted, None
            
        except Exception as e:
            return [], str(e)

    async def _push_members_to_api(self, members: List[Dict]) -> str:
        """推送成员数据到API"""
        if not members:
            return "没有可推送的成员数据"
            
        try:
            url = f"{self.API_BASE_URL}push_group_members.php"
            payload = {
                "bot_qq": str(self.context.bot.self_id),
                "members": members
            }
            
            response = self.api_session.post(
                url,
                json=payload,
                timeout=self.API_TIMEOUT
            )
            response.raise_for_status()
            result = response.json()
            
            if result.get("status") == "success":
                group_id = members[0]["group_id"]
                return f"群 {group_id} 成功推送 {len(members)} 名成员"
            else:
                return f"推送失败: {result.get('message', '未知错误')}"
                
        except requests.exceptions.RequestException as e:
            return f"API请求失败: {str(e)}"
        except Exception as e:
            return f"处理失败: {str(e)}"

    def _extract_group_id(self, text: str, event: AiocqhttpMessageEvent) -> Optional[str]:
        """从命令文本中提取群号，优先使用参数，其次使用当前群"""
        match = re.search(r'(\d+)', text)
        if match:
            return match.group(1)
            
        # 从事件中获取当前群号
        if event.message_obj.type == MessageType.GROUP_MESSAGE:
            return str(event.get_group_id())
            
        return None

    @staticmethod
    def _extract_activation_key(comment: str) -> str:
        """从备注中提取12位卡密"""
        match = re.search(r'[A-Za-z0-9]{12}', comment)
        return match.group(0) if match else ""

    def _is_admin(self, user_id: str) -> bool:
        """检查用户是否为管理员"""
        return int(user_id) in self.ADMIN_QQS

    @staticmethod
    async def _is_bot_in_group(bot, group_id: str) -> bool:
        """检查机器人是否在指定群聊中"""
        try:
            groups = await bot.get_group_list(no_cache=True)
            return any(str(g["group_id"]) == group_id for g in groups)
        except Exception:
            return False

    async def _mark_key_used(self, group_id: str, key: str, user_id: str) -> None:
        """标记卡密已使用"""
        try:
            url = f"{self.API_BASE_URL}mark_key.php?group_id={group_id}&key={key}&used_by={user_id}"
            await self.context.loop.run_in_executor(
                None, 
                lambda: self.api_session.get(url, timeout=self.API_TIMEOUT)
            )
        except Exception as e:
            logger.error(f"标记卡密使用状态失败: {str(e)}")
