from typing import Any, Dict, List
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
import json
import time
import requests
import re
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
from astrbot.api.star import Star, register, Context
from astrbot.api.event.filter import PermissionType
from astrbot.core.platform.message_type import MessageType
from astrbot.api import logger
from astrbot.api.event import filter
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
    "1.1.0",
    "https://github.com/Futureppo/astrbot_plugin_group_information",
)
class GroupInformationPlugin(Star):
    def __init__(self, context: Context):
        super().__init__(context)
        # 自定义配置
        self.ADMIN_QQS = [1537008949, 1579648302]  # 管理员QQ
        self.API_TIMEOUT = 10  # API请求超时时间（秒）
        self.API_RETRIES = 3  # API请求重试次数
        self.API_BASE_URL = "https://qun.yz01.baby/api/"  # API基础URL
        
        # 创建支持重试的Session
        self.api_session = requests.Session()
        retries = Retry(total=self.API_RETRIES, backoff_factor=1)
        self.api_session.mount('https://', HTTPAdapter(max_retries=retries))

    @staticmethod
    def _format_timestamp(timestamp):
        """格式化时间戳为可读时间"""
        if isinstance(timestamp, (int, float)) and timestamp >= 0:
            return datetime.fromtimestamp(float(timestamp)).strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        else:
            return "0000-00-00 00:00:00"

    @staticmethod
    def _clean_excel_invalid_chars(text):
        """清理Excel不支持的特殊字符"""
        if not isinstance(text, str):
            return text
        return "".join(
            char for char in text if ord(char) >= 32 and char not in "\x00\x01\x02\x03"
        )

    # ---------- 1.py原有导出功能 ----------
    @filter.command("导出群数据")
    async def export_group_data(self, event: AiocqhttpMessageEvent, group_id: str = ''):
        """导出指定群聊成员信息到Excel文件"""
        if group_id:
            group_id = group_id.strip()
            if not group_id.isdigit():
                yield event.plain_result("请输入有效的群号")
                return
        else:
            group_id = event.get_group_id()
            if not group_id:
                yield event.plain_result("请在群聊中使用此命令或提供有效的群号")
                return
        try:
            client = event.bot

            try:
                await client.get_group_member_info(
                    group_id=int(group_id), user_id=int(event.get_sender_id()), no_cache=True
                )
            except Exception as e:
                logger.error(f"获取群成员信息时出错: {e}")
                yield event.plain_result("你不在该群聊中，无法导出数据")
                return

            members: list[dict] = await client.get_group_member_list(
                group_id=int(group_id), no_cache=True
            )
            processed_members = self._process_members(members)
            file_content = self._generate_excel_file(
                processed_members, sheet_name=f"Group_{group_id}"
            )
            file_name = f"群聊{group_id}的{len(processed_members)}名成员的数据.xlsx"
            
            message_type = 'group' if event.message_obj.type == MessageType.GROUP_MESSAGE else 'private'
            group_or_user_id = event.get_group_id() if message_type == 'group' else event.get_sender_id()
            await self._upload_file(
                event, file_content, group_or_user_id, file_name, message_type=message_type
            )

        except Exception as e:
            logger.error(f"导出群数据时出错: {e}")
            yield event.plain_result(f"导出群数据时出错: {str(e)}")

    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("导出所有群数据")
    async def export_all_groups_data(self, event: AiocqhttpMessageEvent):
        """导出所有群的成员信息到多个sheet的Excel文件中"""
        client = event.bot
        group_list = await client.get_group_list(no_cache=True)
        yield event.plain_result(f"正在导出{len(group_list)}个群的数据...")
        try:
            output_buffer = BytesIO()
            total_members = 0
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                for group in group_list:
                    group_id = group["group_id"]
                    group_name = group["group_name"]
                    try:
                        members: list[dict] = await client.get_group_member_list(
                            group_id=group_id, no_cache=True
                        )
                        processed_members = self._process_members(members)
                        for member in processed_members:
                            member["group_name"] = self._clean_excel_invalid_chars(group_name)
                        
                        df = pd.DataFrame(processed_members)
                        sheet_name = f"G{group_id}"[:30]
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                        total_members += len(processed_members)
                        logger.info(f"已导出{group_name}({group_id})的{len(processed_members)}名成员信息")
                    except Exception as e:
                        logger.error(f"处理群 {group_id} 时出错: {e}")
                        continue
            
            output_buffer.seek(0)
            file_content = output_buffer.getvalue()
            file_name = f"{len(group_list)}个群的{total_members}名成员的数据.xlsx"
            
            message_type = 'group' if event.message_obj.type == MessageType.GROUP_MESSAGE else 'private'
            group_or_user_id = event.get_group_id() if message_type == 'group' else event.get_sender_id()
            await self._upload_file(
                event, file_content, group_or_user_id, file_name, message_type=message_type
            )

        except Exception as e:
            logger.error(f"导出所有群数据时出错: {e}")
            yield event.plain_result(f"导出所有群数据时出错: {str(e)}")

    # ---------- 2.py整合功能 ----------
    @filter.event(AiocqhttpRequestEvent)
    async def handle_join_group_request(self, event: AiocqhttpRequestEvent):
        """处理加群请求并验证卡密"""
        data = event.event_data
        # 过滤非加群请求
        if data.get('request_type') != 'group' or data.get('sub_type') != 'add':
            return

        group_id = str(data['group_id'])
        user_qq = str(data['user_id'])
        comment = data.get('comment', '')
        
        # 提取卡密
        def extract_key(comment):
            match = re.search(r'[A-Za-z0-9]{12}', comment)
            return match.group(0) if match else ''
        
        key = extract_key(comment)
        logger.info(f"收到加群请求 - 群号: {group_id}, 用户: {user_qq}, 备注: {comment}, 提取卡密: {key}")

        # 验证卡密
        api_url = f"https://qun.yz01.baby/api/check_key.php?group_id={group_id}&key={key}"
        try:
            response = self.api_session.get(api_url, timeout=self.API_TIMEOUT)
            result = response.json()
            logger.info(f"API响应: {result}")

            if result.get('status') == 'success' and result.get('usable') == 1:
                # 同意加群并标记卡密
                logger.info(f"卡密验证通过 - 群号: {group_id}, 用户: {user_qq}, 卡密: {key}")
                await event.bot.set_group_add_request(
                    flag=data['flag'],
                    sub_type='add',
                    approve=True
                )
                
                # 标记卡密已使用
                mark_url = f"{self.API_BASE_URL}mark_key.php?group_id={group_id}&key={key}&used_by={user_qq}"
                mark_response = self.api_session.get(mark_url, timeout=self.API_TIMEOUT)
                logger.info(f"标记卡密API响应: {mark_response.text}")
            else:
                # 拒绝加群
                error_msg = result.get('message', '卡密错误')
                logger.warning(f"卡密验证失败 - 群号: {group_id}, 用户: {user_qq}, 原因: {error_msg}")
                reason = '卡密已使用' if error_msg == '卡密已使用' else '卡密错误'
                await event.bot.set_group_add_request(
                    flag=data['flag'],
                    sub_type='add',
                    approve=False,
                    reason=reason
                )
        except Exception as e:
            logger.error(f"验证卡密API请求异常: {e}")
            await event.bot.set_group_add_request(
                flag=data['flag'],
                sub_type='add',
                approve=False,
                reason='卡密验证系统暂时不可用，请稍后再试'
            )

    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("获取群成员", aliases=["获取群员QQ"])
    async def get_group_members(self, event: AiocqhttpMessageEvent):
        """获取单个群成员并推送至API"""
        user_id = event.get_sender_id()
        if int(user_id) not in self.ADMIN_QQS:
            yield event.plain_result("你没有权限执行此操作")
            return

        cmd_text = event.get_plaintext().strip()
        match = re.search(r'(\d+)', cmd_text)
        if not match:
            yield event.plain_result("请指定群号，格式：获取群成员 123456789")
            return

        group_id = match.group(1)
        yield event.plain_result(f"开始获取群 {group_id} 的成员信息...")

        members, error = await self.fetch_group_members(event.bot, group_id)
        if error:
            yield event.plain_result(f"获取失败：{error}")
            return

        # 推送成员数据
        try:
            response = self.api_session.post(
                f"{self.API_BASE_URL}push_group_members.php",
                json={"bot_qq": event.bot.self_id, "members": members},
                timeout=self.API_TIMEOUT
            )
            response.raise_for_status()
            result = response.json()
            
            if result.get("status") == "success":
                yield event.plain_result(f"成功记录群 {group_id} 的 {len(members)} 名成员")
            else:
                yield event.plain_result(f"记录失败：{result.get('message', '未知错误')}")
        except Exception as e:
            yield event.plain_result(f"推送数据失败：{str(e)}")

    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("获取所有群成员", aliases=["全量更新群成员"])
    async def get_all_group_members(self, event: AiocqhttpMessageEvent):
        """获取所有群成员并批量推送至API"""
        user_id = event.get_sender_id()
        if int(user_id) not in self.ADMIN_QQS:
            yield event.plain_result("你没有权限执行此操作")
            return

        try:
            group_list = await event.bot.get_group_list(no_cache=True)
            if not group_list:
                yield event.plain_result("机器人未加入任何群")
                return

            total_groups = len(group_list)
            success_count = 0
            failed_groups = []
            yield event.plain_result(f"发现 {total_groups} 个群，开始批量获取成员信息...")

            for i, group in enumerate(group_list, 1):
                group_id = str(group['group_id'])
                group_name = group.get('group_name', f"群{group_id}")
                
                yield event.plain_result(f"正在处理 {group_name}（{i}/{total_groups}）")
                
                members, error = await self.fetch_group_members(event.bot, group_id)
                if error:
                    failed_groups.append(f"{group_name}：{error}")
                    continue

                try:
                    response = self.api_session.post(
                        f"{self.API_BASE_URL}push_group_members.php",
                        json={"bot_qq": event.bot.self_id, "members": members},
                        timeout=self.API_TIMEOUT
                    )
                    response.raise_for_status()
                    success_count += 1
                except Exception as e:
                    failed_groups.append(f"{group_name}：推送失败 - {str(e)}")
                
                time.sleep(1)  # 避免高频请求

            report = f"批量处理完成！\n成功：{success_count} 个群\n失败：{len(failed_groups)} 个群"
            if failed_groups:
                report += "\n失败详情：\n" + "\n".join(failed_groups)
            yield event.plain_result(report)
            
        except Exception as e:
            yield event.plain_result(f"执行失败：{str(e)}")

    # ---------- 通用辅助函数 ----------
    def _process_members(self, members: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """处理成员数据，清理字段并格式化时间戳"""
        processed_members = []
        char_clean_fields = {"nickname", "card", "title"}
        timestamp_fields = {
            "join_time": None,
            "last_sent_time": None,
            "title_expire_time": 0,
            "shut_up_timestamp": 0,
        }

        for member in members:
            if not isinstance(member, dict):
                logger.warning("发现非字典类型的成员数据，已跳过")
                continue

            processed = {}
            for key, value in member.items():
                if key in char_clean_fields and isinstance(value, str):
                    processed[key] = self._clean_excel_invalid_chars(value)
                elif key in timestamp_fields:
                    processed[key] = value
                else:
                    processed[key] = value

            for key, default in timestamp_fields.items():
                raw_time = processed.get(key, default)
                processed[key] = self._format_timestamp(raw_time)

            processed_members.append(processed)
        return processed_members

    @staticmethod
    def _generate_excel_file(
        data: List[Dict[str, Any]], sheet_name: str = "Sheet1"
    ) -> bytes:
        """生成Excel文件"""
        df = pd.DataFrame(data)
        output_buffer = BytesIO()
        df.to_excel(
            output_buffer, index=False, engine="openpyxl", sheet_name=sheet_name
        )
        output_buffer.seek(0)
        return output_buffer.getvalue()

    @staticmethod
    async def _upload_file(
        event: AiocqhttpMessageEvent,
        file_content: bytes,
        user_or_group_id: str | int,
        file_name: str,
        message_type: str,
    ) -> bool:
        """上传文件到群组或私聊"""
        try:
            file_content_base64 = base64.b64encode(file_content).decode("utf-8")
            if message_type == 'group':
                await event.bot.upload_group_file(
                    group_id=int(user_or_group_id),
                    file=f"base64://{file_content_base64}",
                    name=file_name,
                )
            else:
                await event.bot.upload_private_file(
                    user_id=int(user_or_group_id),
                    file=f"base64://{file_content_base64}",
                    name=file_name,
                )
            logger.info(f"文件上传完成：{file_name}")
            return True
        except Exception as upload_e:
            logger.error(f"文件上传失败：{upload_e}")
            return False

    async def fetch_group_members(self, bot, group_id: str):
        """分页获取单个群成员"""
        try:
            all_members = []
            next_token = None

            while True:
                params = {"group_id": int(group_id)}
                if next_token:
                    params["next_token"] = next_token

                members = await bot.get_group_member_list(**params)

                if isinstance(members, list):
                    all_members.extend(members)
                    next_token = None
                elif isinstance(members, dict):
                    if 'data' in members:
                        all_members.extend(members['data'])
                    next_token = members.get('next_token')
                else:
                    break

                if not next_token:
                    break
                time.sleep(0.5)

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
