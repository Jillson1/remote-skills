import json
import requests
import urllib3
import sys
import struct

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import hashlib
import hmac
import base64
import urllib.parse
import re
from datetime import datetime
import os
import configparser
import mimetypes

class WpsClient:
    def __init__(self, config_path=None):
        self.config = configparser.ConfigParser()
        
        # 默认加载同级目录下 ../config/airpage.properties
        if not config_path:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(base_dir, 'config', 'airpage.properties')
        
        if os.path.exists(config_path):
            self.config.read(config_path, encoding='utf-8')
            self.access_key = self.config['DEFAULT'].get('ACCESS_KEY')
            self.secret_key = self.config['DEFAULT'].get('SECRET_KEY')
            self.domain = self.config['DEFAULT'].get('DOMAIN', 'https://openapi.wps.cn')
            self.drive_id = self.config['DEFAULT'].get('DEFAULT_DRIVE_ID')
            self.parent_id = self.config['DEFAULT'].get('DEFAULT_PARENT_ID', '0')
            self.mention_dept_id = self.config['DEFAULT'].get('MENTION_DEPT_ID')
        else:
            # Fallback to env vars or empty
            self.access_key = os.getenv('WPS_ACCESS_KEY')
            self.secret_key = os.getenv('WPS_SECRET_KEY')
            self.domain = os.getenv('WPS_DOMAIN', 'https://openapi.wps.cn')
            self.drive_id = os.getenv('WPS_DRIVE_ID')
            self.parent_id = '0'
            self.mention_dept_id = os.getenv('WPS_MENTION_DEPT_ID')

        self.access_token = None
        # 用户名到 userId 的缓存（避免重复查询部门成员）
        self._user_name_cache = None

    def resolve_file_id(self, input_str):
        r"""
        从输入字符串中解析 file_id。
        支持多种输入格式：
        - 完整 URL: https://kdocs.cn/l/cgXXO4wMXj5T
        - 带查询参数: https://kdocs.cn/l/cgXXO4wMXj5T?from=xxx
        - shell 转义的 URL: https://kdocs.cn/l/cgXXO4wMXj5T\?from\=xxx
        - 纯 ID: cgXXO4wMXj5T
        """
        if not input_str or not isinstance(input_str, str):
            return None
        
        url = input_str.strip()
        
        # 处理 shell 转义字符（如 \? \= \& 等）
        url = url.replace('\\?', '?').replace('\\=', '=').replace('\\&', '&')
        
        # 去除可能存在的查询参数和锚点
        base_url = url.split('?')[0].split('#')[0]
        
        # 取 URL 路径的最后一段作为 ID
        # 例如 https://kdocs.cn/l/cgXXO4wMXj5T -> cgXXO4wMXj5T
        # 例如 https://www.kdocs.cn/p/123456 -> 123456
        if '/' in base_url:
            parts = base_url.rstrip('/').split('/')
            return parts[-1]
            
        return url

    @staticmethod
    def decode_unique_id(unique_id):
        """
        解码 WPS 文件 ID (Unique ID) 获取整型 fileID 和 linkID
        """
        if not unique_id or len(unique_id) < 24: # minLength check
            return None, None

        encodeTable = "M9x1rB8cboz5NCsVDyh7FWjYvLAXfZn3tupgaJeUPGET4KdiHS2wm6RkqQ"
        decodeTable = [0xFF] * 256
        for i, c in enumerate(encodeTable):
            decodeTable[ord(c)] = i

        magicNumber = 190001718551341
        blockSize = 8
        encodeBlockSize = 11

        def decode_block(block):
            base = len(encodeTable)
            max_index = 255
            x = 0
            for i in range(len(block) - 1, -1, -1):
                char_code = ord(block[i])
                if char_code > max_index or decodeTable[char_code] == 0xFF:
                    raise ValueError(f"Illegal character: {block[i]}")
                x = x * base + decodeTable[char_code]
            x ^= magicNumber
            return struct.pack('<Q', x)

        def decode_last_block(block):
            base = len(encodeTable)
            max_index = 255
            x = 0
            for i in range(len(block) - 1, -1, -1):
                char_code = ord(block[i])
                if char_code > max_index or decodeTable[char_code] == 0xFF:
                    raise ValueError(f"Illegal character: {block[i]}")
                x = x * base + decodeTable[char_code]
            result = bytearray(struct.pack('<Q', x))
            j = blockSize - 1
            while j >= 0:
                if result[j] != 0: break
                j -= 1
            return bytes(result[:j+1])

        def mix(data):
            length = len(data)
            mid = length >> 1
            for i in range(0, mid, 2):
                data[i], data[length-1-i] = data[length-1-i], data[i]

        try:
            length = len(unique_id)
            remainder = length % encodeBlockSize
            process_length = length
            if remainder > 0:
                process_length -= remainder
            
            decoded = bytearray()
            for i in range(0, process_length, encodeBlockSize):
                decoded.extend(decode_block(unique_id[i : i+encodeBlockSize]))
            
            mix(decoded)
            
            if remainder > 0:
                decoded.extend(decode_last_block(unique_id[process_length:]))
            
            file_id = struct.unpack('<Q', decoded[0:8])[0]
            link_id_len = decoded[8]
            link_id_end = 9 + link_id_len
            link_id = decoded[9:link_id_end].decode('utf-8')
            
            return file_id, link_id
        except Exception as e:
            return None, None


    def _kso1_sign(self, method, uri, content_type, request_body):
        kso_date = datetime.utcnow().strftime('%a, %d %b %Y %H:%M:%S GMT')
        
        sha256_hex = ''
        if request_body:
            sha256_obj = hashlib.sha256()
            if isinstance(request_body, str):
                sha256_obj.update(request_body.encode('utf-8'))
            elif isinstance(request_body, bytes):
                sha256_obj.update(request_body)
            else:
                sha256_obj.update(b'')
            sha256_hex = sha256_obj.hexdigest()

        signature_string = f'KSO-1{method}{uri}{content_type}{kso_date}{sha256_hex}'
        mac = hmac.new(bytes(self.secret_key, 'utf-8'),
                       bytes(signature_string, 'utf-8'),
                       hashlib.sha256)
        kso_signature = mac.hexdigest()
        
        authorization = f'KSO-1 {self.access_key}:{kso_signature}'

        headers = {
            'X-Kso-Date': kso_date,
            'Content-Type': content_type,
            'X-Kso-Authorization': authorization
        }
        
        if self.access_token and not uri.endswith("/oauth2/token"):
             headers['Authorization'] = "Bearer " + self.access_token
             
        return headers

    def get_token(self):
        if self.access_token:
            return self.access_token
            
        uri = "/oauth2/token"
        url = self.domain + uri
        data = {
            "grant_type": "client_credentials",
            "client_id": self.access_key,
            "client_secret": self.secret_key
        }
        body = urllib.parse.urlencode(data)
        headers = self._kso1_sign("POST", uri, "application/x-www-form-urlencoded", body)
        
        response = requests.post(url, data=data, timeout=30)
        if response.status_code == 200:
            self.access_token = response.json().get('access_token')
            return self.access_token
        else:
            raise Exception(f"获取 Token 失败: {response.text}")

    # 支持的 AP 文档类型后缀
    AP_SUPPORTED_EXTENSIONS = ['.otl']
    
    # 常见的 WPS 文档类型映射（用于友好提示）
    FILE_TYPE_NAMES = {
        '.docx': 'Word 文档',
        '.doc': 'Word 文档',
        '.wps': 'WPS 文字文档',
        '.xlsx': 'Excel 表格',
        '.xls': 'Excel 表格',
        '.et': 'WPS 表格',
        '.ksheet': 'WPS 智能表格',
        '.pptx': 'PowerPoint 演示',
        '.ppt': 'PowerPoint 演示',
        '.dpp': 'WPS 演示',
        '.pdf': 'PDF 文档',
        '.otl': 'WPS 智能文档 (AP)',
    }

    def get_file_meta(self, file_id, return_error=False):
        """
        根据 file_id 获取文件信息
        API: GET /v7/files/{file_id}/meta
        
        Args:
            file_id: 文件 ID
            return_error: 如果为 True，失败时返回包含错误信息的字典而非 None
            
        Returns:
            成功时返回 API 响应 JSON
            失败时：
            - return_error=False: 返回 None
            - return_error=True: 返回 {"_error": True, "status": 400, "code": xxx, "msg": xxx}
        """
        self.get_token()
        uri = f"/v7/files/{file_id}/meta"
        url = self.domain + uri
        
        headers = self._kso1_sign("GET", uri, "application/json", None)
        
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            # 打印fileid 
            print(f"[WpsClient] get_file_meta 请求失败: Status={response.status_code}, Msg={response.text} FileId={file_id}")
            if return_error:
                # 尝试解析 JSON 错误响应
                error_info = {"_error": True, "status": response.status_code}
                try:
                    err_json = response.json()
                    error_info["code"] = err_json.get("code")
                    error_info["msg"] = err_json.get("msg")
                except:
                    error_info["msg"] = response.text
                return error_info
            return None

    def check_ap_document(self, file_id, doc_url=None, silent=False):
        """
        检查文件是否为支持的 AP 文档格式。
        
        Args:
            file_id: 文件 ID
            doc_url: 原始文档 URL（用于错误提示）
            silent: 是否静默模式（不打印错误信息）
            
        Returns:
            (is_supported: bool, file_name: str, error_msg: str)
            - is_supported: 是否为支持的 AP 文档
            - file_name: 文件名（如果获取成功）
            - error_msg: 错误提示信息（如果不支持）
        """
        meta = self.get_file_meta(file_id, return_error=True)
        
        # 处理 API 请求失败的情况
        if meta and meta.get('_error'):
            status = meta.get('status')
            code = meta.get('code')
            
            # 根据错误码生成友好提示
            if status == 400 or status == 404 or code == 400000004:
                # 文档链接无效或已删除
                error_msg = f"""
❌ 无法访问该文档

   可能的原因:
   - 文档链接无效（ID 格式错误）
   - 文档已被删除
   - 文档不存在
   
   💡 请手动打开链接检查是否能正常访问:
   {doc_url or f'https://kdocs.cn/l/{file_id}'}
"""
                return False, None, error_msg
            
            elif status == 403 or code == 403000001:
                # 无权限访问
                error_msg = f"""
❌ 无权限访问该文档

   当前应用没有该文档的访问权限。
   
   💡 请手动打开链接检查是否能正常访问:
   {doc_url or f'https://kdocs.cn/l/{file_id}'}
   
   如果您能正常访问，可能需要检查应用的权限配置。
"""
                return False, None, error_msg
            
            else:
                # 其他未知错误，返回 None 让后续操作继续尝试
                return None, None, f"获取文件信息失败: {meta.get('msg', '未知错误')}"
        
        if not meta or meta.get('code') != 0:
            # 无法获取元数据，返回 None 让调用方决定如何处理
            return None, None, "无法获取文件信息"
        
        data = meta.get('data', {})
        file_name = data.get('name', '')
        
        # 检查文件后缀
        file_ext = ''
        for ext in self.AP_SUPPORTED_EXTENSIONS + list(self.FILE_TYPE_NAMES.keys()):
            if file_name.lower().endswith(ext):
                file_ext = ext
                break
        
        if file_ext in self.AP_SUPPORTED_EXTENSIONS:
            return True, file_name, None
        
        # 不支持的格式，生成友好提示
        type_name = self.FILE_TYPE_NAMES.get(file_ext, f"'{file_ext}' 格式文件" if file_ext else "未知格式文件")
        error_msg = f"""
❌ 不支持的文档格式

   文件名称: {file_name}
   文件类型: {type_name}
   
   AirpageFile Skill 仅支持操作 WPS 智能文档 (.otl)。
   
   💡 提示:
   - 如果您需要读取/操作表格文件，请使用 AirsheetFile Skill
   - 如果您需要处理其他格式的文档，请先在 WPS 中将其转换为智能文档格式
"""
        return False, file_name, error_msg

    def create_drive(self, name="Application Drive"):
        self.get_token()
        uri = "/v7/drives/create"
        url = self.domain + uri
        body = {
            "allotee_type": "app",
            "name": name
        }
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        response = requests.post(url, data=body_bytes, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] create_drive 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def create_file(self, name, file_type="file", drive_id=None, parent_id=None):
        self.get_token()
        target_drive = drive_id or self.drive_id
        target_parent = parent_id or self.parent_id
        
        if not target_drive:
            raise ValueError("需要 Drive ID。请先创建 Drive。")

        uri = f"/v7/drives/{target_drive}/files/{target_parent}/create"
        url = self.domain + uri
        
        if not name.endswith('.otl'):
            name += '.otl'
            
        body = {
            "file_type": file_type,
            "name": name,
            "on_name_conflict": "rename"
        }
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        response = requests.post(url, data=body_bytes, headers=headers)
        if response.status_code == 200:
            return response.json()
        raise Exception(f"创建文件失败: {response.text}")

    def convert_markdown_to_blocks(self, file_id, markdown_content):
        self.get_token()
        uri = f"/v7/airpage/{file_id}/blocks/convert"
        url = self.domain + uri
        
        markdown_content = markdown_content.replace('&nbsp;', ' ')
        
        # 预处理：将 @用户名 转换为特殊标记 **@@用户名@@**，便于后续识别
        markdown_content = self._preprocess_user_mentions(markdown_content)

        arg_data = {
            "format": "markdown",
            "content": markdown_content
        }
        arg_str = json.dumps(arg_data)
        arg_base64 = base64.b64encode(arg_str.encode('utf-8')).decode('utf-8')
        
        body = {"arg": arg_base64}
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        response = requests.post(url, data=body_bytes, headers=headers)
        if response.status_code == 200:
            res_json = response.json()
            if 'data' in res_json and 'result' in res_json['data']:
                result_bytes = base64.b64decode(res_json['data']['result'])
                blocks = json.loads(result_bytes.decode('utf-8'))
                # 处理自定义图片语法
                blocks = self._process_custom_images(blocks, markdown_content)
                # 处理 WPSDocument 链接 (自动转换 kdocs.cn 链接为 WPSDocument block)
                blocks = self._process_kdocs_links(blocks)
                # 处理 @用户提及，将特殊标记转换为 WPSUser block
                blocks = self._process_wps_user_mentions(blocks)
                #接口化打印 jump
                # print(f"[WpsClient] blocks: {json.dumps(blocks, indent=2, ensure_ascii=False)}")
                return blocks
            else:
                print(f"[WpsClient] convert_markdown_to_blocks 解析失败: {res_json}")
        else:
            print(f"[WpsClient] convert_markdown_to_blocks 请求失败: Status={response.status_code}, Msg={response.text}")
        return None

    def _process_custom_images(self, blocks_data, markdown_content, debug=False):
        """
        修正 Picture Block 的 sourceKey。
        原理：服务端转换会将 ![alt](sourceKey:xxx) 识别为 picture 类型，但会丢失 sourceKey 字段。
        我们需要从原始 Markdown 中按顺序提取 sourceKey，并回填到对应的 Picture Block 中。
        """
        if not markdown_content:
            return blocks_data

        # 1. 确定要遍历的目标列表
        target_list = []
        if isinstance(blocks_data, list):
            target_list = blocks_data
        elif isinstance(blocks_data, dict):
            if 'blocks' in blocks_data:
                target_list = blocks_data['blocks']
            elif 'content' in blocks_data:
                target_list = blocks_data['content']
            # 如果是单个 Block (如 heading/paragraph)，放入列表处理
            elif 'type' in blocks_data:
                target_list = [blocks_data]
        
        if not target_list:
            if debug:
                print("[WpsClient] _process_custom_images: 输入数据结构不包含可遍历的 Block List")
            return blocks_data

        # 2. 从 Markdown 中提取所有 sourceKey (按顺序)
        # 匹配格式: ![...](sourceKey:xxxx)
        source_keys = re.findall(r'!\[.*?\]\(sourceKey:([a-zA-Z0-9]+)\)', markdown_content)
        
        if not source_keys:
            if debug:
                print("[WpsClient] 未检测到自定义 sourceKey 图片。")
            return blocks_data
        
        if debug:
            print(f"[WpsClient] 检测到 {len(source_keys)} 个自定义图片 Key: {source_keys}")
        
        # 使用迭代器方便按顺序取值
        key_iter = iter(source_keys)
        self._matched_count = 0  # 计数器
        
        # 加载图片尺寸数据
        image_sizes = {}
        try:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            sizes_file = os.path.join(base_dir, 'output', 'image_sizes.json')
            if os.path.exists(sizes_file):
                with open(sizes_file, 'r', encoding='utf-8') as f:
                    image_sizes = json.load(f)
        except:
            pass
        
        def fill_keys(block_list):
            for block in block_list:
                # 递归处理容器
                if 'content' in block and isinstance(block['content'], list):
                    fill_keys(block['content'])
                
                # 找到 picture block 且 sourceKey 为空
                if block.get('type') == 'picture':
                    attrs = block.get('attrs', {})
                    # 只有当 sourceKey 为空时才回填，避免覆盖可能正确的值
                    current_key = attrs.get('sourceKey', '')
                    if not current_key:
                        try:
                            key = next(key_iter)
                            attrs['sourceKey'] = key
                            attrs['borderType'] = 1
                            # remove version (用户之前手动添加的逻辑保留)
                            attrs.pop('version', None)
                            
                            # --- 尝试回填尺寸 ---
                            if key in image_sizes:
                                size = image_sizes[key]
                                w = size.get('width')
                                h = size.get('height')
                                if w and h:
                                    attrs['width'] = w
                                    attrs['height'] = h
                                    # renderWidth/Height 通常也需要设置
                                    attrs['renderWidth'] = w
                                    attrs['renderHeight'] = h
                                    if debug:
                                        print(f"[WpsClient] -> 图片尺寸回填: {w}x{h}")
                            # ------------------
                            
                            self._matched_count += 1
                            if debug:
                                print(f"[WpsClient] -> 正在修复第 {self._matched_count} 张图片 (Block ID: {block.get('id')}) => sourceKey: {key}")
                        except StopIteration:
                            # Markdown 里的 key 用完了，但还有 picture block (可能是非 sourceKey 图片)
                            if debug:
                                print(f"[WpsClient] 警告: 剩余 Picture Block 数量超过了提取到的 Key 数量。")
                            pass
                    else:
                        if debug:
                            print(f"[WpsClient] 跳过已有 sourceKey 的图片 (Key: {current_key})")
        
        fill_keys(target_list)
        if debug and self._matched_count > 0:
            print(f"[WpsClient] 图片 Key 修复完成，共处理 {self._matched_count} 个。")
        return blocks_data

    def _process_kdocs_links(self, blocks_data):
        """
        遍历 Blocks，查找带 kdocs.cn 链接的 text 节点，将其转换为 WPSDocument Block。
        规则：
        1. 遍历所有 paragraph 的 content
        2. 找到 type=text 且带有 link 属性，href 包含 kdocs.cn 的节点
        3. 将该节点替换为 WPSDocument block
        """
        target_list = []
        if isinstance(blocks_data, list):
            target_list = blocks_data
        elif isinstance(blocks_data, dict):
            if 'blocks' in blocks_data:
                target_list = blocks_data['blocks']
            elif 'content' in blocks_data:
                target_list = blocks_data['content']
            elif 'type' in blocks_data:
                target_list = [blocks_data]
        
        if not target_list:
            return blocks_data

        for i, block in enumerate(target_list):
            # 递归处理子内容 (如表格内)
            if 'content' in block and isinstance(block['content'], list):
                # 递归调用会修改引用对象的内容
                self._process_kdocs_links(block['content'])

            if block.get('type') == 'paragraph':
                content_list = block.get('content', [])
                new_content = []
                
                for item in content_list:
                    if item.get('type') == 'text':
                        attrs = item.get('attrs', {})
                        link_info = attrs.get('link', {})
                        href = link_info.get('href', '')
                        text_content = item.get('content', '')
                        
                        if href and 'kdocs.cn' in href:
                            # 转换为 WPSDocument block
                            wps_block = self._create_wps_document_block(href, text_content)
                            if wps_block:
                                new_content.append(wps_block)
                                continue
                    
                    # 保留原节点
                    new_content.append(item)
                
                block['content'] = new_content

        return blocks_data

    def _create_wps_document_block(self, href, text_content):
        """
        根据 kdocs 链接创建 WPSDocument block。
        """
        # 提取 ID
        file_id = self.resolve_file_id(href)
        
        # 尝试获取文件真实元数据
        meta_res = self.get_file_meta(file_id)
        real_file_id = file_id  # 默认使用 link_id
        
        # 获取文件元数据以确定文档类型
        wps_document_type = "otl"  # 默认值
        wps_document_name = text_content
        wps_document_id = str(real_file_id)
        meta_failed = False  # 标记元数据获取是否失败
        
        if meta_res and meta_res.get('code') == 0 and meta_res.get('data'):
            meta_data = meta_res['data']
            # 优先使用 link_id，如果没有则使用 id
            real_file_id = meta_data.get('link_id') or meta_data.get('id') or file_id
            wps_document_name = meta_data.get('name', text_content)
            
            # 尝试从长 ID 解码出整型 ID
            long_id = meta_data.get('id')
            if long_id:
                decoded_fid, _ = self.decode_unique_id(long_id)
                if decoded_fid:
                    wps_document_id = str(decoded_fid)
            
            # 根据文件名后缀判断文档类型
            file_name_lower = wps_document_name.lower()
            if ".et" in file_name_lower or ".xlsx" in file_name_lower or ".xls" in file_name_lower or ".ksheet" in file_name_lower:
                wps_document_type = "et"
            elif ".dpp" in file_name_lower or ".ppt" in file_name_lower or ".pptx" in file_name_lower:
                wps_document_type = "dpp"
            elif ".otl" in file_name_lower:
                wps_document_type = "otl"
            elif ".wps" in file_name_lower or ".doc" in file_name_lower or ".docx" in file_name_lower:
                wps_document_type = "wps"
        else:
            # 元数据获取失败（通常是 403 无权限，文档未开启分享）
            meta_failed = True
            wps_document_name = f"{text_content}（未开启分享）"
        
        # 构造 WPSDocument block
        return {
            "type": "WPSDocument",
            "attrs": {
                "version": 1,
                "wpsDocumentId": str(wps_document_id),
                "wpsDocumentLink": href,
                "wpsDocumentName": wps_document_name,
                "wpsDocumentType": wps_document_type
            }
        }

    def _preprocess_user_mentions(self, markdown_content):
        """
        预处理 Markdown 内容，将 @用户名 转换为特殊标记 **@@用户名@@**
        这样在 WPS API 转换后，会生成粗体的 text block，便于后续识别和替换为 WPSUser block。
        
        支持的格式：
        - @用户名 (用户名为中文、英文、数字、下划线的组合)
        - @用户名 必须前面是空格或行首
        - @用户名 后跟空格、标点或行尾
        
        注意：避免匹配已有的链接或邮箱中的 @
        """
        if not markdown_content:
            return markdown_content
        
        # 匹配 @用户名 的正则
        # 用户名支持：中文、英文、数字、下划线，长度 1-20
        # 前置条件：@ 前面必须是空格、行首或特定标点（如括号）
        # 排除：邮箱中的 @、Markdown 链接中的 @、连续的中英文字符
        pattern = r'(?:^|(?<=\s)|(?<=[，。！？、；：""''（）【】]))@([\u4e00-\u9fa5a-zA-Z0-9_]{1,20})(?![a-zA-Z0-9@.])'
        
        def replace_mention(match):
            username = match.group(1)
            # 使用 @@username@@ 作为特殊标记，外层加粗
            return f'**@@{username}@@**'
        
        result = re.sub(pattern, replace_mention, markdown_content, flags=re.MULTILINE)
        return result

    def _get_user_name_map(self, force_refresh=False):
        """
        获取用户名到 userId 的映射字典。
        从配置的部门获取所有成员，构建 {name: userId} 的映射。
        结果会缓存，避免重复查询。
        
        Args:
            force_refresh: 是否强制刷新缓存
            
        Returns:
            dict: {用户名: userId} 的映射字典
        """
        if self._user_name_cache is not None and not force_refresh:
            return self._user_name_cache
        
        if not self.mention_dept_id:
            print("[WpsClient] 未配置 MENTION_DEPT_ID，无法解析 @用户提及")
            self._user_name_cache = {}
            return self._user_name_cache
        
        print(f"[WpsClient] 正在获取部门成员以构建用户名映射 (DeptID: {self.mention_dept_id})...")
        
        # 获取部门所有成员（递归获取子部门）
        members = self.get_all_dept_members(
            dept_id=self.mention_dept_id,
            recursive=True,
            with_user_detail=True
        )
        
        # 构建用户名到 userId 的映射
        name_map = {}
        for member in members:
            user_info = member.get('user_info', {})
            # 字段名是 user_name 而不是 name
            name = user_info.get('user_name')
            
            # userId 需要是数字形式，从 avatar URL 中提取
            # avatar 格式: https://img.qwps.cn/1385941926?...
            user_id = None
            avatar = user_info.get('avatar', '')
            if avatar:
                # 从 avatar URL 中提取数字 ID
                match = re.search(r'qwps\.cn/(\d+)', avatar)
                if match:
                    user_id = match.group(1)
            
            if name and user_id:
                # 如果有重名，保留第一个（或可以改为列表处理）
                if name not in name_map:
                    name_map[name] = str(user_id)
        
        print(f"[WpsClient] 用户名映射构建完成，共 {len(name_map)} 个用户")
        self._user_name_cache = name_map
        return self._user_name_cache

    def _process_wps_user_mentions(self, blocks_data):
        """
        后处理 blocks，将特殊标记 @@用户名@@ 的粗体 text 替换为 WPSUser block。
        
        处理逻辑：
        1. 先扫描是否存在 @@xxx@@ 标记（避免无目标时的多余请求）
        2. 找到 type=text 且内容包含 @@xxx@@ 的节点
        3. 查询用户名对应的 userId
        4. 将 text 节点拆分/替换为 WPSUser block
        """
        # 匹配 @@用户名@@ 的正则
        mention_pattern = re.compile(r'@@([\u4e00-\u9fa5a-zA-Z0-9_]{1,20})@@')
        
        # 先扫描是否存在标记，避免无目标时发起多余的用户映射请求
        def has_mention_markers(data):
            """递归检查是否存在 @@xxx@@ 标记"""
            if isinstance(data, list):
                return any(has_mention_markers(item) for item in data)
            elif isinstance(data, dict):
                if data.get('type') == 'text':
                    content = data.get('content', '')
                    if mention_pattern.search(content):
                        return True
                if 'content' in data:
                    return has_mention_markers(data['content'])
                if 'blocks' in data:
                    return has_mention_markers(data['blocks'])
            return False
        
        if not has_mention_markers(blocks_data):
            # 没有找到任何 @@xxx@@ 标记，直接返回
            return blocks_data
        
        # 确定存在标记，获取用户名映射
        user_map = self._get_user_name_map()
        if not user_map:
            # 没有用户映射，清理标记但不转换
            return self._clean_mention_markers(blocks_data)
        
        def process_content_list(content_list):
            """处理 content 列表，返回新的列表"""
            if not content_list or not isinstance(content_list, list):
                return content_list
            
            new_content = []
            for item in content_list:
                if not isinstance(item, dict):
                    new_content.append(item)
                    continue
                
                item_type = item.get('type')
                
                # 递归处理嵌套结构
                if 'content' in item and isinstance(item['content'], list):
                    # 对于包含 content 的节点，递归处理
                    # 包括：段落、标题、列表项、引用、表格、表格行、表格单元格等
                    if item_type in ['paragraph', 'heading', 'listItem', 'blockquote', 
                                     'table', 'tableRow', 'tableCell', 'bulletList', 'orderedList']:
                        item['content'] = process_content_list(item['content'])
                    new_content.append(item)
                    continue
                
                # 处理 text 节点
                if item_type == 'text':
                    text_content = item.get('content', '')
                    attrs = item.get('attrs', {})
                    is_bold = attrs.get('bold', False)
                    
                    # 查找 @@xxx@@ 标记
                    matches = list(mention_pattern.finditer(text_content))
                    
                    if not matches:
                        new_content.append(item)
                        continue
                    
                    # 有匹配，需要拆分 text 节点
                    last_end = 0
                    for match in matches:
                        username = match.group(1)
                        start, end = match.start(), match.end()
                        
                        # 前面的普通文本
                        if start > last_end:
                            prefix_text = text_content[last_end:start]
                            if prefix_text:
                                prefix_item = {'type': 'text', 'content': prefix_text}
                                # 保留原有属性（除了 bold，因为 @@xx@@ 外的文本不应是粗体）
                                if attrs:
                                    prefix_attrs = {k: v for k, v in attrs.items() if k != 'bold'}
                                    if prefix_attrs:
                                        prefix_item['attrs'] = prefix_attrs
                                new_content.append(prefix_item)
                        
                        # 查找用户 ID
                        user_id = user_map.get(username)
                        if user_id:
                            # 创建 WPSUser block
                            wps_user_block = {
                                'type': 'WPSUser',
                                'attrs': {
                                    'name': username,
                                    'userId': user_id
                                }
                            }
                            new_content.append(wps_user_block)
                            print(f"[WpsClient] @{username} -> WPSUser (userId: {user_id})")
                        else:
                            # 找不到用户，保留原文但去掉标记
                            fallback_text = f'@{username}'
                            fallback_item = {'type': 'text', 'content': fallback_text}
                            if attrs:
                                clean_attrs = {k: v for k, v in attrs.items() if k != 'bold'}
                                if clean_attrs:
                                    fallback_item['attrs'] = clean_attrs
                            new_content.append(fallback_item)
                            print(f"[WpsClient] 警告: 未找到用户 '{username}'，保留为普通文本")
                        
                        last_end = end
                    
                    # 后面剩余的普通文本
                    if last_end < len(text_content):
                        suffix_text = text_content[last_end:]
                        if suffix_text:
                            suffix_item = {'type': 'text', 'content': suffix_text}
                            if attrs:
                                suffix_attrs = {k: v for k, v in attrs.items() if k != 'bold'}
                                if suffix_attrs:
                                    suffix_item['attrs'] = suffix_attrs
                            new_content.append(suffix_item)
                else:
                    new_content.append(item)
            
            return new_content
        
        # 处理 blocks 数据结构
        if isinstance(blocks_data, list):
            for block in blocks_data:
                if 'content' in block and isinstance(block['content'], list):
                    block['content'] = process_content_list(block['content'])
        elif isinstance(blocks_data, dict):
            if 'blocks' in blocks_data:
                for block in blocks_data['blocks']:
                    if 'content' in block and isinstance(block['content'], list):
                        block['content'] = process_content_list(block['content'])
            elif 'content' in blocks_data:
                blocks_data['content'] = process_content_list(blocks_data['content'])
        
        return blocks_data

    def _clean_mention_markers(self, blocks_data):
        """
        清理 @@xxx@@ 标记，将其还原为 @xxx 普通文本。
        用于没有用户映射时的降级处理。
        """
        pattern = re.compile(r'@@([\u4e00-\u9fa5a-zA-Z0-9_]{1,20})@@')
        
        def clean_content(content_list):
            if not content_list or not isinstance(content_list, list):
                return content_list
            
            for item in content_list:
                if not isinstance(item, dict):
                    continue
                
                if 'content' in item:
                    if isinstance(item['content'], list):
                        clean_content(item['content'])
                    elif isinstance(item['content'], str):
                        item['content'] = pattern.sub(r'@\1', item['content'])
                        # 同时去掉粗体属性
                        if item.get('attrs', {}).get('bold'):
                            del item['attrs']['bold']
                            if not item['attrs']:
                                del item['attrs']
            
            return content_list
        
        if isinstance(blocks_data, list):
            for block in blocks_data:
                if 'content' in block:
                    clean_content(block['content'])
        elif isinstance(blocks_data, dict):
            if 'blocks' in blocks_data:
                for block in blocks_data['blocks']:
                    if 'content' in block:
                        clean_content(block['content'])
            elif 'content' in blocks_data:
                clean_content(blocks_data['content'])
        
        return blocks_data

    def update_block(self, file_id, block_id, content=None, attrs=None):
        """
        更新 Block 的内容 (content) 或属性 (attrs)。
        
        Args:
            file_id: 文档 ID
            block_id: 目标 Block ID
            content: 新的 Block 内容列表 (对应 operation: update_content)
            attrs: 新的属性字典 (对应 operation: update_attrs)
            
        注意: content 和 attrs 只能二选一，如果都提供，优先处理 content。
        """
        self.get_token()
        uri = f"/v7/airpage/{file_id}/blocks/update"
        url = self.domain + uri
        
        arg_data = {
            "blockId": block_id
        }
        
        if content is not None:
            arg_data["operation"] = "update_content"
            arg_data["content"] = content
        elif attrs is not None:
            arg_data["operation"] = "update_attrs"
            arg_data["attrs"] = attrs
        else:
            print("[WpsClient] update_block 错误: 必须提供 content 或 attrs")
            return None

        arg_str = json.dumps(arg_data)
        arg_base64 = base64.b64encode(arg_str.encode('utf-8')).decode('utf-8')
        
        body = {"arg": arg_base64}
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在更新 Block (ID: {block_id}, Op: {arg_data.get('operation')})...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] update_block 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def delete_blocks(self, file_id, parent_id, start_index, end_index):
        """
        删除指定范围的 Block (删除 parent_id 下的子节点 [start_index, end_index))
        """
        self.get_token()
        uri = f"/v7/airpage/{file_id}/blocks/delete"
        url = self.domain + uri
        
        arg_data = {
            "blockId": parent_id,
            "startIndex": start_index,
            "endIndex": end_index
        }
        arg_str = json.dumps(arg_data)
        arg_base64 = base64.b64encode(arg_str.encode('utf-8')).decode('utf-8')
        
        body = {"arg": arg_base64}
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在删除 Blocks (Parent: {parent_id}, Range: [{start_index}, {end_index}))...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] delete_blocks 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def insert_blocks(self, file_id, blocks, index=1, block_id="doc"):
        self.get_token()
        uri = f"/v7/airpage/{file_id}/blocks/create"
        url = self.domain + uri
        
        arg_data = {
            "blockId": block_id,
            "index": index,
            "content": blocks
        }
        arg_str = json.dumps(arg_data)
        
        # 调试输出
        # print(json.dumps(arg_data, indent=2, ensure_ascii=False))

        arg_base64 = base64.b64encode(arg_str.encode('utf-8')).decode('utf-8')
        
        body = {"arg": arg_base64}
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在操作 Block (Parent: {block_id}, Index: {index})...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] 操作失败: Status={response.status_code}, Msg={response.text}")
            return None

    def update_title(self, file_id, new_title, align=2):
        """
        更新文档的主标题。
        流程: 获取文档结构 -> 找到 Title Block ID -> 更新其文本内容
        """
        print(f"正在尝试更新文档 {file_id} 的标题为: {new_title}")
        
        # 1. 获取文档结构
        blocks_data = self.get_blocks(file_id)
        if not blocks_data or 'blocks' not in blocks_data:
            print("无法获取文档结构，跳过标题更新。")
            return False
            
        # 2. 查找 Title Block ID
        title_id = None
        
        # 辅助函数：递归查找 title block
        def find_title(blocks_list):
            for block in blocks_list:
                if block.get('type') == 'title':
                    return block.get('id')
                # 如果有子内容，继续递归查找
                if 'content' in block and isinstance(block['content'], list):
                    res = find_title(block['content'])
                    if res: return res
            return None

        if blocks_data and 'blocks' in blocks_data:
            title_id = find_title(blocks_data['blocks'])
        
        if not title_id:
            print("未在文档中找到 Title Block (type=title)，无法更新。")
            # 调试：打印一下结构方便排查
            # print(json.dumps(blocks_data, indent=2, ensure_ascii=False))
            return False
            
        print(f"找到 Title Block ID: {title_id}")
        
        # 3. 构造更新 Payload
        # 注意：这里是向 Title Block 插入/替换文本内容
        # 格式: { "type": "text", "attrs": {"align": 2}, "content": "新标题" }
        update_content = [
            {
                "type": "text",
                "attrs": {
                    "align": align
                },
                "content": new_title
            }
        ]
        
        # 调用 insert_blocks (其实是 update)，传入 title_id 作为 parent
        # index=0 表示覆盖/插入到第一个位置 (对于文本内容通常是替换效果)
        res = self.insert_blocks(file_id, update_content, index=0, block_id=title_id)
        return res is not None

    def get_blocks(self, file_id, block_id="doc"):
        self.get_token()
        uri = f"/v7/airpage/{file_id}/blocks"
        url = self.domain + uri
        
        arg_data = {
            "blockId": block_id
        }
        arg_str = json.dumps(arg_data)
        arg_base64 = base64.b64encode(arg_str.encode('utf-8')).decode('utf-8')
        
        body = {"arg": arg_base64}
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        response = requests.post(url, data=body_bytes, headers=headers)
        if response.status_code == 200:
            res_json = response.json()
            if 'data' in res_json and 'result' in res_json['data']:
                result_bytes = base64.b64decode(res_json['data']['result'])
                return json.loads(result_bytes.decode('utf-8'))
            else:
                print(f"[WpsClient] get_blocks 解析失败: {res_json}")
        else:
            print(f"[WpsClient] get_blocks 请求失败: Status={response.status_code}, Msg={response.text}")
        return None

    def get_attachment_info(self, file_id, attachment_id):
        self.get_token()
        uri = f"/v7/coop/files/{file_id}/attachments/{attachment_id}"
        url = self.domain + uri
        
        # GET 请求通常没有 Body，Content-Type 可以为空或 application/json
        headers = self._kso1_sign("GET", uri, "application/json", None)
        
        print(f"[WpsClient] 正在获取附件信息 (File: {file_id}, Attachment: {attachment_id})...")
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] get_attachment_info 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def batch_get_attachment_info(self, file_id, attachment_ids):
        """
        批量查询附件文件信息
        
        Args:
            file_id: 文件ID
            attachment_ids: 附件ID列表，可以是 list 或单个字符串
        
        Returns:
            dict: 包含 attachment_infos 列表的响应数据，格式：
                {
                    "data": {
                        "attachment_infos": [
                            {
                                "attachment_id": "string",
                                "download_url": "string",
                                "name": "string",
                                "size": 0
                            }
                        ]
                    },
                    "code": 0,
                    "msg": "string"
                }
            失败时返回 None
        """
        self.get_token()
        uri = f"/v7/documents/{file_id}/attachments/batch_get"
        url = self.domain + uri
        
        # 确保 attachment_ids 是列表格式
        if isinstance(attachment_ids, str):
            attachment_ids = [attachment_ids]
        elif not isinstance(attachment_ids, list):
            print(f"[WpsClient] batch_get_attachment_info 参数错误: attachment_ids 必须是 list 或 str")
            return None
        
        body = {
            "attachment_ids": attachment_ids
        }
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在批量获取附件信息 (File: {file_id}, 附件数量: {len(attachment_ids)})...")
        response = requests.post(url, data=body_bytes, headers=headers, verify=False)
        
        if response.status_code == 200:
            result = response.json()
            if result.get('code') == 0:
                attachment_infos = result.get('data', {}).get('attachment_infos', [])
                print(f"[WpsClient] 成功获取 {len(attachment_infos)} 个附件信息")
                return result
            else:
                print(f"[WpsClient] batch_get_attachment_info 业务错误: Code={result.get('code')}, Msg={result.get('msg')}")
                return None
        else:
            print(f"[WpsClient] batch_get_attachment_info 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def open_share(self, file_id, drive_id=None):
        self.get_token()
        target_drive = drive_id or self.drive_id
        uri = f"/v7/drives/{target_drive}/files/{file_id}/open_link"
        url = self.domain + uri
        
        body = {
            "opts": {
                "allow_perm_apply": True,
                "close_after_expire": True,
                "expire_period": 0,
                "expire_time": 0
            },
            "role_id": "12", # Default read-only role usually
            "scope": "anyone"
        }
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        response = requests.post(url, data=body_bytes, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] open_share 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def get_users_by_emails(self, emails, status=["active"], with_dept=False):
        """
        根据邮箱获取用户列表
        API: POST /v7/users/by_emails
        
        Args:
            emails (list): 邮箱列表
            status (list): 用户状态列表，默认为 ["active"]。可选值: active, notactive, disabled
            with_dept (bool): 是否返回部门信息，默认为 False
            
        Returns:
            dict: API 响应结果
        """
        self.get_token()
        uri = "/v7/users/by_emails"
        url = self.domain + uri
        
        # 兼容单个状态传入
        if isinstance(status, str):
            status = [status]
            
        body = {
            "emails": emails,
            "status": status,
            "with_dept": with_dept
        }
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在根据邮箱获取用户信息 (Count: {len(emails)})...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            # 检查是否返回空数据（可能是权限问题）
            if result.get('code') == 0:
                items = result.get('data', {}).get('items', [])
                if not items:
                    self._print_contact_permission_hint(empty_result=True)
            return result
        else:
            print(f"[WpsClient] get_users_by_emails 请求失败: Status={response.status_code}, Msg={response.text}")
            self._print_contact_permission_hint(response)
            return None

    def get_users_by_ex_user_ids(self, ex_user_ids, status=["active"]):
        """
        根据外部身份源ID (ex_user_id) 获取用户信息
        API: POST /v7/users/by_ex_user_ids
        
        Args:
            ex_user_ids (list): 外部身份源ID列表
            status (list): 用户状态列表，默认为 ["active"]。可选值: active, notactive, disabled
            
        Returns:
            dict: API 响应结果
        """
        self.get_token()
        uri = "/v7/users/by_ex_user_ids"
        url = self.domain + uri
        
        # 兼容单个值传入
        if isinstance(ex_user_ids, str):
            ex_user_ids = [ex_user_ids]
        if isinstance(status, str):
            status = [status]
            
        body = {
            "ex_user_ids": ex_user_ids,
            "status": status
        }
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在根据外部ID获取用户信息 (Count: {len(ex_user_ids)})...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            # 检查是否返回空数据（可能是权限问题）
            if result.get('code') == 0:
                items = result.get('data', {}).get('items', [])
                if not items:
                    self._print_contact_permission_hint(empty_result=True)
            return result
        else:
            print(f"[WpsClient] get_users_by_ex_user_ids 请求失败: Status={response.status_code}, Msg={response.text}")
            self._print_contact_permission_hint(response)
            return None

    def get_user_depts(self, user_id):
        """
        获取用户所在部门列表
        API: GET /v7/users/{user_id}/depts
        
        Args:
            user_id (str): 用户 ID
            
        Returns:
            dict: API 响应结果，包含用户所在的部门列表
                {
                    "code": 0,
                    "msg": "",
                    "data": {
                        "items": [
                            {
                                "abs_path": "部门绝对路径",
                                "ctime": 创建时间戳,
                                "ex_dept_id": "外部身份源部门ID",
                                "id": "部门ID",
                                "leaders": [{"order": 0, "user_id": "领导用户ID"}],
                                "name": "部门名称",
                                "order": 排序值,
                                "parent_id": "父部门ID"
                            }
                        ]
                    }
                }
        """
        self.get_token()
        uri = f"/v7/users/{user_id}/depts"
        url = self.domain + uri
        
        headers = self._kso1_sign("GET", uri, "application/json", None)
        
        print(f"[WpsClient] 正在获取用户所在部门列表 (UserID: {user_id})...")
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            # 检查是否返回空数据（可能是权限问题）
            if result.get('code') == 0:
                items = result.get('data', {}).get('items', [])
                if not items:
                    self._print_contact_permission_hint(empty_result=True)
            return result
        else:
            print(f"[WpsClient] get_user_depts 请求失败: Status={response.status_code}, Msg={response.text}")
            self._print_contact_permission_hint(response)
            return None

    def get_dept_members(self, dept_id, status=["active"], page_size=50, page_token=None, 
                         recursive=False, with_user_detail=True):
        """
        查询部门下用户列表
        API: GET /v7/depts/{dept_id}/members
        
        Args:
            dept_id (str): 部门 ID
            status (list): 用户状态列表。可选值: active(正常), notactive(未激活), disabled(禁用)
            page_size (int): 分页大小，默认 50，最大 50
            page_token (str): 分页标记，首次请求不填
            recursive (bool): 是否递归查询子部门，默认 False
            with_user_detail (bool): 是否返回用户详情信息，默认 True
            
        Returns:
            dict: API 响应结果
                {
                    "code": 0,
                    "data": {
                        "items": [
                            {
                                "user_id": "用户ID",
                                "dept_id": "部门ID",
                                "is_leader": false,
                                "user_info": {...}  # with_user_detail=True 时返回
                            }
                        ],
                        "next_page_token": "下一页token"
                    }
                }
        """
        self.get_token()
        
        # 兼容单个状态传入
        if isinstance(status, str):
            status = [status]
        
        # 构建查询参数
        query_parts = []
        for s in status:
            query_parts.append(f"status={s}")
        query_parts.append(f"page_size={page_size}")
        if page_token:
            query_parts.append(f"page_token={page_token}")
        query_parts.append(f"recursive={str(recursive).lower()}")
        query_parts.append(f"with_user_detail={str(with_user_detail).lower()}")
        
        query_string = "&".join(query_parts)
        uri = f"/v7/depts/{dept_id}/members?{query_string}"
        url = self.domain + uri
        
        headers = self._kso1_sign("GET", uri, "application/json", None)
        
        print(f"[WpsClient] 正在查询部门成员列表 (DeptID: {dept_id}, Recursive: {recursive})...")
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            # 检查是否返回空数据（可能是权限问题）
            if result.get('code') == 0:
                items = result.get('data', {}).get('items', [])
                if not items:
                    self._print_contact_permission_hint(empty_result=True)
            return result
        else:
            print(f"[WpsClient] get_dept_members 请求失败: Status={response.status_code}, Msg={response.text}")
            self._print_contact_permission_hint(response)
            return None

    def get_all_dept_members(self, dept_id, status=["active"], recursive=False, with_user_detail=True):
        """
        获取部门下所有成员（自动分页获取全部）
        
        Args:
            dept_id (str): 部门 ID
            status (list): 用户状态列表
            recursive (bool): 是否递归查询子部门
            with_user_detail (bool): 是否返回用户详情
            
        Returns:
            list: 所有成员列表
        """
        all_members = []
        page_token = None
        page_count = 0
        
        while True:
            page_count += 1
            result = self.get_dept_members(
                dept_id=dept_id,
                status=status,
                page_size=50,
                page_token=page_token,
                recursive=recursive,
                with_user_detail=with_user_detail
            )
            
            if not result or result.get('code') != 0:
                print(f"[WpsClient] 获取部门成员失败，已获取 {len(all_members)} 个成员")
                break
            
            data = result.get('data', {})
            items = data.get('items', [])
            all_members.extend(items)
            
            # 检查是否还有下一页
            page_token = data.get('next_page_token')
            if not page_token:
                break
            
            print(f"[WpsClient] 已获取 {len(all_members)} 个成员，继续获取下一页...")
        
        print(f"[WpsClient] 部门成员获取完成，共 {len(all_members)} 个成员")
        return all_members

    def _print_contact_permission_hint(self, response=None, empty_result=False):
        """
        打印通讯录接口权限提示信息
        
        Args:
            response: HTTP 响应对象（接口调用失败时传入）
            empty_result: 是否为空结果提示（接口成功但数据为空）
        """
        if empty_result:
            print("\n⚠️  [权限提示] 查询结果为空，可能原因：")
            print("   1. 查询的用户/邮箱/外部ID 确实不存在")
            print("   2. 数据权限未开通：权限管理 -> 数据权限 -> 选择可访问的部门/人员范围")
            print("   注意：能力权限和数据权限都需要走审批流程。")
            print("   详情请访问 WPS 开放平台 (open.wps.cn) 进行配置。\n")
            return
            
        try:
            err_json = response.json()
            err_code = err_json.get('code')
            
            # 403 权限不足 或 特定错误码
            if response.status_code == 403 or err_code == 403000001:
                print("\n⚠️  [权限提示] 通讯录接口访问失败，可能需要开通以下权限：")
                print("   1. 能力权限：权限管理 -> 能力权限 -> kso.contact.read / kso.contact.readwrite")
                print("   2. 数据权限：权限管理 -> 数据权限 -> 选择可访问的部门/人员范围")
                print("   注意：以上两项权限都需要走审批流程。")
                print("   详情请访问 WPS 开放平台 (open.wps.cn) 进行配置。\n")
        except:
            pass

    def _calculate_md5(self, file_path):
        """计算文件MD5"""
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    def upload_attachment(self, file_id, file_path):
        """
        上传附件的完整流程：
        1. 申请上传地址
        2. 执行上传 (PUT/POST 到云存储)
        3. 提交上传完成
        """
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return None

        filename = os.path.basename(file_path)
        filesize = os.path.getsize(file_path)
        mime_type, _ = mimetypes.guess_type(file_path)
        if not mime_type:
            mime_type = "application/octet-stream"
        
        file_md5 = self._calculate_md5(file_path)

        # Step 1: 申请上传地址
        self.get_token()
        uri_addr = f"/v7/documents/{file_id}/attachments/upload/address"
        url_addr = self.domain + uri_addr
        
        body_addr = {
            "name": filename,
            "size": filesize,
            "content_type": mime_type,
            "md5": file_md5 
        }
        body_bytes_addr = json.dumps(body_addr)
        headers_addr = self._kso1_sign("POST", uri_addr, "application/json", body_bytes_addr)
        
        print(f"[WpsClient] 1/3 正在申请上传地址 ({filename}, {filesize} bytes)...")
        res_addr = requests.post(url_addr, data=body_bytes_addr, headers=headers_addr, verify=False)
        
        if res_addr.status_code != 200:
            print(f"申请上传地址失败: {res_addr.text}")
            if "invalid_scope" in res_addr.text and "kso.documents.readwrite" in res_addr.text:
                print("\n⚠️  [权限提示] 您的应用似乎未开通【文档管理能力】。")
                print("请前往 WPS 开放平台 (open.wps.cn) -> 应用管理 -> 权限管理，申请/开启 'kso.documents.readwrite' 权限。")
            return None
            
        data_addr = res_addr.json().get('data', {})
        # print(f"DEBUG: address_response={json.dumps(data_addr, indent=2)}")

        upload_req = data_addr.get('request', {})
        upload_url = upload_req.get('url')
        upload_method = upload_req.get('method', 'PUT')
        upload_headers = upload_req.get('headers', {})
        
        # 强制确保 Content-Type 一致
        if 'Content-Type' not in upload_headers and 'content-type' not in upload_headers:
             upload_headers['Content-Type'] = mime_type

        upload_id = data_addr.get('upload_id')
        send_back_params = data_addr.get('send_back_params', {})
        
        if not upload_url or not upload_id:
            print("获取上传地址响应异常 (缺少 url 或 upload_id)")
            return None

        # Step 2: 执行上传
        print(f"[WpsClient] 2/3 正在上传文件到云存储...")
        with open(file_path, 'rb') as f:
            # 注意：这里直接使用 requests 请求云存储，不需要 KSO 签名，但需要带上 API 返回的 headers
            res_upload = requests.request(upload_method, upload_url, headers=upload_headers, data=f, verify=False)
            
        if not (200 <= res_upload.status_code < 300):
            print(f"云存储上传失败: Status={res_upload.status_code}, Msg={res_upload.text}")
            return None
            
        # 尝试从响应头获取 etag (S3 通常会返回 ETag)
        # 调试：WPS 文档服务有时候对 ETag 校验敏感，如果 500 错误，尝试不回传 ETag
        # etag = res_upload.headers.get('ETag')
        # if etag:
            # ETag 通常带引号，视情况去除
            # etag = etag.strip('"')
            # send_back_params['etag'] = etag # 恢复 ETag 回传
            # pass
        
        # 动态处理 send_back_params 中的 header 指令
        # 例如: {"etag": "header.ETag", "key": "header.newfilename"}
        final_params = {}
        for k, v in send_back_params.items():
            if isinstance(v, str) and v.startswith("header."):
                header_key = v.split(".", 1)[1]
                # 尝试从 headers 获取，忽略大小写
                header_val = None
                for hk, hv in res_upload.headers.items():
                    if hk.lower() == header_key.lower():
                        header_val = hv
                        break
                
                if header_val:
                     # ETag 特殊处理去引号
                    if header_key.lower() == 'etag':
                        header_val = header_val.strip('"')
                    final_params[k] = header_val
                    # print(f"DEBUG: 解析 Params: {k} -> {header_val} (from header {header_key})")
                else:
                    print(f"WARNING: 未在响应头中找到 {header_key}")
                    final_params[k] = v # Fallback
            else:
                final_params[k] = v

        # Step 3: 提交上传完成
        print(f"[WpsClient] 3/3 提交上传完成通知...")
        uri_done = f"/v7/documents/{file_id}/attachments/upload/complete"
        url_done = self.domain + uri_done
        
        body_done = {
            "upload_id": upload_id,
            "params": final_params
        }
        body_bytes_done = json.dumps(body_done)
        headers_done = self._kso1_sign("POST", uri_done, "application/json", body_bytes_done)
        
        res_done = requests.post(url_done, data=body_bytes_done, headers=headers_done, verify=False)
        
        if res_done.status_code == 200:
            result = res_done.json()
            # print(f"DEBUG: result={json.dumps(result, indent=2)}")
            if result.get('code') == 0 or (result.get('data') and 'attachment_id' in result.get('data')):
                 print("✅ 上传成功！")
                 
                 # 尝试记录图片尺寸
                 try:
                     att_id = result.get('data', {}).get('attachment_id')
                     if att_id:
                         w, h = self.get_image_size(file_path)
                         if w and h:
                             self._save_image_size(att_id, w, h)
                 except Exception: 
                     pass
                 
                 return result
            else:
                 # 即使 HTTP 200，API 返回的 code 可能不是 0
                 print(f"提交上传完成失败 (Business Error): {json.dumps(result, ensure_ascii=False)}")
                 return None
        else:
            print(f"提交上传完成失败: {res_done.text}")
            return None

    @staticmethod
    def get_image_size(file_path):
        """获取图片尺寸 (width, height)"""
        try:
            from PIL import Image
            with Image.open(file_path) as img:
                return img.width, img.height
        except ImportError:
            # Fallback for macOS without Pillow
            if sys.platform == 'darwin':
                import subprocess
                try:
                    cmd = ['sips', '-g', 'pixelWidth', '-g', 'pixelHeight', file_path]
                    output = subprocess.check_output(cmd, stderr=subprocess.DEVNULL).decode()
                    w_match = re.search(r'pixelWidth: (\d+)', output)
                    h_match = re.search(r'pixelHeight: (\d+)', output)
                    if w_match and h_match:
                        return int(w_match.group(1)), int(h_match.group(1))
                except Exception:
                    pass
        except Exception:
            pass
        return None, None

    def _save_image_size(self, source_key, w, h):
        try:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            output_dir = os.path.join(base_dir, 'output')
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            sizes_file = os.path.join(output_dir, 'image_sizes.json')
            
            data = {}
            if os.path.exists(sizes_file):
                try:
                    with open(sizes_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                except: pass
            
            data[source_key] = {'width': w, 'height': h}
            
            with open(sizes_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except: pass
