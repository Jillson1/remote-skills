# -*- coding: utf-8 -*-
import sys
import io
import json
import requests
import urllib3
import hashlib
import hmac
import base64
import urllib.parse
import os
import configparser
from datetime import datetime

# Fix Windows console encoding
if sys.platform == 'win32':
    if hasattr(sys.stdout, 'buffer') and sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'buffer') and sys.stderr.encoding != 'utf-8':
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class WpsClient:
    def __init__(self, config_path=None):
        self.config = configparser.ConfigParser()
        
        # 默认加载同级目录下 ../config/airsheet.properties
        if not config_path:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(base_dir, 'config', 'airsheet.properties')
        
        if os.path.exists(config_path):
            self.config.read(config_path, encoding='utf-8')
            self.access_key = self.config['DEFAULT'].get('ACCESS_KEY')
            self.secret_key = self.config['DEFAULT'].get('SECRET_KEY')
            self.domain = self.config['DEFAULT'].get('DOMAIN', 'https://openapi.wps.cn')
            self.drive_id = self.config['DEFAULT'].get('DEFAULT_DRIVE_ID')
            self.parent_id = self.config['DEFAULT'].get('DEFAULT_PARENT_ID', '0')
        else:
            # 回退到环境变量或空值
            self.access_key = os.getenv('WPS_ACCESS_KEY')
            self.secret_key = os.getenv('WPS_SECRET_KEY')
            self.domain = os.getenv('WPS_DOMAIN', 'https://openapi.wps.cn')
            self.drive_id = os.getenv('WPS_DRIVE_ID')
            self.parent_id = '0'

        self.access_token = None

    def resolve_file_id(self, input_str):
        """
        简单从输入字符串中解析 file_id。
        如果输入包含 /，则取最后一段作为 ID；否则直接返回输入值。
        """
        if not input_str or not isinstance(input_str, str):
            return None
        
        # 简单粗暴：取 URL 路径的最后一段
        base_url = input_str.split('?')[0].split('#')[0]
        
        if '/' in base_url:
            parts = base_url.rstrip('/').split('/')
            return parts[-1]
            
        return input_str.strip()

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

    def create_drive(self, name="应用驱动器"):
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

    def create_airsheet(self, name, drive_id=None, parent_id=None, parent_path=None):
        """
        创建智能表格
        API: POST /v7/airsheet/files
        """
        self.get_token()
        uri = "/v7/airsheet/files"
        url = self.domain + uri
        
        target_drive = drive_id or self.drive_id
        target_parent = parent_id or self.parent_id
        
        if not target_drive:
             raise ValueError("需要 Drive ID。请先配置或创建 Drive。")
             
        # 处理文件名后缀
        if not name.endswith('.ksheet'):
             name += '.ksheet'
             
        body = {
            "drive_id": target_drive,
            "name": name,
            "on_name_conflict": "rename"
        }
        
        if parent_path:
             body["parent_path"] = parent_path
        elif target_parent and target_parent != '0':
             body["parent_id"] = target_parent
        
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在创建智能表格: {name} ...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
             print(f"[WpsClient] create_airsheet 请求失败: Status={response.status_code}, Msg={response.text}")
             return None

    def get_file_meta(self, file_id):
        """
        获取文件元数据
        API: GET /v7/files/{file_id}/meta
        """
        self.get_token()
        uri = f"/v7/files/{file_id}/meta"
        url = self.domain + uri
        
        headers = self._kso1_sign("GET", uri, "application/json", None)
        
        # print(f"[WpsClient] 正在获取文件信息: {file_id} ...")
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] get_file_meta 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def detect_file_type(self, file_id):
        """
        根据文件元数据自动检测表格类型
        :return: 'airsheet' 或 'sheets'
        """
        res = self.get_file_meta(file_id)
        if res and res.get('code') == 0:
            filename = res['data'].get('name', '').lower()
            if filename.endswith('.ksheet'):
                return 'airsheet'
            elif filename.endswith(('.xlsx', '.xls', '.et')):
                return 'sheets'
            else:
                # 无法识别后缀时，默认回退到 airsheet 或报错，这里暂时默认 airsheet 但打印警告
                print(f"[WpsClient] 警告: 无法根据文件名 '{filename}' 识别表格类型，默认使用 'airsheet'")
                return 'airsheet'
        return 'airsheet' # 请求失败时的兜底

    def create_worksheet(self, file_id, name=None, position=None, col_width=None, file_type='auto'):
        """
        创建工作表
        API: POST /v7/{file_type}/{file_id}/worksheets
        :param file_type: 'auto'(自动检测), 'airsheet' (智能表格) 或 'sheets' (传统表格)
        """
        self.get_token()
        
        # 自动检测逻辑
        if file_type == 'auto':
            file_type = self.detect_file_type(file_id)
            print(f"[WpsClient] 自动识别文件类型为: {file_type}")
        
        # 兼容处理：确保 file_type 有效
        if file_type not in ['airsheet', 'sheets']:
            print(f"[WpsClient] 警告: 未知的 file_type '{file_type}'，默认使用 'airsheet'")
            file_type = 'airsheet'

        uri = f"/v7/{file_type}/{file_id}/worksheets"
        url = self.domain + uri
        
        # 默认插入到最后
        if not position:
             position = {"end": True}
             
        body = {
            "position": position
        }
        
        if name:
             body["name"] = name
        if col_width is not None:
             body["col_width"] = col_width
             
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        type_name = "智能表格" if file_type == 'airsheet' else "电子表格"
        print(f"[WpsClient] 正在为{type_name}创建工作表: {name if name else '未指定'} ...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] create_worksheet 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def get_worksheets(self, file_id, file_type='auto'):
        """
        获取Sheet列表信息
        API: GET /v7/{file_type}/{file_id}/worksheets
        """
        self.get_token()
        
        # 自动检测逻辑
        if file_type == 'auto':
            file_type = self.detect_file_type(file_id)
            # print(f"[WpsClient] 自动识别文件类型为: {file_type}")
        
        # 兼容处理
        if file_type not in ['airsheet', 'sheets']:
            file_type = 'airsheet'

        uri = f"/v7/{file_type}/{file_id}/worksheets"
        url = self.domain + uri
        
        headers = self._kso1_sign("GET", uri, "application/json", None)
        
        # print(f"[WpsClient] 正在获取工作表列表: {file_id} ...")
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] get_worksheets 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def add_rows(self, file_id, worksheet_id, range_data, file_type='auto'):
        """
        创建行 (追加数据)
        API: POST /v7/{file_type}/{file_id}/worksheets/{worksheet_id}/rows
        """
        self.get_token()

        # 自动检测逻辑
        if file_type == 'auto':
            file_type = self.detect_file_type(file_id)

        # 兼容处理
        if file_type not in ['airsheet', 'sheets']:
            file_type = 'airsheet'

        uri = f"/v7/{file_type}/{file_id}/worksheets/{worksheet_id}/rows"
        url = self.domain + uri
        
        body = {
            "range_data": range_data
        }
        
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        type_name = "智能表格" if file_type == 'airsheet' else "电子表格"
        print(f"[WpsClient] 正在向{type_name}工作表 {worksheet_id} 添加 {len(range_data)} 行数据...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] add_rows 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def get_range_data(self, file_id, worksheet_id, row_from, row_to, col_from, col_to, file_type='auto'):
        """
        获取单元格选区数据
        API: GET /v7/{file_type}/{file_id}/worksheets/{worksheet_id}/range_data
        """
        self.get_token()

        # 自动检测逻辑
        if file_type == 'auto':
            file_type = self.detect_file_type(file_id)

        # 兼容处理
        if file_type not in ['airsheet', 'sheets']:
            file_type = 'airsheet'

        uri = f"/v7/{file_type}/{file_id}/worksheets/{worksheet_id}/range_data"
        url = self.domain + uri
        
        params = {
            "row_from": row_from,
            "row_to": row_to,
            "col_from": col_from,
            "col_to": col_to
        }
        
        headers = self._kso1_sign("GET", uri, "application/json", None)
        
        # print(f"[WpsClient] 正在读取选区: Sheet {worksheet_id}, R{row_from}C{col_from}:R{row_to}C{col_to} ...")
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            #print(f"[WpsClient] get_range_data 请求成功: Status={response.status_code}, Msg={response.text}")
            return response.json()
        else:
            print(f"[WpsClient] get_range_data 请求失败: Status={response.status_code}, Msg={response.text}")
            return None

    def update_range_data(self, file_id, worksheet_id, range_data, file_type='auto'):
        """
        更新单元格选区数据
        API: POST /v7/{file_type}/{file_id}/worksheets/{worksheet_id}/range_data/batch_update
        """
        self.get_token()

        # 自动检测逻辑
        if file_type == 'auto':
            file_type = self.detect_file_type(file_id)

        # 兼容处理
        if file_type not in ['airsheet', 'sheets']:
            file_type = 'airsheet'

        uri = f"/v7/{file_type}/{file_id}/worksheets/{worksheet_id}/range_data/batch_update"
        url = self.domain + uri
        
        body = {
            "range_data": range_data
        }
        
        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)
        
        print(f"[WpsClient] 正在更新选区: Sheet {worksheet_id}, {len(range_data)} 个单元格 ...")
        response = requests.post(url, data=body_bytes, headers=headers)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WpsClient] update_range_data 请求失败: Status={response.status_code}, Msg={response.text}")
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
            "role_id": "12", # 默认只读角色
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

    def download_file(self, file_id, save_path=None):
        """
        获取文件下载链接并下载
        """
        self.get_token()
        uri = f"/v7/files/{file_id}/download"
        url = self.domain + uri
        
        headers = self._kso1_sign("GET", uri, "application/json", None)
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            res_json = response.json()
            if 'data' in res_json and 'url' in res_json['data']:
                download_url = res_json['data']['url']
                
                # 下载文件
                print(f"正在下载文件: {download_url[:50]}...")
                r = requests.get(download_url)
                if r.status_code == 200:
                    if save_path:
                        with open(save_path, 'wb') as f:
                            f.write(r.content)
                        return save_path
                    return r.content
                else:
                    print(f"下载文件内容失败: {r.status_code}")
                    return None
            else:
                print(f"获取下载链接失败: {res_json}")
                return None
        else:
            print(f"请求下载接口失败: {response.status_code} {response.text}")
            return None

    def upload_file(self, file_path, drive_id=None, parent_id=None, parent_path=None, on_name_conflict="rename"):
        """
        上传本地文件到云空间（三步上传）
        1. 请求文件上传信息 (request_upload)
        2. 上传实体文件到云存储
        3. 提交文件上传完成 (commit_upload)
        
        :param file_path: 本地文件路径
        :param drive_id: 驱动盘 ID（不传则使用配置默认值）
        :param parent_id: 父目录 ID（不传则使用配置默认值）
        :param parent_path: 相对路径数组，若指定路径不存在将自动创建
        :param on_name_conflict: 文件名冲突处理方式: rename(默认), overwrite, fail
        :return: 上传成功后的文件信息 dict, 失败返回 None
        """
        if not os.path.exists(file_path):
            print(f"[WpsClient] 文件不存在: {file_path}")
            return None

        target_drive = drive_id or self.drive_id
        target_parent = parent_id or self.parent_id or '0'

        if not target_drive:
            raise ValueError("需要 Drive ID。请先配置或创建 Drive。")

        filename = os.path.basename(file_path)
        filesize = os.path.getsize(file_path)

        # 计算文件哈希（公网必传，至少 md5 和 sha256 中的一种）
        print(f"[WpsClient] 正在计算文件哈希: {filename} ({filesize} bytes)...")
        sha256_hash = hashlib.sha256()
        md5_hash = hashlib.md5()
        with open(file_path, 'rb') as f:
            while True:
                chunk = f.read(8192)
                if not chunk:
                    break
                sha256_hash.update(chunk)
                md5_hash.update(chunk)

        hashes = [
            {"sum": sha256_hash.hexdigest(), "type": "sha256"},
            {"sum": md5_hash.hexdigest(), "type": "md5"}
        ]

        # ========== 第 1 步：请求文件上传信息 ==========
        self.get_token()
        uri = f"/v7/drives/{target_drive}/files/{target_parent}/request_upload"
        url = self.domain + uri

        body = {
            "name": filename,
            "size": filesize,
            "hashes": hashes,
            "on_name_conflict": on_name_conflict
        }
        if parent_path:
            body["parent_path"] = parent_path

        body_bytes = json.dumps(body)
        headers = self._kso1_sign("POST", uri, "application/json", body_bytes)

        print(f"[WpsClient] 第 1 步: 请求上传信息...")
        res_request = requests.post(url, data=body_bytes, headers=headers)

        if res_request.status_code != 200:
            print(f"[WpsClient] request_upload 失败: Status={res_request.status_code}, Msg={res_request.text}")
            return None

        res_json = res_request.json()
        if res_json.get('code', 0) != 0:
            print(f"[WpsClient] request_upload 业务错误: {res_json}")
            return None

        data = res_json.get('data', {})
        store_request = data.get('store_request', {})
        upload_id = data.get('upload_id')
        upload_url = store_request.get('url')
        upload_method = store_request.get('method', 'PUT')
        # 云存储可能需要特定请求头（如签名等），从响应中提取
        upload_headers = store_request.get('headers', {})

        if not upload_url or not upload_id:
            print(f"[WpsClient] 未获取到上传地址或 upload_id: {data}")
            return None

        print(f"[WpsClient] 获取到上传地址, method={upload_method}, upload_id={upload_id}")

        # ========== 第 2 步：上传实体文件到云存储 ==========
        print(f"[WpsClient] 第 2 步: 上传实体文件 ({filesize} bytes)...")

        # 读取文件内容（用于签名和上传）
        with open(file_path, 'rb') as f:
            file_content = f.read()

        # 云存储网关同样需要 KSO-1 签名认证
        # 从上传 URL 中提取 URI 路径
        parsed_url = urllib.parse.urlparse(upload_url)
        upload_uri = parsed_url.path
        
        content_type = 'application/octet-stream'
        
        # 使用已有的签名方法生成认证头（传入文件二进制内容计算 SHA256）
        upload_sign_headers = self._kso1_sign(upload_method, upload_uri, content_type, file_content)
        
        # 合并云存储返回的 headers（如果有的话）
        if upload_headers:
            upload_sign_headers.update(upload_headers)

        res_upload = requests.request(
            upload_method, upload_url,
            headers=upload_sign_headers,
            data=file_content,
            verify=False
        )

        if not (200 <= res_upload.status_code < 300):
            print(f"[WpsClient] 上传实体文件失败: Status={res_upload.status_code}, Msg={res_upload.text}")
            return None

        print(f"[WpsClient] 实体文件上传成功 (Status={res_upload.status_code})")

        # ========== 第 3 步：提交文件上传完成 ==========
        uri_commit = f"/v7/drives/{target_drive}/files/{target_parent}/commit_upload"
        url_commit = self.domain + uri_commit

        body_commit = {
            "upload_id": upload_id
        }
        body_commit_bytes = json.dumps(body_commit)
        headers_commit = self._kso1_sign("POST", uri_commit, "application/json", body_commit_bytes)

        print(f"[WpsClient] 第 3 步: 提交上传完成...")
        res_commit = requests.post(url_commit, data=body_commit_bytes, headers=headers_commit)

        if res_commit.status_code != 200:
            print(f"[WpsClient] commit_upload 失败: Status={res_commit.status_code}, Msg={res_commit.text}")
            return None

        commit_json = res_commit.json()
        if commit_json.get('code', 0) != 0:
            print(f"[WpsClient] commit_upload 业务错误: {commit_json}")
            return None

        file_data = commit_json.get('data', {})
        print(f"[WpsClient] ✅ 文件上传成功！")
        print(f"[WpsClient]   文件名: {file_data.get('name')}")
        print(f"[WpsClient]   文件ID: {file_data.get('id')}")
        print(f"[WpsClient]   大小: {file_data.get('size')} bytes")

        return commit_json

    def update_file_content(self, file_id, file_path):
        """
        更新文件内容 (上传新版本)
        """
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return None

        filename = os.path.basename(file_path)
        filesize = os.path.getsize(file_path)
        
        # 1. 获取上传地址
        self.get_token()
        uri_addr = f"/v7/files/{file_id}/upload/address"
        url_addr = self.domain + uri_addr
        
        body_addr = {
            "name": filename,
            "size": filesize
        }
        body_bytes = json.dumps(body_addr)
        headers = self._kso1_sign("POST", uri_addr, "application/json", body_bytes)
        
        res_addr = requests.post(url_addr, data=body_bytes, headers=headers)
        if res_addr.status_code != 200:
            print(f"申请上传地址失败: {res_addr.text}")
            return None
            
        data_addr = res_addr.json().get('data', {})
        upload_method = 'PUT'
        upload_url = None
        upload_headers = {}
        upload_id = data_addr.get('upload_id')
        
        if 'request' in data_addr:
             req = data_addr['request']
             upload_url = req.get('url')
             upload_method = req.get('method', 'PUT')
             upload_headers = req.get('headers', {})
        else:
             upload_url = data_addr.get('url')
             upload_method = data_addr.get('method', 'PUT')
             upload_headers = data_addr.get('headers', {})

        if not upload_url:
            print("未获取到上传 URL")
            return None

        # 2. 上传文件
        print("正在上传新版本...")
        with open(file_path, 'rb') as f:
            res_upload = requests.request(upload_method, upload_url, headers=upload_headers, data=f, verify=False)
            
        if not (200 <= res_upload.status_code < 300):
            print(f"上传数据失败: {res_upload.status_code} {res_upload.text}")
            return None
            
        # 3. 完成上传（提交）
        uri_done = f"/v7/files/{file_id}/upload/complete"
        url_done = self.domain + uri_done
        
        body_done = {
            "upload_id": upload_id
        }
        
        body_bytes_done = json.dumps(body_done)
        headers_done = self._kso1_sign("POST", uri_done, "application/json", body_bytes_done)
        
        res_done = requests.post(url_done, data=body_bytes_done, headers=headers_done)
        if res_done.status_code == 200:
            print("✅ 更新成功！")
            return res_done.json()
        else:
            print(f"提交更新失败: {res_done.text}")
            return None
