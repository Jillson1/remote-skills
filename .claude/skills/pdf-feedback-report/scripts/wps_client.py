# -*- coding: utf-8 -*-
"""
WPS 在线表格 API 客户端（本 skill 内置，用于读取反馈数据源表格）。
仅包含：解析 file_id、获取 token、检测表格类型、读取选区数据。
"""
import sys
import io
import requests
import urllib3
import urllib.parse
import hashlib
import hmac
import os
import configparser
from datetime import datetime

if sys.platform == "win32" and hasattr(sys.stdout, "buffer"):
    if sys.stdout.encoding != "utf-8":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if sys.stderr.encoding != "utf-8":
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class WpsClient:
    """WPS OpenAPI 客户端，仅实现本 skill 所需的在线表格读取能力。"""

    def __init__(self, config_path=None):
        self.config = configparser.ConfigParser()
        if not config_path:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(base_dir, "config", "config.properties")
        if os.path.exists(config_path):
            self.config.read(config_path, encoding="utf-8")
            self.access_key = (self.config["DEFAULT"].get("ACCESS_KEY") or "").strip()
            self.secret_key = (self.config["DEFAULT"].get("SECRET_KEY") or "").strip()
            self.domain = (self.config["DEFAULT"].get("DOMAIN") or "https://openapi.wps.cn").strip().rstrip("/")
            self.drive_id = (self.config["DEFAULT"].get("DEFAULT_DRIVE_ID") or "").strip()
        else:
            self.access_key = os.getenv("WPS_ACCESS_KEY", "")
            self.secret_key = os.getenv("WPS_SECRET_KEY", "")
            self.domain = os.getenv("WPS_DOMAIN", "https://openapi.wps.cn").rstrip("/")
            self.drive_id = os.getenv("WPS_DRIVE_ID", "")
        self.access_token = None

    def resolve_file_id(self, input_str):
        if not input_str or not isinstance(input_str, str):
            return None
        base_url = input_str.split("?")[0].split("#")[0]
        if "/" in base_url:
            return base_url.rstrip("/").split("/")[-1]
        return input_str.strip()

    def _kso1_sign(self, method, uri, content_type, request_body):
        kso_date = datetime.utcnow().strftime("%a, %d %b %Y %H:%M:%S GMT")
        sha256_hex = ""
        if request_body:
            h = hashlib.sha256()
            if isinstance(request_body, str):
                h.update(request_body.encode("utf-8"))
            elif isinstance(request_body, bytes):
                h.update(request_body)
            else:
                h.update(b"")
            sha256_hex = h.hexdigest()
        signature_string = f"KSO-1{method}{uri}{content_type}{kso_date}{sha256_hex}"
        mac = hmac.new(
            self.secret_key.encode("utf-8"),
            signature_string.encode("utf-8"),
            hashlib.sha256,
        )
        authorization = f"KSO-1 {self.access_key}:{mac.hexdigest()}"
        headers = {
            "X-Kso-Date": kso_date,
            "Content-Type": content_type,
            "X-Kso-Authorization": authorization,
        }
        if self.access_token and not uri.endswith("/oauth2/token"):
            headers["Authorization"] = "Bearer " + self.access_token
        return headers

    def get_token(self):
        if self.access_token:
            return self.access_token
        uri = "/oauth2/token"
        url = self.domain + uri
        data = {
            "grant_type": "client_credentials",
            "client_id": self.access_key,
            "client_secret": self.secret_key,
        }
        body = urllib.parse.urlencode(data)
        headers = self._kso1_sign("POST", uri, "application/x-www-form-urlencoded", body)
        response = requests.post(url, data=data, timeout=30, headers=headers)
        if response.status_code == 200:
            self.access_token = response.json().get("access_token")
            return self.access_token
        raise Exception(f"获取 Token 失败: {response.text}")

    def get_file_meta(self, file_id):
        self.get_token()
        uri = f"/v7/files/{file_id}/meta"
        url = self.domain + uri
        headers = self._kso1_sign("GET", uri, "application/json", None)
        response = requests.get(url, headers=headers, timeout=30)
        if response.status_code == 200:
            return response.json()
        return None

    def detect_file_type(self, file_id):
        res = self.get_file_meta(file_id)
        if res and res.get("code") == 0:
            name = (res.get("data") or {}).get("name", "").lower()
            if name.endswith(".ksheet"):
                return "airsheet"
            if name.endswith((".xlsx", ".xls", ".et")):
                return "sheets"
        return "airsheet"

    def get_range_data(self, file_id, worksheet_id, row_from, row_to, col_from, col_to, file_type="auto"):
        self.get_token()
        if file_type == "auto":
            file_type = self.detect_file_type(file_id)
        if file_type not in ("airsheet", "sheets"):
            file_type = "airsheet"
        uri = f"/v7/{file_type}/{file_id}/worksheets/{worksheet_id}/range_data"
        url = self.domain + uri
        params = {
            "row_from": row_from,
            "row_to": row_to,
            "col_from": col_from,
            "col_to": col_to,
        }
        headers = self._kso1_sign("GET", uri, "application/json", None)
        response = requests.get(url, headers=headers, params=params, timeout=60)
        if response.status_code == 200:
            return response.json()
        return None
