#!/usr/bin/env python3
# coding=utf-8
"""
华为 WAF（海外版）API 用户登录 + 接口调用 
依据本地源码还原：
  - 公钥下发  : api/cms/production       (production_utils.py:234  pb=base64(PEM), pb_len=4096)
  - 登录加密  : RSA-OAEP(MGF1-SHA256, SHA256) 分段400B拼接后base64  (cryptor.py / rsa_tool.py:112)
  - 验证码    : API 用户组(group_id=6)豁免                          (user_func.py:832)
  - CSRF/签名 : API 组免 CSRF(jwt.py:701)；peer/* 免签名(jwt.py:1218)
  - 参数校验  : HaSyncTaskUpdateForm  (validators/ha.py:427)
依赖: pip install requests pycryptodome
用法:
  python waf_api_fuzz.py --ip 172.24.1.159 -u api15 -p '密码'            # 基线调用
"""
import argparse
import base64
import json
import sys

import requests
import urllib3
from Crypto.Cipher import PKCS1_OAEP
from Crypto.Hash import SHA256
from Crypto.PublicKey import RSA

urllib3.disable_warnings()


UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "Content-Type": "application/json"}

CHUNK = {1024: 100, 2048: 200, 3072: 300, 4096: 400}
# r_i18n 随设备语言走，中英文都匹配 (user_func.py:732 'Please change your password for the first time!')
FIRST_LOGIN_HINTS = ("first time", "初次登录", "修改密码")



def get_public_key(ip):
    r = requests.post(f"https://{ip}/api/cms/production", json={}, headers=UA, verify=False, timeout=15)
    d = r.json()["data"]
    return d["pb"], d["pb_len"]


def rsa_encrypt(pub_b64, bits, payload: bytes) -> str:
    cipher = PKCS1_OAEP.new(RSA.import_key(base64.b64decode(pub_b64)), SHA256)
    n = CHUNK.get(bits, 400)
    ct = b"".join(cipher.encrypt(payload[i:i + n]) for i in range(0, len(payload), n))
    return base64.b64encode(ct).decode()


def login(ip, user, pwd):
    pub, bits = get_public_key(ip)
    msg = rsa_encrypt(pub, bits, json.dumps(
        {"username": user, "password": pwd, "code": "ABCDE"}).encode())
    r = requests.post(f"https://{ip}/api/cms/user/login", json={"message": msg},
                      headers=UA, verify=False, timeout=15)
    d = r.json()
    token = d.get("token", {}).get("access_token", "")
    print(f"[*] login: success={d.get('success')} msg={d.get('msg')!r} token={'OK' if token else 'EMPTY'}")
    return token, d


def is_first_login(login_resp):
    msg = str(login_resp.get("msg", ""))
    return any(h in msg for h in FIRST_LOGIN_HINTS)


def changepass(ip, user, token, pwd, newpwd):
    pub, bits = get_public_key(ip)
    # PasswordForm 要求 code 必填(validators/account.py:78)；API 组服务端免真实验证码核对(account_func.py:1203)
    msg = rsa_encrypt(pub, bits, json.dumps(
            {"username": user, "oldpwd": pwd, "newpwd": newpwd, "confirmpwd": newpwd,
             "code": "ABCDE"}).encode())
    h = {**UA, "Authorization": f"Bearer {token}"}
    r = requests.post(f"https://{ip}/api/v1/account/changepwd", json={"message": msg}, headers=h, verify=False, timeout=15)
    d = r.json()
    print(f"[*] changepwd: success={d.get('success')} msg={d.get('msg')!r}")
    return d.get("success") is True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ip", required=True)
    ap.add_argument("-u", "--user", required=True)
    ap.add_argument("-p", "--password", required=True)
    ap.add_argument("-np", "--newpwd", required=True)
    args = ap.parse_args()

    token, resp = login(args.ip, args.user, args.password)
    if not token:
        sys.exit("[!] 未获取到 token，检查凭据/网络/API_LISTS 白名单")
    print(f"Authorization : Bearer {token}")
    
    if is_first_login(resp):
        print("[*] 检测到首次登录，自动调用改密接口")
        if not changepass(args.ip, args.user, token, args.password, args.newpwd):
            sys.exit("[!] 改密失败，检查新密码复杂度(≥8位, 特殊字符仅限 !@#%-, 不含用户名/弱密码表)")
        # 改密后 user.random 轮换、旧 token 作废(account_func.py:1325)，必须用新密码重登
        token, resp = login(args.ip, args.user, args.newpwd)
        if not token:
            sys.exit("[!] 改密后重登未获取到 token")
        print("[*] 改密后重登成功，继续后续调用")
    

if __name__ == "__main__":
    main()
