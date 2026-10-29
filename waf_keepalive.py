# -*- coding: utf-8 -*-
"""
WAF 会话保活：原样重放 Burp 导出的请求包
用法:  python waf_keepalive.py <Burp请求包文件>
       未指定文件则打印用法并退出
停止:  Ctrl+C
说明:
  - 每次新建 TCP/TLS 连接发送原始报文，忽略证书校验
  - 仅替换 3 个与业务无关的 hop-by-hop 头:
      Connection      -> close     (让服务器回完即关, 便于收取完整响应)
      Accept-Encoding -> identity  (响应不压缩, 便于明文检查失效特征)
      Content-Length  -> 按实际 body 长度重算 (防止文件尾部换行差异导致卡住)
  - 其余内容(含 Cookie / X-CSRFtoken / X-Sign 等)逐字节原样发送
"""
import socket
import ssl
import sys
import time

if len(sys.argv) < 2:
    sys.exit(f"用法: python {sys.argv[0]} <Burp请求包文件>\n"
             f"示例: python {sys.argv[0]} packet-body.txt")
PACKET_FILE = sys.argv[1]
INTERVAL = 60          # 重放间隔(秒), 远小于会话超时(该账号 30 分钟)即可
PORT = 443


def load_packet(path):
    raw = open(path, "rb").read().replace(b"\r\n", b"\n")
    head, sep, body = raw.partition(b"\n\n")
    if not sep:
        raise ValueError("文件不是合法的 HTTP 请求(缺少头部与 body 之间的空行)")
    lines = head.split(b"\n")
    request_line = lines[0]
    host = None
    out_headers = []
    for ln in lines[1:]:
        if b":" not in ln:
            continue
        name = ln.split(b":", 1)[0].strip().lower()
        if name == b"host":
            host = ln.split(b":", 1)[1].strip().decode()
        if name == b"connection":
            ln = b"Connection: close"
        elif name == b"accept-encoding":
            ln = b"Accept-Encoding: identity"
        elif name == b"content-length":
            ln = b"Content-Length: %d" % len(body)
        out_headers.append(ln)
    if not host:
        raise ValueError("请求文件中缺少 Host 头")
    packet = b"\r\n".join([request_line] + out_headers) + b"\r\n\r\n" + body
    return host, packet


def replay(host, packet):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    resp = b""
    with socket.create_connection((host, PORT), timeout=15) as sock:
        with ctx.wrap_socket(sock, server_hostname=host) as ssock:
            ssock.sendall(packet)
            while True:
                chunk = ssock.recv(65536)
                if not chunk:
                    break
                resp += chunk
    return resp


def main():
    host, packet = load_packet(PACKET_FILE)
    print(f"[*] 目标 https://{host}:{PORT}  每 {INTERVAL}s 重放一次, Ctrl+C 停止")
    n = 0
    while True:
        n += 1
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        try:
            resp = replay(host, packet)
            status = resp.split(b"\r\n", 1)[0].decode(errors="replace")
            text = resp.partition(b"\r\n\r\n")[2].decode("utf-8", errors="replace")
            # error_code 10000 = 认证失败(会话已失效), 9000 段为各类 token 异常
            if '"error_code": 10000' in text or "CSRF" in text or not resp.startswith(b"HTTP/"):
                print(f"[{ts}] #{n} {status} !! 会话疑似失效: {text[:200]}")
            else:
                print(f"[{ts}] #{n} {status} OK")
        except Exception as e:
            print(f"[{ts}] #{n} 发送失败: {e}")
        time.sleep(INTERVAL)


if __name__ == "__main__":
    main()
