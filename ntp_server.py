#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简易 NTP 服务器 (RFC 5905, IPv4/IPv6) —— 用于 WAF 客户端协议测试
功能:
  1. 正常模式: 作为 stratum N 的时间服务器应答, 支持人为时间偏移注入
  2. fuzz 模式: --fuzz <preset> 按预设返回畸形报文, 观察客户端行为
  3. 记录每个客户端请求(地址/端口/报文hex), 便于抓包比对
用法示例:
  python ntp_server.py                          # 正常服务器, 端口123, stratum 2
  python ntp_server.py --port 1230 --stratum 1  # 非特权端口测试
  python ntp_server.py --offset 3600            # 时间快1小时(测时间跳变防护)
  python ntp_server.py --fuzz bad_stratum       # 畸形stratum=255
  python ntp_server.py --fuzz kiss_of_death     # KoD RATE 限速应答
Fuzz 预设:
  bad_li_vn_mode   li_vn_mode 非法 (leap=3,version=7,mode=9)
  bad_stratum      stratum=255, poll=99, precision=-99
  zero_timestamp   Transmit Timestamp 全零
  year1927         Transmit Timestamp 置极端历史时间
  bad_root         root_delay/root_dispersion 极值
  kiss_of_death    stratum=0, REFID="RATE" (KoD)
  oversized        追加 512B 非法扩展字段(超长报文)
  random_garbage   随机垃圾数据
"""
import argparse
import random
import socket
import struct
import sys
import time

NTP_DELTA = 2208988800  # 1900 -> 1970

def to_ntp(t: float) -> int:
    if t <= 0:
        t = 1e-6
    return (int(t) + NTP_DELTA) << 32 | int((t % 1) * (1 << 32))

def build_packet(mode=4, version=4, leap=0, stratum=2, poll=6, precision=-20,
                 root_delay=0x1000, root_disp=0x1000, ref_id=b'GPS\0',
                 ref_ts=None, orig_ts=0, recv_ts=None, xmit_ts=None,
                 extra=b''):
    p = struct.pack('!BBbb', (leap << 6) | (version << 3) | mode,
                    stratum, poll, precision)
    p += struct.pack('!III', root_delay, root_disp, int.from_bytes(ref_id[:4], 'big'))
    now = time.time()
    p += struct.pack('!QQQ', to_ntp(ref_ts or now), orig_ts,
                     to_ntp(recv_ts if recv_ts is not None else now))
    p += struct.pack('!Q', to_ntp(xmit_ts if xmit_ts is not None else now))
    return p + extra

def reply_for(req: bytes, args) -> bytes:
    orig = int.from_bytes(req[40:48], 'big') if len(req) >= 48 else 0
    now = time.time() + args.offset
    kw = dict(mode=4, version=args.version, leap=0, stratum=args.stratum,
              poll=args.poll, ref_id=args.refid.encode()[:4].ljust(4, b'\x00'),
              orig_ts=orig, recv_ts=now, xmit_ts=now)
    if args.fuzz == 'bad_li_vn_mode':
        return struct.pack('!BBbb', (3 << 6) | (7 << 3) | 9, 2, 6, -20) + b'\x00' * 44
    if args.fuzz == 'bad_stratum':
        kw.update(stratum=255, poll=99, precision=-99)
    elif args.fuzz == 'zero_timestamp':
        kw.update(recv_ts=0.000001, xmit_ts=0.000001)
    elif args.fuzz == 'year1927':
        kw.update(ref_ts=-86400*365*73, recv_ts=-86400*365*73, xmit_ts=-86400*365*73)
    elif args.fuzz == 'bad_root':
        kw.update(root_delay=0xFFFFFFFF, root_disp=0xFFFFFFFF)
    elif args.fuzz == 'kiss_of_death':
        kw.update(stratum=0, ref_id=b'RATE', leap=3)
    elif args.fuzz == 'oversized':
        return build_packet(**kw) + b'\xde\xad\xbe\xef' * 128
    elif args.fuzz == 'random_garbage':
        return bytes(random.randint(0, 255) for _ in range(48))
    return build_packet(**kw)

def main():
    ap = argparse.ArgumentParser(description='NTP test server for WAF client testing')
    ap.add_argument('--bind', default='0.0.0.0')
    ap.add_argument('--port', type=int, default=123)
    ap.add_argument('--stratum', type=int, default=2)
    ap.add_argument('--version', type=int, default=4)
    ap.add_argument('--poll', type=int, default=6)
    ap.add_argument('--refid', default='GPS')
    ap.add_argument('--offset', type=float, default=0.0,
                    help='时间偏移秒数, 如 3600 / -86400')
    ap.add_argument('--fuzz', choices=['bad_li_vn_mode', 'bad_stratum', 'zero_timestamp',
                                       'year1927', 'bad_root', 'kiss_of_death',
                                       'oversized', 'random_garbage'])
    ap.add_argument('--log', default='ntp_server.log')
    args = ap.parse_args()

    fam = socket.AF_INET6 if ':' in args.bind else socket.AF_INET
    s = socket.socket(fam, socket.SOCK_DGRAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    s.bind((args.bind, args.port))
    print(f'NTP server on [{args.bind}]:{args.port} stratum={args.stratum} '
          f'version={args.version} offset={args.offset:+g}s fuzz={args.fuzz or "normal"}', flush=True)
    log = open(args.log, 'a', encoding='utf-8')
    while True:
        data, addr = s.recvfrom(4096)
        resp = reply_for(data, args)
        s.sendto(resp, addr)
        line = f'{time.strftime("%F %T")} {addr[0]}:{addr[1]} req={data.hex()} resp={resp.hex()[:96]}'
        print(line, flush=True)
        log.write(line + '\n'); log.flush()

if __name__ == '__main__':
    try:
        main()
    except PermissionError:
        sys.exit('绑定端口无权限: 用管理员/root 运行, 或 --port 1230')
