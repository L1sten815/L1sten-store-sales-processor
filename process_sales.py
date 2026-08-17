#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单机直跑版（无需启动服务）：直接对原始 xlsx 跑加权日均销量处理。
复用 server.py 中同一套已校验的快速管线：
  - 读取：calamine（Rust 解析，惰性迭代，远快于 openpyxl）
  - 写入：自研流式 inline-string XLSX 写器（不经共享字符串表，内存低、速度快）
用法：
  python process_sales.py                        # 用默认权重 0.6/0.3/0.1 与内置路径
  python process_sales.py 输入.xlsx 输出.xlsx   # 指定输入输出
  python process_sales.py 输入.xlsx 输出.xlsx 0.6 0.3 0.1   # 指定权重
"""
import os, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import server  # 复用已验证的 process_file

DEFAULT_SRC = r"D:\Google\download\export6a7fc0a1df2da9.95077228_45542.xlsx"
DEFAULT_DST = r"D:\Google\download\export6a7fc0a1df2da9.95077228_45542_处理结果.xlsx"

def main():
    args = sys.argv[1:]
    src = args[0] if len(args) >= 1 else DEFAULT_SRC
    dst = args[1] if len(args) >= 2 else DEFAULT_DST
    w3 = float(args[2]) if len(args) >= 3 else server.DEF[0]
    w7 = float(args[3]) if len(args) >= 4 else server.DEF[1]
    w14 = float(args[4]) if len(args) >= 5 else server.DEF[2]

    if not os.path.exists(src):
        print("输入文件不存在:", src)
        sys.exit(1)

    print("输入 :", src)
    print("输出 :", dst)
    print("权重 : 三天=%.2f 七天=%.2f 十四天=%.2f" % (w3, w7, w14))
    t0 = time.time()
    h, n, _ = server.process_file(src, w3, w7, w14, dst)
    dt = time.time() - t0
    print("完成 ✅  处理 %d 行，耗时 %.1f 秒（%.0f 行/秒）" % (n, dt, n / dt if dt else 0))
    print("输出列数:", len(h))
    print("输出大小: %.1f MB" % (os.path.getsize(dst) / 1024 / 1024))

if __name__ == '__main__':
    main()
