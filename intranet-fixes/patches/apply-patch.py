#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
字节补丁工具：app.asar 中 "disabled" → "onDemand" 等长替换
用法:
    python3 apply-patch.py <app.asar路径> apply      # 打补丁（自动备份到同目录 .bak）
    python3 apply-patch.py <app.asar路径> rollback   # 回滚（使用 .bak 恢复）
    python3 apply-patch.py <app.asar路径> verify     # 校验当前状态
"""
import sys, os, shutil

FROM = b"disabled"
TO   = b"onDemand"
OFFSET = 280109971  # 用户内网实测偏移；verify 时若不符会自动搜索

def load(path):
    with open(path, "rb") as f:
        return bytearray(f.read())

def save(path, data):
    with open(path, "wb") as f:
        f.write(data)

def find_offset(data, target, start=0):
    idx = data.find(target, start)
    return idx if idx >= 0 else None

def verify(path, verbose=True):
    data = load(path)
    at_off = data[OFFSET:OFFSET+len(FROM)] == FROM
    search = find_offset(data, FROM, 0)
    search_to = find_offset(data, TO, 0)
    if verbose:
        print(f"offset {OFFSET}: content={bytes(data[OFFSET:OFFSET+8]).decode('ascii', 'replace')!r} is_disabled={at_off}")
        print(f"full-scan: first 'disabled' at {search}, first 'onDemand' at {search_to}")
    return at_off, search

def apply_patch(path):
    data = load(path)
    # 优先定点
    if data[OFFSET:OFFSET+len(FROM)] == FROM:
        data[OFFSET:OFFSET+len(FROM)] = TO
        save(path, data)
        print(f"OK: 定点补丁 @ {OFFSET}")
        return
    # 否则全文件搜索
    idx = data.find(FROM)
    if idx >= 0:
        data[idx:idx+len(FROM)] = TO
        save(path, data)
        print(f"OK: 全文件搜索补丁 @ {idx}（注意：实际偏移 ≠ 预期 {OFFSET}，请记录）")
        return
    print("FAIL: 未找到 'disabled'（可能已打过补丁或版本不同）")
    sys.exit(1)

def rollback(path):
    bak = path + ".bak"
    if not os.path.exists(bak):
        print("FAIL: 无备份文件", bak)
        sys.exit(1)
    shutil.copyfile(bak, path)
    print(f"OK: 已回滚 {bak}")

def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    path, op = sys.argv[1], sys.argv[2]
    if not os.path.exists(path):
        print("FAIL: 文件不存在", path)
        sys.exit(1)
    if op == "apply":
        if not os.path.exists(path + ".bak"):
            shutil.copyfile(path, path + ".bak")
            print(f"已备份: {path}.bak")
        apply_patch(path)
        verify(path)
    elif op == "rollback":
        rollback(path)
    elif op == "verify":
        verify(path)
    else:
        print("未知操作:", op)
        sys.exit(1)

if __name__ == "__main__":
    main()
