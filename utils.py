#!/usr/bin/env python3
"""上线相关的小工具：结果文件清理、简单限流、文件校验。"""

import os
import threading
import time
from collections import defaultdict, deque


def start_cleanup_thread(result_dir, ttl_seconds=1800, interval_seconds=300):
    """后台线程：周期性删除 result_dir 下超过 ttl 的文件，防止磁盘被撑爆。"""
    def _loop():
        while True:
            now = time.time()
            try:
                for name in os.listdir(result_dir):
                    path = os.path.join(result_dir, name)
                    try:
                        if os.path.isfile(path) and now - os.path.getmtime(path) > ttl_seconds:
                            os.remove(path)
                    except OSError:
                        pass
            except OSError:
                pass
            time.sleep(interval_seconds)

    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    return t


class RateLimiter:
    """简单的滑动窗口限流：每个 IP 在 window 秒内最多 max_calls 次。"""

    def __init__(self, max_calls=30, window=60):
        self.max_calls = max_calls
        self.window = window
        self._hits = defaultdict(deque)
        self._lock = threading.Lock()

    def allow(self, key):
        now = time.time()
        with self._lock:
            dq = self._hits[key]
            while dq and now - dq[0] > self.window:
                dq.popleft()
            if len(dq) >= self.max_calls:
                return False
            dq.append(now)
            return True


def ext_ok(filename, allowed):
    """校验文件扩展名是否在白名单内。"""
    if not filename:
        return False
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    return ext in allowed
