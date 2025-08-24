from PySide6.QtCore import QObject, QThread, Signal, Slot, Qt, QTimer
from typing import Dict, Any, Tuple
import logging
import os
from collections import deque
import time


class BackgroundLogWorker(QObject):
    log_request = Signal(list)  # 接收日志列表：[(level, msg, args, kwargs), ...]

    def __init__(self, log_file_path: str):
        super().__init__()
        os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
        self.logger = self._setup_logger(log_file_path)
        self.log_request.connect(self._process_log_request, Qt.QueuedConnection)

    def _setup_logger(self, log_file: str) -> logging.Logger:
        logger = logging.getLogger("OptimizedLogger")
        logger.setLevel(logging.DEBUG)
        logger.propagate = False
        if not logger.handlers:
            handler = logging.FileHandler(log_file, encoding="utf-8")
            formatter = logging.Formatter(
                "%(asctime)s - %(levelname)s - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger

    @Slot(list)
    def _process_log_request(self, log_list: list):
        """在后台线程中完成日志格式化（耗时操作移到此处）"""
        for level, msg, args, kwargs in log_list:
            try:
                # 后台线程中执行字符串格式化（不阻塞主线程）
                formatted_msg = msg.format(*args, **kwargs) if args or kwargs else msg
                log_method = getattr(self.logger, level, None)
                if log_method and callable(log_method):
                    log_method(formatted_msg)
                    print(f"formatted_msg = {formatted_msg}")
            except Exception as e:
                print(f"日志处理失败：{e}")


class BackgroundLogManager(QObject):
    def __init__(self, log_file_path: str):
        super().__init__()
        self.log_buffer = deque()  # 轻量缓冲区
        self.log_thread = QThread()
        self.log_worker = BackgroundLogWorker(log_file_path)
        self.log_worker.moveToThread(self.log_thread)
        self.log_thread.start()

        # 重复日志去重缓存
        self._log_cache = {}  # {msg_template: (last_time, count)}
        self._cache_clean_timer = QTimer(self)
        self._cache_clean_timer.setInterval(50)  # 每5秒清理一次过期缓存
        self._cache_clean_timer.timeout.connect(self._clean_log_cache)
        self._cache_clean_timer.start()

    def _clean_log_cache(self):
        """清理5秒以上的缓存，避免内存占用过大"""
        current_time = time.time()
        to_remove = [k for k, (t, _) in self._log_cache.items() if current_time - t > 5]
        for k in to_remove:
            del self._log_cache[k]

    def _send_log_request(self, level: str, msg: str, *args, **kwargs):
        if not self.log_thread.isRunning():
            return

        # 1. 过滤重复日志（1秒内相同模板的日志合并）
        # current_time = time.time()
        # if msg in self._log_cache:
        #     last_time, count = self._log_cache[msg]
        #     if current_time - last_time < 1:
        #         self._log_cache[msg] = (last_time, count + 1)
        #         return
        #     else:
        #         # 补充重复次数
        #         msg += f"（重复{count + 1}次）"
        
        # self._log_cache[msg] = (current_time, 1)

        # 2. 限制消息长度，避免大字符串
        # if len(msg) > 500:
        #     msg = msg[:500] + "..."

        # 3. 只传递原始参数，不在主线程格式化
        self.log_buffer.append( (level, msg, args, kwargs) )

        # 4. 缓冲区满时主动刷新（避免堆积）
        if len(self.log_buffer) >= 50:
            self._flush_log_buffer()

    def _flush_log_buffer(self):
        if not self.log_buffer:
            return
        self.log_worker.log_request.emit(list(self.log_buffer))
        print(f"def _flush_log_buffer(self): = {self.log_buffer}")
        self.log_buffer.clear()

    # 对外接口（保持不变）
    def debug(self, msg: str, *args, **kwargs):
        self._send_log_request("debug", msg, *args, **kwargs)

    def info(self, msg: str, *args, **kwargs):
        self._send_log_request("info", msg, *args, **kwargs)
        # print(f"info(self, msg: str, *args, **kwargs): {msg}")

    def warning(self, msg: str, *args, **kwargs):
        self._send_log_request("warning", msg, *args, **kwargs)

    def error(self, msg: str, *args, **kwargs):
        self._send_log_request("error", msg, *args, **kwargs)

    def critical(self, msg: str, *args, **kwargs):
        self._send_log_request("critical", msg, *args, **kwargs)

    def task_finished(self):
        self._flush_log_buffer()
    # ... 其他日志级别方法

    def shutdown(self):
        self._flush_log_buffer()
        if self.log_thread.isRunning():
            self.log_thread.quit()
            self.log_thread.wait(3000)
