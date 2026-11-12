# -*- coding: utf-8 -*-
"""只读抽取任务池。守护线程不写数据库，不阻挡程序退出。

取消只能阻止待执行任务；已运行的第三方解析不能被Python硬杀。
关闭任务池后不再接收新任务；协调线程丢弃迟到结果。
"""
import queue
import threading
from concurrent.futures import Future


class ExtractionPool:
    def __init__(self, max_workers):
        self._queue = queue.Queue()
        self._closed = threading.Event()
        self._lock = threading.Lock()
        self._threads = []
        for i in range(max_workers):
            thread = threading.Thread(target=self._worker, name=f'DocExtract-{i}', daemon=True)
            thread.start()
            self._threads.append(thread)

    def submit(self, function, *args):
        with self._lock:
            if self._closed.is_set():
                raise RuntimeError('解析任务池已经关闭')
            future = Future()
            self._queue.put((future, function, args))
            return future

    def _worker(self):
        while not self._closed.is_set():
            try:
                future, function, args = self._queue.get(timeout=0.2)
            except queue.Empty:
                continue
            try:
                if self._closed.is_set():
                    future.cancel()
                elif future.set_running_or_notify_cancel():
                    try:
                        value = function(*args)
                    except BaseException as error:
                        future.set_exception(error)
                    else:
                        future.set_result(value)
            finally:
                self._queue.task_done()

    def shutdown(self, wait=False, cancel_futures=True):
        with self._lock:
            self._closed.set()
            if cancel_futures:
                while True:
                    try:
                        future, function, args = self._queue.get_nowait()
                    except queue.Empty:
                        break
                    future.cancel()
                    self._queue.task_done()
        if wait:
            for thread in self._threads:
                thread.join()
