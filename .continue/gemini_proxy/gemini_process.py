#!/usr/bin/env python3
"""
Gemini Process Manager
管理 gemini-cli 进程的启动、通信和生命周期
"""

import os
import json
import re
import time
import logging
import subprocess
import threading
import queue
from typing import Generator

# 配置日志 - 写入到 console.log 文件
log_dir = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(log_dir, 'console.log')

# 创建日志格式
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# 配置文件处理器
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)

# 配置控制台处理器
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)

# 配置根日志器
logging.basicConfig(
    level=logging.INFO,
    handlers=[file_handler, console_handler],
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

class GeminiProcess:
    def __init__(self, provider: str = "deepseek", model: str = "deepseek-chat", api_key: str = None):
        self.provider = provider
        self.model = model
        self.api_key = api_key
        self.process = None
        self.running = False
        self.output_thread = None
        self.output_queue = queue.Queue()
        
        # 启动进程
        self.start_process()

    def start_process(self):
        """启动 gemini 进程"""
        try:
            # 设置环境变量
            env = os.environ.copy()
            if self.api_key:
                env['GEMINI_API_KEY'] = self.api_key
            
            # 设置 provider 环境变量
            env['GEMINI_PROVIDER'] = self.provider
            
            # 构建命令行参数
            cmd = [
                'gemini',
                '--provider=' + self.provider,
                '--model=' + self.model,
                '-y',
                '--plain'
            ]
            
            logger.info(f"Starting gemini process with provider={self.provider}, model={self.model}")
            
            # 使用 subprocess.Popen 替代 pexpect
            self.process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
                universal_newlines=True,
                env=env
            )
            
            # 启动输出监听线程
            self.running = True
            self.output_thread = threading.Thread(target=self._monitor_output)
            self.output_thread.daemon = True
            self.output_thread.start()
            
            # 等待初始化
            time.sleep(3)
            logger.info(f"Started gemini process in interactive JSON mode with provider={self.provider}, model={self.model}")
            
        except Exception as e:
            logger.error(f"Failed to start gemini process: {e}")
            self.process = None
            raise Exception(f"Failed to start gemini process: {e}")

    def _monitor_output(self):
        """监听进程输出的线程"""
        while self.running and self.process:
            try:
                line = self.process.stdout.readline()
                if not line:
                    break
                    
                if line:
                    # 过滤掉初始化和调试信息
                    if not line.startswith('Data collection is disabled'):
                        self.output_queue.put(line)
                        logger.debug(f"Output: {line}")
            except Exception as e:
                logger.error(f"Output monitoring error: {e}")
                break
        
        logger.info("Output monitoring thread stopped")

    def is_process_alive(self):
        return self.process is not None and self.process.poll() is None

    def send_prompt_stream(self, prompt: str) -> Generator[str, None, None]:
        """流式发送 prompt 并返回生成器（模拟流式输出）"""
        if not self.is_process_alive():
            logger.warning("Gemini process is not alive, restarting...")
            self.restart()
            if not self.is_process_alive():
                raise Exception("Gemini process could not be restarted.")
        
        try:
            logger.info(f"Sending prompt to gemini ({self.provider}/{self.model}): {prompt[:50]}...")
            
            # 发送 prompt
            self.process.stdin.write(prompt + '\n')
            self.process.stdin.flush()
            
            # 等待完整响应
            response_content = ""
            start_time = time.time()
            timeout = 120  # 2分钟超时
            
            while time.time() - start_time < timeout:
              # 非阻塞方式获取输出
              try:
                  line = self.output_queue.get_nowait()
                  line = re.sub(r'🤖 Output:s?', '', line)
                  if not line.strip():
                      continue
                  elif line.strip() == '👤 Input:':
                      return
                  response_content += line
                  yield line
              except queue.Empty:
                  time.sleep(0.1)
                  continue
            
            # 超时处理
            if not response_content:
                yield "Sorry, the response timed out. Please try again."
                
        except Exception as e:
            logger.error(f"Error in send_prompt_stream: {e}")
            yield f"Error: {str(e)}"

    def send_prompt(self, prompt: str) -> str:
        """发送 prompt 并返回完整响应"""
        response_parts = []
        for chunk in self.send_prompt_stream(prompt):
            response_parts.append(chunk)
        return ''.join(response_parts)

    def restart(self):
        """重启进程"""
        logger.info("Restarting gemini process...")
        self.running = False
        if self.process:
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
            except:
                self.process.kill()
        
        self.process = None
        self.start_process()

    def update_config(self, provider: str = None, model: str = None, api_key: str = None):
        """更新配置并重启进程"""
        if provider:
            self.provider = provider
        if model:
            self.model = model
        if api_key:
            self.api_key = api_key
        
        logger.info(f"Updating config: provider={self.provider}, model={self.model}")
        self.restart()
