import os
import time
import logging
import sys
import fcntl
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import pandas as pd
from typing import Dict, Optional, List, Set
import pyinotify  # For Linux file monitoring
import threading
from queue import Queue, Empty

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('excel_monitor.log'),
        logging.StreamHandler()
    ]
)

class ExcelMonitor:
    def __init__(self, file_paths: List[str] = None):
        """
        初始化Excel监控器
        :param file_paths: 要监控的Excel文件路径列表
        """
        self.file_paths = [os.path.abspath(path) for path in file_paths] if file_paths else []
        self.watch_manager = pyinotify.WatchManager()
        self.observer = None
        self.handled_files = set()
        self.lock = threading.Lock()
        self.event_queue = Queue()
        self.running = True
        self.TRIGGER_TEXT = "触发"
        self.GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
    def is_file_open(self, filepath: str) -> bool:
        """检查文件是否被其他进程打开"""
        if not os.path.exists(filepath):
            return False
            
        try:
            # 尝试以独占模式打开文件
            with open(filepath, 'a+') as f:
                fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)
                return False
        except (IOError, BlockingIOError):
            return True
        except Exception as e:
            logging.error(f"检查文件 {filepath} 状态时出错: {str(e)}")
            return False

    def start_monitoring(self):
        """开始监控文件变化"""
        # 为每个文件添加inotify监控
        for file_path in self.file_paths:
            if not os.path.exists(file_path):
                logging.warning(f"文件不存在，将被创建: {file_path}")
                try:
                    open(file_path, 'a').close()
                except Exception as e:
                    logging.error(f"创建文件 {file_path} 失败: {str(e)}")
                    continue
            
            # 添加文件监控
            self.watch_manager.add_watch(
                file_path, 
                pyinotify.IN_MODIFY | pyinotify.IN_CLOSE_WRITE,
                rec=False
            )
            logging.info(f"已添加监控: {file_path}")
        
        # 启动事件处理线程
        event_handler = ExcelFileHandler(self.event_queue)
        self.observer = pyinotify.ThreadedNotifier(self.watch_manager, event_handler)
        self.observer.start()
        
        # 启动处理线程
        process_thread = threading.Thread(target=self._process_events, daemon=True)
        process_thread.start()
        
        logging.info("监控已启动，等待文件被打开...")
        
        try:
            while self.running:
                # 检查文件是否被打开
                for file_path in self.file_paths:
                    if self.is_file_open(file_path):
                        logging.info(f"检测到文件被打开: {file_path}")
                        # 立即处理一次文件
                        self.process_excel(file_path)
                
                time.sleep(2)  # 每2秒检查一次文件状态
                
        except KeyboardInterrupt:
            logging.info("正在停止监控...")
            self.stop()
    
    def _process_events(self):
        """处理文件变化事件"""
        while self.running:
            try:
                file_path = self.event_queue.get(timeout=1)
                if file_path and os.path.exists(file_path):
                    self.process_excel(file_path)
            except Empty:
                continue
            except Exception as e:
                logging.error(f"处理事件时出错: {str(e)}")
    
    def stop(self):
        """停止监控"""
        self.running = False
        if self.observer:
            self.observer.stop()
        logging.info("监控已停止")
    
    def process_excel(self, file_path: str):
        """处理Excel文件"""
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            return
            
        # 检查文件是否被其他进程打开
        if not self.is_file_open(file_path):
            return
            
        # 使用锁确保同一时间只有一个线程处理文件
        with self.lock:
            try:
                logging.info(f"正在处理文件: {file_path}")
                
                # 创建备份
                backup_path = f"{file_path}.backup_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                
                # 使用临时文件进行修改
                temp_path = f"{file_path}.tmp"
                
                # 复制原文件到临时文件
                import shutil
                shutil.copy2(file_path, temp_path)
                
                # 处理临时文件
                wb = openpyxl.load_workbook(temp_path)
                modified = False
                
                for ws_name in wb.sheetnames:
                    ws = wb[ws_name]
                    
                    # 检查每个单元格是否包含触发文本
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value and self.TRIGGER_TEXT in str(cell.value):
                                # 高亮显示触发单元格
                                cell.fill = self.GREEN_FILL
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                modified = True
                                logging.info(f"在 {file_path} 的 {ws_name} 工作表中找到触发文本: {cell.coordinate}")
                
                # 如果有修改，则保存文件
                if modified:
                    # 创建备份
                    shutil.copy2(file_path, backup_path)
                    # 保存修改到原文件
                    wb.save(file_path)
                    logging.info(f"文件已更新: {file_path}")
                
                # 关闭工作簿
                wb.close()
                
                # 删除临时文件
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    
            except Exception as e:
                logging.error(f"处理文件 {file_path} 时出错: {str(e)}", exc_info=True)
                # 如果出错，确保临时文件被删除
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            
            # 如果有修改，则保存文件
            if modified:
                backup_path = f"{file_path}.backup_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                os.rename(file_path, backup_path)
                wb.save(file_path)
                logging.info(f"文件已处理并保存: {file_path}")
                
        except Exception as e:
            logging.error(f"处理文件 {file_path} 时出错: {str(e)}", exc_info=True)
    
class ExcelFileHandler(pyinotify.ProcessEvent):
    """处理文件变化事件"""
    def __init__(self, event_queue):
        super().__init__()
        self.event_queue = event_queue
    
    def process_IN_MODIFY(self, event):
        """处理文件修改事件"""
        if not event.dir:
            self.event_queue.put(event.pathname)
    
    def process_IN_CLOSE_WRITE(self, event):
        """处理文件关闭写入事件"""
        if not event.dir:
            self.event_queue.put(event.pathname)

def get_file_paths() -> List[str]:
    """获取用户输入的要监控的文件路径"""
    print("请输入要监控的Excel文件路径（每行一个，空行结束）:")
    file_paths = []
    
    while True:
        try:
            path = input().strip()
            if not path:  # 空行结束输入
                break
                
            path = os.path.expanduser(path)  # 处理 ~ 符号
            path = os.path.abspath(path)     # 转换为绝对路径
            
            # 检查文件是否存在
            if not os.path.exists(path):
                create = input(f"文件 {path} 不存在，是否创建? (y/n): ").strip().lower()
                if create == 'y':
                    try:
                        open(path, 'a').close()
                        print(f"已创建文件: {path}")
                    except Exception as e:
                        print(f"创建文件失败: {str(e)}")
                        continue
                else:
                    continue
            
            if path not in file_paths:
                file_paths.append(path)
                print(f"已添加: {path}")
            else:
                print(f"文件已添加: {path}")
                
        except KeyboardInterrupt:
            print("\n输入结束")
            break
        except Exception as e:
            print(f"输入错误: {str(e)}")
    
    return file_paths

if __name__ == "__main__":
    # 安装依赖: pip install openpyxl pandas pyinotify
    print("=== Excel文件监控器 ===")
    print("功能: 监控Excel文件，当单元格包含'触发'时自动高亮")
    print("按 Ctrl+C 停止监控\n")
    
    # 获取要监控的文件路径
    # file_paths = get_file_paths()
    file_paths = ["/home/niejie/work/Code/temp/test.xlsx"]
    
    if not file_paths:
        print("未指定要监控的文件，程序退出")
        sys.exit(0)
    
    print("\n开始监控以下文件:")
    for path in file_paths:
        print(f"- {path}")
    print("\n等待文件被打开...")
    
    try:
        # 启动监控
        monitor = ExcelMonitor(file_paths=file_paths)
        monitor.start_monitoring()
    except KeyboardInterrupt:
        print("\n正在停止监控...")
        monitor.stop()
    except Exception as e:
        logging.error(f"程序出错: {str(e)}", exc_info=True)
    finally:
        print("程序已退出")