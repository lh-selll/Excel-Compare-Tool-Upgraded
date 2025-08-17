import os
import sys
import shutil
import logging
import argparse
import subprocess
from typing import List, Optional, Union
from pathlib import Path


class FileHandler:
    """文件操作工具类，封装常见的文件读写、管理功能"""
    
    def __init__(self, verbose=False):
        """初始化Excel文件打开器"""
        self.verbose = verbose
        self.logger = self._setup_logger()
        self.excel_valid_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv']
        self.text_valid_extensions = (
            '.txt', '.md', '.csv', '.log', '.py', 
            '.html', '.css', '.js', '.json', '.xml'
        )
        
    def _setup_logger(self):
        """配置日志记录器"""
        logger = logging.getLogger('ExcelFileOpener')
        logger.setLevel(logging.INFO if self.verbose else logging.WARNING)
        
        # 创建控制台处理器
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        
        # 创建日志格式
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        
        # 添加处理器到logger
        if not logger.handlers:
            logger.addHandler(ch)
        
        return logger
    
    def validate_file_path(self, file_path, format):
        """验证文件路径是否有效"""
        path = Path(file_path)
        
        # 检查路径是否存在
        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        # 检查是否为文件
        if not path.is_file():
            raise ValueError(f"路径不是文件: {file_path}")
        
        # 检查文件扩展名是否为Excel格式
        if format == "text":
            if path.suffix.lower() not in self.text_valid_extensions:
                raise ValueError(f"文件不是Text格式: {path.suffix}")
            
        elif format == "excel":
            if path.suffix.lower() not in self.excel_valid_extensions:
                raise ValueError(f"文件不是Excel格式: {path.suffix}")
        
        return path
    
    @staticmethod
    def create_text_file(file_path: str, content: str = "", encoding: str = "utf-8") -> bool:
        """
        创建新文件（若已存在则覆盖）
        :param file_path: 文件路径
        :param content: 初始内容（默认空）
        :param encoding: 编码格式
        :return: 操作是否成功
        """
        try:
            # 确保父目录存在
            dir_path = os.path.dirname(file_path)
            if dir_path and not os.path.exists(dir_path):
                os.makedirs(dir_path, exist_ok=True)
            
            with open(file_path, "w", encoding=encoding) as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"创建文件失败: {e}")
            return False

    @staticmethod
    def read_text_file(file_path: str, encoding: str = "utf-8", is_binary: bool = False) -> Union[str, bytes, None]:
        """
        读取文件内容
        :param file_path: 文件路径
        :param encoding: 编码格式（文本文件用）
        :param is_binary: 是否为二进制文件
        :return: 文件内容（文本返回str，二进制返回bytes，失败返回None）
        """
        try:
            mode = "rb" if is_binary else "r"
            with open(file_path, mode, encoding=encoding if not is_binary else None) as f:
                return f.read()
        except Exception as e:
            print(f"读取文件失败: {e}")
            return None

    @staticmethod
    def append_text_content(file_path: str, content: str, encoding: str = "utf-8") -> bool:
        """
        向文件追加内容（文本文件）
        :param file_path: 文件路径
        :param content: 追加的内容
        :param encoding: 编码格式
        :return: 操作是否成功
        """
        try:
            with open(file_path, "a", encoding=encoding) as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"追加内容失败: {e}")
            return False

    @staticmethod
    def modify_text_line(file_path: str, line_num: int, new_content: str, encoding: str = "utf-8") -> bool:
        """
        修改文件指定行（从1开始计数）
        :param file_path: 文件路径
        :param line_num: 行号（1-based）
        :param new_content: 新内容
        :param encoding: 编码格式
        :return: 操作是否成功
        """
        if line_num < 1:
            print("行号必须大于等于1")
            return False
        
        content = FileHandler.read_text_file(file_path, encoding)
        if content is None:
            return False
        
        lines = content.splitlines()
        if line_num > len(lines):
            print(f"行号超出范围（最大行号: {len(lines)}）")
            return False
        
        # 保持原换行符风格（默认添加\n）
        lines[line_num - 1] = new_content
        return FileHandler.create_text_file(file_path, "\n".join(lines), encoding)

    @staticmethod
    def delete_file(file_path: str) -> bool:
        """
        删除文件
        :param file_path: 文件路径
        :return: 操作是否成功
        """
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return False
        
        try:
            os.remove(file_path)
            return True
        except Exception as e:
            print(f"删除文件失败: {e}")
            return False

    @staticmethod
    def copy_file(src_path: str, dest_path: str, overwrite: bool = False) -> bool:
        """
        复制文件
        :param src_path: 源文件路径
        :param dest_path: 目标路径
        :param overwrite: 是否覆盖已存在文件
        :return: 操作是否成功
        """
        if not os.path.exists(src_path):
            print(f"源文件不存在: {src_path}")
            return False
        
        if os.path.exists(dest_path) and not overwrite:
            print(f"目标文件已存在: {dest_path}（未覆盖）")
            return False
        
        try:
            # 确保目标目录存在
            dest_dir = os.path.dirname(dest_path)
            if dest_dir and not os.path.exists(dest_dir):
                os.makedirs(dest_dir, exist_ok=True)
            
            shutil.copy2(src_path, dest_path)  # 保留元数据
            return True
        except Exception as e:
            print(f"复制文件失败: {e}")
            return False

    @staticmethod
    def move_file(src_path: str, dest_path: str, overwrite: bool = False) -> bool:
        """
        移动文件
        :param src_path: 源文件路径
        :param dest_path: 目标路径
        :param overwrite: 是否覆盖已存在文件
        :return: 操作是否成功
        """
        if not os.path.exists(src_path):
            print(f"源文件不存在: {src_path}")
            return False
        
        if os.path.exists(dest_path):
            if overwrite:
                os.remove(dest_path)
            else:
                print(f"目标文件已存在: {dest_path}（未移动）")
                return False
        
        try:
            # 确保目标目录存在
            dest_dir = os.path.dirname(dest_path)
            if dest_dir and not os.path.exists(dest_dir):
                os.makedirs(dest_dir, exist_ok=True)
            
            shutil.move(src_path, dest_path)
            return True
        except Exception as e:
            print(f"移动文件失败: {e}")
            return False

    @staticmethod
    def get_file_info(file_path: str) -> Optional[dict]:
        """
        获取文件基本信息
        :param file_path: 文件路径
        :return: 包含文件信息的字典（不存在返回None）
        """
        if not os.path.exists(file_path):
            return None
        
        try:
            stat = os.stat(file_path)
            return {
                "size": stat.st_size,  # 大小（字节）
                "create_time": stat.st_ctime,  # 创建时间（时间戳）
                "modify_time": stat.st_mtime,  # 修改时间（时间戳）
                "is_file": os.path.isfile(file_path)
            }
        except Exception as e:
            print(f"获取文件信息失败: {e}")
            return None

    def _detect_os(self) -> str:
        """检测        检测当前操作系统
        返回:
            str: 操作系统类型（'windows'、'macos' 或 'linux'）
        """
        if sys.platform.startswith('win32'):
            return 'windows'
        elif sys.platform.startswith('darwin'):
            return 'macos'
        elif sys.platform.startswith(('linux', 'freebsd', 'openbsd')):
            return 'linux'
        else:
            return 'unknown'
    
    def open_text_file(self, file_path: str) -> bool:
        """
        使用系统默认应用打开文件
        参数:
            file_path: 文件路径
            check_text: 是否仅允许打开文本文件
        返回:
            bool: 操作是否成功
        """
        # 初始化文件打开器，自动识别当前操作系统
        self.os_type = self._detect_os()
        

        try:
            self.validate_file_path(file_path, "text")
            # 根据不同操作系统执行打开命令
            if self.os_type == 'windows':
                os.startfile(file_path)  # Windows 特有方法
            elif self.os_type == 'macos':
                subprocess.run(['open', file_path], check=True)
            elif self.os_type == 'linux':
                subprocess.run(['xdg-open', file_path], check=True)
            else:
                print(f"错误：不支持的操作系统 - {sys.platform}")
                error = f"错误：不支持的操作系统 - {sys.platform}"
                return False, error


            print(f"成功打开Text文件: {file_path}")
            error = f"成功打开Text文件: {file_path}"
            return True, error


        except Exception as e:
            print(f"打开文件失败：{str(e)}")
            error = f"打开文件失败：{str(e)}"
            return False, error
        
    def open_excel_file(self, file_path):
        """使用默认Excel应用打开文件"""
        try:
            # 验证文件路径
            file_path = self.validate_file_path(file_path, "excel")
            
            # 检查操作系统
            if not sys.platform.startswith('win'):
                raise OSError("此功能仅适用于Windows系统")
                
            # 规范化文件路径
            normalized_path = os.path.normpath(file_path)
            
            # 尝试打开文件
            os.startfile(normalized_path)
            error = f"成功打开Excel文件: {file_path}"
            return True, error
            
        except (FileNotFoundError, ValueError, OSError) as e:
            self.logger.error(f"文件打开发生错误: {e}")
            error = f"文件打开发生错误: {e}"
            return False, error
        except Exception as e:
            self.logger.error(f"文件打开时发生未知错误: {e}")
            error = f"文件打开时发生未知错误: {e}"
            return False, error
