# __init__.py
"""
Excel行颜色同步工具包
快速生成带VBA一键同步按钮的xlsm文件
"""
from .excel_maker import ExcelSyncColorMaker
from .vba_config import FULL_VBA_CODE, BTN_CONFIG, COLOR_RULES

__version__ = "1.0.0"
__all__ = ["ExcelSyncColorMaker", "FULL_VBA_CODE", "BTN_CONFIG", "COLOR_RULES"]