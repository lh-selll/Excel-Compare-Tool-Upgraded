# excel_maker.py
import os
import pythoncom
import win32com.client as win32
from .vba_config import FULL_VBA_CODE, BTN_CONFIG


class ExcelSyncColorMaker:
    def __init__(self):
        self.workbook = None

    def add_macro_to_excel(self, file_path: str, sheet_name_list, save_xlsm_path: str):
        """
        对外统一入口：传入文件路径+工作表名，自动添加宏与按钮
        :param file_path: 原始xlsx文件完整路径
        :param sheet_name: 需要添加按钮的工作表名称
        :param save_xlsm_path: 输出带宏文件路径，后缀必须 .xlsm
        """# 兼容：传入单个字符串 或 list
        if isinstance(sheet_name_list, str):
            sheet_name_list = [sheet_name_list]
            
        # 格式校验
        if not save_xlsm_path.lower().endswith(".xlsm"):
            raise ValueError("输出文件必须为 .xlsm 格式，xlsx无法存储VBA宏")

        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            # 打开原始表格（只读，不破坏原文件）
            # self.workbook = excel.Workbooks.Open(Filename=file_path, ReadOnly=True)
            abs_file_path = os.path.abspath(file_path)
            abs_save_xlsm_path = os.path.abspath(save_xlsm_path)
            
            print(f"转换后的abs_file_path绝对路径：{abs_file_path}")
            print(f"转换后的abs_save_xlsm_path绝对路径：{abs_save_xlsm_path}")
            self.workbook = excel.Workbooks.Open(Filename=abs_file_path, ReadOnly=True)

            # ========== 【修正】VBA模块只处理一次，放到循环外面 ==========
            self._remove_old_vba()
            self._inject_vba()

            # 循环工作表：仅添加按钮，不再重复注入宏
            index = 0
            for sheet_name in sheet_name_list:
                try:
                    target_sheet = self.workbook.Worksheets(sheet_name)
                except Exception:
                    raise ValueError(f"工作表【{sheet_name}】不存在，请检查名称")

                self._remove_old_button(target_sheet)
                self._add_button(target_sheet, index)
                index += 1
                
            # return self.workbook

            # 另存为启用宏文件
            print(f"save_xlsm_path = {abs_save_xlsm_path}")
            self.workbook.SaveAs(Filename=abs_save_xlsm_path, FileFormat=52)
            print(f"✅ 处理完成，文件输出至：{abs_save_xlsm_path}")

        except Exception as err:
            raise RuntimeError(f"添加宏失败：{str(err)}") from err
        finally:
            # 强制释放Excel进程，避免后台残留
            if self.workbook is not None:
                self.workbook.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()

    # ---------------------- 内部私有方法 ----------------------
    def _inject_vba(self) -> None:
        """注入全套VBA同步宏代码"""
        vba_project = self.workbook.VBProject
        code_mod = vba_project.VBComponents.Add(1)
        code_mod.Name = "Module_SyncColor"
        code_mod.CodeModule.AddFromString(FULL_VBA_CODE)

    def _remove_old_vba(self) -> None:
        """删除旧同名VBA模块"""
        try:
            vba_project = self.workbook.VBProject
            for comp in vba_project.VBComponents:
                if comp.Name == "Module_SyncColor":
                    vba_project.VBComponents.Remove(comp)
                    break
        except Exception:
            pass

    def _add_button(self, ws, index):
        """在指定工作表创建表单按钮"""
        left_range = ws.Range(BTN_CONFIG["left_cell"])
        btn_shape = ws.Shapes.AddFormControl(
            Type=0,
            Left=left_range.Left,
            Top=left_range.Top,
            Width=BTN_CONFIG["width"],
            Height=BTN_CONFIG["height"]
        )
        btn_shape.Name = BTN_CONFIG["name"]+str(index)
        print(f"btn_shape.Name = {btn_shape.Name}")
        btn_shape.OnAction = BTN_CONFIG["bind_macro"]
        # 调用VBA修改按钮文字，规避win32com文本属性报错
        # self.workbook.Application.Run(BTN_CONFIG["text_macro"])
        txt = btn_shape.TextFrame.Characters()
        txt.Text = "同步本行颜色与文字"

    def _remove_old_button(self, ws):
        """删除工作表内旧的同步按钮"""
        try:
            for shp in ws.Shapes:
                if shp.Name == BTN_CONFIG["name"]:
                    shp.Delete()
                    break
        except Exception:
            pass