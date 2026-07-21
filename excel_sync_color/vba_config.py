# vba_config.py
"""
VBA宏配置文件，所有Excel内置宏统一存放于此，修改无需改动生成逻辑

========================================================================================
执行逻辑：
a（最高优先级）：本行任意单元格 =【差异点】→ A 列标记「差异点」
b：无差异点，整行全部【删除行】→「删除行」
c：无差异点，整行全部【新增】→「新增」
d：无差异点，单元格只有【新增】或【一致】 →「一致」
其余混合场景（删除 + 新增 / 删除 + 一致 / 删除 + 新增 + 一致等）→ A 列清空，无填充色
========================================================================================
"""

# 主同步宏：跳过隐藏列、读取条件格式、A列同步文字底色
SYNC_MAIN_VBA = '''
Sub SyncFirstCellColorAndText()
    Dim ws As Worksheet
    Dim targetRow As Long, lastCol As Long, col As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ws = ActiveSheet
    Const FirstCol As Integer = 1          ' A列输出文字与颜色
    Const ScanStartCol As Integer = 2       ' 从B列开始扫描
    
    ' 颜色映射（与Python配色严格保持一致）
    Dim colorRule As Variant
    colorRule = Array( _
        Array(RGB(175, 255, 175), "一致"), _
        Array(RGB(229, 115, 115), "差异点"), _
        Array(RGB(211, 227, 253), "新增"), _
        Array(RGB(105, 105, 105), "删除行") _
    )
    
    lastCol = ws.UsedRange.Columns.Count
    
    For targetRow = 2 To ws.UsedRange.Rows.Count
        Dim hasDiff As Boolean
        Dim hasDel As Boolean, hasNew As Boolean, hasAgree As Boolean
        
        hasDiff = False
        hasDel = False
        hasNew = False
        hasAgree = False
        
        For col = ScanStartCol To lastCol
            Dim targetCell As Range
            Set targetCell = ws.Cells(targetRow, col)
            
            ' 跳过隐藏列
            If targetCell.EntireColumn.Hidden = True Then
                GoTo NextColLoop
            End If
            
            Dim cellColor As Long
            Dim foundType As String
            foundType = ""
            cellColor = targetCell.DisplayFormat.Interior.Color
            
            ' 匹配单元格颜色对应的类型
            Dim ruleItem As Variant
            For Each ruleItem In colorRule
                If ruleItem(0) = cellColor Then
                    foundType = ruleItem(1)
                    Exit For
                End If
            Next ruleItem
            
            Select Case foundType
                Case "差异点"
                    hasDiff = True
                Case "删除行"
                    hasDel = True
                Case "新增"
                    hasNew = True
                Case "一致"
                    hasAgree = True
            End Select
            
NextColLoop:
        Next col
        
        Dim outputText As String
        Dim outputColor As Long
        outputText = ""
        outputColor = xlColorIndexNone
        
        ' 按优先级判断
        If hasDiff Then
            outputText = "差异点"
            outputColor = RGB(229, 115, 115)
        ElseIf hasDel And Not hasNew And Not hasAgree Then
            ' 全部删除行
            outputText = "删除行"
            outputColor = RGB(105, 105, 105)
        ElseIf hasNew And Not hasDel And Not hasAgree Then
            ' 全部新增
            outputText = "新增"
            outputColor = RGB(211, 227, 253)
        ElseIf Not hasDel And (hasNew Or hasAgree) Then
            ' 不含删除行，只有新增/一致（单独一致、单独新增、新增+一致）
            outputText = "一致"
            outputColor = RGB(175, 255, 175)
        End If
        
        ' 写入首列
        ws.Cells(targetRow, FirstCol).Value = outputText
        If outputText <> "" Then
            ws.Cells(targetRow, FirstCol).Interior.Color = outputColor
        Else
            ws.Cells(targetRow, FirstCol).Interior.Color = xlColorIndexNone
        End If
    Next targetRow
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "同步完成！"
End Sub


'''

# 辅助宏：修改按钮文字（规避win32com操作Shape文本报错）
SET_BTN_TEXT_VBA = '''
Sub SetSyncBtnText()
    Dim shp As Shape
    On Error Resume Next
    Set shp = ActiveSheet.Shapes("Btn_SyncColor")
    If Not shp Is Nothing Then
        shp.TextFrame.Characters.Text = "同步本行颜色与文字"
    End If
End Sub
'''

# 拼接完整VBA模块代码
FULL_VBA_CODE = SYNC_MAIN_VBA + "\n" + SET_BTN_TEXT_VBA

# 按钮基础配置（可统一修改位置、尺寸、宏绑定名）
BTN_CONFIG = {
    "name": "Btn_SyncColor",
    "bind_macro": "SyncFirstCellColorAndText",
    "text_macro": "SetSyncBtnText",
    "left_cell": "A1",
    "width": 140,
    "height": 28
}

# 颜色映射配置（如需新增颜色规则，仅修改此处即可）
COLOR_RULES = [
    {"rgb": (175, 255, 175), "text": "一致"},
    {"rgb": (229, 115, 115), "text": "差异点"},
    {"rgb": (211, 227, 253), "text": "新增"},
    {"rgb": (105, 105, 105), "text": "删除行"},
]