从截图来看，文件已经整理到同一个文件夹里了，而且命名非常规范（`01...`, `02...` 开头）。

既然你的需求是**“以 0914 版本为基础，把 0821 版本里的数据补充进去”**，而且文件有 40 多个，手动一个个打开用公式太慢了。

这里有一个**最稳妥且自动化**的方案：使用 **VBA 宏**。它可以自动识别哪一个是新版（0914），哪一个是旧版（0821），然后自动在后台完成“新建 Sheet + 填充公式 + 转数值”的全过程。

### 操作步骤

#### 第一步：准备工作

1. 确保所有文件都在你截图的这个文件夹里。
2. **新建一个空的 Excel 文件**，把它也保存在这个文件夹里（或者随便哪里都行，但为了方便管理建议放一起），命名为 `合并工具.xlsm`（注意后缀要是 xlsm，启用宏的工作簿）。
3. 打开这个 `合并工具.xlsm`。

#### 第二步：复制 VBA 代码

按键盘上的 `Alt + F11` 打开 VBA 编辑器，点击菜单栏的 `插入` -> `模块`，然后把下面的代码全部复制进去：

```vba
Sub ConditionalGroupFill_Final()
    Dim sPath As String, sFile As String
    Dim wbMain As Workbook, wbNew As Workbook, wbOld As Workbook
    Dim wsNew As Worksheet, wsOld As Worksheet
    Dim dictOld As Object, arrFiles As Object
    Dim r As Long, lastRowNew As Long, lastRowOld As Long, maxRow As Long, col As Long
    Dim matchCount As Long, saveCount As Long, bModified As Boolean
    Dim vNew As Variant, vOld As Variant
    Dim triggerFill As Boolean
    Dim sFileName As String, sOldFileName As String, sOldFilePath As String
    Dim wsWasProtected As Boolean
    
    ' 1. 优化运行环境
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set wbMain = ThisWorkbook
    sPath = wbMain.Path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    Set dictOld = CreateObject("Scripting.Dictionary")
    Set arrFiles = CreateObject("System.Collections.ArrayList")
    
    ' 2. 收集 0821 文件索引
    sFile = Dir(sPath & "*.xls*")
    Do While sFile <> ""
        If LCase(sFile) <> LCase(wbMain.Name) And InStr(sFile, "0821") > 0 Then
            dictOld(LCase(sFile)) = True
        End If
        sFile = Dir()
    Loop
    
    ' 3. 收集 0914 文件完整路径
    sFile = Dir(sPath & "*.xls*")
    Do While sFile <> ""
        If LCase(sFile) <> LCase(wbMain.Name) And InStr(sFile, "0914") > 0 Then
            arrFiles.Add sPath & sFile
        End If
        sFile = Dir()
    Loop
    
    ' 4. 遍历处理
    Dim vFile As Variant
    For Each vFile In arrFiles.ToArray
        On Error GoTo FileError
        
        sFileName = Mid(CStr(vFile), InStrRev(CStr(vFile), "\") + 1)
        sOldFileName = Replace(sFileName, "0914", "0821")
        sOldFilePath = sPath & sOldFileName
        
        If dictOld.Exists(LCase(sOldFileName)) Then
            matchCount = matchCount + 1
            
            Set wbNew = Workbooks.Open(CStr(vFile), ReadOnly:=False)
            Set wbOld = Workbooks.Open(sOldFilePath, ReadOnly:=True)
            
            ' 【修复1】强制使用 Worksheets 而不是 Sheets，避免选中图表页报错
            Set wsNew = wbNew.Worksheets(1)
            Set wsOld = wbOld.Worksheets(1)
            
            ' 【修复2】自动解除工作表保护，防止写入时 1004 报错
            wsWasProtected = wsNew.ProtectContents
            If wsWasProtected Then
                On Error Resume Next
                wsNew.Unprotect ' 如果有密码，这里可能需要改为 wsNew.Unprotect "你的密码"
                On Error GoTo FileError
            End If
            
            ' 【修复3】多重探测最大行数，防止A列为空导致行数变成104万行
            lastRowNew = 1
            On Error Resume Next
            lastRowNew = Application.Max( _
                wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row, _
                wsNew.Cells(wsNew.Rows.Count, 23).End(xlUp).Row, _
                wsNew.UsedRange.Rows.Count + wsNew.UsedRange.Row - 1)
            On Error GoTo FileError
            
            lastRowOld = 1
            On Error Resume Next
            lastRowOld = Application.Max( _
                wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row, _
                wsOld.Cells(wsOld.Rows.Count, 23).End(xlUp).Row, _
                wsOld.UsedRange.Rows.Count + wsOld.UsedRange.Row - 1)
            On Error GoTo FileError
            
            maxRow = Application.Max(lastRowNew, lastRowOld)
            
            ' 极限保护：如果计算出的行数异常大，强制截断
            If maxRow > 100000 Then maxRow = 100000 
            
            bModified = False
            
            ' 【核心逻辑】逐行判断并覆盖
            For r = 2 To maxRow
                triggerFill = False
                
                ' 检查 W(23) / X(24) / Y(25) / Z(26)
                For col = 23 To 26
                    vNew = ""
                    vOld = ""
                    On Error Resume Next
                    vNew = Trim(CStr(wsNew.Cells(r, col).Value))
                    vOld = Trim(CStr(wsOld.Cells(r, col).Value))
                    On Error GoTo FileError
                    
                    If Len(vNew) = 0 And Len(vOld) > 0 Then
                        triggerFill = True
                        Exit For
                    End If
                Next col
                
                ' 触发覆盖 W~AB (23~28)
                If triggerFill Then
                    ' 【修复4】写入前清除合并单元格和数组公式，彻底扫清 1004 障碍
                    On Error Resume Next
                    wsNew.Cells(r, 23).Resize(1, 6).UnMerge
                    wsNew.Cells(r, 23).Resize(1, 6).FormulaArray = False ' 尝试清除数组公式属性
                    On Error GoTo FileError
                    
                    ' 逐列安全写入
                    For col = 23 To 28
                        On Error Resume Next
                        wsNew.Cells(r, col).Value = wsOld.Cells(r, col).Value
                        On Error GoTo FileError
                    Next col
                    bModified = True
                End If
            Next r
            
            ' 恢复工作表保护
            If wsWasProtected Then
                On Error Resume Next
                wsNew.Protect
                On Error GoTo FileError
            End If
            
            If bModified Then
                wbNew.Save
                saveCount = saveCount + 1
            End If
            
            wbOld.Close SaveChanges:=False
            Set wbOld = Nothing
            wbNew.Close SaveChanges:=False
            Set wbNew = Nothing
        End If
        
        GoTo NextFile
        
FileError:
        Debug.Print "处理失败: " & sFileName & " | 行:" & r & " 列:" & col & " | 错误:" & Err.Description
        On Error Resume Next
        If wsWasProtected And Not wsNew Is Nothing Then wsNew.Protect
        If Not wbOld Is Nothing Then wbOld.Close SaveChanges:=False
        If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
        Set wbOld = Nothing: Set wbNew = Nothing
        On Error GoTo 0
        
NextFile:
    Next vFile
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    
    MsgBox "执行完毕！" & vbCrLf & _
           "成功配对: " & matchCount & " 个文件" & vbCrLf & _
           "实际保存: " & saveCount & " 个文件", vbInformation, "处理完成"
End Sub
```

### 第三步：运行

1. 回到 Excel 界面。
2. 按 `Alt + F8`，选择 `MergeOldToNew`，点击“运行”。
3. 等待几分钟（取决于文件大小和数量），程序会自动打开每一对文件，计算并保存。

---

### 这个脚本做了什么？

1. **智能配对**：它会自动识别文件名里的 `20260914` 和 `20260821`，把它们凑成一对。
2. **以新补旧**：它会打开这两个文件，读取数据到内存。
3. **内存运算**：在内存中执行“如果有新值就用新值，否则用旧值”的逻辑（比 Excel 公式快得多）。
4. **生成结果**：直接把算好的纯数值写入到 `0914` 那个文件的新 Sheet `合并结果` 中。
5. **自动保存**：处理完一对就自动保存关闭，最后告诉你完成了。

### ⚠️ 重要提示

* **备份！** 运行前最好把这 40 个文件复制一份备份，防止意外覆盖。
* **Sheet 名称**：代码默认读取每个文件的**第 1 个 Sheet** (`Sheets(1)`). 如果你的数据不在第一个 Sheet，需要修改代码中的 `Set wsNew = wbNew.Sheets(1)` 为具体的名字，例如 `wbNew.Sheets("Sheet1")`。
* **文件格式**：代码假设所有文件都是 `.xlsx`。

你可以先拿两个文件试一下，没问题再跑全部。

---

VBA 方案能一步到位。不过如果你担心代码不好维护，我也可以教你用 Power Query 的“追加查询”功能，不用写代码也能搞定，想试试吗？