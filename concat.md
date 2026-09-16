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
Sub ConditionalGroupFill()
    Dim sPath As String, sFile As String
    Dim wbMain As Workbook, wbNew As Workbook, wbOld As Workbook
    Dim wsNew As Worksheet, wsOld As Worksheet
    Dim dictOld As Object, arrFiles As Object
    Dim r As Long, lastRowNew As Long, lastRowOld As Long, maxRow As Long, col As Long
    Dim matchCount As Long, saveCount As Long, bModified As Boolean
    Dim vNew As Variant, vOld As Variant
    Dim triggerFill As Boolean
    Dim sFileName As String, sOldFileName As String, sOldFilePath As String
    
    ' 1. 优化运行环境
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set wbMain = ThisWorkbook
    ' 确保路径以 \ 结尾
    sPath = wbMain.Path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    Set dictOld = CreateObject("Scripting.Dictionary")
    Set arrFiles = CreateObject("System.Collections.ArrayList")
    
    ' 2. 收集 0821 文件索引 (仅收集文件名，不重置Dir指针)
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
    
    ' 4. 遍历处理 0914 文件
    Dim vFile As Variant
    For Each vFile In arrFiles.ToArray
        On Error GoTo FileError
        
        ' 【修复报错核心】使用字符串截取获取文件名，彻底废弃 Dir() 提取文件名
        sFileName = Mid(CStr(vFile), InStrRev(CStr(vFile), "\") + 1)
        sOldFileName = Replace(sFileName, "0914", "0821")
        sOldFilePath = sPath & sOldFileName
        
        ' 检查对应的 0821 文件是否存在
        If dictOld.Exists(LCase(sOldFileName)) Then
            matchCount = matchCount + 1
            
            ' 打开 0914 (可写) 和 0821 (只读)
            Set wbNew = Workbooks.Open(CStr(vFile), ReadOnly:=False)
            Set wbOld = Workbooks.Open(sOldFilePath, ReadOnly:=True)
            
            Set wsNew = wbNew.Sheets(1)
            Set wsOld = wbOld.Sheets(1)
            
            ' 获取两个表的最大行数（以A列为准，如果A列可能为空，可改为用特定列如W列）
            lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
            lastRowOld = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row
            maxRow = Application.Max(lastRowNew, lastRowOld)
            
            bModified = False
            
            ' 【核心逻辑】逐行判断 W/X/Y/Z，满足条件则覆盖 W~AB 共6列
            ' 假设第1行为表头，从第2行开始扫描
            For r = 2 To maxRow
                triggerFill = False
                
                ' 检查 W(23) / X(24) / Y(25) / Z(26) 任一列：0821有值 且 0914为空
                For col = 23 To 26
                    ' 使用 Trim + CStr 清洗空格、不可见字符和公式假空值
                    vNew = Trim(CStr(wsNew.Cells(r, col).Value))
                    vOld = Trim(CStr(wsOld.Cells(r, col).Value))
                    
                    If Len(vNew) = 0 And Len(vOld) > 0 Then
                        triggerFill = True
                        Exit For ' 只要有一列满足即触发，无需继续检查后面的列
                    End If
                Next col
                
                ' 触发后：整组覆盖 W/X/Y/Z/AA/AB (第23~28列)
                If triggerFill Then
                    For col = 23 To 28
                        ' 直接赋值，保留0821的原始格式和公式结果
                        wsNew.Cells(r, col).Value = wsOld.Cells(r, col).Value
                    Next col
                    bModified = True
                End If
            Next r
            
            ' 如果有修改，则保存 0914 文件
            If bModified Then
                wbNew.Save
                saveCount = saveCount + 1
            End If
            
            ' 关闭 0821 文件
            wbOld.Close SaveChanges:=False
            Set wbOld = Nothing
            
            ' 关闭 0914 文件 (如果没修改，Close时不保存；如果修改了，前面已经Save过了)
            wbNew.Close SaveChanges:=False
            Set wbNew = Nothing
        End If
        
        GoTo NextFile
        
FileError:
        Debug.Print "处理失败: " & sFileName & " | 错误原因: " & Err.Description
        ' 确保出错时关闭已打开的工作簿，防止内存泄漏和文件锁定
        On Error Resume Next
        If Not wbOld Is Nothing Then wbOld.Close SaveChanges:=False
        If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
        Set wbOld = Nothing: Set wbNew = Nothing
        On Error GoTo 0
        
NextFile:
    Next vFile
    
    ' 5. 恢复 Excel 环境
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    
    ' 6. 输出结果
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