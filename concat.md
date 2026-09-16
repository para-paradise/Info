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
Sub BatchMergeData()
    Dim sPath As String
    Dim sFile As String
    Dim wbTool As Workbook
    Dim wsResult As Worksheet
    Dim wbNew As Workbook, wbOld As Workbook
    Dim wsNew As Worksheet, wsOld As Worksheet
    Dim arrData() As Variant
    Dim i As Long, lLastRow As Long, lLastCol As Long
    Dim r As Long, c As Long
    Dim dictOldFiles As Object
    Dim sBaseName As String
    Dim sOldName As String
    
    ' 关闭屏幕刷新，提升速度
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbTool = ThisWorkbook
    sPath = wbTool.Path & "\"
    
    ' 1. 扫描文件夹，提取所有0821旧版文件的文件名，存入字典
    Set dictOldFiles = CreateObject("Scripting.Dictionary")
    sFile = Dir(sPath & "*.xlsx")
    Do While sFile <> ""
        If InStr(1, sFile, "0821") > 0 Then
            ' 提取前缀作为Key，例如 "01VSACN-产品_数据通信DFX测试模式库"
            sBaseName = Left(sFile, InStrRev(sFile, "_") - 1)
            If Not dictOldFiles.Exists(sBaseName) Then
                dictOldFiles.Add sBaseName, sFile
            End If
        End If
        sFile = Dir
    Loop
    
    ' 2. 创建一个新的结果Sheet，命名为 "最终合并结果"
    ' 如果已经存在，直接清空；如果不存在，新建
    On Error Resume Next
    Set wsResult = wbTool.Sheets("最终合并结果")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = wbTool.Sheets.Add(After:=wbTool.Sheets(wbTool.Sheets.Count))
        wsResult.Name = "最终合并结果"
    Else
        wsResult.Cells.Clear
    End If
    
    Dim lDestRow As Long
    lDestRow = 1
    
    ' 3. 遍历所有0914新版文件进行合并
    sFile = Dir(sPath & "*.xlsx")
    Do While sFile <> ""
        ' 排除自身、排除0821文件、排除已经生成的结果文件
        If sFile <> "合并工具.xlsm" And sFile <> "合并结果.xlsx" And sFile <> "最终合并结果.xlsx" Then
            If InStr(1, sFile, "0914") > 0 Then
                ' 提取当前0914文件的基础名称
                sBaseName = Left(sFile, InStrRev(sFile, "_") - 1)
                
                ' 尝试打开新版文件
                On Error Resume Next
                Set wbNew = Workbooks.Open(sPath & sFile, False, True) ' 只读打开
                If Err.Number <> 0 Then
                    Err.Clear
                    GoTo NextFile
                End If
                On Error GoTo 0
                
                Set wsNew = wbNew.Sheets(1)
                lLastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
                lLastCol = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column
                
                ' 4. 核心逻辑：如果该0914有对应的0821旧文件，进行数据补充
                If dictOldFiles.Exists(sBaseName) Then
                    sOldName = dictOldFiles(sBaseName)
                    
                    On Error Resume Next
                    Set wbOld = Workbooks.Open(sPath & sOldName, False, True)
                    On Error GoTo 0
                    
                    If Not wbOld Is Nothing Then
                        Set wsOld = wbOld.Sheets(1)
                        arrData = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(lLastRow, lLastCol)).Value
                        
                        ' 遍历新版数据，如果单元格为空且旧版对应位置有值，则填充旧版值
                        For r = 1 To UBound(arrData, 1)
                            For c = 1 To UBound(arrData, 2)
                                If Trim(arrData(r, c)) = "" Then
                                    If Not IsEmpty(wsOld.Cells(r, c)) And Trim(wsOld.Cells(r, c).Value) <> "" Then
                                        arrData(r, c) = wsOld.Cells(r, c).Value
                                    End If
                                End If
                            Next c
                        Next r
                        
                        ' 将处理好的数据写入结果Sheet
                        wsResult.Range(wsResult.Cells(lDestRow, 1), wsResult.Cells(lDestRow + lLastRow - 1, lLastCol)).Value = arrData
                        lDestRow = lDestRow + lLastRow
                        
                        wbOld.Close False
                        Set wbOld = Nothing
                    Else
                        ' 如果没有对应的旧版文件，直接把0914的数据复制过去
                        wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(lLastRow, lLastCol)).Copy Destination:=wsResult.Cells(lDestRow, 1)
                        lDestRow = lDestRow + lLastRow
                    End If
                Else
                    ' 如果没有对应的旧版文件，直接把0914的数据复制过去
                    wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(lLastRow, lLastCol)).Copy Destination:=wsResult.Cells(lDestRow, 1)
                    lDestRow = lDestRow + lLastRow
                End If
                
                wbNew.Close False
                Set wbNew = Nothing
            End If
        End If
NextFile:
        sFile = Dir
    Loop
    
    ' 恢复屏幕刷新
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "恭喜！所有0914数据已根据0821增量补充完毕，结果已保存至【最终合并结果】Sheet中！", vbInformation, "批量合并完成"
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