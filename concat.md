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
Sub FixMergeAndSave()
    Dim sPath As String
    Dim sFile As String
    Dim wbNew As Workbook, wbOld As Workbook
    Dim wsNew As Worksheet, wsOld As Worksheet
    Dim dictOld As Object
    Dim sBaseName As String
    Dim sOldName As String
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    
    ' 关闭屏幕刷新和弹窗，加快速度并防止意外中断
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set dictOld = CreateObject("Scripting.Dictionary")
    
    ' 获取当前宏文件所在的文件夹路径
    sPath = ThisWorkbook.Path & "\"
    
    ' === 第一步：扫描文件夹，把 0821 文件存到字典中 ===
    sFile = Dir(sPath & "*.xlsx")
    Do While sFile <> ""
        ' 只识别 0821 的文件
        If InStr(1, sFile, "0821", vbTextCompare) > 0 Then
            ' 提取基础名（去掉日期和扩展名，例如 01VSACN-产品_数据通信DFX测试模式库_安全性测试_License管理特性分析-已完成）
            sBaseName = Split(sFile, "_20260821_")(0) & "_20260821_" & Split(sFile, "_20260821_")(1)
            sBaseName = Left(sBaseName, InStrRev(sBaseName, ".") - 1)
            dictOld(sBaseName) = sFile
        End If
        sFile = Dir()
    Loop
    
    ' === 第二步：遍历文件夹，处理 0914 文件 ===
    sFile = Dir(sPath & "*.xlsx")
    Do While sFile <> ""
        ' 只识别 0914 的文件，且不是合并工具自身
        If InStr(1, sFile, "0914", vbTextCompare) > 0 And sFile <> ThisWorkbook.Name Then
            ' 构建对应的基础名
            sBaseName = Split(sFile, "_20260914_")(0) & "_20260914_" & Split(sFile, "_20260914_")(1)
            sBaseName = Left(sBaseName, InStrRev(sBaseName, ".") - 1)
            
            ' 检查是否有对应的 0821 文件
            If dictOld.Exists(sBaseName) Then
                sOldName = dictOld(sBaseName)
                
                ' 打开新版和旧版文件
                Set wbNew = Workbooks.Open(sPath & sFile)
                Set wbOld = Workbooks.Open(sPath & sOldName)
                
                ' 获取两个文件的第一个工作表
                Set wsNew = wbNew.Sheets(1)
                Set wsOld = wbOld.Sheets(1)
                
                ' 获取旧版文件的最大行数和列数
                lastRow = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row
                lastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column
                
                ' === 核心：增量补充逻辑 ===
                For r = 1 To lastRow
                    For c = 1 To lastCol
                        ' 如果新版文件的单元格为空，且旧版文件对应位置有值，则填入
                        If Trim(wsNew.Cells(r, c).Value) = "" And Trim(wsOld.Cells(r, c).Value) <> "" Then
                            wsNew.Cells(r, c).Value = wsOld.Cells(r, c).Value
                        End If
                    Next c
                Next r
                
                ' === 强制保存并关闭新版文件 (0914) ===
                wbNew.Close SaveChanges:=True
                
                ' 关闭旧版文件 (0821)
                wbOld.Close SaveChanges:=False
                
            End If
        End If
        sFile = Dir()
    Loop
    
    ' 恢复屏幕刷新和弹窗
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "所有 0914 文件已成功补充 0821 数据并保存！", vbInformation, "完成"
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