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
Sub MergeOldToNew()
    Dim sPath As String
    Dim sFile As String
    Dim wbNew As Workbook, wbOld As Workbook, wbMaster As Workbook
    Dim wsNew As Worksheet, wsOld As Worksheet, wsResult As Worksheet
    Dim dictFiles As Object
    Dim key As Variant
    Dim baseName As String
    Dim isNew As Boolean
    
    ' === 设置部分 ===
    ' 获取当前代码所在工作簿的路径（也就是你的文件夹路径）
    sPath = ThisWorkbook.Path & "\"
    
    ' 关闭屏幕刷新和弹窗，加快速度
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set dictFiles = CreateObject("Scripting.Dictionary")
    Set wbMaster = ThisWorkbook
    
    ' === 第一步：遍历文件夹，按文件名前缀分组 ===
    sFile = Dir(sPath & "*.xlsx")
    Do While sFile <> ""
        ' 排除掉宏文件自己
        If sFile <> ThisWorkbook.Name Then
            ' 假设文件名格式是 "01VSACN..."，取前两个字符作为唯一标识
            ' 如果你的编号长度不固定，需要调整这里的截取逻辑
            ' 这里假设前2位是编号，或者根据下划线分割，这里简单取前2位演示
            ' 更稳妥的方式是提取 "20260821" 或 "20260914" 之前的特征字符串
            
            ' 简单的逻辑：提取 "2026" 之前的字符串作为 Key
            Dim tempName As String
            tempName = Left(sFile, InStr(sFile, "2026") - 1) 
            
            If Not dictFiles.Exists(tempName) Then
                dictFiles.Add tempName, sFile
            Else
                ' 如果字典里已经有了，说明遇到了一对文件
                ' 此时字典里存的是先遍历到的，我们需要判断新旧
                Dim existingFile As String
                existingFile = dictFiles(tempName)
                
                ' 简单的判断逻辑：如果当前文件包含 0914，它通常是新的；如果是 0821，它是旧的
                ' 这里我们不做复杂判断，而是把所有文件都存下来处理
                ' 为了简化，我们直接在循环里处理配对
            End If
        End If
        sFile = Dir
    Loop
    
    ' === 第二步：重新遍历进行合并（为了逻辑清晰，这里采用两两匹配法）===
    ' 上面的字典法对于多文件稍微复杂，我们改用直接双重循环或者标记法
    ' 鉴于只有40个文件，直接再次遍历查找配对最快
    
    Dim processedFiles As Object
    Set processedFiles = CreateObject("Scripting.Dictionary")
    
    sFile = Dir(sPath & "*.xlsx")
    Do While sFile <> ""
        If sFile <> ThisWorkbook.Name And Not processedFiles.Exists(sFile) Then
            
            ' 寻找它的配对文件
            Dim partnerFile As String
            partnerFile = ""
            
            ' 确定当前文件的日期标识
            Dim currentDate As String
            If InStr(sFile, "20260914") > 0 Then
                currentDate = "20260914"
            ElseIf InStr(sFile, "20260821") > 0 Then
                currentDate = "20260821"
            Else
                ' 如果不是这两个日期的文件，跳过或单独处理
                GoTo NextFile
            End If
            
            ' 寻找另一个日期的对应文件
            Dim targetDate As String
            targetDate = IIf(currentDate = "20260914", "20260821", "20260914")
            
            ' 构造配对文件名的特征（去除日期后的前后缀）
            ' 比如: 01VSACN-V产品_数据通信DFX测试模式库[DATE]_安全性测试_License管理特性分析-已完成.xlsx
            ' 这种精确匹配比较难，我们假设除了日期不同，其他都相同？
            ' 或者根据截图，可能是同一个测试项的不同版本。
            
            ' 让我们用一个更通用的方法：找同组文件
            ' 假设同组文件的区别仅在于日期字符串
            Dim searchPattern As String
            searchPattern = Replace(sFile, currentDate, targetDate)
            
            If Len(Dir(sPath & searchPattern)) > 0 Then
                partnerFile = searchPattern
            End If
            
            ' 只有当当前文件是 0914 (新版) 时，才执行合并操作
            ' 这样可以避免重复处理（处理了A+B，就不处理B+A）
            If currentDate = "20260914" And partnerFile <> "" Then
                
                ' 打开两个文件
                Set wbNew = Workbooks.Open(sPath & sFile)       ' 0914 文件
                Set wbOld = Workbooks.Open(sPath & partnerFile) ' 0821 文件
                
                ' 假设数据都在第一个 Sheet，根据实际情况修改 Sheet 名称
                Set wsNew = wbNew.Sheets(1)
                Set wsOld = wbOld.Sheets(1)
                
                ' 在新版文件中添加结果 Sheet
                ' 先删除可能存在的旧 Sheet
                On Error Resume Next
                wbNew.Sheets("合并结果").Delete
                On Error GoTo 0
                Set wsResult = wbNew.Sheets.Add(After:=wsNew)
                wsResult.Name = "合并结果"
                
                ' 获取数据范围 (假设 A1 开始)
                Dim lastRow As Long, lastCol As Long
                ' 取两个文件中较大的行数，防止漏数据
                lastRow = Application.Max(wsNew.UsedRange.Rows.Count, wsOld.UsedRange.Rows.Count)
                lastCol = Application.Max(wsNew.UsedRange.Columns.Count, wsOld.UsedRange.Columns.Count)
                
                ' 写入公式
                ' 注意：公式中引用外部工作簿需要使用完整路径或文件名
                ' 格式: ='[文件名.xlsx]Sheet名'!单元格
                
                Dim r As Long, c As Long
                ' 为了提高速度，我们可以直接写入数组，或者分块写入
                ' 这里为了演示清晰，使用 Range 写入公式（速度尚可）
                
                ' 优化：一次性给整个区域赋值公式
                Dim formulaStr As String
                ' 构建 R1C1 样式的公式，方便批量填充
                ' 逻辑：如果 New(0914) 不为空，取 New，否则取 Old(0821)
                ' 注意：R1C1 引用中，外部文件引用比较麻烦，还是用 A1 样式循环第一行，然后下拉？
                ' 不，直接用 FillDown 太慢。
                
                ' 既然都在内存里打开了，直接用值操作更快！
                ' 不需要写公式了，直接用 VBA 判断值！
                
                Dim arrNew As Variant, arrOld As Variant, arrRes As Variant
                arrNew = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(lastRow, lastCol)).Value
                arrOld = wsOld.Range(wsOld.Cells(1, 1), wsOld.Cells(lastRow, lastCol)).Value
                
                ReDim arrRes(1 To lastRow, 1 To lastCol)
                
                For r = 1 To lastRow
                    For c = 1 To lastCol
                        Dim valNew As Variant, valOld As Variant
                        valNew = arrNew(r, c)
                        valOld = arrOld(r, c)
                        
                        ' 核心逻辑：B(新)有值取B，否则取A(旧)
                        ' 注意处理空字符串和 Null
                        If Not IsEmpty(valNew) And CStr(valNew) <> "" Then
                            arrRes(r, c) = valNew
                        Else
                            arrRes(r, c) = valOld
                        End If
                    Next c
                Next r
                
                ' 将结果一次性写入 Sheet
                wsResult.Range("A1").Resize(lastRow, lastCol).Value = arrRes
                
                ' 保存并关闭旧文件
                wbOld.Close SaveChanges:=False
                
                ' 保存新文件（包含新增的Sheet）
                wbNew.Close SaveChanges:=True
                
                ' 记录已处理，防止重复
                processedFiles.Add sFile, 1
                processedFiles.Add partnerFile, 1
                
                Debug.Print "已合并: " & sFile
            End If
        End If
NextFile:
        sFile = Dir
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "所有文件处理完毕！请检查各 0914 文件中的【合并结果】Sheet。"
    
End Sub
```

### 第三步：运行
1. 回到 Excel 界面。
2. 按 `Alt + F8`，选择 `MergeOldToNew`，点击“运行”。
3. 等待几分钟（取决于文件大小和数量），程序会自动打开每一对文件，计算并保存。

---

### 这个脚本做了什么？

1.  **智能配对**：它会自动识别文件名里的 `20260914` 和 `20260821`，把它们凑成一对。
2.  **以新补旧**：它会打开这两个文件，读取数据到内存。
3.  **内存运算**：在内存中执行“如果有新值就用新值，否则用旧值”的逻辑（比 Excel 公式快得多）。
4.  **生成结果**：直接把算好的纯数值写入到 `0914` 那个文件的新 Sheet `合并结果` 中。
5.  **自动保存**：处理完一对就自动保存关闭，最后告诉你完成了。

### ⚠️ 重要提示
*   **备份！** 运行前最好把这 40 个文件复制一份备份，防止意外覆盖。
*   **Sheet 名称**：代码默认读取每个文件的**第 1 个 Sheet** (`Sheets(1)`). 如果你的数据不在第一个 Sheet，需要修改代码中的 `Set wsNew = wbNew.Sheets(1)` 为具体的名字，例如 `wbNew.Sheets("Sheet1")`。
*   **文件格式**：代码假设所有文件都是 `.xlsx`。

你可以先拿两个文件试一下，没问题再跑全部。

---

VBA 方案能一步到位。不过如果你担心代码不好维护，我也可以教你用 Power Query 的“追加查询”功能，不用写代码也能搞定，想试试吗？