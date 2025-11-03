' ===========================
' 作用：导入VBA宏至Excel个人宏工作簿
' 作者：IceYer
' 原理：批量导入 modules/*.bas 到 PERSONAL.XLSB
' 1.通过注册表决定是否开启 "信任对 VBA 工程对象模型的访问"；
' 2.结束 Excel 进程，创建/合并 PERSONAL.XLSB；
' 3.依据 *.bas 首行判断 Attribute VB_Name 决定导入/覆盖/跳过；
' 4.通过注册表决定是否还原 "信任对 VBA 工程对象模型的访问"；
' 5.弹窗提示。
' ===========================

Option Explicit

' -------- 实用函数区域 --------

' 读取注册表 DWORD，不存在则返回默认值
Function RegReadDwordOrDefault(keyPath, valueName, defaultVal)
    Dim wsh : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    Dim val : val = wsh.RegRead(keyPath & valueName)
    If Err.Number <> 0 Then
        RegReadDwordOrDefault = defaultVal
        Err.Clear
    Else
        RegReadDwordOrDefault = CLng(val)
    End If
    On Error GoTo 0
End Function

' 写注册表 DWORD（自动创建分支）
Sub RegWriteDword(keyPath, valueName, byVal dwordVal)
    Dim wsh : Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite keyPath & valueName, CLng(dwordVal), "REG_DWORD"
End Sub

' 检测注册表中 Excel 各版本的“信任对 VBA 工程对象模型的访问”设置
' 统一查找 Excel 的 AccessVBOM，策略优先；未命中则回退到最高版本用户路径
Function GetExcelVBOMPath(isPolicy)
    Dim wsh : Set wsh = CreateObject("WScript.Shell")
    Dim vers : vers  = Array("16.0","15.0","14.0","12.0","11.0","10.0","9.0","8.0")
    Dim roots: roots = Array("HKCU\Software\Policies\Microsoft\Office\", _
                             "HKCU\Software\Microsoft\Office\")
    Dim r, v, path, ri

    For ri = 0 To UBound(roots)
        For Each v In vers
            path = roots(ri) & v & "\Excel\Security\"
            On Error Resume Next
            wsh.RegRead path & "AccessVBOM"
            If Err.Number = 0 Then
                On Error GoTo 0
                isPolicy = (ri = 0)
                GetExcelVBOMPath = path
                Exit Function
            End If
            On Error GoTo 0
        Next
    Next

    ' 未找到则回退到最高版本的用户路径（写入时会自动创建项）
    isPolicy = False
    GetExcelVBOMPath = "HKCU\Software\Microsoft\Office\" & vers(0) & "\Excel\Security\"
End Function


' 结束所有 EXCEL.EXE（防止已打开实例锁住 PERSONAL.XLSB）
Sub KillExcelProcesses()
    On Error Resume Next
    Dim svc : Set svc = GetObject("winmgmts:\\.\root\cimv2")
    Dim procs : Set procs = svc.ExecQuery("Select * from Win32_Process Where Name='EXCEL.EXE'")
    Dim p
    For Each p In procs
        p.Terminate
    Next
    On Error GoTo 0
End Sub

' 正则匹配 bas 文件中的 Attribute VB_Name = "模块名"
Function ExtractVBNameFromBas(filePath)
    ' 使用系统默认编码读取文件
    On Error Resume Next
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts, firstLine
    Set ts = fso.OpenTextFile(filePath, 1, False, -2) ' TristateUseDefault = -2
    If Err.Number = 0 Then
        If Not ts.AtEndOfStream Then firstLine = ts.ReadLine
        ts.Close
    End If
    On Error GoTo 0

    ' 去掉 UTF-8 BOM（若存在）
    If VarType(firstLine) = vbString Then
        firstLine = Replace(firstLine, ChrW(65279), "")
    Else
        firstLine = ""
    End If

    ' 正则匹配，如果有注释符号则跳过导入模块
    Dim re : Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*Attribute\s+VB_Name\s*=\s*""([^""]+)"""
    re.IgnoreCase = True
    re.Global = False
    Dim m : Set m = re.Execute(firstLine)
    If m.Count > 0 Then
        ExtractVBNameFromBas = m(0).SubMatches(0)
    Else
        ExtractVBNameFromBas = ""
    End If
End Function


' 模块是否存在
Function VBComponentExists(vbproj, compName)
    Dim comp
    For Each comp In vbproj.VBComponents
        If StrComp(comp.Name, compName, vbTextCompare) = 0 Then
            VBComponentExists = True
            Exit Function
        End If
    Next
    VBComponentExists = False
End Function

' 删除模块
Sub RemoveVBComponentByName(vbproj, compName)
    Dim comp
    For Each comp In vbproj.VBComponents
        If StrComp(comp.Name, compName, vbTextCompare) = 0 Then
            vbproj.VBComponents.Remove comp
            Exit Sub
        End If
    Next
End Sub

' -------- 主流程 --------

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim baseDir : baseDir = fso.GetParentFolderName(WScript.ScriptFullName) ' 脚本所在目录
Dim modulesDir : modulesDir = fso.BuildPath(baseDir, "modules")

If Not fso.FolderExists(modulesDir) Then
    WScript.Echo "未找到 modules 文件夹：" & modulesDir
    WScript.Quit 1
End If

' 1) 处理 VBOM 设置：读取原值 -> 写入 1（开启）
Dim secPath, isPolicy
secPath = GetExcelVBOMPath(isPolicy)
Dim vbomOriginal : vbomOriginal = RegReadDwordOrDefault(secPath, "AccessVBOM", 0)

Dim policyMsg
If isPolicy Then
    policyMsg = vbCrLf & "注意：当前路径为策略路径：" & secPath & _
                "（策略优先，若 AccessVBOM=0 可能阻止导入操作）。"
Else
    policyMsg = ""
End If

' 写临时值为 1
On Error Resume Next
RegWriteDword secPath, "AccessVBOM", 1
Dim writeErr : writeErr = Err.Number
On Error GoTo 0

Dim vbomTemp : vbomTemp = RegReadDwordOrDefault(secPath, "AccessVBOM", vbomOriginal)

' 2) 结束 Excel 进程
KillExcelProcesses()

' 3) 启动 Excel，创建/合并 PERSONAL.XLSB 并导入模块
Dim excel, wbPersonal, personalPath, importLog : importLog = ""
Dim succeededCount : succeededCount = 0
Dim replacedCount  : replacedCount  = 0
Dim failedCount    : failedCount    = 0
Dim skippedCount   : skippedCount   = 0 

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If excel Is Nothing Then
    MsgBox "无法启动 Excel。请确认已安装。", vbCritical, "导入失败"
    ' 还原设置后再退出
    RegWriteDword secPath, "AccessVBOM", vbomOriginal
    WScript.Quit 2
End If
On Error GoTo 0

excel.DisplayAlerts = False
excel.Visible = False  ' 如需观察过程可改为 True

' XLSTART 路径（个人宏工作簿默认存放处）
personalPath = excel.StartupPath & "\PERSONAL.XLSB"

' 打开或创建 PERSONAL.XLSB
If fso.FileExists(personalPath) Then
    Set wbPersonal = excel.Workbooks.Open(personalPath, False, False) ' 读写打开
Else
    Dim wbNew : Set wbNew = excel.Workbooks.Add()
    ' 50 = xlExcel12 (二进制工作簿 .xlsb)
    wbNew.SaveAs personalPath, 50
    Set wbPersonal = wbNew
End If

' 取 VBA 工程
Dim vbProj : Set vbProj = wbPersonal.VBProject

' 遍历 modules 目录下的所有 .bas
Dim file
For Each file In fso.GetFolder(modulesDir).Files
    If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
        Dim fullPath : fullPath = file.Path

        ' 只看首行且必须以 Attribute VB_Name 开头
        Dim vbName : vbName = ExtractVBNameFromBas(fullPath)

        If vbName = "" Then
            ' 无有效 VB_Name -> 跳过导入
            importLog    = importLog & "× 跳过 " & file.Name & " (首行未启用)" & vbCrLf
            skippedCount = skippedCount + 1

        Else
            On Error Resume Next

            ' 覆盖：如已有同名组件则先删
            Dim isReplace : isReplace = False
            If VBComponentExists(vbProj, vbName) Then
                RemoveVBComponentByName vbProj, vbName
                isReplace = True
            End If

            ' 执行导入
            vbProj.VBComponents.Import fullPath

            If Err.Number = 0 Then
                If isReplace Then
                    importLog = importLog & "√ 覆盖 " & file.Name & " (" & vbName & ")" & vbCrLf
                    replacedCount = replacedCount + 1
                Else
                    importLog = importLog & "√ 导入 " & file.Name & " (" & vbName & ")" & vbCrLf
                    succeededCount = succeededCount + 1
                End If
            Else
                importLog = importLog & "× 失败 " & file.Name & " (" & vbName & ") -> " & Err.Description & vbCrLf
                failedCount = failedCount + 1
                Err.Clear
            End If

            On Error GoTo 0
        End If
    End If
Next

' 保存并退出
On Error Resume Next
excel.Windows("PERSONAL.XLSB").Visible = False ' 保存前隐藏 PERSONAL.XLSB，以防下次启动显示
wbPersonal.Save
wbPersonal.Close False
excel.Quit
WScript.Sleep 800   ' 等待进程完全退出，否则还原注册表失败
Set excel = Nothing
On Error GoTo 0

' 4) 还原 VBOM 原值
RegWriteDword secPath, "AccessVBOM", vbomOriginal
Dim vbomRestored : vbomRestored = RegReadDwordOrDefault(secPath, "AccessVBOM", vbomOriginal)

' 汇总信息
Dim msg
msg = "统计：导入 " & succeededCount & "，覆盖 " & replacedCount & "，失败 " & failedCount & "，跳过 " & skippedCount & vbCrLf & String(40,"-") & vbCrLf & _
      "位置: " & personalPath & vbCrLf & vbCrLf & _
      "Attribute VB_Name 详情：" & vbCrLf & importLog & vbCrLf & _
      "信任对 VBA 工程对象模型的访问：" & vbCrLf & _
      "原始值=" & vbomOriginal & " → 临时值=" & vbomTemp & " → 还原值=" & vbomRestored & _
      policyMsg

If writeErr <> 0 Then
    msg = "警告：尝试写入 VBOM 设置失败（Err=" & writeErr & "）。" & vbCrLf & _
          "可能需要以管理员身份运行，或被组策略限制。" & vbCrLf & vbCrLf & msg
End If

MsgBox msg, vbInformation, "PERSONAL.XLSB 导入完成"
