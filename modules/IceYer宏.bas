Attribute VB_Name = "IceYer宏"
Sub Excel压字调整高度()
'
' Excel压字调整高度 宏
' 在自动识别行高基础上增加指定高度，可以有效解决压字情况。
'

    Dim sht As Worksheet, r As Range
    Dim padPt As Double
    Dim totalRows As Long, i As Long, stepCount As Long
    Dim userInput As Variant

    '=== 让用户输入行高加值 ===
    userInput = InputBox("请输入每行额外增加的高度（单位：pt）" & vbCrLf & _
                         "（按“取消”可中止执行）", _
                         "调整行高", 8.504)

    If StrPtr(userInput) = 0 Then Exit Sub    ' 取消直接退出

    If IsNumeric(userInput) Then
        padPt = CDbl(userInput)
    Else
        padPt = 0
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "正在计算，请稍候..."

    '=== 主体逻辑 ===
    For Each sht In ActiveWindow.SelectedSheets
        With sht.UsedRange
            .WrapText = True
            .EntireRow.AutoFit

            totalRows = .Rows.Count
            stepCount = Application.Max(1, totalRows \ 100) ' 每1%刷新一次状态栏

            For i = 1 To totalRows
                If .Rows(i).RowHeight > 0 Then
                    .Rows(i).RowHeight = .Rows(i).RowHeight + padPt
                End If

                ' 每隔若干行刷新进度提示
                If i Mod stepCount = 0 Then
                    Application.StatusBar = "正在调整行高：" & _
                        Format(i / totalRows, "0%") & " 完成"
                    DoEvents   ' 防止假死
                End If
            Next i
        End With
    Next sht

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "已完成自动行高调整，每行增加 " & padPt & " pt。", vbInformation

End Sub


