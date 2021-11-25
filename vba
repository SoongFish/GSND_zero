Public Function findEndX(sht As Worksheet, y As Integer) As Long 'y행의 끝 return

    findEndX = sht.Cells(y, Columns.Count).End(1).Column
    
End Function

Public Function FindEndY(sht As Worksheet, x As Integer) As Long 'x열의 끝 return

    FindEndY = sht.Cells(Rows.Count, x).End(3).Row
    
End Function

Sub start_sign()

    Call 전처리_출장
    UserForm1.ProgressBar1.Value = 5
    UserForm1.Repaint
    
    Call 전처리_결재
    UserForm1.ProgressBar1.Value = 10
    UserForm1.Repaint
    
    Call 작업_결재
    UserForm1.ProgressBar1.Value = 90
    UserForm1.Repaint
    
    Call 시간_결재
    UserForm1.ProgressBar1.Value = 100
    UserForm1.Repaint
    
    UserForm1.CommandButton1.Caption = "작업 완료"
    MsgBox "작업이 완료되었습니다."

End Sub

Sub start_restaurant()
    
    Call 전처리_출장
    UserForm1.ProgressBar1.Value = 5
    UserForm1.Repaint
    
    Call 전처리_식당
    UserForm1.ProgressBar1.Value = 10
    UserForm1.Repaint
    
    Call 작업_식당
    UserForm1.ProgressBar1.Value = 100
    UserForm1.Repaint
    
    UserForm1.CommandButton1.Caption = "작업 완료"
    MsgBox "작업이 완료되었습니다."

End Sub

Sub 전처리_출장()

' 표준소속 변환
    Sheets("출장").Select
    Columns("A:A").Select
    Selection.Replace What:="재난안전건설본부", Replacement:="재난안전건설국", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IFERROR(MID(RC[1],FIND(""실 "",RC[1])+2, LEN(RC[1])),MID(RC[1],FIND(""국 "",RC[1])+2, LEN(RC[1]))), RC[1])"
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "표준소속"
    
' 키 생성
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[1],RC[4])"
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "키"
    
' 키에 날짜 추가 (최종키)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[1], RC[7])"
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "최종키"

End Sub

Sub 전처리_결재()

' 표준소속 변환
    Sheets("결재").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(RC[1], FIND(""-"", RC[1])-1)"
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "표준소속"

' 키 생성
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[1],RC[5])"
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "키"
    
    
' 결재시간 추가
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Value = "=RIGHT(RC[5], 4)"
    Range("A2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Selection.NumberFormatLocal = "0_ "
    Range("A1").Value = "결재시간"
    
' 키에 날짜 추가 (최종키)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[2], LEFT(RC[6], 8))"
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "최종키"

    
' 최종키 정렬
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), findEndX(ActiveSheet, 1))).Sort Key1:=Cells(1, 1), Order1:=xlAscending
    
End Sub

Sub 전처리_식당()

    '최종키 생성
    Sheets("식당").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[2],RC[3], RC[5])"
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 2), 1))
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "최종키"
    
    ' 최종키 정렬
    Range(Cells(2, 1), Cells(FindEndY(ActiveSheet, 1), findEndX(ActiveSheet, 1))).Sort Key1:=Cells(1, 1), Order1:=xlAscending
    
End Sub

Sub 작업_결재()

    Dim flag_pass As Boolean
    flag_pass = False
    Dim loop_end As Long
    
    Sheets("결재").Select
    Range("L1").Value = "출장여부"
    
    loop_end = FindEndY(ActiveSheet, 1)
    For i = 2 To loop_end
        Range("A1").Value = "=IFNA(VLOOKUP(R[" & i - 1 & "]C, 출장!C, 1, FALSE),0)"
        If Range("A1").Value = 0 Then ' 출장내역이 없으면
            Do While Cells(i, 1).Value = Cells(i + 1, 1).Value
                i = i + 1
            Loop
        Else
            Cells(i, 12).Value = "O"
        End If
continue:
    'UserForm1.ProgressBar1.Value = UserForm1.ProgressBar1.Value + (95 / loop_end)
    'UserForm1.Repaint
    Next
    
    Range("A1").Value = "최종키"
    
    ' 출장여부 필터링
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(FindEndY(ActiveSheet, 1), findEndX(ActiveSheet, 1))).AutoFilter Field:=12, Criteria1:="O"
    Range(Cells(1, 1), Cells(FindEndY(ActiveSheet, 1), findEndX(ActiveSheet, 1))).Select
    Selection.Copy
    
    ' "분석결과(결재)" 시트로 결과 이동
    Sheets("분석결과(결재)").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("결재").Select
    Selection.AutoFilter
    Range("A1").Select
    Sheets("분석결과(결재)").Select
    Range("A1").Select
    
End Sub

Sub 작업_식당()

    Dim flag_pass As Boolean
    flag_pass = False
    Dim loop_end As Long
    
    Sheets("식당").Select
    Range("M1").Value = "출장여부"
    
    loop_end = FindEndY(ActiveSheet, 1)
    For i = 2 To loop_end
        Range("A1").Value = "=IFNA(VLOOKUP(R[" & i - 1 & "]C, 출장!C, 1, FALSE),0)"
        If Range("A1").Value = 0 Then ' 출장내역이 없으면
            Do While Cells(i, 1).Value = Cells(i + 1, 1).Value
                i = i + 1
            Loop
        Else
            Cells(i, 13).Value = "O"
        End If
continue:
    UserForm1.ProgressBar1.Value = UserForm1.ProgressBar1.Value + (95 / loop_end)
    UserForm1.Repaint
    Next
    
    Range("A1").Value = "최종키"
    
    ' 출장여부 필터링
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(FindEndY(ActiveSheet, 1), findEndX(ActiveSheet, 1))).AutoFilter Field:=13, Criteria1:="O"
    Range(Cells(1, 1), Cells(FindEndY(ActiveSheet, 1), findEndX(ActiveSheet, 1))).Select
    Selection.Copy
    
    ' "분석결과(결재)" 시트로 결과 이동
    Sheets("분석결과(식당)").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("식당").Select
    Selection.AutoFilter
    Range("A1").Select
    Sheets("분석결과(식당)").Select
    Range("A1").Select

End Sub


Sub 시간_결재()

    Range("M1").Value = "출장시작"
    Range("M2").Select
    Range("M2").Value = "=VLOOKUP(RC[-12], 출장!C[-12]:C[-2], 10, FALSE)"
    Selection.AutoFill Destination:=Range(Cells(2, 13), Cells(FindEndY(ActiveSheet, 2), 13))
    Range(Cells(2, 13), Cells(FindEndY(ActiveSheet, 2), 13)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(Cells(2, 13), Cells(FindEndY(ActiveSheet, 2), 13)).Select
    Selection.Replace What:=":", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("M:M").Select
    Selection.NumberFormatLocal = "0_ "
    
    Range("N1").Value = "출장종료"
    Range("N2").Select
    Range("N2").Value = "=VLOOKUP(RC[-13], 출장!C[-13]:C[-3], 11, FALSE)"
    Selection.AutoFill Destination:=Range(Cells(2, 14), Cells(FindEndY(ActiveSheet, 2), 14))
    Range(Cells(2, 14), Cells(FindEndY(ActiveSheet, 2), 14)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(Cells(2, 14), Cells(FindEndY(ActiveSheet, 2), 14)).Select
    Selection.Replace What:=":", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("N:N").Select
    Selection.NumberFormatLocal = "0_ "
        
    Range("O1").Value = "적발"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-13]-RC[-2] >= 0, IF(RC[-1]-RC[-13] >= 0, ""O"", """"), """")"
    Selection.AutoFill Destination:=Range(Cells(2, 15), Cells(FindEndY(ActiveSheet, 2), 15))
    Range(Cells(2, 15), Cells(FindEndY(ActiveSheet, 2), 15)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(FindEndY(ActiveSheet, 1), findEndX(ActiveSheet, 1))).AutoFilter Field:=15, Criteria1:="O"
    
    Range("A1").Select
    
End Sub

