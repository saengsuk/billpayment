Sub OpenWorkbook()
 
Workbooks.Open ThisWorkbook.Path & "\QB.xlsx"

End Sub

Sub CloseWorkbook()

Workbooks("QB.xlsx").Close SaveChanges:=True

End Sub

Sub ClearEntireSheet()
Application.DisplayAlerts = False

Sheets("QB").Cells.Clear
Sheets("QB").Delete
Application.DisplayAlerts = True

End Sub
Sub ClearSheet1()
With Sheets("Sheet1")
    .Rows(2 & ":" & .Rows.Count).Delete
End With
End Sub

Sub CopyfromQB()

OpenWorkbook
Workbooks("QB.xlsx").Worksheets("Sheet1").Copy _
    Before:=Workbooks("asb_datasource.xlsm").Sheets(1)
Sheets(1).Select
    Sheets(1).Name = "QB"
CloseWorkbook
    
End Sub







Sub ChangeInvFormat()
Dim LastRow As Long
Sheets("QB").Activate
Columns("H").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Columns("I:I").Select
Selection.NumberFormat = "General"
LastRow = Cells(Rows.Count, 8).End(xlUp).Row
Range("I2:I" & LastRow).Formula = "=SUBSTITUTE(RC[-1],""/"","""")"




End Sub


Sub ChangeAmountFormat()
Dim LastRow As Long
Sheets("QB").Activate

Columns("W:W").Select
Selection.NumberFormat = "General"
LastRow = Cells(Rows.Count, 8).End(xlUp).Row
Range("W2:W" & LastRow).Formula = "=ROUND(RC[-1],0)"




End Sub


Sub ChangeTermFormat()

'2020 - 2021

Columns("N:N").Select
Selection.Replace What:="1 Year/2020-21", Replacement:="'1002020", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="1st2020-21", Replacement:="'1012020", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="2nd2020-21", Replacement:="'1022020", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
Selection.Replace What:="Jap 1/2020", Replacement:="'2012020", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Jap 2/2020", Replacement:="'2022020", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


'2019 - 2020

Columns("N:N").Select
Selection.Replace What:="1 Year/2019-20", Replacement:="'1002019", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="1st2019-20", Replacement:="'1012019", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="2nd2019-20", Replacement:="'1022019", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
Selection.Replace What:="Jap 1/2019", Replacement:="'2012019", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Jap 2/2019", Replacement:="'2022019", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
      
      
 '2018-2019
 
Selection.Replace What:="1 Year/2018-19", Replacement:="'1002018", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="1st2018-19", Replacement:="'1012018", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="2nd2018-19", Replacement:="'1022018", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        
Selection.Replace What:="Jap 1/2018", Replacement:="'2012018", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Jap 2/2018", Replacement:="'2022018", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   
   '2017-2018
'Selection.Replace What:="1 Year/2017-18", Replacement:="'002017", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'Selection.Replace What:="1st2017-18", Replacement:="'012017", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'Selection.Replace What:="2nd2017-18", Replacement:="'022017", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


End Sub

Sub CopyName()
Sheets("QB").Activate

LastRow = Range("J2").End(xlDown).Row
Range("J2:J" & LastRow).Copy Sheets("Sheet1").Range("A2")

End Sub


Sub CopyInv()
Sheets("QB").Activate
LastRow = Range("I2").End(xlDown).Row
Range("I2:I" & LastRow).Copy
Sheets("Sheet1").Range("B2").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


End Sub


Sub CopyTerm()
Sheets("QB").Activate

LastRow = Range("N2").End(xlDown).Row
Range("N2:N" & LastRow).Copy Sheets("Sheet1").Range("C2")

End Sub


Sub CopyAmount()
Sheets("QB").Activate

LastRow = Range("W2").End(xlDown).Row
Range("W2:W" & LastRow).Copy
Sheets("Sheet1").Range("D2").PasteSpecial Paste:=xlPasteValues

End Sub

Sub CopyAll()
CopyName
CopyInv
CopyTerm
CopyAmount
End Sub





Sub HumanFormula()
Dim LastRow As Long
Sheets("Sheet1").Activate
LastRow = Range("A2").End(xlDown).Row
Range("E2:E" & LastRow).Formula = "=CONCATENATE(""0994002514350"","" "",""01"","" "",RC[-3],"" "",RC[-2],"" "",RC[-1],""00"")"


End Sub

Sub RawFormula()
Dim FormulaStr As String
Dim Formula2 As String
Dim LastRow As Long
'FormulaStr = "=CONCATENATE(""" & Chr(124) & "|099400251435001"",""~013"",RC[-4],""~013"",RC[-3],""~013"",IF(RC[-2]=0,0,CONCATENATE(RC[-2],""00"")))"
Formula2 = "IF((RC[-2]-INT(RC[-2]))=0,CONCATENATE(RC[-2],""00""),SUBSTITUTE(RC[-2],""."",""""))"
'FormulaStr = "=CONCATENATE(""|099400251435001"",""~013"",RC[-4],""~013"",RC[-3],""~013"",IF(RC[-2]=0,0,CONCATENATE(RC[-2],""00"")))"
FormulaStr = "=CONCATENATE(""|099400251435001"",""~013"",RC[-4],""~013"",RC[-3],""~013""," & Formula2 & ")"
'MsgBox (FormulaStr)

Sheets("Sheet1").Activate
LastRow = Range("A2").End(xlDown).Row
Range("F2:F" & LastRow).Formula = FormulaStr

End Sub


Sub EncodeFormula()
Dim LastRow As Long
Sheets("Sheet1").Activate
LastRow = Range("A2").End(xlDown).Row
Range("G2:G" & LastRow).Formula = "=code128(RC[-1],0,1)"

End Sub

Sub QrcodeFormula()
Dim FormulaStr As String
Dim Formula2 As String
Dim LastRow As Long
Sheets("Sheet1").Activate
LastRow = Range("A2").End(xlDown).Row
Formula2 = "IF((RC[-4]-INT(RC[-4]))=0,CONCATENATE(RC[-4],""00""),SUBSTITUTE(RC[-4],""."",""""))"
FormulaStr = "=CONCATENATE(""|099400251435001"",CHAR(10),RC[-6],CHAR(10),RC[-5],CHAR(10)," & Formula2 & ")"
Range("H2:H" & LastRow).Formula = FormulaStr
Range("H2:H" & LastRow).WrapText = True

End Sub

Sub Test()
Sheets("Sheet1").Activate
If Range("D21") = Int(Range("D21")) Then
  'x is an Integer!'
  MsgBox ("hello")
Else
  'x is not an Integer!'
End If
End Sub

Sub Test2()
Sheets("QB").Activate
LastRow = Cells(Rows.Count, 8).End(xlUp).Row
MsgBox (LastRow)
LastRow2 = Range("H2").End(xlDown).Row
MsgBox (LastRow2)
End Sub



Sub StartOVer()
ClearEntireSheet
ClearSheet1
CopyfromQB
ChangeInvFormat
ChangeAmountFormat
ChangeTermFormat
CopyAll
HumanFormula
RawFormula
EncodeFormula
QrcodeFormula
End Sub

