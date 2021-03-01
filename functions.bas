Function LastRow(wsheet As String, col As String) As Long
Dim ws As Worksheet
Set ws = ActiveWorkbook.Sheets(wsheet)

LastRow = ws.Cells(Rows.Count, col).End(xlUp).row

End Function

Function LastColumn(wsheet As String, row As String) As String

Dim ws As Worksheet
Set ws = ActiveWorkbook.Sheets(wsheet)

LastColumn = Split(Columns(ws.Cells(row, Columns.Count).End(xlToLeft).Column).Address(, False), ":")(1)

End Function

Function InjectDate(FilePath As String) As String

Dim FPArray() As String

FPArray = Split(FilePath, "\")

For x = 0 To UBound(FPArray)

    If x = UBound(FPArray) Then
        InjectDate = InjectDate + Format(Now(), "YYYYMMDD") + " - " + FPArray(x)
    Else
        InjectDate = InjectDate + FPArray(x) + "\"
    End If

Next x

End Function

Private Function concat(ByVal arr, Optional ByVal delim$ = " ") As String
' Purpose: build string from 2-dim array row, delimited by 2nd argument
' Note:    concatenation via JOIN needs a "flat" 1-dim array via double transposition
  concat = Join(Application.Transpose(Application.Transpose(arr)), delim)
End Function

Public Function GetLength(a As Variant) As Integer
   If IsEmpty(a) Then
      GetLength = 0
   Else
      GetLength = UBound(a) - LBound(a) + 1
   End If
End Function

Function getRange(dtaSheet As String) As Integer

'Finds the last non-blank cell on a sheet/range.
Sheets(dtaSheet).Select
Dim lRow As Long
Dim lCol As Long
    
    lRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
      'MsgBox lRow
    
    
    getRange = lRow

End Function

Public Function IEWindowFromTitle(sTitle As String) As SHDocVw.InternetExplorer

    Dim objShellWindows As New SHDocVw.ShellWindows
    Dim win As Object, rv As SHDocVw.InternetExplorer

    For Each win In objShellWindows
        If TypeName(win.document) = "HTMLDocument" Then
            If UCase(win.document.title) = UCase(sTitle) Then
                Set rv = win
                Exit For
            End If
        End If
    Next

    Set IEWindowFromTitle = rv

End Function
