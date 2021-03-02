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

Public Function valueIsInArray(val As Variant, arr As Variant) As Boolean
    Dim rivi As Long
    valueIsInArray = False
    If Not IsArray(arr) Then Exit Function
    For rivi = LBound(arr) To UBound(arr)
        If arr(rivi) = val Then
            valueIsInArray = True
            Exit Function
        End If
    Next rivi
    valueIsInArray = False
End Function
    
Public Function arrMatch(ByVal arrname As Variant, ByVal value As Variant, Optional col As Long = 1)

    Dim rivi As Long

    For rivi = 1 To UBound(arrname)

        If arrname(rivi, col) = value Then
            arrMatch = rivi
            Exit Function
        End If

    Next rivi

    arrMatch = -1

End Function
    
Public Function CharCount(OrigString As String, _
                          Chars As String, Optional CaseSensitive As Boolean = False) _
                          As Long

'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string

'PARAMETERS:
'OrigString: String to Search in
'Chars: Character(s) to search for
'CaseSensitive (Optional): Do a case sensitive search
'Defaults to false

'RETURNS:
'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2
''**********************************************

    Dim lLen As Long
    Dim lCharLen As Long
    Dim lAns As Long
    Dim sInput As String
    Dim sChar As String
    Dim lCtr As Long
    Dim lEndOfLoop As Long
    Dim bytCompareType As Byte

    sInput = OrigString
    If sInput = vbNullString Then Exit Function
    lLen = Len(sInput)
    lCharLen = Len(Chars)
    lEndOfLoop = (lLen - lCharLen) + 1
    bytCompareType = IIf(CaseSensitive, vbBinaryCompare, _
                         vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid$(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then _
           lAns = lAns + 1
    Next

    CharCount = lAns

End Function
