Attribute VB_Name = "txtDataB"
Option Explicit
Type Entry
Author As String
Quote As String
End Type
Public TotalEntrys As Long
Public Entrys() As Entry



Function LoadFile() As String
Open "quotes.txt" For Binary As #1
LoadFile = Input$(LOF(1), 1)
Close #1
End Function
Sub LoadCates()
Form1.Combo1.Clear
    Dim A1 As String, Str As String, A2() As String, a4 As Long, A3() As String
    
    A1 = LoadFile
    
    A2 = Split(A1, "<##QUOTE##>")
    TotalEntrys = UBound(A2) - 1
    For a4 = 0 To UBound(A2) - 1
    A3 = Split(LCase(A2(a4)), "<blockquote>")
    If InStr(1, Str, Replace(A3(0), " ", "_")) > 0 Then GoTo OK
    Form1.Combo1.AddItem A3(0)
    Str = Str & " " & Replace(A3(0), " ", "_")
    
OK:
    
    Next
    Str = vbNullString
    End Sub
    'Add Data SubRoutine

'Load List in the Var Entrys
Public Sub LoadList(Cate As String)
Dim A1 As String, Str As String, A2() As String, a4 As Long, A3() As String
    
    A1 = LoadFile
    A2 = Split(A1, "<##QUOTE##>")
    ReDim Entrys(0)
    For a4 = 0 To UBound(A2) - 1
    A3 = Split(LCase(A2(a4)), "<blockquote>")
    If A3(0) = Cate Then
        ReDim Preserve Entrys(UBound(Entrys) + 1)
        Entrys(UBound(Entrys)).Author = A3(0)
        Entrys(UBound(Entrys)).Quote = A3(1)
        End If
        Next

Exit Sub

err:
    MsgBox "CANNOT LOAD LIST" & vbCrLf & vbCrLf & "Error : " & err.Description
    
End Sub

