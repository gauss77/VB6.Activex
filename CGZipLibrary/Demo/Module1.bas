Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
  Dim sPar As String
  Dim c() As String
  
  sPar = Command$
  If sPar <> "" Then
    c = Split(sPar, "/")
    
    Dim i As Integer
    For i = 0 To UBound(c)
      Debug.Print c(i)
      
      Select Case UCase(Left(c(i), 2))
        Case Is = "B:"
          Form1.Text1(0).Text = Mid(c(i), 3, Len(c(i)))
        Case Is = "R:"
          Form1.Text1(1).Text = Mid(c(i), 3, Len(c(i)))
      End Select
    Next
  End If
  
  Form1.Show
End Sub
