VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ten_skoroszyt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim shapeTlo As Shape
    Dim shapePrzycisk As Shape
    'Ustawienie arkusza
    Set ws = ThisWorkbook.Sheets("tetris")
On Error Resume Next
    Set shapeTlo = ws.Shapes("Tlo")
    Set shapePrzycisk = ws.Shapes("Rectangle: Rounded Corners 1")
On Error GoTo 0
'Dopasowanie grafiki do rozmiaru arkusza
    If Not shapeTlo Is Nothing Then
With ActiveWindow.VisibleRange
    shapeTlo.LockAspectRatio = msoFalse
    shapeTlo.Top = .Top
    shapeTlo.Left = .Left
    shapeTlo.Width = .Width
    shapeTlo.Height = .Height
End With
'Dopasowanie przycisku
    If Not shapePrzycisk Is Nothing Then
    shapePrzycisk.Left = shapeTlo.Left + (shapeTlo.Width / 2) - (shapePrzycisk.Width / 2)
    shapePrzycisk.Top = shapeTlo.Top + (shapeTlo.Height / 2) - (shapePrzycisk.Height / 2) + 150
    End If
End If
End Sub
