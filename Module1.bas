Attribute VB_Name = "Module1"
Option Explicit

Sub RunForm()
Dim tWB As Workbook
Dim i, n  As Integer
Set tWB = ThisWorkbook
tWB.Activate
Worksheets("Sheet1").Visible = True
For i = 1 To 51
    UserForm1.state1select.AddItem Worksheets("Sheet1").Range("E" & i)
    UserForm1.state2select.AddItem Worksheets("Sheet1").Range("E" & i)
Next i

With UserForm1
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
End With

Worksheets("Sheet1").Visible = False
End Sub

Sub PopulateStates()
'YOUR CODE GOES HERE
End Sub

