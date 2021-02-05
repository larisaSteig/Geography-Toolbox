VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Geography Toolbox"
   ClientHeight    =   4608
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   14832
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1



Private Sub GoButton1_Click()
Dim d As Double
If City1Label = "" Or City2Label = "" Then
    MsgBox "One or more cities are not selected. Please select two cities."
ElseIf City1Label = City2Label Then
    MsgBox "Same city is selected twice. Please select different cities."
Else

d = Distance(City1Label, State1Label, City2Label, State2Label)
'The previous line should reference the Distance function below!
MsgBox (UserForm1.City1Label & " and " & UserForm1.City2Label & " are " & FormatNumber(d, 0) & " miles apart as the crow flies.")
End If
End Sub

Private Sub GoButton2_Click()
Dim d As Double
If city1select.Text = "" Or city2select.Text = "" Then
    MsgBox "One or more cities are not selected. Please select two cities."
ElseIf city1select.Text = city2select.Text Then
    MsgBox "Same city is selected twice. Please select different cities."
Else

d = Distance(city1select.Value, state1select.Text, city2select.Value, state2select.Text)
'The previous line should reference the Distance function below!
MsgBox (UserForm1.city1select.Text & " and " & UserForm1.city2select.Text & " are " & FormatNumber(d, 0) & " miles apart as the crow flies.")
End If
End Sub

Private Sub state1select_Change()
Dim LastRow As Long
Dim i, j As Integer
UserForm1.city1select.Clear
Worksheets("Sheet1").Visible = True
Worksheets("Sheet1").Select
Range("A1").Select

For i = 1 To 857
    j = 1
    If Worksheets("Sheet1").Range("A1:A857").Cells(i, 1).Value = UserForm1.state1select.Text Then
    Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            UserForm1.city1select.AddItem ActiveCell.Offset(j, 0)
            j = j + 1
        Loop
        UserForm1.city1select.Text = ActiveCell.Offset(1, 0)
        Exit For
    End If

Next i
Worksheets("Sheet1").Visible = False

End Sub

Private Sub state2select_Change()
Dim LastRow As Long
Dim i, j As Integer
UserForm1.city2select.Clear
Worksheets("Sheet1").Visible = True
Worksheets("Sheet1").Select
Range("A1").Select

For i = 1 To 857
    j = 1
    If Range("A1:A857").Cells(i, 1).Value = UserForm1.state2select.Text Then
    Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            UserForm1.city2select.AddItem ActiveCell.Offset(j, 0)
            j = j + 1
        Loop
        UserForm1.city2select.Text = ActiveCell.Offset(1, 0)
        Exit For
    End If

Next i
Worksheets("Sheet1").Visible = False
End Sub

Private Sub SearchButton1_Click()
Dim i As Integer, j As Integer, k As Integer
Dim Cities() As String, States() As String
Dim Ans As Integer, commaposition

For i = 1 To 855
Worksheets("Sheet1").Visible = True
Worksheets("Sheet1").Select
    commaposition = InStr(UserForm1.city1input.Text, ",")

    If (UserForm1.city1input = "") Or IsNumeric(UserForm1.city1input.Value) = True Then
        Worksheets("Sheet1").Visible = False
        MsgBox "Cant be empty or number"
        Exit Sub
    End If
    If commaposition = 0 Then
        If UCase(Left(Worksheets("Sheet1").Range("A1:A855").Cells(i, 1), Len(UserForm1.city1input.Text))) = UCase(UserForm1.city1input.Text) _
        And Not UCase(Worksheets("Sheet1").Range("A1:A855").Cells(i, 1)) = Worksheets("Sheet1").Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
            ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not UCase(ActiveCell.Offset(-j, 0)) = ActiveCell.Offset(-j, 0) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j, 0)
                    Exit Do
                End If
            Loop
        End If
    Else
        If UCase(Left(Range("A1:A855").Cells(i, 1), commaposition - 1)) = UCase(Left(UserForm1.city1input.Text, commaposition - 1)) _
        And Not UCase(Range("A1:A855").Cells(i, 1)) = Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
        ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not IsEmpty(ActiveCell.Offset(-j, 0)) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j + 1, 0)
                    Exit Do
                End If
            Loop
        End If
    End If
Next i

If k = 0 Then
    Worksheets("Sheet1").Visible = False
    MsgBox "No cities were found. Please check your spelling and try again."
ElseIf k > 1 Then
    UserForm2.ComboBox1.Clear
    For j = 1 To k
        UserForm2.ComboBox1.AddItem Cities(j) & ", " & States(j)
    Next j
    UserForm2.ComboBox1.Text = Cities(1) & ", " & States(1)
    Worksheets("Sheet1").Visible = False
    UserForm2.Show
    
    If UserForm2.ComboBox1.ListIndex = -1 Then
    
    
    UserForm1.city1input = ""
    Else
    UserForm1.city1input = ""
    City1Label = Cities(UserForm2.ComboBox1.ListIndex + 1)
    State1Label = States(UserForm2.ComboBox1.ListIndex + 1)
    End If
    
    'State1Label = States(UserForm2.ComboBox1.ListIndex + 1)
Else
    Worksheets("Sheet1").Visible = False
    Ans = MsgBox("Did you mean " & Cities(1) & ", " & States(1) & "?", vbYesNo)
    If Ans = 7 Then
        Worksheets("Sheet1").Visible = False
        MsgBox "Sorry, that's the only location we could find meeting your search criterion."
        Exit Sub
    End If
    UserForm1.city1input = ""
    
    City1Label = Cities(1)
    State1Label = States(1)
End If
Worksheets("Sheet1").Visible = False


End Sub

Private Sub SearchButton2_Click()
Dim i As Integer, j As Integer, k As Integer
Dim Cities() As String, States() As String
Dim Ans As Integer, commaposition

For i = 1 To 855
Worksheets("Sheet1").Visible = True
Worksheets("Sheet1").Select
    commaposition = InStr(UserForm1.city2input.Text, ",")
    If (UserForm1.city2input = "") Or IsNumeric(UserForm1.city2input.Value) = True Then
    Worksheets("Sheet1").Visible = False
    MsgBox "Cant be empty or number"
    Exit Sub
    End If
    If commaposition = 0 Then
        If UCase(Left(Worksheets("Sheet1").Range("A1:A855").Cells(i, 1), Len(UserForm1.city2input.Text))) = UCase(UserForm1.city2input.Text) _
        And Not UCase(Worksheets("Sheet1").Range("A1:A855").Cells(i, 1)) = Worksheets("Sheet1").Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
            ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not UCase(ActiveCell.Offset(-j, 0)) = ActiveCell.Offset(-j, 0) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j, 0)
                    Exit Do
                End If
            Loop
        End If
    Else
        If UCase(Left(Range("A1:A855").Cells(i, 1), commaposition - 1)) = UCase(Left(UserForm1.city2input.Text, commaposition - 1)) _
        And Not UCase(Range("A1:A855").Cells(i, 1)) = Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
        ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not IsEmpty(ActiveCell.Offset(-j, 0)) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j + 1, 0)
                    Exit Do
                End If
            Loop
        End If
    End If
Next i

If k = 0 Then
    Worksheets("Sheet1").Visible = False
    MsgBox "No cities were found. Please check your spelling and try again."
ElseIf k > 1 Then
    UserForm2.ComboBox1.Clear
    For j = 1 To k
        UserForm2.ComboBox1.AddItem Cities(j) & ", " & States(j)
    Next j
    UserForm2.ComboBox1.Text = Cities(1) & ", " & States(1)
      Worksheets("Sheet1").Visible = False
    UserForm2.Show
  
    If UserForm2.ComboBox1.ListIndex = -1 Then
    Worksheets("Sheet1").Visible = False
    
    UserForm1.city2input = ""
    Else
    UserForm1.city2input = ""
    City2Label = Cities(UserForm2.ComboBox1.ListIndex + 1)
    State2Label = States(UserForm2.ComboBox1.ListIndex + 1)
    End If
Else
    Worksheets("Sheet1").Visible = False
    Ans = MsgBox("Did you mean " & Cities(1) & ", " & States(1) & "?", vbYesNo)
    
    If Ans = 7 Then
        MsgBox "Sorry, that's the only location we could find meeting your search criterion."
        Exit Sub
    End If
    UserForm1.city2input = ""
    City2Label = Cities(1)
    State2Label = States(1)
End If
Worksheets("Sheet1").Visible = False
End Sub

Function Distance(city1 As String, state1 As String, city2 As String, state2 As String) As Double
Dim i As Integer
Dim FindCity1, FindCity2 As Range
Dim lat1, lat2, lon1, lon2, A, b, c, x, z, Rad As Double
Worksheets("Sheet1").Visible = True
Worksheets("Sheet1").Select
x = Application.WorksheetFunction.pi() / 180

Rad = 3960
    For i = 1 To 1000
        If Worksheets("Sheet1").Range("A1:A100").Cells(i, 1) = state1 Then
            For z = i To 1000
            Set FindCity1 = Worksheets("Sheet1").Range("A1:A" & z).Find(What:=city1)
                If Not FindCity1 Is Nothing Then
                    lat1 = FindCity1.Offset(0, 1)
                    lon1 = FindCity1.Offset(0, 2)
                    Exit For
                End If
            Next z
            Exit For
         End If
    Next i
    For i = 1 To 1000
        If Worksheets("Sheet1").Range("A1:A100").Cells(i, 1) = state2 Then
            For z = i To 1000
                Set FindCity2 = Worksheets("Sheet1").Range("A1:A" & z).Find(What:=city2)
                    If Not FindCity2 Is Nothing Then
                        lat2 = FindCity2.Offset(0, 1)
                        lon2 = FindCity2.Offset(0, 2)
                        Exit For
                    End If
            Next z
            Exit For
        End If
    Next i
    
A = Cos(lat1 * x) * Cos(lat2 * x) * Cos(lon1 * x) * Cos(lon2 * x)
b = Cos(lat1 * x) * Sin(lon1 * x) * Cos(lat2 * x) * Sin(lon2 * x)
c = Sin(lat1 * x) * Sin(lat2 * x)

Distance = Application.WorksheetFunction.Acos(A + b + c) * Rad
Worksheets("Sheet1").Visible = False
End Function

Private Sub QuitButton_Click()
Unload Me
End Sub

Private Sub ResetButton_Click()
Unload UserForm1
RunForm
End Sub


Private Sub UserForm_Click()

End Sub
