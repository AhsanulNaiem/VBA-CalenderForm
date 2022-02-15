VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} calenderForm 
   Caption         =   "Click in Date to pick a Date"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3405
   OleObjectBlob   =   "calenderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "calenderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub todayButton_Click()
Me.currentMonth.caption = Date
End Sub

Private Sub CommandButton1_Click()
btn_click ("CommandButton1")
End Sub
Private Sub CommandButton2_Click()
btn_click ("CommandButton2")
End Sub
Private Sub CommandButton3_Click()
btn_click ("CommandButton3")
End Sub
Private Sub CommandButton4_Click()
btn_click ("CommandButton4")
End Sub
Private Sub CommandButton5_Click()
btn_click ("CommandButton5")
End Sub
Private Sub CommandButton6_Click()
btn_click ("CommandButton6")
End Sub
Private Sub CommandButton7_Click()
btn_click ("CommandButton7")
End Sub
Private Sub CommandButton8_Click()
btn_click ("CommandButton8")
End Sub
Private Sub CommandButton9_Click()
btn_click ("CommandButton9")
End Sub
Private Sub CommandButton10_Click()
btn_click ("CommandButton10")
End Sub
Private Sub CommandButton11_Click()
btn_click ("CommandButton11")
End Sub
Private Sub CommandButton12_Click()
btn_click ("CommandButton12")
End Sub
Private Sub CommandButton13_Click()
btn_click ("CommandButton13")
End Sub
Private Sub CommandButton14_Click()
btn_click ("CommandButton14")
End Sub
Private Sub CommandButton15_Click()
btn_click ("CommandButton15")
End Sub
Private Sub CommandButton16_Click()
btn_click ("CommandButton16")
End Sub
Private Sub CommandButton17_Click()
btn_click ("CommandButton17")
End Sub
Private Sub CommandButton18_Click()
btn_click ("CommandButton18")
End Sub
Private Sub CommandButton19_Click()
btn_click ("CommandButton19")
End Sub
Private Sub CommandButton20_Click()
btn_click ("CommandButton20")
End Sub
Private Sub CommandButton21_Click()
btn_click ("CommandButton21")
End Sub
Private Sub CommandButton22_Click()
btn_click ("CommandButton22")
End Sub
Private Sub CommandButton23_Click()
btn_click ("CommandButton23")
End Sub
Private Sub CommandButton24_Click()
btn_click ("CommandButton24")
End Sub
Private Sub CommandButton25_Click()
btn_click ("CommandButton25")
End Sub
Private Sub CommandButton26_Click()
btn_click ("CommandButton26")
End Sub
Private Sub CommandButton27_Click()
btn_click ("CommandButton27")
End Sub
Private Sub CommandButton28_Click()
btn_click ("CommandButton28")
End Sub
Private Sub CommandButton29_Click()
btn_click ("CommandButton29")
End Sub
Private Sub CommandButton30_Click()
btn_click ("CommandButton30")
End Sub
Private Sub CommandButton31_Click()
btn_click ("CommandButton31")
End Sub
Private Sub CommandButton32_Click()
btn_click ("CommandButton32")
End Sub
Private Sub CommandButton33_Click()
btn_click ("CommandButton33")
End Sub
Private Sub CommandButton34_Click()
btn_click ("CommandButton34")
End Sub
Private Sub CommandButton35_Click()
btn_click ("CommandButton35")
End Sub
Private Sub CommandButton36_Click()
btn_click ("CommandButton36")
End Sub
Private Sub CommandButton37_Click()
btn_click ("CommandButton37")
End Sub
Private Sub CommandButton38_Click()
btn_click ("CommandButton38")
End Sub
Private Sub CommandButton39_Click()
btn_click ("CommandButton39")
End Sub
Private Sub CommandButton40_Click()
btn_click ("CommandButton40")
End Sub
Private Sub CommandButton41_Click()
btn_click ("CommandButton41")
End Sub
Private Sub CommandButton42_Click()
btn_click ("CommandButton42")
End Sub

Private Sub comboboxMonth_Change()
    btn_setups
    label_currentMonth
End Sub

Private Sub comboyear_Change()
    If Me.comboboxMonth.Value = "" Then Exit Sub
    btn_setups
    label_currentMonth
End Sub

Private Sub LavelMonthHigher_Click()
    On Error GoTo abc
    If Me.comboboxMonth.ListIndex = 11 Then 'if december
        Me.comboboxMonth.Value = Me.comboboxMonth.List(0) 'january
        Me.comboyear.Value = Me.comboyear.Value + 1 'next year
    Else:
        Me.comboboxMonth.Value = Me.comboboxMonth.List(Me.comboboxMonth.ListIndex + 1) 'next month in same year
    End If
    Exit Sub
abc:
     MsgBox "Reached Limit"
End Sub

Private Sub LavelMonthLower_Click()
    On Error GoTo abc
    If Me.comboboxMonth.ListIndex = 0 Then
        Me.comboboxMonth.Value = Me.comboboxMonth.List(11)
        Me.comboyear.Value = Me.comboyear.Value - 1
    Else:
        Me.comboboxMonth.Value = Me.comboboxMonth.List(Me.comboboxMonth.ListIndex - 1)
    End If
    Exit Sub
abc:
    MsgBox "Reached Limit"
End Sub


Private Sub UserForm_Initialize()
    Me.CommandButton1.Width = 23
    
'=============set up comboyear================'
    Me.comboyear.Width = Me.CommandButton1.Width * 3
    Me.comboyear.Top = Me.comboyear.Height
    Me.comboyear.Left = Me.CommandButton1.Left
    
    Dim y As Long
    For y = 1900 To 2099
        Me.comboyear.AddItem (y)
    Next
    Me.comboyear.Value = Format(Date, "yyyy")
    
'==========set up comboboxMonth==========='
    Me.comboboxMonth.Width = Me.CommandButton1.Width * 4
    Me.comboboxMonth.Left = Me.CommandButton1.Left + Me.CommandButton1.Width * 3 'setup position and width
    Me.comboboxMonth.Top = Me.comboboxMonth.Height
    
    Dim m As Integer
    For m = 1 To 12
        Me.comboboxMonth.AddItem (Format(m & "/1/2022", "mmmm"))
    Next
    Me.comboboxMonth.Value = (Format(Date, "mmmm")) 'setup default value


'=========== set up currentMonth label =============
    label_currentMonth
    Me.LavelMonthLower.Left = Me.currentMonth.Left 'setup angels position
    Me.LavelMonthLower.Top = Me.currentMonth.Top
    Me.LavelMonthHigher.Left = Me.currentMonth.Left + Me.currentMonth.Width - Me.LavelMonthHigher.Width
    Me.LavelMonthHigher.Top = Me.currentMonth.Top
'==========set placement of date buttons==========
    Dim btn As Object
    Dim caption As Integer, EndOfMont As Integer
    EndOfMont = Int(Day(Me.comboboxMonth.Value & "/1/" & Me.comboyear.Value))
    For Each btn In Me.Controls
        If Left(btn.Name, 13) = "CommandButton" Then 'make sure this is button
            With btn
                .Width = Me.CommandButton1.Width
                If (Int(Mid(.Name, 14, 2)) Mod 7) = 0 Then 'last column
                    .Left = 10 + btn.Width * ((Int(Mid(.Name, 14, 2)) Mod 7) + 7 - 1)
                Else: ' Normal date buttons
                    .Left = 10 + btn.Width * ((Int(Mid(.Name, 14, 2)) Mod 7) - 1)
                End If
            End With
        End If
    Next
    
    
'===todayButton
Me.todayButton.Left = Me.CommandButton1.Left + Me.CommandButton1.Width * 7 - Me.todayButton.Width

'===checkbox
Me.CheckBox1.Left = Me.CommandButton1.Left
'=======button caption
'==========set form width and height=========='
    Me.Width = Me.CommandButton1.Width * 8 + 10
End Sub

Private Sub btn_setups()
    Dim btn As Object
    Dim EndOfMont As Integer, firstDt As Date
    Dim smonth As String, syear As Integer
    Dim caption As Variant
    
    smonth = Me.comboboxMonth.Value
    syear = Me.comboyear.Value
    firstDt = smonth & "/1/" & syear
    EndOfMont = Day(Application.WorksheetFunction.EoMonth(firstDt, 0))
    For Each btn In Me.Controls
        With btn
            If Left(.Name, 13) = "CommandButton" Then 'make sure it is button
                If Int(Mid(.Name, 14, 2)) < 43 Then 'make sure buttons are not weakdays button, set dates caption
                    caption = Int(Mid(.Name, 14, 2)) - Int(Weekday(Me.comboboxMonth & "/1/" & Me.comboyear, vbSaturday)) + 1
                    If caption > 0 And caption < (EndOfMont + 1) Then
                        .caption = caption
                        .BackColor = 16777215
                    ElseIf caption < 1 Then 'prev month
                        .caption = Day(firstDt + caption - 1)
                        .BackColor = 15527148
                    Else: 'next month
                        .caption = caption - EndOfMont
                        .BackColor = 15527148
                    End If
                End If
            End If
        End With
    Next
End Sub

Sub label_currentMonth()
    Me.currentMonth.Left = 10
    Me.currentMonth.Width = Me.CommandButton1.Width * 7 + 1
    Me.currentMonth.Top = Me.comboyear.Top + Me.comboyear.Height + 2
    Me.currentMonth.caption = Me.comboboxMonth.Value & "-" & Me.comboyear.Value
End Sub

Function getDate() As Variant
    label_currentMonth
    Me.Show
    If Me.CheckBox1.Value = True Then
        If Me.currentMonth.caption <> "" Then
            getDate = Me.currentMonth.caption & " " & Time
        Else:
            getDate = ""
            End If
    Else:
        getDate = Me.currentMonth.caption
    End If
End Function

Sub btn_click(btnName As String)
    Dim sdate As Date 'selected date
    Dim endOfMonth As Integer, caption As Integer, firstOfMotnh As Date
    
    firstOfMotnh = (Me.comboboxMonth.Value & "/1/" & Me.comboyear.Value) 'first day of month
    endOfMonth = Int(Day(Application.WorksheetFunction.EoMonth(firstOfMotnh, 0))) 'last day of month
    caption = Int(Mid(Me.Controls(btnName).Name, 14, 2)) - Int(Weekday(Me.comboboxMonth & "/1/" & Me.comboyear, vbSaturday)) + 1
    
    If caption < 1 Then
        sdate = Me.comboboxMonth.List(Me.comboboxMonth.ListIndex - 1) & "/" & Me.Controls(btnName).caption & "/" & Me.comboyear.Value
    ElseIf caption > endOfMonth Then
        sdate = Me.comboboxMonth.List(Me.comboboxMonth.ListIndex + 1) & "/" & Me.Controls(btnName).caption & "/" & Me.comboyear.Value
    Else:
        sdate = Me.comboboxMonth.Value & "/" & Me.Controls(btnName).caption & "/" & Me.comboyear.Value
    End If
    Me.currentMonth.caption = sdate
    Me.Hide
End Sub

Private Sub UserForm_Terminate()
Me.currentMonth.caption = ""
End Sub
