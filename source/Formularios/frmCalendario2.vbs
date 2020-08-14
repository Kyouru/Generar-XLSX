


Option Explicit
Dim bCmbSel As Boolean           'Flag
Dim lFirstDayInMonth As Long     'Weekday number of first day
Dim lDayPos As Long              'Day position in date
Dim lMonthPos As Long            'Month position in date
Dim sMonth As String             'Month name
Dim sDateFormat As String        'The date format
Dim datFirstDay As Date          'The first date
Dim datLastDay As Date           'The second date

Private bancoSaldos() As String

Private Sub UserForm_Initialize()
'This procedure executes before
'the userform opens.
Dim ctl As Control               'Userform control variable
Dim lCount As Long               'Counter
Dim InputLblEvt As clLabelClass2 'Temporary class

On Error GoTo ErrorHandle

'The collections colLabelEvent and colLabels
'are declared in Module1.
'colLabelEvent is a collection of classes,
'clLabelClasses, that control the event
'driven action, when a date label is clicked.
'colLabels is a collection of the date labels
'used for identifying labels, setting their
'properties and more.
Set colLabelEvent = New Collection
Set colLabels = New Collection

'Loop through the date labels in Frame1
'and add them to the collections.
For Each ctl In Frame1.Controls
   'If the control element is a label
   If TypeOf ctl Is MSForms.Label Then
      'Make a new instance of the clLabel class
      Set InputLblEvt = New clLabelClass2
      
      'and assign it to this Label
      Set InputLblEvt.InputLabel = ctl
      
      'which we add to the collection, colLabelEvent.
      'Any click event on a label (day) in Frame1
      'will now be handled by the class,
      'because it declares:
      'Public WithEvents InputLabel As MSForms.Label
      'That way we avoid writing click events for
      'every label.
      colLabelEvent.Add InputLblEvt
      
      'and to the colLabels collection
      colLabels.Add ctl, ctl.Name
   End If
Next

'We have no use for InputLblEvent anymore
'and set it to Nothing to save memory.
Set InputLblEvt = Nothing

'Add month names to the month combobox.
'By using the VBA function MonthName it
'will automatically be in the user's
'language as defined in the country
'settings.
For lCount = 1 To 12
   With cmbMonth
      .AddItem MonthName(lCount)
   End With
Next

'Add years to the years combo box. VBA doesn't
'handle older years than 1900.
For lCount = 1900 To Year(Now) + 100
   With cmbYear
      .AddItem lCount
   End With
Next

'Weekday labels to local settings (first day of the week) and language.
'If for instance the country is the USA and the language is English,
'the first day of the week will be Sunday, and the labels from left to
'right will say: "SU" "MO" "TU" "WE" "TH" "FR" "SA"
'The VBA function StrConv(String,1) converts to upper case.
lblDay1.Caption = StrConv(Left(WeekdayName(1, , vbUseSystemDayOfWeek), 2), 1)
lblDay2.Caption = StrConv(Left(WeekdayName(2, , vbUseSystemDayOfWeek), 2), 1)
lblDay3.Caption = StrConv(Left(WeekdayName(3, , vbUseSystemDayOfWeek), 2), 1)
lblDay4.Caption = StrConv(Left(WeekdayName(4, , vbUseSystemDayOfWeek), 2), 1)
lblDay5.Caption = StrConv(Left(WeekdayName(5, , vbUseSystemDayOfWeek), 2), 1)
lblDay6.Caption = StrConv(Left(WeekdayName(6, , vbUseSystemDayOfWeek), 2), 1)
lblDay7.Caption = StrConv(Left(WeekdayName(7, , vbUseSystemDayOfWeek), 2), 1)

'Tag the labels. The tags are used by clLabelClass to check,
'if a date is in the selected month, the previous or next.
With colLabels
   For lCount = 1 To .count
      .Item(lCount).Tag = lCount
   Next
End With

'The LabelCaptions procedure will arrange
'the calendar's look depending on month and year.
LabelCaptions Month(Now), Year(Now)

'Find the system settings for sequence of day,
'and month.
lDayPos = Day("01-02-03")
lMonthPos = Month("01-02-03")

Exit Sub
ErrorHandle:
MsgBox Err.Description
End Sub
Sub LabelCaptions(lMonth As Long, lYear As Long)
Dim lCount As Long            'Counter
Dim lNumber As Long           'Counter
Dim lMonthPrev As Long        'Previous month
Dim lDaysPrev As Long         'Days in previous month
Dim lYearPrev As Long         'Previous year

'Get the month name from the month number
sMonth = MonthName(lMonth)

'Save month number in variable
lSelMonth = lMonth

'Save year in variable
lSelYear = lYear

'Save month and year first date
If bSecondDate = False Then
    lSelMonth1 = lSelMonth
    lSelYear1 = lSelYear
End If

'Prepare for getting days in previous month
Select Case lMonth
   Case 2 To 11
      lMonthPrev = lMonth - 1
      lYearPrev = lYear
   Case 1
      lMonthPrev = 12
      lYearPrev = lYear - 1
   Case 12
      lMonthPrev = 11
      lYearPrev = lYear
End Select
   
'Days in month
lDays = DaysInMonth(lMonth, lYear)
'Days in previous month
lDaysPrev = DaysInMonth(lMonthPrev, lYearPrev)

'If it is Jan. 1900 the
'back button is disabled.
If lSelYear >= 1900 And lSelMonth > 1 Then
   lblBack.Enabled = True
ElseIf lSelYear = 1900 And lSelMonth = 1 Then
   lblBack.Enabled = False
End If

'If this wasn't started by a selection
'in one of the combo boxes (month, year).
If bCmbSel = False Then
   cmbMonth.Text = sMonth
   cmbYear.Text = lYear
End If

'Find the first date in the month.
lFirstDayInMonth = DateSerial(lSelYear, lSelMonth, 1)

'Find the weekday number using local settings for
'first day of the week. We want to know if it is a
'Monday etc. for putting the first day of the month
'in the right weekday position.
'The first day of a week varies from country to country.
'In USA it is Sunday, in Denmark it is Monday.
'So we use vbUseSystemDayOfWeek to get the local settings.
lFirstDayInMonth = Weekday(lFirstDayInMonth, vbUseSystemDayOfWeek)

If lFirstDayInMonth = 1 Then
   lStartPos = 8
Else
   lStartPos = lFirstDayInMonth
End If

'Days from previous month if the
'first day in the month is not a monday.
lNumber = lDaysPrev + 1
For lCount = lStartPos - 1 To 1 Step -1
   lNumber = lNumber - 1
   With colLabels.Item(lCount)
      .Caption = lNumber
      .ForeColor = &HE0E0E0
   End With
Next

'The labels/buttons for the days of the month.
lNumber = 0
For lCount = lStartPos To lDays + lStartPos - 1
   lNumber = lNumber + 1
   With colLabels.Item(lCount)
      .Caption = lNumber
      .ForeColor = &H80000012
   End With
Next

'The days (labels) in next month
lNumber = 0
For lCount = lDays + lStartPos To 42
   lNumber = lNumber + 1
   With colLabels.Item(lCount)
      .Caption = lNumber
      .ForeColor = &HE0E0E0
   End With
Next

End Sub
Function DaysInMonth(lMonth As Long, lYear As Long) As Long

'Number of days in month
Select Case lMonth
   Case 1, 3, 5, 7, 8, 10, 12
      DaysInMonth = 31
   Case 2
      'Leap year?
      If IsDate("29/2/" & lYear) = False Then
         DaysInMonth = 28
      Else
         DaysInMonth = 29
      End If
   Case Else
      DaysInMonth = 30
End Select

End Function

Private Sub lblBack_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   lblBack.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub lblBack_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   lblBack.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub lblForward_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   lblForward.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub lblForward_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   lblForward.SpecialEffect = fmSpecialEffectFlat
End Sub
Sub lblForward_Click()

'If the user clicked forward
'arrow, we display the next month
If lSelMonth < 12 Then
   lSelMonth = lSelMonth + 1
Else
   lSelMonth = 1
   lSelYear = lSelYear + 1
End If

If Len(sActiveDay) > 0 Then
   'Make the previously selected look normal
   With colLabels.Item(sActiveDay)
      .BorderColor = &H8000000E
      .BorderStyle = fmBorderStyleNone
   End With
End If

'Update the calendar's look
LabelCaptions lSelMonth, lSelYear

End Sub
Sub lblBack_Click()

'If the user clicked the back arrow,
'we display the previous month.
If lSelMonth > 1 Then
   lSelMonth = lSelMonth - 1
Else
   lSelMonth = 12
   lSelYear = lSelYear - 1
End If

If Len(sActiveDay) > 0 Then
   'Make the previously selected look normal
   With colLabels.Item(sActiveDay)
      .BorderColor = &H8000000E
      .BorderStyle = fmBorderStyleNone
   End With
End If

'Update the calendar's look
LabelCaptions lSelMonth, lSelYear

End Sub
Private Sub cmbMonth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'If the month combo box is activated directly
bCmbSel = True
End Sub
Private Sub cmbMonth_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'If the month combo box is activated directly
bCmbSel = True
End Sub
Private Sub cmbMonth_Change()
Dim lOldMonth As Long

If bCmbSel Then
   'The month written by the user must match
   'one on the list.
   If cmbMonth.MatchFound = False Then Exit Sub
   
   lOldMonth = lSelMonth
   lSelMonth = Month(DateValue("01 " & cmbMonth.Text & " 2015"))
   If lSelMonth <> lOldMonth Then
      LabelCaptions lSelMonth, lSelYear
   End If
   bCmbSel = False
   If Len(sActiveDay) > 0 Then
      'Make the previously selected look normal
      colLabels.Item(sActiveDay).SpecialEffect = fmSpecialEffectFlat
   End If

End If
End Sub
Private Sub cmbMonth_AfterUpdate()
'The tricky user will paste a value in
'the cmbMonth's text. The value will be
'disregarded, because it doesn't match
'a value on the list, and if he leaves
'the combo, we reinsert the last
'selected month name.

If cmbMonth.MatchFound = False Then
   MsgBox "The month name must match one on the list."
   cmbMonth.Text = MonthName(lSelMonth)
End If

End Sub

Private Sub cmbYear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'If the year combo box is activated directly
bCmbSel = True
End Sub
Private Sub cmbYear_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'If the year combo box is activated directly
bCmbSel = True
End Sub
Private Sub cmbYear_Change()
Dim lOldYear As Long

If bCmbSel Then
   lOldYear = lSelYear
   If Val(cmbYear.Text) < 1900 Then
      cmbYear.Text = lSelYear
      bCmbSel = False
      Exit Sub
   End If
   lSelYear = Year("01 " & MonthName(lSelMonth) & " " & cmbYear.Text)
   'Call LabelCaptions
   If lSelYear <> lOldYear Then
      LabelCaptions lSelMonth, lSelYear
   End If
   bCmbSel = False
   If Len(sActiveDay) > 0 Then
      'Make the previously selected look normal
      colLabels.Item(sActiveDay).SpecialEffect = fmSpecialEffectFlat
   End If
End If

End Sub
Sub FillFirstDay()

'Unhide label
lblFirstDate.Visible = True
lblStartDate.Visible = True

'Insert first date.
datFirstDay = ReturnDate(lFirstDay, lSelMonth, lSelYear)
lblStartDate.Caption = Format(datFirstDay, sDateFormat)

'Enable command button for finding next day
cmdNext.Enabled = True

End Sub
Sub FillSecondDay()

lblLastDate.Visible = True
lblStopDate.Visible = True
datLastDay = ReturnDate(lFirstDay, lSelMonth, lSelYear)
lblStopDate.Caption = Format(datLastDay, sDateFormat)

cmdNext.Enabled = False

'Enable OK-button
cmdOK.Enabled = True

End Sub
Function ReturnDate(ByVal lDay As Long, ByVal lMonth As Long, ByVal lYear As Long) As Date
'Returns the date with day, month and year in
'the sequence defined by the system's settings.

If lDayPos = 1 And lMonthPos = 2 Then
   ReturnDate = lDay & "/" & lMonth & "/" & lYear
   Exit Function
ElseIf lDayPos = 2 And lMonthPos = 1 Then
   ReturnDate = lMonth & "/" & lDay & "/" & lYear
   Exit Function
ElseIf lDayPos = 3 And lMonthPos = 2 Then
   ReturnDate = lYear & "/" & lMonth & "/" & lDay
   Exit Function
ElseIf lDayPos = 2 And lMonthPos = 3 Then
   ReturnDate = lYear & "/" & lDay & "/" & lMonth
   Exit Function
ElseIf lDayPos = 1 And lMonthPos = 3 Then
   ReturnDate = lDay & "/" & lYear & "/" & lMonth
   Exit Function
ElseIf lMonthPos = 1 And lDayPos = 3 Then
   ReturnDate = lMonth & "/" & lYear & "/" & lDay
End If

End Function

Private Sub cmdNext_Click()
'The look of the first selected date
'(label) is set to normal.
With colLabels.Item(sActiveDay)
   .BackColor = &H8000000E
   .SpecialEffect = fmSpecialEffectFlat
End With

'Set flag for handling second date
bSecondDate = True

colLabels.Item(sActiveDay).BorderStyle = fmBorderStyleNone
cmdNext.Enabled = False
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'If the user clicks the cross in the upper right corner
If CloseMode = 0 Then cmdCancel_Click
End Sub
Private Sub cmdCancel_Click()

Set colLabelEvent = Nothing
Set colLabels = Nothing
bSecondDate = False
sActiveDay = Empty
lFirstDay = 0
        [TMP_FECHA] = ""
        [TMP_FECHA2] = ""
Unload Me

End Sub

Private Sub cmdOK_Click()
'datFirstDay and datLastDay are the
'two dates selected, and if you want
'to use them elsewhere, they should
'not be declared in this procedure,
'but e.g. as Public variables in
'Module1.

'Put your code here. For instance you can
'find the difference between the 2 dates,
'or you can copy the dates to cells.

If datFirstDay > 0 And datLastDay > 0 Then
    If datLastDay >= datFirstDay Then
        [TMP_FECHA] = fechaDateStr(datFirstDay)
        [TMP_FECHA2] = fechaDateStr(datLastDay)
        Unload Me
        
    Else
        MsgBox "Segunda Fecha es Inferior"
    End If
End If
    bSecondDate = False
End Sub

Private Sub PrintArray(Data As Variant, Cl As Range)
    Cl.Resize(UBound(Data, 2), UBound(Data, 1)) = Application.Transpose(Data)
End Sub
