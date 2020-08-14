Option Explicit

'By declaring Public WithEvents we can handle
'events "collectively". In this case it is
'the click event on a date label, and by
'doing it this way we avoid writing click
'events for each and every data label.
Public WithEvents InputLabel As MSForms.Label
Private Sub InputLabel_click()

'We change the look of the selected day
With InputLabel
   'If selected already, we exit
   If .BorderColor = vbBlue And lSelMonth1 = lSelMonth And lSelYear1 = lSelYear Then Exit Sub
   
   .BorderColor = vbBlue
   .BorderStyle = fmBorderStyleSingle
   
   'If another day was chosen before this
   'one, we make that label look normal.
   If Len(sActiveDay) > 0 Then
      If sActiveDay <> InputLabel.Name Then
         With colLabels.Item(sActiveDay)
            .BorderColor = &H8000000E
            .BorderStyle = fmBorderStyleNone
         End With
      End If
   End If
   sActiveDay = InputLabel.Name
   lFirstDay = Val(InputLabel.Caption)
End With

End Sub
