Attribute VB_Name = "ProgressExamples"
Public Sub example()
' example()
' no parameters
' Provides a time based example for the ProgressForm

    Dim i As Long
    
    ' Displays the ProgressForm and initializes it
    ProgressForm.start
    
    'simply a loop to do something for showing the ProgressForm
    For i = 0 To 100 Step 6
        ' Updates the % on the ProgressForm
        ProgressForm.update i
        
        Application.Wait (Now + TimeValue("0:00:01"))
    Next
    
    ' hides the ProfessForm and cleans up
    ProgressForm.done
End Sub


Public Sub example2()
' example2()
' no parameters
' Provides a time based example for the ProgressForm with a different Caption

    Dim i As Long
    
    ' Displays the ProgressForm and initializes it
    ProgressForm.start ("Are we done yet?")
    
    For i = 0 To 100 Step 6
        ProgressForm.update i
        Application.Wait (Now + TimeValue("0:00:01"))
    Next
    ProgressForm.done
End Sub


Public Sub example3()
' example3()
' no parameters
' Provides a time based example for the ProgressForm with macro interruption

    Dim i As Long
    
    ' Displays the ProgressForm and initializes it
    Call ProgressForm.start(, True)
    
    ' or any of the following
    'ProgressForm.start , True
    'ProgressForm.start "With Interrupt", True
    'ProgressForm.start bInterrupt:=True
    
    
    For i = 0 To 100 Step 6
        ProgressForm.update i
        Application.Wait (Now + TimeValue("0:00:01"))
    Next
    ProgressForm.done
End Sub


Public Sub example4()
' example4()
' no parameters
' Provides a time based example using a label with the progress

    Dim i As Long
    ProgressForm.start
       
    For i = 0 To 5
        ' Updates the % on the ProgressForm with a label
        ProgressForm.update (i / 5 * 100), "Performing Task: " & i
        
        Application.Wait (Now + TimeValue("0:00:01"))
    Next
    
    ' hides the ProfessForm and cleans up
    ProgressForm.done
End Sub



'finally an example that does something useful
Public Sub clearBlankCells()
' clearBlankCells
' parameters: none are provided, but uses Selection to specify the range of cells to clear
'       if selection is a single cell, the entire active range is used
' Removes blank cells (cells that are empty) and blank rows (rows that have no non-empty cells)
'
    Dim r As Range              ' area to manipulate
    Dim i As Long, j As Long    ' indicies
    Dim x As Long, y As Long    ' dimensions
    Dim isBlank As Boolean      ' blank row text
    
    
    Set r = Selection
    If (r.rows.Count = 1 And r.columns.Count = 1) Then
        Set r = Range(Range("A1"), Range("A1").SpecialCells(xlCellTypeLastCell))
    End If
    
    x = r.rows.Count
    y = r.columns.Count
    
    Set r = r(1)
    
' Displays the ProgressForm and initializes it
ProgressForm.start
    
    For i = x - 1 To 0 Step -1
        isBlank = True
        For j = y - 1 To 0 Step -1
            If (IsEmpty(r.Offset(i, j))) Then
                r.Offset(i, j).Delete (xlToLeft)
            Else
                isBlank = False
            End If
        Next
        If isBlank Then
            r.Offset(i, y).EntireRow.Delete
        End If
    
' Updates the % on the ProfressForm
ProgressForm.update CLng((x - i) / x * 100)

    Next
    
' hides the ProfessForm and cleans up
ProgressForm.done

End Sub
