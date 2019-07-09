# ExcelVBA_ProgressBar
Creates a simple non-modal dialog in Excel
Used to effectively show the progress of a macro

## Installation
1) Download
2) In Excel VBA Editor:
  + File -> Import File ...
  + Select ProgressForm.frm
  + Optionally, repeat for ProgressExamples.bas

## Usage
```
Sub example()
    ProgressForm.start        # displays the ProgressForm
    ProgressForm.update pct   # updates the ProgressForm to pct%
    ProgressForm.done         # hides the ProgressForm
End Sub
```

## License
GNU Public License
