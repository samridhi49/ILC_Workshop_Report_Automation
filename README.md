# ILC Workshop Report Automation

## Charlotte
You will have to change the path of the final slideshow

```
 Sub CreateSlides()
     Dim pptApp As Object
     Dim pptPres As Object
     Dim pptSlide As Object
     
     ' Create PowerPoint application
     Set pptApp = CreateObject("PowerPoint.Application")
     pptApp.Visible = True
     
     ' Create a new presentation
     Set pptPres = pptApp.Presentations.Add
     
     ' Add slides
     Set pptSlide = pptPres.Slides.Add(1, 11) ' 11 represents the slide layout
     
     ' Slide 1
     With pptSlide
         .Shapes.Title.TextFrame.TextRange.Text = "Slide 1"
     
     End With
     
     ' Slide 2
     Set pptSlide = pptPres.Slides.Add(2, 11)
     With pptSlide
         .Shapes.Title.TextFrame.TextRange.Text = "Slide 2"
     End With
     
     ' Slide 3
     Set pptSlide = pptPres.Slides.Add(3, 11)
     With pptSlide
         .Shapes.Title.TextFrame.TextRange.Text = "Slide 3"
     End With
     
     ' Continue adding slides as needed
     
     ' Save and close the presentation
     pptPres.SaveAs "C:\Users\ces16\Downloads\presentation.pptx"
     pptPres.Close
     
     ' Quit PowerPoint application
     pptApp.Quit
     
     ' Clean up
     Set pptSlide = Nothing
     Set pptPres = Nothing
     Set pptApp = Nothing
 End Sub
```

## Saloni

## Samridhi 

https://python-pptx.readthedocs.io/en/latest/
<br>
https://support.microsoft.com/en-us/word/use-mail-merge-for-bulk-email-letters-labels-and-envelopes

**VBA**
<br>
https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office
<br>
<br>
**Documentation:** https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference

__This might be something worth exploring:__

### Steps to Connect Excel and PowerPoint
- Open VBA Editor ``alt + F11``
- Tools -> References in the top menu
- Scroll down to find `Microsoft PowerPoint 16.0 Object Library`
- Select the box next to it and click `OK`
- Insert -> Module in the top menu
- Paste the following code (from https://forum.aiprm.com/t/creating-powerpoint-presentations-with-vba-code-automation-using-chatgpt/35252):
 
  ```
  Sub CreateSlides()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    
    ' Use generic Object instead of Excel.Worksheet
    Dim xlApp As Object
    Dim xlWB As Object
    Dim ws As Object
    
    Dim lastRow As Long
    Dim i As Long
    Dim slideIndex As Long
    
    ' Look for an already open instance of Excel
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    ' If Excel isn't open, alert user and exit
    If xlApp Is Nothing Then
        MsgBox "Please open your Excel file first!", vbCritical
        Exit Sub
    End If
    
    ' Reference the active worksheet in the open Excel window
    Set ws = xlApp.ActiveSheet
    
    ' Find the last row with data in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(-4162).Row ' -4162 is the raw value for xlUp
    
    ' Create PowerPoint application connection
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    slideIndex = 1
    
    ' Loop through rows starting from row 1 (change to 2 if you have headers)
    For i = 1 To lastRow
        ' Add a new slide (11 represents title/object layout)
        Set pptSlide = pptPres.Slides.Add(slideIndex, 11)
        
        ' Populate slide text from Columns A and B
        With pptSlide
            If .Shapes.HasTitle Then
                .Shapes.Title.TextFrame.TextRange.Text = CStr(ws.Cells(i, "A").Value)
            End If
            
            If .Shapes.Count >= 2 Then
                .Shapes(2).TextFrame.TextRange.Text = CStr(ws.Cells(i, "B").Value)
            End If
        End With
        
        slideIndex = slideIndex + 1
    Next i
    
    ' Save and close the presentation
    pptPres.SaveAs "C:\Users\sv49\Downloads\presentation.pptx"
    pptPres.Close
    
    ' Quit PowerPoint application
    ' pptApp.Quit
    
    ' Clean up
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    Set ws = Nothing
    Set xlApp = Nothing
  End Sub

- Go back to the Excel sheet
- Add text in Column A and Column B starting at row 2
- Press ``alt + F8``
- Select the macro
- Click ``Run`` to create the slide

There's also ways to add graphs and customize the ppt based on a selected range of data but might have to look more into that.

Power Query and Power BI--didn't read too much about them but might be better alternative to VBA? Looks like it might involve some software download though.
