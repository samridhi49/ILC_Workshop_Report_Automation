# ILC Workshop Report Automation

## Charlotte


## Saloni

## Samridhi 

https://python-pptx.readthedocs.io/en/latest/

**VBA**
<br>
https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office

<br>

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
        .Shapes(2).TextFrame.TextRange.Text = "Insert content here"
    End With
    
    ' Slide 2
    Set pptSlide = pptPres.Slides.Add(2, 11)
    With pptSlide
        .Shapes.Title.TextFrame.TextRange.Text = "Slide 2"
        .Shapes(2).TextFrame.TextRange.Text = "Insert content here"
    End With
    
    ' Slide 3
    Set pptSlide = pptPres.Slides.Add(3, 11)
    With pptSlide
        .Shapes.Title.TextFrame.TextRange.Text = "Slide 3"
        .Shapes(2).TextFrame.TextRange.Text = "Insert content here"
    End With
    
    ' Continue adding slides as needed
    
    ' Save and close the presentation
    pptPres.SaveAs "C:\Path\to\save\presentation.pptx"
    pptPres.Close
    
    ' Quit PowerPoint application
    pptApp.Quit
    
    ' Clean up
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
  End Sub 
- Go back to the Excel sheet
- Add text in Column A and Column B starting at row 2
- Press ``alt + F8``
- Select the macro
- Click ``Run`` to create the slide

There's also ways to add graphs and customize the ppt based on a selected range of data but might have to look more into that.
