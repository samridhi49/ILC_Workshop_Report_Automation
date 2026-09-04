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

XML Code for Workshop Template Slide Master layout
```
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg>
<p:bgRef idx="1001">
<a:schemeClr val="bg1"/>
</p:bgRef>
</p:bg>
<p:spTree>
<p:nvGrpSpPr>
<p:cNvPr id="1" name=""/>
<p:cNvGrpSpPr/>
<p:nvPr/>
</p:nvGrpSpPr>
<p:grpSpPr>
<a:xfrm>
<a:off x="0" y="0"/>
<a:ext cx="0" cy="0"/>
<a:chOff x="0" y="0"/>
<a:chExt cx="0" cy="0"/>
</a:xfrm>
</p:grpSpPr>
<p:sp>
<p:nvSpPr>
<p:cNvPr id="2" name="Title Placeholder 1">
<a:extLst>
<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{1E3D4D1F-EB46-E2FA-E56A-D50D5F25C6B4}"/>
</a:ext>
</a:extLst>
</p:cNvPr>
<p:cNvSpPr>
<a:spLocks noGrp="1"/>
</p:cNvSpPr>
<p:nvPr>
<p:ph type="title"/>
</p:nvPr>
</p:nvSpPr>
<p:spPr>
<a:xfrm>
<a:off x="838200" y="365126"/>
<a:ext cx="10515600" cy="1325563"/>
</a:xfrm>
<a:prstGeom prst="rect">
<a:avLst/>
</a:prstGeom>
</p:spPr>
<p:txBody>
<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr">
<a:normAutofit/>
</a:bodyPr>
<a:lstStyle/>
<a:p>
<a:r>
<a:rPr lang="en-US"/>
<a:t>Click to edit Master title style</a:t>
</a:r>
</a:p>
</p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr>
<p:cNvPr id="3" name="Text Placeholder 2">
<a:extLst>
<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6777A0AC-0685-7E03-9867-2C172EFBCA8E}"/>
</a:ext>
</a:extLst>
</p:cNvPr>
<p:cNvSpPr>
<a:spLocks noGrp="1"/>
</p:cNvSpPr>
<p:nvPr>
<p:ph type="body" idx="1"/>
</p:nvPr>
</p:nvSpPr>
<p:spPr>
<a:xfrm>
<a:off x="838200" y="1825625"/>
<a:ext cx="10515600" cy="4351338"/>
</a:xfrm>
<a:prstGeom prst="rect">
<a:avLst/>
</a:prstGeom>
</p:spPr>
<p:txBody>
<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0">
<a:normAutofit/>
</a:bodyPr>
<a:lstStyle/>
<a:p>
<a:pPr lvl="0"/>
<a:r>
<a:rPr lang="en-US"/>
<a:t>Click to edit Master text styles</a:t>
</a:r>
</a:p>
<a:p>
<a:pPr lvl="1"/>
<a:r>
<a:rPr lang="en-US"/>
<a:t>Second level</a:t>
</a:r>
</a:p>
<a:p>
<a:pPr lvl="2"/>
<a:r>
<a:rPr lang="en-US"/>
<a:t>Third level</a:t>
</a:r>
</a:p>
<a:p>
<a:pPr lvl="3"/>
<a:r>
<a:rPr lang="en-US"/>
<a:t>Fourth level</a:t>
</a:r>
</a:p>
<a:p>
<a:pPr lvl="4"/>
<a:r>
<a:rPr lang="en-US"/>
<a:t>Fifth level</a:t>
</a:r>
</a:p>
</p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr>
<p:cNvPr id="4" name="Date Placeholder 3">
<a:extLst>
<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{F211305A-54B1-3C84-A6CB-390DD5ADD74E}"/>
</a:ext>
</a:extLst>
</p:cNvPr>
<p:cNvSpPr>
<a:spLocks noGrp="1"/>
</p:cNvSpPr>
<p:nvPr>
<p:ph type="dt" sz="half" idx="2"/>
</p:nvPr>
</p:nvSpPr>
<p:spPr>
<a:xfrm>
<a:off x="838200" y="6356350"/>
<a:ext cx="2743200" cy="365125"/>
</a:xfrm>
<a:prstGeom prst="rect">
<a:avLst/>
</a:prstGeom>
</p:spPr>
<p:txBody>
<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>
<a:lstStyle>
<a:lvl1pPr algn="l">
<a:defRPr sz="1200">
<a:solidFill>
<a:schemeClr val="tx1">
<a:tint val="82000"/>
</a:schemeClr>
</a:solidFill>
</a:defRPr>
</a:lvl1pPr>
</a:lstStyle>
<a:p>
<a:fld id="{59EF4C3E-BA27-45B4-BD12-089335B5238F}" type="datetimeFigureOut">
<a:rPr lang="en-US" smtClean="0"/>
<a:t>9/4/2026</a:t>
</a:fld>
<a:endParaRPr lang="en-US"/>
</a:p>
</p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr>
<p:cNvPr id="5" name="Footer Placeholder 4">
<a:extLst>
<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6EB18DD4-55CE-681F-416B-862429E4E6B3}"/>
</a:ext>
</a:extLst>
</p:cNvPr>
<p:cNvSpPr>
<a:spLocks noGrp="1"/>
</p:cNvSpPr>
<p:nvPr>
<p:ph type="ftr" sz="quarter" idx="3"/>
</p:nvPr>
</p:nvSpPr>
<p:spPr>
<a:xfrm>
<a:off x="4038600" y="6356350"/>
<a:ext cx="4114800" cy="365125"/>
</a:xfrm>
<a:prstGeom prst="rect">
<a:avLst/>
</a:prstGeom>
</p:spPr>
<p:txBody>
<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>
<a:lstStyle>
<a:lvl1pPr algn="ctr">
<a:defRPr sz="1200">
<a:solidFill>
<a:schemeClr val="tx1">
<a:tint val="82000"/>
</a:schemeClr>
</a:solidFill>
</a:defRPr>
</a:lvl1pPr>
</a:lstStyle>
<a:p>
<a:endParaRPr lang="en-US"/>
</a:p>
</p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr>
<p:cNvPr id="6" name="Slide Number Placeholder 5">
<a:extLst>
<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{C82EF878-0CC8-F762-7FD4-CAA85D19F760}"/>
</a:ext>
</a:extLst>
</p:cNvPr>
<p:cNvSpPr>
<a:spLocks noGrp="1"/>
</p:cNvSpPr>
<p:nvPr>
<p:ph type="sldNum" sz="quarter" idx="4"/>
</p:nvPr>
</p:nvSpPr>
<p:spPr>
<a:xfrm>
<a:off x="8610600" y="6356350"/>
<a:ext cx="2743200" cy="365125"/>
</a:xfrm>
<a:prstGeom prst="rect">
<a:avLst/>
</a:prstGeom>
</p:spPr>
<p:txBody>
<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>
<a:lstStyle>
<a:lvl1pPr algn="r">
<a:defRPr sz="1200">
<a:solidFill>
<a:schemeClr val="tx1">
<a:tint val="82000"/>
</a:schemeClr>
</a:solidFill>
</a:defRPr>
</a:lvl1pPr>
</a:lstStyle>
<a:p>
<a:fld id="{60D42A0F-D5B0-4D7C-9E14-0317AEF46BEC}" type="slidenum">
<a:rPr lang="en-US" smtClean="0"/>
<a:t>‹#›</a:t>
</a:fld>
<a:endParaRPr lang="en-US"/>
</a:p>
</p:txBody>
</p:sp>
</p:spTree>
<p:extLst>
<p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="2018536124"/>
</p:ext>
</p:extLst>
</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst>
<p:sldLayoutId id="2147483655" r:id="rId1"/>
</p:sldLayoutIdLst>
<p:txStyles>
<p:titleStyle>
<a:lvl1pPr algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPct val="0"/>
</a:spcBef>
<a:buNone/>
<a:defRPr sz="4400" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mj-lt"/>
<a:ea typeface="+mj-ea"/>
<a:cs typeface="+mj-cs"/>
</a:defRPr>
</a:lvl1pPr>
</p:titleStyle>
<p:bodyStyle>
<a:lvl1pPr marL="228596" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="1000"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="2800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl1pPr>
<a:lvl2pPr marL="685788" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="2400" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl2pPr>
<a:lvl3pPr marL="1142980" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="2000" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl3pPr>
<a:lvl4pPr marL="1600172" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl4pPr>
<a:lvl5pPr marL="2057364" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl5pPr>
<a:lvl6pPr marL="2514556" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl6pPr>
<a:lvl7pPr marL="2971748" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl7pPr>
<a:lvl8pPr marL="3428940" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl8pPr>
<a:lvl9pPr marL="3886132" indent="-228596" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc>
<a:spcPct val="90000"/>
</a:lnSpc>
<a:spcBef>
<a:spcPts val="501"/>
</a:spcBef>
<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
<a:buChar char="•"/>
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl9pPr>
</p:bodyStyle>
<p:otherStyle>
<a:defPPr>
<a:defRPr lang="en-US"/>
</a:defPPr>
<a:lvl1pPr marL="0" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl1pPr>
<a:lvl2pPr marL="457192" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl2pPr>
<a:lvl3pPr marL="914384" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl3pPr>
<a:lvl4pPr marL="1371576" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl4pPr>
<a:lvl5pPr marL="1828768" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl5pPr>
<a:lvl6pPr marL="2285960" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl6pPr>
<a:lvl7pPr marL="2743152" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl7pPr>
<a:lvl8pPr marL="3200344" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl8pPr>
<a:lvl9pPr marL="3657536" algn="l" defTabSz="914384" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200">
<a:solidFill>
<a:schemeClr val="tx1"/>
</a:solidFill>
<a:latin typeface="+mn-lt"/>
<a:ea typeface="+mn-ea"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl9pPr>
</p:otherStyle>
</p:txStyles>
</p:sldMaster>
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
