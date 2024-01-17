Sub ObjectCloner()
    ActiveDocument.Unit = cdrCentimeter
    Dim currentDocument As Document
    Set currentDocument = ActiveDocument

    Dim paperWidth As Double
    Dim paperHeight As Double
    
    Dim OffsetT As Double
    Dim OffsetR As Double
    Dim OffsetB As Double
    Dim OffsetL As Double
    
    OffsetT = 0.5
    OffsetR = 1
    OffsetB = 0.5
    OffsetL = 1
    
    paperWidth = currentDocument.Pages.First.SizeWidth - (OffsetR + OffsetL)
    paperHeight = currentDocument.Pages.First.SizeHeight - (OffsetT + OffsetB)
    
    Dim selected As Shape
    Set selected = ActiveSelection.Shapes(1)
          
    Dim widthInput As Double
    Dim heightInput As Double
          
    widthInput = selected.SizeWidth
    heightInput = selected.SizeHeight
          
    Dim mataX As Integer
    Dim mataY As Integer
    mataX = 1
    mataY = 1
         
    Do While (mataX * widthInput) < paperWidth
        mataX = mataX + 1
    Loop
           
    If mataX > 1 Then
        mataX = mataX - 1
    End If
            
    Do While (mataY * heightInput) < paperHeight
        mataY = mataY + 1
    Loop
          
    If mataY > 1 Then
        mataY = mataY - 1
    End If
         
    Dim xGroup As New ShapeRange
    xGroup.Add selected
            
    Dim i As Integer
    For i = 1 To mataX - 1
        xGroup.Add selected.Duplicate
        xGroup(xGroup.Count).Move widthInput * i, 0
    Next i
            
    Dim yGroup As New ShapeRange

    Dim j As Integer
    For j = 1 To mataY
        Dim yRow As New ShapeRange
        yRow.AddRange xGroup.Duplicate
        yRow.Move 0, heightInput
        yGroup.AddRange yRow
    Next j
            
    xGroup.Delete
    yGroup.Group
End Sub
Sub GridMaker()
    ActiveDocument.Unit = cdrCentimeter
    Dim currentDocument As Document
    Set currentDocument = ActiveDocument
    
    Dim paperWidth As Double
    Dim paperHeight As Double
        
    paperWidth = currentDocument.Pages.First.SizeWidth - 0.5
    paperHeight = currentDocument.Pages.First.SizeHeight - 1
            
    Dim selected As Shape
    Set selected = ActiveSelection.Shapes(1)
            
    Dim widthInput As Double
    Dim heightInput As Double
            
    widthInput = selected.SizeWidth
    heightInput = selected.SizeHeight
            
    Dim mataX As Integer
    Dim mataY As Integer
    mataX = 1
    mataY = 1
               
    Do While (mataX * widthInput) < paperWidth
        mataX = mataX + 1
    Loop
            
    If mataX > 1 Then
        mataX = mataX - 1
    End If
            
    Do While (mataY * heightInput) < paperHeight
        mataY = mataY + 1
    Loop
            
    If mataY > 1 Then
        mataY = mataY - 1
    End If
            
    Dim rectangle As Shape
    Set rectangle = currentDocument.ActiveLayer.CreateRectangle(0, 0, widthInput, heightInput)
            
    Dim i As Integer
    Dim xGroup As New ShapeRange
            
    xGroup.Add rectangle
            
    For i = 1 To mataX - 1
        xGroup.Add rectangle.Duplicate
        xGroup(xGroup.Count).Move widthInput * i, 0
    Next i
            
    Dim yGroup As New ShapeRange
            
    Dim j As Integer
    For j = 1 To mataY
        Dim yRow As New ShapeRange
        yRow.AddRange xGroup.Duplicate
        yRow.Move 0, heightInput
        yGroup.AddRange yRow
    Next j
            
    xGroup.Delete
    yGroup.Group
End Sub
Sub CutMarker()
    ActiveDocument.Unit = cdrMillimeter
    
    Dim selected As Shape
    Set selected = ActiveSelection.Shapes(1)
    
    Dim selectedPosX As Double
    Dim selectedPosY As Double
    Dim selectedWidth As Double
    Dim selectedHeight As Double
    
    Dim cutmarks As New ShapeRange
    Dim withobject As New ShapeRange
    Dim cutmarkWeight As Double
    cutmarkWeight = 0.075
    
    withobject.Add selected
    
    selectedPosX = selected.PositionX
    selectedPosY = selected.PositionY
    selectedWidth = selected.SizeWidth
    selectedHeight = selected.SizeHeight
    
    Dim cutline1 As Shape
    Set cutline1 = ActiveLayer.CreateLineSegment(selectedPosX + 3, selectedPosY, selectedPosX, selectedPosY - 3)
    cutline1.Fill.ApplyNoFill
    cutline1.Outline.SetPropertiesEx cutmarkWeight, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 60), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse
    
    Dim cutmark1 As Curve
    Set cutmark1 = ActiveDocument.CreateCurve
    With cutmark1.CreateSubPath(selectedPosX + 3, selectedPosY)
        .AppendLineSegment selectedPosX, selectedPosY
        .AppendLineSegment selectedPosX, selectedPosY - 3
    End With
    cutline1.Curve.CopyAssign cutmark1
    
    Dim cutline2 As Shape
    Set cutline2 = ActiveLayer.CreateLineSegment((selectedPosX + selectedWidth) - 3, selectedPosY, (selectedPosX + selectedWidth), selectedPosY - 3)
    cutline2.Fill.ApplyNoFill
    cutline2.Outline.SetPropertiesEx cutmarkWeight, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 60), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse
    
    Dim cutmark2 As Curve
    Set cutmark2 = ActiveDocument.CreateCurve
    With cutmark2.CreateSubPath((selectedPosX + selectedWidth) - 3, selectedPosY)
        .AppendLineSegment (selectedPosX + selectedWidth), selectedPosY
        .AppendLineSegment (selectedPosX + selectedWidth), selectedPosY - 3
    End With
    cutline2.Curve.CopyAssign cutmark2
    
    Dim cutline3 As Shape
    Set cutline3 = ActiveLayer.CreateLineSegment((selectedPosX + selectedWidth), (selectedPosY - selectedHeight) + 3, (selectedPosX + selectedWidth) - 3, (selectedPosY - selectedHeight))
    cutline3.Fill.ApplyNoFill
    cutline3.Outline.SetPropertiesEx cutmarkWeight, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 60), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse
    
    Dim cutmark3 As Curve
    Set cutmark3 = ActiveDocument.CreateCurve
    With cutmark3.CreateSubPath((selectedPosX + selectedWidth), (selectedPosY - selectedHeight) + 3)
        .AppendLineSegment (selectedPosX + selectedWidth), (selectedPosY - selectedHeight)
        .AppendLineSegment (selectedPosX + selectedWidth) - 3, (selectedPosY - selectedHeight)
    End With
    cutline3.Curve.CopyAssign cutmark3
    
    Dim cutline4 As Shape
    Set cutline4 = ActiveLayer.CreateLineSegment(selectedPosX + 3, (selectedPosY - selectedHeight), selectedPosX, (selectedPosY - selectedHeight) - 3)
    cutline4.Fill.ApplyNoFill
    cutline4.Outline.SetPropertiesEx cutmarkWeight, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 60), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse
    
    Dim cutmark4 As Curve
    Set cutmark4 = ActiveDocument.CreateCurve
    With cutmark4.CreateSubPath(selectedPosX + 3, (selectedPosY - selectedHeight))
        .AppendLineSegment selectedPosX, (selectedPosY - selectedHeight)
        .AppendLineSegment selectedPosX, (selectedPosY - selectedHeight) + 3
    End With
    cutline4.Curve.CopyAssign cutmark4
    
    cutmarks.Add cutline1
    cutmarks.Add cutline2
    cutmarks.Add cutline3
    cutmarks.Add cutline4
    
    withobject.Add cutmarks.Group
    withobject.Group
End Sub
