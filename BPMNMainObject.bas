Attribute VB_Name = "Module1"


'To draw rectangle : activity, pool, swimlane, sub-process sign, grouping
Public Sub DrawRectangle(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As fRectangles, _
    ByVal strText As String, ByVal dblSizeFactor As Double, _
    ByVal dblPositionX As Double, ByVal dblPositionY As Double, _
    Optional dblCustomLength, Optional dblCustomHeight)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case fRectangles.fStandard 'full rounded rectangle for standard task/ activity
            dblSizeLength = dblCustomLength * dblSizeFactor
            dblSizeHeight = dblCustomHeight * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0.05658
            intTransparency = 0
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False

        Case fRectangles.fStandardDash 'full dash rounded rectangle for event sub-process activity
            dblSizeLength = dblCustomLength * dblSizeFactor
            dblSizeHeight = dblCustomHeight * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSysDash
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0.05658
            intTransparency = 0
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False

        Case fRectangles.fThick 'full bold rounded rectangle for call activity
            dblSizeLength = dblCustomLength * dblSizeFactor
            dblSizeHeight = dblCustomHeight * dblSizeFactor
            dblWeight = OUTLINE_BOLD_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0.05658
            intTransparency = 0
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False
        
        Case fRectangles.fThinOuter 'full rounded rectangle for transactional activity
            dblSizeLength = dblCustomLength * dblSizeFactor
            dblSizeHeight = dblCustomHeight * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0.05658
            intTransparency = 0
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False

        Case fRectangles.fThinInner 'less rounded rectangle for transactional activity
            dblSizeLength = (dblCustomLength - 2) * dblSizeFactor
            dblSizeHeight = (dblCustomHeight - 2) * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0.035
            intTransparency = 1
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False
            
        Case fRectangles.fSmall 'small rectangle for sub-process activity, usually attached plus sign into it
            dblSizeLength = 5 * dblSizeFactor
            dblSizeHeight = 5 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0
            intTransparency = 1
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False
       
        Case fRectangles.fGroup 'full dash dot rounded rectangle for grouping activities
            dblSizeLength = dblCustomLength * dblSizeFactor
            dblSizeHeight = dblCustomHeight * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineDashDot
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0.05658
            intTransparency = 1
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False
        
        Case Else 'full rounded rectangle for standard task/ activity
            dblSizeLength = dblCustomLength * dblSizeFactor
            dblSizeHeight = dblCustomHeight * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationHorizontal
            dblAdjustment = 0.05658
            intTransparency = 0
            intOutlineColor = GetRGB(ACTIVITY_OUTLINE_COLOR)
            intFillColor = GetRGB(ACTIVITY_FILL_COLOR)
            bolGradientColor = False
           
    End Select
    
    With myDocument.Shapes.AddShape(msoShapeRoundedRectangle, _
        dblPositionX - (dblSizeLength - dblWeight) / 2, _
        dblPositionY - (dblSizeHeight - dblWeight) / 2, _
        dblSizeLength - dblWeight, dblSizeHeight - dblWeight)
        
        With .line
            .Visible = msoTrue
            .Weight = dblWeight
            .DashStyle = intDashStyle
            .ForeColor.RGB = intOutlineColor
        End With
        
        With .TextFrame
            'text format
            .TextRange.Text = strText
            .TextRange.font.Size = FONT_NORMAL_SIZE * dblSizeFactor / 10
            .TextRange.font.Name = BASE_FONT_NAME
            .TextRange.font.Color.RGB = GetRGB(ACTIVITY_FONT_COLOR)
            .Orientation = intTextOrientation
            
            'margin setting
            .MarginBottom = 0
            .MarginLeft = 2 * dblSizeFactor
            .MarginRight = 2 * dblSizeFactor
            .MarginTop = 0
            
            'setting width and word wrap
            .WordWrap = True
            .TextRange.ParagraphFormat.Alignment = ppAlignCenter
            
        End With
        
        'to adjust indentation of the corners
        .Adjustments.Item(1) = dblAdjustment
        
        'to fill the rectangle
        With .Fill
            .Transparency = intTransparency
            .ForeColor.RGB = intFillColor
            If bolGradientColor Then .TwoColorGradient msoGradientHorizontal, 1
        End With
        
        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fRectangle)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & Trim(Str(dblWeight)) & ";" & _
            Trim(Str(dblSizeFactor))
            
    End With
    
End Sub
'end draw rectangle


'To draw diamond : gateway
Public Sub DrawDiamond(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As fDiamonds, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double, Optional dblRotationDegree = 0)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case fDiamonds.fStandard 'full diamond for gateway
            dblSizeLength = NORMAL_GATEWAY_WIDTH * dblSizeFactor
            dblSizeHeight = NORMAL_GATEWAY_HEIGHT * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intTransparency = 0
            intOutlineColor = GetRGB(GATEWAY_OUTLINE_COLOR)
            intFillColor = GetRGB(GATEWAY_FILL_COLOR)
        
        Case fDiamonds.fSmall 'for sequence flow
            dblSizeLength = 4 * dblSizeFactor
            dblSizeHeight = 5 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intTransparency = 0
            intOutlineColor = GetRGB(GATEWAY_OUTLINE_COLOR)
            intFillColor = GetRGB(GATEWAY_FILL_COLOR)
            
        Case Else 'full diamond for gateway
            dblSizeLength = NORMAL_GATEWAY_WIDTH * dblSizeFactor
            dblSizeHeight = NORMAL_GATEWAY_HEIGHT * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intTransparency = 0
            intOutlineColor = GetRGB(GATEWAY_OUTLINE_COLOR)
            intFillColor = GetRGB(GATEWAY_FILL_COLOR)
    End Select

    
    With myDocument.Shapes.AddShape(msoShapeDiamond, _
        dblPositionX - (dblSizeLength - dblWeight) / 2, _
        dblPositionY - (dblSizeHeight - dblWeight) / 2, _
        dblSizeLength - dblWeight, dblSizeHeight - dblWeight)
        
        With .line
            .Visible = msoTrue
            .Weight = dblWeight
            .ForeColor.RGB = intOutlineColor
        End With
                        
        With .Fill
            .Transparency = 0
            .ForeColor.RGB = intFillColor
        End With
        
        'rotate to get arrow
        .IncrementRotation dblRotationDegree

        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fDiamond)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & Trim(Str(dblWeight)) & ";" & _
            Trim(Str(dblSizeFactor))

    End With

End Sub
'end draw diamond


'To draw circle : events, event-based gateway
Public Sub DrawCircle(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As fCircles, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case fCircles.fStandard 'full circle solid for start event
            dblSize = NORMAL_EVENT_SIZE * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(START_EVENT_FILL_COLOR) 'start event color
            
        Case fCircles.fThinOuter 'full circle solid for catch-interrupt-throw
            dblSize = NORMAL_EVENT_SIZE * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(EVENT_FILL_COLOR)

        Case fCircles.fThinInner 'less solid circle, to make another circle for catch-interrupt-throw
            dblSize = (NORMAL_EVENT_SIZE - 2) * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(EVENT_FILL_COLOR)
            
        Case fCircles.fThick 'bold circle solid for end event
            dblSize = NORMAL_EVENT_SIZE * dblSizeFactor
            dblWeight = OUTLINE_BOLD_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(END_EVENT_FILL_COLOR) 'end event color
            
        Case fCircles.fSolidInner 'less circle solid for termination event
            dblSize = (NORMAL_EVENT_SIZE - 7) * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(EVENT_OUTLINE_COLOR) 'same color with outline to get solid circle
                      
        Case fCircles.fThinInnerInner 'less circle solid for exclusive event based gateway
            dblSize = (NORMAL_EVENT_SIZE - 3.5) * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(EVENT_FILL_COLOR)
            
        Case fCircles.fDashOuter 'full circle dash for non-interrupt
            dblSize = NORMAL_EVENT_SIZE * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSysDash
            intTransparency = 0
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(EVENT_FILL_COLOR)
            
        Case fCircles.fDashInner 'less circle dash, to make another circle for non-interrupt
            dblSize = (NORMAL_EVENT_SIZE - 2) * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSysDash
            intTransparency = 1
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(EVENT_FILL_COLOR)
        
        Case fCircles.fSmall 'small circle for anchor of connector
            dblSize = 4 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(BACKGROUND_COLOR)
        
        Case Else 'full circle solid for start event
            dblSize = NORMAL_EVENT_SIZE * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            intOutlineColor = GetRGB(EVENT_OUTLINE_COLOR)
            intFillColor = GetRGB(START_EVENT_OUTLINE_COLOR)
            
    End Select
    
    With myDocument.Shapes.AddShape(msoShapeOval, _
        dblPositionX - (dblSize - dblWeight) / 2, _
        dblPositionY - (dblSize - dblWeight) / 2, _
        dblSize - dblWeight, dblSize - dblWeight)
        
        With .line
            .Visible = msoTrue 'all circles have outline
            .Weight = dblWeight
            .DashStyle = intDashStyle
            .ForeColor.RGB = intOutlineColor
        End With
        

        With .Fill
            .Transparency = intTransparency
            .ForeColor.RGB = intFillColor
        End With
        
        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fCircle)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & Trim(Str(dblWeight)) & ";" & _
            Trim(Str(dblSizeFactor))

    End With

End Sub
'end draw circles


'To draw line for connector
Public Sub DrawLine(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As fLines, _
    ByVal dblSizeFactor As Double, ByVal dblBeginPositionX As Double, _
    ByVal dblBeginPositionY As Double, ByVal dblEndPositionX As Double, _
    ByVal dblEndPositionY As Double, Optional strDots = "")
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case fLines.fSolid  'solid line
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10 * 0.75
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intStyle = MsoLineStyle.msoLineSingle
        
        Case fLines.fDot 'dot lines
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10 * 0.75
            intDashStyle = MsoLineDashStyle.msoLineSysDot
            intStyle = MsoLineStyle.msoLineSingle
            
        Case fLines.fDash 'dash dot lines
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10 * 0.75
            intDashStyle = MsoLineDashStyle.msoLineDash
            intStyle = MsoLineStyle.msoLineSingle
            
        Case fLines.fDouble 'thin thin
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10 * 0.75
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intStyle = MsoLineStyle.msoLineThinThin
        
        Case Else
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10 * 0.75
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intStyle = MsoLineStyle.msoLineSingle
    End Select
    
    dblLengthX = Abs(dblEndPositionX - dblBeginPositionX)
    dblLengthY = Abs(dblEndPositionY - dblBeginPositionY)
    
    If dblLengthX = 0 And strDots = "" Then strDots = "0,1"
    If dblLengthY = 0 And strDots = "" Then strDots = "1"

    arrDots = Split(strDots, ",")
    'to parse items into activity
    dblTotalDotsX = 0
    dblTotalDotsY = 0
    
    '"0.25,2.25,-1.25,-1.25"
    For icount = LBound(arrDots) To UBound(arrDots)
        
        If icount Mod 2 = 0 Then 'for x movement
            
            With myDocument.Shapes.AddLine(dblBeginPositionX, dblBeginPositionY, _
                dblBeginPositionX + arrDots(icount) * dblLengthX, dblBeginPositionY)
        
                With .line
                    .Visible = msoTrue
                    .Weight = dblWeight
                    .DashStyle = intDashStyle
                    .Style = intStyle
                    .ForeColor.RGB = GetRGB(GATEWAY_OUTLINE_COLOR)
                End With
                
                'set the name
                .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
                    IIf(intSubIndex + icount = 0, "", "-" & Trim(Str(intSubIndex + icount)))
                .Title = Trim(Str(fShapeTypes.fLine)) & ";" & Trim(Str(dblBeginPositionX)) & ";" & _
                    Trim(Str(dblBeginPositionY)) & ";" & Trim(Str(dblBeginPositionX + arrDots(icount) * dblLengthX)) & ";" & _
                    Trim(Str(dblBeginPositionY)) & ";" & Trim(Str(dblWeight)) & ";" & _
                    Trim(Str(dblSizeFactor))

            End With
            
            dblTotalDotsX = dblTotalDotsX + arrDots(icount)
            'next beginning position x
            dblBeginPositionX = dblBeginPositionX + arrDots(icount) * dblLengthX
            
        Else 'for y movement
        
            With myDocument.Shapes.AddLine(dblBeginPositionX, dblBeginPositionY, _
                dblBeginPositionX, dblBeginPositionY + arrDots(icount) * dblLengthY)
        
                With .line
                    .Visible = msoTrue
                    .Weight = dblWeight
                    .DashStyle = intDashStyle
                    .Style = intStyle
                    .ForeColor.RGB = GetRGB(GATEWAY_OUTLINE_COLOR)
                End With
                
                'set the name
                .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
                    IIf(intSubIndex + icount = 0, "", "-" & Trim(Str(intSubIndex + icount)))
                .Title = Trim(Str(fShapeTypes.fLine)) & ";" & Trim(Str(dblBeginPositionX)) & ";" & _
                    Trim(Str(dblBeginPositionY)) & ";" & Trim(Str(dblBeginPositionX)) & ";" & _
                    Trim(Str(dblBeginPositionY + arrDots(icount) * dblLengthY)) & ";" & Trim(Str(dblWeight)) & ";" & _
                    Trim(Str(dblSizeFactor))

            End With
            
            dblTotalDotsY = dblTotalDotsY + arrDots(icount)
            'next beginning position y
            dblBeginPositionY = dblBeginPositionY + arrDots(icount) * dblLengthY

        End If
    Next

End Sub
'end draw line


'To draw plus and cross sign : event, gateway
Public Sub DrawCross(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As fCrosses, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case fCrosses.fSolidPlusSmall 'small plus sign for sub-process marker
            dblSize = 3 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 20
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            bolRotation = False
            
        Case fCrosses.fSolidPlusLarge 'solid plus sign for complex gateway, parallel gateway
            dblSize = 12 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            bolRotation = False
            
        Case fCrosses.fSolidMultiplyLarge 'solid multiply sign for complex gateway, exclusive marker gateway
            dblSize = 12 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            bolRotation = True
            
        Case fCrosses.fSolidPlus 'solid plus sign for parallel event-based gateway
            dblSize = 10 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            bolRotation = False
        
        Case fCrosses.fSolidMultiply  'solid multiply sign for throw/ end event
            dblSize = 10 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            bolRotation = True
        
        Case fCrosses.fOutlinePlus 'outlined plus sign for catch/ non-interrupting/ interrupting
            dblSize = 10 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            bolRotation = False
            
        Case fCrosses.fOutlineMultiply 'oulined multiply sign for catch/ non-interrupting/ interrupting event
            dblSize = 10 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            bolRotation = True
        
        Case Else 'solid plus sign for event
            dblSize = 12 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            bolRotation = False
    
    End Select

    With myDocument.Shapes.AddShape(msoShapeCross, _
        dblPositionX - (dblSize - dblWeight) / 2, _
        dblPositionY - (dblSize - dblWeight) / 2, _
        dblSize - dblWeight, dblSize - dblWeight)
        
        With .line
            .Visible = msoTrue
            .Weight = dblWeight
            .DashStyle = intDashStyle
            .ForeColor.RGB = GetRGB(GATEWAY_OUTLINE_COLOR)
        End With
        

        With .Fill
            .Transparency = intTransparency
            .ForeColor.RGB = GetRGB(GATEWAY_OUTLINE_COLOR)
        End With
        
        'set the thickness of the cross
        .Adjustments.Item(1) = 0.45
        
        'rotate 45 degrees to get multiply sign
        If bolRotation Then
            .IncrementRotation 45
        End If
        
        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fCross)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & Trim(Str(dblWeight)) & ";" & _
            Trim(Str(dblSizeFactor))

    End With

End Sub
'end draw cross


'To draw triangle : event, connector
Public Sub DrawTriangle(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double, Optional dblRotationDegree = 0)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case 1 'outlined triangle for event
            dblSizeLength = 8 * dblSizeFactor
            dblSizeHeight = 8 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intTransparency = 1
            
        Case 2 'solid triangle for event
            dblSizeLength = 8 * dblSizeFactor
            dblSizeHeight = 8 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intTransparency = 0
            
        Case 3 'outlined small triangle for connector
            dblSizeLength = 3.5 * dblSizeFactor
            dblSizeHeight = 3.5 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intTransparency = 1
            
        Case 4 'solid small triangle for connector
            dblSizeLength = 3.5 * dblSizeFactor
            dblSizeHeight = 3.5 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intTransparency = 0
            
        Case Else 'outlined triangle for event
            dblSizeLength = 8 * dblSizeFactor
            dblSizeHeight = 8 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intTransparency = 1
            
    End Select
    
    With myDocument.Shapes.AddShape(msoShapeIsoscelesTriangle, _
        dblPositionX - (dblSizeLength - dblWeight) / 2, _
        dblPositionY - (dblSizeHeight - dblWeight) / 2, _
        dblSizeLength - dblWeight, dblSizeHeight - dblWeight)
        
        With .line
            .Visible = msoTrue
            .Weight = dblWeight
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With

        With .Fill
            .Transparency = intTransparency
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
        'rotate to get arrow
        .IncrementRotation dblRotationDegree

        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fTriangle)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & Trim(Str(dblWeight)) & ";" & _
            Trim(Str(dblSizeFactor))

    End With
    
End Sub
'end draw triangle


'To draw pentagon : event, gateway
Public Sub DrawPentagon(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case 1 'outlined pentagon for gateway
            dblSizeLength = 9 * dblSizeFactor
            dblSizeHeight = 9 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            
        Case 2 'outlined pentagon for event
            dblSizeLength = 10 * dblSizeFactor
            dblSizeHeight = 10 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            
        Case 3 'solid pentagon for event
            dblSizeLength = 10 * dblSizeFactor
            dblSizeHeight = 10 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            
        Case Else 'outlined pentagon for gateway
            dblSizeLength = 9 * dblSizeFactor
            dblSizeHeight = 9 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            
    End Select
    
    With myDocument.Shapes.AddShape(msoShapeRegularPentagon, _
        dblPositionX - (dblSizeLength - dblWeight) / 2, _
        dblPositionY - (dblSizeHeight - dblWeight) / 2, _
        dblSizeLength - dblWeight, dblSizeHeight - dblWeight)
        
        With .line
            .Visible = msoTrue
            .Weight = dblWeight
            .DashStyle = intDashStyle
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With

        With .Fill
            .Solid
            .Transparency = intTransparency
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fPentagon)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & Trim(Str(dblWeight)) & ";" & _
            Trim(Str(dblSizeFactor))

    End With
    
End Sub
'end draw pentagon


'To draw arrow : event
Public Sub DrawArrow(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case 1 'outlined arrow for event
            dblSizeLength = 8 * dblSizeFactor
            dblSizeHeight = 8 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            
        Case 2 'solid arrow for event
            dblSizeLength = 8 * dblSizeFactor
            dblSizeHeight = 8 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
            
        Case 3 'outlined arrow for data input
            dblSizeLength = 4 * dblSizeFactor
            dblSizeHeight = 3 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
        
        Case 4 'solid arrow for data output
            dblSizeLength = 4 * dblSizeFactor
            dblSizeHeight = 3 * dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 0
        
        Case Else 'outlined arrow for event
            dblSizeLength = 8 * dblSizeFactor
            dblSizeHeight = 8 * dblSizeFactor
            dblWeight = OUTLINE_NORMAL_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTransparency = 1
            
    End Select
    
    With myDocument.Shapes.AddShape(msoShapeRightArrow, _
        dblPositionX - (dblSizeLength - dblWeight) / 2, _
        dblPositionY - (dblSizeHeight - dblWeight) / 2, _
        dblSizeLength - dblWeight, dblSizeHeight - dblWeight)
        
        With .line
            .Visible = msoTrue
            .Weight = dblWeight
            .DashStyle = intDashStyle
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With

        With .Fill
            .Transparency = intTransparency
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fArrow)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & Trim(Str(dblWeight)) & ";" & _
            Trim(Str(dblSizeFactor))

    End With
    
End Sub
'end draw arrow

'To draw caption : event, gateway, artifact
Public Sub DrawCaption(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal strText As String, ByVal dblSizeFactor As Double, _
    ByVal dblPositionX As Double, ByVal dblPositionY As Double)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case 1 'caption for gateways
            dblSizeLength = 100 * dblSizeFactor
            dblSizeHeight = 38 * dblSizeFactor
            dblWeight = 0
            intTextOrientation = msoTextOrientationHorizontal
            
        Case Else 'caption for gateways
            dblSizeLength = 100 * dblSizeFactor
            dblSizeHeight = 38 * dblSizeFactor
            dblWeight = 0
            intTextOrientation = msoTextOrientationHorizontal
    End Select
    
    'if string text is empty, do not create the shape of text
    If Trim(strText) <> "" Then
        With myDocument.Shapes.AddShape(msoShapeRoundedRectangle, _
            dblPositionX - (dblSizeLength - dblWeight) / 2, _
            dblPositionY - (dblSizeHeight - dblWeight) / 2, _
            dblSizeLength - dblWeight, dblSizeHeight - dblWeight)
            
            With .line
                .Visible = msoFalse
            End With
            
            With .TextFrame
                'text format
                .TextRange.Text = strText
                .TextRange.font.Size = FONT_NORMAL_SIZE * dblSizeFactor / 10
                .TextRange.font.Name = BASE_FONT_NAME
                .TextRange.font.Color.RGB = GetRGB(CAPTION_FONT_COLOR)
                .Orientation = intTextOrientation
                
                'margin setting
                .MarginBottom = 0
                .MarginLeft = 2 * dblSizeFactor
                .MarginRight = 2 * dblSizeFactor
                .MarginTop = 0
                
                'setting width
                .WordWrap = False
                .AutoSize = ppAutoSizeShapeToFitText
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                
            End With
            
            With .Fill
                .Transparency = 1
            End With
            
            'set the name
            .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
                IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
            .Title = Trim(Str(fShapeTypes.fObject)) & ";" & _
                Trim(Str(dblPositionX)) & ";" & Trim(Str(dblPositionY)) & ";" & _
                Trim(Str(.Width)) & ";" & Trim(Str(.Height)) & ";" & "0" & ";" & _
                Trim(Str(dblSizeFactor))
    
        End With
    End If
    
End Sub
'end draw caption


'To draw pool and lane : pool, lane
Public Sub DrawPoolLane(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal strText As String, ByVal dblSizeFactor As Double, _
    ByVal dblPositionX As Double, ByVal dblPositionY As Double, _
    ByVal dblWidth As Double, ByVal dblHeight As Double)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case 1 'full rectangle for label actor/ caption
            dblSizeLength = 12 * dblSizeFactor 'set fix 12
            dblSizeHeight = dblHeight 'in dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationUpward
            dblAdjustment = 0
            intTransparency = 1
            intOutlineColor = GetRGB(POOL_OUTLINE_COLOR)
            intFillColor = GetRGB(POOL_FILL_COLOR)
            bolGradientColor = False

        Case 2 'full rectangle for flowchart content
            dblSizeLength = dblWidth - 12 * dblSizeFactor 'adjusted by actor width
            dblSizeHeight = dblHeight 'in dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationUpward
            dblAdjustment = 0
            intTransparency = 1
            intOutlineColor = GetRGB(POOL_OUTLINE_COLOR)
            intFillColor = GetRGB(POOL_OUTLINE_COLOR)
            bolGradientColor = False
            
        Case Else 'full rectangle for label actor/ caption
            dblSizeLength = 12 * dblSizeFactor 'set fix 12
            dblSizeHeight = dblHeight '* dblSizeFactor
            dblWeight = OUTLINE_THIN_WEIGHT * dblSizeFactor / 10
            intDashStyle = MsoLineDashStyle.msoLineSolid
            intTextOrientation = msoTextOrientationUpward
            dblAdjustment = 0
            intTransparency = 1
            intOutlineColor = GetRGB(POOL_OUTLINE_COLOR)
            intFillColor = GetRGB(POOL_OUTLINE_COLOR)
            bolGradientColor = False
           
    End Select
    
    With myDocument.Shapes.AddShape(msoShapeRoundedRectangle, _
        dblPositionX - (dblSizeLength - dblWeight) / 2, _
        dblPositionY - (dblSizeHeight - dblWeight) / 2, _
        dblSizeLength - dblWeight, dblSizeHeight - dblWeight)
        
        With .line
            .Visible = msoTrue
            .Weight = dblWeight
            .DashStyle = intDashStyle
            .ForeColor.RGB = intOutlineColor
        End With
        
        With .TextFrame
            'text format
            .TextRange.Text = strText
            .TextRange.font.Size = FONT_HUGE_SIZE * dblSizeFactor / 10
            .TextRange.font.Name = BASE_FONT_NAME
            .TextRange.font.Color.RGB = GetRGB(POOL_FONT_COLOR)
            .TextRange.font.Bold = True
            .Orientation = intTextOrientation
            
            'margin setting
            .MarginBottom = 0
            .MarginLeft = 5
            .MarginRight = 5
            .MarginTop = 0
            
            'setting width and word wrap
            .WordWrap = True
            .TextRange.ParagraphFormat.Alignment = ppAlignCenter
            
        End With
        
        'to adjust indentation of the corners
        .Adjustments.Item(1) = dblAdjustment
        
        'to fill the rectangle
        With .Fill
            .Transparency = intTransparency
            .ForeColor.RGB = intFillColor
            If bolGradientColor Then .TwoColorGradient msoGradientHorizontal, 1
        End With
        
        'set the name
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex))) 'name with format i****
        .Title = Trim(Str(fShapeTypes.fPoolLane)) & ";" & _
            Trim(Str(dblPositionX)) & ";" & Trim(Str(dblPositionY)) & ";" & _
            Trim(Str(.Width)) & ";" & Trim(Str(.Height)) & ";" & _
            Trim(Str(dblWeight)) & ";" & Trim(Str(dblSizeFactor))
    End With
    
    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name)).ZOrder msoSendToBack
End Sub


