Attribute VB_Name = "Module1"
'To draw escalation : event
Public Sub DrawEscalation(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
        
    Select Case intOption
        Case 1 'outlined escalation
            intSizeAdjustment = 300
            intCombineStyle = 0
        
        Case 2 'solid escalation
            intSizeAdjustment = 300
            intCombineStyle = 1
            
        Case Else 'outlined escalation
            intSizeAdjustment = 300
            intCombineStyle = 0
    End Select
    
    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=0 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-4.23344999999995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-1.0003 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-1.44974999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-1.6352 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=1.44974999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-2.6356 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=4.23344999999995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-1.7178 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=3.32114999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-0.9177 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=2.29905000000008 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=0 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=1.38665000000015 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=0.878500000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=2.33555000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=1.7571 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=3.28455000000008 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=2.6356 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=4.23344999999995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=1.7403 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=1.41644999999994 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=0.895200000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-1.41634999999997 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=0 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-4.23344999999995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=1.92000000000014E-02 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-2.48164999999995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=0.510700000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-0.934950000000072 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=0.974600000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=0.620249999999942 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=1.4661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=2.16685000000007 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=0.9838 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=1.64585000000011 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=0.5015 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=1.12495000000013 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=1.92000000000014E-02 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=0.603949999999941 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-0.388599999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=0.961549999999988 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-1.4476 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=2.32954999999993 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-1.3797 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=2.03255000000013 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-0.858499999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=0.545149999999921 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-0.513699999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-0.997949999999946 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=1.92000000000014E-02 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-2.48164999999995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    
    If intCombineStyle = 0 Then
        ActivePresentation.Slides(intSlideNumber).Shapes. _
            Range(Array(shp1.Name, shp2.Name)).MergeShapes msoMergeCombine
    Else
        ActivePresentation.Slides(intSlideNumber).Shapes. _
            Range(Array(shp1.Name, shp2.Name)).MergeShapes msoMergeUnion
    End If
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Solid
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With

End Sub
'end draw escalation


'To draw hand (manual task) : activity
Sub DrawHand(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-78.9843499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-601.75055 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-101.28345 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-601.74355 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-121.93725 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-593.49265 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-139.55075 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-582.13535 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-139.58005 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-582.11775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-139.60735 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-582.09825 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-234.30115 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-520.77875 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-558.47475 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-295.80815 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-629.73435 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-246.89705 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-629.73635 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-246.89515 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-629.73825 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-246.89515 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-694.10945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-202.69895 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-737.47115 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-135.99135 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-762.40035 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-56.95955 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-762.40425 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-56.9439500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-762.40815 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-56.9302500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-788.36055 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=25.58525 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-785.03025 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=115.47305 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-784.77925 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=190.75525 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-784.77735 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=190.77485 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-784.77735 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=190.79435 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-784.52255 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=247.47995 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-783.17115 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=296.94955 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-768.46875 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=356.94275 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-768.46775 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=356.94775 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-768.46585 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=356.95315 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-768.46485 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=356.95845 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-747.09195 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=444.76045 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-705.69685 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=508.29445 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-647.23435 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=547.24355 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-588.77455 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=586.19085 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-516.62025 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=600.01115 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-435.66015 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=600.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-143.37675 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=601.35835 * dblSizeFactor / 300 + dblPositionY, _
            X2:=149.08095 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=601.75085 * dblSizeFactor / 300 + dblPositionY, _
            X3:=441.52155 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=600.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=441.54695 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=441.57225 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=473.78095 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600.04995 * dblSizeFactor / 300 + dblPositionY, _
            X2:=504.53165 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=586.50875 * dblSizeFactor / 300 + dblPositionY, _
            X3:=523.70705 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=562.58345 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=542.88105 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=538.65995 * dblSizeFactor / 300 + dblPositionY, _
            X2:=550.72375 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=507.71845 * dblSizeFactor / 300 + dblPositionY, _
            X3:=551.13675 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=473.77285 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=551.42765 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=451.00485 * dblSizeFactor / 300 + dblPositionY, _
            X2:=548.22605 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=429.48585 * dblSizeFactor / 300 + dblPositionY, _
            X3:=540.83015 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=410.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=573.20515 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=410.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=604.73625 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=410.26115 * dblSizeFactor / 300 + dblPositionY, _
            X2:=633.63885 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=401.27715 * dblSizeFactor / 300 + dblPositionY, _
            X3:=654.35745 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=382.34515 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=675.06905 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=363.41955 * dblSizeFactor / 300 + dblPositionY, _
            X2:=685.66185 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=337.72425 * dblSizeFactor / 300 + dblPositionY, _
            X3:=691.16605 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=310.66545 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=691.16605 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=310.65765 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=698.84835 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=273.01785 * dblSizeFactor / 300 + dblPositionY, _
            X2:=695.22045 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=235.85115 * dblSizeFactor / 300 + dblPositionY, _
            X3:=680.25005 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=204.66155 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=713.19565 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=202.00005 * dblSizeFactor / 300 + dblPositionY, _
            X2:=742.28325 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=189.55875 * dblSizeFactor / 300 + dblPositionY, _
            X3:=760.99025 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=166.75525 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=782.01345 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=141.12845 * dblSizeFactor / 300 + dblPositionY, _
            X2:=788.13765 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=108.54985 * dblSizeFactor / 300 + dblPositionY, _
            X3:=788.24615 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=74.03455 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=788.36055 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=38.0011499999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=779.61695 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=3.61414999999988 * dblSizeFactor / 300 + dblPositionY, _
            X3:=759.43365 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-23.7310500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=739.25045 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-51.0762500000001 * dblSizeFactor / 300 + dblPositionY, _
            X2:=705.66225 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-69.6870500000001 * dblSizeFactor / 300 + dblPositionY, _
            X3:=667.86725 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-69.74085 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=667.85745 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-69.74085 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=667.84575 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-69.74085 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=654.08655 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-69.7522500000001 * dblSizeFactor / 300 + dblPositionY, _
            X2:=641.68605 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-69.7325500000001 * dblSizeFactor / 300 + dblPositionY, _
            X3:=626.06445 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-69.74085 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=633.58115 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-88.48495 * dblSizeFactor / 300 + dblPositionY, _
            X2:=636.88555 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-109.56735 * dblSizeFactor / 300 + dblPositionY, _
            X3:=636.81645 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-131.74665 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=636.81645 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-131.80335 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=636.81645 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-131.85995 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=636.58865 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-167.63235 * dblSizeFactor / 300 + dblPositionY, _
            X2:=627.56555 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-201.81345 * dblSizeFactor / 300 + dblPositionY, _
            X3:=607.28125 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-228.98695 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=586.99735 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-256.15985 * dblSizeFactor / 300 + dblPositionY, _
            X2:=553.47315 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-274.74015 * dblSizeFactor / 300 + dblPositionY, _
            X3:=515.69535 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-274.74085 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=291.27555 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-275.62635 * dblSizeFactor / 300 + dblPositionY, _
            X2:=69.44395 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-272.58385 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-119.20705 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-273.53765 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-109.37645 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-283.42265 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-99.82945 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-293.06265 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-89.33595 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-303.49865 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-56.85465 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-335.80165 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-25.63675 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-366.24135 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-9.50974999999994 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-387.10995 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=29.04535 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-436.73235 * dblSizeFactor / 300 + dblPositionY, _
            X2:=34.79805 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-503.96165 * dblSizeFactor / 300 + dblPositionY, _
            X3:=2.33984999999996 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-553.41465 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-14.77295 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-579.55795 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-39.86735 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-597.24655 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-66.81055 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-600.92835 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-70.17945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-601.38865 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-73.51845 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-601.64895 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-76.8222499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-601.72715 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-76.8222499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-601.72515 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-77.54495 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-601.74225 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-78.26505 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-601.75085 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-78.9843499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-601.75055 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-78.43745 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-531.72325 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-77.66945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-531.71465 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-76.95215 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-531.66155 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-76.28905 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-531.57085 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-70.98375 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-530.84595 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-65.30345 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-528.95065 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-56.21875 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-515.06505 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-56.20115 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-515.03965 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-56.18555 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-515.01425 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-43.49375 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-495.68785 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-44.05425 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-456.70895 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-64.80855 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-430.02205 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-64.84375 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-429.97715 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-64.8788499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-429.93225 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-72.42535 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-420.16245 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-106.33275 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-385.31635 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-138.69725 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-353.12945 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-171.06175 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-320.94265 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-201.83115 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-291.23395 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-217.05465 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-272.68805 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-231.61155 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-254.95415 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-226.27045 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-240.04075 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-222.32225 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-230.70565 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-218.37405 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-221.37055 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-213.85755 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-210.46895 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-193.72455 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-205.68615 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-187.69665 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-204.25415 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-187.36475 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-204.66635 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-185.43555 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-204.53375 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-183.50635 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-204.40125 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-181.51215 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-204.30985 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-179.16795 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-204.22515 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-174.47945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-204.05585 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-168.48055 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-203.92775 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-161.07415 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-203.81895 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=65.3907499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-202.43975 * dblSizeFactor / 300 + dblPositionY, _
            X2:=289.34835 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-205.05955 * dblSizeFactor / 300 + dblPositionY, _
            X3:=515.63475 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-204.73885 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=515.66605 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-204.73885 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=515.69725 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-204.73885 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=532.90915 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-204.73885 * dblSizeFactor / 300 + dblPositionY, _
            X2:=542.25425 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-199.08185 * dblSizeFactor / 300 + dblPositionY, _
            X3:=551.18945 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-187.11195 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=560.11875 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-175.15005 * dblSizeFactor / 300 + dblPositionY, _
            X2:=566.65795 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-155.29865 * dblSizeFactor / 300 + dblPositionY, _
            X3:=566.81835 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-131.45955 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=566.88435 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-106.72935 * dblSizeFactor / 300 + dblPositionY, _
            X2:=560.96175 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-91.6396500000001 * dblSizeFactor / 300 + dblPositionY, _
            X3:=553.76175 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-83.3267500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=546.56175 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-75.01395 * dblSizeFactor / 300 + dblPositionY, _
            X2:=536.91645 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-69.8061500000001 * dblSizeFactor / 300 + dblPositionY, _
            X3:=516.28715 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-69.7388500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.63675 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-69.7388500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.63675 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=0.26114999999993 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=516.37895 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=0.26114999999993 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=516.42775 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=0.26114999999993 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=582.90615 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=0.267349999999965 * dblSizeFactor / 300 + dblPositionY, _
            X2:=617.96855 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=0.219549999999913 * dblSizeFactor / 300 + dblPositionY, _
            X3:=667.79105 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=0.26114999999993 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=685.09835 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=0.290649999999914 * dblSizeFactor / 300 + dblPositionY, _
            X2:=694.31985 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=5.92274999999995 * dblSizeFactor / 300 + dblPositionY, _
            X3:=703.11525 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=17.8392499999999 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=711.91455 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=29.76095 * dblSizeFactor / 300 + dblPositionY, _
            X2:=718.32475 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=49.6660499999999 * dblSizeFactor / 300 + dblPositionY, _
            X3:=718.24805 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=73.8138499999999 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=718.24805 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=73.8158499999998 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=718.16705 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=99.63855 * dblSizeFactor / 300 + dblPositionY, _
            X2:=712.90195 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=115.01165 * dblSizeFactor / 300 + dblPositionY, _
            X3:=706.87305 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=122.36075 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=700.84425 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=129.70985 * dblSizeFactor / 300 + dblPositionY, _
            X2:=692.23495 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=135.09275 * dblSizeFactor / 300 + dblPositionY, _
            X3:=667.39065 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=135.26315 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=658.96085 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=135.32095 * dblSizeFactor / 300 + dblPositionY, _
            X2:=598.58275 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=135.27675 * dblSizeFactor / 300 + dblPositionY, _
            X3:=573.52735 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=135.30605 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=573.13035 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=135.30105 * dblSizeFactor / 300 + dblPositionY, _
            X2:=572.74175 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=135.26435 * dblSizeFactor / 300 + dblPositionY, _
            X3:=572.34375 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=135.26315 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=392.11725 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=134.41255 * dblSizeFactor / 300 + dblPositionY, _
            X2:=211.88915 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=135.26115 * dblSizeFactor / 300 + dblPositionY, _
            X3:=31.66215 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=135.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.63675 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=135.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.61135 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=205.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.63675 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=205.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=31.6797499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=205.26115 * dblSizeFactor / 300 + dblPositionY, _
            X2:=414.15345 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=205.49975 * dblSizeFactor / 300 + dblPositionY, _
            X3:=572.96685 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=205.30805 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=592.66685 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=205.58715 * dblSizeFactor / 300 + dblPositionY, _
            X2:=604.47775 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=213.51865 * dblSizeFactor / 300 + dblPositionY, _
            X3:=613.90045 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=228.90955 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=623.46395 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=244.53055 * dblSizeFactor / 300 + dblPositionY, _
            X2:=628.22365 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=269.01905 * dblSizeFactor / 300 + dblPositionY, _
            X3:=622.58015 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=296.66545 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=622.58015 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=296.67915 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=622.58015 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=296.69275 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=618.84375 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=315.07325 * dblSizeFactor / 300 + dblPositionY, _
            X2:=613.16755 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=325.16575 * dblSizeFactor / 300 + dblPositionY, _
            X3:=607.14455 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=330.66935 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=592.81005 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=340.26115 * dblSizeFactor / 300 + dblPositionY, _
            X2:=573.21095 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=340.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=392.68745 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=340.17505 * dblSizeFactor / 300 + dblPositionY, _
            X2:=212.16645 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=339.80895 * dblSizeFactor / 300 + dblPositionY, _
            X3:=31.6426500000001 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=340.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.55085 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=340.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.6426500000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=410.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=31.73445 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=410.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=441.04105 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=410.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=441.05665 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=410.26115 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=455.76375 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=410.35695 * dblSizeFactor / 300 + dblPositionY, _
            X2:=462.30345 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=414.26955 * dblSizeFactor / 300 + dblPositionY, _
            X3:=469.01765 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=423.23775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=475.73185 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=432.20585 * dblSizeFactor / 300 + dblPositionY, _
            X2:=481.45705 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=448.88015 * dblSizeFactor / 300 + dblPositionY, _
            X3:=481.14845 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=472.89005 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=481.14845 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=472.90375 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=481.14845 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=472.91545 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=480.86435 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=496.32705 * dblSizeFactor / 300 + dblPositionY, _
            X2:=475.19675 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=511.18705 * dblSizeFactor / 300 + dblPositionY, _
            X3:=469.09185 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=518.80415 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=462.98705 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=526.42115 * dblSizeFactor / 300 + dblPositionY, _
            X2:=456.68685 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=530.16105 * dblSizeFactor / 300 + dblPositionY, _
            X3:=441.11915 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=530.26315 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=148.99575 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=531.75095 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-143.23685 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=531.35985 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-435.39055 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=530.26315 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-435.40425 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=530.26315 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-435.41595 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=530.26315 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-508.46895 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=530.04005 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-565.87185 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=517.33155 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-608.41595 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=488.98775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-650.96645 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=460.64375 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-682.02365 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=416.11925 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-700.45505 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=340.39005 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-700.46285 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=340.36075 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-700.46875 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=340.33145 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-713.33895 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=287.83885 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-714.52795 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=246.32115 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-714.77925 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=190.52095 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-714.77925 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=190.47995 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-715.03495 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=113.90935 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-716.50375 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=30.43035 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-695.63275 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-35.92835 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-674.24095 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-103.72955 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-639.64505 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-155.18055 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-590.12105 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-189.18415 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-590.11715 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-189.18615 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-517.99265 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-238.69085 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-191.88815 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-464.84495 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-101.61715 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-523.30525 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-101.61325 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-523.30725 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-93.13715 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-528.77195 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-86.11395 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-531.09595 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-80.89445 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-531.60995 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-80.02445 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-531.69565 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-79.20555 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-531.73175 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-78.43745 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-531.72325 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    
    ActivePresentation.Slides(intSlideNumber). _
        Shapes.Range(Array(shp1.Name, shp2.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw hand


'To draw user : activity
Sub DrawUser(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=0 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-625 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-177.2774 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-625 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-304.0974 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-488.7193 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-304.4531 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-333.4863 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-304.4531 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-333.4414 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-304.4531 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-333.3984 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-304.4419 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-286.4536 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-291.7391 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-236.9086 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-271.8926 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-192.2207 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-257.5623 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-159.9536 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-239.7456 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-130.2892 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-218.0352 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-106.5254 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-347.0641 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-62.3821 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-498.2278 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=10.2699 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-574.6328 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=153.5293 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-578.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=161.25 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-578.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=625 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=578.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=625 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=578.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=161.25 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=574.6328 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=153.5293 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=499.3267 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=12.3304000000001 * dblSizeFactor / 300 + dblPositionY, _
            X2:=351.434 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-60.2205 * dblSizeFactor / 300 + dblPositionY, _
            X3:=223.6523 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-104.5469 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=286.0317 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-167.9795 * dblSizeFactor / 300 + dblPositionY, _
            X2:=304.4331 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-249.8527 * dblSizeFactor / 300 + dblPositionY, _
            X3:=304.4531 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-333.3984 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=304.4531 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-333.4414 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=304.4531 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-333.4863 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=304.0974 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-488.7193 * dblSizeFactor / 300 + dblPositionY, _
            X2:=177.2774 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-625 * dblSizeFactor / 300 + dblPositionY, _
            X3:=0 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-625 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-121.041 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-459.7012 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-112.8466 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-459.6806 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-103.8499 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-459.4069 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-93.9316 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-458.8223 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-14.9086 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-454.1646 * dblSizeFactor / 300 + dblPositionY, _
            X2:=11.6873000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-439.9401 * dblSizeFactor / 300 + dblPositionY, _
            X3:=32.1309 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-426.4492 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=52.5744 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-412.9584 * dblSizeFactor / 300 + dblPositionY, _
            X2:=66.9862000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-400.116 * dblSizeFactor / 300 + dblPositionY, _
            X3:=121.0137 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-398.541 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=121.0297 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-398.541 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=121.0477 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-398.541 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=163.1494 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-400.1158 * dblSizeFactor / 300 + dblPositionY, _
            X2:=183.3968 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-407.6221 * dblSizeFactor / 300 + dblPositionY, _
            X3:=197.9168 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-416.125 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=203.8005 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-419.5705 * dblSizeFactor / 300 + dblPositionY, _
            X2:=208.7409 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-423.1522 * dblSizeFactor / 300 + dblPositionY, _
            X3:=213.8035 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-426.5723 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=227.1849 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-397.89 * dblSizeFactor / 300 + dblPositionY, _
            X2:=234.3706 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-366.1833 * dblSizeFactor / 300 + dblPositionY, _
            X3:=234.452 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-333.3691 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=234.424 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-239.9292 * dblSizeFactor / 300 + dblPositionY, _
            X2:=218.3601 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-174.4926 * dblSizeFactor / 300 + dblPositionY, _
            X3:=132.6844 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-120.9453 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=141.0653 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-57.7754 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=158.898 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-52.36 * dblSizeFactor / 300 + dblPositionY, _
            X2:=177.3311 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-46.4181 * dblSizeFactor / 300 + dblPositionY, _
            X3:=196.0399 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-39.9023 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=198.6527 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-29.0022 * dblSizeFactor / 300 + dblPositionY, _
            X2:=201.6715 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-14.706 * dblSizeFactor / 300 + dblPositionY, _
            X3:=203.702 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=0.492200000000025 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=205.8291 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=16.4146 * dblSizeFactor / 300 + dblPositionY, _
            X2:=206.6657 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=32.981 * dblSizeFactor / 300 + dblPositionY, _
            X3:=205.3543 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=45.4199000000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=200.1809 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=64.5737999999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=199.5028 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=65.252 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=155.9559 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=108.7988 * dblSizeFactor / 300 + dblPositionY, _
            X2:=78.6564000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=134.1953 * dblSizeFactor / 300 + dblPositionY, _
            X3:=0.250800000000027 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=134.1953 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-78.1548 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=134.1953 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-155.4543 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=108.7988 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-199.0011 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=65.252 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-199.6793 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=64.5737999999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-203.5413 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=57.8588 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-204.8527 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=45.4199000000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-206.1641 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=32.981 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-205.3275 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=16.4146 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-203.2004 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=0.492200000000025 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-201.1587 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-14.7899 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-198.1159 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-29.1704 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-195.4933 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-40.0918 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-176.9714 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-46.5329 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-158.7236 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-52.4125 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-141.0636 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-57.7754 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-136.0676 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-125.4629 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-140.1356 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-130.68 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-144.2805 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-134.1361 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-149.2277 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-137.8477 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-168.349 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-152.193 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-191.5377 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-183.7514 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-207.9172 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-220.6328 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-224.2908 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-257.501 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-234.4351 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-299.6423 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-234.4504 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-333.3769 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-234.3486 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-373.8213 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-223.4646 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-412.5879 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-203.4738 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-446.0742 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-199.8969 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-447.4056 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-196.1251 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-448.8127 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-191.8859 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-450.207 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-176.9781 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-455.1103 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-156.5494 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-459.7904 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-121.0402 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-459.7012 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-272.3887 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-10.1445 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-272.4516 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-9.6857 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-272.5225 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-9.23789999999997 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-272.584 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-8.77729999999997 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-275.1709 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=10.5862 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-276.6909 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=31.6632 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-274.4668 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=52.7598 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=-266.6761 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=96.5699999999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-248.4981 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=114.748 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-186.4503 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=176.7958 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-92.5471 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=204.1953 * dblSizeFactor / 300 + dblPositionY, _
            X3:=0.25 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=204.1953 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=93.0471 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=204.1953 * dblSizeFactor / 300 + dblPositionY, _
            X2:=186.9503 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=176.7958 * dblSizeFactor / 300 + dblPositionY, _
            X3:=248.998 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=114.748 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=267.1761 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=96.5699999999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=272.7427 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=73.8563999999999 * dblSizeFactor / 300 + dblPositionY, _
            X3:=274.9668 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=52.7598 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=277.1909 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=31.6631 * dblSizeFactor / 300 + dblPositionY, _
            X2:=275.6709 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=10.5862 * dblSizeFactor / 300 + dblPositionY, _
            X3:=273.084 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-8.77729999999997 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=273.033 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-9.15899999999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=272.9738 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-9.52970000000005 * dblSizeFactor / 300 + dblPositionY, _
            X3:=272.9219 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-9.91020000000003 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=367.2556 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=31.4588 * dblSizeFactor / 300 + dblPositionY, _
            X2:=458.4367 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=90.7053000000001 * dblSizeFactor / 300 + dblPositionY, _
            X3:=508.75 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=179.1641 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=508.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=555 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=341.25 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=555 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=341.25 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=290 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=271.25 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=290 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=271.25 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=555 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-272.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=555 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-272.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=290 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-342.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=290 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-342.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=555 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-508.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=555 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-508.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=179.1641 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-458.3419 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=90.5385000000001 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-366.912 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=31.2349 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-272.3887 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-10.1445 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw user


'To draw data source : artifact
Sub DrawDataSource(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-660.915 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-157.49255 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-660.915 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-315.04595 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-647.9045 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-438.44155 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-620.9932 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-500.13935 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-607.5375 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-553.27475 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-590.8123 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-594.93175 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-568.8604 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-635.05595 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-547.7162 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-667.70705 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-520.445 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-677.77555 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-481.0342 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-678.46885 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-478.9757 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-678.96905 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-476.8572 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-679.26965 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-474.7061 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-680.35565 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-469.4365 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-679.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-465.8662 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-680.13465 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-150.8249 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-679.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=148.9633 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-679.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=473.0049 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-678.89855 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=476.501 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-670.26555 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=518.3927 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-636.58885 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=546.9084 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-594.93175 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=568.8604 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-553.27475 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=590.8123 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-500.13935 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=607.5355 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-438.44155 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=620.9912 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-315.04595 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=647.9026 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-157.49255 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=660.915 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=660.915 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=158.13285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=660.915 * dblSizeFactor / 100 + dblPositionY, _
            X2:=315.68635 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=647.9026 * dblSizeFactor / 100 + dblPositionY, _
            X3:=439.08185 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=620.9912 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=500.77965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=607.5355 * dblSizeFactor / 100 + dblPositionY, _
            X2:=553.91515 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=590.8123 * dblSizeFactor / 100 + dblPositionY, _
            X3:=595.57215 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=568.8604 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=670.90785 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=518.3927 * dblSizeFactor / 100 + dblPositionY, _
            X2:=679.54085 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=476.501 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=680.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=473.0049 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=680.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=163.1467 * dblSizeFactor / 100 + dblPositionY, _
            X2:=680.26965 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-174.3283 * dblSizeFactor / 100 + dblPositionY, _
            X3:=680.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-467.8428 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=680.35565 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-469.3393 * dblSizeFactor / 100 + dblPositionY, _
            X2:=680.35565 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-470.8404 * dblSizeFactor / 100 + dblPositionY, _
            X3:=680.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-472.3369 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=680.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-473.0049 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=680.14045 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-473.5791 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=679.83675 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-476.4431 * dblSizeFactor / 100 + dblPositionY, _
            X2:=679.18075 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-479.2586 * dblSizeFactor / 100 + dblPositionY, _
            X3:=678.18735 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-481.9619 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=667.84705 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-520.876 * dblSizeFactor / 100 + dblPositionY, _
            X2:=635.38245 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-547.8815 * dblSizeFactor / 100 + dblPositionY, _
            X3:=595.57215 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-568.8604 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=553.91515 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-590.8123 * dblSizeFactor / 100 + dblPositionY, _
            X2:=500.77965 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-607.5375 * dblSizeFactor / 100 + dblPositionY, _
            X3:=439.08185 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-620.9932 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=315.68635 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-647.9045 * dblSizeFactor / 100 + dblPositionY, _
            X2:=158.13285 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-660.915 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-660.915 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-590.915 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=154.22855 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-590.915 * dblSizeFactor / 100 + dblPositionY, _
            X2:=308.39615 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-577.8488 * dblSizeFactor / 100 + dblPositionY, _
            X3:=424.16585 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-552.6006 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=482.05075 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-539.9764 * dblSizeFactor / 100 + dblPositionY, _
            X2:=530.35905 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-524.1004 * dblSizeFactor / 100 + dblPositionY, _
            X3:=562.93735 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-506.9326 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=588.93745 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-493.2314 * dblSizeFactor / 100 + dblPositionY, _
            X2:=602.54785 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-479.6193 * dblSizeFactor / 100 + dblPositionY, _
            X3:=608.11705 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-469.4365 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=602.54755 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-459.2539 * dblSizeFactor / 100 + dblPositionY, _
            X2:=588.93665 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-445.6412 * dblSizeFactor / 100 + dblPositionY, _
            X3:=562.93735 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-431.9404 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=530.35905 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-414.7727 * dblSizeFactor / 100 + dblPositionY, _
            X2:=482.05075 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-398.8985 * dblSizeFactor / 100 + dblPositionY, _
            X3:=424.16585 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-386.2744 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=308.39615 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-361.0261 * dblSizeFactor / 100 + dblPositionY, _
            X2:=154.22855 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-347.958 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-347.958 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-153.58815 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-347.958 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-307.75585 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-361.0261 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-423.52555 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-386.2744 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-481.41035 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-398.8985 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-529.71875 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-414.7727 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-562.29705 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-431.9404 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-588.29625 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-445.6412 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-601.90715 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-459.2539 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-607.47675 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-469.4365 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-601.90755 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-479.6193 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-588.29705 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-493.2314 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-562.29705 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-506.9326 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-529.71875 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-524.1004 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-481.41035 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-539.9764 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-423.52555 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-552.6006 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-307.75585 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-577.8488 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-153.58815 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-590.915 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-590.915 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-609.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-378.2725 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-604.85435 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-375.42 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-599.95915 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-372.662 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-594.93175 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-370.0127 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-553.27475 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-348.0607 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-500.13935 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-331.3375 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-438.44155 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-317.8818 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-315.04595 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-290.9705 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-157.49255 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-277.958 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-277.958 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=158.13285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-277.958 * dblSizeFactor / 100 + dblPositionY, _
            X2:=315.68635 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-290.9705 * dblSizeFactor / 100 + dblPositionY, _
            X3:=439.08185 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-317.8818 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=500.77965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-331.3375 * dblSizeFactor / 100 + dblPositionY, _
            X2:=553.91515 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-348.0607 * dblSizeFactor / 100 + dblPositionY, _
            X3:=595.57215 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-370.0127 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=600.59945 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-372.662 * dblSizeFactor / 100 + dblPositionY, _
            X2:=605.49465 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-375.42 * dblSizeFactor / 100 + dblPositionY, _
            X3:=610.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-378.2725 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=610.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-296.21 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=606.83915 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-285.801 * dblSizeFactor / 100 + dblPositionY, _
            X2:=593.31795 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-269.9032 * dblSizeFactor / 100 + dblPositionY, _
            X3:=562.93735 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-253.8936 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=530.35905 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-236.7258 * dblSizeFactor / 100 + dblPositionY, _
            X2:=482.05075 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-220.8517 * dblSizeFactor / 100 + dblPositionY, _
            X3:=424.16585 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-208.2275 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=308.39615 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-182.9793 * dblSizeFactor / 100 + dblPositionY, _
            X2:=154.22855 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-169.9111 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-169.9111 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-153.58815 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-169.9111 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-307.75585 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-182.9793 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-423.52555 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-208.2275 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-481.41035 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-220.8517 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-529.71875 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-236.7258 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-562.29705 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-253.8936 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-592.67755 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-269.9032 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-606.19885 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-285.801 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-609.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-296.21 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-609.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-200.2256 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-604.85435 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-197.3732 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-599.95915 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-194.6151 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-594.93175 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-191.9658 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-553.27475 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-170.0138 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-500.13935 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-153.2907 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-438.44155 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-139.835 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-315.04595 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-112.9236 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-157.49255 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-99.9111 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-99.9111 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=158.13285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-99.9111 * dblSizeFactor / 100 + dblPositionY, _
            X2:=315.68635 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-112.9236 * dblSizeFactor / 100 + dblPositionY, _
            X3:=439.08185 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-139.835 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=500.77965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-153.2907 * dblSizeFactor / 100 + dblPositionY, _
            X2:=553.91515 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-170.0138 * dblSizeFactor / 100 + dblPositionY, _
            X3:=595.57215 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-191.9658 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=600.59945 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-194.6151 * dblSizeFactor / 100 + dblPositionY, _
            X2:=605.49465 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-197.3732 * dblSizeFactor / 100 + dblPositionY, _
            X3:=610.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-200.2256 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=610.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-118.1631 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=606.83915 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-107.7541 * dblSizeFactor / 100 + dblPositionY, _
            X2:=593.31795 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-91.8563 * dblSizeFactor / 100 + dblPositionY, _
            X3:=562.93735 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-75.8467 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=530.35905 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-58.6789 * dblSizeFactor / 100 + dblPositionY, _
            X2:=482.05075 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-42.8048 * dblSizeFactor / 100 + dblPositionY, _
            X3:=424.16585 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-30.1807 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=308.39615 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-4.9324 * dblSizeFactor / 100 + dblPositionY, _
            X2:=154.22855 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=8.13569999999999 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=8.13569999999999 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-153.58815 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=8.13569999999999 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-307.75585 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-4.9324 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-423.52555 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-30.1807 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-481.41035 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-42.8048 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-529.71875 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-58.6789 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-562.29705 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-75.8467 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-592.67755 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-91.8563 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-606.19885 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-107.7541 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-609.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-118.1631 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-609.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-22.1768 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-604.85435 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-19.3243 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-599.95915 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-16.5682 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-594.93175 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-13.9189 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-553.27475 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=8.03299999999999 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-500.13935 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=24.7562 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-438.44155 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=38.2119 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-315.04595 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=65.1233 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-157.49255 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=78.1357 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=78.1357 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=158.13285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=78.1357 * dblSizeFactor / 100 + dblPositionY, _
            X2:=315.68635 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=65.1233 * dblSizeFactor / 100 + dblPositionY, _
            X3:=439.08185 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=38.2119 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=500.77965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=24.7562 * dblSizeFactor / 100 + dblPositionY, _
            X2:=553.91515 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=8.03299999999999 * dblSizeFactor / 100 + dblPositionY, _
            X3:=595.57215 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-13.9189 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=600.59945 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-16.5682 * dblSizeFactor / 100 + dblPositionY, _
            X2:=605.49465 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-19.3243 * dblSizeFactor / 100 + dblPositionY, _
            X3:=610.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-22.1768 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=610.25965 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=464.6162 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=606.83915 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=475.0252 * dblSizeFactor / 100 + dblPositionY, _
            X2:=593.31795 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=490.923 * dblSizeFactor / 100 + dblPositionY, _
            X3:=562.93735 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=506.9326 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=530.35905 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=524.1004 * dblSizeFactor / 100 + dblPositionY, _
            X2:=482.05075 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=539.9765 * dblSizeFactor / 100 + dblPositionY, _
            X3:=424.16585 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=552.6006 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=308.39615 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=577.8488 * dblSizeFactor / 100 + dblPositionY, _
            X2:=154.22855 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=590.915 * dblSizeFactor / 100 + dblPositionY, _
            X3:=0.320149999999899 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=590.915 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-153.58815 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=590.915 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-307.75585 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=577.8488 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-423.52555 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=552.6006 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-481.41035 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=539.9765 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-529.71875 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=524.1004 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-562.29705 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=506.9326 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-592.67755 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=490.923 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-606.19885 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=475.0252 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-609.61925 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=464.6162 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    Set shp4 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 3)
    Set shp5 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 4)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name, shp4.Name, shp5.Name)). _
        MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With

End Sub
'end data source


'To draw business rule : activity
Sub DrawBusinessRule(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-821.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-617 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-821.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=617 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=821.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=617 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=821.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-617 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-199 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-435 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-199 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-435 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=145.4355 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=145.4355 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-365 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-199 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-199 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=145.4355 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-365 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=145.4355 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=215.4355 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-435 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=215.4355 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-435 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=547 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=547 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-365 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=215.4355 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=215.4355 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=751.8945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=547 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-365 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=547 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    Set shp4 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 3)
    Set shp5 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 4)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name, shp4.Name, shp5.Name)). _
        MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end business rule


'To draw script : activity
Sub DrawScript(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-254.20895 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-622 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-262.47075 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-617.0957 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-265.16995 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-615.4922 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-381.96135 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-546.2796 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-462.33475 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-482.8145 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-515.39065 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-420.8398 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-568.51995 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-358.7794 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-594.20715 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-296.2919 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-594.93755 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-236.0527 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-596.39095 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-116.1723 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-507.96525 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-26.3341 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-429.51175 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=54.0137 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-351.11795 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=134.3003 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-281.18915 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=210.0155 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-274.60745 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=272.6777 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-271.25275 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=304.6164 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-278.80225 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=336.6966 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-313.62695 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=380.6211 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-348.30515 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=424.3609 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-411.24365 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=476.7493 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-511.86715 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=535.4062 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-660.42575 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=622 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=256.42775 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=622 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=267.27735 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=615.6875 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=267.28905 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=615.6777 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=373.92445 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=553.5161 * dblSizeFactor / 300 + dblPositionY, _
            X2:=445.76575 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=496.0867 * dblSizeFactor / 300 + dblPositionY, _
            X3:=491.54295 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=438.3477 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=537.45665 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=380.4366 * dblSizeFactor / 300 + dblPositionY, _
            X2:=556.13655 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=319.8482 * dblSizeFactor / 300 + dblPositionY, _
            X3:=550.16205 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=262.9688 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=538.32445 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=150.2668 * dblSizeFactor / 300 + dblPositionY, _
            X2:=446.51685 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=68.1422 * dblSizeFactor / 300 + dblPositionY, _
            X3:=369.33395 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-10.9043 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=292.09115 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-90.0122 * dblSizeFactor / 300 + dblPositionY, _
            X2:=229.48155 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-164.6251 * dblSizeFactor / 300 + dblPositionY, _
            X3:=230.33395 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-234.9277 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=230.76435 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-270.4188 * dblSizeFactor / 300 + dblPositionY, _
            X2:=244.40575 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-310.0121 * dblSizeFactor / 300 + dblPositionY, _
            X3:=287.55275 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-360.4121 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=330.62585 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-410.726 * dblSizeFactor / 300 + dblPositionY, _
            X2:=403.24835 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-469.5933 * dblSizeFactor / 300 + dblPositionY, _
            X3:=514.58395 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-535.5703 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=514.58395 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-535.5723 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=660.42575 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-622 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-228.36725 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-529.082 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=336.27535 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-529.082 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=286.54615 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-492.0037 * dblSizeFactor / 300 + dblPositionY, _
            X2:=247.22695 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-456.1845 * dblSizeFactor / 300 + dblPositionY, _
            X3:=216.96875 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-420.8398 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=163.83945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-358.7794 * dblSizeFactor / 300 + dblPositionY, _
            X2:=138.15225 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-296.2919 * dblSizeFactor / 300 + dblPositionY, _
            X3:=137.42185 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-236.0527 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=135.96845 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-116.1712 * dblSizeFactor / 300 + dblPositionY, _
            X2:=224.39845 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-26.3337 * dblSizeFactor / 300 + dblPositionY, _
            X3:=302.85155 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=54.0137 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=381.24495 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=134.2998 * dblSizeFactor / 300 + dblPositionY, _
            X2:=451.16815 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=210.0143 * dblSizeFactor / 300 + dblPositionY, _
            X3:=457.74995 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=272.6777 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=461.10465 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=304.6165 * dblSizeFactor / 300 + dblPositionY, _
            X2:=453.55705 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=336.6966 * dblSizeFactor / 300 + dblPositionY, _
            X3:=418.73245 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=380.6211 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=385.33615 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=422.7439 * dblSizeFactor / 300 + dblPositionY, _
            X2:=325.27695 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=473.0068 * dblSizeFactor / 300 + dblPositionY, _
            X3:=230.94335 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=529.082 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-335.14645 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=529.082 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-295.87635 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=498.4553 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-264.67825 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=468.4466 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-240.81645 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=438.3496 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-194.90265 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=380.4385 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-176.22295 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=319.8501 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-182.19725 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=262.9707 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-194.03495 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=150.2687 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-285.84255 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=68.1442 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-363.02535 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-10.9023 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-440.26815 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-90.0102000000001 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-502.87585 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-164.6231 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-502.02345 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-234.9258 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-501.59315 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-270.4167 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-487.95355 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-310.0102 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-444.80665 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-360.4102 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-403.17995 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-409.0343 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-333.53245 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-465.7565 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-228.36725 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-529.082 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-360.48825 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-370.6074 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-360.48825 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-327.6875 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=80.9101499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-327.6875 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=80.9101499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-370.6074 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-376.89065 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-139.4316 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-376.89065 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-96.5137 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=79.49415 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-96.5137 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=79.49415 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-139.4316 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-170.09575 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=91.7383 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-170.09575 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=134.6562 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=271.89255 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=134.6562 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=271.89255 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=91.7383 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-106.89845 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=322.9121 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-106.89845 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=365.8301 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=351.08785 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=365.8301 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=351.08785 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=322.9121 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    Set shp4 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 3)
    Set shp5 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 4)
    Set shp6 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 5)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name, shp4.Name, shp5.Name, shp6.Name)). _
        MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw script

'To draw adhoc : activity
Sub DrawAdhoc(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-700 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=45.3135 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-641.6676 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-92.915 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-565.1057 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-236.9172 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-433.454 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-314.7723 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-335.6702 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-373.1636 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-214.7527 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-337.1912 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-125.0255 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-279.9533 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=13.1616 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-194.7462 * dblSizeFactor / 300 + dblPositionY, _
            X2:=121.2656 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-68.1105 * dblSizeFactor / 300 + dblPositionY, _
            X3:=257.5801 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=19.5545 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=339.9153 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=67.8188 * dblSizeFactor / 300 + dblPositionY, _
            X2:=442.3132 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=28.2719 * dblSizeFactor / 300 + dblPositionY, _
            X3:=502.3275 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-38.502 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=574.729 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-122.9474 * dblSizeFactor / 300 + dblPositionY, _
            X2:=657.5431 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-202.5245 * dblSizeFactor / 300 + dblPositionY, _
            X3:=700 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-308.483 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=700 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=21.5547 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=638.6687 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=143.2249 * dblSizeFactor / 300 + dblPositionY, _
            X2:=559.7687 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=269.7074 * dblSizeFactor / 300 + dblPositionY, _
            X3:=433.608 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=328.724 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=330.3799 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=373.1636 * dblSizeFactor / 300 + dblPositionY, _
            X2:=210.4599 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=346.5129 * dblSizeFactor / 300 + dblPositionY, _
            X3:=121.0842 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=282.1383 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-9.93520000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=194.1594 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-106.402 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=58.4485 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-248.7703 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-12.6413 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-317.9421 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-48.6459 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-406.147 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-40.1869 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-464.101 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=13.9813 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-568.588 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=106.4851 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-628.4657 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=237.3012 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-700 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=355.4432 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-700 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=252.0666 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-700 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=148.69 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-700 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=45.3135 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw adhoc


'To draw service : activity
Sub DrawService(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-206.2539 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-591.44095 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-206.3002 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-554.44245 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-206.2437 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-517.44315 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-206.1484 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-480.44485 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-237.7017 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-471.51815 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-266.5348 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-459.06065 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-293.7637 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-443.80025 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-373.3457 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-522.42525 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-522.1523 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-372.75735 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-442.5762 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-294.14595 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-457.9478 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-266.59445 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-469.9513 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-237.31535 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-478.3281 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-206.90385 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-591.0527 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-206.69875 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-591.0527 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=4.09034999999994 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-477.1504 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=3.68015000000014 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-466.7345 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=44.97855 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-444.0658 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=82.10535 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-421.4883 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=115.30325 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-421.4883 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-66.5210499999999 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-521.0527 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-66.16355 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-521.0527 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-136.82565 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-421.9844 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-137.00535 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-416.4238 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-165.17725 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-408.3143 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-206.26185 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-392.254 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-245.42675 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-369.1797 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-280.47025 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-353.3945 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-304.44285 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-422.998 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-373.20455 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-372.9043 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-423.58935 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-303.1738 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-354.69675 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-279.4434 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-370.53275 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-243.1795 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-394.45825 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-203.5155 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-410.49625 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-163.9629 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-418.84915 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-136.0039 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-424.65385 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-136.2539 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-521.43895 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-64.2871 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-521.43895 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-64.8437 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-425.15775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-36.4160000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-425.15775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=123.9902 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-425.15775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=106.2768 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-446.30135 * dblSizeFactor / 300 + dblPositionY, _
            X2:=32.8697999999999 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-473.81535 * dblSizeFactor / 300 + dblPositionY, _
            X3:=5.4824000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-481.48195 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=6.11719999999991 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-591.43895 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-72.2828 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-591.44095 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-133.7483 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-591.43455 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-206.2539 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-591.44095 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-7.09570000000008 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-389.78275 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-6.80860000000007 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-278.79055 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-38.3630000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-269.86385 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-67.1962 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-257.40805 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-94.4258 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-242.14595 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-174.0098 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-320.77095 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-322.8164 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-171.10305 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-243.2383 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-92.4917499999999 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-258.6106 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-64.9386499999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-270.6154 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-35.66045 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-278.9922 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-5.24755000000005 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-391.7148 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-5.04444999999987 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-391.7148 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=205.74465 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-277.8145 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=205.33645 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-268.849 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=236.59345 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-256.4203 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=265.13625 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-241.1895 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=292.08255 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-322.6836 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=373.25825 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-172.2188 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=521.13725 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-91.0176 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=440.31295 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-63.2033 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=455.75765 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-33.6371 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=467.74895 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-2.98440000000005 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=476.09425 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-2.93159999999989 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=590.80715 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=69.4601 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=591.44095 * dblSizeFactor / 300 + dblPositionY, _
            X2:=146.5354 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=591.11445 * dblSizeFactor / 300 + dblPositionY, _
            X3:=208.4043 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=591.10795 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=208.4043 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=475.03755 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=239.9821 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=466.13375 * dblSizeFactor / 300 + dblPositionY, _
            X2:=268.9053 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=453.58665 * dblSizeFactor / 300 + dblPositionY, _
            X3:=296.1406 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=438.33255 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=377.2754 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=518.32865 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=526.1797 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=368.84815 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=444.875 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=288.72505 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=460.282 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=261.09945 * dblSizeFactor / 300 + dblPositionY, _
            X2:=472.2874 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=231.77135 * dblSizeFactor / 300 + dblPositionY, _
            X3:=480.6738 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=201.32665 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=591.0527 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=200.64505 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=591.0527 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-9.94094999999993 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=479.4355 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-9.2612499999999 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=470.4631 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-40.5482499999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=458.2142 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-68.89805 * dblSizeFactor / 300 + dblPositionY, _
            X3:=442.7812 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-96.0229499999999 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=519.9023 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-173.34135 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=369.5215 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-321.47215 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=292.5 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-244.18705 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=264.7997 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-259.50595 * dblSizeFactor / 300 + dblPositionY, _
            X2:=235.3537 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-271.46935 * dblSizeFactor / 300 + dblPositionY, _
            X3:=204.8203 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-279.82565 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=205.4531 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-389.78275 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=63.0840000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-319.78275 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=135.0508 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-319.78275 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=134.4961 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-223.49955 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=162.9219 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-217.92135 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=204.1943 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-209.82235 * dblSizeFactor / 300 + dblPositionY, _
            X2:=243.7525 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-193.78195 * dblSizeFactor / 300 + dblPositionY, _
            X3:=278.9375 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-170.84915 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=302.7617 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-155.31985 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=369.9785 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-222.76515 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=420.5957 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-172.90385 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=353.2852 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-105.42135 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=369.4805 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-81.5151499999999 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=393.2452 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-45.64225 * dblSizeFactor / 300 + dblPositionY, _
            X2:=409.0088 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-6.86524999999995 * dblSizeFactor / 300 + dblPositionY, _
            X3:=417.9082 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=33.06495 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=423.7637 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=61.07665 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=521.0449 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=60.4848500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=521.0449 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=131.07475 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=424.3066 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=131.67045 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=418.75 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=159.68215 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=410.5902 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=200.81285 * dblSizeFactor / 300 + dblPositionY, _
            X2:=394.5649 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=239.97625 * dblSizeFactor / 300 + dblPositionY, _
            X3:=371.4512 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=275.08055 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=355.6484 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=299.08645 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=426.9062 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=369.30905 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=376.8203 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=419.58835 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=305.5547 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=349.32275 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=281.8652 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=365.06105 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=245.59 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=388.76755 * dblSizeFactor / 300 + dblPositionY, _
            X2:=206.2406 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=405.18375 * dblSizeFactor / 300 + dblPositionY, _
            X3:=166.541 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=413.27195 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=138.3965 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=418.89305 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=138.3965 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=521.10795 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=119.1555 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=521.15465 * dblSizeFactor / 300 + dblPositionY, _
            X2:=98.1656 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=521.12455 * dblSizeFactor / 300 + dblPositionY, _
            X3:=67.0273 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=521.08055 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=66.9824000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=419.63135 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=38.7109 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=414.10985 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-2.61920000000009 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=406.03905 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-42.1123 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=390.07235 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-77.2363 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=367.04345 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-100.9902 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=351.46925 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-172.5313 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=422.68015 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-223.1563 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=372.92235 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-151.5156 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=301.56105 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-167.8633 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=277.55715 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-191.6501 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=241.76295 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-207.3862 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=202.94485 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-216.3008 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=163.06105 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-222.1504 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=135.13525 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-321.7148 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=135.49075 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-321.7148 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=64.8305500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-222.6465 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=64.64895 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-217.0859 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=36.4809500000001 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-208.9764 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-4.60394999999994 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-192.9158 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-43.77095 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-169.8418 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-78.81395 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-154.0566 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-102.78855 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-223.6621 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-171.55025 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-173.5684 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-221.93505 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-103.8359 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-153.04245 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-80.1055 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-168.87645 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-43.7956 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-192.64615 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-5.26459999999997 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-208.41785 * dblSizeFactor / 300 + dblPositionY, _
            X3:=35.375 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-217.19095 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=63.3359 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-222.99565 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=100.25 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-62.5795499999999 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=13.6031 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-62.5795499999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-57.3887 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=8.41405000000009 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-57.3887 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=95.06105 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=13.6031 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=252.69775 * dblSizeFactor / 300 + dblPositionY, _
            X2:=100.25 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=252.69775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=257.8887 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=181.70805 * dblSizeFactor / 300 + dblPositionY, _
            X2:=257.8887 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=95.06105 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=186.8969 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-62.5795499999999 * dblSizeFactor / 300 + dblPositionY, _
            X2:=100.25 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-62.5795499999999 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=100.25 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=7.42045000000007 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=149.0662 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=7.42045000000007 * dblSizeFactor / 300 + dblPositionY, _
            X2:=187.8887 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=46.24495 * dblSizeFactor / 300 + dblPositionY, _
            X3:=187.8887 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=95.06105 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=149.0662 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=182.69775 * dblSizeFactor / 300 + dblPositionY, _
            X2:=100.25 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=182.69775 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=12.6133 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=143.87715 * dblSizeFactor / 300 + dblPositionY, _
            X2:=12.6133 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=95.06105 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=51.4338 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=7.42045000000007 * dblSizeFactor / 300 + dblPositionY, _
            X2:=100.25 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=7.42045000000007 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    Set shp4 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 3)
    Set shp5 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 4)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name, shp4.Name, shp5.Name)). _
        MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw service


'To draw loop marker : activity
Sub DrawLoop(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=74.2094500000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-593.66805 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-177.75125 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-597.20345 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-419.41355 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-421.43935 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-491.72665 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-179.52045 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-547.31085 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-5.54395 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-514.83215 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=193.54055 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-406.20455 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=340.50585 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-495.89995 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=323.14255 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-585.59525 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=305.77925 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-675.29055 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=288.41595 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-682.89205 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=327.68745 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-690.49365 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=366.95895 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-698.09525 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=406.23045 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-536.80685 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=437.45175 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-375.51845 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=468.67315 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-214.22995 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=499.89445 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-183.00085 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=339.89905 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-151.77165 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=179.90355 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-120.54245 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=19.90815 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-159.80225 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=12.24535 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-199.06205 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=4.58264999999999 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-238.32175 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-3.08015 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-257.27555 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=94.02535 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-276.22935 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=191.13075 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-295.18315 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=288.23625 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-433.91605 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=122.63515 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-431.91265 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-139.53675 * dblSizeFactor / 300 + dblPositionY, _
            X3:=-290.81635 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-303.14275 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-153.39875 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-474.85875 * dblSizeFactor / 300 + dblPositionY, _
            X2:=108.38705 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-524.15035 * dblSizeFactor / 300 + dblPositionY, _
            X3:=299.91735 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-417.14275 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=483.14905 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-322.57485 * dblSizeFactor / 300 + dblPositionY, _
            X2:=584.80485 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-98.54645 * dblSizeFactor / 300 + dblPositionY, _
            X3:=534.81345 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=101.60305 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=490.04365 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=310.51385 * dblSizeFactor / 300 + dblPositionY, _
            X2:=287.40965 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=468.94285 * dblSizeFactor / 300 + dblPositionY, _
            X3:=74.2094500000001 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=464.65625 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=18.7244499999999 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=460.72095 * dblSizeFactor / 300 + dblPositionY, _
            X2:=-9.16525000000001 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=540.85195 * dblSizeFactor / 300 + dblPositionY, _
            X3:=36.77275 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=572.21705 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=76.8770500000001 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=597.20345 * dblSizeFactor / 300 + dblPositionY, _
            X2:=127.61925 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=579.58065 * dblSizeFactor / 300 + dblPositionY, _
            X3:=171.58125 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=576.69175 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=419.76215 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=538.94395 * dblSizeFactor / 300 + dblPositionY, _
            X2:=629.10065 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=327.23995 * dblSizeFactor / 300 + dblPositionY, _
            X3:=661.10115 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=77.78225 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=698.09525 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-160.24305 * dblSizeFactor / 300 + dblPositionY, _
            X2:=569.71735 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-411.15345 * dblSizeFactor / 300 + dblPositionY, _
            X3:=356.29785 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-522.45935 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=270.12975 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-569.22785 * dblSizeFactor / 300 + dblPositionY, _
            X2:=172.22975 * dblSizeFactor / 300 + dblPositionX, _
            Y2:=-593.82345 * dblSizeFactor / 300 + dblPositionY, _
            X3:=74.2094500000001 * dblSizeFactor / 300 + dblPositionX, _
            Y3:=-593.66805 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw loop


'To draw compensation marker
Sub DrawCompensation(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
        
    Select Case intOption
        Case 1 'outlined compensation
            intSizeAdjustment = 2
            intCombineStyle = 0
        
        Case 2 'solid compensation
            intSizeAdjustment = 2
            intCombineStyle = 1
            
        Case 3 'outlined compensation for activity bottom marker
            intSizeAdjustment = 1.1
            intCombineStyle = 0
        
        Case Else 'outlined compensation
            intSizeAdjustment = 2
            intCombineStyle = 0
    End Select
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-13.8874499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-480.20135 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-53.49545 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-474.40735 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-82.01615 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-441.64025 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-116.03215 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-422.68535 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-302.40895 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-297.18615 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-489.24935 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-172.34065 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-675.33635 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-46.43265 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-711.26915 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-18.83685 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-699.18495 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=39.99795 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-659.26075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=57.14915 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-452.93565 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=195.57275 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-247.04115 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=334.67315 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-40.44685 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=472.67375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=1.38454999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=497.48695 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=56.21045 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=456.49295 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=47.8820499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=409.59785 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=47.8820499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=303.78325 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=47.8820499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=197.96875 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=47.8820499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=92.15415 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=236.84585 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=218.90855 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=425.37855 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=346.34055 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=614.61165 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=472.67185 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=656.44315 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=497.48515 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=711.26915 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=456.49105 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=702.94065 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=409.59585 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=702.73995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=130.99265 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=703.34195 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-147.63065 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=702.63975 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-426.22125 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=699.91185 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-474.75265 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=635.26435 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-497.48695 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=601.01835 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-464.34825 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=416.63965 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-340.43335 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=232.26085 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-216.51855 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=47.8820499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-92.60375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=47.0932499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-205.76255 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=49.4664499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-319.11985 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=46.6836499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-432.15375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=41.47955 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-460.17155 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=14.5878500000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-481.50915 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-13.8874499999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-480.20135 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-72.11795 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-307.61745 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-72.11795 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=307.16775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-224.57825 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=204.70355 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-377.03855 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=102.23945 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-529.49875 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.224750000000014 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-377.03855 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-102.68895 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-224.57825 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-205.15325 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-72.11795 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-307.61735 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=582.94065 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-307.61545 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=582.94065 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=307.16585 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=430.48035 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=204.70235 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=278.02005 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=102.23875 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=125.55975 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.224750000000014 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=278.02005 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-102.68835 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=430.48035 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-205.15195 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=582.94065 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-307.61545 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    
    If intCombineStyle = 0 Then
        ActivePresentation.Slides(intSlideNumber).Shapes. _
            Range(Array(shp1.Name, shp2.Name, shp3.Name)).MergeShapes msoMergeCombine
    Else
        ActivePresentation.Slides(intSlideNumber).Shapes. _
            Range(Array(shp1.Name, shp2.Name, shp3.Name)).MergeShapes msoMergeUnion
    End If
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw compensation


'To draw multi instance : activity
Sub DrawMultiInstance(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-500 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-600 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-500 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-300 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-300 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-600 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-100 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-600 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-100 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=100 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=100 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-600 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=300 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-600 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=300 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=500 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=600 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=500 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-600 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw multiinstance


'To draw receive : event, activity
Sub DrawReceive(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
        
    Select Case intOption
        Case 1 'send icon for activity
            intSizeAdjustment = 1
        
        Case 2 'send icon for event
            intSizeAdjustment = 1.7
            
        Case Else 'send icon for activity
            intSizeAdjustment = 1
    End Select

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-811 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-535 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-811 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=535 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=811 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=535 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=811 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-535 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-643.1191 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-465 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=643.1191 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-465 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=0 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-41.8945 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-741 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-445.6035 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=0 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=41.8945000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=741 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-445.6035 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=741 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=465 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-741 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=465 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw receive


'To draw send : event, activity
Sub DrawSend(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case 1 'send icon for activity
            intSizeAdjustment = 1
        
        Case 2 'send icon for event
            intSizeAdjustment = 1.7
            
        Case Else ''send icon for activity
            intSizeAdjustment = 1
    End Select
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-803 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-532 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-1 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-78 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=801 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-532 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-801 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-364 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-801 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=532 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=803 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=532 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=803 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-364 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-1 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=84 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw send


'To draw error : event
Sub DrawError(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
        
    Select Case intOption
        Case 1 'outlined error
            intSizeAdjustment = 350
            intCombineStyle = 0
        
        Case 2 'solid error
            intSizeAdjustment = 350
            intCombineStyle = 1
            
        Case Else 'outlined error
            intSizeAdjustment = 350
            intCombineStyle = 0
    End Select
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=3.6643 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-4.07390000000009 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=2.94 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-2.60570000000007 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=2.2157 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-1.13740000000007 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=1.4914 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=0.330799999999954 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=0.5908 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-0.840599999999995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-0.309800000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-2.01200000000017 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-1.2104 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-3.18340000000012 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-2.0284 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-0.764300000000048 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-2.8463 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=1.65480000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-3.6643 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=4.07389999999987 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-2.7472 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=2.89519999999993 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-1.8302 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=1.7165 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-0.913200000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=0.537799999999834 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=3.26999999999984E-02 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=1.62009999999987 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=0.978599999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=2.70249999999987 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=1.9246 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=3.78489999999988 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=2.5045 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=1.16529999999989 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=3.0844 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-1.4543000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=3.6643 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-4.07390000000009 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-0.9702 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-1.74990000000003 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-0.113800000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-0.664300000000139 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=0.7425 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=0.421399999999949 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=1.5988 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=1.50699999999983 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=1.8127 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=1.03089999999997 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=2.0266 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=0.554899999999861 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=2.2405 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=7.88000000000011E-02 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=2.0357 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=0.902599999999893 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=1.831 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=1.72640000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=1.6262 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=2.5501999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=0.752199999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=1.52199999999993 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-0.1219 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=0.49369999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-0.995900000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.53449999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-1.3505 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=1.36999999999716E-02 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-1.7051 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=0.561899999999923 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-2.0597 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=1.11019999999985 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-1.6965 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=0.156799999999976 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-1.3333 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-0.796600000000126 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-0.9702 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-1.74990000000003 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    
    If intCombineStyle = 0 Then
        ActivePresentation.Slides(intSlideNumber).Shapes. _
            Range(Array(shp1.Name, shp2.Name)).MergeShapes msoMergeCombine
    Else
        ActivePresentation.Slides(intSlideNumber).Shapes. _
            Range(Array(shp1.Name, shp2.Name)).MergeShapes msoMergeUnion
    End If
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw error


'To draw clock : event
Sub DrawClock(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    Select Case intOption
        Case 1 'clock for event
            intSizeAdjustment = 7
        
        Case Else 'clock for event
            intSizeAdjustment = 7
    End Select
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-229.391 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-0.912500000000023 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-229.25 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-46.6305 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-215.144 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-91.1535 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-188.837 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-128.5525 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-163.197 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-165.0005 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-126.177 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-193.0055 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-84.397 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-208.3405 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-40.928 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-224.2965 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=7.267 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-225.7555 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=51.783 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-213.2815 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=94.686 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-201.2605 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=133.074 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-175.7115 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=161.543 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-141.5595 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=221.06 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-70.1735 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=227.955 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=34.2505 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=180.155 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=113.5505 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=157.263 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=151.5295 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=122.72 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=181.7855 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=82.475 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=200.2295 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=40.29 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=219.5625 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-7.96899999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=224.5055 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-53.345 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=215.3045 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-142.766 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=197.1745 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-212.585 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=122.8145 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-226.815 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=33.0345 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-228.591 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=21.8145 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-229.35 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=10.4345 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-229.385 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.915500000000009 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-229.395 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-4.13950000000003 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-234.395 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-4.14050000000003 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-234.385 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.915500000000009 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-234.247 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=43.8415 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-221.025 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=87.7095 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-196.13 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=124.9245 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-172 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=160.9985 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-137.301 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=189.5635 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-97.463 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=206.6805 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-55.99 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=224.4995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-9.58499999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=229.0685 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=34.687 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=220.6535 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=77.742 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=212.4705 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=117.756 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=191.0465 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=149.127 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=160.5725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=214.094 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=97.4715 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=234.395 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-1.57750000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=200.717 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-85.4475 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=167.457 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-168.2985 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=85.977 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-224.1075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-3.023 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-226.5775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-92.643 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-229.0685 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-176.853 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-175.8965 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-214.113 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-94.5775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-227.558 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-65.2475 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-234.283 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-33.1475 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-234.383 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.917500000000018 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-234.393 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=2.30249999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-229.393 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=2.30249999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-229.383 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.917500000000018 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=183.1575 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=211.6875 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=218.1355 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=218.1355 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=211.6875 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=183.1575 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=176.7075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=176.7075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=183.1575 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-104.171 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=156.2475 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-109.207 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=164.4035 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-114.244 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=172.5605 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-119.28 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=180.7175 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-122.678 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=186.2205 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-114.026 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=191.2385 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-110.645 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=185.7645 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-105.609 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=177.6085 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-100.572 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=169.4515 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-95.536 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=161.2945 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-92.136 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=155.7845 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-100.786 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=150.7745 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-104.166 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=156.2445 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-170.592 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=86.3175 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-195.342 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=101.0655 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-200.875 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=104.3625 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-195.848 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=113.0095 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-190.295 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=109.7005 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-165.545 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=94.9525 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-160.012 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=91.6525 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-165.038 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=83.0025 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-170.592 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=86.3125 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-193.166 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-221.261 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-227.71 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-227.71 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-221.261 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-193.166 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-186.718 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-186.718 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-193.166 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-165.184 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-97.5725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-173.554 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-102.2235 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-181.924 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-106.8735 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-190.294 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-111.5245 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-195.936 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-114.6595 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-200.978 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-106.0215 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-195.341 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-102.8895 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-186.971 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-98.2385 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-178.601 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-93.5885 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-170.231 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-88.9375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-164.59 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-85.8075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-159.548 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-94.4375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-165.184 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-97.5675 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-96.971 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-163.3805 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-101.528 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-171.4515 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-106.085 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-179.5215 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-110.642 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-187.5915 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-113.812 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-193.2055 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-122.452 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-188.1675 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-119.277 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-182.5445 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-114.72 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-174.4735 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-110.163 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-166.4035 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-105.606 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-158.3335 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-102.446 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-152.7195 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-93.806 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-157.7585 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-96.976 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-163.3805 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-185.3485 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-213.5115 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-219.9605 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-219.9605 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-213.5115 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-185.3485 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-13.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-178.9005 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-178.9005 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-3.661 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-185.3485 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=88.459 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-157.5375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=92.957 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-165.8735 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=97.455 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-174.2095 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=101.953 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-182.5455 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=105.015 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-188.2195 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=96.381 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-193.2685 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=93.318 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-187.5925 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=88.82 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-179.2565 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=84.322 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-170.9205 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=79.824 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-162.5845 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=76.774 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-156.9095 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=85.404 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-151.8615 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=88.464 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-157.5375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=153.269 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-88.1425 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=161.519 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-93.0595 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=169.769 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-97.9755 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=178.019 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-102.8915 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=183.551 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-106.1885 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=178.526 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-114.8355 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=172.972 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-111.5265 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=164.722 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-106.6095 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=156.472 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-101.6935 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=148.222 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-96.7775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=142.692 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-93.4775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=147.712 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-84.8275 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=153.272 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-88.1375 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=175.339 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=203.937 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=210.385 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=210.385 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=203.937 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=175.339 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=168.889 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-5.91250000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=168.889 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=175.339 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=4.08749999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=148.579 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=95.7775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=156.709 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=100.4215 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=164.839 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=105.0635 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=172.969 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=109.7075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=178.577 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=112.9105 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=183.614 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=104.2695 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=178.016 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=101.0725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=169.886 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=96.4285 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=161.756 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=91.7865 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=153.626 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=87.1425 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=148.016 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=83.9425 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=142.976 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=92.5825 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=148.576 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=95.7825 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=79.109 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=160.4775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=83.846 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=168.9075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=88.583 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=177.3365 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=93.32 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=185.7655 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=96.479 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=191.3865 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=105.118 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=186.3465 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=101.955 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=180.7185 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=97.218 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=172.2885 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=92.481 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=163.8605 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=87.744 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=155.4305 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=84.594 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=149.8105 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=75.954 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=154.8505 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=79.114 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=160.4805 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-5.80099999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-0.422500000000014 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-5.80099999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-112.7425 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-5.80099999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-115.9665 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-10.801 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-115.9665 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-10.801 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-112.7425 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-10.801 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-0.422500000000014 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-10.801 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=2.79749999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-5.80099999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=2.79749999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-5.80099999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-0.422500000000014 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-6.941 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-2.43250000000003 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-31.628 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-45.1495 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-56.315 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-87.8665 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-81.002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-130.5825 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-82.977 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-133.9995 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-84.952 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-137.4165 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-86.927 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-140.8345 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-88.539 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-143.6245 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-92.861 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-141.1075 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-91.244 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-138.3115 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-66.557 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-95.5945 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-41.87 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-52.8775 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-17.183 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-10.1615 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-15.208 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-6.74450000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-13.233 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-3.32750000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-11.258 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=9.04999999999916E-02 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-9.62799999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=2.88049999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-5.30799999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=0.370499999999993 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-6.928 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-2.42950000000002 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-7.511 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-13.6725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-8.21199999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-13.6725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-15.026 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-13.6725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-20.712 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-7.98650000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=-20.712 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-1.17250000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=-15.026 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=11.3275 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-8.21199999999999 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=11.3275 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-7.511 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=11.3275 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-0.697000000000003 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=11.3275 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=4.989 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=5.64149999999998 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X3:=4.989 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y3:=-1.17250000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingSmooth, _
            X1:=-0.700999999999993 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-13.6725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY, _
            X2:=-7.511 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y2:=-13.6725 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    Set shp4 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 3)
    Set shp5 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 4)
    Set shp6 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 5)
    Set shp7 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 6)
    Set shp8 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 7)
    Set shp9 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 8)
    Set shp10 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 9)
    Set shp11 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 10)
    Set shp12 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 11)
    Set shp13 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 12)
    Set shp14 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 13)
    Set shp15 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 14)
    Set shp16 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 15)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name, shp4.Name, shp5.Name, shp6.Name, _
        shp7.Name, shp8.Name, shp9.Name, shp10.Name, shp11.Name, shp12.Name, _
        shp13.Name, shp14.Name, shp15.Name, shp16.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
            
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw clock


'To draw sequential marker : activity
Sub DrawSequential(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-500 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-300 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-300 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-500 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-500 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-100 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=100 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=100 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-100 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=-100 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=300 * dblSizeFactor / 300 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=500 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=500 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=300 * dblSizeFactor / 300 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-600 * dblSizeFactor / 300 + dblPositionX, _
            Y1:=300 * dblSizeFactor / 300 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
        
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Solid
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With

End Sub
'end draw sequential


'To draw conditional : event
Sub DrawConditional(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
        
    Select Case intOption
        Case 1 'conditional for event
            intSizeAdjustment = 3
            intCombineStyle = 0

        Case Else 'conditional for event
            intSizeAdjustment = 3
            intCombineStyle = 0
    End Select
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-310.0059 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-390.0352 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-310.0059 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-364.9082 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-310.0059 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=390.0352 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=390.0352 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=390.0352 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=310.0059 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=390.0352 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=310.0059 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-390.0352 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-260.1191 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-339.9648 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=259.9277 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-339.9648 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=259.9277 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=339.9551 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=339.9551 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=339.9551 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-260.1191 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=339.9551 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-263.4238 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-213.5391 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-213.5391 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-263.4238 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-112.8184 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-62.7402 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-62.7402 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=-112.8184 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=61.2207000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=111.1152 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=111.1152 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=61.2207000000001 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=218.4434 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-206.4512 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=268.3398 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=268.3398 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=206.2617 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionX, _
            Y1:=218.4434 * dblSizeFactor / 300 * intSizeAdjustment + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    Set shp4 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 3)
    Set shp5 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 4)
    Set shp6 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 5)
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name, shp4.Name, shp5.Name, shp6.Name)). _
        MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
        
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor)) & ";"
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw conditional


'To draw data object : artifact
Sub DrawDataObject(ByVal intSlideNumber As Integer, ByVal intIndex As Integer, _
    ByVal intSubIndex As Integer, ByVal intOption As Integer, _
    ByVal dblSizeFactor As Double, ByVal dblPositionX As Double, _
    ByVal dblPositionY As Double)

    Set myDocument = ActivePresentation.Slides(intSlideNumber)
    
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=251.9355 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-792.1065 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-34.2949 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-792.1058 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-320.5254 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-792.1055 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-606.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-792.1045 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=-606.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=792.1065 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=606.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=792.1065 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=606.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=385.488 * dblSizeFactor / 100 + dblPositionY, _
            X2:=606.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-21.1305 * dblSizeFactor / 100 + dblPositionY, _
            X3:=606.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-427.749 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=488.4824 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-549.2015 * dblSizeFactor / 100 + dblPositionY, _
            X2:=370.209 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-670.6539 * dblSizeFactor / 100 + dblPositionY, _
            X3:=251.9355 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-792.1064 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=147.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-702.1065 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=147.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-330.458 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=516.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-330.458 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _
            X1:=516.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=702.1065 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=172.252 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=702.1065 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-172.2519 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=702.1065 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-516.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=702.1065 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-516.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=234.0362 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-516.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-234.0342 * dblSizeFactor / 100 + dblPositionY, _
            X3:=-516.7559 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-702.1045 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=-295.2611 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-702.1051 * dblSizeFactor / 100 + dblPositionY, _
            X2:=-73.7663 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-702.1055 * dblSizeFactor / 100 + dblPositionY, _
            X3:=147.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-702.1065 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
            X1:=237.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-677.6944 * dblSizeFactor / 100 + dblPositionY)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=321.2298 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-591.9489 * dblSizeFactor / 100 + dblPositionY, _
            X2:=404.7311 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-506.2035 * dblSizeFactor / 100 + dblPositionY, _
            X3:=488.2324 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-420.458 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=404.7311 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-420.458 * dblSizeFactor / 100 + dblPositionY, _
            X2:=321.2298 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-420.458 * dblSizeFactor / 100 + dblPositionY, _
            X3:=237.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-420.458 * dblSizeFactor / 100 + dblPositionY
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _
            X1:=237.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y1:=-506.2035 * dblSizeFactor / 100 + dblPositionY, _
            X2:=237.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y2:=-591.9489 * dblSizeFactor / 100 + dblPositionY, _
            X3:=237.7285 * dblSizeFactor / 100 + dblPositionX, _
            Y3:=-677.6944 * dblSizeFactor / 100 + dblPositionY
        .ConvertToShape
    End With

    Set shp1 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 0)
    Set shp2 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 1)
    Set shp3 = ActivePresentation.Slides(intSlideNumber). _
        Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count - 2)
    
    ActivePresentation.Slides(intSlideNumber).Shapes. _
        Range(Array(shp1.Name, shp2.Name, shp3.Name)).MergeShapes msoMergeCombine
    
    With ActivePresentation.Slides(intSlideNumber). _
         Shapes(ActivePresentation.Slides(intSlideNumber).Shapes.Count)
         
        .Name = SHAPE_PREFIX_NAME & Trim(Str(intIndex)) & _
            IIf(intSubIndex = 0, "", "-" & Trim(Str(intSubIndex)))
        .Title = Trim(Str(fShapeTypes.fObject)) & ";" & Trim(Str(dblPositionX)) & ";" & _
            Trim(Str(dblPositionY)) & ";" & Trim(Str(.Width)) & ";" & _
            Trim(Str(.Height)) & ";" & "0" & ";" & Trim(Str(dblSizeFactor))
            
        With .line
            .Visible = msoFalse
        End With

        With .Fill
            .Transparency = 0
            .ForeColor.RGB = GetRGB(EVENT_OUTLINE_COLOR)
        End With
        
    End With
    
End Sub
'end draw data object
