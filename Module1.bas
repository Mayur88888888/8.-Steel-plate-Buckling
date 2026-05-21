Attribute VB_Name = "Module1"
Option Explicit

' =========================================================
' DYNAMIC PLATE BUCKLING SIMULATION
' =========================================================

Const totalSteps As Integer = 100

Const E As Double = 200000
Const nu As Double = 0.3
Const PI As Double = 3.14159265358979

Const PLATE_LEFT As Double = 120
Const PLATE_TOP As Double = 280

Sub SimulateDynamicBuckling()

    Dim ws As Worksheet

    Dim width As Double
    Dim thickness As Double
    Dim length As Double
    Dim totalLoadTons As Double
    Dim totalLoadN As Double
    Dim maxArcHeight As Double

    Dim sigma_cr As Double
    Dim Pcr As Double

    Dim stepNum As Integer

    Dim loadFraction As Double
    Dim loadN As Double

    Dim deflectionRaw() As Double

    Dim maxDeflectionRaw As Double

    Dim scaleFactor As Double
    Dim arcHeight As Double

    Dim plateWidthPx As Double
    Dim plateHeightPx As Double

    Dim k As Double
    Dim factor As Double

    Dim s As Shape

    ' =====================================================
    ' CREATE / ACCESS SHEET
    ' =====================================================

    On Error Resume Next
    Set ws = Worksheets("Simulation")
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = Worksheets.Add
        ws.Name = "Simulation"

    End If

    ws.Activate

    ' =====================================================
    ' CLEAN PREVIOUS CONTENT
    ' =====================================================

    ws.Range("D2:I200").ClearContents

    For Each s In ws.Shapes
        s.Delete
    Next s

    On Error Resume Next
    ws.ChartObjects("DeflectionChart").Delete
    On Error GoTo 0

    ' =====================================================
    ' INPUT SECTION
    ' =====================================================

    ws.Range("A1").Value = "PLATE BUCKLING SIMULATION"

    ws.Range("A2").Value = "Plate Width (mm)"
    ws.Range("A3").Value = "Applied Load (tons)"
    ws.Range("A4").Value = "Thickness (mm)"
    ws.Range("A5").Value = "Plate Length (mm)"
    ws.Range("A6").Value = "Max Arc Height (px)"

    ' DEFAULT VALUES
 ' =====================================================
' CHECK USER INPUTS
' =====================================================

    If ws.Range("B2").Value = "" _
    Or ws.Range("B3").Value = "" _
    Or ws.Range("B4").Value = "" _
    Or ws.Range("B5").Value = "" _
    Or ws.Range("B6").Value = "" Then

        MsgBox "Please enter all input values in cells B2:B6.", vbExclamation
        Exit Sub

    End If
 
 
 

    ' FORMAT INPUT CELLS
    With ws.Range("B2:B6")
        .Interior.Color = RGB(255, 255, 180)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    ' =====================================================
    ' READ USER INPUTS
    ' =====================================================

    width = CDbl(ws.Range("B2").Value)
    totalLoadTons = CDbl(ws.Range("B3").Value)
    thickness = CDbl(ws.Range("B4").Value)
    length = CDbl(ws.Range("B5").Value)
    maxArcHeight = CDbl(ws.Range("B6").Value)

    ' =====================================================
    ' VALIDATION
    ' =====================================================

    If width <= 0 _
    Or thickness <= 0 _
    Or length <= 0 _
    Or totalLoadTons <= 0 Then

        MsgBox "Please enter valid positive values.", vbCritical
        Exit Sub

    End If

    totalLoadN = totalLoadTons * 1000 * 9.81

    plateWidthPx = 450
    plateHeightPx = 20

    ' =====================================================
    ' BUCKLING CALCULATIONS
    ' =====================================================

    k = 4

    factor = (k * PI ^ 2 * E) / _
    (12 * (1 - nu ^ 2))

    sigma_cr = factor * (thickness / length) ^ 2

    Pcr = sigma_cr * thickness * width

    ' VISUAL REDUCTION FACTOR
    Pcr = Pcr * 0.02

    ReDim deflectionRaw(1 To totalSteps)

    maxDeflectionRaw = 0

    For stepNum = 1 To totalSteps

        loadFraction = stepNum / totalSteps

        loadN = loadFraction * totalLoadN

        ' =============================================
        ' ELASTIC + POST BUCKLING BEHAVIOR
        ' =============================================

        If loadN < Pcr Then

            deflectionRaw(stepNum) = _
            (loadN / Pcr) * 5

        Else

            deflectionRaw(stepNum) = _
            5 + Sqr(loadN - Pcr) * 0.08

        End If

        If deflectionRaw(stepNum) > maxDeflectionRaw Then
            maxDeflectionRaw = deflectionRaw(stepNum)
        End If

    Next stepNum

    If maxDeflectionRaw <= 0 Then
        maxDeflectionRaw = 1
    End If

    scaleFactor = maxArcHeight / maxDeflectionRaw

    ' =====================================================
    ' OUTPUT INFORMATION
    ' =====================================================

    ws.Range("D2").Value = "Critical Load (kN)"
    ws.Range("D3").Value = "Current Step"
    ws.Range("D4").Value = "Current Load (kN)"
    ws.Range("D5").Value = "Deflection (mm)"

    ws.Range("E2").Value = Round(Pcr / 1000, 2)

    ws.Range("H1").Value = "Load (tons)"
    ws.Range("I1").Value = "Deflection (mm)"

    With ws.Range("D2:E5")
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 240, 255)
    End With

    ' =====================================================
    ' DRAW PLATE
    ' =====================================================

    DrawPlate ws, plateWidthPx, plateHeightPx

    ' =====================================================
    ' START SIMULATION
    ' =====================================================

    For stepNum = 1 To totalSteps

        loadFraction = stepNum / totalSteps

        loadN = loadFraction * totalLoadN

        arcHeight = _
        deflectionRaw(stepNum) * scaleFactor

        DrawBucklingArc _
        ws, _
        arcHeight, _
        plateWidthPx, _
        stepNum

        Dim deflectionMM As Double

        deflectionMM = deflectionRaw(stepNum) * 0.05

        ' LIVE READINGS
        ws.Range("E3").Value = _
        stepNum & "/" & totalSteps

        ws.Range("E4").Value = _
        Round(loadN / 1000, 2)

        ws.Range("E5").Value = _
        Round(deflectionMM, 2)

        ' STORE TABLE DATA
        ws.Range("H" & stepNum + 1).Value = _
        Round(loadN / 9810, 2)

        ws.Range("I" & stepNum + 1).Value = _
        Round(deflectionMM, 2)

        DoEvents

        WaitMilliseconds 40

        ' DELETE PREVIOUS ARC
        If stepNum > 1 Then

            On Error Resume Next

            ws.Shapes( _
            "BucklingArc" & stepNum - 1).Delete

            On Error GoTo 0

        End If

    Next stepNum

    ' =====================================================
    ' CREATE CHART
    ' =====================================================

    CreateDeflectionChart ws

    MsgBox "Simulation Completed Successfully", vbInformation

End Sub

' =========================================================
' DRAW PLATE
' =========================================================

Sub DrawPlate( _
    ws As Worksheet, _
    plateWidthPx As Double, _
    plateHeightPx As Double)

    Dim plateShape As Shape

    Set plateShape = ws.Shapes.AddShape( _
    msoShapeRectangle, _
    PLATE_LEFT, _
    PLATE_TOP, _
    plateWidthPx, _
    plateHeightPx)

    With plateShape

        .Name = "PlateShape"

        .Fill.ForeColor.RGB = RGB(180, 180, 180)

        .Line.ForeColor.RGB = RGB(0, 0, 0)

        .Line.Weight = 1.5

    End With

End Sub

' =========================================================
' DRAW BUCKLING ARC
' =========================================================

Sub DrawBucklingArc( _
    ws As Worksheet, _
    arcHeight As Double, _
    widthPx As Double, _
    stepNum As Integer)

    Dim arcShape As Shape

    Dim arcLeft As Double
    Dim arcTop As Double

    arcLeft = PLATE_LEFT

    arcTop = PLATE_TOP - arcHeight + 10

    Set arcShape = ws.Shapes.AddShape( _
    msoShapeArc, _
    arcLeft, _
    arcTop, _
    widthPx, _
    arcHeight)

    With arcShape

        .Name = "BucklingArc" & stepNum

        .Line.ForeColor.RGB = RGB(255, 0, 0)

        .Line.Weight = 3

        .Adjustments.Item(1) = 0
        .Adjustments.Item(2) = 180

    End With

End Sub

' =========================================================
' CREATE CHART
' =========================================================

Sub CreateDeflectionChart(ws As Worksheet)

    Dim chObj As ChartObject

    Set chObj = ws.ChartObjects.Add( _
    Left:=650, _
    Top:=40, _
    width:=420, _
    Height:=320)

    chObj.Name = "DeflectionChart"

    With chObj.Chart

        .ChartType = xlXYScatterSmooth

        .SetSourceData _
        Source:=ws.Range("H1:I101")

        .HasTitle = True

        .ChartTitle.Text = _
        "Load vs Deflection"

        .Axes(xlCategory).HasTitle = True

        .Axes(xlCategory).AxisTitle.Text = _
        "Load (tons)"

        .Axes(xlValue).HasTitle = True

        .Axes(xlValue).AxisTitle.Text = _
        "Deflection (mm)"

    End With

End Sub

' =========================================================
' WAIT FUNCTION
' =========================================================

Sub WaitMilliseconds(ms As Long)

    Dim endTime As Double

    endTime = Timer + ms / 1000#

    Do While Timer < endTime
        DoEvents
    Loop

End Sub

