Attribute VB_Name = "Module1"
Option Explicit

Const totalSteps As Integer = 100
Const E As Double = 200000 ' Young's Modulus in MPa (N/mm²)
Const nu As Double = 0.3   ' Poisson's ratio for steel
Const PI As Double = 3.14159265358979

Sub SimulateDynamicBuckling()
    Dim thickness As Double, width As Double, length As Double
    Dim totalLoadTons As Double, totalLoadN As Double
    Dim maxRadiusPx As Double
    Dim plateWidthPx As Double
    Dim sigma_cr As Double, Pcr As Double
    Dim stepNum As Integer
    Dim loadFraction As Double, loadN As Double
    Dim deflectionRaw() As Double
    Dim maxDeflectionRaw As Double
    Dim arcRadiusPx As Double
    Dim scaleFactor As Double
    Dim arcShape As Shape
    Dim k As Double
    Dim factor As Double

    ' === Read Inputs ===
    width = CDbl(Range("B1").value)
    totalLoadTons = CDbl(Range("B2").value)
    maxRadiusPx = CDbl(Range("B3").value)
    thickness = CDbl(Range("B5").value)
    length = CDbl(Range("B6").value)
    plateWidthPx = 400 ' for drawing

    ' === Validate Inputs ===
    If length = 0 Or (1 - nu ^ 2) = 0 Then
        MsgBox "Invalid input values. Please check plate dimensions or material properties.", vbCritical
        Exit Sub
    End If

    totalLoadN = totalLoadTons * 1000 * 9.81 ' Convert tons to Newtons
    k = 4 ' Simply supported plate

    ' === Calculate critical buckling stress and total load ===
    factor = (k * PI ^ 2 * E) / (12 * (1 - nu ^ 2))
    sigma_cr = factor * (thickness / length) ^ 2
    Pcr = sigma_cr * thickness * width

    ReDim deflectionRaw(1 To totalSteps)
    maxDeflectionRaw = 0

    ' === Calculate raw deflection values ===
    For stepNum = 1 To totalSteps
        loadFraction = stepNum / totalSteps
        loadN = loadFraction * totalLoadN

        If loadN <= Pcr Then
            deflectionRaw(stepNum) = 0
        Else
            deflectionRaw(stepNum) = Sqr(loadN - Pcr)
        End If

        If deflectionRaw(stepNum) > maxDeflectionRaw Then
            maxDeflectionRaw = deflectionRaw(stepNum)
        End If
    Next stepNum

    If maxDeflectionRaw = 0 Then
        MsgBox "The applied load is below the critical buckling load. No deflection observed.", vbInformation
        Exit Sub
    End If

    scaleFactor = maxRadiusPx / maxDeflectionRaw

    ' === Display headers ===
    Range("A10").value = "Dynamic Plate Buckling Simulation"
    Range("A12").value = "Plate: L=" & length & " mm, W=" & width & " mm, t=" & thickness & " mm"
    Range("A13").value = "Critical Buckling Load: " & Format(Pcr / 1000, "0.0") & " kN"
    Range("A14").value = "Applied Load: " & Format(totalLoadN / 1000, "0.0") & " kN"
    Range("E1").value = "Load (tons)"
    Range("F1").value = "Deflection (mm approx)"

    ' === Cleanup old shapes and chart ===
    Dim s As Shape
    For Each s In ActiveSheet.Shapes
        If s.Name Like "BucklingArc*" Then s.Delete
    Next s
    On Error Resume Next
    ActiveSheet.ChartObjects("DeflectionChart").Delete
    On Error GoTo 0

    ' === Animate deflection ===
    For stepNum = 1 To totalSteps
        loadFraction = stepNum / totalSteps
        loadN = loadFraction * totalLoadN

        arcRadiusPx = deflectionRaw(stepNum) * scaleFactor

        Set arcShape = DrawBucklingArc(arcRadiusPx, plateWidthPx, stepNum)

        Dim deflectionMM As Double
        deflectionMM = (deflectionRaw(stepNum) / maxDeflectionRaw) * 20

        Range("A6").value = "Step: " & stepNum & "/" & totalSteps
        Range("A7").value = "Load: " & Format(loadN / 1000, "0.0") & " kN"
        Range("A8").value = "Deflection: " & Format(deflectionMM, "0.00") & " mm"

        Range("E" & stepNum + 1).value = loadN / 9810 ' convert N to tons
        Range("F" & stepNum + 1).value = deflectionMM

        DoEvents
        WaitMilliseconds 50

        If stepNum > 1 Then
            On Error Resume Next
            ActiveSheet.Shapes("BucklingArc" & (stepNum - 1)).Delete
            On Error GoTo 0
        End If
    Next stepNum

    CreateDeflectionChart
End Sub

Function DrawBucklingArc(radius As Double, widthPx As Double, stepNum As Integer) As Shape
    Dim cx As Double, cy As Double
    Dim arcLeft As Double, arcTop As Double
    Dim arcWidth As Double, arcHeight As Double

    cx = 100 + widthPx / 2
    cy = 300
    arcWidth = widthPx
    arcHeight = radius * 2
    arcLeft = cx - arcWidth / 2
    arcTop = cy - radius

    Set DrawBucklingArc = ActiveSheet.Shapes.AddShape(msoShapeArc, arcLeft, arcTop, arcWidth, arcHeight)
    With DrawBucklingArc
        .Name = "BucklingArc" & stepNum
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = 3
        .Adjustments.Item(1) = 0
        .Adjustments.Item(2) = 180
    End With
End Function

Sub CreateDeflectionChart()
    Dim chObj As ChartObject
    Set chObj = ActiveSheet.ChartObjects.Add(Left:=500, Top:=50, width:=400, Height:=300)
    chObj.Name = "DeflectionChart"
    With chObj.Chart
        .ChartType = xlXYScatterSmooth
        .SetSourceData Source:=Range("E1:F" & totalSteps + 1)
        .HasTitle = True
        .ChartTitle.Text = "Load vs Approximate Deflection"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Load (tons)"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Deflection (mm approx)"
    End With
End Sub

Sub WaitMilliseconds(ms As Long)
    Dim endTime As Double
    endTime = Timer + ms / 1000#
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

