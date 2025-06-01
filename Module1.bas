Attribute VB_Name = "Module1"
Option Explicit

' === Plate Properties ===
Const plateLengthMM = 500
Const plateWidthMM = 400
Const plateThicknessMM = 80
Const totalLoadN = 5000000 ' 500 tons

' === Grid Visualization ===
Const rows = 20   ' Y direction
Const cols = 25   ' X direction
Const maxSteps = 100

Dim deflection(cols, rows) As Double

Sub SimulateBuckling()
    Dim i As Integer, j As Integer, t As Integer
    Dim xRatio As Double, yRatio As Double
    Dim ampl As Double
    Dim loadFraction As Double
    Dim cellValue As String
    
    ClearVisualization
    InitializeGrid

    For t = 1 To maxSteps
        loadFraction = t / maxSteps
        ampl = 10 * loadFraction ' Max amplitude = 10 units

        For i = 1 To cols
            For j = 1 To rows
                xRatio = i / cols
                yRatio = j / rows

                ' Sinusoidal deformation (first buckling mode)
                deflection(i, j) = ampl * Sin(Application.WorksheetFunction.Pi() * xRatio) * Sin(Application.WorksheetFunction.Pi() * yRatio)
                
                ' Map deflection to color
                Call SetCellColor(i, j, deflection(i, j))
            Next j
        Next i

        ' Display time/load
        Range("A25").value = "Step: " & t & "/" & maxSteps
        Range("A26").value = "Load: " & Round(loadFraction * totalLoadN / 1000, 0) & " kN"
        
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01")) ' 0.1 second delay
    Next t
End Sub

Sub InitializeGrid()
    Dim i As Integer, j As Integer
    For i = 1 To cols
        For j = 1 To rows
            With Cells(j + 1, i + 1)
                .ColumnWidth = 2
                .RowHeight = 15
                .Interior.Color = RGB(200, 200, 255) ' Neutral blue
                .value = ""
            End With
        Next j
    Next i
    
    Range("A23").value = "Plate Buckling Simulation"
    Range("A24").value = "Plate: 500mm x 400mm, 80mm thick, Load = 500 tons"
End Sub

Sub SetCellColor(i As Integer, j As Integer, value As Double)
    Dim r As Integer, g As Integer, b As Integer
    Dim ratio As Double
    ratio = (value + 10) / 20 ' Normalize between 0 and 1

    ' Color mapping: Blue (low) to Red (high)
    If ratio < 0 Then ratio = 0
    If ratio > 1 Then ratio = 1
    r = 255 * ratio
    g = 0
    b = 255 * (1 - ratio)

    Cells(j + 1, i + 1).Interior.Color = RGB(r, g, b)
End Sub

Sub ClearVisualization()
    Dim i As Integer, j As Integer
    For i = 1 To cols
        For j = 1 To rows
            Cells(j + 1, i + 1).Interior.ColorIndex = xlNone
            Cells(j + 1, i + 1).value = ""
        Next j
    Next i
End Sub

