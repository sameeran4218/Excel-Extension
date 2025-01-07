
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegressionMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub MyMacro(ByRef control As Office.IRibbonControl)
    Dim n As Integer
    Dim m As Integer
    
    ' Prompt user for input and output columns
    Dim xColumn As String, yColumn As String
    On Error GoTo UserCancel
    xColumn = InputBox("Enter the column letter for X (independent variable):", "Select X Column")
    If xColumn = "" Then Exit Sub ' Exit if user cancels
    yColumn = InputBox("Enter the column letter for Y (dependent variable):", "Select Y Column")
    If yColumn = "" Then Exit Sub ' Exit if user cancels
    
    ' Validate column inputs
    If Not IsNumeric(Cells(2, Columns(xColumn).Column).Value) Or Not IsNumeric(Cells(2, Columns(yColumn).Column).Value) Then
        MsgBox "The selected columns contain non-numeric data. Please select columns with numeric data only.", vbExclamation
        Exit Sub
    End If
    
    ' Determine last row and column
    On Error GoTo DataError
    n = ActiveSheet.Cells(Rows.Count, xColumn).End(xlUp).Row - 1
    If n < 1 Then
        MsgBox "The selected column does not contain enough data. Please check the input.", vbExclamation
        Exit Sub
    End If
    m = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Create vectors
    Dim x() As Double
    Dim y() As Double
    Dim z() As Double
    ReDim x(1 To n)
    ReDim y(1 To n)
    ReDim z(1 To n)

    ' Input x & y Points
    Dim i As Integer
    For i = 1 To n
        If IsEmpty(Cells(1 + i, Columns(xColumn).Column)) Or IsEmpty(Cells(1 + i, Columns(yColumn).Column)) Then
            MsgBox "Missing value detected in row " & (1 + i) & ". Please fill all rows in the selected columns.", vbExclamation
            Exit Sub
        End If
        x(i) = Cells(1 + i, Columns(xColumn).Column)
        y(i) = Cells(1 + i, Columns(yColumn).Column)
    Next i

    ' Create Sums
    Dim sumx As Double, Sumy As Double, sumxy As Double, sumx2 As Double
    sumx = 0
    Sumy = 0
    sumxy = 0
    sumx2 = 0

    For i = 1 To n
        sumx = sumx + x(i)
        Sumy = Sumy + y(i)
        sumxy = sumxy + x(i) * y(i)
        sumx2 = sumx2 + x(i) * x(i)
    Next i

    ' Calculating a1
    Dim a1 As Double
    a1 = (n * sumxy - sumx * Sumy) / (n * sumx2 - (sumx) ^ 2)

    ' Calculating a0
    Dim a0 As Double
    Dim ym As Double, xm As Double
    ym = Sumy / n
    xm = sumx / n
    a0 = ym - a1 * xm

    ' Calculating r^2
    Dim r2 As Double
    Dim st As Double, sr As Double
    st = 0
    sr = 0

    For i = 1 To n
        st = st + (y(i) - ym) ^ 2
        sr = sr + (y(i) - a0 - a1 * x(i)) ^ 2
    Next i

    r2 = (st - sr) / st

    ' Display Data
    ActiveSheet.Cells(7, m + 9).Value = "Slope :"
    ActiveSheet.Cells(8, m + 9).Value = "Intercept :"
    ActiveSheet.Cells(32, m + 9).Value = "R2 Score :"
    ActiveSheet.Cells(7, m + 10).Value = a1
    ActiveSheet.Cells(8, m + 10).Value = a0
    ActiveSheet.Cells(32, m + 10).Value = r2
    
    ActiveSheet.Cells(7, m + 9).Font.Bold = True
    ActiveSheet.Cells(8, m + 9).Font.Bold = True
    ActiveSheet.Cells(32, m + 9).Font.Bold = True
    ActiveSheet.Cells(7, m + 10).Font.Bold = True
    ActiveSheet.Cells(8, m + 10).Font.Bold = True
    ActiveSheet.Cells(32, m + 10).Font.Bold = True
    

    Cells(1, m + 1) = "Predicted Data" ' Label for the column
    ActiveSheet.Cells(1, m + 1).Font.Bold = True
    For i = 1 To n
        Cells(1 + i, m + 1) = a0 + a1 * x(i)
    Next i
    For i = 1 To n
        z(i) = Cells(1 + i, m + 1)
    Next i

    ' Plot Graph
    Call CreateRegressionChart(n, m, x, y, z, a0, a1, xColumn, yColumn)

    MsgBox "Regression calculation and graph plotting completed successfully!"
    Exit Sub

UserCancel:
    MsgBox "Operation cancelled by the user.", vbExclamation
    Exit Sub

DataError:
    MsgBox "An error occurred while processing the data. Please check your inputs.", vbCritical
    Exit Sub
End Sub

Private Sub CreateRegressionChart(n As Integer, m As Integer, x() As Double, y() As Double, z() As Double, a0 As Double, a1 As Double, xColumn As String, yColumn As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim chtObj As ChartObject
    Set chtObj = ws.ChartObjects.Add(ws.Cells(1, m + 5).Left, ws.Cells(10, 1).Top, 500, 300)
    Dim ch As Chart
    Set ch = chtObj.Chart
    Dim xRange As Range, yRange As Range, zRange As Range
    Dim zColumn As String
    
    Set xRange = Range(xColumn & "2:" & xColumn & n + 1)
    Set yRange = Range(yColumn & "2:" & yColumn & n + 1)
    Set zRange = Range(Cells(2, m + 1), Cells(n + 1, m + 1))

    Dim ser As Series
    Set ser = ch.SeriesCollection.NewSeries
    ser.Values = yRange
    ser.XValues = xRange
    ser.ChartType = xlXYScatter
    ser.Name = "Actual Values"

    Dim ser2 As Series
    Set ser2 = ch.SeriesCollection.NewSeries
    ser2.Values = zRange
    ser2.XValues = xRange
    ser2.ChartType = xlXYScatterLines
    ser2.MarkerStyle = xlMarkerStyleNone
    ser2.Name = "Regression Line"

    With ch
        ' Adding axis titles
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "X-Axis"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Y-Axis"
    End With
End Sub
