Attribute VB_Name = "Run"
' Name: Beier (Benjamin) Liu
' Date:

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================


Function PopulateArray() As Collection
' ==================================================================================================
' ARGUMENTS:
'
' RETURNS:
' PopulateArray     --Collection, collection of Derivatives
' ==================================================================================================

' Preparation Phrase
Dim cDers As Collection
Dim Der As Derivative
Set cDers = New Collection

Dim i As Integer
Dim rowCount As Integer
rowCount = Sheets("Output").Cells(Rows.Count, "B").End(xlUp).Row

' Handling Phrase
For i = 2 To rowCount
    Set Der = New Derivative
    Der.InstrumentType = Sheets("Output").Cells(i, 3)
    Der.COB = Sheets("Output").Cells(i, 1)
    Der.Value = Sheets("Output").Cells(i, 4)
    cDers.Add Der
Next i

' Checking Phrase
Set PopulateArray = cDers
End Function

Sub GetOutput()
' ==================================================================================================
' ARGUMENTS:
'
' RETURNS:
'
' OPERATIONS:
' Create a chart with specific format
' ==================================================================================================

' Preparation Phrase
Dim i As Integer
Dim path As String, strSQL As String, outputCell As String
Dim ticker As String
Dim rowCount As Integer
Dim cnt As Integer
cnt = 0

Dim freq As Double, discountRate As Double, couponMargin As Double, couponRate As Double, notional As Double
Dim discountRates As Variant, forwardRates As Variant
Dim nper As Integer
rowCount = Sheets("output").Cells(Rows.Count, "B").End(xlUp).Row

' Handling Phrase
For i = 2 To 502
    If Sheets("Output").Range("C" & i).Value = "Equity" Then
        If cnt = 0 Then 'If the first time to price the stock
            ticker = WorksheetFunction.Index(Range("Ticker"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
            path = Sheets("Input").Range("M2").Value
            strSQL = "SELECT Date, " & Chr(39) & ticker & Chr(39) & ", Close_Price FROM " & ticker & " ORDER BY Date DESC"
            outputCell = "A2"
            Utils.DBQuery path, strSQL, outputCell
        End If
        Sheets("Output").Range("D" & i).Value = WorksheetFunction.Index(Range("Close_Price"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("Stock_Date"), -1))  'Temporary fixed for non-trading date, so choose 1
        cnt = cnt + 1
    End If
    If Sheets("Output").Range("C" & i).Value = "Bond" Then
        discountRate = WorksheetFunction.Index(Range("Discount_Rate"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("Riskfree_Date"), 0))
        couponRate = WorksheetFunction.Index(Range("Coupon"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        freq = 2#
        nper = 0.5 * freq
        notional = WorksheetFunction.Index(Range("Notional"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        ' Sheets("Output").Range("D" & i).Value = Pricer.BondPrice(discountRate, nper, freq, couponRate, notional)
        Sheets("Output").Range("D" & i).Value = WorksheetFunction.pv(discountRate, nper, -couponRate * notional, -notional)
        cnt = 0
    End If
    If Sheets("Output").Range("C" & i).Value = "Floater" Then
        couponMargin = WorksheetFunction.Index(Range("Coupon_Margin"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        freq = 360
        notional = WorksheetFunction.Index(Range("Notional"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        Sheets("Output").Range("D" & i).Value = Pricer.FloatingRateNotePrice(WorksheetFunction.Index(Range("Yield"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("FRN_Date"), 1)), WorksheetFunction.Index(Range("Forward_Rate"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("FRN_Date"), 1)), freq, couponMargin, notional)
        cnt = 0
    End If
    If Sheets("Output").Range("C" & i).Value = "Cash" Then
        Sheets("Output").Range("D" & i).Value = WorksheetFunction.Index(Range("Quantity"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        cnt = 0
    End If
Next i


' Checking Phrase
End Sub


Sub GetReport()
' ==================================================================================================
' ARGUMENTS:
'
' RETURNS:
'
' OPERATIONS:
' compute portfolio and marginal VaR
' ==================================================================================================

' Preparation Phrase
Dim newPort As Portfolio
Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray()

' Handling Phrase
Sheets("Report").Range("E12").Value = newPort.VaR(Sheets("Report").Range("B12").Value, 0.01)
Sheets("Report").Range("E13").Value = newPort.VaR(Sheets("Report").Range("B13").Value, 0.01)
Sheets("Report").Range("E14").Value = newPort.VaR(Sheets("Report").Range("B14").Value, 0.01)
Sheets("Report").Range("E15").Value = newPort.VaR(Sheets("Report").Range("B15").Value, 0.01)
Sheets("Report").Range("E16").Value = newPort.VaR(Sheets("Report").Range("B16").Value, 0.01)
Sheets("Report").Range("E17").Value = newPort.VaR(Sheets("Report").Range("B17").Value, 0.01)

Sheets("Report").Range("D12").Value = newPort.retun(Sheets("Report").Range("B12").Value)
Sheets("Report").Range("D13").Value = newPort.retun(Sheets("Report").Range("B13").Value)
Sheets("Report").Range("D14").Value = newPort.retun(Sheets("Report").Range("B14").Value)
Sheets("Report").Range("D15").Value = newPort.retun(Sheets("Report").Range("B15").Value)
Sheets("Report").Range("D16").Value = newPort.retun(Sheets("Report").Range("B16").Value)
Sheets("Report").Range("D17").Value = newPort.retun(Sheets("Report").Range("B17").Value)


' Checking Phrase
End Sub

Sub Run()

GetOutput

GetReport

End Sub




