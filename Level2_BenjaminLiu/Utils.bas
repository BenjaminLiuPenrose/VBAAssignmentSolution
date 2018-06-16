Attribute VB_Name = "Utils"
' Name: Beier (Benjamin) Liu
' Date:

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' solution to homewoek 2
' ===================================================================================================

Function Interpolate(oldRates As Variant, freq As Double) As Variant
' ==================================================================================================
' ARGUMENTS:
' oldRates  --variant, a matrix of data input

' RETURNS:
' Interpolate --variant, a matrix of rates after interpolation (dates, rates)
' ==================================================================================================

' ==================================================Preparation Phrase
Dim i As Integer, j As Integer, oldCount As Integer, maturity As Integer
maturity = oldRates(oldRates.Rows.Count, 1)
maturity = maturity * (freq * 1#)

Dim Mat As Variant
ReDim Mat(1 To maturity, 1 To 2)

' ==================================================Handling Phrase
' Fills in the years
For i = 1 To maturity
    Mat(i, 1) = i / (freq * 1#)
    Mat(i, 2) = 0
Next i

' Copies over known rates
For i = 1 To maturity
    For j = 1 To oldRates.Rows.Count
        If oldRates(j, 1) = i / (freq * 1#) Then
            Mat(i, 2) = oldRates(j, 2)
        End If
    Next j
Next i

' Rates interpolation
For i = 2 To maturity
    If Mat(i, 2) = 0 Then
        ' Find the next filled rate to interpolate with
        For j = i + 1 To maturity
            If Mat(j, 2) <> 0 Then Exit For
        Next j

        Mat(i, 2) = Mat(i - 1, 2) + (Mat(j, 2) - Mat(i - 1, 2)) / (j - i + 1)
        
    End If
Next i

' ===============================================================Checking Phrase
Interpolate = Mat

End Function

Function Interpolate_helper() As String

Interpolate_helper = "oldRates as variant, freq as integer"

End Function



Function forwardRates(oldRates As Variant, freq As Double) As Variant
' ==================================================================================================
' ARGUMENTS:
' oldRates  --variant, a matrix of data input
' freq  --double, frequency of interpolation e.g. 1, 2, 0.25

' RETURNS:
' ForwardRates   --variant, computed value of matrix of implied forward rate (dates, rates)
' ==================================================================================================

' Preparation Phrase
Dim date1 As Double, date2 As Double, rate1 As Double, rate2 As Double
Dim compoundRate1 As Double, compoundRate2 As Double, deltaPeriod As Double

Dim yield As Variant
yield = Interpolate(oldRates, freq)

Dim i As Integer, maturity As Integer
maturity = UBound(yield, 1) - LBound(yield, 1) + 1

Dim Mat As Variant
ReDim Mat(1 To maturity, 1 To 2)

' Handling Phrase
Mat = yield

For i = 2 To maturity
    rate1 = yield(i - 1, 2)
    rate2 = yield(i, 2)
    date1 = yield(i - 1, 1)
    date2 = yield(i, 1)
    
    compoundRate2 = (1 + rate2 / freq) ^ (freq * date2)
    compoundRate1 = (1 + rate1 / freq) ^ (freq * date1)
    deltaPeriod = freq * (date2 - date1)
    'CompoundRate2=CompoundRate1*(1+ForwardRate/freq)^(DeltaPeriod)
    Mat(i, 2) = (compoundRate2 / compoundRate1 - 1) ^ (1# / deltaPeriod)
    Mat(i, 2) = Mat(i, 2) * freq
Next i

' Checking Phrase
forwardRates = Mat

End Function

Function ForwardRates_helper() As String

ForwardRates_helper = "oldRate as variant, freq as Double"

End Function


Function FRNPrice(oldRates As Variant, freq As Double, couponMargin As Double, notional As Double, nper As Integer) As Double
' ==================================================================================================
' ARGUMENTS:
' oldRates       --a matrix of data input
' freq          --double, frequency of payment
' couponMargin  --double, coupon rate=forwardRate+couponMargin
' notional      --double, notional amount
' nper          --integer, number of payment

' RETURNS:
' FRNPrice      --double, computed price of floating rate note
' ==================================================================================================

' Preparation Phrase
Dim pv As Double
Dim pmt As Double
Dim discountFactor As Double
Dim i As Integer

pv = 0#

Dim discountRates As Variant, forwRates As Variant
discountRates = Interpolate(oldRates, freq)
forwRates = forwardRates(oldRates, freq)

' Dim nper As Integer
' nper = UBound(discountRates, 1) - LBound(discountRates, 1) + 1

' Handling Phrase
For i = 1 To nper
    pmt = notional * (forwRates(i, 2) + couponMargin) / freq
    discountFactor = (1 + discountRates(i, 2) / freq) ^ (-i)
    pv = pv + pmt * discountFactor
Next i

pv = pv + notional * discountFactor

' Checking Phrase
FRNPrice = pv

End Function

Function FRNPrice_helper() As String

FRNPrice_helper = "oldRates As Variant, freq As Double, couponMargin As Double, notional As Double, nper As Integer"

End Function

