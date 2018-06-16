' Name: Beier (Benjamin) Liu
' Date: 5/14/2108

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Function FloatingRateNotePrice(discountRate As Variant, forwardRate As Variant, freq As Double, couponMargin As Double, notional As Double) As Double
' ==================================================================================================
' ARGUMENTS:
' discountRate  --range, array of money market rate
' forwardRate   --range, array of implied forward rate
' freq          --double, frequency of payment
' couponMargin  --double, coupon rate=forwardRate+couponMargin
' notional      --double, notional amount
' RETURNS:
' FloatingRateNotePrice_helper--double, computed price of floating rate note
' ==================================================================================================

' Preparation Phrase
Dim pv As Double
Dim pmt As Double
Dim discountFactor As Double
Dim maturity As Double
Dim i As Integer, nper As Integer
nper = discountRate.Rows.Count
maturity = nper / freq
pv = 0#

' Handling Phrase
For i = 1 To nper
    pmt = notional * (forwardRate(i, 1) + couponMargin) / freq
    discountFactor = (1 + discountRate(i, 1) / freq) ^ (-i)
    pv = pv + pmt * discountFactor
Next i

pv = pv + notional * discountFactor

' Checking Phrase
FloatingRateNotePrice = pv
End Function

Function FloatingRateNotePrice_helper() As String

FloatingRateNotePrice_helper = "discountRate as variant, forwardRate as variant, freq as double, couponMargin as double, notional as double"

End Function
