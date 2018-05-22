Attribute VB_Name = "Pricer"
' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' FRNPrice, BondPrice, BSPrice functions
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


Function BSPrice(flavor As String, S As Double, period As Double, T As Double, r As Double, sigma As Double, K As Double, q As Double) As Double
' ==================================================================================================
' ARGUMENTS:
' flavor    --string, call or put
' S         --double, underlying stock spot price
' period    --double, current period
' T         --double, maturity date, s.t. time to maturity=T-priod
' r         --double, risk free rate, used as continuous discount rate
' sigma     --double, vol of underlying stock price
' K         --double, strike price
' q         --double, divident yield, used as equivalent dividnet payment rate
' RETURNS:
' BSPrice   --double, option price due to BS Pricing Formula
' ==================================================================================================

' Preparation Phrase
Dim d1 As Double, d2 As Double
Dim price As Double
price = 0#

' Handling Phrase
d1 = WorksheetFunction.Ln(S / K) + (r - q + 0.5 * sigma ^ 2)
d1 = d1 / (sigma * (T - period) ^ (1 / 2))
d2 = d1 - sigma * (T - period) ^ (1 / 2)

If LCase(flavor) = "call" Or LCase(flavor) = "c" Then
    price = S * Exp(-q * (T - period)) * WorksheetFunction.NormDist(d1, 0, 1, True)
    price = price - K * Exp(-r * (T - period)) * WorksheetFunction.NormDist(d2, 0, 1, True)
Else
    price = K * WorksheetFunction.Exp(-r * (T - period)) * WorksheetFunction.NORM.DIST(-d2, 0, 1, True)
    price = price - S * WorksheetFunction.Exp(-q * (T - period)) * WorksheetFunction.NORM.DIST(-d1, 0, 1, True)
End If

' Checking Phrase
BSPrice = price
End Function

Function BSPrice_helper() As String

BSPrice_helper = "flavor As String, S As Double, period As Double, T As Double, r As Double, sigma As Double, K As Double, q As Double"

End Function


Function BondPrice(discountRate As Double, nper As Integer, freq As Double, couponRate As Double, notional As Double) As Double
' ==================================================================================================
' ARGUMENTS:
' discountRate  --double, discount rate used to discount CF to PV
' nper          --integer, number of payments
' freq          --double, frequency of payments, s.t. maturity=nper*freq
' couponRate    --double, coupon rate used to calculate coupon
' notional      --double, notional amount
' RETURNS:
' BondPrice     --double, price of the bond
' ==================================================================================================

' Preparation Phrase
Dim pv As Double
Dim pmt As Double
Dim discountFactor As Double
Dim maturity As Double
Dim i As Integer
pv = 0#
discountFactor = 1#
pmt = notional * couponRate / freq

' Handling Phrase
For i = 1 To nper
    discountFactor = (1 + discountRate / freq) ^ (-i)
    pv = pv + pmt * discountFactor
Next i

pv = pv + notional * discountFactor

' Checking Phrase

BondPrice = pv
End Function

Function BondPrice_helper() As String

BondPrice_helper = "discountRate as double, nper as integer, freq as double, couponRate as double, notional as double"

End Function


