' Name: Beier (Benjamin) Liu
' Date:

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

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

