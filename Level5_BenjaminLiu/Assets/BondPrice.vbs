' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Function BondPrice(discountRate as double, nper as integer, freq as double, couponRate as double, notional as double) as double
' ==================================================================================================
' ARGUMENTS:
' discountRate	--double, discount rate used to discount CF to PV
' nper			--integer, number of payments
' freq			--double, frequency of payments, s.t. maturity=nper*freq
' couponRate	--double, coupon rate used to calculate coupon 
' notional		--double, notional amount
' RETURNS:
' BondPrice 	--double, price of the bond 
' ==================================================================================================

' Preparation Phrase
Dim pv As Double
Dim pmt As Double
Dim discountFactor As Double
Dim maturity As Double
Dim i As Integer
pv=0.0
discountFactor = 1.0
pmt = notional * couponMRate / freq

' Handling Phrase
for i =1 to nper
    discountFactor = (1 + discountRate / freq) ^ (-i)
    pv = pv + pmt * discountFactor
next i

pv = pv + notional * discountFactor

' Checking Phrase

BondPrice=pv 
End Function 

Function BondPrice_helper() as String

BondPrice_helper="discountRate as double, nper as integer, freq as double, couponRate as double, notional as double"

End Function 
