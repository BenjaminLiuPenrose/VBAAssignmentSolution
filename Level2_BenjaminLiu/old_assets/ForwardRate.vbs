' Name: Beier (Benjamin) Liu
' Date: 

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Function ForwardRate(freq As Double, date1 As Double, date2 As Double, rate1 As Double, rate2 As Double) As Double
' ==================================================================================================
' ARGUMENTS:
' freq  --double, frequency of interpolation e.g. 1, 2, 0.25
' date1 --string, one date
' date2 --string, another different latter date
' rate1 --double, forward rate of date1
' rate2 --double, forward rate of date2
' RETURNS:
' ForwardRate   --double, computed value of implied forward rate
' ==================================================================================================

' Preparation Phrase
Dim CompoundRate1 As Double, compoundrate2 As Double, deltaperiod As Double

' Handling Phrase
compoundrate2 = (1 + rate2 / freq) ^ (freq * date2)
CompoundRate1 = (1 + rate1 / freq) ^ (freq * date1)
deltaperiod = freq * (date2 - date1)
'CompoundRate2=CompoundRate1*(1+ForwardRate/freq)^(DeltaPeriod)
ForwardRate = (compoundrate2 / CompoundRate1 - 1) ^ (1# / deltaperiod)
ForwardRate = ForwardRate * freq
' Checking Phrase


End Function

Function ForwardRate_helper() As String

ForwardRate_helper = "freq as Double, date1 as string, date2 as string, rate1 as double, rate2 as double"

End Function
