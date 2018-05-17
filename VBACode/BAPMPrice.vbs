' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Function BAPMPrice(flavor As String, N As Integer, S As Double, period As Double, T As Double, r As Double, sigma As Double, K As Double, q As Double) As Double
' ==================================================================================================
' ARGUMENTS:
' flavor    --string, call or put
' N         --integer, number of subperiods between current period and maturity date
' S         --double, underlying stock spot price
' period    --double, current period
' T         --double, maturity date, s.t. time to maturity=T-priod
' r         --double, risk free rate, used as continuous discount rate
' sigma     --double, vol of underlying stock price
' K         --double, strike price
' q         --double, divident yield, used as equivalent dividnet payment rate
' RETURNS:
' BAPMPrice --double, option price due to Binominal tree pricing simulations
' ==================================================================================================

' Preparation Phrase
Dim i As Integer, j As Integer, steps As Integer
Dim deltaT As Double, u As Double, d As Double, Pu As Double, Pd As Double, DF As Double, EquiS As Double
deltaT = (T - period) / (N * 1#)
u = Exp(sigma * (deltaT) ^ (0.5))
d = 1 / u
Pu = Exp((r - q) * deltaT) - d
Pu = Pu / (u - d)
Pd = 1 - Pu
DF = Exp(-deltaT * r)
EquiS = S * Exp(-deltaT * N * q)
steps = N + 1 'Since we need one more step in our array to include the first node

Dim arr As Variant
ReDim arr(1 To steps)

' Handling Phrase
For i = 1 To steps
	arr(i) = EquiS * u ^ (steps - i)
	arr(i) = arr(i) * d ^ (i - 1)   'compute possible future stock price
	If LCase(flavor) = "call" Or LCase(flavor) = "c" Then
		arr(i) = WorksheetFunction.Max(arr(i) - K, 0#)
	Else
		arr(i) = WorksheetFunction.Max(K - arr(i), 0#)
	End If
Next i
	
steps = steps - 1

For j = steps - 1 To 1 Step -1
	For i = 1 To j
		arr(i) = (arr(i + 1) * Pd + arr(i) * Pu) * DF
	Next i
Next j


' Checking Phrase
BAPMPrice = arr(1)
End Function

Function BAPMPrice_helper() As String

BAPMPrice_helper = "flavor as string, N as integer, S as double, period as double, T as double, r as double, sigma as double, K as double, q as double"

End Function

Sub FindOptimalN()

'Preparation Phrase
dim N as integer
dim res as double,  S As Double, period As Double, T As Double, r As Double, sigma As Double, K As Double, q As Double
dim flavor As String
res=Sheets("Homework").Range("D20").Value
flavor=Sheets("Homework").Range("D11").Value
S=Sheets("Homework").Range("D12").Value
period=Sheets("Homework").Range("D13").Value
T=Sheets("Homework").Range("D14").Value
r=Sheets("Homework").Range("D15").Value
sigma=Sheets("Homework").Range("D16").Value
K=Sheets("Homework").Range("D17").Value
q=Sheets("Homework").Range("D18").Value

'Handling Phrase
For N = 20 to 1000 step 5
	If WorksheetFunction.ABS(BAPMPrice(flavor, N, S, period, T, r, sigma, K, q)-res)<0.005 then exit for 
Next N

'Checking Phrase
Sheets("Homework").Range("D22").Value=N
Sheets("Homework").Range("D22").Value=BAPMPrice(flavor, N, S, period, T, r, sigma, K, q)
End Sub