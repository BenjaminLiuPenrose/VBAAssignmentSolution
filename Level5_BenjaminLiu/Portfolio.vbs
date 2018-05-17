' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Class Portfolio
' ===================================================================================================

' Class member
Private Derivatives_ As Collection

' Getter and setter
Public Property Get Derivatives() As Collection
	Set Derivatives = Derivatives_
End Property

Public Property Set Derivatives(iDerivatives As Collection)
	Set Derivatives_ = iDerivatives
End Property

' Class method
Function Dates_OneClass(strInstrumentType As String) As Integer
' ==================================================================================================
' ARGUMENTS:
' strInstrumentType -- string, name of the Instrument Type
' RETURNS:
' Dates_OneClass    -- integer, number of datesof the Instrument Type
' ==================================================================================================

' Preparation Phrase
Dim Dates As Integer
Dim Der As Derivative
Dates = 0

' Handling Phrase
For Each Der In Derivatives
	If strInstrumentType = Der.InstrumentType Then
		Dates = Dates + 1
	End If
Next Der

' Checking Phrase
Dates_OneClass = Dates
End Function


Function GetData_OneClass(strInstrumentType As String) As Variant
' ==================================================================================================
' ARGUMENTS:
' strInstrumentType -- string, name of the Instrument Type
' RETURNS:
' GetData_OneClass  -- variant, a table of historical data of one Instrument Type
					' e.g. format Date| Value| Returns
' ==================================================================================================

' Preparation Phrase
Dim Mat As Variant
Dim i As Integer
Dim Der As Derivative
Dim Dates As Integer
Dim weight As Double
Dim typ As String
If strInstrumentType = "Portfolio" Then
	strInstrumentType = "Equity"
	typ = "Portfolio"
End If
Dates = Dates_OneClass(strInstrumentType)
ReDim Mat(1 To Dates, 1 To 3)                   ' Date, Value, Returns
i = 1

' Handling Phrase
For Each Der In Derivatives
	If strInstrumentType = Der.InstrumentType Then
		Mat(i, 1) = Der.COB
		Mat(i, 2) = Der.Value
		If typ = "Portfolio" Then
			weight = GetWeight_OneClass_OneDate(strInstrumentType, Der.COB)
			Mat(i, 2) = Der.Value / (1# * weight)
		End If
		i = i + 1
	End If
Next Der

For i = 2 To Dates
	Mat(i - 1, 3) = Mat(i - 1, 2) / Mat(i, 2) - 1
Next i


' Checking Phrase
GetData_OneClass = Mat
End Function

Function GetWeight_OneClass_OneDate(strInstrumentType As String, COB As Double) As Double
' ==================================================================================================
' ARGUMENTS:
' strInstrumentType             --string, the Instrument Type
' COB                           --double, the date
' RETURNS:
' GetWeight_OneClass_OneDate    --double, the weight of one asset at one date
' ==================================================================================================

' Preparation Phrase
Dim Der As Derivative
Dim totalWeight As Double
Dim marginalWeight As Double
Dim weight As Double
totalWeight = 0#

' Handling Phrase
For Each Der In Derivatives
	If COB = Der.COB Then totalWeight = totalWeight + Der.Value
	If COB = Der.COB And strInstrumentType = Der.InstrumentType Then marginalWeight = Der.Value
Next Der

' Checking Phrase
GetWeight_OneClass_OneDate = marginalWeight / totalWeight

End Function

Function VaR(strInstrumentType As String, pct As Double) As Double
' ==================================================================================================
' ARGUMENTS:
' strInstrumentType --string, the Instrument Type
' pct               --double, the percentage of the VaR, e.g. 0.01 for 99%
' RETURNS:
' VaR               --double, the VaR of one asset of percentage given
' ==================================================================================================

' Preparation Phrase
Dim returns As Variant
Dim Dates As Integer
Dim i As Integer
Dim Mat As Variant
Dim val As Double
Mat = GetData_OneClass(strInstrumentType)
If strInstrumentType = "Portfolio" Then strInstrumentType = "Equity"
Dates = Dates_OneClass(strInstrumentType)
ReDim returns(1 To Dates)

' Handling Phrase
For i = 2 To Dates
	returns(i) = Mat(i - 1, 3)
Next i

val = WorksheetFunction.Percentile(returns, pct)

' Checking Phrase
VaR = val
End Function
