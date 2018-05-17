' Name: Beier (Benjamin) Liu
' Date:

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Function Interpolate(oldRates As Variant, freq As Double) As Variant
' ==================================================================================================
' ARGUMENTS:
' oldRates  --variant, a matrix of data input
' freq  --double, frequency of interpolation e.g. 1, 2, 0.25
' RETURNS:
' Interpolate --variant, a matrix of rates after interpolation
' ==================================================================================================

' ==================================================Preparation Phrase
Dim i As Integer, j As Integer, oldCount As Integer, Maturity As Integer
Maturity = oldRates(oldRates.Rows.Count, 1)
Maturity = Maturity * (freq * 1#)

Dim Mat As Variant
ReDim Mat(1 To Maturity, 1 To 2)

' ==================================================Handling Phrase
' Fills in the years
For i = 1 To Maturity
	Mat(i, 1) = i / (freq * 1#)
	Mat(i, 2) = 0
Next i

' Copies over known rates
For i = 1 To Maturity
	For j = 1 To oldRates.Rows.Count
		If oldRates(j, 1) = i / (freq * 1#) Then
			Mat(i, 2) = oldRates(j, 2)
		End If
	Next j
Next i

' Rates interpolation
For i = 2 To Maturity
	If Mat(i, 2) = 0 Then
		' Find the next filled rate to interpolate with
		For j = i + 1 To Maturity
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