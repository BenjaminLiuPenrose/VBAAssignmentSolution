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

' Handling Phrase
For i = 5 To 64
    Set Der = New Derivative
    Der.InstrumentType = Sheets("MarketData").Cells(i, 2)
    Der.COB = Sheets("MarketData").Cells(i, 3)
    Der.Value = Sheets("MarketData").Cells(i, 4)
    cDers.Add Der
Next i

' Checking Phrase
Set PopulateArray = cDers
End Function


Function PopulateArray2(remove As String) As Collection
' ==================================================================================================
' ARGUMENTS:
' remove 			--string, name of the asset to be removed 
' RETURNS:
' PopulateArray     --Collection, collection of Derivatives
' ==================================================================================================

' Preparation Phrase
Dim cDers As Collection
Dim Der As Derivative
Set cDers = New Collection

Dim i As Integer

' Handling Phrase
if remove <>"Equity" then
	For i = 5 To 14
		Set Der = New Derivative
		Der.InstrumentType = Sheets("MarketData").Cells(i, 2)
		Der.COB = Sheets("MarketData").Cells(i, 3)
		Der.Value = Sheets("MarketData").Cells(i, 4)
		cDers.Add Der
	Next i
end if

if remove <>"Commodity" then
	For i = 15 To 24
		Set Der = New Derivative
		Der.InstrumentType = Sheets("MarketData").Cells(i, 2)
		Der.COB = Sheets("MarketData").Cells(i, 3)
		Der.Value = Sheets("MarketData").Cells(i, 4)
		cDers.Add Der
	Next i
end if 

if remove <>"Fixed Income" then
	For i = 25 To 34
		Set Der = New Derivative
		Der.InstrumentType = Sheets("MarketData").Cells(i, 2)
		Der.COB = Sheets("MarketData").Cells(i, 3)
		Der.Value = Sheets("MarketData").Cells(i, 4)
		cDers.Add Der
	Next i
end if 

if remove <>"CDS" then 
	For i = 35 To 44
		Set Der = New Derivative
		Der.InstrumentType = Sheets("MarketData").Cells(i, 2)
		Der.COB = Sheets("MarketData").Cells(i, 3)
		Der.Value = Sheets("MarketData").Cells(i, 4)
		cDers.Add Der
	Next i
end if 

if remove <>"Futures" then
	For i = 45 To 54
		Set Der = New Derivative
		Der.InstrumentType = Sheets("MarketData").Cells(i, 2)
		Der.COB = Sheets("MarketData").Cells(i, 3)
		Der.Value = Sheets("MarketData").Cells(i, 4)
		cDers.Add Der
	Next i
end if 

if remove <>"FX" then
	For i = 55 To 64
		Set Der = New Derivative
		Der.InstrumentType = Sheets("MarketData").Cells(i, 2)
		Der.COB = Sheets("MarketData").Cells(i, 3)
		Der.Value = Sheets("MarketData").Cells(i, 4)
		cDers.Add Der
	Next i
end if 

' Checking Phrase
Set PopulateArray2 = cDers
End Function


Sub TestAB()
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

 ' Dim i As Integer, j As Integer
 ' Dim Mat As Variant
 ' Dim Dates As Integer
 ' Mat = newPort.GetData_OneClass("Equity")
 ' Dates = newPort.Dates_OneClass("Equity")


' Handling Phrase
Sheets("Homework").Range("C16").Value = newPort.VaR("Portfolio", 0.01)
Sheets("Homework").Range("F16").Value = newPort.VaR("Equity", 0.01)
Sheets("Homework").Range("I16").Value = newPort.VaR("Commodity", 0.01)
Sheets("Homework").Range("L16").Value = newPort.VaR("Fixed Income", 0.01)
Sheets("Homework").Range("O16").Value = newPort.VaR("CDS", 0.01)
Sheets("Homework").Range("R16").Value = newPort.VaR("Futures", 0.01)

 ' For i = 1 To Dates
 '   For j = 1 To 3
	'    Range("Output").Cells(i, j).Value = Mat(i, j)
 '   Next j
 ' Next i

 ' Sheets("MarketData").Range("O16").Value = newPort.GetWeight_OneClass_OneDate("Equity", 43235#)
' Checking Phrase
End Sub

Sub TestCD ()
' ==================================================================================================
' ARGUMENTS:
'
' RETURNS:
'
' OPERATIONS:
' Remove Instrument with best return or VaR 
' ==================================================================================================

' Preparation Phrase
Dim newPort As Portfolio

Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray2("Equity")
Sheets("Homework").Range("G21").Value = newPort.retun("Portfolio")
Sheets("Homework").Range("H21").Value = newPort.VaR("Portfolio", 0.01)

Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray2("Commodity")
Sheets("Homework").Range("G22").Value = newPort.retun("Portfolio")
Sheets("Homework").Range("H22").Value = newPort.VaR("Portfolio", 0.01)

Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray2("Fixed Income")
Sheets("Homework").Range("G23").Value = newPort.retun("Portfolio")
Sheets("Homework").Range("H23").Value = newPort.VaR("Portfolio", 0.01)

Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray2("CDS")
Sheets("Homework").Range("G24").Value = newPort.retun("Portfolio")
Sheets("Homework").Range("H24").Value = newPort.VaR("Portfolio", 0.01)

Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray2("Futures")
Sheets("Homework").Range("G25").Value = newPort.retun("Portfolio")
Sheets("Homework").Range("H25").Value = newPort.VaR("Portfolio", 0.01)

' Handling Phrase
' Checking Phrase

End Sub



