Attribute VB_Name = "Utils"
' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Interpolate, forwardRate, DBQuery functions
' ===================================================================================================

Function Interpolate(oldRates As Variant, freq As Double) As Variant
' ==================================================================================================
' ARGUMENTS:
' oldRates  --variant, a matrix of data input
' RETURNS:
' Interpolate --variant, a matrix of rates after interpolation
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



Function forwardRate(freq As Double, date1 As Double, date2 As Double, rate1 As Double, rate2 As Double) As Double
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
forwardRate = (compoundrate2 / CompoundRate1 - 1) ^ (1# / deltaperiod)
forwardRate = forwardRate * freq
' Checking Phrase


End Function

Function ForwardRate_helper() As String

ForwardRate_helper = "freq as Double, date1 as string, date2 as string, rate1 as double, rate2 as double"

End Function


Sub DBQuery(path As String, strSQL As String, outputCell As String)
' ======================================================================================================
' ARGUMENTS:
' path        --string, contains the folder path for the Access database
' strSQL         --string, the SQL sentense
' outputCell    --string, the start cell of output
' RETURNS:
'
' USAGES:
' Qurey a database with strSQL
'======================================================================================================

' Preparation Phrase
Dim strFile As String
Dim strCon As String
Dim cn, rs As Object
strFile = path & "\MarketData.accdb"
strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strFile

Set cn = CreateObject("ADODB.Connection")
cn.Open strCon
Set rs = CreateObject("ADODB.RECORDSET")
rs.activeconnection = cn

' Handling Phrase
rs.Open strSQL
Sheets("Market Data").Range(outputCell).CopyFromRecordset rs

' Checking Phrase
rs.Close
cn.Close
Set cn = Nothing
End Sub



Function DBQuery_helper() As String

SQL_helper = "Qurey a database with keyword and threshold"

End Function
