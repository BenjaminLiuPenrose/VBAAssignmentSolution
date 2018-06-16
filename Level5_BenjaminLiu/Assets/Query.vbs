' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Format selected range for "BIS OTC.xlsx"
' ===================================================================================================


Sub DBQuery1(thresInput As Double, keyInput As String, range As String, path As String)
' ======================================================================================================
' ARGUMENTS:
' thresInput  --double is the threshold for market cap
' keyInput    --string is the search keyword for s_holder
' range       --string is either column s_holder or s_holder_name
' path        --string contains the path for the access database
' RETURNS:
'
' USAGES:
' Qurey a database with keyword and threshold
'======================================================================================================

' Preparation Phrase==================================================
Dim strFile As String
Dim strCon As String
Dim strSQL As String
Dim strThres, strKey As String
' valInput depreciated
' Dim valInput As String
Dim cn, rs As Object

' strFile depreciated
' strFile = "C:\Users\liubeier\Desktop\¸ß·çÏÕ¹ÉÆ±¹Ø×¢Ãûµ¥\test.accdb"
strFile = path & "\AllAShares.accdb"
strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strFile

Set cn = CreateObject("ADODB.Connection")
cn.Open strCon

' Handling Phrase=======================================================
' valInput depreciated
' valInput = InputBox("Input")
thresInput = thresInput * 100000000
strThres = "liq_cap<=" & thresInput
strKey = range & " LIKE " & Chr(34) & "%" & keyInput & "%" & Chr(34)
strSQL = "SELECT trade_code, s_name, s_holder_name FROM AllAShares_dynamic WHERE " & strThres & " AND " & strKey
' strSQL depreciated
' strSQL = "select s_name from AllAShares"

Set rs = CreateObject("ADODB.RECORDSET")
rs.activeconnection = cn
rs.Open strSQL
Sheets("Res").range("A21").CopyFromRecordset rs

' Checking Phrase=======================================================
rs.Close
cn.Close
Set cn = Nothing
End Sub

Sub DBQuery2 (path As String, strSQL As String, outputCell As String)
' ======================================================================================================
' ARGUMENTS:
' path        --string, contains the folder path for the Access database
' strSQL		 --string, the SQL sentense
' outputCell	--string, the start cell of output 
' RETURNS:	
'
' USAGES:
' Qurey a database with strSQL
'======================================================================================================

' Preparation Phrase
Dim strFile As String
Dim strCon As String
Dim cn, rs As Object
strFile = path & "\AllAShares.accdb"
strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strFile

Set cn = CreateObject("ADODB.Connection")
cn.Open strCon
Set rs = CreateObject("ADODB.RECORDSET")
rs.activeconnection = cn

' Handling Phrase
rs.Open strSQL
Sheets("Res").range("A21").CopyFromRecordset rs

' Checking Phrase
rs.Close
cn.Close
Set cn = Nothing
End Sub



Function DBQuery_helper() As String

SQL_helper = "Qurey a database with keyword and threshold"

End Function




module genStrSQL

module genPath

sub Finder


