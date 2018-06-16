VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Derivative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Class Derivative
' ===================================================================================================

' Class member
Private Value_ As Double
Private InstrumentType_ As String
Private COB_ As Double

' Getter and setter
Public Property Get Value() As Double
    Value = Value_
End Property

Public Property Let Value(iValue As Double)
    Value_ = iValue
End Property

Public Property Get InstrumentType() As String
    InstrumentType = InstrumentType_
End Property

Public Property Let InstrumentType(iInstrumentType As String)
    InstrumentType_ = iInstrumentType
End Property

Public Property Get COB() As Double
    COB = COB_
End Property

Public Property Let COB(iCOB As Double)
    COB_ = iCOB
End Property
