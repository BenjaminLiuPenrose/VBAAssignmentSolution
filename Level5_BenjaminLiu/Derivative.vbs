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
private Value_ as double
private InstrumentType_ as string
private COB_ as double

' Getter and setter
public property get Value() as double
	Value=Value_
end property

public property let Value(iValue as double)
	Value_=iValue
end property

public property get InstrumentType() as string
	InstrumentType=InstrumentType_
end property

public property let InstrumentType(iInstrumentType as string)
	InstrumentType_=iInstrumentType
end property

public property get COB() as double
	COB=COB_
end property

public property let COB(iCOB as double)
	COB_=iCOB
end property




