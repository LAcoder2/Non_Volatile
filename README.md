Let's start with this, but there are some ideas on how to improve it..
```vba
    [DllExport] 'Non-volatile analogue of the Offset sheet function
    Function Offset_nv(xRef As XLOPER12, xRowOff As XLOPER12, xColOff As XLOPER12, xHeight As XLOPER12, xWidth As XLOPER12) As LongPtr
        'Dim lret As Long = '_
        pExcel12p(xlfOffset, xRef, 5, xRef, xRowOff, xColOff, xHeight, xWidth)
    
        Return VarPtr(xRef)
    End Function
```
