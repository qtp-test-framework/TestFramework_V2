'========== Declare and initialize the variables ==========
Dim sBatchSheetPath, chkBox
 
 
sBatchSheetPath = Wscript.Arguments(0)

chkBox = Wscript.Arguments(1)

a=Split(chkBox,"*")
for each x in a
    MsgBox "x = " & x
next
 
 'MsgBox "chkBox = " & chkBox
 