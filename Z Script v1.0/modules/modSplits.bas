Attribute VB_Name = "modSplits"
Public vCom1 As String
Public vCom2 As String
Public vCom3 As String
Public vCom4 As String
Public vCom5 As String
Public vCom6 As String

Public vSpc1 As String
Public vSpc2 As String
Public vSpc3 As String
Public vSpc4 As String
Public vSpc5 As String

Public vEqu1 As String
Public vEqu2 As String
Public vEqu3 As String

Public vBrack As String

Public vStp1 As String
Public vStp2 As String
Public vStp3 As String


Public Function CommaSplit(vString As String)

vCurStep = 0
r = Split(vString, ", ")

For i = 0 To UBound(r)
    vCurStep = vCurStep + 1
    
    If vCurStep = 1 Then vCom1 = r(i)
    If vCurStep = 2 Then vCom2 = r(i)
    If vCurStep = 3 Then vCom3 = r(i)
    If vCurStep = 4 Then vCom4 = r(i)
    If vCurStep = 5 Then vCom5 = r(i)
    If vCurStep = 6 Then vCom6 = r(i)
    
Next i


End Function


Public Function SpaceSplit(vString As String)
vCurStep = 0
r = Split(vString, " ")

For i = 0 To UBound(r)
    vCurStep = vCurStep + 1
    
    If vCurStep = 1 Then vSpc1 = r(i)
    If vCurStep = 2 Then vSpc2 = r(i)
    If vCurStep = 3 Then vSpc3 = r(i)
    If vCurStep = 4 Then vSpc4 = r(i)
    If vCurStep = 5 Then vSpc5 = r(i)
    
Next i
End Function

Public Function FullstopSplit(vString As String)
vCurStep = 0
r = Split(vString, ".")

For i = 0 To UBound(r)
    vCurStep = vCurStep + 1
    
    If vCurStep = 1 Then vStp1 = r(i)
    If vCurStep = 2 Then vStp2 = r(i)
    If vCurStep = 3 Then vStp3 = r(i)
    
Next i
End Function

Public Function EqualsSplit(vString As String)
vCurStep = 0
r = Split(vString, " = ")

For i = 0 To UBound(r)
    vCurStep = vCurStep + 1
    
    If vCurStep = 1 Then vEqu1 = r(i)
    If vCurStep = 2 Then vEqu2 = r(i)
    If vCurStep = 3 Then vEqu3 = r(i)
    If vCurStep = 4 Then vEqu4 = r(i)
    
Next i
End Function

Public Function BracketSplit(vString As String)
vCurStep = 0
r = Split(vString, "(")

For i = 0 To UBound(r)
vCurStep = vCurStep + 1
    If vCurStep = 2 Then vString = r(i)
Next i

vCurStep = 0
r = Split(vString, ")")

For i = 0 To UBound(r)
vCurStep = vCurStep + 1
 If vCurStep = 1 Then vString = r(i)
Next i
vBrack = vString
End Function



Public Function CleanVars()
vBrack = ""

vCom1 = ""
vCom2 = ""
vCom3 = ""
vCom4 = ""
vCom5 = ""
vCom6 = ""

vSpc1 = ""
vSpc2 = ""
vSpc3 = ""
vSpc4 = ""
vSpc5 = ""

vEqu1 = ""
vEqu2 = ""
vEqu3 = ""
vEqu4 = ""

vStp1 = ""
vStp2 = ""
vStp3 = ""

End Function








