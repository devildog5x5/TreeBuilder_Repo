Attribute VB_Name = "ParsingFunction"
'**************************************
'Name: Get_After_Comma
'Description:Use to parse comma delimited strings.
'Gets the value of a string after the intCommaNumber th comma
'By: Ian Ippolito(psc)
'Inputs:strString --string to search
'intCommaNumber=comma preceeding string
'i.e. intCommaNumber=2 will find the string after the 2nd comma
     'inputparms:
     'intCommaNumber = 0 to ?
     'strString = string to search
'Returns:returns value of string after the intCommaNumberth comma, or NOT FOUND if item doesn't exist
'Assumes: None
'Side Effects: None
'This code is copyrighted and has limited warranties.
'Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.14/lngWId.-1/qx/vb/scripts/ShowCode.htm
'for details.
'**************************************

Function Get_After_Comma(ByVal intCommaNumber As Integer, ByVal strString As String) As String
Dim intIndex As Integer
Dim intStartOfString As Integer
Dim intEndOfString As Integer
Dim boolNotFound As Integer
'check for intCommaNumber = 0--i.e. firs
     t one
If (intCommaNumber = 0) Then
Get_After_Comma = Left$(strString, InStr(strString, ",") - 1)
Else
     'not the first one
     'init start of string on first comma
intStartOfString = InStr(strString, ",")
'place start of string after intCommaNumber-th comma (-1 since
     'already did one
boolNotFound = 0
For intIndex = 1 To intCommaNumber - 1
     'get next comma
intStartOfString = InStr(intStartOfString + 1, strString, ",")
     'check for not found
If (intStartOfString = 0) Then
boolNotFound = 1
End If
Next intIndex
     'put start of string past 1st comma
intStartOfString = intStartOfString + 1
     'check for ending in a comma
If (intStartOfString > Len(strString)) Then
boolNotFound = 1
End If
If (boolNotFound = 1) Then
Get_After_Comma = "NOT FOUND"
Else
intEndOfString = InStr(intStartOfString, strString, ",")
'check for no second comma (i.e. end of string)
If (intEndOfString = 0) Then
intEndOfString = Len(strString) + 1
Else
intEndOfString = intEndOfString - 1
End If
Get_After_Comma = Mid$(strString, intStartOfString, intEndOfString - intStartOfString + 1)
End If
End If
End Function
