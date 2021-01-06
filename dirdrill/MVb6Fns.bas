Attribute VB_Name = "MVb6Functions"
' *************************************************************************
'  Copyright ©2000-2005 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
'  Originally stolen, and heavily modified, from:
'    Implement VB6 string functions in VB5 (Improved Over MSKB Version)
'    http://www.freevbcode.com/ShowCode.Asp?ID=17
'    http://support.microsoft.com/default.aspx?scid=kb;en-us;188007
' *************************************************************************
'  VB6 functions for VB5...
'    Join          Joins an array of strings into one string.
'    Split         Split a string into a variant array.
'    InStrRev      Similar to InStr but searches from end of string.
'    Replace       To find a particular string and replace it.
'    Reverse       To reverse a string.
' *************************************************************************
Option Explicit

Public Function InStrRev(ByVal StringCheck As String, ByVal StringMatch As String, Optional Start As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
   Dim nPos As Long

   ' Start from end if negative.
   If Start < 0 Then
      Start = Len(StringCheck)
   End If

   ' Truncate StringCheck to maximum possible position.
   If Start < Len(StringCheck) Then
      StringCheck = Left$(StringCheck, Start)
   End If

   ' Find last occurance of StringMatch.
   nPos = InStr(1, StringCheck, StringMatch, Compare)
   Do While nPos
      InStrRev = nPos
      nPos = InStr(nPos + 1, StringCheck, StringMatch, Compare)
   Loop
End Function

Public Function Join(SourceArray() As String, Optional Delimiter As String = " ") As String
   Dim i As Long, n As Long
   Dim nDelimLen As Long

   ' Cache this value for frequent use.
   nDelimLen = Len(Delimiter)

   ' Determine final string length by
   ' examining all descriptors.
   For i = LBound(SourceArray) To UBound(SourceArray)
      n = n + Len(SourceArray(i))
   Next i

   ' Add required delimiters to overall length.
   If nDelimLen Then
      n = n + (nDelimLen * (UBound(SourceArray) - LBound(SourceArray)))
   End If

   ' Create buffer for results, and set initial
   ' position of injection pointer.
   Join = Space$(n)
   n = 1

   ' Inject all but last element, each followed
   ' by (optional) delimiter.
   For i = LBound(SourceArray) To UBound(SourceArray) - 1
      Mid$(Join, n) = SourceArray(i)
      n = n + Len(SourceArray(i))
      If nDelimLen Then
         Mid$(Join, n) = Delimiter
         n = n + nDelimLen
      End If
   Next i

   ' Inject final element, and return.
   Mid$(Join, n) = SourceArray(i)
End Function

Public Function Replace(ByVal Expression As String, ByVal Find As String, ByVal Replase As String, Optional Start As Long = 1, Optional Count As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
   Dim nC As Long, nPos As Long
   Dim nFindLen As Long, nReplaceLen As Long

   nFindLen = Len(Find)
   nReplaceLen = Len(Replase)

   If (Find <> "") And (Find <> Replase) Then
      nPos = InStr(Start, Expression, Find, Compare)
      Do While nPos
         nC = nC + 1
         Expression = Left(Expression, nPos - 1) & Replase & Mid(Expression, nPos + nFindLen)
         If Count <> -1 And nC >= Count Then Exit Do
         nPos = InStr(nPos + nReplaceLen, Expression, Find, Compare)
      Loop
   End If

   Replace = Expression
End Function

Public Function StrReverse(ByVal Expression As String) As String
   Dim i As Long, n As Long
   ' Just flop one character at a time...
   StrReverse = Space$(Len(Expression))
   For i = Len(Expression) To 1 Step -1
      n = n + 1
      Mid$(StrReverse, n, 1) = Mid$(Expression, i, 1)
   Next i
End Function

Public Function Split(ByVal Expression As String, Optional Delimiter As String = " ", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
   Dim nCount As Long
   Dim nPos As Long
   Dim nDelimLen As Long
   Dim nStart As Long
   Dim sRet() As String

   ' Special case #1, Limit=0.
   If Limit = 0 Then
      ' Return unbound Variant array.
      Split = Array()
      Exit Function
   End If

   ' Special case #2, no delimiter.
   nDelimLen = Len(Delimiter)
   If nDelimLen = 0 Then
      ' Return expression in single-element Variant array.
      Split = Array(Expression)
      Exit Function
   End If

   ' Always start at beginning of Expression.
   nStart = 1

   ' Find first delimiter instance.
   nPos = InStr(nStart, Expression, Delimiter, Compare)
   Do While nPos
      ' Extract this element into enlarged array.
      ReDim Preserve sRet(0 To nCount) As String
      ' Bail if we hit the limit, or increment
      ' to next search start position.
      If nCount + 1 = Limit Then
         sRet(nCount) = Mid$(Expression, nStart)
         Exit Do
      Else
         sRet(nCount) = Mid$(Expression, nStart, nPos - nStart)
         nStart = nPos + nDelimLen
      End If
      ' Increment element counter
      nCount = nCount + 1
      ' Find next delimiter instance.
      nPos = InStr(nStart, Expression, Delimiter, Compare)
   Loop

   ' Grab last element.
   ReDim Preserve sRet(0 To nCount) As String
   sRet(nCount) = Mid$(Expression, nStart)

   ' Assign results and return.
   Split = sRet
End Function


