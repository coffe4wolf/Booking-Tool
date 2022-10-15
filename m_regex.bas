Attribute VB_Name = "m_regex"
' Module-level declaration. The object will persist between calls
Public pCachedRegexes As Dictionary
 
Public Function GetRegex( _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As RegExp
     
    ' Ensure the dictionary has been initialized
    If pCachedRegexes Is Nothing Then Set pCachedRegexes = New Dictionary
     
    ' Build the unique key for the regex: a combination
    ' of the boolean properties and the pattern itself
    Dim rxKey As String
    rxKey = IIf(IgnoreCase, "1", "0") & _
            IIf(MultiLine, "1", "0") & _
            IIf(MatchGlobal, "1", "0") & _
            Pattern
             
    ' If the RegExp object doesn't already exist, create it
    If Not pCachedRegexes.Exists(rxKey) Then
        Dim oRegExp As New RegExp
        With oRegExp
            .Pattern = Pattern
            .IgnoreCase = IgnoreCase
            .MultiLine = MultiLine
            .Global = MatchGlobal
        End With
        Set pCachedRegexes(rxKey) = oRegExp
    End If
 
    ' Fetch and return the pre-compiled RegExp object
    Set GetRegex = pCachedRegexes(rxKey)
End Function

Public Function RxTest( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Boolean
 
    ' Wow, that was easy:
    RxTest = GetRegex(Pattern, IgnoreCase, MultiLine, False).test(SourceString)
     
End Function

Public Function RxMatch( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant
 
    Dim oMatches As MatchCollection
    With GetRegex(Pattern, IgnoreCase, MultiLine, False)
        Set oMatches = .Execute(SourceString)
        If oMatches.Count > 0 Then
            RxMatch = oMatches(0).value
        Else
            RxMatch = ""
        End If
    End With
 
End Function

Public Function RxMatches( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As Variant
 
    Dim oMatch As Match
    Dim arrMatches
    Dim lngCount As Long
     
    arrMatches = Array()
    With GetRegex(Pattern, IgnoreCase, MultiLine, MatchGlobal)
        For Each oMatch In .Execute(SourceString)
            ReDim Preserve arrMatches(lngCount)
            arrMatches(lngCount) = oMatch.value
            lngCount = lngCount + 1
        Next
    End With
 
    RxMatches = arrMatches
End Function

Public Function RxReplace( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    ByVal ReplacePattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As String
 
    ' A single statement!
    RxReplace = GetRegex( _
        Pattern, IgnoreCase, MultiLine, MatchGlobal).Replace( _
        SourceString, ReplacePattern)
 
End Function



