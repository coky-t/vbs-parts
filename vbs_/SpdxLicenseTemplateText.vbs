'Attribute VB_Name = "SpdxLicenseTemplateText"
Option Explicit

'
' Copyright (c) 2020,2023,2024 Koki Takeyama
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included
' in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'

'
' SPDX License List Matching Guidelines, v2.3
' https://spdx.github.io/spdx-spec/v2.3/license-matching-guidelines-and-templates/
'

Public Function GetMatchingLines(TemplateText)
    
    ' TemplateTextArray
    
    Dim TemplateTextArray
    TemplateTextArray = Split(Replace(TemplateText, vbCrLf, vbLf), vbLf)
    
    Dim LB
    Dim UB
    LB = LBound(TemplateTextArray)
    UB = UBound(TemplateTextArray)
    
    ' ResultArray
    
    Dim ResultArray()
    ReDim ResultArray(UB)
    
    Dim Index
    For Index = LB To UB
        ResultArray(Index) = GetMatchingText(Trim(TemplateTextArray(Index)))
    Next
    
    ' MatchingLines
    
    Dim MatchingLines
    MatchingLines = Join(ResultArray, vbNewLine) & vbNewLine
    
    GetMatchingLines = MatchingLines
End Function

Public Function GetMatchingText(TemplateText)
    If TemplateText = "" Then Exit Function
    
    ' B.15.3 Legacy Text Template format
    'Const Pattern = "(?:<<var;name=""([^""]+)"";original=""(.+)"";match=""(.+)"">>|<<beginOptional>>(.+)<<endOptional>>)"
    Const Pattern = "(?:<<var;name=""([^""]+)"";original=""|<<beginOptional>>)"
    
    If Not RegExp_Test(TemplateText, Pattern, True, False) Then
        GetMatchingText = GetSimpleMatchingPattern(TemplateText)
        Exit Function
    End If
    
    Dim Matches
    Set Matches = RegExp_Execute(TemplateText, Pattern, True, False, False)
    
    Dim Match
    Set Match = Matches.Item(0)
    
    Dim PreviousPattern
    If Match.FirstIndex > 0 Then
        Dim PreviousText
        PreviousText = Left(TemplateText, Match.FirstIndex)
        PreviousPattern = GetSimpleMatchingPattern(PreviousText)
    End If
    
    Dim MiddlePattern
    Dim MiddleTextTemp
    Dim PostPattern
    Dim PostTextTemp
    
    If Match.SubMatches.Item(0) <> "" Then
        ' B.3.4 Guideline: replaceable text
        ' B.8 Bullets and numbering
        ' B.11 Copyright notice
        ' <<var;name="([^"]+)";original="(.+)";match="(.+)">>
        
        'Dim VarName
        'Dim VarOriginal
        'Dim VarMatch
        'VarName = Match.SubMatches.Item(0)
        'VarOriginal = Match.SubMatches.Item(1)
        'VarMatch = Match.SubMatches.Item(2)
        
        'MiddlePattern = VarMatch
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostPattern = GetMatchingText(PostTextTemp)
        'End If
        
        Dim VarMatchPos
        Dim VarEndPos
        
        VarMatchPos = _
            InStr( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText, _
                """;match=""")
        
        If VarMatchPos > 0 Then
            VarEndPos = _
                InStr( _
                    VarMatchPos + Len(""";match="""), _
                    TemplateText, _
                    """>>")
        End If
        
        If VarEndPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    VarMatchPos + Len(""";match="""), _
                    VarEndPos - (VarMatchPos + Len(""";match=""")))
            MiddlePattern = MiddleTextTemp
            
            If VarEndPos - 1 + Len(""">>") < Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - (VarEndPos - 1 + Len(""">>")))
                PostPattern = GetMatchingText(PostTextTemp)
            End If
            
        End If
        
    Else
        ' B.3.5 Guideline: omittable text
        ' B.12 License name or title
        ' B.13 Extraneous text at the end of a license
        ' "<<beginOptional>>(.+)<<endOptional>>"
        
        'Dim OptText
        'OptText = Match.SubMatches.Item(3)
        
        'Dim OptPattern
        'OptPattern = GetSimpleMatchingPattern(OptText)
        
        'MiddlePattern = "(?:" & OptPattern & ")?"
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostPattern = GetMatchingText(PostTextTemp)
        'End If
        
        Dim EndOptionalPos
        'EndOptionalPos = _
        '    InStr( _
        '        Match.FirstIndex + Match.Length + 1, _
        '        TemplateText, _
        '        "<<endOptional>>")
        EndOptionalPos = _
            GetEndOptionalPos( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText)
        If EndOptionalPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    EndOptionalPos - 1 - (Match.FirstIndex + Match.Length))
            MiddlePattern = "(?:" & GetMatchingText(MiddleTextTemp) & ")?"
            
            If EndOptionalPos - 1 + Len("<<endOptional>>") < _
                Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - _
                            (EndOptionalPos - 1 + Len("<<endOptional>>")))
                PostPattern = GetMatchingText(PostTextTemp)
            End If
            
        End If
        
    End If
    
    GetMatchingText = PreviousPattern & MiddlePattern & PostPattern
End Function

Public Function GetPlainText(TemplateText)
    If TemplateText = "" Then Exit Function
    
    Dim PlainText
    PlainText = TemplateText
    
    ' For Custom Template
    PlainText = _
        RegExp_Replace( _
            PlainText, _
            "", _
            "<<var;name=""([^""]+)"";original=""([^""]+)"";" & _
            "match=""(\(\?<![^""]+)"">>", _
            False, True, False)
    PlainText = _
        RegExp_Replace( _
            PlainText, _
            "", _
            "<<var;name=""([^""]+)"";original=""([^""]+)"";" & _
            "match=""(\(\?![^""]+)"">>", _
            False, True, False)
    
    PlainText = Replace(PlainText, "<<beginOptional>>", "")
    PlainText = Replace(PlainText, "<<endOptional>>", "")
    PlainText = _
        RegExp_Replace( _
            PlainText, _
            "", "<<var;name=""([^""]+)"";original=""", _
            False, True, False)
    PlainText = _
        RegExp_Replace( _
            PlainText, _
            "", """;match=""([^""]+)"">>", _
            False, True, False)
    
    GetPlainText = PlainText
End Function

Public Function GetPlainTextEx(TemplateText)
    If TemplateText = "" Then Exit Function
    
    ' For Custom Template
    TemplateText = _
        RegExp_Replace( _
            TemplateText, _
            "", _
            "<<var;name=""([^""]+)"";original=""([^""]+)"";" & _
            "match=""(\(\?<![^""]+)"">>", _
            False, True, False)
    TemplateText = _
        RegExp_Replace( _
            TemplateText, _
            "", _
            "<<var;name=""([^""]+)"";original=""([^""]+)"";" & _
            "match=""(\(\?![^""]+)"">>", _
            False, True, False)
    
    ' B.15.3 Legacy Text Template format
    'Const Pattern = "(?:<<var;name=""([^""]+)"";original=""(.+)"";match=""(.+)"">>|<<beginOptional>>(.+)<<endOptional>>" & ")"
    Const Pattern = "(?:<<var;name=""([^""]+)"";original=""|<<beginOptional>>)"
    
    If Not RegExp_Test(TemplateText, Pattern, True, False) Then
        GetPlainTextEx = TemplateText
        Exit Function
    End If
    
    Dim Matches
    Set Matches = RegExp_Execute(TemplateText, Pattern, True, False, False)
    
    Dim Match
    Set Match = Matches.Item(0)
    
    Dim PreviousText
    If Match.FirstIndex > 0 Then
        PreviousText = Left(TemplateText, Match.FirstIndex)
    End If
    
    Dim MiddleText
    Dim MiddleTextTemp
    Dim PostText
    Dim PostTextTemp
    
    If Match.SubMatches.Item(0) <> "" Then
        ' B.3.4 Guideline: replaceable text
        ' B.8 Bullets and numbering
        ' B.11 Copyright notice
        ' <<var;name="([^"]+)";original="(.+)";match="(.+)">>
        
        'Dim VarName
        'Dim VarOriginal
        'Dim VarMatch
        'VarName = Match.SubMatches.Item(0)
        'VarOriginal = Match.SubMatches.Item(1)
        'VarMatch = Match.SubMatches.Item(2)
        
        'MiddleText = VarOriginal
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetPlainTextEx(PostTextTemp)
        'End If
        
        Dim VarMatchPos
        Dim VarEndPos
        
        VarMatchPos = _
            InStr( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText, _
                """;match=""")
        
        If VarMatchPos > 0 Then
            VarEndPos = _
                InStr( _
                    VarMatchPos + Len(""";match="""), _
                    TemplateText, _
                    """>>")
        End If
        
        If VarEndPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    VarMatchPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = MiddleTextTemp
            
            If VarEndPos - 1 + Len(""">>") < Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - (VarEndPos - 1 + Len(""">>")))
                PostText = GetPlainTextEx(PostTextTemp)
            End If
            
        End If
        
    Else
        ' B.3.5 Guideline: omittable text
        ' B.12 License name or title
        ' B.13 Extraneous text at the end of a license
        ' "<<beginOptional>>(.+)<<endOptional>>"
        
        'MiddleTextTemp = Match.SubMatches.Item(3)
        'MiddleText = GetPlainTextEx(MiddleTextTemp)
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetPlainTextEx(PostTextTemp)
        'End If
        
        Dim EndOptionalPos
        'EndOptionalPos = _
        '    InStr( _
        '        Match.FirstIndex + Match.Length + 1, _
        '        TemplateText, _
        '        "<<endOptional>>")
        EndOptionalPos = _
            GetEndOptionalPos( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText)
        If EndOptionalPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    EndOptionalPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = GetPlainTextEx(MiddleTextTemp)
            
            If EndOptionalPos - 1 + Len("<<endOptional>>") < _
                Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - _
                            (EndOptionalPos - 1 + Len("<<endOptional>>")))
                PostText = GetPlainTextEx(PostTextTemp)
            End If
            
        End If
        
    End If
    
    GetPlainTextEx = PreviousText & MiddleText & PostText
End Function

Private Function GetEndOptionalPos(StartPos, TemplateText)
    Dim BeginOptionalPos
    BeginOptionalPos = InStr(StartPos, TemplateText, "<<beginOptional>>")
    
    Dim EndOptionalPos
    EndOptionalPos = InStr(StartPos, TemplateText, "<<endOptional>>")
    
    If BeginOptionalPos = 0 Then
        GetEndOptionalPos = EndOptionalPos
        Exit Function
    End If
    
    If BeginOptionalPos > EndOptionalPos Then
        GetEndOptionalPos = EndOptionalPos
        Exit Function
    End If
    
    Dim EndOptionalPosTemp
    EndOptionalPosTemp = _
        GetEndOptionalPos( _
            BeginOptionalPos + Len("<<beginOptional>>"), _
            TemplateText)
    
    EndOptionalPos = _
        GetEndOptionalPos( _
            EndOptionalPosTemp + Len("<<endOptional>>"), _
            TemplateText)
    
    GetEndOptionalPos = EndOptionalPos
End Function

Public Function GetFontText(TemplateText)
    If TemplateText = "" Then Exit Function
    
    ' For Custom Template
    TemplateText = _
        RegExp_Replace( _
            TemplateText, _
            "", _
            "<<var;name=""([^""]+)"";original=""([^""]+)"";" & _
            "match=""(\(\?<![^""]+)"">>", _
            False, True, False)
    TemplateText = _
        RegExp_Replace( _
            TemplateText, _
            "", _
            "<<var;name=""([^""]+)"";original=""([^""]+)"";" & _
            "match=""(\(\?![^""]+)"">>", _
            False, True, False)
    
    ' B.15.3 Legacy Text Template format
    'Const Pattern = "(?:<<var;name=""([^""]+)"";original=""(.+)"";match=""(.+)"">>|<<beginOptional>>(.+)<<endOptional>>)"
    Const Pattern = "(?:<<var;name=""([^""]+)"";original=""|<<beginOptional>>)"
    
    If Not RegExp_Test(TemplateText, Pattern, True, False) Then
        GetFontText = Space(Len(TemplateText))
        Exit Function
    End If
    
    Dim Matches
    Set Matches = RegExp_Execute(TemplateText, Pattern, True, False, False)
    
    Dim Match
    Set Match = Matches.Item(0)
    
    Dim PreviousText
    If Match.FirstIndex > 0 Then
        PreviousText = Space(Match.FirstIndex)
    End If
    
    Dim MiddleText
    Dim MiddleTextTemp
    Dim PostText
    Dim PostTextTemp
    
    If Match.SubMatches.Item(0) <> "" Then
        ' B.3.4 Guideline: replaceable text
        ' B.8 Bullets and numbering
        ' B.11 Copyright notice
        ' <<var;name="([^"]+)";original="(.+)";match="(.+)">>
        
        'Dim VarName
        'Dim VarOriginal
        'Dim VarMatch
        'VarName = Match.SubMatches.Item(0)
        'VarOriginal = Match.SubMatches.Item(1)
        'VarMatch = Match.SubMatches.Item(2)
        
        'MiddleText = StrRepeat("r", Len(VarOriginal))
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetFontText(PostTextTemp)
        'End If
        
        Dim VarMatchPos
        Dim VarEndPos
        
        VarMatchPos = _
            InStr( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText, _
                """;match=""")
        
        If VarMatchPos > 0 Then
            VarEndPos = _
                InStr( _
                    VarMatchPos + Len(""";match="""), _
                    TemplateText, _
                    """>>")
        End If
        
        If VarEndPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    VarMatchPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = StrRepeat("r", Len(MiddleTextTemp))
            
            If VarEndPos - 1 + Len(""">>") < Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - (VarEndPos - 1 + Len(""">>")))
                PostText = GetFontText(PostTextTemp)
            End If
            
        End If
        
    Else
        ' B.3.5 Guideline: omittable text
        ' B.12 License name or title
        ' B.13 Extraneous text at the end of a license
        ' "<<beginOptional>>(.+)<<endOptional>>"
        
        'MiddleTextTemp = Match.SubMatches.Item(3)
        'MiddleText = StrRepeat("b", Len(GetFontText(MiddleTextTemp)))
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetFontText(PostTextTemp)
        'End If
        
        Dim EndOptionalPos
        'EndOptionalPos = _
        '    InStr( _
        '        Match.FirstIndex + Match.Length + 1, _
        '        TemplateText, _
        '        "<<endOptional>>")
        EndOptionalPos = _
            GetEndOptionalPos( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText)
        If EndOptionalPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    EndOptionalPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = StrRepeat("b", Len(GetFontText(MiddleTextTemp)))
            
            If EndOptionalPos - 1 + Len("<<endOptional>>") < _
                Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - _
                            (EndOptionalPos - 1 + Len("<<endOptional>>")))
                PostText = GetFontText(PostTextTemp)
            End If
            
        End If
        
    End If
    
    GetFontText = PreviousText & MiddleText & PostText
End Function

Private Function StrRepeat(Text, Length)
    If Text = "" Then Exit Function
    If Length <= 0 Then Exit Function
    
    Dim Result
    
    Dim Index
    For Index = 0 To Length - 1
        Result = Result & Text
    Next
    
    StrRepeat = Result
End Function
