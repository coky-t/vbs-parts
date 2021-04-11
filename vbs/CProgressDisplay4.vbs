'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "CProgressDisplay4"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit

'
' Copyright (c) 2021 Koki Takeyama
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

Class CProgressDisplay4

Public Counter
Public CounterEnd

Public IndicatorDoneSymbol
Public IndicatorDoingSymbols
Public IndicatorNotYetSymbol
Public IndicatorEnd

Public Ruler
Public bRulerAlready

Public Title
Public Comment

Public StartTime

Private Sub Class_Initialize()
    Reset
End Sub

Private Sub Class_Terminate()
End Sub

Private Sub Reset()
    Counter = 0
    CounterEnd = 100
    
    IndicatorDoneSymbol = "*"
    IndicatorDoingSymbols = Array("|", "/", "-", "\")
    IndicatorNotYetSymbol = " "
    
    IndicatorEnd = 51
    Ruler = _
        "0%   10   20   30   40   50   60   70   80   90   100%" & vbCrLf & _
        "|----|----|----|----|----|----|----|----|----|----|"
    
    'IndicatorEnd = 52
    'Ruler = "|0----------25-----------50----------75---------100|"
    
    bRulerAlready = False
    
    Title = ""
    Comment = ""
    
    StartTime = Now
End Sub

Public Sub Increment()
    Counter = Counter + 1
    Display GetIndicator, GetTimes
End Sub

Private Function GetIndicator()
    Dim Indicator
    
    Dim IndicatorDoneCount
    IndicatorDoneCount = (Counter * IndicatorEnd) \ CounterEnd
    
    Dim Index
    For Index = 1 To IndicatorDoneCount
        Indicator = Indicator & IndicatorDoneSymbol
    Next
    
    If IndicatorDoneCount < IndicatorEnd Then
        Dim IndicatorDoingIndex
        IndicatorDoingIndex = Counter Mod (UBound(IndicatorDoingSymbols) + 1)
        
        Indicator = Indicator & IndicatorDoingSymbols(IndicatorDoingIndex)
    End If
    
    For Index = IndicatorDoneCount + 2 To IndicatorEnd
        Indicator = Indicator & IndicatorNotYetSymbol
    Next
    
    GetIndicator = Indicator
End Function

Public Function GetTimes()
    Dim EndTime
    EndTime = Now
    
    Dim ElapsedSeconds
    ElapsedSeconds = DateDiff("s", StartTime, EndTime)
    
    If ElapsedSeconds = 0 Then Exit Function
    
    Dim Times
    'Times = " Elapsed time: " & FormatTime(ElapsedSeconds)
    Times = "                    "
    
    If Counter < CounterEnd Then
        Dim RemainingSeconds
        RemainingSeconds = _
            ElapsedSeconds * (CounterEnd - Counter) \ Counter
        
        'Times = Times & _
        '    " Estimated time remaining: " & FormatTime(RemainingSeconds)
        Times = " Remaining: " & FormatTime(RemainingSeconds)
    End If
    
    GetTimes = Times
End Function

Public Sub Display(Indicator, Times)
    If Not bRulerAlready Then
        WScript.StdOut.WriteLine Ruler
        bRulerAlready = True
    End If
    Dim Control
    Control = Chr(13)
    If Counter >= CounterEnd Then Control = Control & Chr(10)
    WScript.StdOut.Write Title & Indicator & Comment & Times & Control
End Sub

Private Function FormatTime(Seconds)
    Dim SecondDiff
    SecondDiff = Seconds
    
    Dim MinuteDiff
    MinuteDiff = SecondDiff \ 60
    SecondDiff = SecondDiff Mod 60
    
    Dim HourDiff
    HourDiff = MinuteDiff \ 60
    MinuteDiff = MinuteDiff Mod 60
    
    Dim TimeDiffString
    TimeDiffString = _
        Right("0" & HourDiff, 2) & ":" & _
        Right("0" & MinuteDiff, 2) & ":" & _
        Right("0" & SecondDiff, 2)
    
    FormatTime = TimeDiffString
End Function

End Class
