'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "CProgressDisplay"
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

Class CProgressDisplay

Public Counter
Public CounterEnd

Public IndicatorDoneSymbol
Public IndicatorNotYetSymbol
Public IndicatorEnd

Public Title
Public Comment

Private Sub Class_Initialize()
    Reset
End Sub

Private Sub Class_Terminate()
End Sub

Private Sub Reset()
    Counter = 0
    CounterEnd = 100
    
    IndicatorDoneSymbol = "###"
    IndicatorNotYetSymbol = "___"
    IndicatorEnd = 10
    
    Title = ""
    Comment = ""
End Sub

Public Sub Increment()
    Counter = Counter + 1
    Display GetIndicator, GetPercent
End Sub

Private Function GetIndicator()
    Dim Indicator
    
    Dim IndicatorDoneCount
    IndicatorDoneCount = (Counter * IndicatorEnd) \ CounterEnd
    
    Dim Index
    For Index = 1 To IndicatorDoneCount
        Indicator = Indicator & IndicatorDoneSymbol
    Next
    For Index = IndicatorDoneCount + 1 To IndicatorEnd
        Indicator = Indicator & IndicatorNotYetSymbol
    Next
    
    GetIndicator = Indicator
End Function

Private Function GetPercent()
    GetPercent = Space(1) & CStr((Counter * 100) \ CounterEnd) & "%"
End Function

Public Sub Display(Indicator, Percent)
    Dim Control
    Control = Chr(13)
    If Counter >= CounterEnd Then Control = Control & Chr(10)
    WScript.StdOut.Write Title & Indicator & Percent & Comment & Control
End Sub

End Class
