'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "CProgressDisplay2"
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

Class CProgressDisplay2

Public Counter
Public CounterEnd

Public IndicatorDoneSymbol
Public IndicatorNotYetSymbol
Public IndicatorEnd

Public Ruler
Public bRulerAlready

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
    
    IndicatorDoneSymbol = "*"
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
End Sub

Public Sub Increment()
    Counter = Counter + 1
    Display GetIndicator
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

Public Sub Display(Indicator)
    If Not bRulerAlready Then
        WScript.StdOut.WriteLine Ruler
        bRulerAlready = True
    End If
    WScript.StdOut.Write Title & Indicator & Comment & Chr(13)
End Sub

End Class
