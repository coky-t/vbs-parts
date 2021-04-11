'Attribute VB_Name = "Test_CProgressDisplay3"
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

'
' --- Test ---
'

Public Sub Test_CProgressDisplay1()
    Dim PDisp
    Set PDisp = New CProgressDisplay3
    PDisp.CounterEnd = 1000
    
    Dim Index
    For Index = 1 To 1000
        PDisp.Increment
        WScript.Sleep 1
    Next
End Sub

Public Sub Test_CProgressDisplay2()
    Dim PDisp
    Set PDisp = New CProgressDisplay3
    PDisp.CounterEnd = 1000
    PDisp.IndicatorDoingSymbols = Array("-", "+", "*")
    
    Dim Index
    For Index = 1 To 1000
        PDisp.Increment
        WScript.Sleep 1
    Next
End Sub

Public Sub Test_CProgressDisplay3()
    Dim PDisp
    Set PDisp = New CProgressDisplay3
    PDisp.CounterEnd = 200
    PDisp.IndicatorDoneSymbol = "+"
    PDisp.IndicatorDoingSymbols = Array("-", "+", "*")
    PDisp.IndicatorNotYetSymbol = "-"
    
    Dim Index
    For Index = 1 To 200
        PDisp.Increment
        WScript.Sleep 100
    Next
End Sub

Public Sub Test_CProgressDisplay4()
    Dim PDisp
    Set PDisp = New CProgressDisplay3
    PDisp.CounterEnd = 100
    PDisp.IndicatorDoneSymbol = "+"
    PDisp.IndicatorNotYetSymbol = "-"
    PDisp.IndicatorEnd = 52
    PDisp.Ruler = "|0----------25-----------50----------75---------100|"
    
    Dim Index
    For Index = 1 To 100
        PDisp.Increment
        WScript.Sleep 100
    Next
End Sub

Public Sub Test_CProgressDisplay5()
    Dim PDisp
    Set PDisp = New CProgressDisplay3
    PDisp.CounterEnd = 300
    PDisp.IndicatorDoneSymbol = "#"
    PDisp.IndicatorDoingSymbols = Array("-", "=", "#")
    PDisp.IndicatorNotYetSymbol = "="
    PDisp.IndicatorEnd = 52
    PDisp.Ruler = "      |0----------25-----------50----------75---------100|"
    PDisp.Title = "Test: "
    
    Dim Index
    For Index = 1 To 300
        PDisp.Increment
        WScript.Sleep 10
    Next
End Sub

Public Sub Test_CProgressDisplay6()
    Dim PDisp
    Set PDisp = New CProgressDisplay3
    PDisp.CounterEnd = 300
    PDisp.IndicatorDoneSymbol = "#"
    PDisp.IndicatorDoingSymbols = Array("-", "=", "#")
    PDisp.IndicatorNotYetSymbol = "="
    PDisp.IndicatorEnd = 52
    PDisp.Ruler = "      |0----------25-----------50----------75---------100|"
    PDisp.Title = "Test: "
    
    Dim Index
    For Index = 1 To 300
        PDisp.Comment = Space(1) & CStr(Index) & "/300"
        PDisp.Increment
        WScript.Sleep 100
    Next
End Sub
