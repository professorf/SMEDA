Attribute VB_Name = "mNeuralNet"
'
' ProfessorF (pf) Neural Network Excel Library
'
' Copyright (c) Nick V. Flor, 2014-2017, All rights reserved
'
' This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
' CC BY-SA, If you use this code for research, you must cite me in your paper references
' To view a copy of the license visit: http://creativecommons.org/licenses/by-sa/4.0/legalcode
' To view a summary of the license visit: http://creativecommons.org/licenses/by-sa/4.0/
'
' This material is based partly upon work supported by the National Science Foundation (NSF)
' under ECCS - 1231046 and - SNM 1635334 . Any opinions, findings, and conclusions or recommendations
' expressed in this material are those of the author and do not necessarily reflect the views of the NSF.
'
' Based on the algorithms in the Explorations in PDP Book (McClelland & Rumelhart, 1988)
'
Option Explicit

Dim nunits As Long
Dim netinput() As Double
Dim activation() As Double
Dim err(), delta() As Double
Dim weight(), wed(), dweight() As Double
Dim first_weight_to() As Double
Dim last_weight_to() As Double
Dim ninputs, nhidden, noutputs As Long
Dim istartinput, iendinput, istarthidden, iendhidden, istartoutput, iendoutput, lastunit As Long
Dim istarttrain, iendtrain As Long
Dim wstrain As Worksheet
Dim trainrow As Long
Dim lrate, momentum As Double

Sub nnInit()
Attribute nnInit.VB_ProcData.VB_Invoke_Func = "i\n14"
Dim i, j As Long
'
' Determine # of units
' Add 2 for input layer & hidden layer biases
'
ninputs = CLng(Range("ninputs"))
nhidden = CLng(Range("nhidden"))
noutputs = CLng(Range("noutputs"))
nunits = ninputs + nhidden + noutputs + 2
'
' To facilitate porting to other languages, we'll use 0 as a base
'
ReDim activation(0 To nunits - 1)
ReDim delta(0 To nunits - 1)
ReDim netinput(0 To nunits - 1)
ReDim weight(0 To nunits - 1, 0 To nunits - 1)
ReDim dweight(0 To nunits - 1, 0 To nunits - 1)
ReDim wed(0 To nunits - 1, 0 To nunits - 1)
ReDim first_weight_to(0 To nunits - 1)
ReDim last_weight_to(0 To nunits - 1)
ReDim err(0 To nunits - 1)
'
' Initialize all unit connections
'
istartinput = 0
iendinput = ninputs
istarthidden = iendinput + 1
iendhidden = istarthidden + nhidden
istartoutput = iendhidden + 1
iendoutput = istartoutput + noutputs - 1
lastunit = nunits - 1

For i = istartinput To iendinput ' inputs have no connections
    first_weight_to(i) = -1
    last_weight_to(i) = -1
Next

For i = istarthidden To iendhidden
    first_weight_to(i) = istartinput
    last_weight_to(i) = iendinput
Next

For i = istartoutput To iendoutput
        first_weight_to(i) = istarthidden
        last_weight_to(i) = iendhidden
Next
'
' Initialize weights
'
For i = istartoutput To iendoutput
    For j = istarthidden To iendhidden
        weight(i, j) = Math.Rnd
    Next
Next

For i = istarthidden To iendhidden
    For j = istartinput To iendinput
        weight(i, j) = Math.Rnd
    Next
Next
'
' Initalize weight delta matrix
'
For i = istartinput To iendoutput
    For j = istartinput To iendoutput
        wed(i, j) = 0
    Next
Next
'
' Initalize learning weight & momentum
'
lrate = CDbl(Range("lrate"))
momentum = CDbl(Range("momentum"))
'
' initialize training indices
'
Set wstrain = Sheets("training")
istarttrain = 1
iendtrain = wstrain.Cells(Rows.Count, 1).End(xlUp).Row
End Sub
Function logistic(net As Double) As Double
    logistic = 1 / (1 + Math.Exp(-net))
End Function
Sub change_weights()
Dim i, j As Long
' sum_linked_weds?
For i = istarthidden To iendoutput
    For j = first_weight_to(i) To last_weight_to(i)
        dweight(i, j) = lrate * wed(i, j) + momentum * dweight(i, j)
        weight(i, j) = weight(i, j) + dweight(i, j)
        wed(i, j) = 0
    Next
Next

End Sub
Sub compute_wed()
Dim i, j As Long
' WARNING: where do weds get zeroed out?
For i = istarthidden To iendoutput
    For j = first_weight_to(i) To last_weight_to(i)
       wed(i, j) = wed(i, j) + delta(i) * activation(j)
    Next
Next
End Sub
Sub compute_error()
Dim tr As Double
Dim i, j, traincol As Long
'
' Clear out all error for hidden and output units
'
For i = istarthidden To iendoutput
    err(i) = 0
Next

'
' Compute errors for output units
'
For i = istartoutput To iendoutput
    traincol = ninputs + (i - istartoutput + 1) ' + noutputs
    tr = wstrain.Cells(trainrow, traincol)
    err(i) = tr - activation(i)
Next
'
' compute errors for all hidden to output
'
For i = iendoutput To istarthidden Step -1
    '
    ' Remember, delta is different for back propagation
    '
    delta(i) = err(i) * activation(i) * (1 - activation(i))
    '
    ' propagate down
    '
    For j = first_weight_to(i) To last_weight_to(i)
        err(j) = err(j) + delta(i) * weight(i, j)
    Next
Next

End Sub
Sub compute_output()
Dim i, j As Long
For i = istarthidden To iendoutput
    netinput(i) = 0
    For j = first_weight_to(i) To last_weight_to(i)
        netinput(i) = netinput(i) + weight(i, j) * activation(j) ' weight to i from j
    Next
    activation(i) = logistic(netinput(i))
Next
End Sub
Sub nnLoadWeights()
Dim i, j As Long
Dim ws As Worksheet

Set ws = Worksheets("weights")
For i = istartinput To iendoutput
    For j = istartinput To iendoutput
        weight(i, j) = ws.Cells(i + 1, j + 1)
    Next
Next
End Sub

Sub nnDumpWeights()
Attribute nnDumpWeights.VB_ProcData.VB_Invoke_Func = "w\n14"
Dim i, j As Long
Dim ws As Worksheet

Set ws = Worksheets("weights")
For i = istartinput To iendoutput
    For j = istartinput To iendoutput
        ws.Cells(i + 1, j + 1) = weight(i, j)
    Next
Next
End Sub
Sub nnDumpActivations()
Attribute nnDumpActivations.VB_ProcData.VB_Invoke_Func = "a\n14"
Dim i, j As Long
Dim ws As Worksheet

Set ws = Worksheets("activations")

For i = istartoutput To iendoutput
    ws.Cells(1, (i - istartoutput + 1)) = activation(i)
Next
For i = istarthidden To iendhidden
    ws.Cells(2, (i - istarthidden + 1)) = activation(i)
Next
For i = istartinput To iendinput
    ws.Cells(3, (i - istartinput + 1)) = activation(i)
Next

End Sub
Sub nnDumpNets()
Attribute nnDumpNets.VB_ProcData.VB_Invoke_Func = "n\n14"
Dim i, j As Long
Dim ws As Worksheet

Set ws = Worksheets("nets")

For i = istartoutput To iendoutput
    ws.Cells(1, (i - istartoutput + 1)) = netinput(i)
Next
For i = istarthidden To iendhidden
    ws.Cells(2, (i - istarthidden + 1)) = netinput(i)
Next
For i = istartinput To iendinput
    ws.Cells(3, (i - istartinput + 1)) = netinput(i)
Next

End Sub
Sub nnLoad(r As Long)
Dim i As Long
For i = istartinput To iendinput 'istartinput must ALWAYS be 0
    If (i = istartinput) Then
        activation(i) = 1 ' bias is always 1
    Else
        activation(i) = wstrain.Cells(r, i)
    End If
Next
End Sub
Sub show_outputs()
Dim i, tcol As Long

For i = istartoutput To iendoutput
    tcol = ninputs + noutputs + 1
    wstrain.Cells(trainrow, (tcol + i - istartoutput)) = activation(i)
Next
End Sub
Sub nnTrain()
Attribute nnTrain.VB_ProcData.VB_Invoke_Func = "r\n14"
Dim i, epochs As Long
'For r = istarttrain To iendtrain
epochs = CLng(Range("epoch"))
For i = 1 To epochs
    lrate = CDbl(Range("lrate"))
    momentum = CDbl(Range("momentum"))
    For trainrow = istarttrain To iendtrain
        nnLoad (trainrow)
        compute_output
        compute_error ' training
        compute_wed ' weight deltas
        change_weights ' actually changes the weights
        show_outputs
        DoEvents
    Next
    Application.StatusBar = CStr(i) + "/" + CStr(epochs)
Next
End Sub

Sub nnRun()
Dim i, epochs As Long
    For trainrow = istarttrain To iendtrain
        nnLoad (trainrow)
        compute_output
        show_outputs
        DoEvents
    Next
End Sub

