Attribute VB_Name = "Module11"
'
' ProfessorF (pf) Vector, Matrix, and Principal Component Analysis Library
'
' Copyright (c) Nick V. Flor, 2014-2017, All rights reserved
'
' This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
' CC BY-SA
' To view a copy of the license visit: http://creativecommons.org/licenses/by-sa/4.0/legalcode
' To view a summary of the license visit: http://creativecommons.org/licenses/by-sa/4.0/
'
' This material is based partly upon work supported by the National Science Foundation (NSF)
' under ECCS - 1231046 and SNM - 1635334. Any opinions, findings, and conclusions or recommendations
' expressed in this material are those of the author and do not necessarily reflect the views of the NSF.
'
' Version: 14May16
'

'
' +++++++++++++++++++++++++++++++++++++++
' ++++++++++ UTILITY FUNCTIONS ++++++++++
' +++++++++++++++++++++++++++++++++++++++
'

'
' reindexRange: Converts an Excel Range to a Matrix with 0 as the Starting Index
'               This makes it easier to reuse later functions with other programming languages
'
Function reindexRange(src As Range) ' returns a matrix with starting index at 0
ReDim dst(0 To src.rows.Count - 1, 0 To src.Columns.Count - 1)

For R = 1 To src.rows.Count
    For C = 1 To src.Columns.Count
        'MsgBox ("M(" & r & "," & c & ")=" & src(r, c).Value)
        dst(R - 1, C - 1) = src(R, C)
    Next
Next
reindexRange = dst
End Function

'
' getNumRowsCols: Returns the number of rows and columns in a matrix
'
Sub getNumRowsCols(ByRef A, ByRef M, ByRef n)
    M = UBound(A, 1) - LBound(A, 1) + 1
    n = UBound(A, 2) - LBound(A, 2) + 1
End Sub

'
' ++++++++++++++++++++++++++++++++++++++
' ++++++++++ MATRIX FUNCTIONS ++++++++++
' ++++++++++++++++++++++++++++++++++++++
'

'
' Matrix Copy
'
Function fmCopy(src)
Dim mRows, nCols, R, C

getNumRowsCols src, mRows, nCols
'mRows = UBound(src, 1) - LBound(src, 1) + 1
'nCols = UBound(src, 2) - LBound(src, 2) + 1

ReDim dst(0 To mRows - 1, 0 To nCols - 1)
For R = 0 To mRows - 1
    For C = 0 To nCols - 1
        dst(R, C) = src(R, C)
    Next
Next
fmCopy = dst
End Function

'
' Matrix Transpose
'
Function fmTrans(src)
Dim mRows, nCols, R, C

getNumRowsCols src, mRows, nCols

'mRows = UBound(src, 1) - LBound(src, 1) + 1
'nCols = UBound(src, 2) - LBound(src, 2) + 1

ReDim dst(0 To nCols - 1, 0 To mRows - 1)

For R = 0 To mRows - 1
    For C = 0 To nCols - 1
        'MsgBox ("M(" & r & "," & c & ")=" & src(r, c).Value)
        dst(C, R) = src(R, C)
    Next
Next
fmTrans = dst
End Function

'
' Matrix Multiply
'
Function fmMult(m1, m2)
Dim m1Rows, m1Cols, m2Rows, m2Cols, mRows, nCols, R, C, k

getNumRowsCols m1, m1Rows, m1Cols
getNumRowsCols m2, m2Rows, m2Cols

'm1Rows = UBound(m1, 1) - LBound(m1, 1) + 1
'm1Cols = UBound(m1, 2) - LBound(m1, 2) + 1
'm2Rows = UBound(m2, 1) - LBound(m2, 1) + 1
'm2Cols = UBound(m2, 2) - LBound(m2, 2) + 1

If m1Cols <> m2Rows Then Return

mRows = m1Rows
nCols = m2Cols

ReDim dst(0 To mRows - 1, 0 To nCols - 1)

For R = 0 To mRows - 1 ' Each Row in m1
    For C = 0 To nCols - 1 ' multiplied by each col in m2
        total = 0
        For k = 0 To m1Cols - 1 'remember m1Cols=m2Rows
            total = total + m1(R, k) * m2(k, C)
        Next
        dst(R, C) = total
    Next
Next
fmMult = dst
End Function

'
' Matrix Extract Minor
'
Function fmMinor(src, diag)
Dim mRows, nCols, D, R, C

getNumRowsCols src, mRows, nCols

'mRows = UBound(src, 1) - LBound(src, 1) + 1
'nCols = UBound(src, 2) - LBound(src, 2) + 1

ReDim dst(0 To mRows - 1, 0 To nCols - 1)

' ZERO OUT MATRIX IF NOT AUTOMATIC
For i = 0 To mRows - 1
    For j = 0 To nCols - 1
        dst(i, j) = 0
    Next
Next


D = 0
While (D < diag)
    dst(D, D) = 1
    D = D + 1
Wend

For R = diag To (mRows - 1)
    For C = diag To (nCols - 1)
        dst(R, C) = src(R, C)
    Next
Next
fmMinor = dst
End Function

'
' Matrix Extract Column Vector
'
Function fmExtract(k, M)
Dim mRows
mRows = UBound(M, 1) - LBound(M, 1) + 1
ReDim v(0 To mRows - 1, 0)

For i = 0 To mRows - 1
    v(i, 0) = M(i, k)
Next
fmExtract = v
End Function

'
' Matrix Create Identity
'
Function createIdentity(mRows, nCols) ' For any rectangular matrix
Dim row, col
ReDim M(mRows - 1, nCols - 1)

    For row = 0 To mRows - 1
        For col = 0 To nCols - 1
            If (row = col) Then
                M(row, col) = 1
            Else
                M(row, col) = 0
            End If
        Next
    Next

createIdentity = M
End Function
'
' CreateMatrix (is really a Matrix)
'
Function createMatrix(rows, cols)
ReDim M(rows - 1, cols - 1)
Dim i, j

    For i = 0 To (rows - 1)
        For j = 0 To (cols - 1)
            M(i, j) = 0
        Next
    Next
    createMatrix = M
End Function
'
' Matrix Create Correlation
'
Function createCorrelationMatrix(A)
Dim M, n, rows, cols, total, col1, col2

' Determine number of rows and cols, index starts at 0
getNumRowsCols A, M, n

' Calculate Means
ReDim mean(n - 1)

For cols = 0 To n - 1
    total = 0
    For rows = 0 To M - 1
        total = total + A(rows, cols)
    Next
    mean(cols) = total / M
Next

' Calculate SDevs SDev=SQRT(SUM(Xi-Mean)^2/N)
ReDim Sdev(n - 1)

For cols = 0 To n - 1
    total = 0
    For rows = 0 To M - 1
        total = total + (A(rows, cols) - mean(cols)) ^ 2
    Next
    Sdev(cols) = Math.Sqr(total / M)
Next

' Now calculate the covariance matrix & divide by stdevs to get correlation
ReDim ACov(n - 1, n - 1)
For col1 = 0 To n - 1
    For col2 = 0 To n - 1
        If (col1 = col2) Then
            ACov(col1, col2) = 1
        Else
            total = 0
            For rows = 0 To M - 1
                total = total + (A(rows, col1) - mean(col1)) * (A(rows, col2) - mean(col2))
            Next
            ACov(col1, col2) = total / M
            ' now do correlation
            ACov(col1, col2) = ACov(col1, col2) / (Sdev(col1) * Sdev(col2))
        End If
    Next
Next

createCorrelationMatrix = ACov

End Function

'
' Matrix Create Covariance
'
Function createCovarianceMatrix(A)
Dim M, n, rows, cols, total, col1, col2

' Determine number of rows and cols, index starts at 0
getNumRowsCols A, M, n

' Calculate Means
ReDim mean(n - 1)

For cols = 0 To n - 1
    total = 0
    For rows = 0 To M - 1
        total = total + A(rows, cols)
    Next
    mean(cols) = total / M
Next

' Calculate SDevs SDev=SQRT(SUM(Xi-Mean)^2/N)
ReDim Sdev(n - 1)

For cols = 0 To n - 1
    total = 0
    For rows = 0 To M - 1
        total = total + (A(rows, cols) - mean(cols)) ^ 2
    Next
    Sdev(cols) = Math.Sqr(total / M)
Next

' Now calculate the covariance matrix
ReDim ACov(n - 1, n - 1)
For col1 = 0 To n - 1
    For col2 = 0 To n - 1
        If (col1 = col2) Then
            ACov(col1, col2) = 1
        Else
            total = 0
            For rows = 0 To M - 1
                total = total + (A(rows, col1) - mean(col1)) * (A(rows, col2) - mean(col2))
            Next
            ACov(col1, col2) = total / M
        End If
    Next
Next

createCovarianceMatrix = ACov

End Function

Function fmInverse(A)
Dim mRows, nCols, n2Cols, R, C

getNumRowsCols A, mRows, nCols

n2Cols = nCols * 2

ReDim G(0 To mRows - 1, 0 To n2Cols - 1)

' Copy the A matrix into G
For R = 0 To mRows - 1
    For C = 0 To nCols - 1
        G(R, C) = A(R, C)
    Next
Next

' Augment G with the Identity
For R = 0 To mRows - 1
    For C = nCols To n2Cols - 1
        If (R = (C - nCols)) Then
            G(R, C) = 1
        Else
            G(R, C) = 0
        End If
    Next
Next

'
' Do Gaussian Elimination
'

' go through each row and eliminate the rows underneath it
Dim key, rscale, r2
For R = 0 To mRows - 1 ' note: one less
    '
    ' Calculate how to scale the key (diagonal) to 1
    '
    key = G(R, R)
    rscale = (1 / key)
    
    '
    ' Now put a 1 in the key position
    '
    For C = R To n2Cols - 1
        G(R, C) = G(R, C) * rscale
    Next
        
    For r2 = (R + 1) To mRows - 1
        '
        ' Similarly calculate how to scale the keys below
        '
        key = G(r2, R) ' note r2, not r
        If (key <> 0) Then
            rscale = (1 / key)
            '
            ' Now put a 1 in the key position
            '
            For C = R To (nCols * 2) - 1
                G(r2, C) = G(R, C) - (G(r2, C) * rscale) ' note r2, not r, and we are subtracting the primary row
            Next
        End If
    Next
Next


'
' Do Back Substitution
'

ReDim v(0 To mRows - 1)
Dim k

For C = nCols To n2Cols - 1
    For R = mRows - 1 To 0 Step -1
        v(R) = G(R, C)
        For k = (nCols - 1) To (R + 1) Step -1
            v(R) = v(R) - v(k) * G(R, k)
        Next
        G(R, C) = v(R)
    Next
Next

'
' Copy the Augmented Matrix to the Inverse Matrix
'
ReDim AI(0 To mRows - 1, 0 To nCols - 1)
' Copy the A matrix into G
For R = 0 To mRows - 1
    For C = nCols To (nCols * 2) - 1
        AI(R, C - nCols) = G(R, C) ' Correct
        ' AI(r, c - nCols) = G(r, c - nCols) ' Debug
    Next
Next

fmInverse = AI
End Function

'
' EXCEL INTERFACES TO MATRIX FUNCTIONS
'
Function pfmInverse(src As Range)
s = reindexRange(src)
pfmInverse = fmInverse(s)
End Function

Function pfmCopy(src As Range)
s = reindexRange(src)
pfmCopy = fmCopy(s)
End Function

Function pfmTrans(src As Range)
s = reindexRange(src)
pfmTrans = fmTrans(s)
End Function

Function pfmMult(m1 As Range, m2 As Range)
Dim mr1, mr2
    mr1 = reindexRange(m1)
    mr2 = reindexRange(m2)
    pfmMult = fmMult(mr1, mr2)
End Function

Function pfmMinor(src As Range, diag As Integer)
s = reindexRange(src)
pfmMinor = fmMinor(s, diag)
End Function

Function pfmExtract(k As Integer, M As Range)
Dim mr
nm = reindexRange(M)
pfmExtract = fmExtract(k, nm)
End Function

Function pfCovariance(R As Range)
Dim A
Dim OutM
Dim LabOutM
Dim rows, cols

' Row 0, Columns 1,..,N contain labels of table
A = reindexRange(R)
OutM = createCovarianceMatrix(A)

ReDim LabOutM(R.Columns.Count, R.Columns.Count)
For cols = 1 To R.Cells.Columns.Count
    LabOutM(0, cols) = R(0, cols)
    LabOutM(cols, 0) = R(0, cols)
Next
For rows = 0 To R.Columns.Count - 1
    For cols = 0 To R.Columns.Count - 1
        LabOutM(rows + 1, cols + 1) = OutM(rows, cols)
    Next
Next
LabOutM(0, 0) = ""

pfCovariance = LabOutM

End Function

Function pfCorrelation(R As Range)
Dim A
Dim OutM
Dim LabOutM
Dim rows, cols

' Row 0, Columns 1,..,N contain labels of table
A = reindexRange(R)
OutM = createCorrelationMatrix(A)

ReDim LabOutM(R.Columns.Count, R.Columns.Count)
For cols = 1 To R.Cells.Columns.Count
    LabOutM(0, cols) = R(0, cols)
    LabOutM(cols, 0) = R(0, cols)
Next
For rows = 0 To R.Columns.Count - 1
    For cols = 0 To R.Columns.Count - 1
        LabOutM(rows + 1, cols + 1) = OutM(rows, cols)
    Next
Next
LabOutM(0, 0) = ""
pfCorrelation = LabOutM

End Function

'
' ++++++++++++++++++++++++++++++++++++++
' ++++++++++ VECTOR FUNCTIONS ++++++++++
' ++++++++++++++++++++++++++++++++++++++
'

'
' Vector Scale
'
Function fvScale(v, s)
Dim mRows
mRows = UBound(v, 1) - LBound(v, 1) + 1
ReDim nv(mRows - 1, 0)

For i = 0 To mRows - 1
    nv(i, 0) = v(i, 0) * s
Next
fvScale = nv
End Function

'
' Vector Add
'
Function fvAdd(v1, v2)
Dim mRows
    mRows = UBound(v1, 1) - LBound(v1, 1) + 1 ' either vector will do
    ReDim nv(mRows - 1, 0) ' remember all vectors have 0 columns
    For i = 0 To mRows - 1
        nv(i, 0) = v1(i, 0) + v2(i, 0)
    Next
fvAdd = nv
End Function
'
' Vector Sub
'
Function fvSub(v1, v2)
Dim mRows
    mRows = UBound(v1, 1) - LBound(v1, 1) + 1 ' either vector will do
    ReDim nv(mRows - 1, 0) ' remember all vectors have 0 columns
    For i = 0 To mRows - 1
        nv(i, 0) = v1(i, 0) - v2(i, 0)
    Next
fvSub = nv
End Function


'
' Vector Normalize
'
Function fvNorm(v) ' Euclidean Norm = Distance from Zero
Dim n, distance, val

n = UBound(v, 1) - LBound(v, 1) + 1 ' size
ReDim nv(0 To n - 1, 0) ' vectors in this modules are 0-columned

distance = 0
For i = 0 To n - 1
    distance = distance + (v(i, 0) * v(i, 0))
Next
distance = Math.Sqr(distance)
fvNorm = distance
End Function

'
' EXCEL INTERFACES TO VECTOR FUNCTIONS
'
Function pfvScale(v As Range, s As Double)
Dim nv
nv = reindexRange(v)
pfvScale = fvScale(nv, s)
End Function

Function pfvAdd(v1 As Range, v2 As Range)
    Dim vr1, vr2
    vr1 = reindexRange(v1)
    vr2 = reindexRange(v2)
    pfvAdd = fvAdd(vr1, vr2)
End Function

Function pfvSub(v1 As Range, v2 As Range)
    Dim vr1, vr2
    vr1 = reindexRange(v1)
    vr2 = reindexRange(v2)
    pfvSub = fvSub(vr1, vr2)
End Function

Function pfvNorm(v As Range)
Dim nv
nv = reindexRange(v)
pfvNorm = fvNorm(nv)
End Function

Function fVar(v, k)
Dim i, mRows, nCols, mean, total
getNumRowsCols v, mRows, nCols

    mean = fAvg(v, k)
    total = 0
    For i = 0 To (mRows - 1)
        total = total + (v(i, k) - mean) ^ 2
    Next
    total = total / mRows
    fVar = total
End Function

Function fAvg(v, k)
Dim i, mRows, nCols, total
    
    getNumRowsCols v, mRows, nCols
    For i = 0 To (mRows - 1)
        total = total + v(i, k)
    Next
    fAvg = (total / mRows)
End Function
Function pfVar(v As Range, Optional ByVal k As Integer = -1)
Dim nv
nv = reindexRange(v)

If (k < 0) Then k = 0
pfVar = fVar(nv, k)
End Function
Function pfAvg(v As Range, Optional ByVal k As Integer = -1)
Dim nv
nv = reindexRange(v)
If (k < 0) Then k = 0
pfAvg = fAvg(nv, k)
End Function
Function fCron(Y)
Dim mRows, nCols, i, j
Dim X, v, sumVY, VX

    getNumRowsCols Y, mRows, nCols
    
    ' Calculate X vector=SUM Y-Rows
    X = createMatrix(mRows, 1)
    For i = 0 To (mRows - 1)
        For j = 0 To (nCols - 1)
            X(i, 0) = X(i, 0) + Y(i, j)
        Next
    Next
    VX = fVar(X, 0) ' Variance of X
    
    ' Calculate Variances of Ys
    VY = createMatrix(1, nCols)
    sumVY = 0
    For j = 0 To (nCols - 1)
        VY(0, j) = fVar(Y, j) ' for debugging
        sumVY = sumVY + VY(0, j) ' Sum of Y Variances
    Next
    
    k = nCols
    fCron = (k / (k - 1)) * (1 - (sumVY / VX))
End Function
Function pfCron(YR As Range)
Dim Y
    Y = reindexRange(YR)
    prCron = fCron(Y)
End Function
'
' +++++++++++++++++++++++++++++++++++++++++++++++++++
' ++++++++++ PRINCIPAL COMPONENT FUNCTIONS ++++++++++
' +++++++++++++++++++++++++++++++++++++++++++++++++++
'

'
' Formulas are based on Wikipedia's QR Algorithm Page: http://en.wikipedia.org/wiki/QR_decomposition, 11/1/2014
'

Function fvIVector(length, k)
ReDim v(length - 1, 0)

' Zero out the vector
For i = 0 To length - 1
    v(i, 0) = 0
Next

v(k, 0) = 1
fvIVector = v
End Function

Function QI2vvT(v) ' Q = I-2vvT
Dim mRows
mRows = UBound(v, 1) - LBound(v, 1) + 1
ReDim Q(mRows - 1, mRows - 1)

For i = 0 To mRows - 1
    For j = 0 To mRows - 1
        Q(i, j) = -2 * v(i, 0) * v(j, 0)
        If (i = j) Then
            Q(i, j) = Q(i, j) + 1
        End If
    Next
Next
QI2vvT = Q
End Function
Function Householder(ByRef A, ByRef Q, ByRef R)
Dim mRows, nCols, QA, QList, QT, t
Dim Am, alpha, X, u, v, E, alphae, uNorm, Qm

    mRows = UBound(A, 1) - LBound(A, 1) + 1 'm.Rows.Count
    nCols = UBound(A, 2) - LBound(A, 2) + 1 'm.Columns.Count
    't is the minimum(mRows-1,nCols)
    If ((mRows - 1) < nCols) Then t = mRows - 1 Else t = nCols
    ReDim QList(t - 1)
    QA = A
    For col = 0 To t - 1
        Am = fmMinor(QA, col)
        X = fmExtract(col, Am)
        alpha = fvNorm(X)
        If A(col, col) > 0 Then alpha = -alpha
        E = fvIVector(mRows, col)
        'u = x+alpha*e ' in Wikipedia it's u=x-alpha*e
        alphae = fvScale(E, alpha)
        u = fvAdd(X, alphae)
        'u = fvSub(x, alphae)
        uNorm = fvNorm(u)
        v = fvScale(u, 1 / uNorm)
        Qm = QI2vvT(v)
        QList(col) = Qm
        'Repeat on the column minor of QA
        QA = fmMult(Qm, A)
    Next
    ' Now calculate Q=Q1T*Q2T..QNT
    ' R=QN...*Q2*Q1*A
    Q = fmTrans(QList(0))
    R = fmMult(QList(0), A)
    For col = 1 To t - 1
        QT = fmTrans(QList(col))
        Q = fmMult(Q, QT)
        R = fmMult(QList(col), R)
    Next
End Function
'
' EXCEL INTERFACES TO PRINCIPAL COMPONENT FUNCTIONS
'
Function pfQI2vvT(v As Range)
Dim VR
VR = reindexRange(v)
pfQI2vvT = QI2vvT(VR)
End Function

Function pfHouseholder(AR As Range)
Dim A, mRows, nCols, Q, R, QR
A = reindexRange(AR)
mRows = UBound(A, 1) - LBound(A, 1) + 1
nCols = UBound(A, 2) - LBound(A, 2) + 1
ReDim Q(mRows - 1, nCols - 1)
ReDim R(mRows - 1, nCols - 1)
ReDim QR(mRows - 1, nCols * 2 - 1)
Householder A, Q, R
For i = 0 To mRows - 1
    For j = 0 To nCols - 1
        QR(i, j) = Q(i, j)
        QR(i, j + nCols) = R(i, j)
    Next
Next
pfHouseholder = QR
End Function

Function pfIterateHouseholder(AR As Range, loops As Integer)
Dim A, mRows, nCols, Q, R, QR, EV
A = reindexRange(AR)
mRows = UBound(A, 1) - LBound(A, 1) + 1
nCols = UBound(A, 2) - LBound(A, 2) + 1
ReDim Q(mRows - 1, nCols - 1)
ReDim R(mRows - 1, nCols - 1)
ReDim QR(mRows - 1, nCols * 3 - 1)

EV = createIdentity(mRows, nCols)

For i = 0 To loops
    Householder A, Q, R
    
    ' fix negatives 01Mar16
    If (False) Then
    For ii = 0 To mRows - 1
        If Q(ii, ii) < 0 Then
            ' fix Q's rows
            For jj = 0 To nCols - 1
                Q(ii, jj) = -Q(ii, jj)
            Next
            ' fix EV's cols
            For jj = 0 To mRows - 1
                EV(jj, ii) = -EV(jj, ii)
            Next
            ' fix r's cols
            For jj = 0 To mRows - 1
                R(jj, ii) = -R(jj, ii)
            Next
        End If
    Next
    End If

    EV = fmMult(EV, Q) ' Eigenvectors
    A = fmMult(R, Q)
Next

' fix negatives 01Mar16
'For i = 0 To mRows - 1
'    If Q(i, i) < 0 Then
'        ' fix Q's rows
'        For j = 0 To ncols - 1
'            Q(i, j) = -Q(i, j)
'        Next
'        ' fix EV's cols
'        For j = 0 To mRows - 1
'            EV(j, i) = -EV(j, i)
'        Next
'    End If
'Next

For i = 0 To mRows - 1
    For j = 0 To nCols - 1
        QR(i, j) = Q(i, j)
        QR(i, j + nCols) = R(i, j)
        QR(i, j + (nCols * 2)) = EV(i, j)
    Next
Next
pfIterateHouseholder = QR
End Function
Function GramSchmidt(ByRef A, ByRef R)
Dim mRows, nCols
Dim i, j, k

getNumRowsCols A, mRows, nCols
R = createMatrix(mRows, nCols)
For j = 0 To (nCols - 1)
    R(j, j) = 0
    For i = 0 To (mRows - 1)
        R(j, j) = R(j, j) + A(i, j) ^ 2
    Next
    R(j, j) = Sqr(R(j, j))
    For i = 0 To (mRows - 1)
        A(i, j) = A(i, j) / R(j, j)
    Next
    For k = (j + 1) To (nCols - 1)
        R(j, k) = 0
        For i = 0 To (mRows - 1)
            R(j, k) = R(j, k) + A(i, j) * A(i, k)
        Next
        For i = 0 To (mRows - 1)
            A(i, k) = A(i, k) - A(i, j) * R(j, k)
        Next
    Next
Next
End Function
Function pfGramSchmidt(AR As Range, Optional loops As Integer = -1)
Dim A, Q, R, i, j, E

If loops < 0 Then loops = 0 ' this is really one iteration
A = reindexRange(AR)
getNumRowsCols A, mRows, nCols
R = createIdentity(mRows, nCols)
E = createIdentity(mRows, nCols)
For i = 0 To loops
    A = fmMult(R, A)
    GramSchmidt A, R
    E = fmMult(E, A)
Next
ReDim res(mRows - 1, (3 * nCols) - 1)
For i = 0 To mRows - 1
    For j = 0 To nCols - 1
        res(i, j) = A(i, j)
        res(i, nCols + j) = R(i, j)
        res(i, 2 * nCols + j) = E(i, j)
    Next
Next
pfGramSchmidt = res
End Function


'
' VARIMAX ROTATION BASED ON IBM SPSS 23 DOCUMENTATION
'
Function calcHinv(F) ' Communality Matrix H^-1/2
Dim mRows, nCols, H
getNumRowsCols F, mRows, nCols

H = createIdentity(mRows, mRows)

For i = 0 To (mRows - 1)
    C = 0
    For j = 0 To (nCols - 1)
        C = C + F(i, j) * F(i, j) ' communality
    Next
    H(i, i) = 1 / Sqr(C) ' 1/sqrt of the communality, H^-1/2
Next
calcHinv = H
End Function
Function calcH(F) ' Communality Matrix H^1/2
Dim mRows, nCols, H
getNumRowsCols F, mRows, nCols

H = createIdentity(mRows, mRows)

For i = 0 To (mRows - 1)
    C = 0
    For j = 0 To (nCols - 1)
        C = C + F(i, j) * F(i, j) ' communality
    Next
    H(i, i) = Sqr(C) ' 1/sqrt of the communality, H^-1/2
Next
calcH = H
End Function
Function pfCommunalitySqrtInv(FR As Range) ' Factors as Range
Dim F
F = reindexRange(FR)

pfCommunalitySqrtInv = calcHinv(F)
End Function
Function calcU(L, pj, pk)
Dim mRows, nCols, i, j
getNumRowsCols L, mRows, nCols
ReDim u(mRows - 1, 0)

For i = 0 To (mRows - 1)
    u(i, 0) = (L(i, pj) ^ 2) - (L(i, pk) ^ 2)
Next
calcU = u
End Function
Function calcV(L, pj, pk)
Dim mRows, nCols, i, j
getNumRowsCols L, mRows, nCols
ReDim v(mRows - 1, 0)

For i = 0 To (mRows - 1)
    v(i, 0) = 2 * L(i, pj) * L(i, pk)
Next
calcV = v
End Function
Function createRot(P)
ReDim rot(1, 1)
        rot(0, 0) = Cos(P)
        rot(0, 1) = -Sin(P)
        rot(1, 0) = Sin(P)
        rot(1, 1) = Cos(P)
createRot = rot
End Function
Function rotateFactors(Lambda, pj, pk, rot)
Dim mRows, nCols
    getNumRowsCols Lambda, mRows, nCols
    ReDim v(mRows - 1, 1) ' remember, +1 in Javascript/C#

    For i = 0 To (mRows - 1)
        v(i, 0) = Lambda(i, pj)
        v(i, 1) = Lambda(i, pk)
    Next
    rotateFactors = fmMult(v, rot)
End Function
Function replaceFactors(Lambda, LL, pj, pk)
Dim i, mRows, nCols
    getNumRowsCols Lambda, mRows, nCols
    For i = 0 To (mRows - 1)
        Lambda(i, pj) = LL(i, 0)
        Lambda(i, pk) = LL(i, 1)
    Next
    replaceFactors = Lambda
End Function
Function fVarimax(F, loops)
Dim H, Hinv, i, j, k, Lambda, LL, mRows, nCols, Omega
Dim u, v, A, B, C, D, X, Y, P, rot
Hinv = calcHinv(F)
H = calcH(F)
Lambda = fmMult(Hinv, F)

getNumRowsCols Lambda, mRows, nCols
For Z = 1 To loops ' force 12
For i = 0 To (nCols - 2)
    For j = (i + 1) To (nCols - 1)
        ' MsgBox (i & "," & j)
        u = calcU(Lambda, i, j)
        v = calcV(Lambda, i, j)
        A = 0
        B = 0
        C = 0
        D = 0
        For k = 0 To (mRows - 1)
            A = A + u(k, 0)
            B = B + v(k, 0)
            C = C + (u(k, 0) ^ 2 - v(k, 0) ^ 2)
             D = D + (2 * u(k, 0) * v(k, 0)) ' IBM one
'            D = D + (u(k, 0) * v(k, 0)) ' WEB
        Next
        X = D - (2 * A * B) / mRows 'IBM
        Y = C - (A ^ 2 - B ^ 2) / mRows ' IBM
'        X = 2 * (D * mRows - A * B) 'WEB
'        Y = C * mRows - (A ^ 2 - B ^ 2) ' WEB
        P = 0.25 * Atn(X / Y)
        rot = createRot(P)
        LL = rotateFactors(Lambda, i, j, rot)
        Lambda = replaceFactors(Lambda, LL, i, j)
    Next
Next
Next
Omega = fmMult(H, Lambda)
fVarimax = Omega
End Function
Function pfVarimax(FR As Range, loops) ' FR is original selection
Dim F
F = reindexRange(FR)
pfVarimax = fVarimax(F, loops)
End Function
Sub allcoms()
cols = 4
For i = 1 To cols - 1
For j = i + 1 To cols
    MsgBox (i & "," & j)
Next
Next
End Sub
'
' END VARIMAX ROTATION CODE
'

'
' UTILITY
'
Sub shadeMatrix() ' shade a selected correlation matrix
Dim s As Range

Set s = Selection

For Each C In s
    If (C.value = 1#) Then
        C.Interior.Color = RGB(0, 0, 0)
    ElseIf (C.value >= 0 And C.value < 0.25) Then
        C.Interior.Color = RGB(255, 255, 255)
    ElseIf (C.value >= 0.25 And C.value < 0.5) Then
        C.Interior.Color = RGB(0, 127, 0)
    ElseIf (C.value >= 0.5 And C.value < 0.75) Then
        C.Interior.Color = RGB(0, 191, 0)
    ElseIf (C.value >= 0.75 And C.value < 1#) Then
        C.Interior.Color = RGB(0, 255, 0)
    ElseIf (C.value >= -1# And C.value < -0.75) Then
        C.Interior.Color = RGB(255, 0, 0)
    ElseIf (C.value >= -0.75 And C.value < -0.5) Then
        C.Interior.Color = RGB(191, 0, 0)
    ElseIf (C.value >= -0.5 And C.value < -0.25) Then
        C.Interior.Color = RGB(127, 0, 0)
    ElseIf (C.value >= -0.25 And C.value < 0) Then
        C.Interior.Color = RGB(255, 255, 255)
    End If
Next
End Sub

Sub calcVariance() ' calculate the total variance accounted for. Input=a square Eigenvalue matrix selection
Dim s As Range
Dim i, rows, cols As Integer
Dim v, F As Double

Set s = Selection
rows = s.rows.Count
cols = s.Columns.Count

For i = 1 To cols
    v = s.Cells(i, i)
    F = v / rows
    s.Cells(rows + 1, i) = F
Next
End Sub

Sub completeMatrix() ' complete the Excel correlation upper triangle
Dim s As Range
Dim R, C As Integer
Set s = Selection

For R = 1 To s.rows.Count
    For C = 1 To R
        s.Cells(C, R) = s.Cells(R, C)
    Next
Next
End Sub

Sub pcprep() ' Given a Selected correlation matrix, Prepares Width*3 selection for output
Dim s As Range
Set s = Selection

newcolL = s.Columns.Count + 2
newcolH = newcolL + s.Columns.Count * 3 - 1

ActiveSheet.Range(s.Cells(1, newcolL), s.Cells(s.rows.Count, newcolH)).Select
End Sub

Sub removeBlanks() ' removes entire row if selected column cell is blank. Input: selected column cell
Dim s As Range
Dim i, mrow As Integer

Set s = Selection
mrow = s.rows.Count
frow = s.row
arow = frow + mrow - 1

For i = mrow To 1 Step -1 ' max row down to
    If (s.Cells(i, 1) = "") Then
        Application.ActiveSheet.rows(arow).Delete
    Else
    End If
    arow = arow - 1
Next
End Sub

Sub simplify() ' simply principal components
Dim s As Range
Dim rows, cols, i, j, v, max, cutoff, quartile
Set s = Selection

rows = s.rows.Count
cols = s.Columns.Count

For j = 1 To cols
    max = 0
    For i = 1 To rows
        If Abs(s.Cells(i, j)) > max Then max = Abs(s.Cells(i, j))
    Next
    cutoff = max * 0.5
    quartile = max * 0.25
    For i = 1 To rows
        If Abs(s.Cells(i, j)) >= cutoff Then
            If s.Cells(i, j) > 0 Then
                s.Cells(i, j) = "+"
            Else
                s.Cells(i, j) = "-"
            End If
        ElseIf Abs(s.Cells(i, j)) >= quartile Then
            If s.Cells(i, j) > 0 Then
                s.Cells(i, j) = "(+)"
            Else
                s.Cells(i, j) = "(-)"
            End If
        Else
            s.Cells(i, j) = ""
        End If
    Next
Next
End Sub
