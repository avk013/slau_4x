Attribute VB_Name = "Module1"
Public X(1 To 5) As Double
Public Z(1 To 5) As String
Public B(1 To 5) As Double
Public A(1 To 5, 1 To 5) As Double

Public Function GaussSolveM(ByRef A_() As Double, _
         ByRef X() As Double) As Boolean
    Dim Result As Boolean
    Dim A() As Double
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim M As Double
    Dim T As Double
    A = A_

    Result = True
    For I = 1 To 4 Step 1
        K = I
        M = Abs(A(I, I))
        For J = I + 1 To 4 Step 1
            If M < Abs(A(J, I)) Then
                M = Abs(A(J, I))
                K = J
            End If
        Next J
        If Abs(M) > 0 Then
            For J = I To 4 + 1 Step 1
                T = A(I, J)
                A(I, J) = A(K, J)
                A(K, J) = T
            Next J
            For K = I + 1 To 4 Step 1
                T = A(K, I) / A(I, I)
                A(K, I) = 0
                For J = I + 1 To 4 + 1 Step 1
                    A(K, J) = A(K, J) - T * A(I, J)
                Next J
            Next K
        Else
            Result = False
            Exit For
        End If
    Next I
    If Result Then
        I = 4
        Do
            X(I) = A(I, 4 + 1)
            J = I + 1
            Do While J <= 4
                X(I) = X(I) - A(I, J) * X(J)
                J = J + 1
            Loop
            X(I) = X(I) / A(I, I)
            I = I - 1
        Loop Until Not I >= 1
    End If

    GaussSolveM = Result
End Function


