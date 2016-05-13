VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Посчитать методом Зеиделя"
      Height          =   615
      Left            =   8400
      TabIndex        =   42
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Посчитать методом Зеиделя"
      Height          =   615
      Left            =   4560
      TabIndex        =   33
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   22
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   16
      Left            =   1920
      TabIndex        =   21
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   3
      Left            =   2760
      TabIndex        =   20
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   19
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   18
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   15
      Left            =   1320
      TabIndex        =   17
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   14
      Left            =   720
      TabIndex        =   16
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   13
      Left            =   120
      TabIndex        =   15
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   12
      Left            =   1920
      TabIndex        =   14
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   11
      Left            =   1320
      TabIndex        =   13
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   10
      Left            =   720
      TabIndex        =   12
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   8
      Left            =   1920
      TabIndex        =   10
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   7
      Left            =   1320
      TabIndex        =   9
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   720
      TabIndex        =   8
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Посчитать Методом Гаусса"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ввод"
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   8040
      ScaleHeight     =   2235
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label25 
      Height          =   495
      Left            =   8640
      TabIndex        =   52
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label24 
      Caption         =   "x4="
      Height          =   255
      Left            =   8160
      TabIndex        =   51
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label23 
      Caption         =   "x3="
      Height          =   255
      Left            =   8160
      TabIndex        =   50
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label22 
      Caption         =   "x2="
      Height          =   255
      Left            =   8160
      TabIndex        =   49
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label21 
      Caption         =   "x1="
      Height          =   255
      Left            =   8160
      TabIndex        =   48
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label20 
      Height          =   495
      Left            =   8640
      TabIndex        =   47
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label19 
      Height          =   495
      Left            =   8640
      TabIndex        =   46
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label18 
      Height          =   495
      Left            =   8640
      TabIndex        =   45
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   10200
      TabIndex        =   44
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "Колличество итераций"
      Height          =   255
      Left            =   8160
      TabIndex        =   43
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Line Line8 
      X1              =   8040
      X2              =   7680
      Y1              =   1200
      Y2              =   960
   End
   Begin VB.Line Line11 
      X1              =   7920
      X2              =   8040
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line9 
      X1              =   7920
      X2              =   7680
      Y1              =   1320
      Y2              =   1560
   End
   Begin VB.Line Line7 
      X1              =   7320
      X2              =   7800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line6 
      X1              =   7320
      X2              =   7800
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   3960
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   3960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label15 
      Caption         =   "x4="
      Height          =   255
      Left            =   480
      TabIndex        =   41
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "x3="
      Height          =   255
      Left            =   480
      TabIndex        =   40
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "x2="
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "x1="
      Height          =   255
      Left            =   480
      TabIndex        =   38
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   37
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   36
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   35
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   34
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   7320
      X2              =   7320
      Y1              =   2520
      Y2              =   9000
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   3960
      Y1              =   2520
      Y2              =   8880
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11640
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label10 
      Caption         =   "Колличество итераций"
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   6600
      TabIndex        =   31
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   4800
      TabIndex        =   30
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   4800
      TabIndex        =   29
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   4800
      TabIndex        =   28
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "x1="
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "x2="
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "x3="
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "x4="
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   4800
      TabIndex        =   23
      Top             =   4440
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub functOpen()
Dim I As Integer, J As Integer
Dim K As Double
Dim Z(1 To 10) As String
Z(1) = "x1"
Z(2) = "x2"
Z(3) = "x3"
Z(4) = "x4"
Z(5) = "x4"
Z(6) = "x5"
Z(7) = "x6"
Z(8) = "x7"
Z(9) = "x8"
Z(10) = "x9"

K = 0
For I = 1 To 4
For J = 1 To 4
K = K + 1
A(I, J) = Val(Text1(K))

Next
A(I, 5) = Val(Text2(I))
Next

For I = 1 To 4
B(I) = A(I, 5)
Next
      Picture1.Cls
      For I = 1 To 4
       For J = 1 To 4
       K = A(I, J)
       If K > 0 Then
        If J > 1 Then
          Picture1.Print "+"; K; Z(J);
        Else
          Picture1.Print K; Z(J);
        End If
       Else
        Picture1.Print K; Z(J);
        End If
        Next
       Picture1.Print "="; B(I)
      Next
End Sub

Private Sub Command1_Click()
Call functOpen

End Sub



Private Sub Command2_Click()
Z(1) = "x1"
Z(2) = "x2"
Z(3) = "x3"
Z(4) = "x4"

Call GaussSolveM(A(), X())
If Result = False Then
Label11(0).Caption = "Система не имеет решений!"
Else
For I = 1 To 4
Label11(I - 1).Caption = X(I)
Next
End If
End Sub

Private Sub Command3_Click()

Dim M(4, 4) As Double
Dim V(4) As Double
Dim X(4) As Double
Dim XP(4) As Double

K = 0
For I = 1 To 4
For J = 1 To 4
K = K + 1
M(I, J) = Val(Text1(K))

Next
V(I) = Val(Text2(I))
Next
X(1) = 0
XP(1) = 1: XP(2) = XP(3) = XP(4) = XP(1)
Do
iter = iter + 1
For ii = 1 To 4
XP(ii) = X(ii)
Next
X(1) = (V(1) - XP(2) * M(1, 2) - XP(3) * M(1, 3) - XP(4) * M(1, 4)) / (M(1, 1))
X(2) = (V(2) - X(1) * M(2, 1) - XP(3) * M(2, 3) - XP(4) * M(2, 4)) / (M(2, 2))
X(3) = (V(3) - X(2) * M(3, 2) - X(1) * M(3, 1) - XP(4) * M(3, 4)) / (M(3, 3))
X(4) = (V(4) - X(2) * M(4, 2) - X(3) * M(4, 3) - X(1) * M(4, 1)) / (M(4, 4))
Loop While Abs(Abs(X(3)) - Abs(XP(3))) > 0.00001
Label5.Caption = iter

Label1.Caption = "0" + Str(X(1))
Label2.Caption = "0" + Str(X(2))
Label3.Caption = "0" + Str(X(3))
Label4.Caption = "0" + Str(X(4))
End Sub


Private Sub Command4_Click()
Dim M(4, 4) As Double
Dim V(4) As Double
Dim X(4) As Double
Dim XP(4) As Double
K = 0
For I = 1 To 4
For J = 1 To 4
K = K + 1
M(I, J) = Val(Text1(K))

Next
V(I) = Val(Text2(I))
Next
X(1) = 0
XP(1) = 1: XP(2) = XP(3) = XP(4) = XP(1)

Do
iter = iter + 1
For ii = 1 To 4
XP(ii) = X(ii)
Next
X(1) = (V(1) - XP(2) * M(1, 2) - XP(3) * M(1, 3) - XP(4) * M(1, 4)) / (M(1, 1))
X(2) = (V(2) - XP(1) * M(2, 1) - XP(3) * M(2, 3) - XP(4) * M(2, 4)) / (M(2, 2))
X(3) = (V(3) - XP(2) * M(3, 2) - XP(1) * M(3, 1) - XP(4) * M(3, 4)) / (M(3, 3))
X(4) = (V(4) - XP(2) * M(4, 2) - XP(3) * M(4, 3) - XP(1) * M(4, 1)) / (M(4, 4))
Loop While Abs(Abs(X(3)) - Abs(XP(3))) > 0.00001
Label17.Caption = iter

Label25.Caption = "0" + Str(X(1))
Label18.Caption = "0" + Str(X(2))
Label19.Caption = "0" + Str(X(3))
Label20.Caption = "0" + Str(X(4))



End Sub

