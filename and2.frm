VERSION 5.00
Begin VB.Form SimpleCalc 
   Caption         =   "Simple Calculator"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton clear 
      Caption         =   "C"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton tans 
      Caption         =   "TAN"
      Height          =   495
      Left            =   3120
      TabIndex        =   21
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cosd 
      Caption         =   "COS"
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton sind 
      Caption         =   "SIN"
      Height          =   495
      Left            =   1200
      TabIndex        =   19
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton logs 
      Caption         =   "LOG"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton sqrs 
      Caption         =   "SQR"
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton fixf 
      Caption         =   "FIX"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton sgns 
      Caption         =   "SGN"
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton absab 
      Caption         =   "ABS"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton arcta 
      Caption         =   "ATN"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton modul 
      Caption         =   "MOD"
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton power 
      Caption         =   "^"
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton intdiv 
      Caption         =   "\"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton divide 
      Caption         =   "/"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton mult 
      Caption         =   "*"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton plus 
      Caption         =   "+"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox num2 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox num1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label result 
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ANSWER:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "INPUT 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "INPUT 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "SimpleCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var1, var2 As Double

Private Sub absab_Click()
    If Not result.Caption = "" Then
        result.Caption = Abs(result.Caption)
    End If
End Sub

Private Sub arcta_Click()
    If Not result.Caption = "" Then
        result.Caption = Atn(result.Caption)
    End If
End Sub

Private Sub clear_Click()
num1.Text = ""
num2.Text = ""
result.Caption = ""
End Sub

Private Sub cosd_Click()
    If Not result.Caption = "" Then
        result.Caption = Cos(result.Caption)
    End If
End Sub

Private Sub divide_Click()
result.Caption = num1.Text / num2.Text
num1.Text = ""
num2.Text = ""
End Sub

Private Sub fixf_Click()
    If Not result.Caption = "" Then
        result.Caption = Fix(result.Caption)
    End If
End Sub

Private Sub intdiv_Click()
    result.Caption = num1.Text \ num2.Text
    num1.Text = ""
    num2.Text = ""
End Sub

Private Sub logs_Click()
    If Not result.Caption = "" Then
        result.Caption = Log(result.Caption)
    End If
End Sub

Private Sub minus_Click()
    result.Caption = num1.Text - num2.Text
    num1.Text = ""
    num2.Text = ""
End Sub

Private Sub modul_Click()
    result.Caption = num1.Text Mod num2.Text
    num1.Text = ""
    num2.Text = ""
End Sub

Private Sub mult_Click()
    result.Caption = num1.Text * num2.Text
    num1.Text = ""
    num2.Text = ""
End Sub

Private Sub plus_Click()
    var1 = Val(num1.Text)
    var2 = Val(num2.Text)
    result.Caption = var1 + var2
    num1.Text = ""
    num2.Text = ""
End Sub

Private Sub power_Click()
    result.Caption = num1.Text ^ num2.Text
    num1.Text = ""
    num2.Text = ""
End Sub

Private Sub sgns_Click()
    If Not result.Caption = "" Then
        result.Caption = Sgn(result.Caption)
    End If
End Sub

Private Sub sind_Click()
    If Not result.Caption = "" Then
        result.Caption = Sin(result.Caption)
    End If
End Sub

Private Sub sqrs_Click()
    If Not result.Caption = "" Then
        result.Caption = Sqr(result.Caption)
    End If
End Sub


Private Sub tans_Click()
    If Not result.Caption = "" Then
        result.Caption = Tan(result.Caption)
    End If
End Sub
