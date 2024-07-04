VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Calculator By Pollob Roy"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command16 
      BackColor       =   &H80000012&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   4080
      MaskColor       =   &H00FFFF80&
      TabIndex        =   18
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H80000012&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   4080
      MaskColor       =   &H00FFFF80&
      TabIndex        =   17
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H80000012&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   16
      Left            =   3120
      MaskColor       =   &H00FFFF80&
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H80000012&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   4080
      MaskColor       =   &H00FFFF80&
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000012&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   3120
      MaskColor       =   &H00FFFF80&
      TabIndex        =   14
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H80000012&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   4080
      MaskColor       =   &H00FFFF80&
      TabIndex        =   13
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H80000012&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   2160
      MaskColor       =   &H00FFFF80&
      TabIndex        =   12
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000012&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   1200
      MaskColor       =   &H00FFFF80&
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   240
      MaskColor       =   &H00FFFF80&
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000012&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   2160
      MaskColor       =   &H00FFFF80&
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000012&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   1200
      MaskColor       =   &H00FFFF80&
      TabIndex        =   8
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   240
      MaskColor       =   &H00FFFF80&
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000012&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   2160
      MaskColor       =   &H00FFFF80&
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000012&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   1200
      MaskColor       =   &H00FFFF80&
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000012&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   240
      MaskColor       =   &H00FFFF80&
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000012&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2160
      MaskColor       =   &H00FFFF80&
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000012&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1200
      MaskColor       =   &H00FFFF80&
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   240
      MaskColor       =   &H00FFFF80&
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op, fn As Integer
Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text & 7
End Sub

Private Sub Command10_Click(Index As Integer)
Text1.Text = Text1.Text & 0
End Sub

Private Sub Command11_Click(Index As Integer)
Text1.Text = Text1.Text & (".")
End Sub

Private Sub Command12_Click(Index As Integer)
Text1.Text = ""
End Sub

Private Sub Command13_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command14_Click(Index As Integer)
op = 1
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command15_Click(Index As Integer)
op = 2
fn = Text1.Text
Text1.Text = ""

End Sub

Private Sub Command16_Click(Index As Integer)
op = 3
fn = Text1.Text
Text1.Text = ""

End Sub

Private Sub Command17_Click(Index As Integer)
op = 4
fn = Text1.Text
Text1.Text = ""

End Sub

Private Sub Command18_Click(Index As Integer)
If op = 1 Then
Text1.Text = Val(fn) + Val(Text1.Text)
ElseIf op = 2 Then
Text1.Text = Val(fn) - Val(Text1.Text)
ElseIf op = 3 Then
Text1.Text = Val(fn) * Val(Text1.Text)
ElseIf op = 4 Then
Text1.Text = Val(fn) / Val(Text1.Text)
End If
End Sub

Private Sub Command2_Click(Index As Integer)
Text1.Text = Text1.Text & 8
End Sub

Private Sub Command3_Click(Index As Integer)
Text1.Text = Text1.Text & 9
End Sub

Private Sub Command4_Click(Index As Integer)
Text1.Text = Text1.Text & 4
End Sub

Private Sub Command5_Click(Index As Integer)
Text1.Text = Text1.Text & 5
End Sub

Private Sub Command6_Click(Index As Integer)
Text1.Text = Text1.Text & 6
End Sub

Private Sub Command7_Click(Index As Integer)
Text1.Text = Text1.Text & 1
End Sub

Private Sub Command8_Click(Index As Integer)
Text1.Text = Text1.Text & 2
End Sub

Private Sub Command9_Click(Index As Integer)
Text1.Text = Text1.Text & 3
End Sub
