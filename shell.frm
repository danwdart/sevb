VERSION 5.00
Begin VB.Form frmShell
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   2460
   ClientTop       =   2070
   ClientWidth     =   8250
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1
      Caption         =   "Turn Off"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   4920
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command1
      Caption         =   "Media"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   4920
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1
      Caption         =   "Games"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1
      Caption         =   "Work"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command1
      Caption         =   "Internet"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1
      Caption         =   "Change"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
MsgBox "Change?? What??", vbCritical + vbOKOnly, "Silent Echoes"
ElseIf Index = 3 Then
MsgBox "For now..."
Shell "Explorer C:\Documents and Settings\All Users\Start Menu\Programs\Games"

ElseIf Index = 5 Then
End
End If
End Sub
