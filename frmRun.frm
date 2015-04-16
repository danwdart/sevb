VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run..."
   ClientHeight    =   1095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "exe"
      DialogTitle     =   "Open File"
      FileName        =   "*.exe"
      Filter          =   "Executable (*.exe)"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
txtFile = CommonDialog1.FileName
End Sub

Private Sub OKButton_Click()
On Error Resume Next
Shell txtFile.Text, vbNormalFocus
End Sub
