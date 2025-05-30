VERSION 5.00
Begin VB.Form confirma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atenção!"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Confirmação.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Repetir"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ignorar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "confirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Hide
Timer1.Interval = 0
Backup.Enabled = True
Backup.SetFocus
Backup.Data4.Recordset.MoveNext
Backup.CoPiA
End Sub

Private Sub Command2_Click()
Me.Hide
Timer1.Interval = 0
Backup.Enabled = True
Backup.SetFocus
Backup.CoPiA
End Sub

Private Sub Command3_Click()
Me.Hide
Timer1.Interval = 0
Backup.Enabled = True
Backup.SetFocus
End Sub

Private Sub Form_Activate()
Left = (Screen.Width - Width) / 2
Top = ((Screen.Height - Height) / 2) - 1000
Backup.Enabled = False
Timer1.Interval = 5
confirma.Enabled = True
End Sub

Private Sub Timer1_Timer()
Beep
End Sub
