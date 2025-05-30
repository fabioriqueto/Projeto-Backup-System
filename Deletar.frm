VERSION 5.00
Begin VB.Form Deletar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exclusão de Projeto"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "Deletar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Projeto a ser excluido"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Deletar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
Backup.Enabled = True
Backup.Data1.Recordset.MoveFirst
Do
   If Backup.Data1.Recordset.EOF Then Exit Do
   If Backup.Data1.Recordset.Fields("Codigo") = Val(Left(Backup.List1.Text, 4)) Then
      Backup.Data1.Recordset.Delete
      Backup.Data1.Refresh
      Backup.List1.RemoveItem (Backup.List1.ListIndex)
      Exit Do
   End If
   Backup.Data1.Recordset.MoveNext
Loop
Backup.Data1.Refresh
Me.Hide
Backup.Label2 = Trim(Backup.List1.ListCount) & " Projetos Elaborados."
End Sub

Private Sub Command7_Click()
Backup.Enabled = True
Me.Hide
End Sub

Private Sub Form_Activate()
Left = (Screen.Width - Width) / 2
Top = ((Screen.Height - Height) / 2) - 500
End Sub
