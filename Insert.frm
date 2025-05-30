VERSION 5.00
Begin VB.Form Insert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserção de Novo Projeto"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "Insert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      MaxLength       =   60
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descreva aqui o nome para o seu novo projeto."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Backup.Enabled = True
If Trim(Text1) <> "" Then
   Backup.Data1.Recordset.AddNew
   Backup.Data1.Recordset.Fields("DescProjeto") = Trim(Text1)
   Backup.Data1.Recordset.Update
   Backup.Data1.Refresh
   Backup.Data1.Recordset.MoveLast
   Backup.List1.AddItem (Right("0000" & Trim(Backup.Data1.Recordset.Fields("Codigo")), 4) & " - " & Backup.Data1.Recordset.Fields("DescProjeto"))
End If
Me.Hide
Backup.Label2 = Trim(Backup.List1.ListCount) & " Projetos Elaborados."
End Sub

Private Sub Command5_Click()
Backup.Enabled = True
Me.Hide
End Sub

Private Sub Form_Activate()
Left = (Screen.Width - Width) / 2
Top = ((Screen.Height - Height) / 2) - 500
Text1 = ""
End Sub

