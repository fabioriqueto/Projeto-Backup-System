VERSION 5.00
Begin VB.Form Insert3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destino da Cópia"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "Insert3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Atribuir destino da cópia ao projeto."
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Megatron Explorer"
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   5775
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   5775
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label12 
         Caption         =   "Diretório de destino da cópia."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Insert3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo erross
ChDir Text1
Backup.Data1.Recordset.MoveFirst
Do While Not Backup.Data1.Recordset.EOF
   If Backup.Data1.Recordset.Fields("Codigo") = CodLink Then
      Backup.Data1.Recordset.Edit
      Backup.Data1.Recordset.Fields("Destino") = Trim(Text1)
      Backup.Data1.Recordset.Update
      Exit Do
   End If
   Backup.Data1.Recordset.MoveNext
Loop
Backup.Data1.Refresh
Backup.Enabled = True
Me.Hide
Backup.attt
erross:
   If Err.Number = 76 Then
      'result = MsgBox("Pasta não encontrada. Pretende cria-la?", vbYesNo)
      MsgBox "Pasta não encontrada."
      'If result = vbYes Then MkDir "c:\fabio\fabio2\teste1"   ' Make new directory or folder.
      Exit Sub
   End If
End Sub

Private Sub Dir1_Change()
Text1 = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveErrs
Dir1.Path = Drive1.Drive
Exit Sub
DriveErrs:
    Select Case Err
        Case 68
            MsgBox prompt:="O drive não esta pronto. Insira o disco no drive.", _
            Buttons:=vbExclamation
            ' Reset path to previous drive.
            Drive1.Drive = Dir1.Path
            Exit Sub
        Case Else
            MsgBox prompt:="Application error.", Buttons:=vbExclamation
    End Select
End Sub

Private Sub Form_Activate()
Left = (Screen.Width - Width) / 2
Top = ((Screen.Height - Height) / 2) - 500
Dir1.Path = Drive1.Drive
Text1 = Dir1.Path
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Backup.Data2.Refresh
Backup.Enabled = True
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Backup.Data2.Refresh
Backup.Enabled = True
Me.Hide
End Sub
