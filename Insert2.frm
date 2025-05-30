VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Insert2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adicionador de Arquivos"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "Insert2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "Insert2.frx":0442
      Height          =   1620
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2858
      _Version        =   327680
      ListField       =   "Desc"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleção de Arquivos / Megatron Explorer"
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command4 
         Caption         =   "Adicionar todos (*.*)"
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   5775
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   5775
      End
      Begin VB.Label Label12 
         Caption         =   "Diretório de Origem:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Arquivo:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   735
      End
   End
End
Attribute VB_Name = "Insert2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Data2.Recordset.AddNew
Data2.Recordset.Fields("Link") = CodLink
Data2.Recordset.Fields("Arquivo") = "*.*"
Data2.Recordset.Fields("Origem") = Trim(Label10)
Data2.Recordset.Fields("Desc") = Trim(Label10) & "\*.*"
Data2.Recordset.Update
Data2.Refresh
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Label10 = Dir1.Path
Label11 = ""
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveErrs
Dir1.Path = Drive1.Drive
Label11 = ""
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

Private Sub File1_Click()
Label11 = File1
End Sub

Private Sub File1_DblClick()
Data2.Recordset.AddNew
Data2.Recordset.Fields("Link") = CodLink
Data2.Recordset.Fields("Arquivo") = Trim(Label11)
Data2.Recordset.Fields("Origem") = Trim(Label10)
Data2.Recordset.Fields("Desc") = Trim(Label10) & "\" & Trim(Label11)
Data2.Recordset.Update
Data2.Refresh
End Sub

Private Sub Form_Activate()
Left = (Screen.Width - Width) / 2
Top = ((Screen.Height - Height) / 2) - 500
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Set rs2 = db1.OpenRecordset("SELECT *FROM Propriedades WHERE Propriedades.Link = " & CodLink)
Set Data2.Recordset = rs2

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
