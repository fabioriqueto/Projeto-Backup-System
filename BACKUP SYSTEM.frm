VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Back-up System Megatron V.1.0"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "BACKUP SYSTEM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Backup System\Backup01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Projetos"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Apresentação"
      TabPicture(0)   =   "BACKUP SYSTEM.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Utilitários"
      TabPicture(1)   =   "BACKUP SYSTEM.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "List1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Propriedades de Cópias"
      TabPicture(2)   =   "BACKUP SYSTEM.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "MaskEdBox1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "MaskEdBox2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Data3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Timer1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Data4"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3480
         Top             =   1680
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   1440
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "hh:mm"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "hh:mm"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73560
         TabIndex        =   14
         Top             =   2760
         Width           =   6135
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Destino da cópia"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71280
         TabIndex        =   12
         Top             =   3240
         Width           =   1695
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   -74880
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   600
         Width           =   7455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir Projeto"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73080
         TabIndex        =   1
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo Projeto"
         Height          =   375
         Left            =   -74880
         TabIndex        =   0
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Efetuar a segunda cópia as :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Efetuar a primeira cópia as :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Horarios para as cópias selecionadas em utilitários."
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label6 
         Caption         =   "Destino da Cópia:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Licenciado à: Licença de uso"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   -74880
         TabIndex        =   10
         Top             =   2520
         Width           =   7455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Back-up System ""MEGATRON"" V.1.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   -74880
         TabIndex        =   9
         Top             =   1080
         Width           =   7455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "( XX ) Projetos elaborados."
         Height          =   255
         Left            =   -69360
         TabIndex        =   8
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Lista dos Projetos para cópia de segurança."
         Height          =   255
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de arquivos"
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   7695
      Begin MSDBCtls.DBList DBList1 
         Bindings        =   "BACKUP SYSTEM.frx":0496
         Height          =   1425
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2514
         _Version        =   393216
         ListField       =   "Desc"
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Excluir Arquivo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Inserir Arquivo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim volta As Integer
Dim SourceFile, DestinationFile

Private Sub Command1_Click()
Insert.Show
Me.Enabled = False
End Sub

Private Sub Command2_Click()
Deletar.Label9.Caption = List1.Text
Deletar.Show
Me.Enabled = False
End Sub

Private Sub Command3_Click()
Insert2.Show
Me.Enabled = False
End Sub

Private Sub Command4_Click()
Me.Enabled = False
Insert3.Show
End Sub

Private Sub Command5_Click()
Data3.Recordset.Edit
Data3.Recordset.Fields("H1") = MaskEdBox1.FormattedText
Data3.Recordset.Fields("H2") = MaskEdBox2.FormattedText
Data3.Recordset.Update
End Sub

Private Sub Command8_Click()
result = MsgBox("Pretende excluir este registro?", vbYesNo)
If result = vbYes Then Data2.Recordset.Delete
If result = vbYes Then Data2.Refresh
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) / 2
Top = ((Screen.Height - Height) / 2) - 1000
CodLink = 0
Set db1 = OpenDatabase(App.Path & "\Backup01.mdb")
Set rs1 = db1.OpenRecordset("Projetos")
Set Data1.Recordset = rs1
If Data1.Recordset.RecordCount <> 0 Then Data1.Recordset.MoveFirst
Set rs3 = db1.OpenRecordset("Hora")
Set Data3.Recordset = rs3
Data3.Recordset.MoveFirst
MaskEdBox1 = Data3.Recordset.Fields("H1")
MaskEdBox2 = Data3.Recordset.Fields("H2")
Dim ConTa As Integer
ConTa = 0
Do While Not Data1.Recordset.EOF
   List1.AddItem (Right("0000" & Trim(Data1.Recordset.Fields("Codigo")), 4) & " - " & Data1.Recordset.Fields("DescProjeto"))
   If Data1.Recordset.Fields("Sel") = True Then List1.Selected(ConTa) = True
   ConTa = ConTa + 1
   Data1.Recordset.MoveNext
Loop
Set rs2 = db1.OpenRecordset("SELECT *FROM Propriedades WHERE Link = " & CodLink & " ORDER BY ARQUIVO")
Set Data2.Recordset = rs2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()
SelCod
attt
End Sub

Private Sub List1_GotFocus()
attt
DBList1.Visible = True
Command3.Enabled = True
Command4.Enabled = True
Command2.Enabled = True
Command8.Enabled = True
End Sub

'**************************************************
'**************************************************
'This example uses the ChDir statement to change the current directory or folder.
' Change current directory or folder to "MYDIR".
'ChDir "MYDIR"
' Assume "C:" is the current drive. The following statement changes
' the default directory on drive "D:". "C:" remains the current drive.
'ChDir "D:\WINDOWS\SYSTEM"
'**************************************************
'**************************************************
'This example uses the CurDir function to return the current path.
' Assume current path on C drive is "C:\WINDOWS\SYSTEM".
' Assume current path on D drive is "D:\EXCEL".
' Assume C is the current drive.
'Dim MyPath
'MyPath = CurDir ' Returns "C:\WINDOWS\SYSTEM".
'MyPath = CurDir("C")    ' Returns "C:\WINDOWS\SYSTEM".
'MyPath = CurDir("D")    ' Returns "D:\EXCEL".
'**************************************************
'**************************************************
'This example uses the MkDir statement to create a directory or folder. If the drive is not specified, the new directory or folder is created on the current drive.
'MkDir "MYDIR"   ' Make new directory or folder.
'**************************************************
'**************************************************
'This example uses the FileCopy statement to copy one file to another. For purposes of this example, assume that SRCFILE is a file containing some data.
'Dim SourceFile, DestinationFile
'SourceFile = "SRCFILE"  ' Define source file name.
'DestinationFile = "DESTFILE"    ' Define target file name.
'FileCopy SourceFile, DestinationFile    ' Copy source to target.
'**************************************************
'**************************************************
'This example uses the Kill statement to delete a file from a disk.
' Assume TESTFILE is a file containing some data.
'Kill "TestFile" ' Delete file.
' Delete all *.TXT files in current directory.
'Kill "*.TXT"
'**************************************************
'**************************************************
'This example uses the Name statement to rename a file. For purposes of this example, assume that the directories or folders that are specified already exist.
'Dim OldName, NewName
'OldName = "OLDFILE": NewName = "NEWFILE"    ' Define filenames.
'Name OldName As NewName ' Rename file.
'*****************************************
'OldName = "C:\MYDIR\OLDFILE": NewName = "C:\YOURDIR\NEWFILE"
'Name OldName As NewName ' Move and rename file.

Private Sub List1_ItemCheck(Item As Integer)
attt
SelCod
Data1.Recordset.MoveFirst
Do
   If Data1.Recordset.EOF = True Then Exit Do
   If Data1.Recordset.Fields("Codigo") = Val(Left(List1.Text, 4)) Then
      Data1.Recordset.Edit
      If List1.Selected(List1.ListIndex) = True Then Data1.Recordset.Fields("Sel") = True
      If List1.Selected(List1.ListIndex) = False Then Data1.Recordset.Fields("Sel") = False
      Data1.Recordset.Update
      If Data1.Recordset.Fields("Destino") <> "" Then
         Text1 = Data1.Recordset.Fields("Destino")
      Else:
         Text1 = ""
      End If
      Exit Do
   End If
   Data1.Recordset.MoveNext
Loop
End Sub

Sub SelCod()
CodLink = Val(Left(List1.Text, 4))
Set rs2 = db1.OpenRecordset("SELECT *FROM Propriedades WHERE Link = " & CodLink & " ORDER BY ARQUIVO")
Set Data2.Recordset = rs2
Label2 = Trim(List1.ListCount) & " Projetos Elaborados."
End Sub

Sub attt()
Data1.Recordset.MoveFirst
Do
   If Data1.Recordset.EOF = True Then Exit Do
   If Data1.Recordset.Fields("Codigo") = Val(Left(List1.Text, 4)) Then
      If Data1.Recordset.Fields("Destino") <> "" Then
         Text1 = Data1.Recordset.Fields("Destino")
      Else:
         Text1 = ""
      End If
      Exit Do
   End If
   Data1.Recordset.MoveNext
Loop
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Then
   Backup.Height = 4290
ElseIf SSTab1.Tab = 1 Then
   Backup.Height = 6690
End If
End Sub

Private Sub Timer1_Timer()
If Time = Data3.Recordset.Fields("H1") Or Time = Data3.Recordset.Fields("H2") Then
   Set rs4 = db1.OpenRecordset("SELECT Projetos.Sel, Projetos.Destino, Propriedades.Origem, Propriedades.Arquivo, Projetos.DescProjeto FROM Projetos INNER JOIN Propriedades ON Projetos.Codigo = Propriedades.Link WHERE Projetos.Sel = True order by PROPRIEDADES.ARQUIVO")
   Set Data4.Recordset = rs4
   Data4.Refresh
   Data4.Recordset.MoveFirst
   CoPiA
End If
End Sub

Sub CoPiA()
On Error GoTo utilizaerro
If Data4.Recordset.RecordCount <> 0 Then
   Do While Not Data4.Recordset.EOF
      SourceFile = Data4.Recordset.Fields("Origem") & "\" & Data4.Recordset.Fields("Arquivo")
      DestinationFile = Data4.Recordset.Fields("Destino") & "\" & Data4.Recordset.Fields("Arquivo")
      FileCopy SourceFile, DestinationFile
      Data4.Recordset.MoveNext
   Loop
End If
utilizaerro:
    If Err.Number = 70 Then
       confirma.Show
       confirma.Label1.Caption = "'" & SourceFile & "'" & " Este arquivo pode estar sendo utilizado, para continuar feche todos os programas e cliqie em <Repetir>."
       confirma.SetFocus
       Exit Sub
    ElseIf Err.Number = 53 Or Err.Number = 76 Then
       Data4.Recordset.MoveNext
       CoPiA
       Exit Sub
    End If
End Sub

Private Sub Timer2_Timer()
Beep
End Sub
