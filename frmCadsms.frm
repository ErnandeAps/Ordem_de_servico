VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadsms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de contatos sms"
   ClientHeight    =   3375
   ClientLeft      =   5505
   ClientTop       =   3945
   ClientWidth     =   6870
   Icon            =   "frmCadsms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInserir 
      Height          =   405
      Left            =   5760
      Picture         =   "frmCadsms.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   150
      Width           =   495
   End
   Begin VB.CommandButton cmdExcItemAc 
      Height          =   405
      Left            =   6270
      Picture         =   "frmCadsms.frx":1493
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   150
      Width           =   495
   End
   Begin VB.TextBox ctFone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3150
      MaxLength       =   25
      TabIndex        =   1
      Top             =   210
      Width           =   2055
   End
   Begin VB.TextBox ctNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      MaxLength       =   25
      TabIndex        =   0
      Top             =   210
      Width           =   3045
   End
   Begin MSDataGridLib.DataGrid dbgCadSms 
      Bindings        =   "frmCadsms.frx":19DE
      Height          =   2295
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "fone"
         Caption         =   "Fone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3119,811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2085,166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DataAdocadsms 
      Height          =   405
      Left            =   30
      Top             =   2940
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Acessorios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fone"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3150
      TabIndex        =   4
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "frmCadsms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdloc_Click()

End Sub

Private Sub cmdExcItemAc_Click()
Dim nIdac As Integer
On Error GoTo E:

nIdac = InputBox("Digite o ID do ítem para continuar.")

Sql = "DELETE from contatossms where id=" & nIdac & ""

conn.Execute Sql

DataAdocadsms.Refresh

E:
If Err.Number = 13 Then
    MsgBox "Digite o ID do ítem para continuar."
    Exit Sub
End If
End Sub

Private Sub cmdInserir_Click()
Sql = "INSERT INTO contatossms (nome, fone) VALUE ('" & ctNome & "', '" & ctFone & "')"

conn.Execute Sql

DataAdocadsms.Refresh
Call sbLimpa_Campos(Me)
ctNome.SetFocus

End Sub

Private Sub ctFone_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctFone.Text = "" Then
        MsgBox "Digite o fone para continuar.", vbCritical
        Exit Sub
    End If
    cmdInserir.SetFocus
End Select
End Sub

Private Sub ctNome_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctNome.Text = "" Then
        MsgBox "Digite o nome para continuar.", vbCritical
        Exit Sub
    End If
    ctFone.SetFocus
End Select
End Sub

Private Sub dbgCadSms_DblClick()
nIDsms = dbgCadSms.Columns(0)
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Call sbViewCadSms

End Sub

Private Sub sbViewCadSms()
With DataAdocadsms
        .ConnectionString = strDns
        .RecordSource = "Select * From contatossms order by id"
End With
DataAdocadsms.Refresh
End Sub
