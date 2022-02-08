VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmtbNserie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de cadastro de N� S�rie"
   ClientHeight    =   3615
   ClientLeft      =   5940
   ClientTop       =   4605
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8505
   Begin MSDataGridLib.DataGrid dbg 
      Bindings        =   "frmtbNserie.frx":0000
      Height          =   1935
      Left            =   60
      TabIndex        =   6
      Top             =   1260
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3413
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
      ColumnCount     =   5
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
         DataField       =   "descricao"
         Caption         =   "Descri��o"
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
         DataField       =   "nserie"
         Caption         =   "N� S�rie"
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
      BeginProperty Column03 
         DataField       =   "status"
         Caption         =   "Status"
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
      BeginProperty Column04 
         DataField       =   "nos"
         Caption         =   "Destino OS"
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
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3240
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1874,835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1094,74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DataAdo 
      Height          =   375
      Left            =   60
      Top             =   3210
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   661
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
      Caption         =   "Rela��o de N� de s�rie por item"
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
   Begin VB.CommandButton cmdInsertPecas 
      Height          =   435
      Left            =   2520
      Picture         =   "frmtbNserie.frx":0016
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   810
      Width           =   525
   End
   Begin VB.CommandButton cmdExcPecas 
      Height          =   435
      Left            =   3090
      Picture         =   "frmtbNserie.frx":05DF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   810
      Width           =   495
   End
   Begin VB.TextBox ctSerie 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      MaxLength       =   8
      TabIndex        =   2
      Top             =   840
      Width           =   2385
   End
   Begin VB.ComboBox cbPecas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   5115
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade"
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
      TabIndex        =   3
      Top             =   630
      Width           =   930
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Descri��o"
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
      TabIndex        =   1
      Top             =   30
      Width           =   795
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relat�rios"
      Begin VB.Menu mnuRelRelSeriePecas 
         Caption         =   "Rela��o de N� de S�rie por pe�as"
      End
   End
End
Attribute VB_Name = "frmtbNserie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nID As Variant
Dim nIDpeca As Variant

Private Sub cbPecas_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    
    If cbPecas.Text = "" Then
        MsgBox "Selecione uma �tem para continuar.", vbCritical
        Exit Sub
    End If
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadpecas where descricao='" & cbPecas & "'", conn
    nIDpeca = rsTabelas!ID
    rsTabelas.Close
    Set rsTabelas = Nothing
    
    With DataAdo
        .ConnectionString = strDns
        .RecordSource = "Select * From tbnserie where idpeca='" & nIDpeca & "' order by id ASC"
   End With
    DataAdo.Refresh
    ctSerie.SetFocus
End Select
End Sub

Private Sub cmdExcPecas_Click()
On Error GoTo E:
nID = InputBox("Digite o id do �tem para continuar.")

Sql = "DELETE FROM tbnserie where id=" & nID & ""
conn.Execute Sql
DataAdo.Refresh

E:
If Err.Number = 13 Then
    MsgBox "�tem inv�lido.", vbCritical
    Exit Sub
End If
End Sub

Private Sub cmdInsertPecas_Click()
Sql = "INSERT INTO tbnserie (idpeca,descricao,nserie,status) VALUES (" & nIDpeca & ",'" & cbPecas & "'," & ctSerie & ", " & _
"'" & "LIVRE" & "')"
conn.Execute Sql
ctSerie.Text = ""
ctSerie.SetFocus

DataAdo.Refresh
DataAdo.Recordset.MoveLast
End Sub

Private Sub ctSerie_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctSerie.Text = "" Then
        MsgBox "Digite uma registro para continuar.", vbCritical
        Exit Sub
    End If
    
    cmdInsertPecas.SetFocus
    
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me

End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from cadpecas order by id ASC", conn
Do While Not rsTabelas.EOF
    cbPecas.AddItem rsTabelas!descricao
    rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing

End Sub
