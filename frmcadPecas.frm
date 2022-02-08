VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcadPecas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Peças"
   ClientHeight    =   5025
   ClientLeft      =   4935
   ClientTop       =   3150
   ClientWidth     =   9315
   ClipControls    =   0   'False
   Icon            =   "frmcadPecas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9315
   Begin VB.TextBox ctserie 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      MaxLength       =   60
      TabIndex        =   14
      Top             =   930
      Width           =   3255
   End
   Begin VB.CommandButton cmdAtualiza 
      Height          =   435
      Left            =   8250
      Picture         =   "frmcadPecas.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   900
      Width           =   525
   End
   Begin VB.TextBox ctqtd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6540
      MaxLength       =   25
      TabIndex        =   11
      Top             =   930
      Width           =   1095
   End
   Begin VB.TextBox ctRef 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5220
      MaxLength       =   30
      TabIndex        =   9
      Top             =   270
      Width           =   2655
   End
   Begin VB.TextBox ctvalcusto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3330
      MaxLength       =   25
      TabIndex        =   7
      Top             =   930
      Width           =   1635
   End
   Begin VB.TextBox ctvalvenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4980
      MaxLength       =   25
      TabIndex        =   5
      Top             =   930
      Width           =   1545
   End
   Begin VB.TextBox ctdescricao 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   90
      MaxLength       =   80
      TabIndex        =   0
      Top             =   270
      Width           =   5115
   End
   Begin VB.CommandButton cmdInsert 
      Height          =   435
      Left            =   7710
      Picture         =   "frmcadPecas.frx":61DB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   900
      Width           =   525
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   435
      Left            =   8790
      Picture         =   "frmcadPecas.frx":67A4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   900
      Width           =   495
   End
   Begin MSAdodcLib.Adodc DataAdo 
      Height          =   435
      Left            =   90
      Top             =   4500
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   767
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Peças"
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
   Begin MSDataGridLib.DataGrid dbg 
      Bindings        =   "frmcadPecas.frx":6CEF
      Height          =   3105
      Left            =   90
      TabIndex        =   1
      Top             =   1350
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   5477
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
         DataField       =   "ref"
         Caption         =   "Referencia"
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
         DataField       =   "descricao"
         Caption         =   "Descrição"
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
         DataField       =   "valcusto"
         Caption         =   "Val. Custo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "valunit"
         Caption         =   "Val. Venda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   585,071
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3704,882
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1289,764
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "NCM"
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
      Left            =   90
      TabIndex        =   15
      Top             =   690
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Qtd Estoque"
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
      Left            =   6570
      TabIndex        =   12
      Top             =   690
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ref.:"
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
      Left            =   5250
      TabIndex        =   10
      Top             =   30
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Custo"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   660
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Venda"
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
      Left            =   5010
      TabIndex        =   6
      Top             =   690
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
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
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   795
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuEntPecas 
         Caption         =   "Entrada de peças"
      End
      Begin VB.Menu mnuSerie 
         Caption         =   "N° de Série"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "sair"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuRelPecas 
         Caption         =   "Relação de peças"
      End
   End
End
Attribute VB_Name = "frmcadPecas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nIdPacas As Variant

Private Sub cmdAtualiza_Click()
Sql = "UPDATE cadpecas SET descricao='" & ctdescricao & "', ref='" & ctRef & "', valcusto=" & converte(ctvalcusto) & ", " & _
"valunit=" & converte(ctvalvenda) & ",serie='" & ctSerie & "' where id=" & nIdPacas & ""

conn.Execute Sql

Call sbLimpa_Campos(Me)
nIdPacas = 0
DataAdo.Refresh
ctdescricao.SetFocus

End Sub

Private Sub cmdExcluir_Click()
Dim nIdac As Integer
On Error GoTo E:

nIdac = InputBox("Digite o ID do ítem para continuar.")

Sql = "DELETE from cadpecas where id=" & nIdac & ""

conn.Execute Sql

DataAdo.Refresh

E:
If Err.Number = 13 Then
    MsgBox "Digite o ID do ítem para continuar."
    Exit Sub
End If
End Sub

Private Sub cmdInsert_Click()
Sql = "INSERT INTO cadpecas (descricao, valcusto, valunit, ref,serie) VALUE ('" & ctdescricao & "'," & converte(ctvalcusto) & ", " & converte(ctvalvenda) & "" & _
",'" & ctRef & "','" & ctSerie & "')"

conn.Execute Sql

DataAdo.Refresh
ctdescricao.Text = ""
ctvalcusto.Text = ""
ctvalvenda.Text = ""
ctRef.Text = ""
DataAdo.Refresh
DataAdo.Recordset.MoveLast
ctdescricao.SetFocus

End Sub

Private Sub ctdescricao_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctdescricao.Text = "" Then
        MsgBox "Digite a descrição para continuar", vbCritical
        Exit Sub
    End If
    
    ctRef.SetFocus
End Select
End Sub

Private Sub ctvalor_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If CTVALOR.Text = "" Then
        MsgBox "Digite o valor para continuar.", vbCritical
        Exit Sub
    End If
    cmdInsert.SetFocus
End Select
End Sub




Private Sub ctRef_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctRef.Text = "" Then
        MsgBox "Digite a referencia para continuar.", vbCritical
        Exit Sub
    End If
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadpecas where ref='" & ctRef & "'", conn
    If Not rsTabelas.EOF Then
        MsgBox "Peça já cadastrada.", vbCritical
        Exit Sub
    End If
    rsTabelas.Close
    Set rsTabelas = Nothing
    
    ctSerie.SetFocus
    
Case vbKeyLeft
    ctdescricao.Text = ""
    ctRef.Text = ""
    ctdescricao.SetFocus
    
End Select
End Sub

Private Sub ctSerie_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
'    If ctserie.Text = "" Then
'        MsgBox "Digite o N° de Série para continuar", vbCritical
'        Exit Sub
'    End If
    
    ctvalcusto.SetFocus
End Select
End Sub

Private Sub ctvalcusto_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctvalcusto.Text = "" Then
        MsgBox "Digite o valor de custo para continuar.", vbCritical
        Exit Sub
    End If
    ctvalvenda.SetFocus
End Select
End Sub

Private Sub ctvalvenda_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctvalvenda.Text = "" Then
        MsgBox "Digite o valor de venda para continuar.", vbCritical
        Exit Sub
    End If
    cmdInsert.SetFocus
End Select
End Sub

Private Sub dbg_DblClick()
nIdPacas = 0
nIdPacas = dbg.Columns(0)
If nIdPacas = 0 Then
    MsgBox "Selecione um componente para continuar.", vbCritical
    Exit Sub
End If
Call sbLimpa_Campos(Me)

On Error Resume Next
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from cadpecas where id=" & nIdPacas & "", conn
ctdescricao = rsTabelas!descricao
ctRef = rsTabelas!ref
ctvalcusto = rsTabelas!valcusto
ctvalvenda = rsTabelas!valunit
ctSerie = rsTabelas!serie
rsTabelas.Close
Set rsTabelas = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True

With DataAdo
        .ConnectionString = strDns
        .RecordSource = "Select * From cadpecas order by id ASC"
   End With
    DataAdo.Refresh
End Sub

Private Sub mnuSerie_Click()
frmtbNserie.Show (1)
End Sub
