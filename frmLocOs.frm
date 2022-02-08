VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLocOs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de consulta"
   ClientHeight    =   5520
   ClientLeft      =   1320
   ClientTop       =   1665
   ClientWidth     =   14205
   Icon            =   "frmLocOs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   14205
   Begin MSDataGridLib.DataGrid dgOs 
      Bindings        =   "frmLocOs.frx":1CFA
      Height          =   3705
      Left            =   30
      TabIndex        =   6
      Top             =   1350
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   6535
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "N° OS"
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
         DataField       =   "modelo"
         Caption         =   "Equip"
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
         Caption         =   "Num.Série"
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
         DataField       =   "Data"
         Caption         =   "Data"
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
         DataField       =   "servico"
         Caption         =   "Serviço"
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
      BeginProperty Column05 
         DataField       =   "Empresa"
         Caption         =   "Cliente"
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
      BeginProperty Column06 
         DataField       =   "endereco"
         Caption         =   "Endereço"
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
      BeginProperty Column07 
         DataField       =   "bairro"
         Caption         =   "Bairro"
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
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1349,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1904,882
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2310,236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3509,858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3780,284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2970,142
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoData 
      Height          =   390
      Left            =   30
      Top             =   5070
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   688
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
      DataSourceName  =   "SuportekLocal"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ordem de serviço"
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
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   30
      TabIndex        =   5
      Top             =   660
      Width           =   14115
      Begin VB.ComboBox cbServSolic 
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
         Left            =   90
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   4035
      End
      Begin MSMask.MaskEdBox maskCriterio 
         Height          =   375
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   645
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   14115
      Begin VB.OptionButton OptServico 
         Caption         =   "Serviço"
         Height          =   285
         Left            =   3300
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optImportadas 
         Caption         =   "Importadas"
         Height          =   285
         Left            =   4590
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSair 
         Height          =   435
         Left            =   6360
         Picture         =   "frmLocOs.frx":1D10
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Filtrar dados"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdloc 
         Height          =   435
         Left            =   5850
         Picture         =   "frmLocOs.frx":225B
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Filtrar dados"
         Top             =   150
         Width           =   495
      End
      Begin VB.OptionButton optdata 
         Caption         =   "Data"
         Height          =   285
         Left            =   2370
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
      Begin VB.OptionButton optcliente 
         Caption         =   "Cliente"
         Height          =   285
         Left            =   270
         TabIndex        =   3
         Top             =   240
         Width           =   825
      End
      Begin VB.OptionButton optos 
         Caption         =   "N° da Os"
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   210
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmLocOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdloc_Click()
If optos.Value = True Then
    If maskCriterio.Text = "" Then
        AdoData.RecordSource = "select * from osdb order by id DESC"
    Else
        AdoData.RecordSource = "select * from osdb where id='" & maskCriterio & "' order by id DESC"
    End If
End If

If optcliente.Value = True Then
    If maskCriterio.Text = "" Then
        AdoData.RecordSource = "select * from osdb order by id DESC"
    Else
        AdoData.RecordSource = "select * from osdb where empresa like '" & maskCriterio & "%' order by id DESC"
    End If
End If

If optdata.Value = True Then
    If maskCriterio.Text = "" Then
        AdoData.RecordSource = "select * from osdb order by id DESC"
    Else
        AdoData.RecordSource = "select * from osdb where data='" & Format(maskCriterio, "yyyy-mm-dd") & "' order by id DESC"
    End If
End If

If optImportadas.Value = True Then
    
   AdoData.RecordSource = "select * from osdb where situacao='" & "Solicitacao em analise" & "' order by id DESC"
   
End If

If OptServico.Value = True Then
    maskCriterio = cbServSolic
   AdoData.RecordSource = "select * from osdb where servico='" & maskCriterio & "' order by id DESC"
   
End If





AdoData.Refresh

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dgOs_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
Dim nNumOs As Variant
Dim nIdParceiro As Integer

nNumOs = dgOs.Columns(0)
On Error Resume Next

Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select * from osdb where id=" & nNumOs & "", conn
frmOs.ctNos = rsTabelas!ID
frmOs.ctcodClient = rsTabelas!idCliente
frmOs.ctrazao = rsTabelas!empresa
frmOs.ctend = rsTabelas!endereco
frmOs.ctbairro = rsTabelas!bairro
frmOs.ctcidade = rsTabelas!cidade
frmOs.ctuf = rsTabelas!estado
frmOs.MASKCEP = rsTabelas!CEP
frmOs.ctcgc = rsTabelas!CNPJ
frmOs.ctinsc = rsTabelas!insc
frmOs.MASKFONE = rsTabelas!Telefone
frmOs.MASKCEL = rsTabelas!celular
frmOs.ctEmail = rsTabelas!email
frmOs.ctroteiro = rsTabelas!roteiro
frmOs.ctcontato = rsTabelas!responsavel
frmOs.ctEquipamento = rsTabelas!equipamento
frmOs.ctNSerie = rsTabelas!nserie
frmOs.ctmarca = rsTabelas!marca
frmOs.ctmodelo = rsTabelas!modelo
frmOs.ctRef = rsTabelas!ref
frmOs.ctTombo = rsTabelas!tombo
frmOs.ctDefeito = rsTabelas!defeito
frmOs.ctDiagnostico = rsTabelas!diagnostico
frmOs.cbServSolic = rsTabelas!servico
frmOs.Cbsituacao = rsTabelas!situacao
frmOs.cbcolaborador = rsTabelas!colaborador
frmOs.maskDataEnt = rsTabelas!Data
frmOs.maskDataPrev = rsTabelas!datasaida
nStatus = rsTabelas!Status
frmOs.ctTotalPago = Format(rsTabelas!totalpago, "##,##0.00")
frmOs.ctDesconto = Format(rsTabelas!desconto, "##,##0.00")
frmOs.cbFormaPag = rsTabelas!formapg
frmOs.ctIdWeb = rsTabelas!idWeb
frmOs.ctObs = rsTabelas!Obs
nIdParceiro = rsTabelas!idcolaborador

rsTabelas.Close
Set rsTabelas = Nothing

Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from clientes where id=" & nIdParceiro & " and titulo='" & "PARCEIRO" & "'", conn
frmOs.cbcolaborador = rsTabelas!nome
rsTabelas.Close
Set rsTabelas = Nothing

If frmOs.ctcodClient.Text = "" Then
    frmOs.cmdImportarCad.Enabled = True
End If

Call fStatusOs(nNumOs)

If nStatus = 3 Then
    frmOs.cmdReceber.Enabled = False
Else
    frmOs.cmdReceber.Enabled = True
End If

Unload Me
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
optcliente.Value = True
maskCriterio.Width = 5900
Call sbAddServico

With AdoData
        .ConnectionString = strDns
        .RecordSource = "Select * From osdb order by id DESC"
   End With
    AdoData.Refresh
End Sub

Private Sub maskCriterio_Change()
Set rsTabelas = New ADODB.Recordset
 
    
End Sub

Private Sub optcliente_Click()
If maskCriterio.Visible = False Then
    maskCriterio.Visible = True
    cbServSolic.Visible = False
End If

maskCriterio.Width = 5900
maskCriterio.Mask = "": maskCriterio.Text = ""
'maskCriterio.SetFocus
End Sub

Private Sub optdata_Click()
If maskCriterio.Visible = False Then
    maskCriterio.Visible = True
    cbServSolic.Visible = False
End If
maskCriterio.Width = 1300
maskCriterio.Text = ""
maskCriterio.Mask = "##/##/####"
'maskCriterio.SetFocus
End Sub

Private Sub optImportadas_Click()
If maskCriterio.Visible = False Then
    maskCriterio.Visible = True
    cbServSolic.Visible = False
End If
maskCriterio.Width = 1300
maskCriterio.Text = ""
maskCriterio.Mask = "##/##/####"
End Sub

Private Sub optos_Click()
If maskCriterio.Visible = False Then
    maskCriterio.Visible = True
    cbServSolic.Visible = False
End If
maskCriterio.Width = 1250
maskCriterio.Mask = "": maskCriterio.Text = ""
'maskCriterio.SetFocus
End Sub


Private Sub OptServico_Click()
maskCriterio.Width = 5900
maskCriterio.Mask = "": maskCriterio.Text = ""
maskCriterio.Visible = False
cbServSolic.Visible = True

End Sub
Private Sub cbServSolic_GotFocus()
   On Error Resume Next
   cbServSolic.SelStart = 0
   cbServSolic.SelLength = Len(cbServSolic.Text)
End Sub

Private Sub sbAddServico()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadservicos order by Descricao", conn
    cbServSolic.Clear
    Do While Not rsTabelas.EOF
        cbServSolic.AddItem rsTabelas!descricao
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
End Sub
