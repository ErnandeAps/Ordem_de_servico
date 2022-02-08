VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntPeças 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de lançamento de entrada de peças"
   ClientHeight    =   6780
   ClientLeft      =   3060
   ClientTop       =   1950
   ClientWidth     =   10440
   Icon            =   "frmEntPeças.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10440
   Begin Crystal.CrystalReport CRP 
      Left            =   5580
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox ctTtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8970
      MaxLength       =   80
      TabIndex        =   22
      Top             =   6240
      Width           =   1395
   End
   Begin MSAdodcLib.Adodc DataAdo 
      Height          =   375
      Left            =   60
      Top             =   5820
      Width           =   10305
      _ExtentX        =   18177
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
      Caption         =   "Assistente de lançamento de peças"
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
   Begin MSDataGridLib.DataGrid DbGrid 
      Bindings        =   "frmEntPeças.frx":5C12
      Height          =   3855
      Left            =   30
      TabIndex        =   21
      Top             =   1920
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   6800
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "nlan"
         Caption         =   "Lan"
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
         Caption         =   "Ref."
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
         DataField       =   "Descricao"
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
         DataField       =   "valunit"
         Caption         =   "Val Custo"
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
         DataField       =   "qtd"
         Caption         =   "Quant."
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
         DataField       =   "total"
         Caption         =   "Valor"
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
            ColumnWidth     =   510,236
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1560,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4515,024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1379,906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   30
      TabIndex        =   10
      Top             =   990
      Width           =   10365
      Begin VB.TextBox ctTotalUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8070
         MaxLength       =   80
         TabIndex        =   24
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton cmdInserir 
         Height          =   495
         Left            =   9090
         Picture         =   "frmEntPeças.frx":5C28
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   270
         Width           =   585
      End
      Begin VB.CommandButton cmdExcItem 
         Height          =   495
         Left            =   9690
         Picture         =   "frmEntPeças.frx":61F1
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox ctQtd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7530
         MaxLength       =   80
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox ctValUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6570
         MaxLength       =   80
         TabIndex        =   15
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox ctCodProd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5130
         MaxLength       =   80
         TabIndex        =   13
         Top             =   360
         Width           =   1425
      End
      Begin VB.ComboBox cbDescricao 
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
         TabIndex        =   12
         Top             =   360
         Width           =   4995
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   8100
         TabIndex        =   25
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Qtd"
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
         Left            =   7590
         TabIndex        =   18
         Top             =   150
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unit."
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
         Left            =   6600
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Produto"
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
         Left            =   5160
         TabIndex        =   14
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição do ítem"
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
         TabIndex        =   11
         Top             =   120
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Nota Fiscal"
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   10365
      Begin VB.TextBox ctSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2010
         MaxLength       =   80
         TabIndex        =   8
         Top             =   480
         Width           =   1875
      End
      Begin VB.CommandButton cmdDtEntrada 
         Height          =   375
         Left            =   5370
         Picture         =   "frmEntPeças.frx":673C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   465
      End
      Begin VB.ComboBox cbFornecedor 
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
         Left            =   5850
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox ctNFe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         MaxLength       =   80
         TabIndex        =   0
         Top             =   480
         Width           =   1875
      End
      Begin MSMask.MaskEdBox maskDataEnt 
         Height          =   375
         Left            =   3930
         TabIndex        =   5
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Série"
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
         Left            =   2010
         TabIndex        =   9
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NF-e N°"
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
         TabIndex        =   7
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Entrada"
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
         Left            =   3990
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
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
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Total ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8220
      TabIndex        =   23
      Top             =   6300
      Width           =   660
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuCadPecas 
         Caption         =   "Cadastro de Peças"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAbrirNfe 
         Caption         =   "Abrir NF-e Cadastrada"
      End
      Begin VB.Menu mnuAtTotal 
         Caption         =   "Atualiza total"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relatorios"
      Begin VB.Menu mnuRelGlan 
         Caption         =   "Relação Geral de Lançamentos"
      End
      Begin VB.Menu mnuRelLanPeriodo 
         Caption         =   "Relação de lançamentos por NF-e"
      End
   End
End
Attribute VB_Name = "frmEntPeças"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nValunit As Double
Dim nValTotal As Double
Dim nQtd As Integer
Private Sub cbDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyReturn
10005           If cbDescricao.Text = "" Then
10010               MsgBox "Selecione um ítem para continuar.", vbCritical
10015               Exit Sub
10020           End If
    
10025           Set rsTabelas = New ADODB.Recordset
10030           rsTabelas.Open "select * from cadpecas where Descricao='" & cbDescricao & "'", conn
10035           ctCodProd = rsTabelas!ref
10040           ctValUnit = rsTabelas!valcusto
10045           rsTabelas.Close
10050           Set rsTabelas = Nothing
    
10055           ctValUnit.SetFocus
    
10060   End Select
End Sub

Private Sub cbFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyReturn
10005           If cbFornecedor.Text = "" Then
10010               MsgBox "Selecione um fornacedor para continuar.", vbCritical
10015               Exit Sub
10020           End If
10025           cbDescricao.SetFocus
10030   End Select
End Sub

Private Sub cmdDtEntrada_Click()
10000   On Error Resume Next
10005   frmCalendario.Show (1)
10010   maskDataEnt = strData
10015   cbFornecedor.SetFocus
End Sub

Private Sub cmdExcItem_Click()
10000   Dim nLan As Integer
10005   On Error GoTo E:

10010   nLan = InputBox("Digite o N° dolançamento para continuar")

10015   Sql = "DELETE FROM tblanpecas where nlan=" & nLan & " and nfe='" & ctNFe & "'"
10020   conn.Execute Sql

10025   Set rsTabelas = New ADODB.Recordset
10030   rsTabelas.Open "select sum(valunit) as nValTotal from tblanpecas where nfe='" & ctNFe & "'", conn
10035   ctTtotal = Format(rsTabelas!nValTotal, "##,##0.00")
10040   rsTabelas.Close
10045   Set rsTabelas = Nothing

10050   DataAdo.Refresh
10055   cbDescricao.SetFocus
10060 E:
10065   If Err.Number = 13 Then
10070       MsgBox "Selecione um lançamento para  continuar.", vbCritical
10075       Exit Sub
10080   End If
End Sub

Private Sub cmdInserir_Click()
10000   Dim nLan As Integer

10005   Set rsTabelas = New ADODB.Recordset
10010   rsTabelas.Open "select  MAX(nlan) as nCountId from tblanpecas where nfe='" & ctNFe & "'", conn
10015   If IsNull(rsTabelas!ncountId) = True Then
10020       nLan = 1
10025   Else
10030       nLan = rsTabelas!ncountId + 1
10035   End If
10040   rsTabelas.Close
10045   Set rsTabelas = Nothing


10050   Sql = "INSERT INTO tblanpecas (nfe, serie, fornecedor, ref, descricao, valunit, qtd, nlan, data, total) VALUES (" & _
            "'" & ctNFe & "','" & ctSerie & "','" & cbFornecedor & "','" & ctCodProd & "','" & cbDescricao & "'," & _
            "" & converte(ctValUnit) & "," & ctQtd & "," & nLan & ",'" & Format$(maskDataEnt, "yyyy-mm-dd hh:mm:ss") & "'," & _
            "" & converte(ctTotalUnit) & ")"

10055   conn.Execute Sql

10060   cbDescricao.Text = ""
10065   ctCodProd.Text = ""
10070   ctValUnit.Text = ""
10075   ctQtd.Text = ""

10080   Set rsTabelas = New ADODB.Recordset
10085   rsTabelas.Open "select sum(total) as nValTotal from tblanpecas where nfe='" & ctNFe & "'", conn
10090   ctTtotal = Format(rsTabelas!nValTotal, "##,##0.00")
10095   rsTabelas.Close
10100   Set rsTabelas = Nothing


10105   DataAdo.Refresh
10110   DataAdo.Recordset.MoveLast
10115   cbDescricao.SetFocus

End Sub

Private Sub ctNFe_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyReturn
10005           If ctNFe.Text = "" Then
10010               MsgBox "Digite o N° da nota fiscal de Entrada para continuar.", vbCritical
10015               Exit Sub
10020           End If
10025           Set rsTabelas = New ADODB.Recordset
10030           rsTabelas.Open "select * from tblanpecas where nfe='" & ctNFe & "'", conn
10035           If Not rsTabelas.EOF Then
10040               MsgBox "Nota fiscal já cadastrada.", vbCritical
10045               ctSerie = rsTabelas!serie
10050               maskDataEnt = rsTabelas!Data
10055               cbFornecedor = rsTabelas!fornecedor
10060               cmdInserir.Enabled = False
10065               cmdExcItem.Enabled = False
10070               Set rsTabelas = New ADODB.Recordset
10075               rsTabelas.Open "select sum(total) as nValTotal from tblanpecas where nfe='" & ctNFe & "'", conn
10080               ctTtotal = Format(rsTabelas!nValTotal, "##,##0.00")
10085               rsTabelas.Close
10090               Set rsTabelas = Nothing
        
10095           End If
    
10100           With DataAdo
10105               .ConnectionString = strDns
10110               .RecordSource = "Select * From tblanpecas where nfe='" & ctNFe & "' order by id ASC"
10115           End With
10120           DataAdo.Refresh
    
10125           ctSerie.SetFocus
10130   End Select
End Sub

Private Sub ctQtd_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyReturn
10005           If ctQtd.Text = "" Then
10010               MsgBox "Digite a quantidade para continuar.", vbCritical
10015               Exit Sub
10020           End If
10025           nValunit = ctValUnit
10030           nQtd = ctQtd
10035           ctTotalUnit = Format(nValunit * nQtd, "##,##0.00")

10040           cmdInserir.SetFocus
10045   End Select
End Sub

Private Sub ctSerie_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyReturn
10005           If ctSerie.Text = "" Then
10010               MsgBox "Digite o n° de Série da Nota Fiscal de Entrada para continuar.", vbCritical
10015               Exit Sub
10020           End If
10025           maskDataEnt.SetFocus
10030   End Select
End Sub

Private Sub ctValUnit_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyReturn
10005           If ctValUnit.Text = "" Then
10010               MsgBox "Digite o valor unitario para continuar.", vbCritical
10015               Exit Sub
10020           End If
10025           Dim nValunit As Currency
10030           nValunit = ctValUnit
                '********************************************Atualiza val unit***********************
10035           Set rsTabelas = New ADODB.Recordset
10040           rsTabelas.Open "select * from cadpecas where ref='" & ctCodProd & "'", conn
10045           If Not rsTabelas.EOF Then
10050               If nValunit <> rsTabelas!valcusto Then
10055                   Ret = MsgBox("O valor de compra cadastrado é ( " & rsTabelas!valunit & " ), diverge do valor digitado. Deseja atualizar o cadastro de peças ???", vbYesNo)
10060                   If Ret = 6 Then
10065                       Sql = "UPDATE cadpecas SET valcusto=" & converte(ctValUnit) & " where ref='" & ctCodProd & "'"
10070                       conn.Execute Sql
10075                   End If
10080               End If
10085           End If
    
10090           ctQtd.SetFocus
    
10095   End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyEscape
10005           Unload Me
10010   End Select
End Sub

Private Sub Form_Load()
10000   KeyPreview = True
10005   Call sbLoadFornecedores
10010   Call sbLoadpecas

End Sub
Private Sub sbLoadpecas()
10000   Set rsTabelas = New ADODB.Recordset
10005   rsTabelas.Open "select * from  cadpecas order by Descricao", conn

10010   cbDescricao.Clear

10015   Do While Not rsTabelas.EOF
10020       cbDescricao.AddItem rsTabelas!descricao
10025       rsTabelas.MoveNext
10030   Loop

10035   rsTabelas.Close
10040   Set rsTabelas = Nothing

End Sub

Private Sub sbLoadFornecedores()
10000   Set rsTabelas = New ADODB.Recordset
10005   rsTabelas.Open "select * from clientes where titulo='" & "FORNECEDOR" & "' order by nome", conn

10010   cbFornecedor.Clear

10015   Do While Not rsTabelas.EOF
10020       cbFornecedor.AddItem rsTabelas!nome
10025       rsTabelas.MoveNext
10030   Loop

10035   rsTabelas.Close
10040   Set rsTabelas = Nothing

End Sub

Private Sub maskDataEnt_KeyDown(KeyCode As Integer, Shift As Integer)
10000   Select Case KeyCode
            Case vbKeyReturn
10005           If maskDataEnt.Text = "__/__/____" Then
10010               MsgBox "Digite a data de Entrada para continuar.", vbCritical
10015               Exit Sub
10020           End If
10025           cbFornecedor.SetFocus
10030   End Select

End Sub
Private Sub sbTotNota()
10000   Dim nValServ As Double
10005   Dim nValPecas As Double

10010   Set rsTabelas = New ADODB.Recordset

10015   rsTabelas.Open "select SUM(valunit) as nValTotal from tblanpecas where nfe='" & ctNFe & "'", conn
10020   If IsNull(rsTabelas!nValTotal) Then
10025       nValServ = 0
10030   Else
10035       nValServ = rsTabelas!nValTotal
10040   End If
10045   rsTabelas.Close
10050   Set rsTabelas = Nothing
End Sub

Private Sub mnuAbrirNfe_Click()
10000   cmdInserir.Enabled = True
10005   cmdExcItem.Enabled = True
End Sub

Private Sub mnuAtTotal_Click()
10000   Dim nValunit As Double
10005   Dim nValTotal As Double
10010   Dim nQtd As Integer

10015   Set rsTabelas = New ADODB.Recordset
10020   rsTabelas.Open "select * from tblanpecas", conn
10025   Do While Not rsTabelas.EOF
10030       nValunit = rsTabelas!valunit
10035       nQtd = rsTabelas!qtd
10040       nValTotal = nValunit * nQtd
10045       Sql = "UPDATE tblanpecas SET total=" & converte(nValTotal) & " where id=" & rsTabelas!ID & ""
10050       conn.Execute Sql
10055       rsTabelas.MoveNext
10060   Loop
10065   rsTabelas.Close
10070   Set rsTabelas = Nothing
End Sub

Private Sub mnuCadPecas_Click()
10000   Call AtivaForm(frmcadPecas)
10005   Call sbLoadpecas

End Sub

Private Sub mnuRelGlan_Click()
10000   crp.Connect = strDns
10005   crp.ReportFileName = App.Path & "\..\relatorios\RelGLanPecas.rpt"
        'CRP.SelectionFormula = "{tblanpecas.nfe}='" & ctNFe & "'"
10010   crp.Destination = crptToWindow
10015   crp.WindowState = crptMaximized
10020   crp.Action = 1
End Sub

Private Sub mnuRelLanPeriodo_Click()
10000   crp.Connect = strDns
10005   crp.ReportFileName = App.Path & "\..\relatorios\LanEntPecas.rpt"
10010   crp.SelectionFormula = "{tblanpecas.nfe}='" & ctNFe & "'"
10015   crp.Destination = crptToWindow
10020   crp.WindowState = crptMaximized
10025   crp.Action = 1
End Sub

Private Sub mnuSair_Click()
10000   Unload Me
End Sub
