VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContasApagar 
   Caption         =   "Assistente de contas a pagar"
   ClientHeight    =   6195
   ClientLeft      =   4170
   ClientTop       =   3450
   ClientWidth     =   11745
   Icon            =   "frmContasApagar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   11745
   Begin MSDataGridLib.DataGrid DBpagamentos 
      Bindings        =   "frmContasApagar.frx":26E2
      Height          =   3405
      Left            =   30
      TabIndex        =   25
      Top             =   1680
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   6006
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
         DataField       =   "datavenc"
         Caption         =   "Vencimento"
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
         DataField       =   "status"
         Caption         =   "Situação"
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
         DataField       =   "grupo"
         Caption         =   "Grupo"
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
         DataField       =   "subgrupo"
         Caption         =   "Conta"
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
         DataField       =   "numParc"
         Caption         =   "Parcelas"
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
         DataField       =   "historico"
         Caption         =   "Histórico"
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
         DataField       =   "valor"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1874,835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2670,236
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1110,047
         EndProperty
      EndProperty
   End
   Begin VB.Frame framConsultar 
      Caption         =   "Período"
      Height          =   1665
      Left            =   30
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   11685
      Begin MSMask.MaskEdBox maskdataini 
         Height          =   360
         Left            =   90
         TabIndex        =   20
         Top             =   450
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   -2147483628
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox maskdatafin 
         Height          =   360
         Left            =   1500
         TabIndex        =   22
         Top             =   450
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   -2147483628
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Data final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1500
         TabIndex        =   23
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   990
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1665
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   11685
      Begin VB.TextBox cthist 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   1050
         Width           =   7905
      End
      Begin VB.TextBox ctvalor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   10
         Top             =   1050
         Width           =   1575
      End
      Begin VB.ComboBox cbGrupo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmContasApagar.frx":26F8
         Left            =   1380
         List            =   "frmContasApagar.frx":26FA
         TabIndex        =   9
         Top             =   390
         Width           =   3825
      End
      Begin VB.ComboBox cbSubGrupo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmContasApagar.frx":26FC
         Left            =   5220
         List            =   "frmContasApagar.frx":26FE
         TabIndex        =   8
         Top             =   390
         Width           =   4935
      End
      Begin VB.TextBox ctParcela 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9660
         TabIndex        =   7
         Top             =   1050
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSMask.MaskEdBox MaskdataVenc 
         Height          =   360
         Left            =   90
         TabIndex        =   12
         Top             =   390
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   -2147483628
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   90
         TabIndex        =   18
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   810
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8010
         TabIndex        =   16
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1380
         TabIndex        =   15
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5220
         TabIndex        =   14
         Top             =   150
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   9660
         TabIndex        =   13
         Top             =   810
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   5100
      Width           =   11685
      Begin MSAdodcLib.Adodc datapag 
         Height          =   375
         Left            =   9630
         Top             =   420
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "Adodc1"
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
      Begin VB.CommandButton cmdEstornaTitulosBaixados 
         Caption         =   "Estorna Tit. Baixados"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8130
         Picture         =   "frmContasApagar.frx":2700
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton BtExcluir 
         Caption         =   "&Baixar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6990
         Picture         =   "frmContasApagar.frx":2A0A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5850
         Picture         =   "frmContasApagar.frx":32D4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exclui o Cliente do Banco de Dados"
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton BtAlterar 
         Caption         =   "Al&terar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4710
         Picture         =   "frmContasApagar.frx":4719
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Altera os dados do Cliente"
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton BtConsultar 
         Caption         =   "&Consultar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3570
         Picture         =   "frmContasApagar.frx":58A1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton BtAdicionar 
         Caption         =   "Novo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         Picture         =   "frmContasApagar.frx":759B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1155
      End
      Begin Crystal.CrystalReport crp 
         Left            =   540
         Top             =   420
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuConsultaTitulosBaixadaos 
         Caption         =   "Consultar titulos baixados"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuRelGeraldeTitulos 
         Caption         =   "Relação geral de titulos"
      End
      Begin VB.Menu mnuRelTitulosAgrupados 
         Caption         =   "Relação de titulos agrupados"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReldeTitulosemDetalhe 
         Caption         =   "Relação de titulos em detalhe"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRelTitulosBaixados 
         Caption         =   "Relação de titulos em detalhe baixados"
      End
   End
End
Attribute VB_Name = "frmContasApagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ret As Integer
Dim mLan As Integer
Dim strtb As String

Private Sub BtAdicionar_Click()
If BtAdicionar.Caption = "Novo" Then
   'strtb = "tbcontasapagartmp"
    With datapag
        .ConnectionString = strDns
        .RecordSource = "Select * from tbcontasapagartmp ORDER BY id DESC"
    End With
    datapag.Refresh
    
    'dataContasApagar.RecordSource = "Select * from " & strtb
    'dataContasApagar.Refresh
    BtAdicionar.Caption = "Salvar"
    MaskdataVenc.SetFocus
    Exit Sub
End If

Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "Select * from tbcontasapagartmp order by Lan ASC", conn
'Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagartmp order by tbCtpgLan")
Do While Not rsTabelas.EOF
    Sql = "INSERT INTO tbcontasapagar (Op,datalan,grupo,subgrupo,historico,datavenc,valor,status,lan,numParc) VALUES ('" & MOP & "', " & _
    "'" & Format$(rsTabelas!datalan, "yyyy-mm-dd hh:mm:ss") & "', '" & rsTabelas!grupo & "', '" & rsTabelas!SubGrupo & "', '" & rsTabelas!Historico & "', " & _
    "'" & Format$(rsTabelas!DataVenc, "yyyy-mm-dd hh:mm:ss") & "'," & converte(rsTabelas!Valor) & ", '" & rsTabelas!Status & "'," & rsTabelas!lan & ", " & _
    "'" & rsTabelas!numParc & "')"
    conn.Execute Sql
        
    rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing

Sql = "DELETE from tbcontasapagartmp"
conn.Execute Sql

Unload Me
Call AtivaForm(frmContasApagar)
'frmContasApagar.Show
End Sub

Private Sub BtAlterar_Click()
If mLan = 0 Then
    MsgBox "Selecione um lançamento p/ continuar", vbCritical, "Assistente de verificação"
    Exit Sub
End If

Sql = "UPDATE tbcontasapagar SET op='" & MOP & "', datalan='" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "', grupo='" & cbGrupo & "', " & _
"subgrupo='" & cbSubGrupo & "', historico='" & cthist & "', datavenc='" & Format(MaskdataVenc, "yyyy-mm-dd hh:mm:ss") & "', " & _
"valor=" & converte(ctvalor) & " where id=" & mLan & ""
conn.Execute Sql

'Set Reg_Dados = BANCO.OpenRecordset("Select * from " & strtb & " where tbCtpgLan=" & mLan & "")
'Reg_Dados.Edit
'Reg_Dados!tbCtpgOp = 0
'Reg_Dados!tbCtpgDataLan = Date
'Reg_Dados!tbCtpgGrupo = cbGrupo
'Reg_Dados!tbCtpgSubGrupo = cbSubGrupo
'Reg_Dados!tbCtpgHistorico = cthist
'Reg_Dados!tbCtpgDataVenc = MaskdataVenc
'Reg_Dados!tbCtpgValor = CTVALOR
'Reg_Dados.Update
datapag.Refresh

Call sbLimpa_Campos(Me)
mLan = 0
BtExcluir.Enabled = False
cmdEstornaTitulosBaixados.Enabled = False
BtAlterar.Enabled = False
cmdExcluir.Enabled = False
End Sub

Private Sub BtConsultar_Click()
BtAdicionar.Enabled = False
framConsultar.Visible = True
maskdataini.SetFocus
End Sub

Private Sub BtExcluir_Click()
'On Error Resume Next

If mLan = 0 Then
    MsgBox "Selecione um lançamento p/ continuar", vbCritical, "Assistente de verificação"
    Exit Sub
End If

Ret = MsgBox("Tem certesa que deseja baixar este lançamento ???", vbYesNo, "Assistente de verificação")
If Ret <> 6 Then
    Exit Sub
End If

Sql = "UPDATE tbcontasapagar SET op='" & MOP & "', databaixa='" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "', status='" & "Pago" & "' " & _
"where id=" & mLan & ""

conn.Execute Sql

'**********************************Lança no caixa diário**************************************
Sql = "INSERT INTO caixa (op,data,doc, tipo, descricao, debito, credito,saldo,num) VALUES ('" & MOP & "', '" & Format$(Date, "yyyy-mm-dd hh:mm:ss") & "', " & _
    "'" & "2.0.01" & "','" & "DÉBITO" & "', '" & cthist & "', " & converte(ctvalor) & ", " & converte(0) & ", " & _
    "" & converte(0) & "," & mLan & ")"
    
conn.Execute Sql

datapag.Refresh
'***********************************************************************************************
Call sbLimpa_Campos(Me)
mLan = 0
BtExcluir.Enabled = False
cmdEstornaTitulosBaixados.Enabled = False
BtAlterar.Enabled = False
cmdExcluir.Enabled = False
End Sub

Private Sub cbGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If cbGrupo.Text = "" Then
        MsgBox "Digite ou selecione um grupo para continuar.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    'Set Reg_Dados = BANCO.OpenRecordset("Select * from tbgrupo where tbgrpDesc='" & cbGrupo & "'")
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "Select * from tbgrupo where tbgrpDesc='" & cbGrupo & "'", conn
    If rsTabelas.EOF Then
        rsTabelas.Close
        Sql = "INSERT INTO tbgrupo (tbgrpdesc) VALUES ('" & cbGrupo & "')"
        conn.Execute Sql
        
        cbGrupo.Clear
        rsTabelas.Open "Select * from tbgrupo order by tbgrpDesc ASC", conn
        'Set Reg_Dados2 = BANCO.OpenRecordset("Select * from tbgrupo order by tbgrpDesc")
        Do While Not rsTabelas.EOF
            cbGrupo.AddItem rsTabelas!tbgrpdesc
            rsTabelas.MoveNext
        Loop
        rsTabelas.Close
        Set rsTabelas = Nothing
    Else
         rsTabelas.Close
        Set rsTabelas = Nothing
    End If

    cbSubGrupo.Clear
    
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "Select * from tbsgrupo where tbgrpdesc='" & cbGrupo.Text & "'", conn
    
    'Set Reg_Dados = BANCO.OpenRecordset("Select * from tbsgrupo where tbgrpdesc='" & cbGrupo.Text & "'")
    Do While Not rsTabelas.EOF
        cbSubGrupo.AddItem rsTabelas!tbsgrpDesc
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
        
    cbSubGrupo.SetFocus

Case vbKeyDelete
    If cbGrupo.Text = "Cartão de Crédito" Then
        MsgBox "Este grupo não pode ser excluido do sistema.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    If cbGrupo.Text = "Cheques" Then
        MsgBox "Este grupo não pode ser excluido do sistema.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    If cbGrupo.Text = "Despesas Fixas" Then
        MsgBox "Este grupo não pode ser excluido do sistema.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    If cbGrupo.Text = "Duplicatas" Then
        MsgBox "Este grupo não pode ser excluido do sistema.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "Select * from tbgrupo where tbgrpDesc='" & cbGrupo & "'", conn
    
    'Set Reg_Dados = BANCO.OpenRecordset("Select * from tbgrupo where tbgrpDesc='" & cbGrupo & "'")
    If Not rsTabelas.EOF Then
        Ret = MsgBox("Tem certeza que deseja excluir este grupo ???", vbYesNo, "Assistente de verificação")
        If Ret = 6 Then
            Sql = "DELETE from tbgrupo where tbgrpDesc='" & cbGrupo & "'"
            conn.Execute Sql
        End If
    End If
    rsTabelas.Close
        
    cbGrupo.Clear
    
    rsTabelas.Open "Select * from tbgrupo order by tbgrpDesc ASC", conn
    'Set Reg_Dados = BANCO.OpenRecordset("Select * from tbgrupo order by tbgrpDesc")
    Do While Not rsTabelas.EOF
        cbGrupo.AddItem rsTabelas!tbgrpdesc
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing

End Select
End Sub

Private Sub cbSubGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If cbSubGrupo.Text = "" Then
        MsgBox "Digite ou selecione um sub-grupo para continuar.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "Select * from tbsgrupo where tbgrpdesc='" & cbGrupo.Text & "' and tbsgrpDesc='" & cbSubGrupo.Text & "'", conn
    
    If rsTabelas.EOF Then
        rsTabelas.Close
        Sql = "INSERT INTO tbsgrupo (tbgrpdesc, tbsgrpdesc) VALUES ('" & cbGrupo & "', '" & cbSubGrupo & "')"
        conn.Execute Sql
    
        cbSubGrupo.Clear
        
        rsTabelas.Open "Select * from tbsgrupo where tbgrpdesc='" & cbGrupo.Text & "' order by tbsgrpdesc ASC", conn
        
        Do While Not rsTabelas.EOF
            cbSubGrupo.AddItem rsTabelas!tbsgrpDesc
            rsTabelas.MoveNext
        Loop
        rsTabelas.Close
        Set rsTabelas = Nothing
    Else
        rsTabelas.Close
        Set rsTabelas = Nothing
    End If
   
    
    cthist.SetFocus

Case vbKeyDelete
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "Select * from tbsgrupo where tbgrpdesc='" & cbGrupo.Text & "' and tbsgrpDesc='" & cbSubGrupo.Text & "'", conn
    
    If Not rsTabelas.EOF Then
        Ret = MsgBox("Tem certeza que deseja excluir este sub-grupo ???", vbYesNo, "Assistente de verificação")
        If Ret = 6 Then
           Sql = "DELETE from tbsgrupo where tbgrpdesc='" & cbGrupo.Text & "' and tbsgrpDesc='" & cbSubGrupo.Text & "'"
           conn.Execute Sql
        End If
    End If
    rsTabelas.Close
    
    cbGrupo.Clear
    
    rsTabelas.Open "Select * from tbsgrupo where tbgrpdesc='" & cbGrupo.Text & "' order by tbgrpDesc ASC", conn
    Do While Not rsTabelas.EOF
        cbSubGrupo.AddItem rsTabelas!tbsgrpDesc
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing

End Select
End Sub

Private Sub cmdEstornaTitulosBaixados_Click()
If mLan = 0 Then
    MsgBox "Selecione um lançamento p/ continuar", vbCritical, "Assistente de verificação"
    Exit Sub
End If

Ret = MsgBox("Tem certesa que deseja estornar este lançamento ???", vbYesNo, "Assistente de verificação")
If Ret <> 6 Then
    Exit Sub
End If

Sql = "UPDATE tbcontasapagar SET op='" & MOP & "', databaixa='" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "', status='" & "Aberto" & "' " & _
"where id=" & mLan & ""

conn.Execute Sql

Sql = "DELETE from caixa where num=" & mLan & ""
conn.Execute Sql

datapag.Refresh
Call sbLimpa_Campos(Me)
mLan = 0
BtExcluir.Enabled = False
cmdEstornaTitulosBaixados.Enabled = False
BtAlterar.Enabled = False
cmdExcluir.Enabled = False
End Sub

Private Sub cmdExcluir_Click()
If mLan = 0 Then
    MsgBox "Selecione um lançamento p/ continuar", vbCritical, "Assistente de verificação"
    Exit Sub
End If

Ret = MsgBox("Tem certesa que deseja excluir este lançamento ???", vbYesNo, "Assistente de verificação")
If Ret <> 6 Then
    Exit Sub
End If

Sql = "DELETE from " & strtb & " where id=" & mLan & ""
conn.Execute Sql

datapag.Refresh

Call sbLimpa_Campos(Me)
mLan = 0
BtExcluir.Enabled = False
cmdEstornaTitulosBaixados.Enabled = False
BtAlterar.Enabled = False
cmdExcluir.Enabled = False
End Sub

Private Sub cthist_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If cthist.Text = "" Then
        MsgBox "Digite o histórico p/ continuar.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    ctvalor.SetFocus

End Select
End Sub

Private Sub ctParcela_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctvalor.SetFocus
End Select
End Sub

Private Sub ctvalor_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctvalor.Text = "" Then
        MsgBox "Digite o valor p/ continuar.", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    If BtAdicionar.Enabled = False Then
        Exit Sub
    End If
    
    Select Case cbGrupo
    Case "Cartão de Crédito"
        Label4.Visible = True
        ctParcela.Visible = True
        ctParcela.SetFocus
        Call sbaddCartaoCredito
        Exit Sub
        
    Case "Cheques"
        Label4.Visible = True
        ctParcela.Visible = True
        Call addOutros
        Exit Sub
        
    Case "Despesas Fixas"
        Label4.Visible = True
        ctParcela.Visible = True
        Call addDespFixas
        Exit Sub
        
    Case "Duplicatas"
        Label4.Visible = True
        ctParcela.Visible = True
        ctParcela.SetFocus
        Call addDuplicatas
        Exit Sub
        
    End Select
    
    Label4.Visible = True
    ctParcela.Visible = True
    Call addOutros
    
End Select
End Sub

Private Sub DBpagamentos_DblClick()
Call sbLimpa_Campos(Me)
framConsultar.Visible = False

mLan = DBpagamentos.Columns(0)
If mLan = 0 Then
    MsgBox "Selecione um titulo p/ continuar.", vbCritical, "Assistente de verificação"
    Exit Sub
End If

If BtAdicionar.Caption = "Novo" Then
    strtb = "tbcontasapagar"
Else
    strtb = "tbcontasapagartmp"
End If

Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "Select * from " & strtb & " where id=" & mLan & "", conn

'Set Reg_Dados = BANCO.OpenRecordset("Select * from " & strtb & " where tbCtpgLan=" & mLan & "")
cbGrupo = rsTabelas!grupo
cbSubGrupo = rsTabelas!SubGrupo
cthist = rsTabelas!Historico
MaskdataVenc = rsTabelas!DataVenc
ctvalor = Format(rsTabelas!Valor, "##,##0.00")
rsTabelas.Close
Set rsTabelas = Nothing

BtExcluir.Enabled = True
cmdEstornaTitulosBaixados.Enabled = True
BtAlterar.Enabled = True
cmdExcluir.Enabled = True


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    If framConsultar.Visible = True Then
        With datapag
            .ConnectionString = strDns
            .RecordSource = "Select * from tbcontasapagartmp ORDER BY id ASC"
        End With
        datapag.Refresh
    
    End If
    
    If BtAdicionar.Caption = "Salvar" Then
        Sql = "DELETE from tbcontasapagartmp"
        conn.Execute Sql
        
        With datapag
            .ConnectionString = strDns
            .RecordSource = "Select * from tbcontasapagartmp ORDER BY id ASC"
        End With
        datapag.Refresh
        
        Call sbLimpa_Campos(Me)
        
        BtAdicionar.Caption = "Novo"
        Exit Sub
    End If
    
    If BtAdicionar.Enabled = False Then
'        Set Reg_Dados = BANCO.OpenRecordset("select * from tbcontasapagartmp")
'        Do While Not Reg_Dados.EOF
'            Reg_Dados.Delete
'            Reg_Dados.MoveNext
'        Loop
'        Reg_Dados.Close
'        Set Reg_Dados = Nothing
        
        
        
        Call sbLimpa_Campos(Me)
        framConsultar.Visible = False
        BtAdicionar.Enabled = True
        BtAdicionar.Caption = "Novo"
        Exit Sub
    End If
    
    
    Unload Me
    

End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
With datapag
        .ConnectionString = strDns
        .RecordSource = "Select * from tbcontasapagartmp ORDER BY id ASC"
End With
datapag.Refresh

Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "Select * from tbgrupo order by tbgrpDesc", conn
Do While Not rsTabelas.EOF
    cbGrupo.AddItem rsTabelas!tbgrpdesc
    rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing

End Sub

Public Sub addDespFixas()
Dim numParcelas As Integer
Dim numParcelasFim As Integer
numParcelas = 1

strtb = "tbcontasapagartmp"

If ctParcela.Text = "" Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If

If ctParcela.Text = 0 Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If
    
Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagar order by tbCtpgLan")
If Reg_Dados.EOF Then
    mLan = 1
Else
    Reg_Dados.MoveLast
    mLan = Reg_Dados!tbCtpgLan + 1
End If
Reg_Dados.Close
Set Reg_Dados = Nothing

numParcelasFim = ctParcela

Do While numParcelas <= numParcelasFim
    
    Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagartmp order by tbCtpgLan")
    Reg_Dados.AddNew
    Reg_Dados!tbCtpgOp = 0
    Reg_Dados!tbCtpgLan = mLan
    Reg_Dados!tbCtpgDataLan = Date
    Reg_Dados!tbCtpgGrupo = cbGrupo
    Reg_Dados!tbCtpgSubGrupo = cbSubGrupo
    Reg_Dados!tbCtpgHistorico = cthist
    Reg_Dados!tbCtpgDataVenc = MaskdataVenc
    Reg_Dados!tbCtpgValor = ctvalor
    Reg_Dados!tbCtpgStatus = "Aberto"
    Reg_Dados.Update
    
    mLan = mLan + 1
    numParcelas = numParcelas + 1
    MaskdataVenc = DateAdd("m", 1, MaskdataVenc)
Loop

Label4.Visible = False
ctParcela.Visible = False

dataContasApagar.RecordSource = "Select * from tbcontasapagartmp"
dataContasApagar.Refresh
Call sbLimpa_Campos(Me)
mLan = 0

End Sub

Public Sub addCheques()

Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagar order by tbCtpgLan")
If Reg_Dados.EOF Then
    mLan = 1
Else
    Reg_Dados.MoveLast
    mLan = Reg_Dados!tbCtpgLan + 1
End If
Reg_Dados.Close
Set Reg_Dados = Nothing


Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagartmp order by tbCtpgLan")
Reg_Dados.AddNew
Reg_Dados!tbCtpgOp = 0
Reg_Dados!tbCtpgLan = mLan
Reg_Dados!tbCtpgDataLan = Date
Reg_Dados!tbCtpgGrupo = cbGrupo
Reg_Dados!tbCtpgSubGrupo = cbSubGrupo
Reg_Dados!tbCtpgHistorico = cthist
Reg_Dados!tbCtpgDataVenc = MaskdataVenc
Reg_Dados!tbCtpgValor = ctvalor
Reg_Dados!tbCtpgStatus = "Aberto"
Reg_Dados.Update
    
dataContasApagar.RecordSource = "Select * from tbcontasapagartmp"
dataContasApagar.Refresh
Call sbLimpa_Campos(Me)
MaskdataVenc.SetFocus
mLan = 0

End Sub

Private Sub maskdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If maskdatafin.Text = "__/__/____" Then
        MsgBox "Digite a final p/ continuar", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    With datapag
        .ConnectionString = strDns
        '.RecordSource = "select * from tbcontasapagar where DataVenc>='" & Format(maskdataini, "yyyy-mm-dd hh:mm:ss") & "'" & _
    "and DataVenc<='" & Format(maskdatafin, "yyyy-mm-dd hh:mm:ss") & "' and Status='" & "Aberto" & "' order by DataVenc ASC"
    .RecordSource = "select * from tbcontasapagar where DataVenc>='" & Format(maskdataini, "yyyy-mm-dd") & "'" & _
    "and DataVenc<='" & Format(maskdatafin, "yyyy-mm-dd") & "' and Status='" & "Aberto" & "' order by DataVenc ASC"
    End With
    datapag.Refresh

    'Call sbLimpa_Campos(Me)
    'framConsultar.Visible = False
End Select
End Sub

Private Sub maskdataini_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If maskdataini.Text = "__/__/____" Then
        MsgBox "Digite a inicial p/ continuar", vbCritical, "Assistente de verificação"
        Exit Sub
    End If
    
    maskdatafin.SetFocus
End Select
End Sub

Private Sub MaskdataVenc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If MaskdataVenc.Text = "__/__/____" Then
        MsgBox "Digite a data do vencimento p/ continuar"
        Exit Sub
    End If
    
    cbGrupo.SetFocus

End Select
End Sub

Public Sub addDuplicatas()
Dim numParcelas As Integer
Dim numParcelasFim As Integer
numParcelas = 1

strtb = "tbcontasapagartmp"

If ctParcela.Text = "" Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If

If ctParcela.Text = 0 Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If
    
Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagar order by tbCtpgLan")
If Reg_Dados.EOF Then
    mLan = 1
Else
    Reg_Dados.MoveLast
    mLan = Reg_Dados!tbCtpgLan + 1
End If
Reg_Dados.Close
Set Reg_Dados = Nothing

numParcelasFim = ctParcela

Do While numParcelas <= numParcelasFim
    
    Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagartmp order by tbCtpgLan")
    Reg_Dados.AddNew
    Reg_Dados!tbCtpgOp = 0
    Reg_Dados!tbCtpgLan = mLan
    Reg_Dados!tbCtpgDataLan = Date
    Reg_Dados!tbCtpgGrupo = cbGrupo
    Reg_Dados!tbCtpgSubGrupo = cbSubGrupo
    Reg_Dados!tbCtpgHistorico = cthist & " - " & numParcelas & "/" & ctParcela
    Reg_Dados!tbCtpgDataVenc = MaskdataVenc
    Reg_Dados!tbCtpgValor = ctvalor
    Reg_Dados!tbCtpgStatus = "Aberto"
    Reg_Dados.Update
    
    mLan = mLan + 1
    numParcelas = numParcelas + 1
    MaskdataVenc = DateAdd("m", 1, MaskdataVenc)
Loop

Label4.Visible = False
ctParcela.Visible = False

dataContasApagar.RecordSource = "Select * from tbcontasapagartmp"
dataContasApagar.Refresh
Call sbLimpa_Campos(Me)
mLan = 0

End Sub

Public Sub addOutros()
Dim numParcelas As Integer
Dim numParcelasFim As Integer
numParcelas = 1

strtb = "tbcontasapagartmp"

If ctParcela.Text = "" Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If

If ctParcela.Text = 0 Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If

numParcelasFim = ctParcela
mLan = 1
Do While numParcelas <= numParcelasFim
    Sql = "INSERT INTO tbcontasapagartmp (Op,datalan,grupo,subgrupo,historico,datavenc,valor,status,lan,numParc) VALUES ('" & MOP & "', " & _
    "'" & Format$(MaskdataVenc, "yyyy-mm-dd hh:mm:ss") & "', '" & cbGrupo & "', '" & cbSubGrupo & "', '" & cthist & "', " & _
    "'" & Format$(MaskdataVenc, "yyyy-mm-dd hh:mm:ss") & "'," & converte(ctvalor) & ", '" & "Aberto" & "'," & mLan & ",'" & numParcelas & "/" & ctParcela & "')"
    conn.Execute Sql
    
    mLan = mLan + 1
    numParcelas = numParcelas + 1
    MaskdataVenc = DateAdd("m", 1, MaskdataVenc)
Loop

Label4.Visible = False
ctParcela.Visible = False
 With datapag
        .ConnectionString = strDns
        .RecordSource = "Select * from tbcontasapagartmp ORDER BY id ASC"
End With
datapag.Refresh

Call sbLimpa_Campos(Me)
mLan = 0


End Sub

Private Sub sbAddPeriodoRel()
If maskdataini.Text = "__/__/____" Then
    Exit Sub
End If

If maskdatafin.Text = "__/__/____" Then
    Exit Sub
End If


Set Reg_Dados = BANCO.OpenRecordset("select * from tbperiodorel")
If Not Reg_Dados.EOF Then
    Reg_Dados.Delete
End If
Reg_Dados.AddNew
Reg_Dados!tbdatData_ini = maskdataini
Reg_Dados!tbdatData_fin = maskdatafin
Reg_Dados.Update
Reg_Dados.Close
Set Reg_Dados = Nothing
End Sub

Private Sub mnuConsultaTitulosBaixadaos_Click()
If maskdataini.Text = "__/__/____" Then
    MsgBox "Digite a data inicial da consulta p/ continuar", vbCritical, "Assistente de verificação"
    Exit Sub
End If

If maskdatafin.Text = "__/__/____" Then
    MsgBox "Digite a data final da consulta p/ continuar", vbCritical, "Assistente de verificação"
    Exit Sub
End If

With datapag
        .ConnectionString = strDns
        .RecordSource = "select * from tbcontasapagar where DataVenc>='" & Format(maskdataini, "yyyy-mm-dd hh:mm:ss") & "'" & _
    "and DataVenc<='" & Format(maskdatafin, "yyyy-mm-dd hh:mm:ss") & "' and Status='" & "Pago" & "' order by DataVenc ASC"
End With
datapag.Refresh


End Sub

Private Sub mnuReldeTitulosemDetalhe_Click()
Dim Mdata_ini As String
Dim Mdata_fin As String
Dim Msit As String
Msit = "Aberto"
Call sbAddPeriodoRel

On Error GoTo E:
Mdata_ini = Format(maskdataini, "yyyy,mm,dd")
Mdata_fin = Format(maskdatafin, "yyyy,mm,dd")
CRP.Password = "EAPS1501"
CRP.DataFiles(0) = BANCO.Name
CRP.SelectionFormula = ""
CRP.ReportFileName = App.Path & "\..\relatorios\RelAgdContasapagar.rpt"
CRP.SelectionFormula = "{relcontasapagar.tbCtpgDataVenc} >= date(" & Mdata_ini & ") and {relcontasapagar.tbCtpgDataVenc} <= date(" & Mdata_fin & ") and {relcontasapagar.tbCtpgStatus}='" & Msit & "'"
CRP.Destination = crptToWindow
CRP.WindowState = crptMaximized
CRP.Action = 0
E:
If Err.Number = 20515 Then
    MsgBox "Consulte os titulos p/ continuar.", vbCritical, "Assistente de Consulta"
    Exit Sub
End If
End Sub

Private Sub mnuRelGeraldeTitulos_Click()
Dim Mdata_ini As String
Dim Mdata_fin As String
Dim Msit As String
Msit = "Aberto"
'Call sbAddPeriodoRel

On Error GoTo E:
Mdata_ini = Format(maskdataini, "yyyy,mm,dd")
Mdata_fin = Format(maskdatafin, "yyyy,mm,dd")
CRP.Connect = strDns
CRP.SelectionFormula = ""
CRP.ReportFileName = App.Path & "\..\relatorios\RelGContasapagar.rpt"
CRP.SelectionFormula = "{tbcontasapagar.DataVenc} >= date(" & Mdata_ini & ") and {tbcontasapagar.DataVenc} <= date(" & Mdata_fin & ") and {tbcontasapagar.Status}='" & Msit & "'"
CRP.Destination = crptToWindow
CRP.WindowState = crptMaximized
CRP.Action = 0
E:
If Err.Number = 20515 Then
    MsgBox "Consulte os titulos p/ continuar.", vbCritical, "Assistente de Consulta"
    Exit Sub
End If

End Sub

Private Sub mnuRelTitulosBaixados_Click()
Dim Mdata_ini As String
Dim Mdata_fin As String
Dim Msit As String
Msit = "Pago"
'Call sbAddPeriodoRel

On Error GoTo E:
Mdata_ini = Format(maskdataini, "yyyy,mm,dd")
Mdata_fin = Format(maskdatafin, "yyyy,mm,dd")
CRP.Connect = strDns
CRP.SelectionFormula = ""
CRP.ReportFileName = App.Path & "\..\relatorios\RelGContaspagas.rpt"
CRP.SelectionFormula = "{tbcontasapagar.DataVenc} >= date(" & Mdata_ini & ") and {tbcontasapagar.DataVenc} <= date(" & Mdata_fin & ") and {tbcontasapagar.Status}='" & Msit & "'"
CRP.Destination = crptToWindow
CRP.WindowState = crptMaximized
CRP.Action = 0
E:
If Err.Number = 20515 Then
    MsgBox "Consulte os titulos p/ continuar.", vbCritical, "Assistente de Consulta"
    Exit Sub
End If
End Sub
Private Sub sbaddCartaoCredito()
Dim numParcelas As Integer
Dim numParcelasFim As Integer
numParcelas = 1

strtb = "tbcontasapagartmp"

If ctParcela.Text = "" Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If

If ctParcela.Text = 0 Then
    MsgBox "Digite a quantidade de parcelas p/ continuar.", vbCritical, "Assistente de verificação"
    ctParcela.SetFocus
    Exit Sub
    
End If
    
Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagar order by tbCtpgLan")
If Reg_Dados.EOF Then
    mLan = 1
Else
    Reg_Dados.MoveLast
    mLan = Reg_Dados!tbCtpgLan + 1
End If
Reg_Dados.Close
Set Reg_Dados = Nothing

numParcelasFim = ctParcela

Do While numParcelas <= numParcelasFim
    
    Set Reg_Dados = BANCO.OpenRecordset("Select * from tbcontasapagartmp order by tbCtpgLan")
    Reg_Dados.AddNew
    Reg_Dados!tbCtpgOp = 0
    Reg_Dados!tbCtpgLan = mLan
    Reg_Dados!tbCtpgDataLan = Date
    Reg_Dados!tbCtpgGrupo = cbGrupo
    Reg_Dados!tbCtpgSubGrupo = cbSubGrupo
    Reg_Dados!tbCtpgHistorico = cthist & " - Parcela: " & numParcelas & "/" & ctParcela
    Reg_Dados!tbCtpgDataVenc = MaskdataVenc
    Reg_Dados!tbCtpgValor = ctvalor
    Reg_Dados!tbCtpgStatus = "Aberto"
    Reg_Dados.Update
    
    mLan = mLan + 1
    numParcelas = numParcelas + 1
    MaskdataVenc = DateAdd("m", 1, MaskdataVenc)
Loop

Label4.Visible = False
ctParcela.Visible = False

dataContasApagar.RecordSource = "Select * from tbcontasapagartmp"
dataContasApagar.Refresh
Call sbLimpa_Campos(Me)
mLan = 0

End Sub




Private Sub mnuSair_Click()
Unload Me
End Sub
