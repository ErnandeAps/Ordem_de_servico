VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOsRemota 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   4125
   ClientTop       =   4230
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   12180
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   210
      TabIndex        =   1
      Top             =   360
      Width           =   675
   End
   Begin MSDataGridLib.DataGrid Dgb 
      Bindings        =   "frmOsRemota.frx":0000
      Height          =   4695
      Left            =   930
      TabIndex        =   0
      Top             =   30
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoData 
      Height          =   405
      Left            =   300
      Top             =   4770
      Width           =   11445
      _ExtentX        =   20188
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
End
Attribute VB_Name = "frmOsRemota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Sql As String
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from osdb where situacao='" & "Solicitacao em analise" & "' order by id", connRemoto
'On Error Resume Next

Do While Not rsTabelas.EOF

    Sql = "INSERT INTO osdb(empresa,cnpj,insc,endereco,bairro,cidade,estado,cep,email,telefone,servico,situacao, " & _
    "equipamento,responsavel,op,tecnico,parecer,news,idWeb,idColaborador,data,hora) VALUE ('" & rsTabelas!empresa & "', " & _
    "'" & rsTabelas!cnpj & "', '" & rsTabelas!insc & "', '" & rsTabelas!endereco & "', '" & rsTabelas!bairro & "', " & _
    "'" & rsTabelas!cidade & "', '" & rsTabelas!estado & "', '" & rsTabelas!cep & "', '" & rsTabelas!email & "', " & _
    "'" & rsTabelas!telefone & "', '" & rsTabelas!servico & "', '" & rsTabelas!situacao & "', '" & rsTabelas!equipamento & "', " & _
    "'" & rsTabelas!responsavel & "', '" & rsTabelas!op & "', '" & rsTabelas!tecnico & "', '" & rsTabelas!parecer & "', " & _
    "'" & rsTabelas!news & "', " & rsTabelas!id & ", " & rsTabelas!idcolaborador & ", '" & rsTabelas!Data & "', '" & rsTabelas!hora & "')"
    
    conn.Execute Sql
    
rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing

End Sub

Private Sub Form_Load()
'With AdoData
'        .ConnectionString = strDnsRemoto
'        .RecordSource = "Select * From osdb"
'   End With
'   AdoData.Refresh
End Sub
