VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUVT 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   2850
   ClientTop       =   3675
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   12795
   Begin VB.CommandButton cmdExcChamado 
      Height          =   645
      Left            =   6960
      Picture         =   "frmUVT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmdAtualizar 
      Height          =   645
      Left            =   6150
      Picture         =   "frmUVT.frx":054B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   30
      Width           =   795
   End
   Begin VB.ComboBox CbChamado 
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
      ItemData        =   "frmUVT.frx":0CD0
      Left            =   60
      List            =   "frmUVT.frx":0CDD
      TabIndex        =   2
      Top             =   210
      Width           =   2055
   End
   Begin VB.CommandButton cmdAddChamado 
      DownPicture     =   "frmUVT.frx":0D01
      Height          =   645
      Left            =   5340
      Picture         =   "frmUVT.frx":15E1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   795
   End
   Begin MSDataGridLib.DataGrid DataGridUvt 
      Bindings        =   "frmUVT.frx":1BAA
      Height          =   3285
      Left            =   60
      TabIndex        =   3
      Top             =   690
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   5794
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
         DataField       =   "ID"
         Caption         =   "IDOS"
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
         DataField       =   "Empresa"
         Caption         =   "Empresa"
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
         DataField       =   "situacao"
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
         DataField       =   "UvtChamado"
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
         DataField       =   "DataFin"
         Caption         =   "Fechamento"
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
         DataField       =   "uvtqtdias"
         Caption         =   "Qtd Dias"
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
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3525,166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3344,882
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   959,811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoDataUvt 
      Height          =   405
      Left            =   60
      Top             =   3960
      Width           =   12705
      _ExtentX        =   22410
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
      Caption         =   "Chamdos UVT"
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
   Begin MSMask.MaskEdBox MaskDataIni 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   210
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
   Begin MSMask.MaskEdBox MaskDataFin 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   210
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Abertura"
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
      Left            =   2190
      TabIndex        =   8
      Top             =   -30
      Width           =   1140
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fechamento"
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
      Left            =   3630
      TabIndex        =   6
      Top             =   -30
      Width           =   1395
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relatorio"
      Begin VB.Menu mnuRelChamados 
         Caption         =   "Relação de Chamado"
      End
   End
End
Attribute VB_Name = "frmUVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nIDOS As Variant
Private Sub cmdAddChamado_Click()
Dim nIdUVT As Integer
    Dim nQtdDias As Integer
    Call sbAtDataUvt
    
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select  MAX(id) as nCountId from tbUvt where IdOs=" & ctNos & "", conn
    If IsNull(rsTabelas!ncountId) = True Then
        nIdUVT = 1
    Else
        nIdUVT = rsTabelas!ncountId + 1
    End If
    rsTabelas.Close
    Set rsTabelas = Nothing
    
    nQtdDias = DateDiff("d", MaskDataIni, MaskDataFin)

    Sql = "INSERT INTO tbUvt (id,IdOs, Status, dataIni, DataFin, QtdDias) VALUES (" & nIdUVT & ", " & ctNos & "," & _
        "'" & CbChamado & "','" & Format(MaskDataIni, "yyyy-mm-dd") & "', '" & Format(MaskDataFin, "yyyy-mm-dd") & "'," & nQtdDias & ")"
    conn.Execute Sql
    AdoDataUvt.Refresh
End Sub

Private Sub cmdAtualizar_Click()
Sql = "UPDATE tbuvt SET StatusOS='" & "Fechado" & "' where IdOs=" & nIDOS & ""
conn.Execute Sql
Call sbAtDataUvt

End Sub

Private Sub cmdExcChamado_Click()
Dim nIdac As Integer
On Error GoTo E:

nIdac = InputBox("Digite o ID do ítem para continuar.")

Sql = "DELETE from tbUvt where id=" & nIdac & " and IdOs=" & ctNos & ""

conn.Execute Sql

AdoDataUvt.Refresh

E:
If Err.Number = 13 Then
    MsgBox "Digite o ID do ítem para continuar."
    Exit Sub
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DataGridUvt_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    nIDOS = 0
    nIDOS = DataGridUvt.Columns(0)
    MsgBox "O chamado da OS de N° " & nIDOS & " foi selecionado."
    
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
Call sbAtDataUvt
End Sub
Private Sub sbAtDataUvt()
Dim nData As Date
'nData = DateAdd("d", 3, Date)
nData = "11/05/2015"
With AdoDataUvt
        .ConnectionString = strDns
        .RecordSource = "Select * From osdb where situacao<>'" & "Fechado" & "'"  ' and DataFin<='" & Format(nData, "yyyy-mm-dd 00:00:00") & "'"
        '.RecordSource = "Select * From tbUvt where StatusOs='" & "Aberto" & "'  order by id"
   End With
   AdoDataUvt.Refresh

End Sub
Private Sub sbVerificaUVT()
Set rsTabelas = New ADODB.Recordset

'rsTabelas.Open "Select * From tbUvt where StatusOs='" & "Aberto" & "'  order by id", conn
Dim nData As Date
nData = DateAdd("d", 3, Date)
rsTabelas.Open "Select * From tbUvt where DataFin> DateTime (" & Format(nData, "yyyy-mm-dd hh:mm:ss") & ")  order by id", conn



End Sub
