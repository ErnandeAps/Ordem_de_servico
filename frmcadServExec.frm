VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcadServExec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Servi�os executados"
   ClientHeight    =   4290
   ClientLeft      =   5175
   ClientTop       =   4065
   ClientWidth     =   7770
   ClipControls    =   0   'False
   Icon            =   "frmcadServExec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ctvalor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5220
      MaxLength       =   25
      TabIndex        =   5
      Top             =   270
      Width           =   1395
   End
   Begin VB.TextBox ctdescricao 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   90
      MaxLength       =   25
      TabIndex        =   0
      Top             =   270
      Width           =   5115
   End
   Begin VB.CommandButton cmdInsert 
      Height          =   435
      Left            =   6660
      Picture         =   "frmcadServExec.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   525
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   435
      Left            =   7200
      Picture         =   "frmcadServExec.frx":61DB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin MSAdodcLib.Adodc DataAdo 
      Height          =   435
      Left            =   90
      Top             =   3810
      Width           =   7605
      _ExtentX        =   13414
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
      Caption         =   "Acess�rios"
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
   Begin MSDataGridLib.DataGrid dbgacessorios 
      Bindings        =   "frmcadServExec.frx":6726
      Height          =   3105
      Left            =   90
      TabIndex        =   1
      Top             =   690
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5477
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
         DataField       =   "valunit"
         Caption         =   "Valor"
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
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4694,74
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1365,165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
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
      TabIndex        =   6
      Top             =   30
      Width           =   420
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   795
   End
End
Attribute VB_Name = "frmcadServExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInsertPecas_Click()

End Sub

Private Sub cmdExcluir_Click()
Dim nIdac As Integer
On Error GoTo E:

nIdac = InputBox("Digite o ID do �tem para continuar.")

Sql = "DELETE from cadServexec where id=" & nIdac & ""

conn.Execute Sql

DataAdo.Refresh

E:
If Err.Number = 13 Then
    MsgBox "Digite o ID do �tem para continuar."
    Exit Sub
End If
End Sub

Private Sub cmdInsert_Click()
Sql = "INSERT INTO cadServexec (descricao, valunit) VALUE ('" & ctdescricao & "'," & converte(ctvalor) & ")"

conn.Execute Sql

DataAdo.Refresh
ctdescricao.Text = ""
ctvalor.Text = ""
ctdescricao.SetFocus

End Sub

Private Sub ctdescricao_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctdescricao.Text = "" Then
        MsgBox "Digite a descri��o para continuar", vbCritical
        Exit Sub
    End If
    
    ctvalor.SetFocus
End Select
End Sub

Private Sub ctvalor_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctvalor.Text = "" Then
        MsgBox "Digite o valor para continuar.", vbCritical
        Exit Sub
    End If
    cmdInsert.SetFocus
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

With DataAdo
        .ConnectionString = strDns
        .RecordSource = "Select * From cadServexec order by id"
   End With
    DataAdo.Refresh
End Sub
