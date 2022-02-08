VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmConCnpj 
   Caption         =   "Consulta de CNPJ Exemplo em VB6 por fernando-mm@hotmail.com"
   ClientHeight    =   8925
   ClientLeft      =   2805
   ClientTop       =   1770
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   787
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   11535
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   240
         TabIndex        =   50
         Top             =   5880
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nome"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Qualificação"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nome do Repres. Legal"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qualif. do Repres. Legal"
            Object.Width           =   4851
         EndProperty
      End
      Begin VB.TextBox MemoAtividadesSecundarias 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   4080
         Width           =   11175
      End
      Begin VB.TextBox EditAtividadePrincipal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   46
         Top             =   3480
         Width           =   11175
      End
      Begin VB.TextBox EditDataSituacao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   44
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox EditSituacao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7680
         TabIndex        =   42
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox EditNaturezaJuridica 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   40
         Top             =   2880
         Width           =   7350
      End
      Begin VB.TextBox EditCapitalSocial 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9360
         TabIndex        =   38
         Top             =   2280
         Width           =   2025
      End
      Begin VB.TextBox EditEFR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   36
         Top             =   2280
         Width           =   2700
      End
      Begin VB.TextBox EditTelefone 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3840
         TabIndex        =   34
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox EditEmail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox EditCEP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         TabIndex        =   30
         Top             =   1680
         Width           =   1905
      End
      Begin VB.TextBox EditUF 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8880
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox EditCidade 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         TabIndex        =   26
         Top             =   1680
         Width           =   4275
      End
      Begin VB.TextBox EditComplemento 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox EditBairro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   22
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox EditNumero 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   20
         Top             =   1080
         Width           =   2070
      End
      Begin VB.TextBox EditEndereco 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox EditFantasia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8280
         TabIndex        =   16
         Top             =   480
         Width           =   3120
      End
      Begin VB.TextBox EditAbertura 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   14
         Top             =   480
         Width           =   1110
      End
      Begin VB.TextBox EditRazaoSocial 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox EditTipoEmpresa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1470
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Quadro de Sócios e Administradores - QSA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   49
         Top             =   5640
         Width           =   3165
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Código e Descrição das Atividades Econômicas Secundárias "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   47
         Top             =   3840
         Width           =   4455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Código e Descrição da Atividade Econômica Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   45
         Top             =   3240
         Width           =   3825
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Data da Situação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9600
         TabIndex        =   43
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Situação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7680
         TabIndex        =   41
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Código e Descrição da Natureza Jurídica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   2640
         Width           =   2940
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Capital Social"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9360
         TabIndex        =   37
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ente Federativo Responsável (EFR)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6600
         TabIndex        =   35
         Top             =   2040
         Width           =   2565
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   33
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Endereço Eletrônico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   2040
         Width           =   1440
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9480
         TabIndex        =   29
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8880
         TabIndex        =   27
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4440
         TabIndex        =   25
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7080
         TabIndex        =   21
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4920
         TabIndex        =   19
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Endereco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8280
         TabIndex        =   15
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Abertura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7080
         TabIndex        =   13
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11580
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   52
         Top             =   870
         Width           =   1815
      End
      Begin VB.CommandButton ButConsulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   7
         Top             =   510
         Width           =   1815
      End
      Begin VB.CommandButton ButLerCaptcha 
         Caption         =   "Ler Captcha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   51
         Top             =   150
         Width           =   1815
      End
      Begin VB.TextBox EditCaptcha 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "00.000.000/0000-00;1;_"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox EditCNPJ 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "00.000.000/0000-00;1;_"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   0
         Text            =   "32.772.873/0001-09"
         Top             =   480
         Width           =   3915
      End
      Begin VB.PictureBox ImgCaptcha 
         Height          =   750
         Left            =   120
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   176
         TabIndex        =   2
         Top             =   240
         Width           =   2700
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Captcha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   7080
         TabIndex        =   6
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Digite o CNPJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label LabBaixarCaptcha 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ATUALIZAR CAPTCHA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   600
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1000
         Width           =   1695
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuImportar 
         Caption         =   "Importar"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuAtCaptcha 
         Caption         =   "Atualizar Captcha"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFechar 
         Caption         =   "Fechar"
      End
   End
End
Attribute VB_Name = "frmConCnpj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButConsulta_Click()
    If UFSistConsultaCNPJ.Consulta(EditCNPJ.Text, EditCaptcha.Text) = 1 Then
        EditTipoEmpresa.Text = UFSistConsultaCNPJ.EmpresaTipo
        EditAbertura.Text = UFSistConsultaCNPJ.Abertura
        EditRazaoSocial.Text = UFSistConsultaCNPJ.RazaoSocial
        EditFantasia.Text = UFSistConsultaCNPJ.Fantasia
        EditEndereco.Text = UFSistConsultaCNPJ.Logradouro
        EditNumero.Text = UFSistConsultaCNPJ.Numero
        EditComplemento.Text = UFSistConsultaCNPJ.Complemento
        EditCEP.Text = UFSistConsultaCNPJ.CEP
        EditBairro.Text = UFSistConsultaCNPJ.BairroDistrito
        EditCidade.Text = UFSistConsultaCNPJ.Municipio
        EditUF.Text = UFSistConsultaCNPJ.UF
        EditEmail.Text = UFSistConsultaCNPJ.EnderecoEletronico
        EditTelefone.Text = UFSistConsultaCNPJ.Telefone
        EditEFR.Text = UFSistConsultaCNPJ.EnteFederativoResponsavel
        EditSituacao.Text = UFSistConsultaCNPJ.SituacaoCadastral
        EditDataSituacao.Text = UFSistConsultaCNPJ.DataSituacaoCadastral
        'EditMotivoSituacaoCadastral.Text = UFSistConsultaCNPJ.MotivoSituacaoCadastral
        'EditSituacaoEspecial.Text = UFSistConsultaCNPJ.SituacaoEspecial
        'EditDataSituacaoEspecial.Text = UFSistConsultaCNPJ.DataSituacaoEspecial
        EditCapitalSocial.Text = UFSistConsultaCNPJ.CapitalSocial
        EditAtividadePrincipal.Text = UFSistConsultaCNPJ.CodigoDescricaoAtividadeEconomicaPrincipal
        EditNaturezaJuridica.Text = UFSistConsultaCNPJ.CodigoDescricaoNaturezaJuridica
        
        Dim Atividades As String
        Dim I As Integer
        I = 0
        While (I < UFSistConsultaCNPJ.CodigoDescricaoAtividadeEconomicaSecundariasCount)
            Atividades = Atividades & _
                UFSistConsultaCNPJ.CodigoDescricaoAtividadeEconomicaSecundarias(I) & vbCrLf
            I = I + 1
        Wend
        MemoAtividadesSecundarias.Text = Atividades
        
        ListView1.ListItems.Clear
        I = 0
        While (I < UFSistConsultaCNPJ.SociosCount)
            With ListView1.ListItems.Add
                .Text = UFSistConsultaCNPJ.SociosNome(I)
                .SubItems(1) = UFSistConsultaCNPJ.SociosQualificacao(I)
                .SubItems(2) = UFSistConsultaCNPJ.SociosNomedoRepresLegal(I)
                .SubItems(3) = UFSistConsultaCNPJ.SociosQualifRepLegal(I)
            End With
            I = I + 1
        Wend
        UFSistConsultaCNPJ.SalvarXML ("Resultado EM XML.xml")
    Else
        MsgBox (UFSistConsultaCNPJ.Erro)
    End If
End Sub

Private Sub ButLerCaptcha_Click()
    EditCaptcha.Text = CaptchaBoss("Captcha.jpg", "A9O981GLMUBG3UY9G6HPKCTPG8Y62P1MMXRWOI56")
End Sub

Private Sub cmdImportar_Click()
frmOs.ctrazao = EditRazaoSocial
frmOs.MASKCEP = EditCEP.Text
frmOs.ctend = EditEndereco & ", " & EditNumero
frmOs.ctbairro = EditBairro
frmOs.ctcidade = EditCidade
frmOs.ctuf = EditUF
frmOs.ctcgc = EditCNPJ
Unload Me
End Sub

Private Sub EditCaptcha_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    Call sbConsultarCnpj
End Select
End Sub

Private Sub EditCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    EditCaptcha.SetFocus
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
    'If UFSistConsultaCNPJ.Proxy("10.1.1.4", "8888", "", "") = 1 Then
    'Else
    '    MsgBox (UFSistConsultaCNPJ.Erro)
    'End If
    If UFSistConsultaCNPJ.CaptchaBaixar("Captcha.jpg") = 1 Then
        ImgCaptcha.Picture = LoadPicture("Captcha.jpg")
    Else
        MsgBox (UFSistConsultaCNPJ.Erro)
    End If
    EditCNPJ.Text = ""
    'EditCNPJ.SetFocus
End Sub

Private Sub LabBaixarCaptcha_Click()
    If UFSistConsultaCNPJ.CaptchaBaixar("Captcha.jpg") = 1 Then
        ImgCaptcha.Picture = LoadPicture("Captcha.jpg")
    Else
        MsgBox (UFSistConsultaCNPJ.Erro)
    End If
    EditCaptcha.Text = ""
    EditCaptcha.SetFocus
End Sub

Private Sub sbConsultarCnpj()
If UFSistConsultaCNPJ.Consulta(EditCNPJ.Text, EditCaptcha.Text) = 1 Then
        EditTipoEmpresa.Text = UFSistConsultaCNPJ.EmpresaTipo
        EditAbertura.Text = UFSistConsultaCNPJ.Abertura
        EditRazaoSocial.Text = UFSistConsultaCNPJ.RazaoSocial
        EditFantasia.Text = UFSistConsultaCNPJ.Fantasia
        EditEndereco.Text = UFSistConsultaCNPJ.Logradouro
        EditNumero.Text = UFSistConsultaCNPJ.Numero
        EditComplemento.Text = UFSistConsultaCNPJ.Complemento
        EditCEP.Text = UFSistConsultaCNPJ.CEP
        EditBairro.Text = UFSistConsultaCNPJ.BairroDistrito
        EditCidade.Text = UFSistConsultaCNPJ.Municipio
        EditUF.Text = UFSistConsultaCNPJ.UF
        EditEmail.Text = UFSistConsultaCNPJ.EnderecoEletronico
        EditTelefone.Text = UFSistConsultaCNPJ.Telefone
        EditEFR.Text = UFSistConsultaCNPJ.EnteFederativoResponsavel
        EditSituacao.Text = UFSistConsultaCNPJ.SituacaoCadastral
        EditDataSituacao.Text = UFSistConsultaCNPJ.DataSituacaoCadastral
        'EditMotivoSituacaoCadastral.Text = UFSistConsultaCNPJ.MotivoSituacaoCadastral
        'EditSituacaoEspecial.Text = UFSistConsultaCNPJ.SituacaoEspecial
        'EditDataSituacaoEspecial.Text = UFSistConsultaCNPJ.DataSituacaoEspecial
        EditCapitalSocial.Text = UFSistConsultaCNPJ.CapitalSocial
        EditAtividadePrincipal.Text = UFSistConsultaCNPJ.CodigoDescricaoAtividadeEconomicaPrincipal
        EditNaturezaJuridica.Text = UFSistConsultaCNPJ.CodigoDescricaoNaturezaJuridica
        
        Dim Atividades As String
        Dim I As Integer
        I = 0
        While (I < UFSistConsultaCNPJ.CodigoDescricaoAtividadeEconomicaSecundariasCount)
            Atividades = Atividades & _
                UFSistConsultaCNPJ.CodigoDescricaoAtividadeEconomicaSecundarias(I) & vbCrLf
            I = I + 1
        Wend
        MemoAtividadesSecundarias.Text = Atividades
        
        ListView1.ListItems.Clear
        I = 0
        While (I < UFSistConsultaCNPJ.SociosCount)
            With ListView1.ListItems.Add
                .Text = UFSistConsultaCNPJ.SociosNome(I)
                .SubItems(1) = UFSistConsultaCNPJ.SociosQualificacao(I)
                .SubItems(2) = UFSistConsultaCNPJ.SociosNomedoRepresLegal(I)
                .SubItems(3) = UFSistConsultaCNPJ.SociosQualifRepLegal(I)
            End With
            I = I + 1
        Wend
        UFSistConsultaCNPJ.SalvarXML ("Resultado EM XML.xml")
    Else
        MsgBox (UFSistConsultaCNPJ.Erro)
    End If
End Sub

Private Sub mnuAtCaptcha_Click()
If UFSistConsultaCNPJ.CaptchaBaixar("Captcha.jpg") = 1 Then
        ImgCaptcha.Picture = LoadPicture("Captcha.jpg")
    Else
        MsgBox (UFSistConsultaCNPJ.Erro)
    End If
    EditCaptcha.Text = ""
    EditCaptcha.SetFocus
End Sub

Private Sub mnuFechar_Click()
Unload Me
End Sub

Private Sub mnuImportar_Click()
frmOs.ctrazao = EditRazaoSocial
frmOs.MASKCEP = EditCEP.Text
frmOs.ctend = EditEndereco & ", " & EditNumero
frmOs.ctbairro = EditBairro
frmOs.ctcidade = EditCidade
frmOs.ctuf = EditUF
frmOs.ctcgc = Format(EditCNPJ, "@@.@@@.@@@/@@@@-@@")
Unload Me
End Sub
