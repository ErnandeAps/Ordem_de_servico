VERSION 5.00
Object = "{CC01E02E-242E-11D8-A6B9-000B231D9747}#1.0#0"; "vertmenu.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordem de Serviço"
   ClientHeight    =   7515
   ClientLeft      =   2655
   ClientTop       =   1995
   ClientWidth     =   14025
   FillColor       =   &H00000080&
   Icon            =   "frmOs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   Begin TabDlg.SSTab SSTab2 
      Height          =   5685
      Left            =   1560
      TabIndex        =   12
      Top             =   1770
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   10028
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   128
      TabCaption(0)   =   "Dados Cliente"
      TabPicture(0)   =   "frmOs.frx":2B06
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Equipamento"
      TabPicture(1)   =   "frmOs.frx":2B22
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Controle de Chamado UVT"
      TabPicture(2)   =   "frmOs.frx":2B3E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label41"
      Tab(2).Control(1)=   "Label42"
      Tab(2).Control(2)=   "Label43"
      Tab(2).Control(3)=   "MaskDataFin"
      Tab(2).Control(4)=   "MaskDataIni"
      Tab(2).Control(5)=   "AdoDataUvt"
      Tab(2).Control(6)=   "DataGridUvt"
      Tab(2).Control(7)=   "CbChamado"
      Tab(2).Control(8)=   "cmdExcChamado"
      Tab(2).Control(9)=   "cmdAddChamado"
      Tab(2).ControlCount=   10
      Begin VB.CommandButton cmdAddChamado 
         Height          =   405
         Left            =   -69480
         Picture         =   "frmOs.frx":2B5A
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   690
         Width           =   495
      End
      Begin VB.CommandButton cmdExcChamado 
         Height          =   405
         Left            =   -68970
         Picture         =   "frmOs.frx":3123
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   690
         Width           =   495
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
         ItemData        =   "frmOs.frx":366E
         Left            =   -74880
         List            =   "frmOs.frx":3678
         TabIndex        =   144
         Top             =   720
         Width           =   2055
      End
      Begin VB.Frame Frame6 
         Height          =   1185
         Left            =   150
         TabIndex        =   70
         Top             =   4380
         Width           =   12105
         Begin VB.CommandButton cmdImportarCad 
            Caption         =   "Importar Cliente"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5940
            Picture         =   "frmOs.frx":3693
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   210
            Width           =   1155
         End
         Begin VB.CommandButton cmdProximo 
            Caption         =   "Próximo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10800
            TabIndex        =   83
            Top             =   660
            Width           =   1215
         End
         Begin VB.CommandButton BtNovo 
            Caption         =   "Cliente Novo"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2430
            Picture         =   "frmOs.frx":46D5
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   210
            Width           =   1155
         End
         Begin VB.CommandButton BtConsultar 
            Caption         =   "&Consultar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3600
            Picture         =   "frmOs.frx":5717
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Consulta Cliente pelo Nome"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4770
            Picture         =   "frmOs.frx":6759
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Altera os dados do Cliente"
            Top             =   210
            Width           =   1155
         End
         Begin Crystal.CrystalReport crp 
            Left            =   870
            Top             =   390
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            PrintFileName   =   "C:\Projeto Bematech\Sistema\Relatorios\OS.rpt"
            PrintFileLinesPerPage=   60
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3855
         Left            =   150
         TabIndex        =   39
         Top             =   510
         Width           =   12105
         Begin VB.ListBox LTNOME 
            Appearance      =   0  'Flat
            Height          =   2370
            ItemData        =   "frmOs.frx":81CB
            Left            =   1440
            List            =   "frmOs.frx":81CD
            TabIndex        =   40
            Top             =   780
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.TextBox MASKCEP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   90
            TabIndex        =   150
            Top             =   990
            Width           =   1515
         End
         Begin VB.ListBox Lt_cep 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00C00000&
            Height          =   1980
            ItemData        =   "frmOs.frx":81CF
            Left            =   90
            List            =   "frmOs.frx":81D1
            TabIndex        =   42
            Top             =   1380
            Visible         =   0   'False
            Width           =   8505
         End
         Begin VB.TextBox CTCGC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3120
            TabIndex        =   115
            Top             =   1590
            Width           =   2325
         End
         Begin VB.ComboBox cbcolaborador 
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
            Left            =   7470
            TabIndex        =   96
            Top             =   390
            Width           =   4515
         End
         Begin VB.TextBox ctcodClient 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   90
            MaxLength       =   8
            TabIndex        =   51
            Top             =   390
            Width           =   1305
         End
         Begin VB.TextBox ctrazao 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1440
            MaxLength       =   39
            TabIndex        =   50
            Top             =   390
            Width           =   4815
         End
         Begin VB.TextBox CTEND 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   49
            Top             =   990
            Width           =   4155
         End
         Begin VB.TextBox CTCIDADE 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   8610
            MaxLength       =   25
            TabIndex        =   48
            Top             =   990
            Width           =   2775
         End
         Begin VB.TextBox ctEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   90
            TabIndex        =   47
            Top             =   2190
            Width           =   3375
         End
         Begin VB.TextBox CTUF 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   11400
            MaxLength       =   2
            TabIndex        =   46
            Top             =   990
            Width           =   345
         End
         Begin VB.TextBox CTINSC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5460
            TabIndex        =   45
            Top             =   1590
            Width           =   2205
         End
         Begin VB.TextBox ctroteiro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   2850
            Width           =   10425
         End
         Begin VB.TextBox ctContato 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3480
            TabIndex        =   43
            Top             =   2190
            Width           =   4275
         End
         Begin VB.TextBox ctBairro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5790
            MaxLength       =   50
            TabIndex        =   41
            Top             =   990
            Width           =   2805
         End
         Begin MSMask.MaskEdBox MASKFONE 
            Height          =   375
            Left            =   90
            TabIndex        =   52
            Top             =   1590
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            MaxLength       =   13
            Mask            =   "(##)####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MASKCEL 
            Height          =   375
            Left            =   1620
            TabIndex        =   53
            Top             =   1590
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            MaxLength       =   13
            Mask            =   "(##)####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskDataNasc 
            Height          =   375
            Left            =   6300
            TabIndex        =   54
            Top             =   390
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            MaxLength       =   8
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label48 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "Parceiros"
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
            Left            =   7470
            TabIndex        =   97
            Top             =   150
            Width           =   735
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Código"
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
            TabIndex        =   69
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social / Nome"
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
            Left            =   1440
            TabIndex        =   68
            Top             =   180
            Width           =   1620
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Endereço/Bairro :"
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
            Left            =   1620
            TabIndex        =   67
            Top             =   780
            Width           =   1380
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Cidade :"
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
            Left            =   8580
            TabIndex        =   66
            Top             =   780
            Width           =   675
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Email :"
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
            TabIndex        =   65
            Top             =   1980
            Width           =   510
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "UF :"
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
            Left            =   11430
            TabIndex        =   64
            Top             =   780
            Width           =   330
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "CEP :"
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
            TabIndex        =   63
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fone :"
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
            TabIndex        =   62
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF"
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
            TabIndex        =   61
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Est / RG"
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
            Left            =   5460
            TabIndex        =   60
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Celular"
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
            Left            =   1650
            TabIndex        =   59
            Top             =   1380
            Width           =   555
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Roteiro"
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
            TabIndex        =   58
            Top             =   2610
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Data nasc."
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
            Left            =   6330
            TabIndex        =   57
            Top             =   150
            Width           =   840
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Contato"
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
            Left            =   3480
            TabIndex        =   56
            Top             =   1980
            Width           =   645
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Bairro :"
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
            Left            =   5790
            TabIndex        =   55
            Top             =   780
            Width           =   570
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5145
         Left            =   -74910
         TabIndex        =   13
         Top             =   420
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9075
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Equipamento"
         TabPicture(0)   =   "frmOs.frx":81D3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmAcessorios"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Defeito Apres."
         TabPicture(1)   =   "frmOs.frx":81EF
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label11"
         Tab(1).Control(1)=   "ctDefeito"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Diagnóstico"
         TabPicture(2)   =   "frmOs.frx":820B
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label12"
         Tab(2).Control(1)=   "Label40"
         Tab(2).Control(2)=   "ctDiagnostico"
         Tab(2).Control(3)=   "ctObs"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Peças"
         TabPicture(3)   =   "frmOs.frx":8227
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label14"
         Tab(3).Control(1)=   "Label13"
         Tab(3).Control(2)=   "Label38"
         Tab(3).Control(3)=   "DataAdoPecas"
         Tab(3).Control(4)=   "DgPecas"
         Tab(3).Control(5)=   "ctqtdPecas"
         Tab(3).Control(6)=   "cbPecas"
         Tab(3).Control(7)=   "cmdInsertPecas"
         Tab(3).Control(8)=   "cmdExcPecas"
         Tab(3).Control(9)=   "cbnserie"
         Tab(3).ControlCount=   10
         TabCaption(4)   =   "Serviços"
         TabPicture(4)   =   "frmOs.frx":8243
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "dataAdoServExec"
         Tab(4).Control(1)=   "Frame8"
         Tab(4).Control(2)=   "Frame9"
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "Histórico"
         TabPicture(5)   =   "frmOs.frx":825F
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "dataAdoHist"
         Tab(5).Control(1)=   "dgHist"
         Tab(5).Control(2)=   "cmdExcAndamentoHist"
         Tab(5).Control(3)=   "cmdInsertAndamentoHist"
         Tab(5).Control(4)=   "cbSituacaoHist"
         Tab(5).ControlCount=   5
         TabCaption(6)   =   "Resumo"
         TabPicture(6)   =   "frmOs.frx":827B
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Label17"
         Tab(6).Control(1)=   "Label33"
         Tab(6).Control(2)=   "framServ"
         Tab(6).Control(3)=   "frampecas"
         Tab(6).Control(4)=   "ctTotalServ"
         Tab(6).Control(5)=   "ctpecasexec"
         Tab(6).Control(6)=   "Frame7"
         Tab(6).ControlCount=   7
         Begin VB.TextBox ctObs 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1980
            Left            =   -74820
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   137
            Top             =   2940
            Width           =   11805
         End
         Begin VB.ComboBox cbnserie 
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
            Left            =   -69660
            TabIndex        =   133
            Top             =   660
            Width           =   2535
         End
         Begin VB.Frame Frame7 
            Caption         =   "Financeiro"
            ForeColor       =   &H00000080&
            Height          =   4035
            Left            =   -69570
            TabIndex        =   119
            Top             =   420
            Width           =   6705
            Begin VB.CommandButton cmdBoleto 
               Height          =   765
               Left            =   4890
               Picture         =   "frmOs.frx":8297
               Style           =   1  'Graphical
               TabIndex        =   132
               ToolTipText     =   "Gerar boleto fatura"
               Top             =   3150
               Width           =   855
            End
            Begin VB.CommandButton cmdRecibo 
               Height          =   765
               Left            =   5730
               Picture         =   "frmOs.frx":8BC8
               Style           =   1  'Graphical
               TabIndex        =   131
               ToolTipText     =   "Imprimir Recibo"
               Top             =   3150
               Width           =   855
            End
            Begin VB.CommandButton cmdReceber 
               Height          =   765
               Left            =   4020
               Picture         =   "frmOs.frx":94FC
               Style           =   1  'Graphical
               TabIndex        =   130
               ToolTipText     =   "Receber pagamento Os."
               Top             =   3150
               Width           =   855
            End
            Begin VB.TextBox ctTotalPago 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4950
               MaxLength       =   8
               TabIndex        =   126
               Top             =   2160
               Width           =   1485
            End
            Begin VB.TextBox ctDesconto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4980
               MaxLength       =   8
               TabIndex        =   124
               Top             =   1590
               Width           =   1485
            End
            Begin VB.ComboBox cbFormaPag 
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
               ItemData        =   "frmOs.frx":9C81
               Left            =   3510
               List            =   "frmOs.frx":9C91
               TabIndex        =   122
               Top             =   1020
               Width           =   2925
            End
            Begin VB.TextBox ctTotalOS 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4980
               MaxLength       =   8
               TabIndex        =   120
               Top             =   360
               Width           =   1485
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Total a Pagar.:"
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
               Left            =   3480
               TabIndex        =   127
               Top             =   2220
               Width           =   1320
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Desconto.:"
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
               Left            =   3810
               TabIndex        =   125
               Top             =   1650
               Width           =   975
            End
            Begin VB.Label Label35 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Forma de Pag.: "
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
               Left            =   2220
               TabIndex        =   123
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Total da OS ="
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
               Left            =   3420
               TabIndex        =   121
               Top             =   420
               Width           =   1320
            End
         End
         Begin VB.TextBox ctpecasexec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -70950
            MaxLength       =   8
            TabIndex        =   117
            Top             =   4500
            Width           =   1305
         End
         Begin VB.TextBox ctTotalServ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -70980
            MaxLength       =   8
            TabIndex        =   112
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Frame frampecas 
            Caption         =   "Peças utilizadas"
            Height          =   1755
            Left            =   -74760
            TabIndex        =   111
            Top             =   2700
            Width           =   5115
            Begin MSDataGridLib.DataGrid dbgpecasexec 
               Bindings        =   "frmOs.frx":9CC0
               Height          =   1365
               Left            =   60
               TabIndex        =   116
               Top             =   210
               Width           =   4965
               _ExtentX        =   8758
               _ExtentY        =   2408
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   "descricao"
                  Caption         =   "Descrição"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   """R$ ""#.##0,00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1046
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "valtotal"
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
                     ColumnWidth     =   3089,764
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     ColumnWidth     =   1289,764
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame framServ 
            Caption         =   "Serviços executados"
            Height          =   1815
            Left            =   -74760
            TabIndex        =   110
            Top             =   420
            Width           =   5115
            Begin MSDataGridLib.DataGrid dbgservexec 
               Bindings        =   "frmOs.frx":9CDB
               Height          =   1485
               Left            =   60
               TabIndex        =   114
               Top             =   210
               Width           =   4965
               _ExtentX        =   8758
               _ExtentY        =   2619
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   "descricao"
                  Caption         =   "Descrição"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   """R$ ""#.##0,00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1046
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "valtotal"
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
                     ColumnWidth     =   3089,764
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     ColumnWidth     =   1289,764
                  EndProperty
               EndProperty
            End
         End
         Begin VB.ComboBox cbSituacaoHist 
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
            Left            =   -74790
            TabIndex        =   109
            Top             =   570
            Width           =   4845
         End
         Begin VB.CommandButton cmdInsertAndamentoHist 
            Height          =   435
            Left            =   -69900
            Picture         =   "frmOs.frx":9CF9
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   510
            Width           =   525
         End
         Begin VB.CommandButton cmdExcAndamentoHist 
            Height          =   435
            Left            =   -69360
            Picture         =   "frmOs.frx":A2C2
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   510
            Width           =   495
         End
         Begin VB.CommandButton cmdExcPecas 
            Height          =   435
            Left            =   -65520
            Picture         =   "frmOs.frx":A80D
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdInsertPecas 
            Height          =   435
            Left            =   -66060
            Picture         =   "frmOs.frx":AD58
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   600
            Width           =   525
         End
         Begin MSDataGridLib.DataGrid dgHist 
            Bindings        =   "frmOs.frx":B321
            Height          =   3495
            Left            =   -74790
            TabIndex        =   100
            Top             =   960
            Width           =   11835
            _ExtentX        =   20876
            _ExtentY        =   6165
            _Version        =   393216
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
               DataField       =   "data"
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
            BeginProperty Column02 
               DataField       =   "eventos"
               Caption         =   "Andamento"
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
               DataField       =   "detalhe"
               Caption         =   "Detalhe"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   705,26
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1860,095
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   3165,166
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2564,788
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2865,26
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame9 
            Caption         =   "Solicitado"
            ForeColor       =   &H00000080&
            Height          =   975
            Left            =   -74880
            TabIndex        =   91
            Top             =   480
            Width           =   12045
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
               Left            =   120
               TabIndex        =   92
               Top             =   420
               Width           =   5955
            End
            Begin VB.Label Label45 
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
               TabIndex        =   93
               Top             =   210
               Width           =   795
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Executado"
            ForeColor       =   &H00000080&
            Height          =   2985
            Left            =   -74880
            TabIndex        =   85
            Top             =   1500
            Width           =   12045
            Begin VB.CommandButton cmdexcServExec 
               Height          =   435
               Left            =   7830
               Picture         =   "frmOs.frx":B33B
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   420
               Width           =   495
            End
            Begin VB.CommandButton cmdinsertServExec 
               Height          =   435
               Left            =   7290
               Picture         =   "frmOs.frx":B886
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   420
               Width           =   525
            End
            Begin VB.TextBox ctqtdServExec 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   6240
               MaxLength       =   8
               TabIndex        =   87
               Top             =   450
               Width           =   1005
            End
            Begin VB.ComboBox cbServExec 
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
               Left            =   120
               TabIndex        =   86
               Top             =   480
               Width           =   5955
            End
            Begin MSDataGridLib.DataGrid DgServExec 
               Bindings        =   "frmOs.frx":BE4F
               Height          =   1995
               Left            =   120
               TabIndex        =   88
               Top             =   900
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   3519
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
               BeginProperty Column02 
                  DataField       =   "qtd"
                  Caption         =   "Qtd"
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
                  Caption         =   "V. Unit"
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
                  DataField       =   "valtotal"
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
                     ColumnWidth     =   840,189
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   6419,906
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   2
                     ColumnWidth     =   854,929
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label16 
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
               Left            =   6240
               TabIndex        =   90
               Top             =   240
               Width           =   930
            End
            Begin VB.Label Label15 
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
               TabIndex        =   89
               Top             =   270
               Width           =   795
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Identificação"
            Height          =   4515
            Left            =   180
            TabIndex        =   23
            Top             =   480
            Width           =   5115
            Begin VB.CommandButton cmdAbrirNF 
               Caption         =   "cc"
               Height          =   405
               Left            =   4560
               Picture         =   "frmOs.frx":BE6D
               Style           =   1  'Graphical
               TabIndex        =   140
               ToolTipText     =   "Visualizar NF Compra Cliente"
               Top             =   3870
               Width           =   495
            End
            Begin MSComDlg.CommonDialog cdiagBox 
               Left            =   4290
               Top             =   3180
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton cmdNfCompra 
               Height          =   405
               Left            =   4050
               Picture         =   "frmOs.frx":E60F
               Style           =   1  'Graphical
               TabIndex        =   139
               ToolTipText     =   "Importar NF Compra Cliente"
               Top             =   3870
               Width           =   495
            End
            Begin VB.CommandButton Command7 
               Height          =   405
               Left            =   4530
               Picture         =   "frmOs.frx":EBD8
               Style           =   1  'Graphical
               TabIndex        =   95
               Top             =   570
               Width           =   495
            End
            Begin VB.ComboBox ctEquipamento 
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
               Left            =   210
               TabIndex        =   94
               Top             =   600
               Width           =   4275
            End
            Begin VB.TextBox ctTombo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   210
               TabIndex        =   28
               Top             =   3900
               Width           =   3795
            End
            Begin VB.TextBox ctRef 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   210
               MaxLength       =   25
               TabIndex        =   27
               Top             =   3210
               Width           =   3795
            End
            Begin VB.TextBox ctModelo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               MaxLength       =   25
               TabIndex        =   26
               Top             =   2520
               Width           =   3795
            End
            Begin VB.TextBox ctMarca 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   210
               MaxLength       =   25
               TabIndex        =   25
               Top             =   1860
               Width           =   3795
            End
            Begin VB.TextBox ctNSerie 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   210
               MaxLength       =   25
               TabIndex        =   24
               Top             =   1200
               Width           =   3795
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "NF Compra"
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
               Left            =   210
               TabIndex        =   34
               Top             =   3660
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Ref."
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
               Left            =   210
               TabIndex        =   33
               Top             =   3000
               Width           =   315
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Modelo"
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
               Left            =   210
               TabIndex        =   32
               Top             =   2310
               Width           =   630
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Marca"
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
               Left            =   210
               TabIndex        =   31
               Top             =   1650
               Width           =   510
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Ne de Série"
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
               Left            =   210
               TabIndex        =   30
               Top             =   990
               Width           =   900
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Equipamento"
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
               Left            =   210
               TabIndex        =   29
               Top             =   360
               Width           =   1050
            End
         End
         Begin VB.Frame frmAcessorios 
            Caption         =   "Acessórios"
            Height          =   4515
            Left            =   5370
            TabIndex        =   19
            Top             =   480
            Width           =   6705
            Begin VB.CommandButton cmdExcItemAc 
               Height          =   405
               Left            =   6060
               Picture         =   "frmOs.frx":F1A1
               Style           =   1  'Graphical
               TabIndex        =   102
               Top             =   540
               Width           =   495
            End
            Begin VB.CommandButton cmdInserir 
               Height          =   405
               Left            =   5550
               Picture         =   "frmOs.frx":F6EC
               Style           =   1  'Graphical
               TabIndex        =   101
               Top             =   540
               Width           =   495
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Próximo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   5370
               TabIndex        =   84
               Top             =   3960
               Width           =   1215
            End
            Begin VB.ComboBox cbAcessorio 
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
               Left            =   150
               TabIndex        =   21
               Top             =   570
               Width           =   5355
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "frmOs.frx":FCB5
               Height          =   2325
               Left            =   150
               TabIndex        =   20
               Top             =   1020
               Width           =   6405
               _ExtentX        =   11298
               _ExtentY        =   4101
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   "id"
                  Caption         =   "Id"
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
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     Alignment       =   2
                     ColumnWidth     =   569,764
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   5235,024
                  EndProperty
               EndProperty
            End
            Begin MSAdodcLib.Adodc DataAdoAcessorio 
               Height          =   405
               Left            =   180
               Top             =   3420
               Width           =   6405
               _ExtentX        =   11298
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
            Begin VB.Label Label10 
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
               Left            =   150
               TabIndex        =   22
               Top             =   360
               Width           =   795
            End
         End
         Begin VB.TextBox ctDefeito 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2580
            Left            =   -74820
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   690
            Width           =   11805
         End
         Begin VB.TextBox ctDiagnostico 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1980
            Left            =   -74820
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   690
            Width           =   11805
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
            Left            =   -74790
            TabIndex        =   15
            Top             =   660
            Width           =   5115
         End
         Begin VB.TextBox ctqtdPecas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   -67110
            MaxLength       =   8
            TabIndex        =   14
            Top             =   630
            Width           =   1005
         End
         Begin MSDataGridLib.DataGrid DgPecas 
            Bindings        =   "frmOs.frx":FCD4
            Height          =   3285
            Left            =   -74790
            TabIndex        =   16
            Top             =   1080
            Width           =   11805
            _ExtentX        =   20823
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
            BeginProperty Column02 
               DataField       =   "numserie"
               Caption         =   "N° Série"
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
               DataField       =   "qtd"
               Caption         =   "Qtd"
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
               DataField       =   "valunit"
               Caption         =   "V. Unit"
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
            BeginProperty Column05 
               DataField       =   "valtotal"
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
                  ColumnWidth     =   750,047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   4605,166
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   1844,787
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   929,764
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc DataAdoPecas 
            Height          =   465
            Left            =   -74760
            Top             =   4470
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   820
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
            Caption         =   "Peças utilizadas"
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
         Begin MSAdodcLib.Adodc dataAdoServExec 
            Height          =   465
            Left            =   -74760
            Top             =   4560
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   820
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
            Caption         =   "Serviços executados"
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
         Begin MSAdodcLib.Adodc dataAdoHist 
            Height          =   465
            Left            =   -74790
            Top             =   4530
            Width           =   11835
            _ExtentX        =   20876
            _ExtentY        =   820
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
            Caption         =   "Histórico"
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
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentário"
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
            Left            =   -74820
            TabIndex        =   138
            Top             =   2670
            Width           =   1065
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "N° de Série"
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
            Left            =   -69690
            TabIndex        =   134
            Top             =   420
            Width           =   885
         End
         Begin VB.Label Label33 
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
            Left            =   -71790
            TabIndex        =   118
            Top             =   4560
            Width           =   660
         End
         Begin VB.Label Label17 
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
            Left            =   -71820
            TabIndex        =   113
            Top             =   2340
            Width           =   660
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
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
            Left            =   -74820
            TabIndex        =   38
            Top             =   420
            Width           =   915
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
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
            Left            =   -74820
            TabIndex        =   37
            Top             =   420
            Width           =   915
         End
         Begin VB.Label Label13 
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
            Left            =   -74790
            TabIndex        =   36
            Top             =   450
            Width           =   795
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
            Left            =   -67110
            TabIndex        =   35
            Top             =   420
            Width           =   930
         End
      End
      Begin MSDataGridLib.DataGrid DataGridUvt 
         Bindings        =   "frmOs.frx":FCEF
         Height          =   2505
         Left            =   -74880
         TabIndex        =   141
         Top             =   1170
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   4419
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
            Caption         =   "Id"
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
            DataField       =   "Status"
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
         BeginProperty Column02 
            DataField       =   "DataIni"
            Caption         =   "Abertura"
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
         BeginProperty Column04 
            DataField       =   "QtdDias"
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
               ColumnWidth     =   569,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1635,024
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   959,811
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoDataUvt 
         Height          =   405
         Left            =   -74880
         Top             =   3720
         Width           =   6855
         _ExtentX        =   12091
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
         Left            =   -72780
         TabIndex        =   142
         Top             =   720
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
         Left            =   -71340
         TabIndex        =   148
         Top             =   720
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
         Left            =   -71310
         TabIndex        =   149
         Top             =   480
         Width           =   1395
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
         Left            =   -74880
         TabIndex        =   145
         Top             =   510
         Width           =   795
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
         Left            =   -72750
         TabIndex        =   143
         Top             =   480
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   1560
      TabIndex        =   1
      Top             =   -60
      Width           =   12435
      Begin VB.TextBox ctIdWeb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   135
         Top             =   420
         Width           =   1305
      End
      Begin VB.CommandButton cmdImportarOs 
         Height          =   765
         Left            =   8040
         Picture         =   "frmOs.frx":FD08
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Autorizar Orde de Serviço"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Height          =   765
         Left            =   9750
         Picture         =   "frmOs.frx":10509
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Autorizar Orde de Serviço"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdAutorizar 
         Height          =   765
         Left            =   8910
         Picture         =   "frmOs.frx":10C8E
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Autorizar Orde de Serviço"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdOrcamento 
         Height          =   765
         Left            =   10620
         Picture         =   "frmOs.frx":11371
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Imprimir Orçamento"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdImpOs 
         Height          =   765
         Left            =   11490
         Picture         =   "frmOs.frx":11CA2
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Imprimir Ordem de Serviço"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdDtPrev 
         Height          =   375
         Left            =   6960
         Picture         =   "frmOs.frx":125D6
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   480
         Width           =   465
      End
      Begin VB.CommandButton cmdDtEntrada 
         Height          =   375
         Left            =   4860
         Picture         =   "frmOs.frx":12AE9
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   480
         Width           =   465
      End
      Begin VB.Frame Frame3 
         Caption         =   "Andamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   795
         Left            =   6720
         TabIndex        =   10
         Top             =   900
         Width           =   5655
         Begin VB.CommandButton cmdAtAndamento 
            Height          =   435
            Left            =   5010
            Picture         =   "frmOs.frx":12FFC
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   240
            Width           =   525
         End
         Begin VB.ComboBox Cbsituacao 
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
            Left            =   120
            TabIndex        =   11
            Top             =   270
            Width           =   4845
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   6585
         Begin VB.OptionButton OptCancelado 
            Caption         =   "Cancelado"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   5190
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptConcluido 
            Caption         =   "Concluido"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2760
            TabIndex        =   8
            Top             =   360
            Width           =   1245
         End
         Begin VB.OptionButton OptFaturado 
            Caption         =   "Faturado"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   4020
            TabIndex        =   7
            Top             =   360
            Width           =   1155
         End
         Begin VB.OptionButton OptAutorizado 
            Caption         =   "Autorizado"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   360
            Width           =   1275
         End
         Begin VB.OptionButton OptSolicitado 
            Caption         =   "Solicitado"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   150
            TabIndex        =   5
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.TextBox ctNos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         MaxLength       =   8
         TabIndex        =   2
         Top             =   420
         Width           =   1305
      End
      Begin MSMask.MaskEdBox maskDataEnt 
         Height          =   375
         Left            =   3420
         TabIndex        =   74
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
      Begin MSMask.MaskEdBox maskDataPrev 
         Height          =   375
         Left            =   5520
         TabIndex        =   76
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
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Web"
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
         Left            =   1470
         TabIndex        =   136
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Prev. conlusão"
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
         Left            =   5550
         TabIndex        =   77
         Top             =   240
         Width           =   1185
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
         Left            =   3450
         TabIndex        =   75
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° OS"
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
         Left            =   120
         TabIndex        =   3
         Top             =   150
         Width           =   645
      End
   End
   Begin VertMenu.VerticalMenu VmnuItem 
      Height          =   7425
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   13097
      MenusMax        =   2
      MenuCaption1    =   "Arquivo"
      MenuItemsMax1   =   7
      MenuItemIcon11  =   "frmOs.frx":135C5
      MenuItemCaption11=   "Novo"
      MenuItemKey11   =   "mnuNovo"
      MenuItemIcon12  =   "frmOs.frx":14317
      MenuItemCaption12=   "Consultar"
      MenuItemKey12   =   "mnuConsultar"
      MenuItemIcon13  =   "frmOs.frx":15069
      MenuItemCaption13=   "Atualizar"
      MenuItemKey13   =   "mnuAlterar"
      MenuItemIcon14  =   "frmOs.frx":15DBB
      MenuItemCaption14=   "Excluir"
      MenuItemKey14   =   "mnuExcluir"
      MenuItemIcon15  =   "frmOs.frx":16B0D
      MenuItemCaption15=   "Fechar"
      MenuItemKey15   =   "mnuFechar"
      MenuItemIcon16  =   "frmOs.frx":1785F
      MenuItemCaption16=   "Cancelar"
      MenuItemKey16   =   "mnuCancelar"
      MenuItemIcon17  =   "frmOs.frx":185B1
      MenuItemCaption17=   "Sair"
      MenuItemKey17   =   "mnuSair"
      MenuCaption2    =   "Relatórios"
      MenuItemsMax2   =   3
      MenuItemIcon21  =   "frmOs.frx":19303
      MenuItemCaption21=   "Imprimir Os"
      MenuItemKey21   =   "mnuImpOs"
      MenuItemIcon22  =   "frmOs.frx":1A055
      MenuItemCaption22=   "Imprimir Orçam."
      MenuItemKey22   =   "mnuImpOrcam"
      MenuItemIcon23  =   "frmOs.frx":1ADA7
      MenuItemCaption23=   "Assist. de Rel."
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuDigDoc 
         Caption         =   "Digitalizar Documentos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuVerifUVT 
         Caption         =   "Verificar Chamado UVT"
      End
      Begin VB.Menu mnuConCnpj 
         Caption         =   "Consultar CNPJ"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuControlUVT 
         Caption         =   "Controle de Chamados UVT"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuRelSolLacra 
         Caption         =   "Solicitação de Lacração"
      End
   End
End
Attribute VB_Name = "frmOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)


Private Sub BtAlterar_Click()
Dim Sql As String
Mcod = ctcodClient.Text

Sql = "UPDATE clientes SET nome='" & ctrazao & "', email='" & ctEmail & "', endereco='" & ctend & "', bairro='" & ctbairro & "'," & _
"cidade='" & ctcidade & "', estado='" & ctuf & "',cep='" & MASKCEP & "', cnpj='" & ctcgc & "', insc='" & ctinsc & "', telefone='" & MASKFONE & "'," & _
"celular='" & MASKCEL & "', roteiro='" & ctroteiro & "', responsavel='" & ctcontato & "' Where id = " & Mcod & ""

conn.Execute Sql

'Call sbLimpa_Campos(Me)
BtConsultar.Enabled = True
BtAlterar.Enabled = False
MsgBox "Dados atualizados com sucesso.", vbInformation, "Assistente de cadastro"

End Sub

Private Sub BtConsultar_Click()
Set rsTabelas = New ADODB.Recordset
If LTNOME.Visible = False Then
    LTNOME.Clear
    LTNOME.Visible = True
Else
    LTNOME.Visible = False
End If

If ctrazao.Text = "" Then
    rsTabelas.Open "select * from clientes order by nome", conn
Else
    rsTabelas.Open "select * from clientes where nome like '" & ctrazao & "%' order by nome", conn
End If

Do While Not rsTabelas.EOF
    LTNOME.AddItem rsTabelas!nome
    rsTabelas.MoveNext
Loop

rsTabelas.Close
Set rsTabelas = Nothing
End Sub

Private Sub BtNovo_Click()
If BtNovo.Caption = "Cliente Novo" Then
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select MAX(id) as nCountId from clientes order by id", conn
    ctcodClient = rsTabelas!ncountId + 1
    rsTabelas.Close
    Set rsTabelas = Nothing
    BtNovo.Caption = "Cliente Salvar"
    ctrazao.SetFocus
Else

    If ctrazao.Text = "" Then
        MsgBox "Preencha todos os dados nos campos para continuar.", vbCritical
        Exit Sub
    End If
    
    Sql = "INSERT INTO clientes (nome, email, endereco, bairro, cidade, estado, cep, cnpj, insc, telefone, celular," & _
        "roteiro, responsavel, titulo) VALUES ('" & ctrazao & "', '" & ctEmail & "', '" & ctend & "', '" & ctbairro & "'," & _
        "'" & ctcidade & "', '" & ctuf & "', '" & MASKCEP & "', '" & ctcgc & "', '" & ctinsc & "', '" & MASKFONE & "'," & _
        "'" & MASKCEL & "', '" & ctroteiro & "', '" & ctcontato & "','" & "CLIENTE" & "')"
        conn.Execute Sql
        
        MsgBox "Cliente cadastrado com sucesso.", vbInformation, "Assistente de verificação."
    BtNovo.Enabled = False
    SSTab2.Tab = 1
    SSTab1.Tab = 0
    BtNovo.Caption = "Cliente Novo"
    ctEquipamento.SetFocus
End If
End Sub

Private Sub CbAndamento_Change()

End Sub

Private Sub CbChamado_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If CbChamado.Text = "" Then
        MsgBox "Selecione uma opção p/ o lançamento.", vbCritical
        Exit Sub
    End If
    MaskDataIni.SetFocus
End Select
End Sub

Private Sub cbcolaborador_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If cbcolaborador.Text = "" Then
        MsgBox "Selecione um colaborador para continuar.", vbCritical
        Exit Sub
    End If
    
    MASKCEP.SetFocus
End Select
End Sub

Private Sub cbFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If cbFormaPag.Text = "" Then
        MsgBox "Selecione uma forma de pagamento para continuar.", vbCritical
        Exit Sub
    End If
    ctDesconto.Enabled = True
    ctDesconto.SetFocus
    
End Select
End Sub

Private Sub cbnserie_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If cbnserie.Text <> "" Then
        ctqtdPecas = 1
        ctqtdPecas.Enabled = False
        cmdInsertPecas.SetFocus
    Else
        ctqtdPecas.Enabled = True
        ctqtdPecas.SetFocus
    End If
    
End Select
End Sub

Private Sub cbPecas_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    Dim nIDpeca As Variant
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadpecas where Descricao='" & cbPecas & "'", conn
    If rsTabelas.EOF Then
        nIDpeca = 0
    Else
        nIDpeca = rsTabelas!ID
    End If
    rsTabelas.Close
    
    cbnserie.Clear
    
    If nIDpeca <> 0 Then
    rsTabelas.Open "select * from tbnserie where idpeca=" & nIDpeca & " and status='" & "LIVRE" & "' order by nserie ASC", conn
    Do While Not rsTabelas.EOF
        cbnserie.AddItem rsTabelas!nserie
        rsTabelas.MoveNext
    Loop
     rsTabelas.Close
     End If
     
     Set rsTabelas = Nothing

    cbnserie.SetFocus
End Select
End Sub

Private Sub cbServExec_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If cbServExec.Text = "" Then
        MsgBox "Selecione uma opção para continuar."
        Exit Sub
    End If
    ctqtdServExec.SetFocus

End Select
End Sub

Private Sub cmdAbrirNF_Click()
Shell "C:\Arquivos de programas\Adobe\Reader 11.0\Reader\AcroRd32.exe \\Vmsuportek-srv\Publico\NFCompraClie\" & ctTombo, vbMaximizedFocus
End Sub

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

Private Sub cmdAtAndamento_Click()
Sql = "UPDATE osdb SET situacao='" & Cbsituacao & "' where id=" & ctNos & ""
conn.Execute Sql



Sql = "INSERT INTO LOGEVENTOS (idcolaborador, data, cnpj, n_os, eventos, servico) VALUES " & _
"(" & ctcodClient & ",'" & Date & " - " & Format(Time, "hh:mm") & "', '" & ctcgc & "', " & ctNos & ", '" & Cbsituacao & "', " & _
"'" & cbServSolic & "')"

conn.Execute Sql

Call sbAtDataHist
End Sub

Private Sub cmdAutorizar_Click()
Call sbAtStatus(ctNos, 2)

End Sub

Private Sub cmdCadPecas_Click()

End Sub

Private Sub cmdDtEntrada_Click()
On Error Resume Next
frmCalendario.Show (1)
maskDataEnt = strData
End Sub

Private Sub cmdDtPrev_Click()
frmCalendario.Show (1)
maskDataPrev = strData
End Sub

Private Sub cmdExcAndamentoHist_Click()
Dim nID As Integer
On Error GoTo E:

nID = InputBox("Digite o ID da peça para continuar.")

Sql = "DELETE FROM logeventos where id=" & nID & " and n_os=" & ctNos & ""

conn.Execute Sql

E:
If Err.Number = 13 Then
    MsgBox "ID inválido, digite novamente para continuar."
    Exit Sub
End If

Call sbAtDataHist
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

Private Sub cmdExcItemAc_Click()
Dim nIdac As Integer
On Error GoTo E:

nIdac = InputBox("Digite o ID do ítem para continuar.")

Sql = "DELETE from osdbacessorios where id=" & nIdac & ""

conn.Execute Sql

Call sbAtDataAdoAcessorio

E:
If Err.Number = 13 Then
    MsgBox "Digite o ID do ítem para continuar."
    Exit Sub
End If
End Sub

Private Sub cmdExcPecas_Click()
Dim nID As Integer
On Error GoTo E:

nID = InputBox("Digite o ID da peça para continuar.")

Sql = "DELETE FROM osdbpecas where id=" & nID & " and n_os=" & ctNos & ""

conn.Execute Sql

E:
If Err.Number = 13 Then
    MsgBox "ID inválido, digite novamente para continuar."
    Exit Sub
End If

Call sbAtDataPecas
Call sbTotalOS(ctNos)

End Sub

Private Sub cmdexcServExec_Click()
Dim nID As Integer
On Error GoTo E:

nID = InputBox("Digite o ID do serviço para continuar.")

Sql = "DELETE FROM osdbservexec where id=" & nID & " and n_os=" & ctNos & ""

conn.Execute Sql

Call sbAtDataServExec
Call sbTotalOS(ctNos)

E:
If Err.Number = 13 Then
    MsgBox "Id inválido. Digite corretamente o ID para continuar."
    Exit Sub
End If
End Sub

Private Sub cmdImportarCad_Click()
On Error Resume Next
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from clientes where cnpj='" & ctcgc & "'", conn
If rsTabelas.EOF Then
    rsTabelas.Close
    rsTabelas.Open "select MAX(id) as nCountId from clientes order by id", conn
    ctcodClient = rsTabelas!ncountId + 1
    rsTabelas.Close
Else
    ctcodClient = rsTabelas!ID
    ctrazao = rsTabelas!nome
    ctend = rsTabelas!endereco
    ctbairro = rsTabelas!bairro
    ctcidade = rsTabelas!cidade
    ctuf = rsTabelas!estado
    MASKCEP = rsTabelas!CEP
    ctcgc = rsTabelas!CNPJ
    ctinsc = rsTabelas!insc
    MASKFONE = rsTabelas!Telefone
    MASKCEL = rsTabelas!celular
    ctEmail = rsTabelas!email
    ctroteiro = rsTabelas!roteiro
    ctcontato = rsTabelas!responsavel
    rsTabelas.Close
End If

Set rsTabelas = Nothing
BtNovo.Caption = "Cliente Salvar"
BtNovo.Enabled = True
cmdImportarCad.Enabled = False
ctrazao.SetFocus
End Sub

Private Sub cmdImportarOs_Click()
Dim Sql As String
Set rsTabelas = New ADODB.Recordset

Call Conecta_BDRemoto
rsTabelas.Open "select * from osdb where situacao='" & "Solicitacao em analise" & "' order by id", connRemoto
'On Error Resume Next

Do While Not rsTabelas.EOF

    Sql = "INSERT INTO osdb(empresa,cnpj,insc,endereco,bairro,cidade,estado,cep,email,servico,situacao, " & _
    "equipamento,responsavel,news,idWeb,idColaborador,data) VALUE ('" & rsTabelas!empresa & "', " & _
    "'" & rsTabelas!CNPJ & "', '" & rsTabelas!insc & "', '" & rsTabelas!endereco & "', '" & rsTabelas!bairro & "', " & _
    "'" & rsTabelas!cidade & "', '" & rsTabelas!estado & "', '" & rsTabelas!CEP & "', '" & rsTabelas!email & "', " & _
    "'" & rsTabelas!servico & "', '" & rsTabelas!situacao & "', '" & rsTabelas!equipamento & "', " & _
    "'" & rsTabelas!responsavel & "', " & _
    "'" & rsTabelas!news & "', " & rsTabelas!ID & ", " & rsTabelas!idcolaborador & ", '" & rsTabelas!Data & "')"
    
    conn.Execute Sql
    
rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing

Call Desconecta_BDRemoto

frmLocOs.Show (1)

'Set rsTabelas = New ADODB.Recordset
'rsTabelas.Open "select  MAX(id) as nCountId from osdb where situacao='" & "Solicitacao em analise" & "'", conn
'MsgBox "Você possui " & rsTabelas!nCountId & " Novas Os em analise.", vbInformation

End Sub

Private Sub cmdImpOs_Click()
If ctNos.Text = "" Then
    MsgBox "Selelcione uma Ordem de serviço para continuar."
    Exit Sub
End If

crp.Connect = strDns
crp.ReportFileName = App.Path & "\..\relatorios\OS.rpt"
crp.SelectionFormula = "{osdb.id}=" & ctNos & ""
crp.Destination = crptToWindow
crp.WindowState = crptMaximized
crp.Action = 1
End Sub

Private Sub cmdInserir_Click()
Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select * from cadacessorios where descricao='" & cbAcessorio & "'", conn
If rsTabelas.EOF Then
    Sql = "INSERT INTO cadacessorios (descricao) VALUES ('" & cbAcessorio & "')"
    conn.Execute Sql
End If

Sql = "INSERT INTO osdbacessorios (descricao, n_os) VALUES ('" & cbAcessorio & "', " & ctNos & ")"
conn.Execute Sql

Call sbAddAcessorios
Call sbAtDataAdoAcessorio
End Sub

Private Sub cmdInsertAndamentoHist_Click()
Sql = "UPDATE osdb SET situacao='" & cbSituacaoHist & "' where id=" & ctNos & ""
conn.Execute Sql

Sql = "INSERT INTO LOGEVENTOS (idcolaborador, data, cnpj, n_os, eventos, servico) VALUES " & _
"(" & ctcodClient & ",'" & Date & " - " & Format(Time, "hh:mm") & "', '" & ctcgc & "', " & ctNos & ", '" & cbSituacaoHist & "', " & _
"'" & cbServSolic & "')"

conn.Execute Sql

Call sbAtDataHist

End Sub

Private Sub cmdInsertPecas_Click()
Dim nQtdpecas As Integer
Dim nValunit As Double
Dim nTotal As Double
Dim strRef As String
Dim nn As String

nQtdpecas = ctqtdPecas

'***************************Consulta valor do peças************************
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from cadpecas where descricao='" & cbPecas & "'", conn
If rsTabelas.EOF Then
    Exit Sub
End If

strRef = rsTabelas!ref

nValunit = Format(rsTabelas!valunit, "##,##0.00")

nTotal = Format(nValunit * nQtdpecas, "##,##0.00")

rsTabelas.Close
Set rsTabelas = Nothing

If cbnserie.Text = "" Then
    cbnserie = 0
End If


Sql = "INSERT INTO osdbpecas (descricao, ref, valunit, qtd, valtotal, n_os,numserie) VALUES ('" & cbPecas & "', '" & strRef & "', " & _
"" & converte(nValunit) & ", " & nQtdpecas & ", " & converte(nValunit) & ", " & ctNos & "," & cbnserie & ")"

conn.Execute Sql

If cbnserie.Text <> "" Then
    Sql = "UPDATE tbnserie set status='" & "USADO" & "', nos=" & ctNos & " where nserie=" & cbnserie & ""
End If

conn.Execute Sql


cbPecas.Text = ""
ctqtdPecas.Text = ""
cbnserie.Text = ""

Call sbAtDataPecas
Call sbTotalOS(ctNos)

End Sub

Private Sub cmdinsertServExec_Click()
Dim nQtdServExec As Integer
Dim nValunit As Double
Dim nTotal As Double

nQtdServExec = ctqtdServExec

'***************************Consulta valor do serviço************************
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from cadservexec where descricao='" & cbServExec & "'", conn
If rsTabelas.EOF Then
    Exit Sub
End If
nValunit = rsTabelas!valunit

rsTabelas.Close
Set rsTabelas = Nothing

nTotal = nValunit * nQtdServExec

'***************************Inclui serviço em osdbservexec ********************
Sql = "INSERT INTO osdbservexec ( descricao, n_os, valunit, qtd, valtotal) VALUES ('" & cbServExec & "', " & _
"" & ctNos & ", " & nValunit & ", " & nQtdServExec & ", " & nTotal & ")"

conn.Execute Sql

Call sbAtDataServExec
Call sbTotalOS(ctNos)
ctqtdServExec.Text = ""
cbServExec.SetFocus

End Sub

Private Sub cmdNfCompra_Click()
10000   Dim strOrigem As String
10005   Dim strDestino As String
10010   Dim strNumNF As String

10015

10020   cdiagBox.ShowOpen
10025   strOrigem = cdiagBox.FileName
strNumNF = ctNos & "-" & cdiagBox.FileTitle
      strDestino = "\\Vmsuportek-srv\Publico\NFCompraClie\" & strNumNF
      
10030   ctTombo = strNumNF
        FileCopy strOrigem, strDestino
10035   MsgBox "NF de compra importada com sucesso."

End Sub

Private Sub cmdOrcamento_Click()
If ctNos.Text = "" Then
    MsgBox "Selelcione uma Ordem de serviço para continuar."
    Exit Sub
End If

crp.Connect = strDns
crp.ReportFileName = App.Path & "\..\relatorios\ORÇAMENTO.rpt"
crp.SelectionFormula = "{osdb.id}=" & ctNos & ""
crp.Destination = crptToWindow
crp.WindowState = crptMaximized
crp.Action = 1

End Sub

Private Sub cmdProximo_Click()
SSTab2.Tab = 1
SSTab1.Tab = 0
ctEquipamento.SetFocus
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub cmdReceber_Click()
Dim Ret As Integer

Sql = "UPDATE osdb SET desconto=" & converte(ctDesconto) & ", totalpago=" & converte(ctTotalPago) & ", " & _
"formapg='" & cbFormaPag & "' where id=" & ctNos & ""

conn.Execute Sql

Sql = "INSERT INTO caixa (op,data,doc, tipo, descricao, debito, credito,saldo) VALUES ('" & MOP & "', '" & Format$(Date, "yyyy-mm-dd hh:mm:ss") & "', " & _
    "'" & "1.0.01" & "','" & "CREDITO" & "', '" & "Recebimento de Os n° " & ctNos & "', " & converte(0) & ", " & converte(ctTotalPago) & ", " & _
    "" & converte(0) & ")"
    
    conn.Execute Sql


Ret = MsgBox("Deseja imprimir o recibo?", vbYesNo)
If Ret = 6 Then
    If ctNos.Text = "" Then
        MsgBox "Selelcione uma Ordem de serviço para continuar."
    Exit Sub
    End If
    
    crp.Connect = strDns
    crp.ReportFileName = App.Path & "\..\relatorios\RECIBO.rpt"
    crp.SelectionFormula = "{osdb.id}=" & ctNos & ""
    crp.Destination = crptToWindow
    crp.WindowState = crptMaximized
    crp.Action = 1
End If

Call sbAtStatus(ctNos, 3)

End Sub

Private Sub cmdRecibo_Click()
If ctNos.Text = "" Then
    MsgBox "Selelcione uma Ordem de serviço para continuar."
Exit Sub
End If

crp.Connect = strDns
crp.ReportFileName = App.Path & "\..\relatorios\RECIBO.rpt"
crp.SelectionFormula = "{osdb.id}=" & ctNos & ""
crp.Destination = crptToWindow
crp.WindowState = crptMaximized
crp.Action = 1
End Sub

Private Sub Command6_Click()
SSTab1.Tab = 1
ctDefeito.SetFocus
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()
'Sql = "UPDATE osdb SET status='" & "3" & "' where id=" & ctNos & ""

'conn.Execute Sql
Call sbAtStatus(ctNos, 3)
'Call fStatusOs(ctNos)
End Sub

Private Sub CTBAIRRO_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctcidade.SetFocus
End Select
End Sub

Private Sub ctBairro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctcgc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctcgc.Text = "" Then
        MsgBox "Digite o CNPJ ou CPF para continuar.", vbCritical
        Exit Sub
    End If
    ctinsc.SetFocus
End Select
End Sub

Private Sub CTCGC_LostFocus()
If Len(ctcgc.Text) > 0 Then
      Select Case Len(ctcgc.Text)
       Case Is = 11
         ctcgc = Format(ctcgc, "@@@.@@@.@@@-@@")
'         If Not calculacpf(CTCGC.Text) Then
'            MsgBox "CPF com DV incorreto !!!"
'            CTCGC = ""
'            CTCGC.Mask = "##############"
'            CTCGC.SetFocus
'         End If
       Case Is = 14
        ctcgc = Format(ctcgc, "@@.@@@.@@@/@@@@-@@")
         'CTCGC.Mask = "##.###.###/####-##"
'         If Not ValidaCGC(CTCGC.Text) Then
'            MsgBox "CGC com DV incorreto !!! "
'            CTCGC = ""
'            CTCGC.Mask = "##############"
'            CTCGC.SetFocus
'         End If
      End Select
    End If
End Sub

Private Sub CTCIDADE_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctuf.SetFocus
End Select
End Sub

Private Sub CTCIDADE_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctcontato_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctcontato.Text = "" Then
        MsgBox "Digite o nome do Responsável para continuar.", vbCritical
        Exit Sub
    End If
    ctroteiro.SetFocus
End Select
End Sub

Private Sub ctContato_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctDefeito_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    Dim nvalapagar As Currency
    Dim nvaldesconto As Currency
    If ctDesconto.Text = "" Then
        ctDesconto = 0
    End If
    nvalapagar = ctTotalOS
    nvaldesconto = ctDesconto
    
    ctDesconto = Format(nvaldesconto, "##,##0.00")
    ctTotalPago = Format(nvalapagar - nvaldesconto, "##,##0.00")
    
End Select
End Sub

Private Sub ctDiagnostico_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctEmail_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctcontato.SetFocus
End Select
End Sub

Private Sub ctend_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctbairro.SetFocus
End Select
End Sub

Private Sub CTEND_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctEquipamento_Click()
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from equipamentos where descricao='" & ctEquipamento & "'", conn
If Not rsTabelas.EOF Then
    ctmarca = rsTabelas!marca
    ctmodelo = rsTabelas!modelo
End If
rsTabelas.Close
Set rsTabelas = Nothing
End Sub

Private Sub ctEquipamento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctNSerie.SetFocus
End Select
End Sub

Private Sub ctinsc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctinsc.Text = "" Then
        MsgBox "Digite a INSC Estadual ou RG para continuar.", vbCritical
        Exit Sub
    End If
    ctEmail.SetFocus
End Select
End Sub

Private Sub ctmarca_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctmodelo.SetFocus
End Select
End Sub

Private Sub ctmodelo_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctRef.SetFocus
End Select
End Sub

Private Sub ctNSerie_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If FsVerifDupNS(ctNSerie) = 1 Then
        MsgBox "Número de série ja exitente."
        Exit Sub
    End If
    ctRef.SetFocus
End Select
End Sub

Private Sub ctNSerie_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctqtdPecas_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctqtdPecas.Text = "" Then
        MsgBox "digite a quantidade para continuar.", vbCritical
        Exit Sub
    End If
    
    cmdInsertPecas.SetFocus
End Select

End Sub

Private Sub ctrazao_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctrazao.Text = "" Then
        MsgBox "Digite o nome ou a razão social da Empresa para continuar.", vbCritical
        Exit Sub
    End If
    ctrazao = UCase(ctrazao)
    
    maskDataNasc.SetFocus
End Select
End Sub

Private Sub ctrazao_KeyPress(KeyAscii As Integer)
If LTNOME.Visible = False Then
    LTNOME.Clear
    LTNOME.Visible = True
Else
    LTNOME.Visible = False
End If

If ctrazao.Text = "" Then
    rsTabelas.Open "select * from clientes order by nome ASC", conn
Else
    rsTabelas.Open "select * from clientes where nome like '" & ctrazao & "%' order by nome ASC", conn
End If

Do While Not rsTabelas.EOF
    LTNOME.AddItem rsTabelas!nome
    rsTabelas.MoveNext
Loop

rsTabelas.Close
Set rsTabelas = Nothing
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ctRef_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctTombo.SetFocus
End Select
End Sub

Private Sub ctroteiro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CTUF_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctuf = UCase(ctuf)
    MASKFONE.SetFocus
End Select
End Sub

Private Sub CTUF_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    If ctNos.Text = "" Then
        Unload Me
    Else
        Unload Me
        Call AtivaForm(frmOs)
    End If
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
SSTab1.Tab = 0
SSTab2.Tab = 0
Call sbAddServico
Call sbAddEquipamento
Call sbAddServico
Call sbAddSituacao
Call sbAddColaborador
Call sbAddPecas
Call sbAddAcessorios
Call sbAddServExec
Call sbAtDataAdoAcessorio
Call sbAtDataHist
Call sbAtDataServExec
Call sbAtDataPecas
Call sbAtDataUvt
End Sub

Private Sub Lt_cep_Click()
'txtEscolha.Text = str(Lt_cep.ListIndex) & " = " & Lt_cep.List(lLt_cep.ListIndex)
End Sub

Private Sub LTNOME_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn:
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from clientes where nome='" & LTNOME.List(LTNOME.ListIndex) & "' order by nome", conn
    
    On Error Resume Next
    
    On Error Resume Next
    ctcodClient = rsTabelas!ID
    ctrazao = rsTabelas!nome
    ctend = rsTabelas!endereco
    ctbairro = rsTabelas!bairro
    ctcidade = rsTabelas!cidade
    ctuf = rsTabelas!estado
    MASKCEP = rsTabelas!CEP
    ctcgc = rsTabelas!CNPJ
    ctinsc = rsTabelas!insc
    MASKFONE = rsTabelas!Telefone
    MASKCEL = rsTabelas!celular
    ctEmail = rsTabelas!email
    ctroteiro = rsTabelas!roteiro
    ctcontato = rsTabelas!responsavel
       
    BtNovo.Enabled = False
    BtAlterar.Enabled = True
    BtConsultar.Enabled = True
    BTEXCLUIR.Enabled = True
    LTNOME.Clear
    LTNOME.Visible = False
    rsTabelas.Close
    Set rsTabelas = Nothing
    
    BtAlterar.Enabled = True
End Select
End Sub

Private Sub MASKCEL_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    ctcgc.SetFocus
End Select
End Sub

Private Sub MASKCEP_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    Call locCep
    MASKCEP = MASKCEP
    ctend.SetFocus
End Select
End Sub

Private Sub MASKCEP_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789" & vbKeyDelete & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub maskDataEnt_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    
End Select
End Sub

Private Sub maskdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    cmdAddChamado.SetFocus
End Select
End Sub

Private Sub maskDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    MaskDataFin = DateAdd("d", 14, MaskDataIni)
    MaskDataFin.SetFocus
End Select
End Sub

Private Sub maskDataNasc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    cbcolaborador.SetFocus
End Select
End Sub

Private Sub MASKFONE_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    
    MASKCEL.SetFocus
End Select
End Sub

Private Sub mnuConCnpj_Click()
Call AtivaForm(frmConCnpj)
MASKFONE.SetFocus

End Sub

Private Sub mnuControlUVT_Click()
Call AtivaForm(frmUVT)
End Sub

Private Sub mnuDigDoc_Click()

Shell "C:\Arquivos de programas\HP\HP Deskjet 3050 J610 series\Bin\HP Deskjet 3050 J610 series.exe -Start UDCDevicePage"
End Sub

Private Sub mnuVerifUVT_Click()
Call sbAtUvt
Call sbAtrasoUVT
crp.Connect = strDns
crp.ReportFileName = App.Path & "\..\relatorios\RelAtrasoUVT.rpt"
'crp.SelectionFormula = "{osdb.id}=" & ctNos & ""
crp.Destination = crptToWindow
crp.WindowState = crptMaximized
crp.Action = 1
End Sub

Private Sub VmnuItem_MenuItemClick(MenuNumber As Long, MenuItem As Long)
Select Case MenuNumber
Case 1
    Select Case MenuItem
    Case 1
        If VmnuItem.MenuItemCaption = "Novo" Then
            VmnuItem.MenuItemCaption = "Salvar"
            Call sbNovaOs
            Cbsituacao.Text = "Solicitacao em analise"
            BtNovo.Enabled = True
            BtConsultar.Enabled = True
            
        Else
            If cbcolaborador.Text = "" Then
                MsgBox "Selecione um colaborador para continuar.", vbCritical
                Exit Sub
            End If
        
            Call sbsalvarDadosOS
            Call sbLimpa_Campos(Me)
            VmnuItem.MenuItemCaption = "Novo"
                
        End If
    Case 2
        SSTab1.Tab = 0
        SSTab2.Tab = 0
        Call sbLimpa_Campos(Me)
        'frmLocOs.Show (1)
        Call AtivaForm(frmLocOs)
        If ctNos.Text = "" Then Exit Sub
        Call sbAtDataHist
        Call sbAtDataAdoAcessorio
        Call sbAtDataServExec
        Call sbAtDataPecas
        'Call sbSumServicos(ctNos)
        'Call sbSumPecas(ctNos)
        Call sbTotalOS(ctNos)
        Call sbAtDataUvt
        Call sbAtChamadoUvt(ctNos)
    Case 3
        If ctIdWeb.Text = "" Then
            ctIdWeb = 0
        End If
        Call sbAtdadosOs
        MsgBox "Dados atualizados com sucesso.", vbInformation
        Call sbLimpa_Campos(Me)
        SSTab1.Tab = 0
        SSTab2.Tab = 0
    Case 4
        Call sbExcOs
        
    Case 5
        Call sbAtStatus(ctNos, 4)
        
    Case 6
        Call sbAtStatus(ctNos, 5)
    Case 7
        Unload Me
    End Select
Case 2
    Select Case MenuItem
    Case 1
       
    Case 2
    
    Case 3
        Call AtivaForm(frmRel)
        
    Case 4
    
    Case 5
        
    End Select

Case 3
    Select Case MenuItem
    Case 1
       
    Case 2
    
    Case 3
    
    Case 4
    
    Case 5
        
    End Select

Case 4

Case 5

End Select

End Sub

Private Sub sbNovaOs()
Call sbLimpa_Campos(Me)
maskDataEnt = Date
SSTab2.Tab = 0
OptSolicitado.Value = True
OptSolicitado.Enabled = True

Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select  MAX(id) as nCountId from osdb", conn
ctNos = rsTabelas!ncountId + 1
rsTabelas.Close
Set rsTabelas = Nothing
ctrazao.SetFocus

End Sub

Private Sub sbsalvarDadosOS()
Dim nIdColaborador As Variant
  
If ctIdWeb.Text = "" Then
    ctIdWeb = 0
End If

nIdColaborador = fLocColaborador(cbcolaborador)

Sql = "INSERT INTO osdb (data, empresa, cnpj, insc, endereco, bairro, cidade, estado, cep, email, telefone, celular, " & _
    "roteiro, responsavel, servico, situacao, equipamento, nserie, marca, modelo, ref, tombo, defeito, diagnostico, news, idcolaborador, " & _
    "colaborador, op, idCliente ,status, Obs) VALUES ('" & Format(maskDataEnt, "yyyy-mm-dd") & "', '" & ctrazao & "', '" & ctcgc & "', '" & ctinsc & "', '" & ctend & "', " & _
    "'" & ctbairro & "', '" & ctcidade & "','" & ctuf & "', '" & MASKCEP & "','" & ctEmail & "', '" & MASKFONE & "', '" & MASKCEL & "', '" & ctroteiro & "'," & _
    "'" & ctcontato & "','" & cbServSolic & "', '" & Cbsituacao & "', '" & ctEquipamento & "', '" & ctNSerie & "', '" & ctmarca & "'," & _
    "'" & ctmodelo & "', '" & ctRef & "', '" & ctTombo & "', '" & ctDefeito & "', '" & ctDiagnostico & "','" & "Novo" & "'," & nIdColaborador & "," & _
    "'" & cbcolaborador & "','" & MOP & "','" & ctcodClient & "', '" & "1" & "','" & ctObs & "')"
    conn.Execute Sql
    

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

Private Sub sbAddEquipamento()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from equipamentos order by Descricao", conn
    ctEquipamento.Clear
    Do While Not rsTabelas.EOF
        ctEquipamento.AddItem rsTabelas!descricao
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
End Sub

Private Sub sbAddSituacao()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadeventos order by Descricao", conn
    Cbsituacao.Clear
    cbSituacaoHist.Clear
    Do While Not rsTabelas.EOF
        Cbsituacao.AddItem rsTabelas!descricao
        cbSituacaoHist.AddItem rsTabelas!descricao
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
End Sub

Private Sub sbAddColaborador()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from clientes where titulo='" & "PARCEIRO" & "' order by nome", conn
    cbcolaborador.Clear
    Do While Not rsTabelas.EOF
        cbcolaborador.AddItem rsTabelas!nome
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from clientes where titulo='" & "ADM" & "' order by nome", conn
    'cbcolaborador.Clear
    'Do While Not rsTabelas.EOF
        cbcolaborador.AddItem rsTabelas!nome
    '    rsTabelas.MoveNext
    'Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
    
    
End Sub

Private Sub sbAddPecas()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadpecas order by Descricao", conn
    cbPecas.Clear
    Do While Not rsTabelas.EOF
        cbPecas.AddItem rsTabelas!descricao
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
End Sub

Private Sub sbAddAcessorios()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadacessorios order by Descricao", conn
    cbAcessorio.Clear
    Do While Not rsTabelas.EOF
        cbAcessorio.AddItem rsTabelas!descricao
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
End Sub

Private Sub sbAddServExec()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from cadservexec order by Descricao", conn
    cbServExec.Clear
    Do While Not rsTabelas.EOF
        cbServExec.AddItem rsTabelas!descricao
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
    
End Sub

Private Function fLocColaborador(strNome As String) As String
    
    Set rsTabelas2 = New ADODB.Recordset
    rsTabelas2.Open "select * from clientes where nome='" & cbcolaborador & "'", conn
    
    fLocColaborador = rsTabelas2!ID
    rsTabelas2.Close
    Set rsTabelas2 = Nothing
    
End Function



Private Sub sbAtdadosOs()
Dim nIdColaborador As Variant
'On Error GoTo E:
If maskDataEnt.Text = "__/__/____" Then
    MsgBox "Informe a data de Entrada do equipamento para continuar", vbCritical
    Exit Sub
End If

If cbcolaborador.Text = "" Then
    MsgBox "Selecione um Parceiro para continuar,", vbCritical
    Exit Sub
End If

nIdColaborador = fLocColaborador(cbcolaborador)

'On Error Resume Next

'Sql = "UPDATE osdb SET data='" & maskDataEnt & "', empresa='" & ctrazao & "', cnpj='" & CTCGC & "', insc='" & CTINSC & "', endereco='" & CTEND & "'," & _
'"bairro='" & ctBairro & "', cidade='" & CTCIDADE & "', estado='" & CTUF & "', cep='" & MASKCEP & "', email='" & ctEmail & "', telefone='" & MASKFONE & "', " & _
'"celular='" & MASKCEL & "', roteiro='" & ctroteiro & "', responsavel='" & ctContato & "', servico='" & cbServSolic & "', situacao='" & Cbsituacao & "', " & _
'"equipamento='" & ctEquipamento & "', nserie='" & ctNSerie & "', marca='" & ctMarca & "', modelo='" & ctModelo & "', ref='" & ctRef & "', " & _
'"tombo='" & ctTombo & "', defeito='" & ctDefeito & "', diagnostico='" & ctDiagnostico & "', idcolaborador=" & nIdColaborador & ", " & _
'"colaborador='" & cbcolaborador & "', op='" & MOP & "', idCliente='" & ctcodClient & "', datasaida='" & maskDataPrev & "', total=" & converte(ctTotalOS) & " where id=" & ctNos & ""

Sql = "UPDATE osdb SET data='" & Format(maskDataEnt, "yyyy-mm-dd") & "', empresa='" & ctrazao & "', cnpj='" & ctcgc & "', insc='" & ctinsc & "', endereco='" & ctend & "'," & _
"bairro='" & ctbairro & "', cidade='" & ctcidade & "', estado='" & ctuf & "', cep='" & MASKCEP & "', email='" & ctEmail & "', telefone='" & MASKFONE & "', " & _
"celular='" & MASKCEL & "', roteiro='" & ctroteiro & "', responsavel='" & ctcontato & "', servico='" & cbServSolic & "', situacao='" & Cbsituacao & "', " & _
"equipamento='" & ctEquipamento & "', nserie='" & ctNSerie & "', marca='" & ctmarca & "', modelo='" & ctmodelo & "', ref='" & ctRef & "', " & _
"tombo='" & ctTombo & "', defeito='" & ctDefeito & "', diagnostico='" & ctDiagnostico & "', idcolaborador=" & nIdColaborador & ", " & _
"colaborador='" & cbcolaborador & "', op='" & MOP & "', idCliente='" & ctcodClient & "', datasaida='" & Format(maskDataPrev, "yyyy-mm-dd") & "', total=" & converte(ctTotalOS) & ", idWeb=" & ctIdWeb & ", Obs='" & ctObs & "' where id=" & ctNos & ""
conn.Execute Sql

'E:
'If Err.Number Then
'    MsgBox "Dados do cadastro incompletos.", vbCritical
'    Exit Sub
'End If
End Sub

Private Sub sbAtStatus(nNos As Variant, nstatua As Integer)
If nNos = "" Then
    MsgBox "Consulte uma Ordem de Serviço para continuar.", vbCritical
    Exit Sub
End If

Sql = "UPDATE osdb SET status='" & nstatua & "' where id=" & nNos & ""

conn.Execute Sql

Call fStatusOs(ctNos)

End Sub

Private Sub sbAtDataHist()
If ctNos.Text = "" Then
    Exit Sub
End If

With dataAdoHist
        .ConnectionString = strDns
        .RecordSource = "Select * From logeventos where n_os=" & ctNos & "  order by id"
   End With
    dataAdoHist.Refresh

End Sub

Private Sub sbAtDataAdoAcessorio()
If ctNos.Text = "" Then
    Exit Sub
End If

With DataAdoAcessorio
        .ConnectionString = strDns
        .RecordSource = "Select * From osdbacessorios where n_os=" & ctNos & "  order by id"
   End With
   DataAdoAcessorio.Refresh

End Sub

Private Sub sbAtDataServExec()
If ctNos.Text = "" Then
    Exit Sub
End If

With dataAdoServExec
        .ConnectionString = strDns
        .RecordSource = "Select * From osdbservexec where n_os=" & ctNos & "  order by id"
   End With
   dataAdoServExec.Refresh

End Sub

Private Sub sbAtDataPecas()
If ctNos.Text = "" Then
    Exit Sub
End If

With DataAdoPecas
        .ConnectionString = strDns
        .RecordSource = "Select * From osdbpecas where n_os=" & ctNos & "  order by id"
   End With
   DataAdoPecas.Refresh

End Sub

Private Sub sbExcOs()
If ctNos.Text = "" Then
    MsgBox "Selecione uma Ordem de Serviço para continuar.", vbCritical
    Exit Sub
Else
    Ret = MsgBox("Tem certeza que deseja excluir esta Ordem de serviço?", vbYesNo)
    If Ret = 6 Then
        Sql = "DELETE from osdb where id=" & ctNos & ""
        conn.Execute Sql
        MsgBox "Ordem de Serviço excluida com sucesso."
    End If
End If
End Sub

Private Sub sbSumServicos(nIDOS As Variant)
Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select SUM(valtotal) as nValTotal from osdbservexec where n_os=" & nIDOS & "", conn
ctTotalServ = Format(rsTabelas!nValTotal, "##,##0.00")
rsTabelas.Close
Set rsTabelas = Nothing

End Sub

Private Sub sbSumPecas(nIDOS As Variant)
Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select SUM(valtotal) as nValTotal from osdbpecas where n_os=" & nIDOS & "", conn
ctpecasexec = Format(rsTabelas!nValTotal, "##,##0.00")
rsTabelas.Close
Set rsTabelas = Nothing

End Sub

Private Sub sbTotalOS(nIDOS As Variant)
Dim nValServ As Double
Dim nValPecas As Double

Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select SUM(valtotal) as nValTotal from osdbservexec where n_os=" & nIDOS & "", conn
If IsNull(rsTabelas!nValTotal) Then
    nValServ = 0
Else
    nValServ = rsTabelas!nValTotal
End If
rsTabelas.Close
Set rsTabelas = Nothing

rsTabelas.Open "select SUM(valtotal) as nValTotal from osdbpecas where n_os=" & nIDOS & "", conn
If IsNull(rsTabelas!nValTotal) Then
    nValPecas = 0
Else
    nValPecas = rsTabelas!nValTotal
End If
rsTabelas.Close
Set rsTabelas = Nothing

ctTotalServ = Format(nValServ, "##,##0.00")
ctpecasexec = Format(nValPecas, "##,##0.00")

ctTotalOS = Format(nValServ + nValPecas, "##,##0.00")


End Sub

Private Sub ctIdWeb_GotFocus()
   On Error Resume Next
   ctIdWeb.SelStart = 0
   ctIdWeb.SelLength = Len(ctIdWeb.Text)
End Sub

Private Sub ctIdWeb_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub sbAtdadosOsWeb()
Dim nIdColaborador As Variant
'On Error GoTo E:


Call Conecta_BDRemoto

nIdColaborador = fLocColaborador(cbcolaborador)

'On Error Resume Next

'Sql = "UPDATE osdb SET data='" & maskDataEnt & "', empresa='" & ctrazao & "', cnpj='" & CTCGC & "', insc='" & CTINSC & "', endereco='" & CTEND & "'," & _
'"bairro='" & ctBairro & "', cidade='" & CTCIDADE & "', estado='" & CTUF & "', cep='" & MASKCEP & "', email='" & ctEmail & "', telefone='" & MASKFONE & "', " & _
'"celular='" & MASKCEL & "', roteiro='" & ctroteiro & "', responsavel='" & ctContato & "', servico='" & cbServSolic & "', situacao='" & Cbsituacao & "', " & _
'"equipamento='" & ctEquipamento & "', nserie='" & ctNSerie & "', marca='" & ctMarca & "', modelo='" & ctModelo & "', ref='" & ctRef & "', " & _
'"tombo='" & ctTombo & "', defeito='" & ctDefeito & "', diagnostico='" & ctDiagnostico & "', idcolaborador=" & nIdColaborador & ", " & _
'"colaborador='" & cbcolaborador & "', op='" & MOP & "', idCliente='" & ctcodClient & "', datasaida='" & maskDataPrev & "', total=" & converte(ctTotalOS) & " where id=" & ctNos & ""

Sql = "UPDATE osdb SET data='" & maskDataEnt & "', empresa='" & ctrazao & "', cnpj='" & ctcgc & "', insc='" & ctinsc & "', endereco='" & ctend & "'," & _
"bairro='" & ctbairro & "', cidade='" & ctcidade & "', estado='" & ctuf & "', cep='" & MASKCEP & "', email='" & ctEmail & "', telefone='" & MASKFONE & "', " & _
"celular='" & MASKCEL & "', roteiro='" & ctroteiro & "', responsavel='" & ctcontato & "', servico='" & cbServSolic & "', situacao='" & Cbsituacao & "', " & _
"equipamento='" & ctEquipamento & "', nserie='" & ctNSerie & "', marca='" & ctmarca & "', modelo='" & ctmodelo & "', ref='" & ctRef & "', " & _
"tombo='" & ctTombo & "', defeito='" & ctDefeito & "', diagnostico='" & ctDiagnostico & "', idcolaborador=" & nIdColaborador & ", " & _
"colaborador='" & cbcolaborador & "', op='" & MOP & "', idCliente='" & ctcodClient & "', datasaida='" & maskDataPrev & "', total=" & converte(ctTotalOS) & ", idWeb=" & ctIdWeb & " where id=" & ctNos & ""
connRemoto.Execute Sql

'E:
'If Err.Number Then
'    MsgBox "Dados do cadastro incompletos.", vbCritical
'End If
End Sub

Private Sub locCep()
Dim Retorno As Variant, x_End() As String
      
   If Len(MASKCEP) > 0 Then
      If Len(MASKCEP) < 7 Then
         MsgBox "Necessário corretamente o CEP ou deixa-lo em branco!"
         MASKCEP.SetFocus
         Exit Sub
      End If
      
      Retorno = Peg_CEP(MASKCEP, "suportek", "suporte")
      
      If Mid(Retorno, 1, 4) <> "ERRO" Then
         x_End = Split(Retorno, ", ")
         
         ctend.Text = x_End(0)
         ctbairro.Text = x_End(1)
         ctcidade.Text = x_End(2)
         ctuf.Text = x_End(3)
      End If
      
   Else
   
      If ctend.Text = "" Then
         MsgBox "Necessário informar Logradouro!"
         ctend.SetFocus
         Exit Sub
      End If
   
      If ctcidade.Text = "" Then
         MsgBox "Necessário informar Localidade!"
         ctcidade.SetFocus
         Exit Sub
      End If
      
      If ctuf.Text = "" Then
         MsgBox "Necessário informar UF!"
         ctuf.SetFocus
         Exit Sub
      End If
      
      Peg_Lg ctend.Text, ctcidade.Text, ctuf.Text, "suportek", "suporte"
      
   End If


End Sub
Private Function Peg_CEP(x_CEP, CepUser, CepPWD) As Variant

  Dim webService As New SoapClient, vaRet As String
      
   If x_CEP = "" Or Len(x_CEP) < 7 Then
      MsgBox "CEP invádo!"
      Peg_CEP = "ERRO"
      Exit Function
   End If

   '
   'Seta o site que provem o serviçde webservice
   '
   webService.mssoapinit ("http://www.byjg.com.br/xmlnuke-php/webservice.php/ws/cep?WSDL")
   '
   'string[] obterCEPAuth(string Logradouro, string Localidade, string UF,Usuario,Senha)
   'Devolve uma lista de CEP àartir de um nome ou parte do nome de um Logradouro. Apenas os 20 primeiros endereç sãretornados. Com o objetivo de se obter uma lista mais precisa possíl, foram acrescentados os parâtros Localidade e UF. Agradecimentos a Paulo Santana pela sugestã  '
   'string obterLogradouro(string CEP)
   'Devolve o nome do Logradouro àartir do CEP fornecido no formato 00000-000 ou 00000000.
   '
   'string obterVersao()
   'Devolve informaçs sobre a versãdo serviçde CEP.
   '

   vaRet = webService.obterLogradouroAuth(x_CEP, CepUser, CepPWD)
    
   If (Mid(UCase(vaRet), 1, 3) = "CEP") Or (Mid(UCase(vaRet), 1, 7) = "USUÁRIO") Then
      MsgBox UCase(vaRet), vbCritical + vbOKOnly
      Peg_CEP = "ERRO"
      Exit Function
   End If
 
   ' Retorna o Array com os dados do endereço
   Peg_CEP = vaRet
  
End Function

Private Sub Peg_Lg(xLogr, xLocal, xUf, CepUser, CepPWD)
   Dim webService As New SoapClient, Retorno As Variant
 
   '
   'Seta o site que provem o serviçde webservice
   '
   webService.mssoapinit ("http://www.byjg.com.br/site/webservice.php/ws/cep?WSDL")

   '
   'string[] obterCEPAuth(string Logradouro, string Localidade, string UF,Usuario,Senha)
   'Devolve uma lista de CEP àartir de um nome ou parte do nome de um Logradouro. Apenas os 20 primeiros endereç sãretornados. Com o objetivo de se obter uma lista mais precisa possíl, foram acrescentados os parâtros Localidade e UF. Agradecimentos a Paulo Santana pela sugestã  '
   'string obterLogradouro(string CEP)
   'Devolve o nome do Logradouro àartir do CEP fornecido no formato 00000-000 ou 00000000.
   '
   'string obterVersao()
   'Devolve informaçs sobre a versãdo serviçde CEP.
   '
   'Retorno = webService.obterLogradouroAuth(txtCEP.Text, Usuario, Senha)
   
   Retorno = webService.obterCEPAuth(xLogr, xLocal, xUf, CepUser, CepPWD)

   '
   'Imprime o retorno da web
   '
   Dim xx As Integer
   
   Lt_cep.Clear
   'txtEscolha.Text = ""
   
   For xx = 0 To UBound(Retorno)
      lstResultado.AddItem Retorno(xx)
   Next

End Sub
Private Sub ctObs_GotFocus()
   On Error Resume Next
   ctObs.SelStart = 0
   ctObs.SelLength = Len(ctObs.Text)
End Sub

Private Sub ctObs_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CbChamado_GotFocus()
   On Error Resume Next
   CbChamado.SelStart = 0
   CbChamado.SelLength = Len(CbChamado.Text)
End Sub

Private Sub sbAtDataUvt()
If ctNos.Text = "" Then
    Exit Sub
End If

With AdoDataUvt
        .ConnectionString = strDns
        .RecordSource = "Select * From tbUvt where IdOs=" & ctNos & "  order by id"
   End With
   AdoDataUvt.Refresh

End Sub

Private Sub sbAtChamadoUvt(nIDOS As Variant)
Dim nQtdDias As Variant

Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select * from tbUvt where idOs=" & ctNos & "", conn

Do While Not rsTabelas.EOF
    nQtdDias = DateDiff("d", Date, rsTabelas!dataFin)
    Sql = "UPDATE tbUvt SET QtdDias=" & nQtdDias & " where id=" & rsTabelas!ID & " and IdOs=" & rsTabelas!IdOs & ""
    
    conn.Execute Sql
    rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing
    
    AdoDataUvt.Refresh
End Sub

Private Sub sbAtUvt()
Dim nQtdDias As Variant

Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select * from tbUvt Order By IdOs", conn

Do While Not rsTabelas.EOF
    nQtdDias = DateDiff("d", Date, rsTabelas!dataFin)
    Sql = "UPDATE tbUvt SET QtdDias=" & nQtdDias & " where id=" & rsTabelas!ID & " and IdOs=" & rsTabelas!IdOs & ""
    
    conn.Execute Sql
    rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing
    
    'AdoDataUvt.Refresh
End Sub

Private Sub sbAtrasoUVT()
Set rsTabelas = New ADODB.Recordset
Set rsTabelas2 = New ADODB.Recordset

Sql = "DELETE from tbatrasouvt"

conn.Execute Sql

rsTabelas.Open "select * from Osdb where ref='" & "ECF" & "' Order By id", conn
Do While Not rsTabelas.EOF
    rsTabelas2.Open "select * from tbUvt where IdOs=" & rsTabelas!ID & " order by id DESC", conn
    If Not rsTabelas2.EOF Then
        If rsTabelas2!qtdDias < 4 Then

            Sql = "INSERT INTO tbatrasouvt (id, IdOs, IE, Empresa, Status, equipamento, dataIni, DataFin, QtdDias) VALUES (" & "1" & ", " & rsTabelas!ID & "," & _
            "'" & rsTabelas!insc & "', '" & rsTabelas!empresa & "', '" & rsTabelas2!Status & "' ,'" & rsTabelas!modelo & "'," & _
            "'" & Format(rsTabelas2!dataIni, "yyyy-mm-dd") & "', '" & Format(rsTabelas2!dataFin, "yyyy-mm-dd") & "'," & rsTabelas2!qtdDias & ")"
            conn.Execute Sql
        
        End If
    End If
    rsTabelas2.Close
    Set rsTabelas2 = Nothing
    rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing
End Sub

Private Function FsVerifDupNS(strNs As String) As Integer
rsTabelas.Open "select * from osdb where nserie='" & strNs & "'", conn
If rsTabelas.EOF Then
    FsVerifDupNS = 0
Else
    FsVerifDupNS = 1
End If
rsTabelas.Close
Set rsTabelas = Nothing
End Function
