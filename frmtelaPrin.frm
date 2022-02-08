VERSION 5.00
Object = "{CC01E02E-242E-11D8-A6B9-000B231D9747}#1.0#0"; "vertmenu.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmtelaPrin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suportek Consultoria e serviços"
   ClientHeight    =   10455
   ClientLeft      =   2220
   ClientTop       =   480
   ClientWidth     =   12465
   Icon            =   "frmtelaPrin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmtelaPrin.frx":27A2
   ScaleHeight     =   0
   ScaleWidth      =   0
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   10020
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "Sistema de gerenciamento"
            TextSave        =   "Sistema de gerenciamento"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8996
            MinWidth        =   8996
         EndProperty
      EndProperty
   End
   Begin VertMenu.VerticalMenu Vtmnu 
      Height          =   8805
      Left            =   30
      TabIndex        =   0
      Top             =   2010
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   15531
      Enabled         =   -1  'True
      MenusMax        =   6
      MenuCaption1    =   "Arquivo"
      MenuItemsMax1   =   4
      MenuItemIcon11  =   "frmtelaPrin.frx":FE26C
      MenuItemCaption11=   "Ordem de Serviço"
      MenuItemKey11   =   "mnuOs"
      MenuItemIcon12  =   "frmtelaPrin.frx":FEFBE
      MenuItemCaption12=   "Enviar Sms"
      MenuItemKey12   =   "mnuSms"
      MenuItemIcon13  =   "frmtelaPrin.frx":FFD10
      MenuItemCaption13=   "Agenda Fone"
      MenuItemKey13   =   "mnuAgenda"
      MenuItemIcon14  =   "frmtelaPrin.frx":100A62
      MenuItemCaption14=   "Sair"
      MenuItemKey14   =   "mnuSair"
      MenuCaption2    =   "Cadastros"
      MenuItemsMax2   =   6
      MenuItemIcon21  =   "frmtelaPrin.frx":1017B4
      MenuItemCaption21=   "Clientes"
      MenuItemKey21   =   "mnuCadClie"
      MenuItemIcon22  =   "frmtelaPrin.frx":102506
      MenuItemCaption22=   "Parceiros"
      MenuItemKey22   =   "mnuCadParc"
      MenuItemIcon23  =   "frmtelaPrin.frx":103258
      MenuItemCaption23=   "Fornecedores"
      MenuItemKey23   =   "mnuCadForn"
      MenuItemIcon24  =   "frmtelaPrin.frx":103FAA
      MenuItemCaption24=   "Funcionários"
      MenuItemKey24   =   "mnuCadFun"
      MenuItemIcon25  =   "frmtelaPrin.frx":104CFC
      MenuItemCaption25=   "Usuários"
      MenuItemKey25   =   "mnuCadUso"
      MenuItemIcon26  =   "frmtelaPrin.frx":105A4E
      MenuItemCaption26=   "Contato Sms"
      MenuItemKey26   =   "mnuContatoSms"
      MenuCaption3    =   "Tabelas"
      MenuItemsMax3   =   8
      MenuItemIcon31  =   "frmtelaPrin.frx":1067A0
      MenuItemCaption31=   "Equipamentos"
      MenuItemKey31   =   "mnuEquipamentos"
      MenuItemIcon32  =   "frmtelaPrin.frx":1074F2
      MenuItemCaption32=   "Acessórios"
      MenuItemKey32   =   "mnuAcessorios"
      MenuItemIcon33  =   "frmtelaPrin.frx":108244
      MenuItemCaption33=   "Serviços Solicitados"
      MenuItemKey33   =   "mnuServSolicitados"
      MenuItemIcon34  =   "frmtelaPrin.frx":108F96
      MenuItemCaption34=   "Serviços Executados"
      MenuItemKey34   =   "mnuServexec"
      MenuItemIcon35  =   "frmtelaPrin.frx":109CE8
      MenuItemCaption35=   "Cadastro de Peças"
      MenuItemKey35   =   "mnuServReal"
      MenuItemIcon36  =   "frmtelaPrin.frx":10AA3A
      MenuItemCaption36=   "Cep Logradouros"
      MenuItemKey36   =   "mnuCep"
      MenuItemIcon37  =   "frmtelaPrin.frx":10B78C
      MenuItemCaption37=   "Eventos"
      MenuItemKey37   =   "mnuEventos"
      MenuItemIcon38  =   "frmtelaPrin.frx":10C4DE
      MenuItemCaption38=   "Entrada Estoq."
      MenuItemKey38   =   "mnuEntradaEstoq"
      MenuCaption4    =   "Relatórios"
      MenuItemsMax4   =   3
      MenuItemIcon41  =   "frmtelaPrin.frx":10D230
      MenuItemCaption41=   "Assistente de Relatórios"
      MenuItemKey41   =   "mnuRelAss"
      MenuItemIcon42  =   "frmtelaPrin.frx":10DF82
      MenuItemCaption42=   "Sol. Lacração"
      MenuItemKey42   =   "mnuSolLacra"
      MenuItemIcon43  =   "frmtelaPrin.frx":10ECD4
      MenuItemCaption43=   "Item3"
      MenuCaption5    =   "Utilitários"
      MenuItemsMax5   =   3
      MenuItemIcon51  =   "frmtelaPrin.frx":10FA26
      MenuItemCaption51=   "Backup"
      MenuItemKey51   =   "mnuBackup"
      MenuItemIcon52  =   "frmtelaPrin.frx":110778
      MenuItemCaption52=   "Restaurar Backup"
      MenuItemKey52   =   "mnuResBackup"
      MenuItemIcon53  =   "frmtelaPrin.frx":1114CA
      MenuItemCaption53=   "Configurações"
      MenuItemKey53   =   "mnuConf"
      MenuCaption6    =   "Financeiro"
      MenuItemsMax6   =   5
      MenuItemIcon61  =   "frmtelaPrin.frx":11221C
      MenuItemCaption61=   "Caixa diário"
      MenuItemKey61   =   "mnuCaixaDiario"
      MenuItemIcon62  =   "frmtelaPrin.frx":112F6E
      MenuItemCaption62=   "Contas a Receber"
      MenuItemKey62   =   "mnuContReceber"
      MenuItemIcon63  =   "frmtelaPrin.frx":113CC0
      MenuItemCaption63=   "Contas a Pagar"
      MenuItemKey63   =   "mnuContApagar"
      MenuItemIcon64  =   "frmtelaPrin.frx":114A12
      MenuItemCaption64=   "Fluxo de Caixa"
      MenuItemKey64   =   "mnuFluxoCaixa"
      MenuItemIcon65  =   "frmtelaPrin.frx":115764
      MenuItemCaption65=   "Plano de Contas"
      MenuItemKey65   =   "mnuPlanoContas"
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "frmtelaPrin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory _
As String, ByVal nShowCmd As Long) As Long


Private Const SW_SHOWNORMAL = 1

Private Sub Vtmnu_MenuItemClick(MenuNumber As Long, MenuItem As Long)
Select Case MenuNumber
Case 1
    Select Case MenuItem
    Case 1
      Call AtivaForm(frmOs)
    Case 2
        Call AtivaForm(frmEnviaSms)
    Case 3
        'Call fcOutlook("Solicitação de retirada de Equipamento.", "DM Distribuidora de peçcas", "contato@suportek.net", "Informa-mos que o equipamentos Imp já se encontra a sua disposição para retirada.")
        Call AtivaForm(frmAgenda)
    Case 4
        Unload Me
    End Select
Case 2
    Select Case MenuItem
    Case 1
       Call AtivaForm(frmCad_CLIENTE)
    Case 2
        Call AtivaForm(frmCad_parceiros)
    Case 3
        Call AtivaForm(frmCad_fornecedores)
    Case 4
        Call AtivaForm(frmCad_FUNCIONARIOS)
    Case 5
        Call AtivaForm(frmCad_Usuario)
    Case 6
        Call AtivaForm(frmCadsms)
    
    End Select

Case 3
    Select Case MenuItem
    Case 1
       Call AtivaForm(frmcadEquipamentos)
    Case 2
        Call AtivaForm(frmcadacessorios)
    Case 3
        Call AtivaForm(frmcadServicos)
    Case 4
        Call AtivaForm(frmcadServExec)
    Case 5
        Call AtivaForm(frmcadPecas)
    Case 6
        ' cep
    Case 7
        Call AtivaForm(frmcadEventos)
    Case 8
        Call AtivaForm(frmEntPeças)
    
    End Select

Case 4
    Select Case MenuItem
    Case 1
        Call AtivaForm(frmRel)
    Case 2
        strArqPdf = App.Path & "\..\pdf\Solicitação de lacracao.pdf"
        Call AtivaForm(frmPdf)
    Case 3
    
    End Select
    
    
Case 5

Case 6
    Select Case MenuItem
    Case 1
       Call AtivaForm(FRM_CAIXA)
    Case 2
        
    Case 3
        Call AtivaForm(frmContasApagar)
    Case 4
        
    Case 5
        Call AtivaForm(frmPlanocontas)
    End Select

End Select
End Sub

Public Sub SendMail(Optional Address As String, _
Optional Subject As String, Optional Body As String, _
Optional CC As String, Optional BCC As String)

Dim strCommand As String

'constroi a string do email
If Len(Subject) Then strCommand = "&Subject=" & Subject
If Len(Body) Then strCommand = strCommand & "&Body=" & Body
If Len(CC) Then strCommand = strCommand & "&CC=" & CC
If Len(BCC) Then strCommand = strCommand & "&BCC=" & BCC

'substitui o primeiro &
'com interrogacao
If Len(strCommand) Then
   Mid(strCommand, 1, 1) = "?"
End If

'Inclui o comando mailto: e o endereço de e-mail
strCommand = "mailto:" & Address & strCommand

'executa o comando via API
Call ShellExecute(Me.hwnd, "open", strCommand, vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub sbEnviarEmail()
Call SendMail("macoratti@riopreto.com.br", "Testando o preenchimento do Outlook", _
  "Esta é a mensagem do seu e-mail...", _
  "copia@carbono.com.br", "copiapara@carbono.com.br")
End Sub
