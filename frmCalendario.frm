VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frmCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendário"
   ClientHeight    =   4035
   ClientLeft      =   6750
   ClientTop       =   3555
   ClientWidth     =   4290
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView montView 
      Height          =   3960
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6985
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   66519041
      CurrentDate     =   41589
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
montView.Value = Date
End Sub

Private Sub montView_DateClick(ByVal DateClicked As Date)
strData = montView
Unload Me
End Sub
