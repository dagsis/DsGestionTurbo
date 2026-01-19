VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CajaChequesTer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques de Terceros"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   3735
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   3510
      Begin VB.CheckBox ChkCartera 
         Caption         =   "Solo en Cartera"
         Height          =   210
         Left            =   165
         TabIndex        =   7
         Top             =   1920
         Width           =   1875
      End
      Begin VB.CheckBox ChkVenc 
         Caption         =   "Por Fecha de Vencimiento"
         Height          =   240
         Left            =   165
         TabIndex        =   6
         Top             =   1605
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   450
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65339393
         CurrentDate     =   36983
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   345
         Left            =   1125
         TabIndex        =   3
         Top             =   945
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65339393
         CurrentDate     =   36983
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   435
         TabIndex        =   5
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   405
         TabIndex        =   4
         Top             =   945
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Aplicar"
      Height          =   420
      Left            =   2625
      TabIndex        =   0
      Top             =   2370
      Width           =   945
   End
End
Attribute VB_Name = "CajaChequesTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdImprimir_Click()
  nImpr = 41
  FrmImpresor.Show
End Sub

Private Sub Form_Load()
  DTPDesde.Value = Date
  DTPHasta.Value = Date
End Sub

