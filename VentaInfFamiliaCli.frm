VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentInfFamiliaCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Ventas por Familia/Clientes"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   6165
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   330
      Left            =   4965
      TabIndex        =   15
      Top             =   3015
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5970
      Begin VB.CheckBox ChkFecha 
         Caption         =   "Limitar a Fechas"
         Height          =   195
         Left            =   1455
         TabIndex        =   6
         Top             =   2370
         Width           =   1980
      End
      Begin VB.CheckBox ChkGeneral 
         Caption         =   "General"
         Height          =   210
         Left            =   270
         TabIndex        =   5
         Top             =   2370
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   255
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "TxtCliente"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4980
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   255
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "TxtCliente"
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4980
         TabIndex        =   1
         Top             =   1260
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   360
         Left            =   900
         TabIndex        =   7
         Top             =   1800
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   54263809
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   360
         Left            =   3135
         TabIndex        =   8
         Top             =   1800
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   54263809
         CurrentDate     =   37021
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1155
         TabIndex        =   14
         Top             =   600
         Width           =   3705
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1155
         TabIndex        =   13
         Top             =   1260
         Width           =   3705
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   285
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   165
         Left            =   2580
         TabIndex        =   11
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   255
         TabIndex        =   10
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1005
         Width           =   1260
      End
   End
End
Attribute VB_Name = "VentInfFamiliaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

