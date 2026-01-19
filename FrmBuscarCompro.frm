VERSION 5.00
Begin VB.Form FrmBuscarCompro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Comprobantes"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   345
      Left            =   2940
      TabIndex        =   7
      Top             =   2250
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   3930
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   315
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1185
         Width           =   1110
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2610
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   525
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   315
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   525
         Width           =   2220
      End
      Begin VB.Label Label3 
         Caption         =   "Número :"
         Height          =   195
         Left            =   300
         TabIndex        =   5
         Top             =   945
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Suc :"
         Height          =   195
         Left            =   2595
         TabIndex        =   3
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobante :"
         Height          =   285
         Left            =   285
         TabIndex        =   1
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmBuscarCompro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

