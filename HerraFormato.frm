VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form HerraFormato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formatos de Comprobantes"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4230
      Left            =   345
      TabIndex        =   4
      Top             =   1275
      Width           =   8385
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   345
      Left            =   7890
      TabIndex        =   3
      Top             =   5835
      Width           =   1110
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4950
      Left            =   165
      TabIndex        =   2
      Top             =   795
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   8731
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PreImpreso"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1425
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   210
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   270
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   1140
   End
End
Attribute VB_Name = "HerraFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

