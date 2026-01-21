VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   8010
      Begin VB.CommandButton CmdValores 
         Caption         =   "Cuotas"
         Height          =   720
         Index           =   2
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdValores 
         Caption         =   "Salir"
         Height          =   720
         Index           =   3
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdValores 
         Caption         =   "Borrar"
         Height          =   720
         Index           =   1
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdValores 
         Caption         =   "Nuevo"
         Height          =   720
         Index           =   0
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3315
         Width           =   945
      End
      Begin VB.TextBox TxtAutorizacion 
         Height          =   315
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "TxtAutorizacion"
         Top             =   2745
         Width           =   1905
      End
      Begin VB.TextBox TxtCupon 
         Height          =   315
         Left            =   4125
         MaxLength       =   30
         TabIndex        =   9
         Text            =   "TxtCupon"
         Top             =   2745
         Width           =   1860
      End
      Begin VB.TextBox TxtTarjeta 
         Height          =   315
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "TxtTarjeta"
         Top             =   2745
         Width           =   2550
      End
      Begin VB.TextBox TxtCheque 
         Height          =   315
         Left            =   4935
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "TxtCheque"
         Top             =   2145
         Width           =   1665
      End
      Begin VB.ComboBox CmbBanco 
         Height          =   315
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2145
         Width           =   2175
      End
      Begin VB.TextBox TxtNombre 
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "TxtNombre"
         Top             =   2145
         Width           =   2550
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "TxtImporte"
         Top             =   1200
         Width           =   1065
      End
      Begin VB.ComboBox CmbPago 
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Width           =   2730
      End
      Begin MSComCtl2.DTPicker DTPVencimiento 
         Height          =   315
         Left            =   2685
         TabIndex        =   8
         Top             =   2745
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   158400513
         CurrentDate     =   36795
      End
      Begin MSComCtl2.DTPicker DTPAcreditacion 
         Height          =   315
         Left            =   6615
         TabIndex        =   6
         Top             =   2145
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   158400513
         CurrentDate     =   36795
      End
      Begin MSDataGridLib.DataGrid Grid2 
         Height          =   1635
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   2884
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7905
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Label LblTPago 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTPago"
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
         Left            =   150
         TabIndex        =   25
         Top             =   3750
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label29 
         Caption         =   "Nº Cupón :"
         Height          =   195
         Left            =   4140
         TabIndex        =   24
         Top             =   2535
         Width           =   870
      End
      Begin VB.Label Label28 
         Caption         =   "Nº Autorización :"
         Height          =   180
         Left            =   6015
         TabIndex        =   23
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label Label27 
         Caption         =   "F.Vencimiento :"
         Height          =   180
         Left            =   2685
         TabIndex        =   22
         Top             =   2535
         Width           =   1110
      End
      Begin VB.Label Label26 
         Caption         =   "Nº Tarjeta :"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "F.Acreditación :"
         Height          =   240
         Left            =   6630
         TabIndex        =   20
         Top             =   1905
         Width           =   1140
      End
      Begin VB.Label Label23 
         Caption         =   "Nº de Cheque :"
         Height          =   195
         Left            =   4935
         TabIndex        =   19
         Top             =   1905
         Width           =   1200
      End
      Begin VB.Label Label22 
         Caption         =   "Banco :"
         Height          =   210
         Left            =   2685
         TabIndex        =   18
         Top             =   1905
         Width           =   750
      End
      Begin VB.Label LblNombre 
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   1905
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Importe :"
         Height          =   210
         Left            =   150
         TabIndex        =   16
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label18 
         Caption         =   "Forma de Pago :"
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   240
         Width           =   1320
      End
   End
End
Attribute VB_Name = "VentaValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Inicializar(ByRef rs As ADODB.Recordset, ByVal totalPagar As Double)
    Set mRs = rs
    mTotal = totalPagar
    Set GridValores.DataSource = mRs
    LblTotalPagar.Caption = Format(mTotal, "#0.00")
    RefrescarSumatoria
End Sub

