VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaReparto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reparto"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   5745
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   330
      Left            =   4440
      TabIndex        =   3
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.TextBox TxtReparto 
         Height          =   315
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   0
         Text            =   "0000000000000"
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   960
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         Format          =   707985409
         CurrentDate     =   37156
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   1305
         TabIndex        =   1
         Top             =   960
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         Format          =   707985409
         CurrentDate     =   37156
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   225
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label9 
         Caption         =   "Reparto :"
         Height          =   225
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "VentaReparto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  nImpr = 58
  FrmImpresor.Show
End Sub

Private Sub Form_Load()
   Limpiar
End Sub

Private Sub Limpiar()
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  TxtReparto.text = "0000000000000"
End Sub

Private Sub TxtReparto_GotFocus()
  TxtReparto.SelStart = 0
  TxtReparto.SelLength = Len(TxtReparto.text)
End Sub

Private Sub TxtReparto_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtReparto_LostFocus()
  If TxtReparto.text = "" Then TxtReparto.text = 0
  TxtReparto.text = Format(TxtReparto.text, "0000000000000")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub
