VERSION 5.00
Begin VB.Form AplicarFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura Aplicar"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   465
      Left            =   3120
      TabIndex        =   1
      Top             =   2355
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Factura a Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   4395
      Begin VB.TextBox TxtFactura 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2100
         TabIndex        =   0
         Text            =   "TxtFactura"
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Nº de Fact a Aplicar :"
         Height          =   180
         Left            =   300
         TabIndex        =   3
         Top             =   855
         Width           =   1635
      End
   End
   Begin VB.Label Label2 
      Caption         =   "ESC = Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   150
      TabIndex        =   4
      Top             =   2370
      Width           =   1335
   End
End
Attribute VB_Name = "AplicarFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()
    nFac = TxtFactura.text
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Unload Me
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
  TxtFactura.text = ""
  nFac = 0
End Sub

Private Sub TxtFactura_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtFactura_LostFocus()
  If TxtFactura.text = "" Then TxtFactura.text = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub DtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
