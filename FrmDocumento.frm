VERSION 5.00
Begin VB.Form FrmDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nº de Documento"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "Aplicar"
      Height          =   360
      Left            =   825
      TabIndex        =   1
      Top             =   840
      Width           =   1485
   End
   Begin VB.TextBox TxtDoc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   525
      MaxLength       =   10
      TabIndex        =   0
      Top             =   195
      Width           =   2115
   End
End
Attribute VB_Name = "FrmDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOk_Click()
   If TxtDoc.text <> "" Then
      nDocAfip = TxtDoc.text
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   TxtDoc.text = ""
   nDocAfip = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
'     KeyAscii = Upper(KeyAscii)
End Sub

Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

