VERSION 5.00
Begin VB.Form FrmClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   90
      TabIndex        =   2
      Top             =   15
      Width           =   4185
      Begin VB.CommandButton CmdOk 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   1410
         TabIndex        =   1
         Top             =   975
         Width           =   1380
      End
      Begin VB.TextBox TxtClave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   855
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   450
         Width           =   2385
      End
   End
End
Attribute VB_Name = "FrmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
   bClave = False
   If TxtClave.text = sClave Or TxtClave.text = sClave1 Then
      If TxtClave.text <> "0" Then
         bClave = True
         bClave1 = True
      Else
         bClave = False
         bClave1 = False
      End If
      Unload Me
   Else
      TxtClave.text = ""
      TxtClave.SetFocus
   End If
End Sub

Private Sub Form_Load()
   TxtClave.text = ""
   bClave = False
   bClave1 = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
'     KeyAscii = Upper(KeyAscii)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      bClave = False
      bClave1 = False
      Unload Me
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If bClave = False Then
     Cancel = True
  End If
End Sub
