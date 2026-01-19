VERSION 5.00
Begin VB.Form HerraDatosCorreo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros del Correo Electronico"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   15870
   Begin VB.CommandButton CmbOk 
      Caption         =   "Aplicar"
      Height          =   420
      Left            =   14445
      TabIndex        =   0
      Top             =   3240
      Width           =   1290
   End
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
      Height          =   3105
      Left            =   270
      TabIndex        =   6
      Top             =   90
      Width           =   15480
      Begin VB.TextBox TxtAsunto 
         Height          =   2610
         Left            =   5340
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "HerraDatosCorreo.frx":0000
         Top             =   330
         Width           =   9870
      End
      Begin VB.TextBox TxtServer 
         Height          =   300
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "TxtServer"
         Top             =   810
         Width           =   3360
      End
      Begin VB.TextBox TxtRemitente 
         Height          =   300
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "TxtRemitente"
         Top             =   2445
         Width           =   3360
      End
      Begin VB.TextBox TxtPass 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1605
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "TxtPass"
         Top             =   1755
         Width           =   2070
      End
      Begin VB.TextBox TxtUser 
         Height          =   300
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "TxtUser"
         Top             =   1245
         Width           =   3360
      End
      Begin VB.ComboBox CmbServidor 
         Height          =   315
         ItemData        =   "HerraDatosCorreo.frx":000F
         Left            =   1605
         List            =   "HerraDatosCorreo.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   375
         Width           =   2505
      End
      Begin VB.Label Label5 
         Caption         =   "SMTP :"
         Height          =   240
         Left            =   270
         TabIndex        =   11
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Remitente :"
         Height          =   240
         Left            =   315
         TabIndex        =   10
         Top             =   2445
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Contraseña :"
         Height          =   240
         Left            =   285
         TabIndex        =   9
         Top             =   1755
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario :"
         Height          =   240
         Left            =   285
         TabIndex        =   8
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Servidor SMTP :"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   375
         Width           =   1320
      End
   End
End
Attribute VB_Name = "HerraDatosCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbOk_Click()
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
              
  sSql = "UPDATE Sistema SET Email=" & CmbServidor.ListIndex
  db.Execute sSql
            
  sSql = "UPDATE EmailDatos SET Servidor='" & TxtServer.text & "',Envia='" & TxtRemitente.text & "',Usuario='" & TxtUser.text & "',Contraseña='" & TxtPass.text & "',Asunto=' " & TxtAsunto.text & "' "
  sSql = sSql & "WHERE Id=" & CmbServidor.ListIndex
  db.Execute sSql
    
  db.CommitTrans
  db.Close
  
  MsgBox "Datos Actualizados Con Exito", vbInformation, "Atención"
  
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Private Sub CmbServidor_Click()
  LinkearTexto CmbServidor.ListIndex
End Sub

Private Sub Form_Load()
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura
    
  CargarCombos
  LinkearTexto cRsl.TraerValorDeUnCampo("Sistema", "Email", "")

End Sub

Private Sub CargarCombos()
  CmbServidor.AddItem "Ninguno"
  CmbServidor.AddItem "Pop3"
  CmbServidor.AddItem "Hotmail"
  CmbServidor.AddItem "Gmail"
  CmbServidor.AddItem "Yahoo"
  CmbServidor.AddItem "General"
End Sub

Private Sub LinkearTexto(pServer As Byte)
  Dim cRsl As ClsLectura, Rs As Recordset
  
  Set cRsl = New ClsLectura
  
  CmbServidor.ListIndex = pServer
  
  Set Rs = cRsl.TraerRsCondi("EmailDatos", "", "Id=" & pServer)
  
  TxtServer.Enabled = False
  If CmbServidor.ListIndex = 1 Then
     TxtServer.Enabled = True
  End If
  
  If Rs.RecordCount <> 0 Then
     TxtServer.text = Rs!Servidor
     TxtUser.text = Rs!Usuario
     TxtPass.text = Rs!Contraseña
     TxtRemitente.text = Rs!Envia
     TxtAsunto.text = "" & Rs!Asunto
  Else
     TxtUser.text = ""
     TxtPass.text = ""
     TxtServer.text = ""
     TxtRemitente.text = ""
     TxtAsunto.text = ""
  End If
  
  
End Sub

