VERSION 5.00
Begin VB.Form CajaRimpZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reimprimir Reportes Z"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4680
   Begin VB.CommandButton CmbTerminar 
      Caption         =   "Terminar"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton CmdReimprimir 
         Caption         =   "Reimprimir"
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Text            =   "0000"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Número de Cierre Z :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "CajaRimpZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbTerminar_Click()
  Unload Me
End Sub


Private Sub CmdReimprimir_Click()
     If nPuerto <> 0 Then
        If nControlador = 1 Then
            Principal.HASAR1.Modelo = MODELO_P320
            Principal.HASAR1.Puerto = nPuerto
            Principal.HASAR1.Comenzar
            If Principal.HASAR1.HuboFaltaPapel = False Then
               Principal.HASAR1.TratarDeCancelarTodo
               Principal.HASAR1.ReporteZPorNumeros Val(TxtNumero.text), Val(TxtNumero.text), False
               Principal.HASAR1.CerrarComprobanteFiscal
               Principal.HASAR1.Finalizar
               MsgBox "Reimpresión Realizada con Exito", vbInformation, "Atención"
            Else
               MsgBox "Controlador sin Papel", vbCritical, "Atención"
            End If
        Else
            If nControlador = 2 Then
               Dim repuesta As Boolean
               Principal.PrinterFiscal1.PortNumber = nPuerto
               repuesta = Principal.PrinterFiscal1.Audit("Z", "T", TxtNumero.text, TxtNumero.text)
            Else
               Dim apertura As Boolean
               bAnswer = True
               Principal.EpsonFPHostControl1.ClosePort
               Principal.EpsonFPHostControl1.CommPort = nPuerto - 1
               Principal.EpsonFPHostControl1.BaudRate = EpsonFPHostControlX.TxBaudRate.br9600
               Principal.EpsonFPHostControl1.ProtocolType = EpsonFPHostControlX.TxProtocolType.protocol_Extended
               apertura = Principal.EpsonFPHostControl1.OpenPort
               If apertura = False Then MsgBox ("Serial Port Open Error")
               Call FPDelay
               
               sCmd = Chr$(&H8) + Chr$(&H13)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(sCmd)
               sCmdExt = Chr$(&H0) + Chr$(&H1)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(sCmdExt)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(TxtNumero.text)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(TxtNumero.text)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.SendCommand
               Call FPDelay
               If Principal.EpsonFPHostControl1.ReturnCode <> 0 Then ShowMsg
               
               sCmd = Chr$(&H8) + Chr$(&H14)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(sCmd)
               sCmdExt = Chr$(&H0) + Chr$(&H0)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(sCmdExt)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.SendCommand
               Call FPDelay
               If Principal.EpsonFPHostControl1.ReturnCode <> 0 Then ShowMsg
               
               sCmd = Chr$(&H8) + Chr$(&H15)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(sCmd)
               sCmdExt = Chr$(&H0) + Chr$(&H0)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.AddDataField(sCmdExt)
               If bAnswer Then bAnswer = Principal.EpsonFPHostControl1.SendCommand
               Call FPDelay
               If Principal.EpsonFPHostControl1.ReturnCode <> 0 Then
                  ShowMsg
               Else
                   MsgBox "Reimpresión Realizada con Exito", vbInformation, "Atención"
               End If
               
             End If
        End If
     End If
End Sub

Private Sub Form_Activate()
  TxtNumero.SetFocus
End Sub

Private Sub Form_Load()
  TxtNumero.text = "0000"
End Sub

Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub
Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPorcentaje_LostFocus()
  TxtNumero.text = Format(TxtNumero.text, "0000")
End Sub
