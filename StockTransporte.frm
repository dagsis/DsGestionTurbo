VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form StockTransporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transporte"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptEstado 
      Caption         =   "En Espera"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   27
      Top             =   5145
      Width           =   1155
   End
   Begin VB.OptionButton OptEstado 
      Caption         =   "Pendiente"
      Height          =   195
      Index           =   1
      Left            =   1335
      TabIndex        =   26
      Top             =   5145
      Width           =   1155
   End
   Begin VB.OptionButton OptEstado 
      Caption         =   "Entregado"
      Height          =   195
      Index           =   2
      Left            =   2550
      TabIndex        =   25
      Top             =   5145
      Width           =   1155
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Height          =   690
      Index           =   1
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   0
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   135
      TabIndex        =   15
      Top             =   30
      Width           =   5745
      Begin VB.TextBox TxtTelefono 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "TxtTelefono"
         Top             =   2190
         Width           =   4155
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "TxtDireccion"
         Top             =   2970
         Width           =   4155
      End
      Begin VB.TextBox TxtHora 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "TxtHora"
         Top             =   645
         Width           =   4155
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   225
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Format          =   308871169
         CurrentDate     =   42650
      End
      Begin VB.TextBox TxtLocalidad 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "TxtLocalidad"
         Top             =   1800
         Width           =   4155
      End
      Begin VB.TextBox TxtCuit 
         Height          =   285
         Left            =   1350
         MaxLength       =   13
         TabIndex        =   7
         Text            =   "TxtCuit"
         Top             =   2580
         Width           =   1530
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "TxtDireccion"
         Top             =   1410
         Width           =   4155
      End
      Begin VB.TextBox TxtDestinatario 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "TxtDestinatario"
         Top             =   1050
         Width           =   4155
      End
      Begin VB.TextBox TxtPatenteAcoplado 
         Height          =   285
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "TxtPatenteAcoplado"
         Top             =   4500
         Width           =   1335
      End
      Begin VB.TextBox TxtPatenteCamion 
         Height          =   285
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "TxtPatente"
         Top             =   4500
         Width           =   1335
      End
      Begin VB.TextBox TxtDomicilio 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "TxtDomicilio"
         Top             =   3780
         Width           =   4155
      End
      Begin VB.TextBox TxtCuitTrans 
         Height          =   285
         Left            =   1350
         MaxLength       =   13
         TabIndex        =   11
         Text            =   "TxtCuitTrans"
         Top             =   4125
         Width           =   1335
      End
      Begin VB.TextBox TxtTransporte 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "TxtTransporte"
         Top             =   3420
         Width           =   4155
      End
      Begin VB.Label Label14 
         Caption         =   "Teléfono  :"
         Height          =   210
         Left            =   165
         TabIndex        =   32
         Top             =   2190
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "Observacion  :"
         Height          =   210
         Left            =   135
         TabIndex        =   30
         Top             =   2970
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "H.de Entrega :"
         Height          =   210
         Left            =   165
         TabIndex        =   29
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label10 
         Caption         =   "F.de Entrega :"
         Height          =   210
         Left            =   165
         TabIndex        =   28
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label9 
         Caption         =   "Localidad  :"
         Height          =   210
         Left            =   165
         TabIndex        =   24
         Top             =   1815
         Width           =   1065
      End
      Begin VB.Label Label8 
         Caption         =   "CUIT :"
         Height          =   210
         Left            =   165
         TabIndex        =   23
         Top             =   2595
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Domicilio  :"
         Height          =   210
         Left            =   165
         TabIndex        =   22
         Top             =   1425
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Destinatario :"
         Height          =   210
         Left            =   165
         TabIndex        =   21
         Top             =   1065
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "Domicilio :"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   3795
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "CUIT :"
         Height          =   240
         Left            =   165
         TabIndex        =   19
         Top             =   4140
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "P.Acoplado :"
         Height          =   255
         Left            =   2850
         TabIndex        =   18
         Top             =   4500
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "P.Camión :"
         Height          =   255
         Left            =   165
         TabIndex        =   17
         Top             =   4485
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Transporte :"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   3435
         Width           =   975
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Localidad  :"
      Height          =   210
      Left            =   285
      TabIndex        =   31
      Top             =   2235
      Width           =   1065
   End
End
Attribute VB_Name = "StockTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset, pMovi As Long

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
            Unload Me
         Case 1
            Grabar
            CmdBotones(1).Enabled = False
  End Select
End Sub

Private Sub Grabar()
  Dim cRle As ClsComprobantesE, nEstado As Byte
  
  Set cRle = New ClsComprobantesE
  rs!Movimiento = pMovi
  rs!fecha = DTPFecha.Value
  rs!Hora = TxtHora.text
  rs!Destinatario = TxtDestinatario.text
  rs!Direccion = TxtDireccion.text
  rs!Localidad = "" & TxtLocalidad.text
  rs!Telefono = "" & TxtTelefono.text
  rs!Cuit = TxtCuit.text
  rs!Transporte = TxtTransporte.text
  rs!Domicilio = TxtDomicilio.text
  rs!CuitTrans = TxtCuitTrans.text
  rs!PatenteCamion = TxtPatenteCamion.text
  rs!PatenteAcoplado = TxtPatenteAcoplado.text
  rs!Obsrevaciones = TxtObservaciones.text
  
  If OptEstado(0).Value = True Then nEstado = 1
  If OptEstado(1).Value = True Then nEstado = 2
  If OptEstado(2).Value = True Then nEstado = 3
  
  rs!Estado = nEstado
  rs.Update
  
  cRle.GrabarTransporte rs, pMovi
  LlenarRs
  Set cRle = Nothing
End Sub

Private Sub LlenarRs()
  Dim cRsl As ClsLectura, sCli As String, nTra As Integer
  
  Set cRsl = New ClsLectura
  sCli = VentaCompRemitos.nClient
  Set rs = cRsl.TraerRsCondi("RegTransporte", "Id", "Movimiento=" & pMovi)
  
  nTra = cRsl.TraerValorDeUnCampo("Clientes", "Transporte", "Cliente='" & sCli & "'")
  
  If rs.RecordCount = 0 Then
     rs.AddNew
     rs!Movimiento = pMovi
     rs!fecha = DTPFecha.Value
     rs!Hora = ""
     rs!Destinatario = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & sCli & "'")
     rs!Direccion = cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & sCli & "'")
     rs!Localidad = cRsl.TraerValorDeUnCampo("Clientes", "Ciudad", "Cliente='" & sCli & "'")
     rs!Cuit = cRsl.TraerValorDeUnCampo("Clientes", "Cuit", "Cliente='" & sCli & "'")
     rs!Transporte = cRsl.TraerValorDeUnCampo("Transporte", "Descripcion", "Id=" & nTra)
     rs!Domicilio = cRsl.TraerValorDeUnCampo("Transporte", "Domicilio", "Id=" & nTra)
     rs!CuitTrans = cRsl.TraerValorDeUnCampo("Transporte", "Cuit", "Id=" & nTra)
     rs!PatenteCamion = cRsl.TraerValorDeUnCampo("Transporte", "PatenteChasis", "Id=" & nTra)
     rs!PatenteAcoplado = cRsl.TraerValorDeUnCampo("Transporte", "PatenteAcoplado", "Id=" & nTra)
     rs!Obsrevaciones = ""
     rs!Telefono = cRsl.TraerValorDeUnCampo("Clientes", "Telefono", "Cliente='" & sCli & "'")
     rs!Estado = 1
     rs.Update
  End If
  
  DTPFecha.Value = rs!fecha
  TxtHora.text = "" & rs!Hora
  TxtDestinatario.text = "" & rs!Destinatario
  TxtDireccion.text = "" & rs!Direccion
  TxtLocalidad.text = "" & rs!Localidad
  TxtTelefono.text = "" & rs!Telefono
  TxtCuit.text = "" & rs!Cuit
  TxtTransporte.text = "" & rs!Transporte
  TxtDomicilio.text = "" & rs!Domicilio
  TxtCuitTrans.text = "" & rs!CuitTrans
  TxtPatenteCamion.text = "" & rs!PatenteCamion
  TxtPatenteAcoplado.text = "" & rs!PatenteAcoplado
  TxtObservaciones.text = "" & rs!Obsrevaciones
  
  If rs!Estado = 1 Then OptEstado(0).Value = True
  If rs!Estado = 2 Then OptEstado(1).Value = True
  If rs!Estado = 3 Then OptEstado(2).Value = True
    
  Set cRsl = Nothing
End Sub

Private Sub DtpFecha_Change()
  TodoBien
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
 CmdBotones(0).Picture = LoadResPicture("Salir", 0)
 CmdBotones(1).Picture = LoadResPicture("Grabar", 0)
 Limpiar
 pMovi = VentaCompRemitos.nMoviTrans
 LlenarRs
 CmdBotones(1).Enabled = False
End Sub

Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  DTPFecha.Value = Date
End Sub

Private Sub TodoBien()
  CmdBotones(1).Enabled = False
  If Len(TxtDestinatario.text) <> 0 And Len(TxtTelefono.text) <> 0 And Len(TxtDireccion.text) <> 0 Then
     CmdBotones(1).Enabled = True
  End If
End Sub

Private Sub OptEstado_Click(Index As Integer)
  TodoBien
End Sub

Private Sub TxtCuit_Change()
  TodoBien
End Sub

Private Sub TxtCuitTrans_Change()
  TodoBien
End Sub

Private Sub TxtDestinatario_Change()
  TodoBien
End Sub

Private Sub TxtDireccion_Change()
  TodoBien
End Sub

Private Sub TxtDomicilio_Change()
  TodoBien
End Sub

Private Sub TxtHora_Change()
  TodoBien
End Sub

Private Sub TxtLocalidad_Change()
  TodoBien
End Sub

Private Sub TxtObservaciones_Change()
  TodoBien
End Sub

Private Sub TxtPatenteAcoplado_Change()
  TodoBien
End Sub

Private Sub TxtPatenteCamion_Change()
  TodoBien
End Sub

Private Sub TxtTelefono_Change()
  TodoBien
End Sub

Private Sub TxtTransporte_Change()
  TodoBien
End Sub

