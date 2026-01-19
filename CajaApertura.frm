VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CajaApertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura de Caja"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   825
      Left            =   3135
      TabIndex        =   8
      Top             =   2160
      Width           =   1170
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   825
      Left            =   4380
      TabIndex        =   7
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   5490
      Begin VB.TextBox TxtSaldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "TxtSald"
         Top             =   1245
         Width           =   1380
      End
      Begin VB.ComboBox CmbCaja 
         Height          =   315
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   300
         Left            =   2130
         TabIndex        =   2
         Top             =   240
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   529
         _Version        =   393216
         Format          =   707985408
         CurrentDate     =   39378
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo de Apertura :"
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   1245
         Width           =   1620
      End
      Begin VB.Label Label2 
         Caption         =   "Transferir a Caja :"
         Height          =   210
         Left            =   195
         TabIndex        =   3
         Top             =   795
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   195
         TabIndex        =   1
         Top             =   285
         Width           =   615
      End
   End
End
Attribute VB_Name = "CajaApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()
  Dim Rx1 As Recordset, Rx2 As Recordset, Rx As Recordset, cRsl As ClsLectura, cRse As ClsComprobantesE
  Dim cNum As ClsComprobantesL, nMovi As Long
  
  Set cRsl = New ClsLectura
  Set cRse = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  
  nMovi = cNum.TraerUltimoMovimiento
  Set Rx1 = cRsl.RsVacio("Caja", "Id", "N")
  Set Rx = cRsl.TraerRsCondi("Caja", "Id", "Fecha='" & FechaAmericana(CajaCerrarZ.DtpFecha.Value) & "' AND Anulado=0 and Suc=" & nSucursal)
  If Rx.RecordCount <> 0 Then
     Do While Not Rx.EOF
        Rx1.AddNew
        Rx1!Suc = nSucursal
        Rx1!Movimiento = nMovi
        Rx1!Comprobante = Rx!Comprobante
        Rx1!Numero = Rx!Numero
        Rx1!Sucursal = Rx!Sucursal
        Rx1!fecha = Rx!fecha
        Rx1!Hora = Time
        Rx1!Caja = Rx!Caja
        Rx1!formapago = Rx!formapago
        Rx1!Nombre = Rx!Nombre
        Rx1!Banco = Rx!Banco
        Rx1!NumeroCheque = Rx!NumeroCheque
        Rx1!FechaAcreditacion = Rx!FechaAcreditacion
        Rx1!NumeroTarjeta = Rx!NumeroTarjeta
        Rx1!FechaVencimiento = Rx!FechaVencimiento
        Rx1!Autorizacion = Rx!Autorizacion
        Rx1!Cupon = Rx!Cupon
        Rx1!Debe = Rx!Haber
        Rx1!Haber = Rx!Debe
        Rx1!Abierta = Rx!Abierta
        Rx1!Anulado = 0
        Rx1.Update
        Rx.MoveNext
     Loop
     Rx.MoveFirst
  
     Do While Not Rx.EOF
        Rx1.AddNew
        Rx1!Suc = nSucursal
        Rx1!Movimiento = nMovi
        Rx1!Comprobante = Rx!Comprobante
        Rx1!Numero = Rx!Numero
        Rx1!Sucursal = Rx!Sucursal
        Rx1!fecha = Rx!fecha
        Rx1!Hora = Time
        Rx1!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
        Rx1!formapago = Rx!formapago
        Rx1!Nombre = Rx!Nombre
        Rx1!Banco = Rx!Banco
        Rx1!NumeroCheque = Rx!NumeroCheque
        Rx1!FechaAcreditacion = Rx!FechaAcreditacion
        Rx1!NumeroTarjeta = Rx!NumeroTarjeta
        Rx1!FechaVencimiento = Rx!FechaVencimiento
        Rx1!Autorizacion = Rx!Autorizacion
        Rx1!Cupon = Rx!Cupon
        Rx1!Debe = Rx!Debe
        If Rx!formapago = "EFECTIVO" Then
           Rx1!Debe = Rx!Debe - CCur(TxtSaldo.text)
        End If
        Rx1!Haber = Rx!Haber
        Rx1!Abierta = Rx!Abierta
        Rx1!Anulado = 0
        Rx1.Update
        Rx.MoveNext
     Loop
  End If
  
  cRse.GrabarSaldosCaja Rx1, DtpFecha.Value, TxtSaldo.text
  nDat = 1
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
   nDat = 0
   Unload Me
End Sub

Private Sub DtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
  bCajaSalida = True

  CargarCombos
  Limpiar
End Sub

Private Sub Limpiar()
  TxtSaldo.text = CajaCerrarZ.LblSub.Caption
  DtpFecha.Value = Date + 1
End Sub

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", "Caja<>1"
  If CmbCaja.ListCount <> 0 Then
     CmbCaja.ListIndex = 0
  Else
     MsgBox "Debe Crear una Caja de Salida", vbInformation, "Atención"
     bCajaSalida = False
     Unload Me
  End If
  
End Sub

Private Sub TxtSaldo_GotFocus()
  TxtSaldo.SelStart = 0
  TxtSaldo.SelLength = Len(TxtSaldo.text)
End Sub

Private Sub TxtSaldo_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtSaldo) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSaldo_LostFocus()
  If TxtSaldo.text = "" Then TxtSaldo.text = 0
  TxtSaldo.text = Format(TxtSaldo.text, "#0.00")
End Sub
