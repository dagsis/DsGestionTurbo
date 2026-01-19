VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaInfVencimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vencimientos de Cuotas"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   6045
   Begin VB.CheckBox ChkOrd 
      Caption         =   "Alfabetico"
      Height          =   210
      Left            =   2910
      TabIndex        =   17
      Top             =   3090
      Width           =   1530
   End
   Begin VB.CheckBox ChkPago 
      Caption         =   "Descartar los Pagos Parciales"
      Height          =   210
      Left            =   150
      TabIndex        =   16
      Top             =   3090
      Width           =   2625
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   405
      Left            =   4680
      TabIndex        =   5
      Top             =   3015
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   5850
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   2505
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4920
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   195
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "TxtCliente"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   660
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   195
         MaxLength       =   15
         TabIndex        =   6
         Text            =   "TxtCliente"
         Top             =   660
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPVencimiento 
         Height          =   315
         Left            =   4305
         TabIndex        =   4
         Top             =   2055
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709885953
         CurrentDate     =   36871
      End
      Begin VB.ComboBox CmbZona 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2055
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   2865
         TabIndex        =   14
         Top             =   2055
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709885953
         CurrentDate     =   36871
      End
      Begin VB.Label Label2 
         Caption         =   "Desde :"
         Height          =   240
         Left            =   2895
         TabIndex        =   15
         Top             =   1785
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
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
         Index           =   1
         Left            =   1095
         TabIndex        =   11
         Top             =   1320
         Width           =   3705
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
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
         Index           =   0
         Left            =   1095
         TabIndex        =   10
         Top             =   660
         Width           =   3705
      End
      Begin VB.Label Label4 
         Caption         =   " Hasta :"
         Height          =   240
         Left            =   4335
         TabIndex        =   3
         Top             =   1785
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Zona :"
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   1785
         Width           =   1215
      End
   End
End
Attribute VB_Name = "VentaInfVencimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsP As ADODB.Recordset
Dim Rz As ADODB.Recordset
Dim Rs As ADODB.Recordset, RsC As ADODB.Recordset
Dim cZona As ClsLectura
Dim cCliente As ClsClienteL

Private Sub CmdImprimir_Click()
  Dim i As Integer
  Me.MousePointer = 11
  CrearRs1
  CrearRs2
  Set Rs = cCliente.ClientesListados(ChkOrd.Value, LblCliente(0).Caption, LblCliente(1).Caption, CmbZona.text, CmbZona.text)
  i = 0
  If Rs.RecordCount <> 0 Then
     ProgressBar1.Min = 0
     ProgressBar1.Max = Rs.RecordCount
  End If
  ProgressBar1.Value = 0
  Do While Not Rs.EOF
     LlenarGrilla
     If RsC.RecordCount <> 0 Then
        RsC.MoveFirst
        Do While Not RsC.EOF
           If RsC!fecha >= DtpDesde.Value And RsC!fecha <= DTPVencimiento.Value Then
              RsP.AddNew
              RsP!Cliente = Rs!Cliente
              RsP!Razon = Rs!RazonSocial
              RsP!Domicilio = Rs!Domicilio
              RsP!Postal = 0 'IIf(rs!CodigoPostal = "", 0, rs!CodigoPostal)
              RsP!Localidad = Rs!Ciudad
              RsP!Provincia = Rs!Provincia
              RsP!fecha = RsC!fecha
              RsP!Comprobante = RsC!Comprobante
              RsP!Numero = RsC!Numero
              RsP!Cuota = RsC!Cuota
              RsP!importe = RsC!importe
              RsP!Cobrado = RsC!Cobrado
              RsP!Cobrar = RsC!Cobrar
              RsP!Venta = RsC!Venta
              RsP!Sucursal = RsC!Sucursal
              RsP.Update
           End If
           RsC.MoveNext
           Loop
      End If
      Rs.MoveNext
      i = i + 1
      ProgressBar1.Value = i
  Loop
  nImpr = 2
  FrmImpresor.Show
  Me.MousePointer = 0
End Sub

Private Sub DTPVencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
  
  Set Rz = New ADODB.Recordset
  Set Rs = New ADODB.Recordset
  
  Set cZona = New ClsLectura
  Set cCliente = New ClsClienteL
  
  DTPVencimiento.Value = Date
  
  CmbZona.AddItem "TODAS"
  CmbZona.ItemData(CmbZona.NewIndex) = 0
  Set Rz = cZona.TraerRS("ZonaTraerTodos", , True)
  Do While Not Rz.EOF
     CmbZona.AddItem Rz!Descripcion
     CmbZona.ItemData(CmbZona.NewIndex) = Rz!Zona
     Rz.MoveNext
  Loop
  CmbZona.ListIndex = 0
  Limpiar
  CargarEtiquetas
End Sub

Private Sub CrearRs1()
 On Error GoTo errHandler

 Set RsC = New ADODB.Recordset
  
 RsC.Fields.Append "Fecha", adDate
 RsC.Fields.Append "Comprobante", adVarChar, 30
 RsC.Fields.Append "Numero", adDouble
 RsC.Fields.Append "Cuota", adInteger
 RsC.Fields.Append "Importe", adCurrency
 RsC.Fields.Append "Cobrado", adCurrency
 RsC.Fields.Append "Cobrar", adCurrency
 RsC.Fields.Append "Venta", adDouble
 RsC.Fields.Append "Sucursal", adInteger
 RsC.CursorType = adOpenKeyset
 RsC.LockType = adLockOptimistic
 RsC.Open
  
 Exit Sub
errHandler:
   ManejaErrores

End Sub

Private Sub CrearRs2()
 On Error GoTo errHandler

 Set RsP = New ADODB.Recordset
  
 RsP.Fields.Append "Cliente", adVarChar, 15
 RsP.Fields.Append "Razon", adVarChar, 45
 RsP.Fields.Append "Domicilio", adVarChar, 45, adFldIsNullable
 RsP.Fields.Append "Postal", adInteger
 RsP.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
 RsP.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
 RsP.Fields.Append "Fecha", adDate
 RsP.Fields.Append "Comprobante", adVarChar, 30
 RsP.Fields.Append "Numero", adDouble
 RsP.Fields.Append "Cuota", adInteger
 RsP.Fields.Append "Importe", adCurrency
 RsP.Fields.Append "Cobrado", adCurrency
 RsP.Fields.Append "Cobrar", adCurrency
 RsP.Fields.Append "Venta", adDouble
 RsP.Fields.Append "Sucursal", adInteger
 RsP.CursorType = adOpenKeyset
 RsP.LockType = adLockOptimistic
 RsP.Open
  
 Exit Sub
errHandler:
   ManejaErrores

End Sub

Private Sub LlenarGrilla()
 Dim cCom As ClsComprobantesL, rCuen As ADODB.Recordset
 Dim sDebe As Currency, sCobrado As Currency, sCobrar As Currency
 Dim cCobrado As Currency, cAux As Currency, nCoti As Currency
 Dim cRx As ClsLectura
 
 Set cRx = New ClsLectura
 Set cCom = New ClsComprobantesL
 Set rCuen = New ADODB.Recordset
 
 Set rCuen = cCom.DameCuentaCorriente(Rs!Cliente)
 CrearRs1
 Set db1 = New Connection
 db1.Open sDb
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = cCom.DameCodComprobante(rCuen!Comprobante)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    nCoti = cRx.TraerValorDeUnCampo("CabComprobantes", "Cotizacion", "Movimiento=" & rCuen!Venta)
    RsC!importe = Format(rCuen!Debe * nCoti, nCantDecimales)
    cCobrado = cCom.DamePagos(rCuen!Cuota, rCuen!Venta)
    RsC!Cobrado = cCobrado
    RsC!Cobrar = Format(rCuen!Debe * nCoti, nCantDecimales) - RsC!Cobrado
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    RsC.Update
    If RsC!Cobrar = 0 Then
       RsC.Delete
    End If
    rCuen.MoveNext
 Loop
 db1.Close
 If ChkPago.Value = 1 Then
    Set rCuen = cCom.DameCuentaCorrienteHAConFecha(Rs!Cliente, DtpDesde.Value)
    If rCuen.RecordCount <> 0 Then
       If RsC.RecordCount <> 0 Then
          RsC.MoveLast
          Do While Not RsC.BOF
             RsC.Delete
             RsC.MovePrevious
          Loop
       End If
    End If
 Else
   Set rCuen = cCom.DameCuentaCorrienteHA(Rs!Cliente)
 End If
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = cCom.DameCodComprobante(rCuen!Comprobante)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    RsC!importe = Format(rCuen!Haber, "#0.00")
    cCobrado = cCom.DamePagosA(rCuen!Venta)
    RsC!Cobrado = cCobrado
    RsC!Cobrar = Format(rCuen!Haber, "#0.00") - RsC!Cobrado
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    If RsC!Cobrar = 0 Then
       RsC.CancelUpdate
    Else
       RsC!Cobrado = (Format(rCuen!Haber, "#0.00") - RsC!Cobrado)
       RsC!importe = 0
       RsC!Cobrar = -RsC!Cobrado
       RsC.Update
    End If
    rCuen.MoveNext
 Loop

End Sub

Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  DtpDesde.Value = Date
  DTPVencimiento.Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim crs As ClsLectura
  Set crs = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Clientes", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ClienteTraerUno", nDat, True)
     LinkearTexto Index
  End If
  Set crs = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  TxtCliente(Index).text = Rs!Cliente
  LblCliente(Index).Caption = Rs!RazonSocial
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
      Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
     KeyAscii = Upper(KeyAscii)
End Sub
Private Sub CargarEtiquetas()
  Set Rs = cCliente.ClientesAlfabetico(0, False)
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
End Sub
Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
  CmdImprimir.Enabled = False
  LblCliente(Index).Caption = ""
  If TxtCliente(Index).text <> "" Then
     If cCliente.BuscarNombreCliente(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cCliente.BuscarNombreCliente(TxtCliente(Index).text)
        CmdImprimir.Enabled = True
     End If
  End If
End Sub

