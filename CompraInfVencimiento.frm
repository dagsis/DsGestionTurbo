VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CompraInfVencimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vencimientos de Pagos"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   6210
   Begin VB.Frame Frame1 
      Caption         =   "Proveedores"
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
      Left            =   150
      TabIndex        =   1
      Top             =   165
      Width           =   5850
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   195
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "TxtCliente"
         Top             =   660
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4920
         TabIndex        =   4
         Top             =   660
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   195
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "TxtCliente"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4920
         TabIndex        =   2
         Top             =   1320
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPVencimiento 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   708837377
         CurrentDate     =   36871
      End
      Begin VB.Label Label4 
         Caption         =   "Con Vencimiento Hasta :"
         Height          =   240
         Left            =   195
         TabIndex        =   10
         Top             =   1785
         Width           =   1845
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
         TabIndex        =   9
         Top             =   660
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
         Index           =   1
         Left            =   1095
         TabIndex        =   8
         Top             =   1320
         Width           =   3705
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Proveedor :"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   390
         Width           =   1350
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   405
      Left            =   4785
      TabIndex        =   0
      Top             =   3120
      Width           =   1200
   End
End
Attribute VB_Name = "CompraInfVencimiento"
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
  Me.MousePointer = 11
  CrearRs1
  CrearRs2
  Set Rs = cCliente.ProveedorListados(False, LblCliente(0).Caption, LblCliente(1).Caption)
  Do While Not Rs.EOF
     LlenarGrilla
     If RsC.RecordCount <> 0 Then
        RsC.MoveFirst
        Do While Not RsC.EOF
           If RsC!fecha <= DTPVencimiento.Value Then
              RsP.AddNew
              RsP!Proveedor = Rs!Proveedor
              RsP!Razon = Rs!RazonSocial
              RsP!Domicilio = Rs!Domicilio
              RsP!Postal = "" & Rs!CodigoPostal
              RsP!Localidad = Rs!Localidad
              RsP!Provincia = "" ' rs![Provincias.Descripcion]
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
  Loop
  nImpr = 32
  FrmImpresor.Show
  Me.MousePointer = 0
End Sub

Private Sub Form_Load()
  
  Set Rz = New ADODB.Recordset
  Set Rs = New ADODB.Recordset
  
  Set cZona = New ClsLectura
  Set cCliente = New ClsClienteL
  
  DTPVencimiento.Value = Date
  
  Limpiar
  CargarEtiquetas
End Sub

Private Sub CrearRs1()
 On Error GoTo errHandler

 Set RsC = New ADODB.Recordset
  
 RsC.Fields.Append "Fecha", adDate
 RsC.Fields.Append "Comprobante", adVarChar, 15
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
  
 RsP.Fields.Append "Proveedor", adVarChar, 15
 RsP.Fields.Append "Razon", adVarChar, 45
 RsP.Fields.Append "Domicilio", adVarChar, 45, adFldIsNullable
 RsP.Fields.Append "Postal", adVarChar, 8, adFldIsNullable
 RsP.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
 RsP.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
 RsP.Fields.Append "Fecha", adDate
 RsP.Fields.Append "Comprobante", adVarChar, 15
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
 
 Set cCom = New ClsComprobantesL
 Set rCuen = New ADODB.Recordset
 Set cRx = New ClsLectura
 
 Set rCuen = cCom.DameCuentaCorrienteP(Rs!Proveedor)
 CrearRs1
 Set db1 = New Connection
 db1.Open sDb
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = Left(cCom.DameCodComprobante(rCuen!Comprobante), 15)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    nCoti = cRx.TraerValorDeUnCampo("CabComprobantes", "Cotizacion", "Movimiento=" & rCuen!Venta)
    RsC!importe = Format(rCuen!Haber * nCoti, nCantDecimales)
    cCobrado = cCom.DamePagos(rCuen!Cuota, rCuen!Venta)
    RsC!Cobrado = cCobrado
    RsC!Cobrar = Format(rCuen!Haber * nCoti, nCantDecimales) - RsC!Cobrado
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    RsC.Update
    If RsC!Cobrar = 0 Then
       RsC.Delete
    End If
    rCuen.MoveNext
 Loop
 db1.Close
 
 Set rCuen = cCom.DameCuentaCorrienteDA(Rs!Proveedor)
 
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = cCom.DameCodComprobante(rCuen!Comprobante)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    RsC!importe = Format(rCuen!Debe, "#0.00")
    cCobrado = cCom.DamePagosA(rCuen!Venta)
    RsC!Cobrado = cCobrado
    RsC!Cobrar = Format(rCuen!Debe, "#0.00") - RsC!Cobrado
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    If RsC!Cobrar = 0 Then
       RsC.CancelUpdate
    Else
       RsC!Cobrado = (Format(rCuen!Debe, "#0.00") - RsC!Cobrado)
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
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim crs As ClsLectura
  Set crs = New ClsLectura
  nDat = 0
  nBuscar = 2
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Proveedores", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ProveedorTraerUno", nDat, True)
     LinkearTexto Index
  End If
  Set crs = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  TxtCliente(Index).text = Rs!Proveedor
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
  Set Rs = cCliente.ProveedorAlfabetico
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Proveedor
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Proveedor
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
     If cCliente.BuscarNombreProveedor(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cCliente.BuscarNombreProveedor(TxtCliente(Index).text)
        CmdImprimir.Enabled = True
     End If
  End If
End Sub


