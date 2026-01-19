VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaInfResumen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Cuenta"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6345
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   180
      TabIndex        =   22
      Top             =   4455
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
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
      Height          =   4380
      Left            =   165
      TabIndex        =   6
      Top             =   45
      Width           =   5970
      Begin VB.CheckBox ChkSalto 
         Caption         =   "Sin Salto de Página"
         Height          =   195
         Left            =   2055
         TabIndex        =   26
         Top             =   4050
         Width           =   1905
      End
      Begin VB.CheckBox ChkTexto 
         Caption         =   "Imprime Texto"
         Height          =   225
         Left            =   60
         TabIndex        =   25
         Top             =   4035
         Width           =   1680
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aplicados"
         Height          =   300
         Index           =   1
         Left            =   4815
         TabIndex        =   24
         Top             =   255
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   300
         Index           =   0
         Left            =   3645
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.CheckBox ChkDesde 
         Caption         =   "Limitar a Fecha Desde"
         Height          =   195
         Left            =   2055
         TabIndex        =   23
         Top             =   3720
         Width           =   1965
      End
      Begin VB.ComboBox CmbZona 
         Height          =   315
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2790
         Width           =   2205
      End
      Begin VB.CheckBox ChkSimpli 
         Caption         =   "Simplificado"
         Height          =   225
         Left            =   4095
         TabIndex        =   19
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox ChkFecha 
         Caption         =   "Limitar a Fecha Hasta"
         Height          =   195
         Left            =   45
         TabIndex        =   18
         Top             =   3720
         Width           =   2070
      End
      Begin VB.CheckBox ChkSaldo 
         Caption         =   "Sin Saldo a Cero"
         Height          =   225
         Left            =   4095
         TabIndex        =   17
         Top             =   3375
         Width           =   1725
      End
      Begin VB.CheckBox ChkAbc 
         Caption         =   "Alfabetico"
         Height          =   210
         Left            =   2055
         TabIndex        =   16
         Top             =   3375
         Width           =   1155
      End
      Begin VB.CheckBox ChkRemitos 
         Caption         =   "Incluido Los Remitos"
         Height          =   225
         Left            =   45
         TabIndex        =   15
         Top             =   3375
         Width           =   1950
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   135
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "TxtCliente"
         Top             =   915
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4860
         TabIndex        =   8
         Top             =   915
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   135
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "TxtCliente"
         Top             =   1575
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4860
         TabIndex        =   7
         Top             =   1575
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   0
         Left            =   750
         TabIndex        =   3
         Top             =   2055
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   709951489
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   1
         Left            =   2985
         TabIndex        =   4
         Top             =   2055
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   709951489
         CurrentDate     =   37021
      End
      Begin VB.Label Label5 
         Caption         =   "Zona :"
         Height          =   255
         Left            =   285
         TabIndex        =   21
         Top             =   2790
         Width           =   615
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
         Left            =   1035
         TabIndex        =   14
         Top             =   915
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
         Left            =   1035
         TabIndex        =   13
         Top             =   1575
         Width           =   3705
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   2055
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   165
         Left            =   2430
         TabIndex        =   11
         Top             =   2055
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   645
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1005
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5175
      TabIndex        =   5
      Top             =   4500
      Width           =   930
   End
End
Attribute VB_Name = "VentaInfResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsClienteL
Public Rs As ADODB.Recordset

Private Sub CmdAplicar_Click()
  If Option1(0).Value = True Then
     If ChkSimpli.Value = 0 Then
        CrearRs
        If ChkDesde.Value = 0 Then
           TraerAnterior
         End If
         TraerSiguiente
         If ChkRemitos.Value = 1 Then
            TraerRemitos
         End If
         nImpr = 1
         FrmImpresor.Show
      Else
         CrearRsSimp
         TraerResumen
         nImpr = 55
         FrmImpresor.Show
      End If
  ElseIf Option1(1).Value = True Then
     CrearRsApli
     TraerRsApli
     nImpr = 65
     FrmImpresor.Show
  Else
     
  End If
End Sub

Private Sub CrearRsApli()
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Cliente", adVarChar, 100
  Rs.Fields.Append "FComprobante", adDate
  Rs.Fields.Append "Comprobante", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Numero", adVarChar, 15, adFldIsNullable
  Rs.Fields.Append "Importe", adDouble
  Rs.Fields.Append "FPago", adDate
  Rs.Fields.Append "Cobranza", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "RNumero", adVarChar, 15, adFldIsNullable
  Rs.Fields.Append "RImporte", adDouble
  Rs.Fields.Append "SaldoAnterior", adDouble
  Rs.Fields.Append "SaldoActual", adDouble
  Rs.Fields.Append "SaldoFecha", adDouble
  Rs.Open

End Sub

Private Sub TraerRsApli()
  Dim RsAux As ADODB.Recordset, cRsl As ClsLectura, cRx As ClsComprobantesL, j As Long, rRes As Recordset
  Dim sOrden As String
  
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set cRx = New ClsComprobantesL

  If TxtCliente(0).text <> TxtCliente(1).text Then
     sOrden = "Cliente"
     If ChkAbc.Value = 1 Then
        sOrden = "RazonSocial"
     End If
     Set RsAux = cRsl.TraerRsCondi("Clientes", sOrden, "RazonSocial>='" & LblCliente(0).Caption & "' AND RazonSocial<='" & LblCliente(1).Caption & "'")
  Else
     Set RsAux = cRsl.TraerRsCondi("Clientes", "Id", "Cliente='" & TxtCliente(0).text & "'")
  End If
  j = 1
  Dim nSaldo As Double
  ProgressBar1.Value = 1
  If RsAux.RecordCount <> 0 Then
     ProgressBar1.Max = RsAux.RecordCount
     Do While Not RsAux.EOF
        Set rRes = cRx.TraerFacturasAplicadas(DtpFecha(0).Value, DtpFecha(1).Value, RsAux!Cliente)
        nSaldo = 0
        If rRes.RecordCount <> 0 Then
           nSaldo = cRx.TraerSaldoAplica(DtpFecha(0).Value, DtpFecha(1).Value, RsAux!Cliente)
        End If
        Do While Not rRes.EOF
           Rs.AddNew
           Rs!Cliente = RsAux!Cliente & " " & RsAux!RazonSocial
           Rs!FComprobante = rRes!FechaFactura
           Rs!Comprobante = rRes!Descripcion
           Rs!Numero = Format(rRes!Numero, "00000000")
           Rs!importe = rRes!Debe
           If Not IsNull(rRes!Cobranza) Then
              Rs!FPago = rRes!FechaPago
              Rs!Cobranza = IIf(IsNull(rRes!Cobranza), "", rRes!Cobranza)
              Rs!RNumero = IIf(IsNull(rRes!Cobranza), "", Format(rRes!NumRec, "00000000"))
              Rs!RImporte = IIf(IsNull(rRes!importe), 0, rRes!importe)
           End If
           Rs!SaldoAnterior = 0
           Rs!SaldoActual = 0
           Rs!SaldoFecha = nSaldo
           Rs.Update
           rRes.MoveNext
        Loop
        ProgressBar1.Value = j
        j = j + 1
        RsAux.MoveNext
     Loop
'     If ChkAbc.Value = 1 Then
'        Rs.Sort = "Razon"
'     End If
  End If
  Set cRsl = Nothing
End Sub

Private Sub CrearRsSimp()
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Cliente", adVarChar, 15
  Rs.Fields.Append "Razon", adVarChar, 50
  Rs.Fields.Append "Domicilio", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Telefono", adVarChar, 40, adFldIsNullable
  Rs.Fields.Append "Saldo", adDouble
  Rs.Open

End Sub

Private Sub TraerResumen()
  Dim RsAux As ADODB.Recordset, cRsl As ClsClienteL, j As Long
  
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsClienteL

  
  Set RsAux = cRsl.ResumenDeCuentaSimpli(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value, ChkFecha.Value, CmbZona.ItemData(CmbZona.ListIndex))
  j = 1
  ProgressBar1.Value = 1
  If RsAux.RecordCount <> 0 Then
     ProgressBar1.Max = RsAux.RecordCount
  
     Do While Not RsAux.EOF
        Rs.AddNew
        Rs!Cliente = RsAux!Cliente
        Rs!Razon = RsAux!RazonSocial
        Rs!Domicilio = RsAux!Domicilio
        Rs!Telefono = RsAux!Telefono
        Rs!saldo = RsAux!saldo
        Rs.Update
        ProgressBar1.Value = j
        j = j + 1
        RsAux.MoveNext
     Loop
     If ChkAbc.Value = 1 Then
        Rs.Sort = "Razon"
     End If
  End If
  Set cRsl = Nothing

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
     If Index = 0 Then
        TxtCliente(1).text = Rs!Cliente
        LblCliente(1).Caption = Rs!RazonSocial
     End If
  End If
  Set cRsl = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  If Rs.RecordCount <> 0 Then
     TxtCliente(Index).text = Rs!Cliente
     LblCliente(Index).Caption = Rs!RazonSocial
  End If
End Sub
Private Sub DtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
     KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  DtpFecha(0).Value = Date - 30
  DtpFecha(1).Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
  ChkAbc.Value = 1
End Sub

Private Sub Form_Load()
  Set cRsl = New ClsClienteL
  Set Rs = New ADODB.Recordset

  Limpiar
  CargarEtiquetas
  
End Sub

Private Sub CargarEtiquetas()
  Dim Rz As ADODB.Recordset
  Dim cZona As ClsLectura

  Set Rz = New ADODB.Recordset
  Set Rs = cRsl.ClientesAlfabetico(0, False)
  Set cZona = New ClsLectura
  
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
  
  CmbZona.AddItem "TODAS"
  CmbZona.ItemData(CmbZona.NewIndex) = 0
  Set Rz = cZona.TraerRS("ZonaTraerTodos", , True)
  Do While Not Rz.EOF
     CmbZona.AddItem Rz!Descripcion
     CmbZona.ItemData(CmbZona.NewIndex) = Rz!Zona
     Rz.MoveNext
  Loop
  CmbZona.ListIndex = 0

End Sub
Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
  CmdAplicar.Enabled = False
  LblCliente(Index).Caption = ""
  Set cRsl = New ClsClienteL
  If TxtCliente(Index).text <> "" Then
     If cRsl.BuscarNombreCliente(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cRsl.BuscarNombreCliente(TxtCliente(Index).text)
        TxtCliente(1).text = TxtCliente(Index).text
        LblCliente(1).Caption = LblCliente(Index).Caption
        CmdAplicar.Enabled = True
     End If
  End If
End Sub

Private Sub CrearRs()
  
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Cliente", adVarChar, 15
  Rs.Fields.Append "Razon", adVarChar, 50
  Rs.Fields.Append "Domicilio", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Postal", adVarChar, 8, adFldIsNullable
  Rs.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Iva", adVarChar, 25, adFldIsNullable
  Rs.Fields.Append "Cuit", adVarChar, 13, adFldIsNullable
  Rs.Fields.Append "Telefono", adVarChar, 40, adFldIsNullable
  Rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs.Fields.Append "Comprobante", adVarChar, 40, adFldIsNullable
  Rs.Fields.Append "Numero", adDouble, , adFldIsNullable
  Rs.Fields.Append "Sucursal", adInteger, , adFldIsNullable
  Rs.Fields.Append "Cuota", adInteger, , adFldIsNullable
  Rs.Fields.Append "Debe", adDouble
  Rs.Fields.Append "Haber", adDouble
  Rs.Fields.Append "Cobrador", adInteger
  Rs.Open
End Sub

Private Sub TraerAnterior()
  Dim RsAux As ADODB.Recordset
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsClienteL
  
  Set RsAux = cRsl.ResumenSaldoAnterior(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, CmbZona.ItemData(CmbZona.ListIndex))
  
  Dim sSaldo As Currency
  
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Cliente = RsAux!Cliente
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = RsAux!CodigoPostal
     Rs!Localidad = RsAux!Ciudad
     Rs!Provincia = RsAux!Provincia
     Rs!Iva = RsAux!TipoIva
     Rs!Cuit = RsAux!Cuit
     Rs!Telefono = RsAux!Telefono
     Rs!fecha = Null
     Rs!Comprobante = "SALDO ANTERIOR"
     Rs!Numero = Null
     Rs!Sucursal = Null
     Rs!Cuota = Null
     sSaldo = IIf(IsNull(RsAux!SumaDeDebe), 0, RsAux!SumaDeDebe) - IIf(IsNull(RsAux!SumaDeHaber), 0, RsAux!SumaDeHaber)
     Rs!Debe = sSaldo 'IIf(IsNull(RsAux!SumaDeDebe), 0, RsAux!SumaDeDebe)
     Rs!Haber = 0      ' IIf(IsNull(RsAux!SumaDeHaber), 0, RsAux!SumaDeHaber)
     Rs!Cobrador = Format(RsAux!Zona, "00")
     Rs.Update
     RsAux.MoveNext
  Loop
  
End Sub

Private Sub TraerSiguiente()
  Dim RsAux As ADODB.Recordset, j As Long, cRcli As ClsClienteL
  Set RsAux = New ADODB.Recordset
  
  Set cRcli = New ClsClienteL
  Set RsAux = cRcli.ResumenDeCuenta(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value, ChkFecha.Value, CmbZona.ItemData(CmbZona.ListIndex))
   
  j = 1
  ProgressBar1.Value = 1
  If RsAux.RecordCount <> 0 Then
     ProgressBar1.Max = RsAux.RecordCount
  End If
  Do While Not RsAux.EOF
     If ChkDesde.Value = 0 Then
        Rs.AddNew
        Rs!Cliente = RsAux!Cliente
        Rs!Razon = RsAux!RazonSocial
        Rs!Domicilio = RsAux!Domicilio
        Rs!Postal = "" & RsAux!CodigoPostal
        Rs!Localidad = RsAux!Ciudad
        Rs!Provincia = RsAux!Provincia
        Rs!Iva = RsAux!TipoIva
        Rs!Cuit = RsAux!Cuit
        Rs!Telefono = RsAux!Telefono
        Rs!fecha = RsAux!fecha
        Rs!Comprobante = RsAux!Comprobante '& " " & RsAux!Motivo
        Rs!Numero = RsAux!Numero
        Rs!Sucursal = RsAux!Sucursal
        Rs!Cuota = RsAux!Cuota
        Rs!Debe = IIf(IsNull(RsAux!Debe), 0, RsAux!Debe)
        Rs!Haber = IIf(IsNull(RsAux!Haber), 0, RsAux!Haber)
        Rs!Cobrador = Format(RsAux!Zona, "00")
        Rs.Update
     Else
        If RsAux!tipo <> 4 Then
           Rs.AddNew
           Rs!Cliente = RsAux!Cliente
           Rs!Razon = RsAux!RazonSocial
           Rs!Domicilio = RsAux!Domicilio
           Rs!Postal = "" & RsAux!CodigoPostal
           Rs!Localidad = RsAux!Ciudad
           Rs!Provincia = RsAux!Provincia
           Rs!Iva = RsAux!TipoIva
           Rs!Cuit = RsAux!Cuit
           Rs!Telefono = RsAux!Telefono
           Rs!fecha = RsAux!fecha
           Rs!Comprobante = RsAux!Comprobante ' & " " & RsAux!Motivo
           Rs!Numero = RsAux!Numero
           Rs!Sucursal = RsAux!Sucursal
           Rs!Cuota = RsAux!Cuota
           Rs!Debe = IIf(IsNull(RsAux!Debe), 0, RsAux!Debe)
           Rs!Haber = IIf(IsNull(RsAux!Haber), 0, RsAux!Haber)
           Rs!Cobrador = Format(RsAux!Zona, "00")
           Rs.Update
     End If
     End If
     ProgressBar1.Value = j
     j = j + 1
     RsAux.MoveNext
  Loop
  If ChkAbc.Value = 1 Then
     Rs.Sort = "Razon"
  End If
   Set cRcli = Nothing
End Sub

Private Sub TraerRemitos()
  Dim rAux As ADODB.Recordset, j As Long, nCantRem As Double, nCantRemAux As Double, RsAux As Recordset
  
  Set rAux = New ADODB.Recordset
  Set RsAux = New Recordset
  Dim cRem As New ClsComprobantesL
      
  Set RsAux = cRem.TraerInfRemitosPend(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value, CmbZona.ItemData(CmbZona.ListIndex))
   j = 1
   ProgressBar1.Value = 1
   ProgressBar1.Max = IIf(RsAux.RecordCount = 0, 1, RsAux.RecordCount)

   Do While Not RsAux.EOF
     Set rAux = cRem.TraerDetallesMovimiento(RsAux!Movimiento)
     Do While Not rAux.EOF
        nCantRem = Round(cRem.TraerCantidadRemito(rAux!Producto, RsAux!Movimiento, rAux!ID), 2)
        nCantRemAux = Round(rAux!Cantidad - Abs(nCantRem), 2)
        If nCantRemAux > 0 Or nCantRem = -1 Then
           Rs.AddNew
           Rs!Cliente = RsAux!Cliente
           Rs!Razon = RsAux!RazonSocial
           Rs!fecha = RsAux!fecha
           Rs!Comprobante = RsAux!Comprobante
           Rs!Numero = RsAux!Numero
           Rs!Sucursal = RsAux!Sucursal
           Rs!Cuota = Null
           Rs!Debe = IIf(IsNull(RsAux!Expr1), 0, Abs(RsAux!Expr1))
           Rs!Haber = 0
           Rs.Update
           Exit Do
        End If
        rAux.MoveNext
     Loop
     RsAux.MoveNext
     ProgressBar1.Value = j
     j = j + 1
  Loop
End Sub

