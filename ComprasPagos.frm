VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos a Proveedores"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   10395
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   90
      TabIndex        =   32
      Top             =   6030
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   720
      Index           =   2
      Left            =   9420
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5460
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Valores"
      Height          =   720
      Index           =   1
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5460
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Grabar"
      Height          =   720
      Index           =   0
      Left            =   7545
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5460
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta Corriente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   10275
      Begin VB.Frame Frame2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   8535
         TabIndex        =   33
         Top             =   195
         Width           =   1635
         Begin VB.OptionButton Option1 
            Caption         =   "Normal"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Detalle"
         Height          =   285
         Left            =   7605
         TabIndex        =   9
         Top             =   1080
         Width           =   810
      End
      Begin VB.TextBox TxtEfectivo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3615
         TabIndex        =   3
         Text            =   "TxtEfectivo"
         Top             =   1455
         Width           =   915
      End
      Begin VB.ComboBox CmbCaja 
         Height          =   315
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1470
         Width           =   2190
      End
      Begin VB.ComboBox CmbComprobante 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   2445
      End
      Begin VB.CommandButton CmdClientes 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   7605
         TabIndex        =   8
         Top             =   750
         Width           =   810
      End
      Begin VB.TextBox TxtCliente 
         Height          =   315
         Left            =   1110
         TabIndex        =   1
         Text            =   "TxtCliente"
         Top             =   750
         Width           =   735
      End
      Begin VB.TextBox TxtCancelar 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5790
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "TxtCancela"
         Top             =   1455
         Width           =   810
      End
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "Aplicar"
         Height          =   285
         Left            =   6675
         TabIndex        =   7
         Top             =   1455
         Width           =   870
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   315
         Left            =   7140
         TabIndex        =   10
         Top             =   270
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   708968449
         CurrentDate     =   36821
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   3000
         Left            =   180
         TabIndex        =   11
         Top             =   1965
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5292
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label LblCancelado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCancelado"
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
         Left            =   7155
         TabIndex        =   27
         Top             =   5040
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "A Pagar :"
         Height          =   195
         Left            =   2835
         TabIndex        =   26
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label LblSeleccionado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblSeleccionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1365
         TabIndex        =   25
         Top             =   5025
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Seleccionado :"
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   5025
         Width           =   1065
      End
      Begin VB.Label LblImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblImporte"
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
         Left            =   4995
         TabIndex        =   23
         Top             =   5040
         Width           =   1125
      End
      Begin VB.Label LblCobrado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCobrado"
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
         Left            =   6150
         TabIndex        =   22
         Top             =   5040
         Width           =   960
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblSaldo"
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
         Left            =   8355
         TabIndex        =   21
         Top             =   5040
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "Totales :"
         Height          =   240
         Left            =   4230
         TabIndex        =   20
         Top             =   5040
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "Caja :"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1455
         Width           =   540
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   6345
         TabIndex        =   18
         Top             =   270
         Width           =   630
      End
      Begin VB.Label LblNumero 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNumero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4215
         TabIndex        =   17
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Nº :"
         Height          =   225
         Left            =   3840
         TabIndex        =   16
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label3 
         Caption         =   "Comprobante :"
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Top             =   270
         Width           =   1155
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
         Height          =   315
         Left            =   1935
         TabIndex        =   14
         Top             =   750
         Width           =   5505
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor :"
         Height          =   240
         Left            =   150
         TabIndex        =   13
         Top             =   735
         Width           =   900
      End
      Begin VB.Label Label10 
         Caption         =   "Sdo a Cancelar :"
         Height          =   210
         Left            =   4560
         TabIndex        =   12
         Top             =   1455
         Width           =   1275
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7425
         Y1              =   1335
         Y2              =   1335
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   720
      Index           =   3
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5460
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Si el pago es solamente en Efectivo grabar sin valores."
      Height          =   420
      Left            =   90
      TabIndex        =   31
      Top             =   5520
      Width           =   2685
   End
End
Attribute VB_Name = "ComprasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim cRsl As ClsLectura
Dim sw As Boolean, cImporte As Currency
Dim rReci As ADODB.Recordset, rCaja As ADODB.Recordset
Dim rValo As ADODB.Recordset, rApli As ADODB.Recordset
Dim rBanco As ADODB.Recordset
Dim rAde As ADODB.Recordset

Public cSaldo As Currency
Public RsC As ADODB.Recordset
Public nCas As Long, sNumero As String
Public nAplicado As Double
Public nAplicadoGan As Double

Private Sub CmbComprobante_Click()
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler

  If sw = False Then
     Set cNum = New ClsComprobantesL
     nSuc = "0000"
     nNum = "00000000"
     LblNumero.Caption = nSuc & "-" & nNum
  End If
  Set cNum = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CmdAplicar_Click()
  Dim nMismo As Currency, cRve As ClsLectura
  
  Set cRve = New ClsLectura
  
  nMismo = TxtEfectivo.text
  
  If RsC.RecordCount <> 0 Then
     If cRve.TraerCantidad("CabRetencion", "Id", "Movimiento=" & RsC!Venta) = 0 Then
        If cRve.TraerValorDeUnCampo("Comprobantes", "Iva", "Descripcion='" & RsC!Comprobante & "'") = 1 Then
           nAplicado = nAplicado + cRve.TraerValorDeUnCampo("CabComprobantes", "Neto", "Movimiento=" & RsC!Venta)
        End If
     End If
     Dim nMoviCta As Long
     nMoviCta = cRve.TraerValorDeUnCampo("Aplicaciones", "MoviCta", "Venta=" & RsC!Venta)
     If cRve.TraerCantidad("TblGananciaAcum", "Id", "Movimiento=" & nMoviCta) = 0 Or cRve.TraerCantidad("TblGananciaAcum", "Importe", "Movimiento=" & nMoviCta) = 1 Then
        If cRve.TraerValorDeUnCampo("Comprobantes", "Iva", "Descripcion='" & RsC!Comprobante & "'") = 1 Then
           nAplicadoGan = nAplicadoGan + cRve.TraerValorDeUnCampo("CabComprobantes", "Neto", "Movimiento=" & RsC!Venta)
        End If
     End If
     
     If TxtCancelar.text > 0 Then
        TxtEfectivo.text = TxtEfectivo.text - TxtCancelar.text
     Else
       TxtEfectivo.text = nMismo
     End If
     LblCancelado.Caption = CDbl(LblCancelado.Caption) + CDbl(TxtCancelar.text)
     RsC!Cancelado = TxtCancelar.text
     RsC!Cobrar = RsC!importe - RsC!Cobrado - RsC!Cancelado
     RsC.Update
     TxtEfectivo.text = Format(TxtEfectivo.text, "#0.00")
     TxtCancelar.text = Format(0, "#0.00")
     LblCancelado.Caption = Format(LblCancelado, "#0.00")
     If TxtEfectivo.text <= 0 Or Option1.Value = False Then
       Botones True, True, True
    End If

  End If
  Set cRve = Nothing
End Sub

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              GrabarTodo
              CrearRs
              LlenarGrilla
              CmdBotones(3).Enabled = True
         Case 1
              Valor
         Case 2
              Salir
         Case 3
              Imprimir
  End Select
End Sub
Private Sub CmdClientes_Click()
  Dim RsC As ADODB.Recordset
  On Error GoTo errHandler
  
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 2
  FrmBuscarEx.Show 1
  CrearRs
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ProveedorTraerUno", nDat, True)
     TxtCliente.text = RsC!Proveedor
     LblCliente.Caption = RsC!RazonSocial
     Option1.Value = True
     LlenarGrilla
     nLlama = 1
     CmbComprobante_Click
     CmbCaja.SetFocus
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CmdDetalle_Click()
 ' ventaPagosDetalles.Show 1
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
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  If CmbComprobante.ListIndex <> -1 Then
     CmbComprobante.ListIndex = 0
  End If
  LblCliente.Caption = ""
'  If CmbCaja.ListIndex <> -1 Then
'     CmbCaja.ListIndex = 0
'  End If
  DtpFecha.Value = Date
  TxtEfectivo.text = "0.00"
  TxtCancelar.text = "0.00"
  LblNumero.Caption = "0000-00000000"
  LblImporte.Caption = "0.00"
  LblCobrado.Caption = "0.00"
  LblCancelado.Caption = "0.00"
  LblSaldo.Caption = "0.00"
  
  LblSeleccionado.Caption = "0.00"
End Sub
Private Sub HabilitarTodo(bEstado As Boolean, bGrid As Boolean)
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
         Or TypeOf ctl Is ComboBox Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  Grid1.Enabled = bGrid
End Sub
Private Sub Form_Load()

 CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
 CmdBotones(1).Picture = LoadResPicture("Valores", 0)
 CmdBotones(2).Picture = LoadResPicture("Salir", 0)
 CmdBotones(3).Picture = LoadResPicture("Imprimir", 0)

 Limpiar
 nLlama = 1
 Option1.Value = True
 HabilitarTodo True, True
 CmdDetalle.Enabled = False
 CargarCombos
 CrearRs
 Botones False, False, True
 Me.Top = 0
 Me.Left = 0
 
End Sub
Private Sub Botones(bGrabar As Boolean, bValor As Boolean, bSalir As Boolean)
  CmdBotones(3).Enabled = False
  CmdBotones(0).Enabled = bGrabar
  CmdBotones(1).Enabled = bValor
  CmdBotones(2).Enabled = bSalir
End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=9 and Suc=" & nSucursal
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  Set cRsl = Nothing
    
  CmbComprobante.ListIndex = 0
  Set cRsl = New ClsLectura
  CmbCaja.text = cRsl.DatoCombo("Cajas", "Caja", "Descripcion", cRsl.RegPorDefecto("Caja"))
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Grid1_Click()
  Dim i As Integer
  On Error GoTo errHandler
  
  LblSeleccionado.Caption = 0
  If TxtEfectivo.text <> 0 Then
     Botones False, False, True
     CmdAplicar.Enabled = False
      RsC.Bookmark = Grid1.Bookmark
      If RsC!Cuota <> 1000 Then
         CmdAplicar.Enabled = True
         LblSeleccionado.Caption = LblSeleccionado.Caption - RsC!Cobrar
         If Grid1.SelBookmarks.Count = 1 Then
            Botones True, True, True
         End If
         If CDbl(TxtEfectivo.text) > CDbl(RsC!Cobrar) Then
            TxtCancelar.text = Format(RsC!Cobrar, "#0.00")
         Else
            TxtCancelar.text = Format(TxtEfectivo.text, "#0.00")
         End If
      End If
  LblSeleccionado.Caption = Format(LblSeleccionado.Caption, "0.00")
   cImporte = CCur(TxtEfectivo.text)
   nCas = 1
 End If
Exit Sub

errHandler:
   ManejaErrores
End Sub
Private Sub TxtCancelar_GotFocus()
  TxtCancelar.SelStart = 0
  TxtCancelar.SelLength = Len(TxtCancelar.text)
End Sub

Private Sub TxtCancelar_LostFocus()
  If TxtCancelar.text = "" Then TxtCancelar.text = 0
  TxtCancelar.text = Format(TxtCancelar.text, "#0.00")
End Sub

Private Sub TxtCliente_GotFocus()
  TxtCliente.SelStart = 0
  TxtCliente.SelLength = Len(TxtCliente.text)
End Sub
Private Sub TxtCliente_LostFocus()
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set cRcL = New ClsClienteL
  LblCliente.Caption = ""
  CrearRs
  CmdBotones(3).Enabled = False
  If TxtCliente.text <> "" Then
     If cRcL.BuscarNombreProveedor(TxtCliente.text) <> "" Then
        LblCliente.Caption = cRcL.BuscarNombreProveedor(TxtCliente.text)
        Option1.Value = True
        LlenarGrilla
        nLlama = 1
        CmbComprobante_Click
        CmbCaja.SetFocus
     Else
        MsgBox "Proveedor no es el Correcto", vbCritical, "Atención"
        CmdClientes.SetFocus
     End If
  End If
  Set cRcL = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub CrearRs()
 On Error GoTo errHandler

 Set RsC = New ADODB.Recordset
  
 RsC.Fields.Append "Fecha", adDate
 RsC.Fields.Append "Comprobante", adVarChar, 30
 RsC.Fields.Append "Numero", adDouble
 RsC.Fields.Append "Cuota", adInteger
 RsC.Fields.Append "Importe", adCurrency
 RsC.Fields.Append "Cobrado", adCurrency
 RsC.Fields.Append "Cancelado", adCurrency
 RsC.Fields.Append "Cobrar", adCurrency
 RsC.Fields.Append "Venta", adDouble
 RsC.Fields.Append "Sucursal", adInteger
 RsC.CursorType = adOpenKeyset
 RsC.LockType = adLockOptimistic
 RsC.Open
 
 CabGrid
 
 Exit Sub
errHandler:
   ManejaErrores

End Sub
Private Sub CabGrid()
  Set Grid1.DataSource = RsC
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
 With Grid1
      .Columns(0).Width = 1400
      .Columns(0).Caption = "Vencimiento"
      .Columns(1).Width = 1600
      .Columns(1).Caption = "Comprobante"
      .Columns(2).Width = 950
      .Columns(2).Caption = "Numero"
      .Columns(2).NumberFormat = "000000"
      .Columns(2).Alignment = dbgRight
      .Columns(3).Width = 700
      .Columns(3).Caption = "Cuota"
      .Columns(3).NumberFormat = "00"
      .Columns(3).Alignment = dbgRight
      .Columns(4).Width = 950
      .Columns(4).Caption = "Importe"
      .Columns(4).NumberFormat = "#0.00"
      .Columns(4).Alignment = dbgRight
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Cobrado"
      .Columns(5).NumberFormat = "#0.00"
      .Columns(5).Alignment = dbgRight
      .Columns(6).Width = 1200
      .Columns(6).Caption = "Cancelado"
      .Columns(6).NumberFormat = "#0.00"
      .Columns(6).Alignment = dbgRight
      .Columns(7).Width = 1300
      .Columns(7).Caption = "A Cobrar"
      .Columns(7).NumberFormat = "#0.00"
      .Columns(7).Alignment = dbgRight
      .Columns(8).Visible = False
      .Columns(9).Visible = False
 End With
End Sub
Private Sub LlenarGrilla()
 Dim cCom As ClsComprobantesL, rCuen As ADODB.Recordset
 Dim sDebe As Currency, sCobrado As Currency, sCobrar As Currency
 Dim sHaber As Currency, hCobrado As Currency, hCobrar As Currency, nCotizacion As Currency
 Dim cRx As ClsLectura
 Dim cCobrado As Currency, cAux As Currency, cAdelanto As Currency, nReg As Long, j As Integer
 
 Set cCom = New ClsComprobantesL
 Set rCuen = New ADODB.Recordset
 Set cRx = New ClsLectura
 Set rCuen = cCom.DameCuentaCorrienteP(TxtCliente.text)
 CmdDetalle.Enabled = False
 LblSeleccionado.Caption = "0.00"
 
 nAplicado = 0
 nAplicadoGan = 0
 nReg = 1
 j = 1
 If rCuen.RecordCount <> 0 Then
    ProgressBar1.Max = rCuen.RecordCount
 End If
 Set db1 = New Connection
 db1.Open sDb
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = Left(cCom.DameCodComprobante(rCuen!Comprobante), 30)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    nCotizacion = cRx.TraerValorDeUnCampo("CabComprobantes", "Cotizacion", "Movimiento=" & rCuen!Venta)
    RsC!importe = Round(rCuen!Haber * nCotizacion, 2)
    cCobrado = cCom.DamePagos(rCuen!Cuota, rCuen!Venta)
    RsC!Cobrado = Round(cCobrado - cAdelanto, 2)
    RsC!Cancelado = 0
    RsC!Cobrar = Round((rCuen!Haber * nCotizacion), 2) - Round(RsC!Cobrado, 2)
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    RsC.Update
    If RsC!Cobrar <= 0 And nReg + 1 <= rCuen.RecordCount Then
       cAdelanto = Round(RsC!Cobrar, 2)
       nReg = nReg + 1
       If rCuen.RecordCount <> 1 Then
'          db.Open sDb
'          db.Execute "UPDATE CuentaCorriente SET Anulado=1 where Venta=" & rCuen!Venta & " AND Cuota=" & rCuen!Cuota
'          db.Close
         RsC.Delete
       End If
    Else
       sDebe = sDebe + Round(RsC!importe, 2)
       sCobrado = sCobrado + Round(RsC!Cobrado, 2)
       sCobrar = sCobrar + Round(RsC!Cobrar, 2)
       If RsC!Cobrar = 0 Then
'          db.Open sDb
'          db.Execute "UPDATE CuentaCorriente SET Anulado=1 where Venta=" & rCuen!Venta & " AND Cuota=" & rCuen!Cuota
'          db.Close
          RsC.Delete
       End If
       cAdelanto = 0
    End If
    ProgressBar1.Value = j
    j = j + 1
    rCuen.MoveNext
 Loop
 db1.Close
 ProgressBar1.Value = 0
 If RsC.RecordCount <> 0 Then
    CmdDetalle.Enabled = True
 End If
 Botones False, False, True

 LblImporte.Caption = Format(sDebe - sHaber, "#0.00")
 LblCobrado.Caption = Format(sCobrado - hCobrado, "#0.00")
 LblSaldo.Caption = Format(sCobrar - hCobrar, "#0.00")
End Sub
Private Sub TxtEfectivo_GotFocus()
  TxtEfectivo.SelStart = 0
  TxtEfectivo.SelLength = Len(TxtEfectivo.text)
End Sub
Private Sub TxtEfectivo_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtEfectivo) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtEfectivo_LostFocus()
  If TxtEfectivo.text = "" Then TxtEfectivo.text = 0
  TxtEfectivo.text = Format(TxtEfectivo.text, "#0.00")
End Sub

Private Sub GrabarTodo()
  On Error GoTo Errores
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cStock As ClsComprobantesL
  Dim RsCv As ADODB.Recordset, RsGan As ADODB.Recordset, RsGanAux As ADODB.Recordset
  
  Dim i As Byte, nNum As String * 8, nSuc As String * 4
  
  Dim Rx As ADODB.Recordset
  Dim Rx1 As ADODB.Recordset

  Set cRsl = New ClsLectura
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set rValo = New ADODB.Recordset
  Set rApli = New ADODB.Recordset
  Set rReci = New ADODB.Recordset
  Set rCaja = New ADODB.Recordset
  Set rBanco = New ADODB.Recordset
  Set RsGan = New ADODB.Recordset
  Set RsGanAux = New ADODB.Recordset
  
  cSaldo = 0

  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
  Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
  
  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
  Set rValo = cRsl.RsVacio("Valores", "Id", "N")
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
  Set rBanco = cRsl.RsVacio("CtaBanco", "Id", "N")
  Set rAde = cRsl.RsVacio("Aplicaciones", "Id", "N")
  
  Set RsGan = cRsl.RsVacio("TblGananciaAcum", "Id", "N")
  If Not ComprasValores.RsG Is Nothing Then
     Set RsGanAux = ComprasValores.RsG
  End If
  
  nNum = Format(cNum.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
  nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
  LblNumero.Caption = nSuc & "-" & nNum

  RsC.MoveFirst
  Do While Not RsC.EOF
     cSaldo = cSaldo + RsC!Cancelado
     RsC.MoveNext
  Loop
  
  RsC.MoveFirst
  
  RsCv.AddNew
  RsCv!Suc = nSucursal
'  RsCv!Id = 0
  RsCv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DtpFecha.Value
  RsCv!FechaIva = DtpFecha.Value
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!Sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = TxtCliente.text
  RsCv!Lista = 0
  RsCv!Vendedor = 0
  RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
  RsCv!CondVenta = 0
  RsCv!Neto = 0
  RsCv!Iva1 = 0
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  RsCv!NoGravados = 0
  Select Case cStock.TraerTipoMovimiento(RsCv!Comprobante)
         Case 15
              RsCv!Debe = cSaldo + TxtEfectivo.text
              RsCv!Haber = 0
         Case 10
              RsCv!Debe = 0
              RsCv!Haber = cSaldo + TxtEfectivo.text
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = ""
  RsCv!COTIZACION = 1
  RsCv!Anulado = 0
  RsCv.Update

     ' Grabo Recibo de Pago
  rReci.AddNew
 ' rReci!Id = 0
  rReci!fecha = DtpFecha.Value
  rReci!Cliente = TxtCliente.text
  rReci!Venta = RsCv!Movimiento
  rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  rReci!Numero = RsCv!Numero
  rReci!Sucursal = RsCv!Sucursal
  Select Case cStock.TraerTipoMovimiento(rReci!Comprobante)
         Case 10
              rReci!Cuota = 0
              rReci!Debe = 0
              rReci!Haber = cSaldo + TxtEfectivo.text
         Case 15
              rReci!Cuota = 1000
              rReci!Debe = cSaldo + TxtEfectivo.text
              rReci!Haber = 0
  End Select
  rReci!Anulado = 0
  rReci.Update
  
  RsC.MoveFirst
  
  If TxtEfectivo.text > 0 Then
     rApli.AddNew
     rApli!Venta = RsC!Venta
     rApli!MoviCta = RsCv!Movimiento
     rApli!Comprobante = rReci!Comprobante
     rApli!Numero = rReci!Numero
     rApli!Sucursal = rReci!Sucursal
     rApli!Cuota = RsC!Cuota
     rApli!importe = TxtEfectivo.text
     rApli!GAnancia = 0
     rApli!Anulado = 0
     rApli.Update
  End If
  
 ' Grabar Valores
 Dim rValoAux As ADODB.Recordset
 RsC.MoveFirst
 If nLlama = 4 Then
    Set rValoAux = ComprasValores.RsV
    rValoAux.MoveFirst
    Do While Not rValoAux.EOF
       rValo.AddNew
      ' rValo!Id = 0
       rValo!Venta = RsCv!Movimiento
       rValo!CodPago = rValoAux!CodPago
       rValo!formapago = rValoAux!formapago
       rValo!Nombre = rValoAux!Nombre
       rValo!Banco = rValoAux!Banco
       rValo!IdCheque = rValoAux!IdCheque
       rValo!NumeroCheque = rValoAux!NCheque
       rValo!FechaAcreditacion = rValoAux!FAcredi
       rValo!NumeroTarjeta = rValoAux!NTarjet
       rValo!FechaVencimiento = rValoAux!FVencim
       rValo!Autorizacion = rValoAux!NAutori
       rValo!Cupon = rValoAux!NCupon
       rValo!importe = Format(rValoAux!importe, "#0.00")
       rValo!PorVenta = " "
       rValo.Update
       rValoAux.MoveNext
   Loop
  Else
      rValo.AddNew
  '    rValo!Id = 0
      rValo!Venta = RsCv!Movimiento
      rValo!CodPago = 1
      rValo!formapago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
      rValo!Nombre = ""
      rValo!importe = cSaldo + CCur(TxtEfectivo.text)
      rValo.Update
  End If
  Dim cTotal As Currency, cDife As Currency
  cDife = CCur(TxtEfectivo.text)
  
  RsC.MoveFirst
  Do While Not RsC.EOF
     If RsC!Cancelado > 0 Then
        rApli.AddNew
        rApli!Venta = RsC!Venta
        rApli!MoviCta = RsCv!Movimiento
        rApli!Comprobante = cRsl.TraerValorDeUnCampo("Comprobantes", "Id", "Descripcion='" & RsC!Comprobante & "'")
        rApli!Numero = RsC!Numero
        rApli!Sucursal = RsC!Sucursal
        rApli!Cuota = RsC!Cuota
        rApli!importe = RsC!Cancelado
        rApli!GAnancia = 0
        rApli!Anulado = 0
        rApli.Update
     End If
     RsC.MoveNext
  Loop
    
' Actualizo Caja
  RsC.MoveFirst
  If cStock.ActualizaCaja(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     If rValo.EOF = False Or rValo.BOF = False Then
        rValo.MoveFirst
        Do While Not rValo.EOF
           If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(rValo!formapago) & "'") = "Caja" Then
              rCaja.AddNew
              rCaja!Suc = nSucursal
              rCaja!Movimiento = RsCv!Movimiento
              rCaja!Comprobante = rReci!Comprobante
              rCaja!Numero = rReci!Numero
              rCaja!Sucursal = rReci!Sucursal
              rCaja!fecha = DtpFecha.Value
              rCaja!Hora = Time
              rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
              rCaja!formapago = rValo!formapago
              rCaja!Nombre = LblCliente.Caption
              rCaja!Banco = rValo!Banco
              rCaja!NumeroCheque = rValo!NumeroCheque
              rCaja!FechaAcreditacion = rValo!FechaAcreditacion
              rCaja!NumeroTarjeta = rValo!NumeroTarjeta
              rCaja!FechaVencimiento = rValo!FechaVencimiento
              rCaja!Autorizacion = rValo!Autorizacion
              rCaja!Cupon = rValo!Cupon
              If cStock.SumaCaja(rCaja!Comprobante) Then
                 rCaja!Debe = rValo!importe
                 rCaja!Haber = 0
              Else
                 rCaja!Debe = 0
                 rCaja!Haber = rValo!importe
              End If
              rCaja!Abierta = "S"
              rCaja!Anulado = 0
              rCaja.Update
           Else
            If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(rValo!formapago) & "'") = "Banco" Then
               rBanco.AddNew
               rBanco!Movimiento = RsCv!Movimiento
               rBanco!Comprobante = rReci!Comprobante
               rBanco!Numero = rReci!Numero
               rBanco!Sucursal = rReci!Sucursal
               rBanco!fecha = DtpFecha.Value
               rBanco!Hora = Time
               rBanco!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
               rBanco!formapago = rValo!formapago
               rBanco!Nombre = LblCliente.Caption
               rBanco!Banco = rValo!Banco
               rBanco!NumeroCheque = rValo!NumeroCheque
               rBanco!FechaAcreditacion = rValo!FechaAcreditacion
               rBanco!NumeroTarjeta = rValo!NumeroTarjeta
               rBanco!FechaVencimiento = rValo!FechaVencimiento
               rBanco!Autorizacion = rValo!Autorizacion
               rBanco!Cupon = rValo!Cupon
               If cStock.SumaCaja(rBanco!Comprobante) Then
                  rBanco!Debe = rValo!importe
                  rBanco!Haber = 0
               Else
                  rBanco!Debe = 0
                  rBanco!Haber = rValo!importe
               End If
               rBanco!Abierta = "S"
               rBanco!Anulado = 0
               rBanco.Update
            End If
           End If
           rValo.MoveNext
       Loop
     End If
  End If
  
  If RsGanAux.State <> 0 Then
     If RsGanAux.RecordCount <> 0 Then
        RsGan.AddNew
        RsGan!Movimiento = RsCv!Movimiento
        RsGan!Proveedor = RsGanAux!Proveedor
        RsGan!fecha = RsGanAux!fecha
        RsGan!importe = RsGanAux!importe
        RsGan!Retenido = RsGanAux!Retenido
        RsGan!Anulado = False
        RsGan.Update
      End If
  End If
  cGrab.AgregarPagos RsCv, rReci, rValo, rApli, rCaja, rBanco, RsGan
  
  sNumero = LblNumero.Caption
  
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cStock = Nothing
  Set RsC = Nothing
  
Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "GrabarTodo"

End Sub

Private Sub Salir()
  If CmdBotones(0).Enabled = True Then
     If MsgBox("El Cobro no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
        Exit Sub
     End If
  End If
  Unload Me
End Sub
Private Sub Valor()
  nLlama = 4
  cSaldo = 0
  
  RsC.MoveFirst
  Do While Not RsC.EOF
     cSaldo = cSaldo + RsC!Cancelado
     RsC.MoveNext
  Loop

  ComprasValores.Show 1
End Sub

Private Sub Imprimir()
  Dim cCli As ClsClienteL, pY As Single, cRsl As ClsLectura
  Dim cNum As ClsComprobantesL
  
  
  Dim Rp As ADODB.Recordset
  
  Set cCli = New ClsClienteL
  Set cRsl = New ClsLectura
  Set cNum = New ClsComprobantesL
  
    Dim P As Printer
  
   For Each P In Printers
      If P.DeviceName = cRsl.TraerValorDeUnCampo("Impresion", "Impresora", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) Then
         Set Printer = P
      '   Printer.FontName = "Courier New"
         Printer.ScaleMode = 7
         Printer.FontBold = cRsl.TraerValorDeUnCampo("Impresion", "Negrita", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Printer.FontSize = cRsl.TraerValorDeUnCampo("Impresion", "Tamaño", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Exit For
      End If
   Next

  cSaldo = 0
  Printer.font.Name = "Ms Sans Seriff"
  Printer.font.Size = 14
  Printer.font.bold = True
  Printer.ScaleMode = 7
  Printer.CurrentY = 1.5
  Printer.CurrentX = 1
  Printer.Print cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
  Printer.font.Size = 14
  Printer.font.bold = False
  Printer.CurrentY = 2.5
  Printer.CurrentX = 1
  Printer.Print "ORDEN DE PAGO"
  Printer.font.Size = 10
  Printer.CurrentY = 3.5
  Printer.CurrentX = 13
  Printer.Print "Nro.de Documento :" & sNumero
  Printer.CurrentY = 4
  Printer.CurrentX = 13
  Printer.Print "Fecha de Pago :" & DtpFecha.Value
  Printer.CurrentY = 4.5
  Printer.CurrentX = 1
  Printer.Print "Proveedor :" & Trim(TxtCliente.text) & " - " & LblCliente.Caption
  Printer.CurrentY = 5
  Printer.CurrentX = 1
  Printer.Print "Domicilio :" & cRsl.TraerValorDeUnCampo("Proveedores", "Domicilio", "Proveedor='" & TxtCliente.text & "'")
  Printer.CurrentY = 5.5
  Printer.CurrentX = 1
  Printer.Print "Localidad :"
  Printer.CurrentY = 6
  Printer.CurrentX = 1
  Printer.Print "CUIT :" & cRsl.TraerValorDeUnCampo("Proveedores", "Cuit", "Proveedor='" & TxtCliente.text & "'")
  Printer.CurrentY = 6.5
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  Printer.CurrentY = 7
  Printer.CurrentX = 1
  Printer.Print "                           FECHA                      TIPO                        NUMERO              CUOTA               IMPORTE"
  Printer.CurrentY = 7.5
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  rApli.MoveFirst
  Printer.CurrentY = 8
  pY = 8
  rApli.MoveFirst
  Do While Not rApli.EOF
     Set Rp = cRsl.TraerComproPago(rApli!Venta)
'     Printer.CurrentX = 1
'     Printer.Print cNum.NumeroAsiento
     Printer.CurrentY = pY
     Printer.CurrentX = 4
     Printer.Print Rp!fecha
     Printer.CurrentY = pY
     Printer.CurrentX = 6.5
     Printer.Print Rp!Comprobante
     Printer.CurrentY = pY
     Printer.CurrentX = 11
     Printer.Print Format(Rp!Sucursal, "0000") & "-"
     Printer.CurrentY = pY
     Printer.CurrentX = 12
     Printer.Print Format(Rp!Numero, "00000000")
     Printer.CurrentY = pY
     Printer.CurrentX = 15
     Printer.Print rApli!Cuota
     Printer.CurrentY = pY
     Printer.CurrentX = 18 - Printer.TextWidth(Format(rApli!importe, "#0.00"))
     Printer.Print Format(rApli!importe, "#0.00")
     cSaldo = cSaldo + rApli!importe
     rApli.MoveNext
     pY = pY + 0.5
  Loop
  If rAde.RecordCount <> 0 Then
     rAde.MoveFirst
     Printer.CurrentX = 1
     Printer.Print Format(rReci!Asiento, "00000000")
     Printer.CurrentY = pY
     Printer.CurrentX = 4
     Printer.Print DtpFecha.Value
     Printer.CurrentY = pY
     Printer.CurrentX = 6.5
     Printer.Print "A Cuenta"
     Printer.CurrentY = pY
     Printer.CurrentX = 11
     Printer.Print Format(rAde!Sucursal, "0000") & "-"
     Printer.CurrentY = pY
     Printer.CurrentX = 12
     Printer.Print Format(rAde!Numero, "00000000")
     Printer.CurrentY = pY
     Printer.CurrentX = 15
     Printer.Print ""
     Printer.CurrentY = pY
     Printer.CurrentX = 18 - Printer.TextWidth(Format(rApli!importe, "#0.00"))
     Printer.Print Format(rAde!importe, "#0.00")
     cSaldo = cSaldo + rAde!importe
  End If
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 8
  Printer.Print "T O T A L   A  P A G A R :"
  Printer.CurrentY = pY
  Printer.CurrentX = 18 - Printer.TextWidth(Format(cSaldo, "#0.00"))
  Printer.Print Format(cSaldo, "#0.00")

  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "CANCELACION :"
  rValo.MoveFirst
  pY = pY + 1
  Do While Not rValo.EOF
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     Printer.Print rValo!formapago
     Printer.CurrentY = pY
     Printer.CurrentX = 7
     If rValo!NumeroCheque <> "" Then
        Printer.Print cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & rValo!Banco) & " " & rValo!NumeroCheque & " " & rValo!FechaAcreditacion
     Else
        If rValo!Cupon <> "" Then
           Printer.Print cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & rValo!Banco) & " " & rValo!NumeroTarjeta & " " & rValo!Cupon
        End If
     End If
     Printer.CurrentY = pY
     Printer.CurrentX = 18 - Printer.TextWidth(Format(rValo!importe, "#0.00"))
     Printer.Print Format(rValo!importe, "#0.00")
     rValo.MoveNext
     pY = pY + 0.5
  Loop
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 8
  Printer.Print "T O T A L   D E L   P A G O :"
  Printer.CurrentY = pY
  Printer.CurrentX = 18 - Printer.TextWidth(Format(cSaldo, "#0.00"))
  Printer.Print Format(cSaldo, "#0.00")
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "CONFECCIONO :"
  Printer.CurrentY = pY
  Printer.CurrentX = 10
  Printer.Print "AUTORIZO :_____________________________________________"
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "Recibimos de " & cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " la cantidad de $ " & Format(cSaldo, "#0.00")
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "En todo de acuerdo con la presente liquidación."
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "ACLARACION_______________________________  DOC:_____________________ FIRMA_________________________"
  Printer.EndDoc

End Sub



'Private Sub GrabarTodo()
'  On Error GoTo Errores
'  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cStock As ClsComprobantesL
'  Dim rValo As ADODB.Recordset, rApli As ADODB.Recordset
'  Dim rReci As ADODB.Recordset, rCaja As ADODB.Recordset
'  Dim i As Byte
'  Dim RsCv As ADODB.Recordset
'
'  Dim Rx As ADODB.Recordset
'  Dim Rx1 As ADODB.Recordset
'  Dim sCuen As String
'
'
'  Set cRsl = New ClsLectura
'  Set cGrab = New ClsComprobantesE
'  Set cNum = New ClsComprobantesL
'  Set cStock = New ClsComprobantesL
'  Set rValo = New ADODB.Recordset
'  Set rApli = New ADODB.Recordset
'  Set rReci = New ADODB.Recordset
'  Set rCaja = New ADODB.Recordset
'
'
'  sCompr = ""
'  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
'  Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
'
'  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
'  Set rValo = cRsl.RsVacio("Valores", "Id", "N")
'  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
'  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
'
'  Set Rx = cRsl.RsVacio("AsientosCabeza", "Id", "N")
'  Set Rx1 = cRsl.RsVacio("AsientosDetalles", "Id", "N")
'
'  RsCv.AddNew
'  RsCv!ID = 0
'  RsCv!Tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
'  RsCv!Movimiento = cNum.TraerUltimoMovimiento
'  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
'  RsCv!Fecha = DTPFecha.Value
'  RsCv!Numero = Right(LblNumero.Caption, 8)
'  RsCv!Sucursal = Left(LblNumero.Caption, 4)
'  RsCv!Cliente = TxtCliente.Text
'  RsCv!Lista = 0
'  RsCv!Vendedor = 0
'  RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
'  RsCv!CondVenta = 0
'  RsCv!Neto = 0
'  RsCv!Iva1 = 0
'  RsCv!Descuentos = 0
'  RsCv!Financiacion = 0
'  RsCv!NoGravados = 0
'  Select Case cStock.TraerTipoMovimiento(RsCv!Comprobante)
'         Case 15
'              RsCv!Debe = CCur(TxtEfectivo.Text)
'              RsCv!Haber = 0
'         Case 10
'              RsCv!Debe = 0
'              RsCv!Haber = CCur(TxtEfectivo.Text)
'  End Select
'  RsCv!Desde = 0
'  RsCv!Hasta = 0
'  RsCv!factura = 0
'  RsCv!Motivo = ""
'  RsCv.Update
'
'
'     ' Grabo Recibo de Pago
'  rReci.AddNew
'  rReci!ID = 0
'  rReci!Fecha = DTPFecha.Value
'  rReci!Cliente = TxtCliente.Text
'  rReci!Venta = RsC!Venta
'  rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
'  rReci!Numero = cNum.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex))
'  rReci!Sucursal = cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex))
'  Select Case cStock.TraerTipoMovimiento(rReci!Comprobante)
'         Case 15
'              rReci!Cuota = 0
'              rReci!Debe = CCur(TxtEfectivo.Text)
'              rReci!Haber = 0
'         Case 10
'              rReci!Cuota = 0
'              rReci!Debe = 0
'              rReci!Haber = CCur(TxtEfectivo.Text)
'  End Select
'  rReci.Update
'
' ' Grabar Valores
'
' Dim rValoAux As ADODB.Recordset, sTi As String, sTin As Byte
' If nLlama = 4 Then
'    Set rValoAux = ComprasValores.RsV
'    rValoAux.MoveFirst
'    Do While Not rValoAux.EOF
'       rValo.AddNew
'       rValo!ID = 0
'       rValo!Venta = RsC!Venta
'       rValo!CodPago = rValoAux!CodPago
'       rValo!formapago = rValoAux!formapago
'       sTi = cRsl.TraerValorDeUnCampo("Cheques", "Titular", "Id=" & rValoAux!IdCheque)
'       sTin = InStr(1, sTi, "-")
'       If sTin <> 0 Then
'          sTi = Mid(sTi, 1, sTin - 1)
'       End If
'       If sTin = 0 Then
'          rValo!Nombre = rValoAux!Nombre
'       Else
'          rValo!Nombre = sTi
'       End If
'       rValo!IdCheque = rValoAux!IdCheque
'       rValo!Banco = rValoAux!Banco
'       rValo!NumeroCheque = rValoAux!NCheque
'       If rValoAux!IdCheque <> 0 Then
'          rValo!FechaAcreditacion = cRsl.TraerValorDeUnCampo("Cheques", "FechaAcre", "Id=" & rValoAux!IdCheque)
'       Else
'         rValo!FechaAcreditacion = rValoAux!FAcredi
'       End If
'       rValo!NumeroTarjeta = rValoAux!NTarjet
'       rValo!FechaVencimiento = rValoAux!FVencim
'       rValo!Autorizacion = rValoAux!NAutori
'       rValo!Cupon = rValoAux!NCupon
'       rValo!importe = Format(rValoAux!importe, "#0.00")
'       rValo!PorVenta = " "
'       rValo.Update
'       rValoAux.MoveNext
'   Loop
'  Else
'      rValo.AddNew
'      rValo!ID = 0
'      rValo!Venta = RsC!Venta
'      rValo!CodPago = 1
'      rValo!formapago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
'      rValo!Nombre = LblCliente.Caption
'      rValo!importe = CCur(TxtEfectivo.Text)
'      rValo.Update
'  End If
'  Dim cTotal As Currency, cDife As Currency
'  cDife = CCur(TxtEfectivo.Text)
'  For i = 0 To Grid1.SelBookmarks.Count - 1
'     RsC.Bookmark = Grid1.SelBookmarks(i)
'     If RsC!Cuota <> 1000 Then
'        If RsC!Cobrar < 0 Then
'           cTotal = Abs(RsC!Cobrar) + cDife
'        Else
'           cTotal = cDife - RsC!Cobrar
'        End If
'     Else
'        cTotal = cDife + RsC!Cobrar
'     End If
'     sCompr = sCompr & Format(RsC!Sucursal, "0000") & "-" & Format(RsC!Numero, "00000000") & "  "
'     rApli.AddNew
'     rApli!Venta = RsC!Venta
'     rApli!Comprobante = rReci!Comprobante
'     rApli!Numero = rReci!Numero
'     rApli!Sucursal = rReci!Sucursal
'     rApli!Cuota = RsC!Cuota
'     If RsC!Cuota <> 1000 Then
'        If cTotal >= 0 Then
'           rApli!importe = RsC!Cobrar
'        Else
'           rApli!importe = Abs(cDife)
'        End If
'     Else
'        If TxtEfectivo.Text >= 0 Then
'           rApli!importe = RsC!Cobrar
'        Else
'         rApli!importe = Abs(cTotal)
'        End If
'     End If
'     cDife = cTotal
'     rApli.Update
'  Next i
'
'  If cTotal > 0 Then
'     rApli.AddNew
'     rApli!Venta = RsC!Venta
'     rApli!Comprobante = rReci!Comprobante
'     rApli!Numero = rReci!Numero
'     rApli!Sucursal = rReci!Sucursal
'     rApli!Cuota = RsC!Cuota
'     rApli!importe = cTotal
'     rApli.Update
'  End If
'  CrearRsPrint
'' Actualizo Caja
'  If cStock.ActualizaCaja(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
'     If rValo.EOF = False Or rValo.BOF = False Then
'        rValo.MoveFirst
'        Do While Not rValo.EOF
'           rCaja.AddNew
'           rCaja!Movimiento = RsC!Venta
'           rCaja!Comprobante = rReci!Comprobante
'           rCaja!Numero = rReci!Numero
'           rCaja!Sucursal = rReci!Sucursal
'           rCaja!Fecha = DTPFecha.Value
'           rCaja!Hora = Time
'           rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
'           rCaja!formapago = rValo!formapago
'           rCaja!Nombre = rValo!Nombre
'           rCaja!Banco = rValo!Banco
'           rCaja!NumeroCheque = rValo!NumeroCheque
'           rCaja!FechaAcreditacion = rValo!FechaAcreditacion
'           rCaja!NumeroTarjeta = rValo!NumeroTarjeta
'           rCaja!FechaVencimiento = rValo!FechaVencimiento
'           rCaja!Autorizacion = rValo!Autorizacion
'           rCaja!Cupon = rValo!Cupon
'           If cStock.SumaCaja(rCaja!Comprobante) Then
'              rCaja!Debe = rValo!importe
'              rCaja!Haber = 0
'           Else
'              rCaja!Debe = 0
'              rCaja!Haber = rValo!importe
'           End If
'           rCaja!Abierta = "S"
'           rCaja.Update
'
'           rPrint2.AddNew
'           rPrint2!Cliente = TxtCliente.Text
'           rPrint2!Razon = LblCliente.Caption
'           rPrint2!Direccion = cRsl.TraerValorDeUnCampo("Proveedores", "Domicilio", "Proveedor='" & TxtCliente.Text & "'")
'           rPrint2!Localidad = cRsl.TraerValorDeUnCampo("Proveedores", "CodigoPostal", "Proveedor='" & TxtCliente.Text & "'") & " - " & cRsl.TraerValorDeUnCampo("Localidades", "Descripcion", "Localidad=" & cRsl.TraerValorDeUnCampo("Proveedores", "Localidad", "Proveedor='" & TxtCliente.Text & "'"))
'           rPrint2!APagar = TxtEfectivo.Text
'           rPrint2!Tipo = rValo!formapago
'           If rValo!Banco <> "" Then
'              rPrint2!Banco = "" & cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & rValo!Banco)
'           End If
'           rPrint2!Detalle = "" & rValo!Nombre
'           rPrint2!Cheque = "" & rValo!NumeroCheque
'           rPrint2!Fecha = rValo!FechaAcreditacion
'           rPrint2!ImporteValor = rValo!importe
'           rPrint2.Update
'           rValo.MoveNext
'       Loop
'     End If
'  End If
'
'
'  Rx.AddNew
'  Rx!ID = 0
'  Rx!Movimiento = cNum.TraerUltimoMovimiento
'  Rx!Empresa = 11
'  Rx!Fecha = DTPFecha.Value
'  Rx!Numero = cNum.NumeroAsiento
'  Rx!Leyenda = CmbComprobante.Text & " " & Format(rReci!Sucursal, "0000") & "-" & Format(rReci!Numero, "00000000")
'  Rx!Anulado = 0
'  Rx.Update
'
'  rValo.MoveFirst
'
'  sCuen = ""
'  Rx1.AddNew
'  Rx1!ID = 0
'  Rx1!Movimiento = cNum.TraerUltimoMovimiento
'  Rx1!Cuenta = cRsl.TraerValorDeUnCampo("Comprobantes", "Cuenta", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
'  Rx1!Nombre = cRsl.TraerValorDeUnCampo("PlanDeCuentas", "Descripcion", "Codigo='" & Rx1!Cuenta & "'")
'  Rx1!Observacion = LblCliente.Caption
'  Rx1!Debe = TxtEfectivo.Text
'  Rx1!Haber = 0
'  Rx1.Update
'
'  Do While Not rValo.EOF
'        Rx1.AddNew
'        Rx1!ID = 0
'        Rx1!Movimiento = cNum.TraerUltimoMovimiento
'        Rx1!Cuenta = cRsl.TraerValorDeUnCampo("CondVenta", "Cuenta", "CondVta=" & rValo!CodPago)
'        Rx1!Nombre = cRsl.TraerValorDeUnCampo("PlanDeCuentas", "Descripcion", "Codigo='" & Rx1!Cuenta & "'")
'        Rx1!Observacion = LblCliente.Caption
'        Rx1!Debe = 0
'        Rx1!Haber = rValo!importe
'        Rx1.Update
'     rValo.MoveNext
'  Loop
'
''
'  CmdBotones(0).Enabled = False
'  CmdBotones(1).Enabled = False
'' Imprimir Recibo de Cobro
'
'
'
'
'  cGrab.AgregarPagos RsCv, rReci, rValo, rApli, rCaja, Rx, Rx1
'
'  Set cGrab = Nothing
'  Set cNum = Nothing
'  Set cStock = Nothing
'  Set RsC = Nothing
'Exit Sub
'Errores:
' MsgBox Err.Description, vbCritical, "Atención"
'
'End Sub
'
