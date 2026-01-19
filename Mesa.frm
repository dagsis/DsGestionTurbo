VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mesas"
   ClientHeight    =   8385
   ClientLeft      =   5040
   ClientTop       =   1545
   ClientWidth     =   6870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4620
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   480
      Width           =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar Venta"
      Height          =   660
      Index           =   4
      Left            =   4005
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7620
      Width           =   1830
   End
   Begin VB.Frame Frame3 
      Caption         =   "Su pago"
      Height          =   4680
      Left            =   135
      TabIndex        =   25
      Top             =   1380
      Width           =   6690
      Begin VB.ComboBox CmbForma 
         Height          =   315
         Left            =   3465
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2550
         Width           =   3015
      End
      Begin VB.ComboBox CmbCompro 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1665
         Width           =   3015
      End
      Begin VB.TextBox TxtPago 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   5265
         TabIndex        =   29
         Text            =   "TxtPago"
         Top             =   1185
         Width           =   1215
      End
      Begin VB.CommandButton CmdVuelto 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5010
         TabIndex        =   33
         Top             =   3465
         Width           =   1470
      End
      Begin VB.ComboBox CmbCond 
         Height          =   315
         Left            =   3465
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2100
         Width           =   3015
      End
      Begin VB.CommandButton BtnBuscar 
         Caption         =   "..."
         Height          =   240
         Left            =   6210
         TabIndex        =   27
         Top             =   285
         Width           =   315
      End
      Begin VB.TextBox TxtCliente 
         Height          =   300
         Left            =   1380
         MaxLength       =   15
         TabIndex        =   26
         Text            =   "TxtCliente"
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label24 
         Caption         =   "C.Pago :"
         Height          =   255
         Left            =   630
         TabIndex        =   48
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "F12"
         Height          =   270
         Left            =   135
         TabIndex        =   47
         Top             =   2535
         Width           =   390
      End
      Begin VB.Label Label22 
         Caption         =   "F6"
         Height          =   270
         Left            =   135
         TabIndex        =   46
         Top             =   1650
         Width           =   345
      End
      Begin VB.Label Label21 
         Caption         =   "F7"
         Height          =   270
         Left            =   135
         TabIndex        =   45
         Top             =   1185
         Width           =   345
      End
      Begin VB.Label Label20 
         Caption         =   "F11"
         Height          =   270
         Left            =   135
         TabIndex        =   44
         Top             =   2115
         Width           =   390
      End
      Begin VB.Label Label17 
         Caption         =   "Compr. :"
         Height          =   255
         Left            =   690
         TabIndex        =   43
         Top             =   1650
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Total de su Mesa :"
         Height          =   255
         Left            =   675
         TabIndex        =   42
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label15 
         Caption         =   "Su Pago :"
         Height          =   255
         Left            =   675
         TabIndex        =   41
         Top             =   1185
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Su Vuelto :"
         Height          =   255
         Left            =   3570
         TabIndex        =   40
         Top             =   3015
         Width           =   1575
      End
      Begin VB.Label LblCompra 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCompra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5265
         TabIndex        =   39
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblVuelto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblVuelto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         TabIndex        =   38
         Top             =   3015
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   630
         TabIndex        =   37
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblImpre 
         Height          =   255
         Left            =   75
         TabIndex        =   36
         Top             =   3045
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "F.Pago :"
         Height          =   255
         Left            =   645
         TabIndex        =   35
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "F5"
         Height          =   270
         Left            =   135
         TabIndex        =   34
         Top             =   315
         Width           =   345
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
         Height          =   330
         Left            =   2190
         TabIndex        =   32
         Top             =   270
         Width           =   3960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Presu."
      Height          =   660
      Index           =   6
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7620
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comanda"
      Enabled         =   0   'False
      Height          =   660
      Index           =   5
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7620
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ingreso de Detalles"
      Height          =   1260
      Left            =   180
      TabIndex        =   16
      Top             =   6225
      Width           =   6720
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   1545
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   660
         Width           =   510
      End
      Begin VB.TextBox Text3 
         Height          =   540
         Left            =   3285
         MaxLength       =   80
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "Mesa.frx":0000
         Top             =   645
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   840
         MaxLength       =   25
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Observación :"
         Height          =   195
         Left            =   2190
         TabIndex        =   20
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad :"
         Height          =   210
         Left            =   135
         TabIndex        =   19
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   300
         Left            =   2145
         TabIndex        =   18
         Top             =   270
         Width           =   3390
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo :"
         Height          =   225
         Left            =   150
         TabIndex        =   17
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   660
      Index           =   3
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7620
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Anular"
      Height          =   660
      Index           =   1
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7620
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nuevo"
      Height          =   660
      Index           =   0
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7620
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle de la Mesa"
      Height          =   4800
      Left            =   180
      TabIndex        =   9
      Top             =   1350
      Width           =   6690
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   4470
         Left            =   120
         TabIndex        =   22
         Top             =   225
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   7885
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
               LCID            =   11274
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
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4620
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   60
      Width           =   2160
   End
   Begin VB.PictureBox HASAR1 
      Height          =   480
      Left            =   4050
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   24
      Top             =   5325
      Width           =   1200
   End
   Begin VB.Label Label19 
      Caption         =   "Descuento :"
      Height          =   255
      Left            =   3540
      TabIndex        =   53
      Top             =   480
      Width           =   960
   End
   Begin VB.Label LblTotalNoGrabado 
      Caption         =   "Label19"
      Height          =   255
      Left            =   3570
      TabIndex        =   52
      Top             =   7935
      Width           =   1080
   End
   Begin VB.Label LblTotalIva 
      Caption         =   "Label19"
      Height          =   195
      Left            =   1950
      TabIndex        =   51
      Top             =   7950
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label LblTotalGrabado 
      Caption         =   "Label19"
      Height          =   255
      Left            =   465
      TabIndex        =   50
      Top             =   7950
      Width           =   1065
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   15
      Top             =   495
      Width           =   1005
   End
   Begin VB.Label Label9 
      Caption         =   "Hora Apertura :"
      Height          =   210
      Left            =   165
      TabIndex        =   14
      Top             =   450
      Width           =   1290
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5415
      TabIndex        =   13
      Top             =   1020
      Width           =   1365
   End
   Begin VB.Label Label7 
      Caption         =   "Sub Total :"
      Height          =   240
      Left            =   3960
      TabIndex        =   12
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Mozo/Moto :"
      Height          =   255
      Left            =   3540
      TabIndex        =   6
      Top             =   75
      Width           =   930
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1515
      TabIndex        =   5
      Top             =   60
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Mesa/Delivery :"
      Height          =   225
      Left            =   165
      TabIndex        =   4
      Top             =   60
      Width           =   1140
   End
End
Attribute VB_Name = "frmMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bVerificar As Boolean

Dim sw As Boolean, sw1 As Boolean, rAux As Recordset, nListaVenta As Integer
Dim Hora As String, nReg As Long, nCant As Integer, bBor As Boolean, bCompro As Boolean

Private KeyRetroceso As Boolean
Dim sConsFinal As String, sNomFinal As String, nMoviAfip As Long, sCondPago As String, dTotalAux As Double
Dim nMozo As Integer, bFoco As Boolean
Dim dNombre As String, dDireccion As String, dObservacion() As String, dHora As String

Private Sub BtnBuscar_Click()
 Dim RsC As ADODB.Recordset, cRsl As ClsLectura
  On Error GoTo errHandler
  
  Dim dTotal As Double
  
  dTotal = dTotalAux
  
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     TxtCliente.text = RsC!Cliente
     LblCliente.Caption = RsC!RazonSocial
     CmbCompro.text = cRsl.TraerComproCliente(TxtCliente.text, Frame3.Caption)
     If TxtCliente.text = "." Then
        CmbCond.text = "Contado"
     Else
        CmbCond.text = sCondPago
     End If
     
     dTotal = dTotalAux - CalcularDescuento(dTotal)
     Label8.Caption = Format(dTotal, "#0.00")
     LblCompra.Caption = Label8.Caption
     TxtPago.SetFocus
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
End Sub


Private Sub CmbCond_Click()
  If VerificarCuenta(TxtCliente.text) = False Then
       bVerificar = False
       Frame3.Visible = False
     '  HabilitarDetalles bCod, False
    '   CmdDetalle(0).Enabled = bBoto
       TxtCliente.text = sConsFinal
       LblCliente.Caption = sNomFinal
       CmbCond.text = sCondPago
       CmbForma.text = "EFECTIVO"
       Label8.Caption = Format(dTotalAux, "#0.00")
       Grid1.Enabled = True
    Else
       bVerificar = True
       Dim dTotal As Double, cRsl As ClsLectura
  
       dTotal = dTotalAux
  
       Set cRsl = New ClsLectura
       
       If TxtCliente <> "." Then
          If CmbCond.text <> "Contado" Then
             Recalcular cRsl.TraerRecargoCliente(TxtCliente.text)
           Else
              Recalcular 0
           End If
           If CmbCond.text = "Contado" Then
              Recalcular -cRsl.TraerDescuentoCliente(TxtCliente.text)
           End If
       Else
          Recalcular 0
       End If

       dTotal = dTotalAux - CalcularDescuento(dTotal)
       Label8.Caption = Format(dTotal, "#0.00")
       LblCompra.Caption = Label8.Caption
       TxtPago_LostFocus
    End If
End Sub

Private Sub CmdVuelto_Click(Index As Integer)
 
 On Error GoTo Errores


  Dim cRsl As ClsLectura, cRle As ClsEscritura, cRv As ADODB.Recordset, nMovimiento As Long
  Dim nInter As Single, sInter As String, nTasa As Single
  Dim cCuit As String, cDire As String, nComp As Long
  Dim nNumero As Long, bCierreControlador As Boolean, bGraboTick As Boolean
  Dim nTotal, nGrabado, nIvas, nNoGrabado As Currency
  Dim cNum As ClsComprobantesL
  Dim nAlicuota As Single, nPercepcion As Currency, sPercep As String
  Dim cPe As ClsComprobantesE
  Dim pInterPorc As Double, vInterno As Double

  
  Set cNum = New ClsComprobantesL
  Set cPe = New ClsComprobantesE
  
  nAlicuota = 0
  nPercepcion = 0
  
  bCierreControlador = False
  bGraboTick = False
  
  Botones False, False, False, False, False, False
  
  Set cRsl = New ClsLectura
  nAlicuota = cRsl.TraerValorDeUnCampo("Clientes", "Percepcion", "Cliente='" & TxtCliente.text & "'")
  
  nPercepcion = CCur(LblTotalNoGrabado.Caption) * nAlicuota / 100
     
  Dim FS As String
  Dim sComando As String

  Dim pParam(1) As String, pRs(0) As ADODB.Recordset
  
  CmdVuelto(0).Enabled = False

  Set cRsl = New ClsLectura
  Set cRle = New ClsEscritura
  
  Set cRv = cRsl.RsVacio("Ventas", "Id", "N")
  
  FS = Chr$(28)
      
  BtnBuscar.Enabled = False
  TxtCliente.Enabled = False
  CmdVuelto(0).Enabled = False
  CmbCond.Enabled = False
     
  nMovimiento = cRsl.ActualizarMovimientoTraer
  nMovimiento = nMovimiento + (cRsl.RegPorDefecto("Caja") * 10000000)
  nMoviAfip = nMovimiento
  
  If rAux.RecordCount <> 0 Then
     rAux.MoveFirst
     Do While Not rAux.EOF
        cRv.AddNew
        cRv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
        cRv!Movimiento = nMovimiento
        cRv!fecha = Date
        cRv!Cliente = IIf(TxtCliente.text = "", ".", TxtCliente.text)
        cRv!Lista = nListaVenta
        cRv!Vendedor = Combo1.ItemData(Combo1.ListIndex)
        cRv!Caja = cRsl.RegPorDefecto("Caja")
        cRv!formapago = CmbForma.ItemData(CmbForma.ListIndex)
        cRv!CondVenta = CmbCond.ItemData(CmbCond.ListIndex)
        If Frame3.Caption = "N" Then
           nNumero = cRsl.TraerValorDeUnCampo("Comprobantes", "Numero", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
           bCompro = True
           cRv!Comprobante = CmbCompro.ItemData(CmbCompro.ListIndex)
           cRv!Numero = nNumero
           cRv!Sucursal = cRsl.TraerValorDeUnCampo("Comprobantes", "Sucursal", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
           cRv!TotalGrabado = CCur(Label8.Caption)
           cRv!TotalNoGrabado = 0
           cRv!TotalIva = 0
           cRv!Grabado = 0
           cRv!Impuestos = 0
           cRv!NoGrabado = 0
           cRv!Percepcion = 0
           cRv!TotalGeneral = CCur(Label8.Caption)
           nComp = CmbCompro.ItemData(CmbCompro.ListIndex)
        Else
           bCompro = False
           cRv!Comprobante = CmbCompro.ItemData(CmbCompro.ListIndex)
           cRv!Numero = 0
           cRv!Sucursal = cRsl.TraerValorDeUnCampo("Comprobantes", "Sucursal", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
           cRv!TotalGrabado = CCur(LblTotalGrabado.Caption)
           cRv!TotalNoGrabado = CCur(LblTotalNoGrabado.Caption)
           cRv!TotalIva = CCur(LblTotalIva.Caption)
           cRv!TotalGeneral = CCur(Label8.Caption)
           sInter = cRsl.TraerValorDeUnCampo("Productos", "TipoInterno", "Producto='" & rAux!a_codart & "'")
           Dim sInter1 As String
           nTasa = cRsl.TraerTasaProducto(rAux!a_codart)
           pInterPorc = 0: vInterno = 0
           Select Case sInter
                  Case "No"
                        nInter = 0
                        sInter = "0"
                  Case "Porcentaje"
                        sInter = ""
                        sInter = "0"
                        pInterPorc = (cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & rAux!a_codart & "'") / 100) + 1
                        vInterno = Int(((rAux!a_precio) / pInterPorc) * 100)
                        nInter = vInterno / 100
                  Case "Fijo"
                       nInter = cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & rAux!a_codart & "'")
                       sInter = Format(Int(nInter), "0000000")
                       sInter1 = sInter & Replace(nInter, "0.", "") & Mid("00000000", 1, Len(Replace(nInter, "0.", "")))
           End Select
           rAux!a_total = rAux!a_total - CalcularDescuento(rAux!a_total)
           cRv!Impuestos = (rAux!a_total - (rAux!a_cant * nInter)) - (rAux!a_total - (rAux!a_cant * nInter)) / ((nTasa / 100) + 1)
           cRv!Grabado = (rAux!a_precio - nInter) / ((nTasa / 100) + 1) * rAux!a_cant
           cRv!Impuestos = cRv!Grabado * nTasa / 100
           cRv!NoGrabado = nNoGrabado + (rAux!a_cant * nInter)
           nComp = CmbCompro.ItemData(CmbCompro.ListIndex)
        End If
        cRv!Producto = rAux!a_codart
        cRv!Descripcion = rAux!A_DESCRI
        cRv!Cantidad = rAux!a_cant
        cRv!PUnitario = rAux!a_precio - CalcularDescuento(rAux!a_precio)
        cRv!Percepcion = nPercepcion
        cRv!pTotal = rAux!a_total + nPercepcion
        cRv!Anulado = 0
        cRv.Update
        rAux.MoveNext
     Loop
     
     Dim sCondiIva As String, nFaNo As String, pCuit As String
     
     If Frame3.Caption = "P" Then
         Dim Respuesta As Boolean
         If nPuerto <> 0 Then
            cCuit = cRsl.TraerValorDeUnCampo("Clientes", "Cuit", "Cliente='" & TxtCliente.text & "'")
            cDire = cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & TxtCliente.text & "'")
            If cCuit = "" Then cCuit = "00000000000"
            If cDire = "" Then cDire = "X"
            If nControlador = 2 Then
               ' Aca Epson
               Principal.PrinterFiscal1.PortNumber = nPuerto
               Principal.PrinterFiscal1.MessagesOn = False
               
               If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 0 Then
                  Select Case cRsl.TraerValorDeUnCampo("Clientes", "TipoIva", "Cliente='" & TxtCliente.text & "'")
                         Case 1
                              sCondiIva = "F"
                         Case 3
                              sCondiIva = "N"
                         Case 4
                              sCondiIva = "M"
                         Case 5
                             sCondiIva = "E"
                  End Select
                  
                  If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 5 Then
                     If cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 83 Then
                        Respuesta = Principal.PrinterFiscal1.OpenTicket("G")
                        If Respuesta = False Then
                         '  MsgBox "Problemas Para Generar El Ticket Factura.." & Chr(13) & Chr(10) & "Verifique Cuit, Condifion de Iva o Cliente", vbCritical, "Atención"
                           GoTo Errores
                        End If
                     Else
                         pCuit = "CUIT"
                         If sCondiIva = "F" Then
                            pCuit = "NDOC"
                            cCuit = "00000000"
                         End If
                        Respuesta = Principal.PrinterFiscal1.OpenInvoice("T", "C", "B", "1", "P", "12", "I", sCondiIva, LblCliente.Caption, "X", pCuit, Trim(cCuit), "N", cDire, "-", "-", "-", CmbCond.text, "G")
                        If Respuesta = False Then
                           MsgBox "Problemas Para Generar El Ticket Factura.." & Chr(13) & Chr(10) & "Verifique Cuit, Condifion de Iva o Cliente", vbCritical, "Atención"
                           GoTo Errores
                        End If
                     End If
                  Else
                      pCuit = "CUIT"
                      If sCondiIva = "F" Then
                         pCuit = "NDOC"
                         cCuit = "00000000"
                      End If
                  ' Nota de Credito B
                     Respuesta = Principal.PrinterFiscal1.OpenInvoice("M", "C", "B", "1", "P", "12", "I", sCondiIva, LblCliente.Caption, "X", pCuit, Trim(cCuit), "N", cDire, "-", "-", Str(nFactura), "XX", "G")
                     If Respuesta = False Then
                        MsgBox "Problemas Para Generar El Ticket Factura.." & Chr(13) & Chr(10) & "Verifique Cuit, Condifion de Iva o Cliente", vbCritical, "Atención"
                        GoTo Errores
                     End If
                  End If
               Else
                  If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 5 Then
                  ' factura A
                     Respuesta = Principal.PrinterFiscal1.OpenInvoice("T", "C", "A", "1", "P", "12", "I", "I", LblCliente.Caption, "X", "CUIT", Trim(cCuit), "N", Trim(cDire), "-", "-", "-", CmbCond.text, "G")
                     If Respuesta = False Then
                        MsgBox "Problemas Para Generar El Ticket Factura.." & Chr(13) & Chr(10) & "Verifique Cuit, Condifion de Iva o Cliente", vbCritical, "Atención"
                        GoTo Errores
                     End If
                  Else
                     ' Nota de Credito A
                      Respuesta = Principal.PrinterFiscal1.OpenInvoice("M", "C", "A", "1", "P", "12", "I", "I", LblCliente.Caption, "X", "CUIT", Trim(cCuit), "N", Trim(cDire), "-", "-", Str(nFactura), Str(nFactura), "G")
                      If Respuesta = False Then
                         MsgBox "Problemas Para Generar El Ticket Factura.." & Chr(13) & Chr(10) & "Verifique Cuit, Condifion de Iva o Cliente", vbCritical, "Atención"
                         GoTo Errores
                     End If
                  End If
              End If
            Else
              ' Aca Hasar
              
               Principal.HASAR1.Modelo = MODELO_615
               Principal.HASAR1.Puerto = nPuerto
               Principal.HASAR1.Comenzar
               Principal.HASAR1.DescripcionesLargas = True
               Principal.HASAR1.TratarDeCancelarTodo
               Principal.HASAR1.Encabezado(11) = "Orientación al Consumidor Prov. Bs.As"
               Principal.HASAR1.Encabezado(12) = "0800-222-9042"
              
               Select Case cRsl.TraerValorDeUnCampo("Clientes", "TipoIva", "Cliente='" & TxtCliente.text & "'")
                      Case 1
                          sCondiIva = "C"
                      Case 2
                           sCondiIva = "I"
                      Case 3
                           sCondiIva = "N"
                      Case 4
                           sCondiIva = "M"
                      Case 5
                          sCondiIva = "E"
               End Select
              
               If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 0 Then
                  If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 5 Then
                     sComando = Chr$(98) & FS & LblCliente.Caption & FS & cCuit & FS & sCondiIva & FS & "C" & FS & cDire
                     Principal.HASAR1.Enviar sComando
                     If cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 83 Then
                        Principal.HASAR1.AbrirComprobanteFiscal TICKET_C
                        nNumero = Val(Principal.HASAR1.Respuesta(3))
                     Else
                         If cCuit = "99999999999" Then
                            MsgBox "Cliente Sin Numero de Cuit", vbCritical, "Atención"
                            Principal.HASAR1.Finalizar
                            CmdVuelto(0).Enabled = True
                            Exit Sub
                         End If
                         Principal.HASAR1.AbrirComprobanteFiscal TICKET_FACTURA_B
                         nNumero = Val(Principal.HASAR1.Respuesta(3))
                     End If
                  Else
                  ' Nota de Credito B
                        If cCuit = "99999999999" Then
                            MsgBox "Cliente Sin Numero de Cuit", vbCritical, "Atención"
                            Exit Sub
                         End If
                       nFaNo = Format(nFactura, "00000000")
                     
                       sComando = Chr$(147) & FS & "1" & FS & "0000-" & nFaNo
                       Principal.HASAR1.Enviar sComando
      
                       sComando = Chr$(98) & FS & LblCliente.Caption & FS & cCuit & FS & "C" & FS & "C" & FS & cDire
                       Principal.HASAR1.Enviar sComando
             
                       sComando = Chr$(128) & FS & "S" & FS & "T"
                       Principal.HASAR1.Enviar sComando
                       nNumero = Val(Principal.HASAR1.Respuesta(3))
                  End If
               Else
                   If cCuit = "99999999999" Then
                      MsgBox "Cliente Sin Numero de Cuit", vbCritical, "Atención"
                      Principal.HASAR1.Finalizar
                      CmdVuelto(0).Enabled = True
                      Exit Sub
                    End If
                  If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 5 Then
                  ' factura A
                    sComando = Chr$(98) & FS & LblCliente.Caption & FS & cCuit & FS & "I" & FS & "C" & FS & Left(cDire, 20)
                    Principal.HASAR1.Enviar sComando
                    If cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 115 Then
                       Principal.HASAR1.Enviar Chr$(64) & FS + "D" + FS + "T"      ' Nota de Debito A
                       Principal.HASAR1.Enviar sComando
                       nNumero = Val(Principal.HASAR1.Respuesta(3))
                    Else
                       Principal.HASAR1.AbrirComprobanteFiscal TICKET_FACTURA_A
                       nNumero = Val(Principal.HASAR1.Respuesta(3))
                    End If
                  Else
                     ' Nota de Credito A
                      nFaNo = Format(nFactura, "00000000")
                      
                      sComando = Chr$(147) & FS & "1" & FS & "0000-" & nFaNo
                      Principal.HASAR1.Enviar sComando

                      sComando = Chr$(98) & FS & LblCliente.Caption & FS & cCuit & FS & "I" & FS & "C" & FS & Left(cDire, 20)
                      Principal.HASAR1.Enviar sComando

                      sComando = Chr$(128) & FS & "R" & FS & "T"
                      Principal.HASAR1.Enviar sComando
                      nNumero = Val(Principal.HASAR1.Respuesta(3))
                  End If
              End If
            End If
         End If
         
         rAux.MoveFirst
         
         Do While Not rAux.EOF
              Dim cIva As Double
              Dim j As Byte

              Dim sDescripcion As String, sCantidad As String, sUnitario As String, sIva As String, cInter As String
              Dim intFijo As String, IntPorc As String, nIva As Currency

              sDescripcion = Left(rAux!A_DESCRI, 20)
              sCantidad = Round(rAux!a_cant, 4) * 1000
              nIva = 0
              Respuesta = True
              sInter1 = ""
              sInter = cRsl.TraerValorDeUnCampo("Productos", "TipoInterno", "Producto='" & rAux!a_codart & "'")
              nTasa = cRsl.TraerTasaProducto(rAux!a_codart)
              nInter = 0
              pInterPorc = 0: vInterno = 0
              Select Case sInter
                     Case "No"
                          nInter = 0
                          sInter = "0"
                     Case "Porcentaje"
                           sInter = ""
'                          sInter = "0"
                           pInterPorc = (cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & rAux!a_codart & "'") / 100) + 1
                           vInterno = Int(((rAux!a_precio) / pInterPorc) * 100)
                           nInter = vInterno / 100
'                          nInter = rAux!a_precio / (rAux!a_precio + (rAux!a_precio * cRsl.TraerValorDeUnCampo("Productos", "Internos", "Producto='" & rAux!a_codart & "'") / 100))
                          sInter1 = Str(vInterno) & "000000"
                     Case "Fijo"
                          nInter = cRsl.TraerValorDeUnCampo("Productos", "Internos", "Producto='" & rAux!a_codart & "'")
                          sInter = Format(Int(nInter), "00000000")
                          sInter1 = ObtenerDecimales(nInter)
                          sInter1 = sInter1 & String$(8 - Len(sInter1), "0")
                          sInter1 = sInter & sInter1
                         '0000000102900000
                         '0000000102900000
                         '0000000088800000
                         '0000000480000000
                         
              End Select

              nIva = cRsl.TraerTasaProducto(rAux!a_codart)
              If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 0 Or nControlador = 1 Then
                 sUnitario = Format(rAux!a_precio, "0.00")
              Else
                 sUnitario = Round((rAux!a_precio - nInter) / ((nIva / 100) + 1), 3)
              End If
              sIva = nIva * 100
              If nControlador = 2 Then
                 If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 0 Then
                    If cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 83 Then
                      ' sInter1 = "0000002839000000"
                      ' sInter1 = "2839000000"
                       sInter1 = Trim(sInter1)
                       If sCantidad <> 0 Then
                          If Respuesta Then Respuesta = Principal.PrinterFiscal1.SendTicketItem(sDescripcion, sCantidad, sUnitario, sIva, "M", "0", "0", sInter1)
                          If Respuesta = False Then
                             MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                             GoTo Errores
                       End If
                     End If
                    Else
                   ' sInter1 = "2839000000"
                     sInter1 = Trim(sInter1)
                     If sCantidad <> 0 Then
                        If Respuesta Then Respuesta = Principal.PrinterFiscal1.SendInvoiceItem(sDescripcion, sCantidad, sUnitario, sIva, "M", "0", "0", "", "", "", "0", sInter1)
                           If Respuesta = False Then
                              MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                              GoTo Errores
                        End If
                     End If
                    End If
                 Else
                  ' sUnitario = "19.78"
                    sInter1 = Trim(sInter1)
                    If Respuesta Then Respuesta = Principal.PrinterFiscal1.SendInvoiceItem(sDescripcion, sCantidad, sUnitario, sIva, "M", "0", "0", "", "", "", "0", sInter1)
                       If Respuesta = False Then
                          MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                          GoTo Errores
                      End If
                 End If
              Else
                  ' Aca va Hasar
                    Dim dCanti, dUnitario, dIva As Double
                    
                    dCanti = sCantidad / 1000
                    dUnitario = sUnitario ' / 100
                    dIva = sIva / 100
                    Principal.HASAR1.ImprimirItem sDescripcion, dCanti, Round(dUnitario, 4), dIva, 0
              End If
      
            rAux.MoveNext
          Loop
          If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = False Then
              If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 5 Then
                 If nControlador = 2 Then
                    If cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 83 Then
                       If Respuesta Then Respuesta = Principal.PrinterFiscal1.CloseTicket
                          If Respuesta = False Then
                             MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                             GoTo Errores
                          End If
                    Else
                       If Respuesta Then Respuesta = Principal.PrinterFiscal1.CloseInvoice("T", "B", "HOLA")
                          If Respuesta = False Then
                             MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                             GoTo Errores
                           End If
                    End If
                    nNumero = Principal.PrinterFiscal1.AnswerField_3
                    GrabarCae nMovimiento, nNumero
                 Else
                     Dim nDes As Double
                     nDes = CalcularDescuento(dTotalAux)
                     
                     If bDescuento = True And nDes <> 0 Then
                        Principal.HASAR1.DescuentoGeneral "Descuento " & dPorcDescuento & " %", nDes, True
                     End If
                     Principal.HASAR1.SubTotal False
                     Principal.HASAR1.CerrarComprobanteFiscal
                     Principal.HASAR1.Finalizar
                     bCierreControlador = True
                 End If
              Else
                 If nControlador = 2 Then
                    If Respuesta Then Respuesta = Principal.PrinterFiscal1.CloseInvoice("M", "B", "HOLA")
                 Else
                    If bDescuento = True And CalcularDescuento(dTotalAux) <> 0 Then
                       Principal.HASAR1.DescuentoGeneral "Descuento " & dPorcDescuento & " %", CalcularDescuento(dTotalAux), True
                    End If
                    Principal.HASAR1.Enviar Chr$(129)
                    Principal.HASAR1.Finalizar
                    bCierreControlador = True
                 End If
              End If

              cRv.MoveFirst
              Do While Not cRv.EOF
                  cRv!Numero = nNumero
                  cRv.Update
                  cRv.MoveNext
              Loop
              cRv.MoveFirst
          Else
             Respuesta = True
             If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 5 Then
                 If nPercepcion <> 0 Then
                    sPercep = Round(nPercepcion, 2) * 100
                    If nControlador = 2 Then
                       If Respuesta Then Respuesta = Principal.PrinterFiscal1.SendInvoicePerception("Percepcion IB Bs.As.", "O", sPercep)
                          If Respuesta = False Then
                             MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                             GoTo Errores
                          End If
                    Else
                       ' Aca Hasar
                       If nPercepcion <> 0 Then
                          Principal.HASAR1.EspecificarPercepcionGlobal "Percepción IB Bs.As.", CDbl(sPercep)
                       End If
                    End If
                 End If
                 If nControlador = 2 Then
                    If Respuesta Then Respuesta = Principal.PrinterFiscal1.CloseInvoice("T", "A", "HOLA")
                       If Respuesta = False Then
                          MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                          GoTo Errores
                      End If
                 Else
                  ' Aca Hasar
                     If bDescuento = True And CalcularDescuento(dTotalAux) <> 0 Then
                       Principal.HASAR1.DescuentoGeneral "Descuento " & dPorcDescuento & " %", CalcularDescuento(dTotalAux), True
                    End If
                     Principal.HASAR1.SubTotal False
                     Principal.HASAR1.CerrarComprobanteFiscal
                 End If
              Else
                 If nControlador = 2 Then
                    If Respuesta Then Respuesta = Principal.PrinterFiscal1.CloseInvoice("M", "A", "HOLA")
                       If Respuesta = False Then
                          MsgBox "Problemas Para Generar El Ticket Factura..", vbCritical, "Atención"
                          GoTo Errores
                      End If
                 Else
                   ' Aca Hasar
                   Principal.HASAR1.SubTotal False
                   Principal.HASAR1.Enviar Chr$(129)
                 End If
              End If
              If nControlador = 2 Then
                 nNumero = Principal.PrinterFiscal1.AnswerField_3
                 GrabarCae nMovimiento, nNumero
              Else
                 Principal.HASAR1.Finalizar
                 bCierreControlador = True
              End If
              cRv!Numero = nNumero
              cRv.Update
          End If
          
          
          If CmbCond.ItemData(CmbCond.ListIndex) = 0 And TxtCliente.text <> sConsFinal Then
              Dim cSal As ClsClienteL
              
              Set cSal = New ClsClienteL

              If nControlador = 2 Then
                 Respuesta = Principal.PrinterFiscal1.OpenNoFiscal
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("Fecha :" & Date)
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("CUENTA CORRIENTE")
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText(Trim(LblCliente.Caption))
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("IMPORTE :" & CCur(Label8.Caption))
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("Saldo Anterior :" & Format(cSal.TraerSaldoCtaCte(TxtCliente.text), "#0.00"))
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("------------------------------------")
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("     Firma Cliente")
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("------------------------------------")
                 Respuesta = Principal.PrinterFiscal1.SendNoFiscalText("        Aclaración")
                 Respuesta = Principal.PrinterFiscal1.CloseNoFiscal
              Else
                 ' Aca va Hasar
                  Principal.HASAR1.Modelo = MODELO_615
                  Principal.HASAR1.Puerto = nPuerto
                  Principal.HASAR1.Comenzar
                  Principal.HASAR1.DescripcionesLargas = True
                  Principal.HASAR1.TratarDeCancelarTodo
                  Principal.HASAR1.AbrirComprobanteNoFiscal
                  Principal.HASAR1.ImprimirTextoNoFiscal ("Fecha :" & Date)
                  Principal.HASAR1.ImprimirTextoNoFiscal ("CUENTA CORRIENTE")
                  Principal.HASAR1.ImprimirTextoNoFiscal (Trim(LblCliente.Caption))
                  Principal.HASAR1.ImprimirTextoNoFiscal ("IMPORTE:" & CCur(Label8.Caption))
                  Principal.HASAR1.ImprimirTextoNoFiscal ("Saldo Anterior :" & Format(cSal.TraerSaldoCtaCte(TxtCliente.text), "#0.00"))
                  Principal.HASAR1.ImprimirTextoNoFiscal ("---------------------------------")
                  Principal.HASAR1.ImprimirTextoNoFiscal ("     Firma Cliente")
                  Principal.HASAR1.ImprimirTextoNoFiscal ("---------------------------------")
                  Principal.HASAR1.ImprimirTextoNoFiscal ("        Aclaración")
                  Principal.HASAR1.CerrarComprobanteNoFiscal
                  Principal.HASAR1.Finalizar
              End If
           End If
        Else
'            Dim nCan As Currency, nUni As Currency, nTot As Currency
'            Dim i As Integer, X As Integer, Y As Single
'
'             Dim P As Printer, sImpre As String
'
'
'             sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & CmbCompro.ItemData(CmbCompro.ListIndex) & " AND Usuario=" & nUsuario)
'
'             If sImpre <> "Ninguna" And sImpre <> "0" Then
'
'                 For Each P In Printers
'                     If P.DeviceName = sImpre Then
'                        Set Printer = P
'                        Printer.FontName = "Courier New"
'                        Printer.ScaleMode = 7
'                        Printer.FontSize = 9
'                        Exit For
'                     End If
'                 Next
'
'                 Y = 1
'                 X = 0
'
'                 Printer.CurrentY = Y
'                 Printer.CurrentX = X
'                 Printer.Print " "
'
'                 rAux.MoveFirst
'
'                 Y = Y + 1
'
'                 Do While Not Rsd.EOF
'                    Printer.CurrentY = Y
'                    Printer.CurrentX = X
'
'                    nCan = Format(rAux!a_cant, "#0.00")
'                    nUni = Format(rAux!a_precio, "#0.00")
'                    nTot = rAux!a_precio * rAux!a_cant
'                    nTot = Format(nTot, "#0.00")
'                    Printer.Print Space(5 - Len(Format(rAux!a_cant, "0.00"))) & Format(rAux!a_cant, "0.00") & " " & Left(rAux!a_descri, 16) & Space(16 - Len(Left(rAux!a_descri, 16))) & Space(7 - Len(Format(nUni, "#0.00"))) & " " & Format(nUni, "0.00") & Space(8 - Len(Format(nTot, "#0.00"))) & " " & Format(nTot, "0.00")
'                       ' Principal.HASAR1.ImprimirTextoNoFiscal (Left(rAux!a_descri, 17) & Space(17 - Len(Left(rAux!a_descri, 17))) & " " & Format(nCan, "0.00") & Space(8 - Len(Format(nUni, "#0.00"))) & " " & Format(nUni, "0.00") & Space(8 - Len(Format(nTot, "#0.00"))) & " " & Format(nTot, "0.00"))
'                    Y = Y + 0.5
'                      ' x = x + 1
'                    rAux.MoveNext
'                 Loop
'
'                '
'                Y = Y + 1
'
'                Printer.CurrentY = Y
'                Printer.CurrentX = X
'
'                Printer.Print "Importe Total :" & Format(label8.caption, "#0.00")
'
'                Y = Y + 2
'
'                Printer.CurrentY = Y
'                Printer.CurrentX = X
'                Printer.Print " "
'
'                Y = Y + 1
'
'                Printer.CurrentY = Y
'                Printer.CurrentX = X
'                Printer.Print " "
'
'                Y = Y + 1
'
'                Printer.CurrentY = Y
'                Printer.CurrentX = X
'                Printer.Print " "
'
'                Y = Y + 1
'                Printer.CurrentY = Y
'                Printer.CurrentX = X
'                Printer.Print " "
'
'                Printer.EndDoc
'             End If
          End If
     End If
     
grabarTicket:
     Dim RsCv As ADODB.Recordset
     Dim RsDa As ADODB.Recordset
     Dim rReci As ADODB.Recordset
     Dim rCant As ADODB.Recordset
     Dim rValo As ADODB.Recordset
     Dim rCaja As ADODB.Recordset
     Dim rApli As ADODB.Recordset
     Dim rRecibo As ADODB.Recordset
     Dim nCuentaResto As Integer
    ' Grabar Cabecera

     Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
     Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
     Set RsDa = cRsl.RsVacio("DetallesComprobantes", "Id", "N")
     Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
     Set rValo = cRsl.RsVacio("Valores", "Id", "N")
     Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
     Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
     Set rRecibo = cRsl.RsVacio("CabComprobantes", "Id", "N")
     
     
     nCuentaResto = cRsl.RegPorDefecto("CuentaRes")
     If bDelivery = True Then
        nCuentaResto = cRsl.RegPorDefecto("CuentaDel")
     End If
           
     cRv.MoveFirst
           
     RsCv.AddNew
     RsCv!Suc = nSucursal
     RsCv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
     RsCv!Movimiento = nMovimiento
     RsCv!Comprobante = cRv!Comprobante
     RsCv!fecha = Date
     RsCv!FechaIva = Date
     RsCv!Numero = nNumero
     RsCv!Sucursal = cRv!Sucursal
     RsCv!Cliente = cRv!Cliente
     RsCv!Lista = cRv!Lista
     RsCv!Vendedor = cRv!Vendedor
     RsCv!Caja = cRv!Caja
     RsCv!CondVenta = CmbCond.ItemData(CmbCond.ListIndex)
           
     Label8.Caption = 0
     LblTotalGrabado.Caption = 0
     LblTotalIva.Caption = 0
     LblTotalNoGrabado.Caption = 0

     nTotal = 0
     nGrabado = 0
     nIvas = 0
     nNoGrabado = 0


     rAux.MoveFirst

     Do While Not rAux.EOF
        RsDa.AddNew
        RsDa!Movimiento = RsCv!Movimiento
        RsDa!Producto = rAux!a_codart
        RsDa!Descripcion = rAux!A_DESCRI
        RsDa!Cantidad = rAux!a_cant
        RsDa!Cuenta = nCuentaResto
        RsDa!Descuento = 0
        RsDa!Tasa = cRsl.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & rAux!a_codart & "'")
        RsDa!Deposito = 1
        RsDa!Medida = cRsl.TraerValorDeUnCampo("Productos", "UMVenta", "Producto='" & rAux!a_codart & "'")
        If Frame3.Caption = "N" Then
           RsDa!Impuesto = 0
           RsDa!PrecioTotal = rAux!a_total
           If rAux!a_cant <> 0 Then
              RsDa!PrecioUnitario = rAux!a_total / rAux!a_cant
           Else
              RsDa!PrecioUnitario = rAux!a_total
           End If
           nTotal = nTotal + RsDa!PrecioTotal
           nGrabado = nGrabado + RsDa!PrecioTotal
        Else
           sInter = cRsl.TraerValorDeUnCampo("Productos", "TipoInterno", "Producto='" & rAux!a_codart & "'")
           nInter = 0
           pInterPorc = 0: vInterno = 0
           Select Case sInter
                  Case "No"
                        nInter = 0
                        sInter = "0"
                   Case "Porcentaje"
                         sInter = ""
'                          sInter = "0"
                         pInterPorc = (cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & rAux!a_codart & "'") / 100) + 1
                         vInterno = Int(((rAux!a_precio) / pInterPorc) * 100)
                         nInter = vInterno / 100
                   Case "Fijo"
                         nInter = cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & rAux!a_codart & "'")
          End Select
          nTasa = cRsl.TraerTasaProducto(rAux!a_codart)
          RsDa!Impuesto = ((rAux!a_total - (rAux!a_cant * nInter)) - (rAux!a_total - (rAux!a_cant * nInter)) / ((nTasa / 100) + 1)) '/ rAux!a_cant
  
          nNoGrabado = nNoGrabado + (rAux!a_cant * nInter)
          If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = False Then
             RsDa!PrecioTotal = rAux!a_total
             RsDa!PrecioUnitario = RsDa!PrecioTotal / rAux!a_cant
           '  RsDa!PrecioUnitario = Round((rAux!a_precio - nInter) / ((nIva / 100) + 1), 3)
             nGrabado = nGrabado + rAux!a_total - (rAux!a_cant * nInter)
          Else
             nGrabado = nGrabado + (rAux!a_total - (rAux!a_cant * nInter)) / ((nTasa / 100) + 1)
             RsDa!PrecioTotal = (rAux!a_total - (rAux!a_cant * nInter)) / ((nTasa / 100) + 1)
             RsDa!PrecioUnitario = RsDa!PrecioTotal / rAux!a_cant
             nIvas = nIvas + (RsDa!Impuesto * rAux!a_cant)
          End If
          nTotal = nTotal + rAux!a_total
        End If
        
        RsDa!MoviRemi = 0
        RsDa.Update

       If cRsl.TraerValorDeUnCampo("Productos", "ActuaStock", "Producto='" & rAux!a_codart & "'") = "Si" Then
          rCant.AddNew
          rCant!Movimiento = RsCv!Movimiento
          rCant!Estado = 1
          rCant!Producto = rAux!a_codart
          rCant!Deposito = 1
          Select Case cNum.QueMoviStock(CmbCompro.ItemData(CmbCompro.ListIndex))
                   Case "Suma"
                        rCant!CantidadDebe = rAux!a_cant
                        rCant!CantidadHaber = 0
                   Case "Resta"
                        rCant!CantidadDebe = 0
                        rCant!CantidadHaber = rAux!a_cant
          End Select
          rCant.Update
        End If
        rAux.MoveNext
     Loop

     Label8.Caption = Round(nTotal, 2)
     LblTotalGrabado.Caption = nGrabado
     LblTotalIva.Caption = nIvas
     LblTotalNoGrabado.Caption = nNoGrabado
     
     RsCv!Neto = CCur(LblTotalGrabado.Caption)
     RsCv!Iva1 = CCur(LblTotalIva.Caption)
     RsCv!Iva2 = 0
     RsCv!Descuentos = 0
     RsCv!Financiacion = 0
     RsCv!NoGravados = CCur(LblTotalNoGrabado.Caption)
     RsCv!PIvaCompras = 0
     RsCv!IBrutosCompras = 0
      Select Case cRsl.TraerValorDeUnCampo("Comprobantes", "TipoMovimiento", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
             Case 1, 11
                  RsCv!Debe = CCur(Label8.Caption)
                  RsCv!Haber = 0
             Case 2, 12, 13
                  RsCv!Debe = 0
                  RsCv!Haber = CCur(Label8.Caption)
      End Select
      RsCv!Desde = 2
      RsCv!Hasta = 0
      RsCv!factura = 0
      RsCv!Motivo = ""
      RsCv!COTIZACION = 1
      RsCv!Anulado = 0
      RsCv.Update

        ' Grabar Cuenta Corriente
     If CmbCond.ItemData(CmbCond.ListIndex) = 0 And TxtCliente.text <> sConsFinal Then
         rReci.AddNew
         rReci!fecha = Date
         rReci!Cliente = cRv!Cliente
         rReci!Venta = RsCv!Movimiento
         rReci!Comprobante = cRv!Comprobante
         rReci!Numero = cRv!Numero
         rReci!Sucursal = cRv!Sucursal
         rReci!Cuota = 1
         Select Case cRsl.TraerValorDeUnCampo("Comprobantes", "TipoMovimiento", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
                Case 1, 11
                     rReci!Debe = Round(nTotal, 2)
                     rReci!Haber = 0
                Case 2, 12, 13
                     rReci!Debe = 0
                     rReci!Haber = Round(nTotal, 2)
         End Select
         rReci!Anulado = 0
         rReci.Update
     Else
'        Dim nNum, nSuc As Long
      
      
       ' Grabar Valores
         rValo.AddNew
         rValo!Venta = RsCv!Movimiento
    '     rValo!Venta = nMoviAfip
         rValo!CodPago = CmbForma.ItemData(CmbForma.ListIndex)
         rValo!formapago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", CmbForma.ItemData(CmbForma.ListIndex))
         rValo!Nombre = ""
         rValo!importe = Label8.Caption
         rValo.Update
      
      
      ' Actualizo Caja
         rValo.MoveFirst
         rCaja.AddNew
         rCaja!Suc = nSucursal
         rCaja!Movimiento = RsCv!Movimiento
         rCaja!Comprobante = cRv!Comprobante
         rCaja!Numero = cRv!Numero
         rCaja!Sucursal = cRv!Sucursal
         rCaja!fecha = Date
         rCaja!Hora = Time
         rCaja!Caja = cRsl.RegPorDefecto("Caja")
         rCaja!formapago = rValo!formapago
         rCaja!Nombre = LblCliente.Caption
         rCaja!Banco = rValo!Banco
         rCaja!NumeroCheque = rValo!NumeroCheque
         rCaja!FechaAcreditacion = rValo!FechaAcreditacion
         rCaja!NumeroTarjeta = rValo!NumeroTarjeta
         rCaja!FechaVencimiento = rValo!FechaVencimiento
         rCaja!Autorizacion = rValo!Autorizacion
         rCaja!Cupon = rValo!Cupon
         If cNum.SumaCaja(CmbCompro.ItemData(CmbCompro.ListIndex)) Then
            rCaja!Debe = rValo!importe
            rCaja!Haber = 0
         Else
            rCaja!Debe = 0
            rCaja!Haber = rValo!importe
         End If
         rCaja!Abierta = "S"
         rCaja!Anulado = 0
         rCaja.Update
      
     End If
                                                                                                                                                                                                         
     cPe.AgregarConReciboResto RsCv, RsDa, rReci, rValo, rApli, rCant, rCaja, rRecibo
      
     bGraboTick = True
      
     Set cRv = Nothing
     Set cRle = Nothing
   
     If bDelivery = False Then
        FrmAdicion.CmdMesa(nMesa).BackColor = &HC000&
     Else
        Delivery.LlenarGrilla
     End If

     bCerrar = True
     Unload Me

   Unload Me
   Set cRsl = Nothing
Exit Sub
Errores:

  If err.Number <> 0 Then
     MsgBox err.Description, vbCritical, "Atención"
     cRsl.GrabarError "CmdVuelto", err.Description, "Numero:" & cRsl.TraerValorDeUnCampo("Comprobantes", "Numero", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) & " Movimiento=" & nMovimiento
  Else
     MsgBox "Controlador Apagado, fuera de Línea, Cable Desconectado o Falta Papel...", vbCritical, "Atención"
  End If

   Frame3.Visible = False
'   HabilitarDetalles bCod, False
'   CmdDetalle(0).Enabled = bBoto
   Label8.Caption = Format(dTotalAux, "#0.00")
   Grid1.Enabled = True
   If nControlador = 1 Then
      If bCierreControlador = False Then
         Principal.HASAR1.Abortar
         Principal.HASAR1.Finalizar
      End If
      If bGraboTick = False Then
         GoTo grabarTicket
      End If
   Else
      
   End If
   Exit Sub
End Sub



Private Sub Combo2_Click()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura

  CalcularDescuentoGrilla cRsl.TraerValorDeUnCampo("Descuentos", "Porcentaje", "Descuento=" & Combo2.ItemData(Combo2.ListIndex))
   LblCompra.Caption = Label8.Caption
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Errores
   Dim cRsl As ClsLectura
   
   Set cRsl = New ClsLectura
   
   If KeyCode = vbKeyF4 Then
      If Frame3.Caption = "N" Then
         Frame3.Caption = "P"
         CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinal"))
         CmbCompro.Enabled = False
      Else
         Frame3.Caption = "N"
         CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinalN"))
         CmbCompro.Enabled = False
      End If
      TxtCliente.text = sConsFinal
      LblCliente.Caption = sNomFinal
      If TxtPago.Visible = True Then
         TxtCliente.SetFocus
      End If
   End If
   If Frame3.Visible = False And sw = False Then
      If KeyCode = vbKeyF1 Then
          Text1.text = ""
          nDat = 0
          nBuscar = 3
          nLlama = 5
          FrmBuscarEx.Show 1
          If nDat <> 0 Then
             Text1.text = cRsl.TraerValorDeUnCampo("Productos", "Producto", "Id=" & nDat)
             Text1_LostFocus
          End If
      End If
  Else
      If KeyCode = vbKeyF5 Then
         TxtCliente.SetFocus
      End If
      If KeyCode = vbKeyF6 Then
         CmbCompro.Enabled = True
         CmbCompro.SetFocus
      End If
      If KeyCode = vbKeyF7 Then
         TxtPago.SetFocus
      End If
      If KeyCode = vbKeyF11 Then
         CmbCond.SetFocus
      End If
      If KeyCode = vbKeyF12 Then
         CmbForma.SetFocus
      End If
  End If
  Exit Sub
Errores:
   MsgBox err.Description, vbCritical, "Atención"
   cRsl.GrabarError "Key Down", err.Description, Str(KeyCode)
End Sub

Private Sub TxtPago_GotFocus()
   If CmbCompro.Enabled = False Then
      TxtPago.SelStart = 0
      TxtPago.SelLength = Len(TxtPago.text)
   Else
       If CmbCompro.Enabled = True And CmbCompro.Visible = True Then
          CmbCompro.SetFocus
       End If
   End If
End Sub

Private Sub TxtPago_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtPago) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPago_LostFocus()
  On Error GoTo Errores
  If TxtPago.text = "" Then TxtPago.text = 0
  TxtPago.text = Format(TxtPago.text, "#0.00")
  LblVuelto.Caption = Format(CDbl(TxtPago.text) - CDbl(LblCompra.Caption), "#0.00")
Exit Sub
Errores:
 TxtPago.text = "0.00"
 TxtPago.SetFocus
 TxtPago.SelStart = 0
 TxtPago.SelLength = Len(TxtPago.text)
End Sub

Private Function VerificarCuenta(pCliente As String) As Boolean

 If UCase(CmbCond.text) <> "CONTADO" Then

   Dim nSaldo As Currency, nLimite As Currency, cRcL As ClsClienteL, cRsl As ClsLectura
   Dim dif As Integer, UltPago As Date, sSaldo As Currency
   Dim nApro As Boolean, nDiferencia As Currency
        
   nApro = True
  
   Set cRcL = New ClsClienteL
   Set cRsl = New ClsLectura
  
  If Not cRsl.TraerValorDeUnCampo("Clientes", "FechaBaja", "Cliente='" & pCliente & "'") = "" Then
     MsgBox "Cliente Con Cuenta Inactiva", vbCritical, "Atención"
     nApro = False
     VerificarCuenta = False
     Exit Function
  End If
  
   UltPago = cRcL.TraerUltimoPago(pCliente)
   dif = DateDiff("d", UltPago, Date)
   If dif > nDias Then
      nApro = False
   End If

   sSaldo = cRcL.ResumenSaldoAnteriorBus(pCliente)
   nLimite = cRsl.TraerValorDeUnCampo("Clientes", "Limite", "Cliente='" & pCliente & "'")
                    
   If sSaldo > 1 And nLimite <> 0 Then
      If sSaldo + CCur(Label8.Caption) > nLimite Then
         nDiferencia = sSaldo + CCur(Label8.Caption) - nLimite
         nApro = False
      End If
   Else
      If CCur(Label8.Caption) > nLimite And nLimite <> 0 Then
         nApro = False
         dif = 0
      Else
         nApro = True
      End If
   End If
  
   If nApro = False Then
     MsgBox "Limite de Fecha de Pago Exedido :" & dif & Chr(10) & Chr(13) & "Limite de Credito Exedido :" & Format(nDiferencia, "#0.00"), vbCritical, "Atención"
     FrmClave.Show 1
     If bClave = True Then
        VerificarCuenta = True
     Else
        VerificarCuenta = False
     End If
   Else
      VerificarCuenta = True
   End If
  Else
    VerificarCuenta = True
  End If
End Function

Private Sub CmbForma_Click()
  Dim dTotal As Double, cRsl As ClsLectura
   
   Set cRsl = New ClsLectura
   dTotal = dTotalAux
  
   dTotal = dTotalAux - CalcularDescuento(dTotal)
   Label8.Caption = Format(dTotal, "#0.00")
   LblCompra.Caption = Label8.Caption
    
   TxtPago_LostFocus
  
   Recalcular cRsl.TraerValorDeUnCampo("CondVenta", "RgTarjeta", "CondVta=" & CmbForma.ItemData(CmbForma.ListIndex))
End Sub

Private Sub Command2_Click()
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura
  nBuscar = 3
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Text1.text = cRsl.TraerValorDeUnCampo("Productos", "Producto", "Id=" & nDat)
     Label5.Caption = cRsl.TraerValorDeUnCampo("Productos", "Descripcion", "Id=" & nDat)
     Text3.SetFocus
  End If
  Set cRsl = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = 27 Then
      Unload Me
              FrmAdicion.TxtMesa.SetFocus
      Exit Sub
     End If
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
     End If
     KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Command1_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Borrar
         Case 3
              salir
         Case 2
              CierreRapido
         Case 4
              CierreMesa
         Case 5
              If bRecibo = True Then
                ImprimirA4
              Else
                ImpimirComanda
              End If
         Case 6
              If bRecibo = True Then
                 Presupuesto
              Else
                 Comanda
              End If
  End Select
End Sub

Private Sub ImprimirA4()
  Dim P As Printer, cRsl As ClsLectura, sImpre As String, i, nCompro As Integer, sSql As String

  
  Set cRsl = New ClsLectura
    
  Dim rOrd As Recordset
  Set rOrd = cRsl.TraerRsCondi("Auxiliar", "Id", "a_mesa=" & Label2.Caption & " and a_mensaje='ABIERTA' and a_impcom='N'")
  
  If rOrd.RecordCount <> 0 Then
    nCompro = cRsl.RegPorDefecto("VentaConsFinalN")
    sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & nCompro & " AND Usuario=" & nUsuario)
    
    If sImpre = "Ninguna" Then
        MsgBox "Debe Definir Un Modelo de Impresora al Comprobante...", vbCritical, "Atención"
        Exit Sub
    End If
    For Each P In Printers
        If P.DeviceName = sImpre Then
           Set Printer = P
           Printer.FontName = "Courier New"
           Printer.ScaleMode = 7
           Printer.FontSize = 9
           Exit For
       End If
    Next
  
       Printer.FontSize = 8
       Printer.ScaleMode = 7
       Printer.FontBold = True
       Printer.CurrentY = 1
       Printer.CurrentX = 0.5
       Printer.Print "FECHA :" & Date
       Printer.CurrentY = 1.5
       Printer.CurrentX = 0.5
       Printer.Print "HORA :" & Time
       Printer.CurrentY = 2
       Printer.CurrentX = 0.5
       Printer.Print "MESA Nº :" & Label2.Caption
       Printer.CurrentY = 2
       Printer.CurrentX = 4
       Printer.Print "MOZO :" & cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & nMozo)
       Printer.CurrentY = 2.5
       Printer.CurrentX = 0.5
       Printer.FontBold = False
       Printer.Print "_________________________________________________________________________________________________________"
       Printer.CurrentY = 2.8
       Printer.CurrentX = 0.5
       Printer.Print "Cantidad         D e s c r i p c i ó n                          Observación "
       Printer.CurrentY = 2.9
       Printer.CurrentX = 0.5
       Printer.Print "_________________________________________________________________________________________________________"
       i = 3.3
  
       db.Open sDb
       db.BeginTrans
  
       Printer.FontSize = 11

       Do While Not rOrd.EOF
          Dim j, k, cc
          If rOrd![a_impcom] = "N" Then
             Printer.CurrentY = i
             Printer.CurrentX = 0.6
             Printer.Print rOrd![a_cant]
             Printer.CurrentY = i
             Printer.CurrentX = 3
             Printer.Print rOrd![A_DESCRI]
             Printer.CurrentY = i
             Printer.CurrentX = 10
             k = 1
             cc = 10
             If Not IsNull(rOrd![a_Observ]) Then
                For j = 1 To Len(rOrd![a_Observ])
                    Printer.CurrentX = cc
                    Printer.CurrentY = i
                    Printer.Print Mid(rOrd![a_Observ], j, 1)
                    If k = 40 Then
                       k = 1
                       cc = 10
                       i = i + 0.2
                    Else
                       cc = cc + 0.18
                       If Mid(rOrd![a_Observ], j, 1) = "I" Then
                          cc = cc - 0.1
                       End If
                       k = k + 1
                     End If
                  Next j
              End If
          
              sSql = "UPDATE Auxiliar set a_impcom='S' WHERE Id=" & rOrd!Id
              db.Execute sSql
              
  '            rOrd![a_impcom] = "S"
  '            rOrd.Update
         End If
nprint:
            rOrd.MoveNext
            If rOrd.EOF Then
               Exit Do
            End If
            i = i + 0.5
      Loop
      db.CommitTrans
      db.Close
  
      Printer.EndDoc
  Else
     MsgBox "La Comanda de Está Mesa ya fué Impresa", vbCritical, "Atencion"
  End If
  rOrd.Close
End Sub

Private Sub ImpimirComanda()
 Dim P As Printer, cRsl As ClsLectura, sImpre As String, i, nCompro As Integer, sSql As String

  
  Set cRsl = New ClsLectura
    
  Dim rOrd As Recordset
  Set rOrd = cRsl.TraerRsCondi("Auxiliar", "Id", "a_mesa=" & Label2.Caption & " and a_mensaje='ABIERTA' and a_impcom='N'")
  
  If rOrd.RecordCount <> 0 Then
    nCompro = cRsl.RegPorDefecto("VentaConsFinalN")
    sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & nCompro & " AND Usuario=" & nUsuario)
    
    If sImpre = "Ninguna" Then
        MsgBox "Debe Definir Un Modelo de Impresora al Comprobante...", vbCritical, "Atención"
        Exit Sub
    End If
    For Each P In Printers
        If P.DeviceName = sImpre Then
           Set Printer = P
           Printer.FontName = "Courier New"
           Printer.ScaleMode = 7
           Printer.FontSize = 9
           Printer.FontBold = True
           Exit For
       End If
    Next
  
       Printer.FontSize = 8
       Printer.ScaleMode = 7
       Printer.CurrentY = 1
       Printer.CurrentX = 0.1
       Printer.Print "PUESTO : COCINA"
       Printer.CurrentY = 1.5
       Printer.CurrentX = 0.1
       Printer.Print "FECHA :" & Date
       Printer.CurrentY = 2
       Printer.CurrentX = 0.1
       Printer.Print "HORA :" & Time
       Printer.CurrentY = 2.5
       Printer.CurrentX = 0.1
       Printer.Print "MESA Nº :" & Label2.Caption
       Printer.CurrentY = 3
       Printer.CurrentX = 0.1
       Printer.Print "MOZO Nº :" & cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & nMozo)
       Printer.CurrentY = 3.5
       Printer.CurrentX = 0.1
       Printer.Print "========================="
       Printer.CurrentY = 3.7
       Printer.CurrentX = 0.1
       Printer.Print "CANT         DETALLE          "
       Printer.CurrentY = 4
       Printer.CurrentX = 0.1
       Printer.Print "========================="
       i = 4.3
  
       db.Open sDb
       db.BeginTrans


       Do While Not rOrd.EOF
          Dim j, k, cc
          If rOrd![a_impcom] = "N" Then
             Printer.CurrentY = i
             Printer.CurrentX = 0.2
             Printer.Print rOrd![a_cant]
             Printer.CurrentY = i
             Printer.CurrentX = 1
             Printer.Print rOrd![A_DESCRI]
             If Not IsNull(rOrd![a_Observ]) Then
                k = 1
                cc = 0.4
                i = i + 0.3
                Printer.Print " -------------------"
                i = i + 0.2
                For j = 1 To Len(rOrd![a_Observ])
                    Printer.CurrentY = i
                    Printer.CurrentX = cc
                    Printer.Print Mid(rOrd![a_Observ], j, 1)
                    If k = 18 Then
                       k = 1
                       cc = 0.3
                       i = i + 0.3
                    Else
                       cc = cc + 0.18
'                       If Mid(rOrd![a_Observ], j, 1) = "I" Then
'                          cc = cc - 0.1
'                       End If
                       k = k + 1
                     End If
                  Next j
                  ' i = i - 0.3
                '  Printer.CurrentY = i
                  Printer.Print " -------------------"
                  i = i + 0.2

              End If
'              rOrd![a_impcom] = "S"
'              rOrd.Update

              sSql = "UPDATE Auxiliar set a_impcom='S' WHERE Id=" & rOrd!Id
              db.Execute sSql

         End If
         rOrd.MoveNext
            If rOrd.EOF Then
               Exit Do
            End If
            i = i + 0.3
      Loop
      Printer.EndDoc
      
      db.CommitTrans
      db.Close
   Else
       MsgBox "La Comanda de Está Mesa ya fué Impresa", vbCritical, "Atencion"
   End If
   rOrd.Close
End Sub


Private Sub CierreMesa()
   Dim cRsl As ClsLectura
   
   Set cRsl = New ClsLectura
   If Command1(4).Caption = "Cerrar Venta" Then
      If rAux.RecordCount <> 0 Then
         bFoco = False
         Habilitar False, False, False, False, False, False, False, False, False
         Botones False, False, False, False, True, False
         Command1(4).Caption = "Regresar"
         Frame3.Visible = True
         Recalcular cRsl.TraerValorDeUnCampo("CondVenta", "RgTarjeta", "CondVta=" & CmbForma.ItemData(CmbForma.ListIndex))
         dTotalAux = Label8.Caption
         CargarFrame
         TxtCliente.Enabled = True
         BtnBuscar.Enabled = False
         CmbCond.Enabled = False
         CmdVuelto(0).Enabled = True
         If bDelivery = True Then
            If rAux.RecordCount <> 0 Then
               rAux.MoveFirst
               TxtCliente.text = rAux![a_puerto2]
            End If
            TxtCliente_LostFocus
         Else
            TxtCliente.SetFocus
            TxtCliente.SelStart = 0
            TxtCliente.SelLength = Len(TxtCliente.text)
         End If
      End If
   Else
      bFoco = True
      Command1(4).Caption = "Cerrar Venta"
      Frame3.Visible = False
      Habilitar True, False, False, False, False, True, True, True, True
      Botones True, True, True, True, True, True
      Label8.Caption = Format(dTotalAux, "#0.00")
      Grid1.Enabled = True
      Recalcular 0
     '  Erase nUnitario
   End If
End Sub

Private Sub Limpiar()

  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  

  TxtCliente.text = sConsFinal
  LblCliente.Caption = sNomFinal
  
  LblCompra.Caption = "0,00"
  TxtPago.text = "0,00"
  LblVuelto.Caption = "0,00"
  Label8.Caption = "0,00"
'  TxtId.text = ""
'  TxtCPrecio.text = ""
'  label8.caption = "0,00"
  LblTotalGrabado.Caption = "0,00"
  LblTotalNoGrabado.Caption = "0,00"
  LblTotalIva.Caption = "0,00"
  
  Set cRsl = Nothing
End Sub

Private Sub CargarFrame()
  
  Me.Refresh
'  bCod = TxtProducto.Enabled
'  bBoto = CmdDetalle(0).Enabled
'
'  HabilitarDetalles False, False
'  CmdDetalle(0).Enabled = False
  Grid1.Enabled = False
 ' TxtPago.SetFocus
  
  LblCompra.Caption = Label8.Caption
End Sub

Private Sub Recalcular(pValor As Double)
'  Dim i As Integer
'  label8.caption = 0
'
'  i = 0
'
'  If Not Rsd Is Nothing Then
'    If Rsd.RecordCount <> 0 Then
'      If EstaArrayVacio(nUnitario) Then
'          ReDim Preserve nUnitario(Rsd.RecordCount)
'          Rsd.MoveFirst
'          Do While Not Rsd.EOF
'             nUnitario(i) = rAux!a_precio
'             i = i + 1
'             Rsd.MoveNext
'          Loop
'
'       End If
'       Rsd.MoveFirst
'       i = 0
'       Do While Not Rsd.EOF
'          rAux!a_precio = nUnitario(i) + (nUnitario(i) * pValor / 100)
'          rAux!a_total = rAux!a_cant * rAux!a_precio
'          Rsd.Update
'          label8.caption = label8.caption + rAux!a_total
'          i = i + 1
'          Rsd.MoveNext
'       Loop
'    End If
'  End If
'
'  label8.caption = Format(label8.caption, "#0.00")
'  dTotalAux = label8.caption

End Sub

Private Sub Form_Load()
  Dim cRsl As ClsLectura, sCondi As String
  
  Set cRsl = New ClsLectura
  sw = True
  Habilitar False, False, False, False, False, False, False, False, False
  Mostrar "", "", 0, ""
  Limpiar
  sw1 = True
  Command1(0).Picture = LoadResPicture("Nuevo", 0)
  Command1(1).Picture = LoadResPicture("Borrar", 0)
  Command1(3).Picture = LoadResPicture("Salir", 0)
  Command1(5).Picture = LoadResPicture("Imprimir", 0)
  CargarCombo
  
  sCondi = "ABIERTA"
  
  If bDelivery = True Then
     sCondi = "DELIVERY"
     Text3.MaxLength = 30
  End If
  
  Set rAux = cRsl.TraerRsCondi("Auxiliar", "Id", "a_mesa=" & nMesa & " AND a_mensaje='" & sCondi & "'")
  
  Hora = rAux![a_hora]
  nMozo = rAux![a_mozo]
  nLista = rAux![a_lista]
  
  ReDim Preserve dObservacion(0)

  If bDelivery = True Then

     Dim rDel As Recordset
     
     Set rDel = cRsl.TraerRsCondi("Clientes", "Id", "Cliente='" & rAux![a_puerto2] & "'")
     dNombre = rDel!RazonSocial
     dDireccion = rDel!Domicilio
     dObservacion(0) = rAux![a_Observ]
     dHora = rAux![a_hora]
     
     Me.Caption = dNombre & " - " & dDireccion
  End If
  
  CabGrid
  Actualiza
  Botones True, False, True, True, True, True
  Frame3.Visible = False
  Frame3.Caption = IIf(cRsl.RegPorDefecto("TipoVenta") = 0, "N", "P")
  If Frame3.Caption = "N" Then
     CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinalN"))
  Else
     CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinal"))
  End If
  
  bFoco = False
  
  Me.Top = 0
  Me.Left = 0
  CmbCompro.Enabled = False

End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = rAux
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
  With Grid1
       .Columns(0).Visible = False
       .Columns(1).Width = 900
       .Columns(1).Caption = "Codigo"
       .Columns(2).Width = 2500
       .Columns(2).Caption = "Descripción"
       .Columns(3).Width = 700
       .Columns(3).Caption = "Cant."
       .Columns(3).Alignment = dbgRight
       .Columns(3).NumberFormat = "####0.00"
       .Columns(3).Width = 600
       .Columns(4).Caption = "P.U."
       .Columns(4).Alignment = dbgRight
       .Columns(4).NumberFormat = "#####0.00"
       .Columns(4).Width = 1000
       .Columns(5).Caption = "Total"
       .Columns(5).Alignment = dbgRight
       .Columns(5).NumberFormat = "######0.00"
       .Columns(5).Width = 1000
       .Columns(6).Visible = False
       .Columns(7).Visible = False
       .Columns(8).Visible = False
       .Columns(9).Visible = False
       .Columns(10).Visible = False
       .Columns(11).Visible = False
       .Columns(12).Visible = False
       .Columns(13).Visible = False
       .Columns(14).Visible = False
       .Columns(15).Visible = False
       .Columns(16).Visible = False
       .Columns(17).Visible = False
       .Columns(18).Visible = False
       .Columns(19).Visible = False
       .Columns(20).Visible = False
       .Columns(21).Visible = False
       .Columns(22).Visible = False
       .Columns(23).Visible = False
       .Columns(24).Visible = False
       .Columns(25).Visible = False
       .Columns(26).Visible = False
       .Columns(27).Visible = False
 End With
End Sub

Private Sub Actualiza()
  Dim cRsl As ClsLectura, cRle As ClsPrograma
  
  Set cRsl = New ClsLectura
  Set cRle = New ClsPrograma
'
  Label2 = Format(nMesa, "000000")
  If rAux.RecordCount <> 0 Then
     Label10.Caption = rAux![a_hora]
     If Not IsNull(rAux![a_mozo]) Then
        Combo1.text = cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & rAux![a_mozo])
     End If
     Label8.Caption = Format(cRle.DameTotalMesa(nMesa), "######0.00")
  End If
End Sub


Private Sub CargarCombo()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Dim nCons As Long, nCondPago As Byte

  cRsl.CargaCombo Combo1, "Vendedores", "Vendedor", "Descripcion", ""
  cRsl.CargaCombo Combo2, "Descuentos", "Descuento", "Descripcion", ""

  
  CmbCond.AddItem "Contado"
  CmbCond.ItemData(CmbCond.NewIndex) = 1
  CmbCond.AddItem "Cuenta Corriente"
  CmbCond.ItemData(CmbCond.NewIndex) = 0
  Combo2.text = "NINGUNO"
 
  cRsl.CargaCombo CmbCompro, "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
  cRsl.CargaCombo CmbForma, "CondVenta", "CondVta", "Descripcion", "Tipo=1 or Tipo=2"
 
 
  nCons = cRsl.TraerValorDeUnCampo("Registros", "ClienteConsFinal", "Suc=" & nSucursal & " and Usuario=" & nUsuario)
  sConsFinal = cRsl.TraerValorDeUnCampo("Clientes", "Cliente", "Id=" & nCons)
  sNomFinal = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & sConsFinal & "'")
  TxtCliente.text = sConsFinal
  LblCliente.Caption = sNomFinal
 
  nFormaPago = cRsl.TraerValorDeUnCampo("Registros", "Cobro", "Suc=" & nSucursal & " and Usuario=" & nUsuario)
  nListaVenta = cRsl.RegPorDefecto("ListaPrecio")
 
  nCondPago = cRsl.RegPorDefecto("CPago")
 
  CmbCond.text = "Cuenta Corriente"
 
  If nCondPago = 1 Then
     CmbCond.text = "Contado"
  End If
 
  sCondPago = CmbCond.text
 
  CmbForma.text = cRsl.TraerValorDeUnCampo("CondVenta", "Descripcion", "CondVta=" & nFormaPago)
  
  
  Set cRsl = Nothing
End Sub
Private Sub Habilitar(bFrame As Boolean, bCodi As Boolean, bCanti As Boolean, bObser As Boolean, bCons As Boolean, _
   bLcod As Boolean, bLDes As Boolean, bLCan As Boolean, bLobser As Boolean)
  Frame2.Enabled = bFrame
  Text1.Enabled = bCodi
  Text2.Enabled = bCanti
  Text3.Enabled = bObser
'  Command2.Enabled = bCons
  Label4.Enabled = bLcod
  Label5.Enabled = bLDes
  Label6.Enabled = bLCan
  Label11.Enabled = bLobser
End Sub

Private Sub Mostrar(sCod As String, sDes As String, sCant As String, sObser As String)
  Text1.text = sCod
  Label5.Caption = sDes
  Text2.text = sCant
  Text3.text = sObser
End Sub

Private Sub Botones(bNueva As Boolean, bBorrar As Boolean, bModi As Boolean, bSalir As Boolean, bCerrar As Boolean, bImpre As Boolean)
  Command1(0).Enabled = bNueva
  Command1(1).Enabled = bBorrar
  Command1(3).Enabled = bSalir
  Command1(4).Enabled = bCerrar
  Command1(5).Enabled = IIf(bDelivery = True, False, bCerrar)
  Command1(6).Enabled = bCerrar
End Sub

Private Sub Nuevo()
 Dim cRle As ClsPrograma
 
 Set cRle = New ClsPrograma
 If sw = True Then
    sw = False
    Botones True, False, False, True, False, False
    Command1(0).Picture = LoadResPicture("Grabar", 0)
    Command1(0).Caption = "&Grabar"
    Habilitar True, True, True, True, True, True, True, True, True
    Mostrar "", "", 1, ""
    Text2.SetFocus
 Else
    sw = True
    Command1(0).Picture = LoadResPicture("Nuevo", 0)
    Command1(0).Caption = "&Nuevo"
    Botones True, False, True, True, True, True
    Habilitar False, False, False, False, False, False, False, False, False
    sw1 = False
    Grabar
    Label8.Caption = Format(cRle.DameTotalMesa(nMesa), "######0.00")
    sw1 = True
 End If
 
 Set cRle = Nothing
End Sub

Private Sub salir()
  If sw = False Then
    sw = True
    Command1(0).Picture = LoadResPicture("Nuevo", 0)
    Command1(0).Caption = "Nuevo"
    Botones True, False, True, True, True, True
    Habilitar False, False, False, False, False, False, False, False, False
  Else
     Unload Me
  End If
End Sub
Private Sub Borrar()
  Dim cRls As ClsLectura, cRle As ClsEscritura, cRsp As ClsPrograma
  
  Set cRls = New ClsLectura
  Set cRle = New ClsEscritura
  Set cRsp = New ClsPrograma
  
  If rAux![a_cant] <> 1 Then
     bBor = True
     Frame2.Enabled = True
     Botones False, False, False, False, False, False
     Text1.text = rAux![a_codart]
     Label5.Caption = rAux![A_DESCRI]
     Text2.text = rAux![a_cant]
     nCant = Text2.text
     Text2.Enabled = True
     Text2.SetFocus
  Else
     If MsgBox("Borra el Plato Elegido ?", 20, "Mensaje de Confitería") = 6 Then
     '   ImpreAnular
        cRle.Actualizar "Auxiliar_B", , nReg
        rAux.Delete
        rAux.MovePrevious
        If rAux.BOF Then
           rAux.MoveFirst
        End If
        Command1(1).Enabled = False
     End If
  End If
  
  Label8.Caption = Format(cRsp.DameTotalMesa(nMesa), "######0.00")
  Set cRls = Nothing
End Sub


Private Sub Grabar()
  Dim cRsl As ClsLectura, cRle As ClsEscritura, sCondi As String
    
  Set cRsl = New ClsLectura
  Set cRle = New ClsEscritura
      
  rAux.AddNew
  rAux![a_codart] = Text1.text
  rAux![A_DESCRI] = Label5.Caption
  rAux![a_cant] = Text2.text
  rAux![a_Mesa] = Label2.Caption
  rAux![a_mozo] = Combo1.ItemData(Combo1.ListIndex)
  rAux![a_impcom] = "N"
  rAux![a_hora] = Time
  rAux![a_fecha] = Date
  rAux![A_MENSAJE] = IIf(bDelivery = False, "ABIERTA", "DELIVERY")
  rAux![a_lista] = nLista
  rAux![a_precio] = cRsl.TraerValorDeUnCampo("Precios", "Precio", "Producto='" & rAux![a_codart] & "' AND Lista=" & nLista)
  rAux![a_puerto] = "LPT1"
  rAux![a_puerto2] = "LPT1"
  rAux![a_puerto3] = "LPT1"
  rAux![a_total] = rAux![a_cant] * rAux![a_precio]
  rAux![a_Observ] = Text3.text
  
  rAux.Update

  cRle.Actualizar "Auxiliar_A", rAux
  
  sCondi = "ABIERTA"
  If bDelivery = True Then
     sCondi = "DELIVERY"
  End If
  
  Set rAux = cRsl.TraerRsCondi("Auxiliar", "Id", "a_mesa=" & nMesa & " AND a_mensaje='" & sCondi & "'")
  CabGrid
  
  Set cRsl = Nothing
End Sub

Private Sub CalcularDescuentoGrilla(pDes As Double)
  Dim nDes As Double, nPrecio As Double, cRsl As ClsLectura, nTotal As Double
  
  
  Set cRsl = New ClsLectura
  
  nDes = pDes
  nTotal = 0
  
  If Not rAux Is Nothing Then
     If rAux.RecordCount <> 0 Then
        rAux.MoveFirst
        Do While Not rAux.EOF
           nPrecio = cRsl.TraerValorDeUnCampo("Precios", "Precio", "Producto='" & rAux![a_codart] & "' AND Lista=" & nLista)
        
           rAux![a_precio] = nPrecio - (nPrecio * nDes / 100)
           rAux![a_total] = rAux![a_cant] * rAux![a_precio]
           nTotal = nTotal + rAux![a_total]
           rAux.Update
           rAux.MoveNext
        Loop
        
        Label8.Caption = Format(nTotal, "#0.00")
        
     End If
  End If
  
End Sub

Private Sub TodoBien()
  Command1(0).Enabled = False
  If Len(Text1.text) <> 0 And Len(Text2.text) <> 0 Then
     Command1(0).Enabled = True
  End If
End Sub

Private Sub Text1_Change()
  TodoBien
End Sub

Private Sub Text1_GotFocus()
  Command1(0).Enabled = False
End Sub
Private Sub Text1_LostFocus()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  If Text1.text <> "" Then
     Label5.Caption = cRsl.TraerValorDeUnCampo("Productos", "Descripcion", "Producto='" & Text1.text & "'")
     If Label5.Caption = "0" Then
        Label5.Caption = ""
        Text1.text = ""
        Text1.SetFocus
     Else
        Text3.SetFocus
        TodoBien
     End If
  End If
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If sw1 = True Then
     Botones True, True, True, True, True, True
     If Grid1.Row <> -1 Then
        If Grid1.Row = 0 Then
           Botones True, False, True, True, True, True
        End If
        rAux.Bookmark = Grid1.Bookmark
        nReg = rAux![Id]
        Text1.text = rAux![a_codart]
        Label5.Caption = rAux![A_DESCRI]
        Text2.text = rAux![a_cant]
        Text3.text = "" & rAux![a_Observ]
     End If
  End If
End Sub

Private Sub Text2_GotFocus()
  Text2.SelStart = 0
  Text2.SelLength = Len(Text2.text)
End Sub

Private Sub Text2_LostFocus()
  Dim cRls As ClsLectura, cRle As ClsEscritura, cRlp As ClsPrograma
  
  Set cRls = New ClsLectura
  Set cRle = New ClsEscritura
  Set cRlp = New ClsPrograma
  
  If Text2.text <> "" Then
     If Not IsNumeric(Text2.text) Then
        Text2.SetFocus
     Else
        If bBor = True Then
           If Text2.text <= nCant Then
              If MsgBox("Borra el Plato Elegido ?", 20, "Mensaje de Confitería") = 6 Then
                 If Text2.text > 0 Then
                    rAux![A_MENSAJE] = "ANULADO"
                    rAux![a_cant] = Text2.text
                    rAux![a_total] = rAux![a_precio] * rAux![a_cant]
                    rAux.Update
                 '    ImpreAnular
                    If nCant > Text2.text Then
                       rAux![A_MENSAJE] = "ABIERTA"
                       rAux![a_cant] = nCant - Text2.text
                       rAux.Update
                       Text2.text = rAux![a_cant]
                       Dim Rx As Recordset
                       Set Rx = New Recordset
                       
                       Rx.Fields.Append "Reg", adInteger
                       Rx.Fields.Append "Cantidad", adSingle

                       Rx.Open
                       Rx.AddNew
                       Rx!Reg = nReg
                       Rx!Cantidad = Text2.text
                       Rx.Update
                       
                       cRle.Actualizar "Auxiliar_M", Rx
                    Else
                       cRle.Actualizar "Auxiliar_B", , nReg
                       rAux.Delete
                       rAux.MovePrevious
                       If rAux.BOF Then
                          rAux.MoveFirst
                       End If
                       Text1.text = ""
                       Text2.text = ""
                       Text3.text = ""
                       Label5.Caption = ""
                    End If
                    bBor = False
                    Frame2.Enabled = False
                    Botones True, True, True, True, True, True
                    Habilitar False, False, False, False, False, False, False, False, False
                 Else
                    Text2.SetFocus
                 End If
                 Command1(1).Enabled = False
              '   Command1(2).Enabled = False
                 Set rAux = cRls.TraerRsCondi("Auxiliar", "Id", "a_mesa=" & nMesa)
                 CabGrid
                 Label8.Caption = Format(cRlp.DameTotalMesa(nMesa), "######0.00")
                 Command1(0).SetFocus
              Else
                 bBor = False
                 Frame2.Enabled = False
                 Botones True, True, True, True, True, True
                 Habilitar False, False, False, False, False, False, False, False, False
              End If
              bBor = False
           Else
               Text2.text = nCant
               Text2.SetFocus
           End If
        End If
     End If
  End If
  
  Set cRle = Nothing
  Set cRls = Nothing
End Sub

Private Sub Presupuesto()
  Dim cImp As ClsPrograma, rMesa As New ADODB.Recordset, cTotal As Currency, i As Single, cRsl As ClsLectura
  Dim nCanti As Single, nUnitario As Currency, nTotal As Currency, nCompro As Integer
  
  Set cImp = New ClsPrograma
  Set cRsl = New ClsLectura
  
  Set rMesa = cImp.DameMesaImp(nMesa)
  
  Dim P As Printer, sImpre As String
        
  nCompro = cRsl.RegPorDefecto("VentaOtraCondicion")
  sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & nCompro & " AND Usuario=" & nUsuario)
  
  If sImpre = "Ninguna" Then
      MsgBox "Debe Definir Un Modelo de Impresora al Comprobante...", vbCritical, "Atención"
      Exit Sub
  End If
  For Each P In Printers
     If P.DeviceName = sImpre Then
        Set Printer = P
 '   Printer.FontName = "Courier New"
        Printer.ScaleMode = 7
        Printer.FontBold = cRsl.TraerValorDeUnCampo("Impresion", "Negrita", "Comprobante=" & nCompro & " AND Usuario=" & nUsuario)
        Printer.FontSize = cRsl.TraerValorDeUnCampo("Impresion", "Tamaño", "Comprobante=" & nCompro & " AND Usuario=" & nUsuario)
        Exit For
     End If
  Next
  
  Printer.FontName = "Courier New"
  Printer.ScaleMode = 7
  Printer.FontSize = 14
 
  Printer.CurrentY = 2
  Printer.CurrentX = 1
  Printer.FontBold = True
  Printer.Print sNomFan
  Printer.FontBold = False
  Printer.FontSize = 12
  Printer.CurrentY = 3.3
  Printer.CurrentX = 1
  Printer.Print "                 PRESUPUESTO SIN VALOR COMERCIAL " & Date
  Printer.CurrentY = 4
  Printer.CurrentX = 1
  If bDelivery = False Then
     Printer.Print "Mesa :" & Label2.Caption
  Else
      Printer.Print "Cliente :" & dNombre & " - " & dDireccion
       Printer.CurrentY = 4.5
       Printer.CurrentX = 1
       Printer.Print "Hora Pedido :" & dHora
  End If
  Printer.CurrentY = 5
  Printer.CurrentX = 1
  Printer.Print "_________________________________________________________________________________________________________"
  Printer.CurrentY = 5.5
  Printer.CurrentX = 1
  Printer.Print "Cantidad   Codigo          D e s c r i p c i ó n           Precio     Total"
  Printer.CurrentY = 5.7
  Printer.CurrentX = 1
  Printer.Print "_________________________________________________________________________________________________________"

  cTotal = 0
  i = 6.2
  rMesa.MoveFirst
  Do While Not rMesa.EOF
         ' Codigo Articulo
         
     nCanti = Format(rMesa!Cantidad, "#0.00")
     nUnitario = Format(rMesa!a_precio, "#0.00")
     nTotal = Format(rMesa!a_precio * rMesa!Cantidad, "#0.00")
                     
     Printer.CurrentY = i
     Printer.CurrentX = 2 - Printer.TextWidth(Format(nCanti, "#0.00"))
     Printer.Print Format(nCanti, "#0.00")
     
     Printer.CurrentY = i
     Printer.CurrentX = 3.5
     Printer.Print rMesa!a_codart

     Printer.CurrentY = i
     Printer.CurrentX = 7
     Printer.Print rMesa!A_DESCRI
                 
     Printer.CurrentY = i
     Printer.CurrentX = 17.5 - Printer.TextWidth(Format(nUnitario, "#0.00"))
     Printer.Print Format(nUnitario, "#0.00")
     
     Printer.CurrentY = i
     Printer.RightToLeft = True
     Printer.CurrentX = 20 - Printer.TextWidth(Format(nTotal, "#0.00"))
     Printer.Print Format(nTotal, "#0.00")
               
     Dim h As Integer
     If Not rMesa.BOF Then
'        If rMesa!a_mensaje <> "" Then
'           ReDim Preserve dObservacion(UBound(dObservacion) + 1)
'           dObservacion(h) = rMesa!a_Observ
'           h = h + 1
'        End If
    End If
    rMesa.MoveNext
    i = i + 0.4
    cTotal = cTotal + nTotal
  Loop
  i = i + 1
  
  Printer.FontBold = True

  Printer.CurrentY = i
  Printer.CurrentX = 15
  Printer.Print "Total :" & Format(cTotal, "#0.00")

  Printer.FontBold = False

  Printer.FontSize = 10

  i = i + 1
  Printer.CurrentY = i
  Printer.CurrentX = 1
  
  Printer.Print "Observación :" & dObservacion(h)

  Printer.CurrentY = i
  Printer.CurrentX = 1

 If bDelivery = True Then
    For h = 0 To UBound(dObservacion)
       Printer.Print "             " & dObservacion(h)
       i = i + 0.3
       Printer.CurrentY = i
       Printer.CurrentX = 1
    Next
  Else
     Printer.Print " "
 End If
'  i = i + 1
'
'  Printer.CurrentY = i
'  Printer.CurrentX = 1
'  Printer.Print "SOLICITE LA FACTURA DE ACUERDO A SU CONDICION FRENTE AL I.V.A."
'
'  i = i + 0.5
'
'  Printer.CurrentY = i
'  Printer.CurrentX = 1
'  Printer.Print "Apellido y Nombre:_____________________________________C.U.I.T Nº____________"
'
'  i = i + 0.5
'
'  Printer.CurrentY = i
'  Printer.CurrentX = 1
'  Printer.Print "RESP.INSCR._______CONSUMIDOR FINAL____MONOTRUBUTISTA___RESP.NO INSCR.____"

  
   Printer.EndDoc


End Sub

Private Sub Comanda()
   Dim nCan As Currency, nUni As Currency, nTot As Currency, Rsd As Recordset, cImp As ClsPrograma, cRsl As ClsLectura
   Dim i As Integer, X As Integer, Y As Single, nCompro As Integer, rMesa As Recordset

   Set cImp = New ClsPrograma
   Set cRsl = New ClsLectura
   
  Set rMesa = cImp.DameMesaImp(nMesa)

   Dim P As Printer, sImpre As String

   nCompro = cRsl.RegPorDefecto("VentaOtraCondicion")

   sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & nCompro & " AND Usuario=" & nUsuario)

   If sImpre <> "Ninguna" Then

      For Each P In Printers
          If P.DeviceName = sImpre Then
             Set Printer = P
             Printer.FontName = "Courier New"
             Printer.ScaleMode = 7
             Printer.FontSize = 9
             Exit For
          End If
      Next

     Printer.FontSize = 11
 
     Printer.CurrentY = 0
     Printer.CurrentX = 0.5
     Printer.FontBold = True
     Printer.Print sNomFan
   '  Printer.FontBold = False
     Printer.CurrentY = 1
     Printer.CurrentX = 0.5
     Printer.Print "PRESUPUESTO " & Label2.Caption
     
     Printer.CurrentY = 2
     Printer.CurrentX = 0.5
     
     Printer.FontSize = 9
     If bDelivery = False Then
        Printer.Print "Mesa :" & Label2.Caption
     Else
        Printer.Print "Cliente :" & dNombre
        Printer.CurrentY = 3.4
        Printer.CurrentX = 0.5
        Printer.Print "Domicilio :" & dDireccion
        Printer.CurrentY = 3.8
        Printer.CurrentX = 0.5
        Printer.Print "Hora Pedido :" & dHora
        
      End If
     
     Y = 3
     X = 0.2


     Printer.CurrentY = Y
     Printer.CurrentX = X
     Printer.Print " "

      rMesa.MoveFirst

       Printer.FontSize = 9

      Do While Not rMesa.EOF
         Printer.CurrentY = Y
         Printer.CurrentX = X

          nCan = Format(rMesa!Cantidad, "#0.00")
          nUni = Format(rMesa!a_precio, "#0.00")
          nTot = rMesa!a_precio * rMesa!Cantidad
          nTot = Format(nTot, "#0.00")
          Printer.Print Space(5 - Len(Format(rMesa!Cantidad, "0.00"))) & Format(rMesa!Cantidad, "0.00") & " " & Left(rMesa!A_DESCRI, 20) & Space(20 - Len(Left(rMesa!A_DESCRI, 20)))
          Y = Y + 0.5
          
          Printer.CurrentX = 1.5
          Printer.Print Space(7 - Len(Format(nUni, "#0.00"))) & " " & Format(nUni, "0.00") & Space(7 - Len(Format(nTot, "#0.00"))) & " " & Format(nTot, "0.00")
          X = 0.2
         Dim h As Integer
'         If Not Rsd.BOF Then
'            If rAux!a_Observ <> "" Then
'               ReDim Preserve dObservacion(UBound(dObservacion) + 1)
'               dObservacion(h) = rMesa!a_Observ
'               h = h + 1
'            End If
'          End If
          Y = Y + 0.2
       '   X = X + 1
          rMesa.MoveNext
       Loop

        Y = Y + 0.5

        Printer.FontSize = 9

        Printer.CurrentY = Y
        Printer.CurrentX = X
    '    Printer.FontBold = True

        Printer.Print "Importe Total :" & Format(Label8.Caption, "#0.00")

        Y = Y + 1
   '     Printer.FontBold = False
        Printer.CurrentY = Y
        Printer.CurrentX = 0.5
  
        Printer.Print dObservacion(h)

        Printer.CurrentY = Y
        Printer.CurrentX = 0.5

        If bDelivery = True Then
           For h = 0 To UBound(dObservacion)
               Printer.Print dObservacion(h)
               Y = Y + 0.3
               Printer.CurrentY = Y
               Printer.CurrentX = 0.5
           Next
         Else
           Printer.Print " "
        End If
        
        Y = Y + 3

        Printer.CurrentY = Y
        Printer.CurrentX = X
        Printer.Print " "

'        Y = Y + 1
'
'        Printer.CurrentY = Y
'        Printer.CurrentX = X
'        Printer.Print " "
'
'        Y = Y + 1
'        Printer.CurrentY = Y
'        Printer.CurrentX = X
'        Printer.Print " "
        
'        Printer.CurrentY = Y
'        Printer.CurrentX = X
'        Printer.Print "----------------------------------------"

        Printer.EndDoc
     End If
End Sub

Private Sub CierreRapido()
'  Venta.Show 1
'  If bCerrar = True Then
'     Unload Me
'  End If
  
End Sub

Private Function CalcularDescuento(pTotal As Double) As Double
  Dim nDesc As Double
  nDesc = 0
  If bDescuento = True Then
     If TxtCliente.text <> sConsFinal Then
        If CmbForma.text = sFormaPago Then
           If CmbCond.text = "Contado" Then
              pTotal = pTotal * dPorcDescuento / 100
              nDesc = pTotal
           End If
        End If
     End If
  End If
  
  CalcularDescuento = nDesc
  
End Function

Private Sub TxtCliente_LostFocus()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Dim dTotal As Double
  
  dTotal = dTotalAux

  If bFoco = False Then
    If TxtPago.Visible = True And TxtPago.Enabled = True And CmdVuelto(0).Enabled = True Then
       If cRsl.TraerCantidad("Clientes", "Cliente", "Cliente='" & TxtCliente.text & "'") = 0 Then
          BtnBuscar.Enabled = True
          CmbCond.Enabled = True
          BtnBuscar.SetFocus
       Else
          LblCliente.Caption = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & TxtCliente.text & "'")
          CmbCompro.text = cRsl.TraerComproCliente(TxtCliente.text, Frame3.Caption)
          If TxtCliente.text = "." Then
             CmbCond.text = "Contado"
          Else
             CmbCond.text = sCondPago
             CmbCond.Enabled = True
          End If
          dTotal = dTotalAux - CalcularDescuento(dTotal)
          Label8.Caption = Format(dTotal, "#0.00")
          LblCompra.Caption = Label8.Caption
          If TxtPago.Visible = True Then
             TxtPago.SetFocus
          End If
       End If
    End If
  End If
End Sub

Function ObtenerDecimales(ByVal n As Double) As String
    Dim pos As Long
    n = Round(n, 3)
    pos = InStr(n, ".") + InStr(n, ",")
    If pos <> 0 Then ObtenerDecimales = Mid$(n, pos + 1)
End Function

Private Sub GrabarCae(pMovimiento As Long, pNumero As Long)
   Dim cRle As ClsEscritura
   Set cRle = New ClsEscritura
        
   Dim rCae As ADODB.Recordset
   Set rCae = New ADODB.Recordset
        
   rCae.Fields.Append "Movimiento", adInteger
   rCae.Fields.Append "Fecha", adVarChar, 10
   rCae.Fields.Append "CAE", adVarChar, 50
   rCae.Fields.Append "Documento", adVarChar, 20
   rCae.Fields.Append "Comprobante", adInteger
   rCae.Fields.Append "Numero", adInteger
   rCae.Open
                                            
   rCae.AddNew
   rCae!Movimiento = pMovimiento
   rCae!fecha = Date
   rCae!CAE = "0000"
   rCae!Documento = "0000000000"
   rCae!Comprobante = "0000000"
   rCae!Numero = pNumero
   rCae.Update
   
   cRle.Actualizar "AfipCaes_A", rCae
   
End Sub
