VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Principal 
   BackColor       =   &H00808080&
   Caption         =   "Ds Gestión Comercial V.1.00"
   ClientHeight    =   6945
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10800
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2070
      Top             =   2580
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2372
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2486
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":259A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":26AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":27C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":28D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3178
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":35CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3A1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6690
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "2001- G&A  -Todos los Derechos Reservados-"
            TextSave        =   "2001- G&A  -Todos los Derechos Reservados-"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clientes"
            Object.ToolTipText     =   "Maestro de Clientes"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Productos"
            Object.ToolTipText     =   "Maestros de Productos"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Proveedores"
            Object.ToolTipText     =   "Maestro de Proveedores"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ventas"
            Object.ToolTipText     =   "Comprobante de Venta"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cobranzas"
            Object.ToolTipText     =   "Cobranzas"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Precios"
            Object.ToolTipText     =   "Consultas de Precios"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PreciosCom"
            Object.ToolTipText     =   "Precios de Compras"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grillaproduc"
            Object.ToolTipText     =   "Grilla de productos"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivos"
      Begin VB.Menu Tablas 
         Caption         =   "Tablas"
         Begin VB.Menu Banco 
            Caption         =   "Banco"
         End
         Begin VB.Menu MenuCaja 
            Caption         =   "Caja"
         End
         Begin VB.Menu MenuCtaBanc 
            Caption         =   "Cuenta Bancaria"
         End
         Begin VB.Menu Comprobante 
            Caption         =   "Comprobantes"
         End
         Begin VB.Menu CodVenta 
            Caption         =   "Forma de Pago"
         End
         Begin VB.Menu Descuento 
            Caption         =   "Descuentos"
         End
         Begin VB.Menu Depositos 
            Caption         =   "Depósitos"
         End
         Begin VB.Menu Familia 
            Caption         =   "Familias"
         End
         Begin VB.Menu FamiliaSub 
            Caption         =   "Familias Sub"
         End
         Begin VB.Menu Impuestos 
            Caption         =   "Impuestos"
         End
         Begin VB.Menu ListaPrecio 
            Caption         =   "Lista de Precios"
         End
         Begin VB.Menu Marca 
            Caption         =   "Marcas"
         End
         Begin VB.Menu Moneda 
            Caption         =   "Monedas"
         End
         Begin VB.Menu MenuPlan 
            Caption         =   "Plan de Cuenta"
         End
         Begin VB.Menu TblRecargo 
            Caption         =   "Recargos"
         End
         Begin VB.Menu MenuTrans 
            Caption         =   "Transporte"
         End
         Begin VB.Menu UniMed 
            Caption         =   "Unidad de Medida"
         End
         Begin VB.Menu Vendedor 
            Caption         =   "Vendedores"
         End
         Begin VB.Menu Zona 
            Caption         =   "Zonas"
         End
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Maestros 
      Caption         =   "&Maestros"
      Begin VB.Menu MClientes 
         Caption         =   "Clientes"
         Begin VB.Menu Clientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu SepCli 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu CliInformes 
            Caption         =   "Listados de Clientes"
         End
      End
      Begin VB.Menu MProductos 
         Caption         =   "Productos"
         Begin VB.Menu Productos 
            Caption         =   "Productos"
         End
         Begin VB.Menu MenProdUnidades 
            Caption         =   "Productos Unidades"
         End
         Begin VB.Menu SepProd1 
            Caption         =   "-"
         End
         Begin VB.Menu ListProduc 
            Caption         =   "Listado de Productos"
         End
         Begin VB.Menu SepProd 
            Caption         =   "-"
         End
         Begin VB.Menu MenuEtiMovi 
            Caption         =   "Etiquetas por Movimiento"
         End
         Begin VB.Menu MenuEtiquetas 
            Caption         =   "Etiquetas"
         End
         Begin VB.Menu SepProd2 
            Caption         =   "-"
         End
         Begin VB.Menu MenuCodBarra 
            Caption         =   "Codigos de Barra"
         End
      End
      Begin VB.Menu MProveedores 
         Caption         =   "Proveedores"
         Begin VB.Menu Proveedores 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu ListProveedores 
            Caption         =   "Listado de Proveedores"
         End
      End
   End
   Begin VB.Menu MenuStock 
      Caption         =   "&Stock"
      Begin VB.Menu StockComprobante 
         Caption         =   "Comprobantes              "
         Shortcut        =   ^C
      End
      Begin VB.Menu MenuConciliar 
         Caption         =   "Conciliar Transferencias"
      End
      Begin VB.Menu StockSep1 
         Caption         =   "-"
      End
      Begin VB.Menu StockInformes 
         Caption         =   "Informes"
         Begin VB.Menu RemitosPendientes 
            Caption         =   "Remitos Pendientes"
         End
         Begin VB.Menu MStockSep1 
            Caption         =   "-"
         End
         Begin VB.Menu MenuStockSaldos 
            Caption         =   "Saldos de Stock"
         End
         Begin VB.Menu StockFaltantesStock 
            Caption         =   "Faltantes de Stock"
         End
         Begin VB.Menu Kardex 
            Caption         =   "Movimiento de Stock"
         End
         Begin VB.Menu MStockSep2 
            Caption         =   "-"
         End
         Begin VB.Menu MStockCostos 
            Caption         =   "Costos"
         End
      End
   End
   Begin VB.Menu MenuCompras 
      Caption         =   "&Compras"
      Begin VB.Menu ComprasFacturacion 
         Caption         =   "Facturación"
      End
      Begin VB.Menu CompRemitos 
         Caption         =   "Remitos"
      End
      Begin VB.Menu ComprasPagos 
         Caption         =   "Pagos"
      End
      Begin VB.Menu MenuAdeProv 
         Caption         =   "Adelantos a Proveedores"
      End
      Begin VB.Menu ComprasAnular 
         Caption         =   "Ver/Anular Pagos"
      End
      Begin VB.Menu ComprasSep 
         Caption         =   "-"
      End
      Begin VB.Menu MenuComprasPrecios 
         Caption         =   "Precios"
      End
      Begin VB.Menu ComprasPrecios 
         Caption         =   "Consultar Precios"
      End
      Begin VB.Menu ComprasSep2 
         Caption         =   "-"
      End
      Begin VB.Menu ComprasInformes 
         Caption         =   "Informes"
         Begin VB.Menu ComprasResCta 
            Caption         =   "Resumen de Cuenta"
         End
         Begin VB.Menu MenuDetCompras 
            Caption         =   "Detalle de Compras"
         End
         Begin VB.Menu ComprasVencimientos 
            Caption         =   "Vencimientos"
         End
         Begin VB.Menu SepCompras 
            Caption         =   "-"
         End
         Begin VB.Menu MenuIvaCompras 
            Caption         =   "Libro de Iva"
         End
         Begin VB.Menu MCitiCompras 
            Caption         =   "Exportar al Citi Compras"
         End
         Begin VB.Menu SepIvaCompras 
            Caption         =   "-"
         End
         Begin VB.Menu InfRetIb 
            Caption         =   "Retención de Ingresos Brutos"
         End
         Begin VB.Menu InfRetGanancia 
            Caption         =   "Retención de Ganancia"
         End
         Begin VB.Menu SepChe 
            Caption         =   "-"
         End
         Begin VB.Menu MnuChequeCart 
            Caption         =   "Cheques en Cartera"
         End
      End
   End
   Begin VB.Menu MenuVentas 
      Caption         =   "&Ventas"
      Begin VB.Menu VentaComprobante 
         Caption         =   "Facturación      "
         Shortcut        =   ^A
      End
      Begin VB.Menu VentaRemitos 
         Caption         =   "Remitos"
      End
      Begin VB.Menu VentaFacRemPend 
         Caption         =   "Facturar Remitos Pendientes"
         Visible         =   0   'False
      End
      Begin VB.Menu MVentaCobranzas 
         Caption         =   "Cobranzas"
         Shortcut        =   ^B
      End
      Begin VB.Menu MenuAdeClientes 
         Caption         =   "Adelantos de Clientes"
      End
      Begin VB.Menu MeVentaAnularCobranza 
         Caption         =   "Ver/Anular/Desimputar Cobranzas"
      End
      Begin VB.Menu VentaSep7 
         Caption         =   "-"
      End
      Begin VB.Menu VentaPrecios 
         Caption         =   "Precios"
      End
      Begin VB.Menu MenuPreciosGrid 
         Caption         =   "Precios Grilla"
      End
      Begin VB.Menu ConsultoPrecios 
         Caption         =   "Consultar Precios x Articulo"
      End
      Begin VB.Menu VentaSep5 
         Caption         =   "-"
      End
      Begin VB.Menu VentasInformes 
         Caption         =   "Informes"
         Begin VB.Menu MenuResCuenta 
            Caption         =   "Resumen de Cuenta"
         End
         Begin VB.Menu MenuResCueDet 
            Caption         =   "Resumen de Cuenta con Detalles"
         End
         Begin VB.Menu MenResCtaFlia 
            Caption         =   "Resumen de Cuenta por Familias"
         End
         Begin VB.Menu MenuVencimientos 
            Caption         =   "Vencimientos"
         End
         Begin VB.Menu MVenProy 
            Caption         =   "Proyección de Saldos"
         End
         Begin VB.Menu MenuDetalles 
            Caption         =   "Detalles Ventas"
         End
         Begin VB.Menu MenuDetCob 
            Caption         =   "Detalles de Cobranza"
         End
         Begin VB.Menu VentaSep1 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MenuInfVenta 
            Caption         =   "Ventas"
         End
         Begin VB.Menu VentaInfFami 
            Caption         =   "Ventas Por Familias"
         End
         Begin VB.Menu VentasUtilidades 
            Caption         =   "Planilla de Utilidades"
         End
         Begin VB.Menu ControlGastos 
            Caption         =   "Control de Gastos"
         End
         Begin VB.Menu Ranking 
            Caption         =   "Ranking Productos Vendidos"
         End
         Begin VB.Menu Consolidados 
            Caption         =   "Consolidado de Mercaderias"
         End
         Begin VB.Menu MenuReparto 
            Caption         =   "Reparto"
         End
         Begin VB.Menu MenuRepRemi 
            Caption         =   "Reparto por Remito"
         End
         Begin VB.Menu MenuRepEntregas 
            Caption         =   "Reparto Entregas"
         End
         Begin VB.Menu SepVent 
            Caption         =   "-"
         End
         Begin VB.Menu VentaFactLine 
            Caption         =   "Ver Facturas On Line"
         End
         Begin VB.Menu MenuIvaVenta 
            Caption         =   "Libro de Iva "
         End
         Begin VB.Menu MCitiVentas 
            Caption         =   "Exportar al Citi Ventas"
         End
         Begin VB.Menu SepIva 
            Caption         =   "-"
         End
         Begin VB.Menu InfPerIB 
            Caption         =   "Percepción de Ingresos Brutos"
         End
         Begin VB.Menu VentaSep 
            Caption         =   "-"
         End
         Begin VB.Menu VentaCosteos 
            Caption         =   "Costeos"
            Begin VB.Menu VentaCostoVenta 
               Caption         =   "Valorización de Inventario"
            End
            Begin VB.Menu VentaProductosVendidos 
               Caption         =   "Rentabilidad"
            End
         End
         Begin VB.Menu VentaSep2 
            Caption         =   "-"
         End
         Begin VB.Menu VentaListaPrecios 
            Caption         =   "Lista de Precios"
         End
      End
   End
   Begin VB.Menu Caja 
      Caption         =   "&Caja"
      Begin VB.Menu MCajaRap 
         Caption         =   "Caja Rápida"
      End
      Begin VB.Menu MCobSaldos 
         Caption         =   "Cobranza Sobre Saldos"
      End
      Begin VB.Menu SepCaja 
         Caption         =   "-"
      End
      Begin VB.Menu MenuComprobanteCaja 
         Caption         =   "Comprobantes"
      End
      Begin VB.Menu CajaSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuInforCaja 
         Caption         =   "Informes"
         Begin VB.Menu MenuInforMoviCaja 
            Caption         =   "Movimientos"
         End
         Begin VB.Menu MenSalCaja 
            Caption         =   "Saldo de Caja"
         End
         Begin VB.Menu SepCa 
            Caption         =   "-"
         End
         Begin VB.Menu CajaChequeTer 
            Caption         =   "Cheques de Terceros"
         End
         Begin VB.Menu MenuSaldoCtaCteBan 
            Caption         =   "Resumen de Ctas Ctes Bancarias"
         End
         Begin VB.Menu MenuCtaCbles 
            Caption         =   "Resumen de Cuentas Contables"
         End
      End
      Begin VB.Menu CajaSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCierreTurno 
         Caption         =   "Cierre de Turno"
      End
      Begin VB.Menu VerTurnos 
         Caption         =   "Ver Turnos"
      End
      Begin VB.Menu SepCierre 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCajaSaldo 
         Caption         =   "Comprobantes de Ventas"
      End
      Begin VB.Menu CierreManual 
         Caption         =   "Cierre de Caja Manual"
      End
   End
   Begin VB.Menu MenuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu HerrUserPer 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu MHerrPer 
         Caption         =   "Permisos"
      End
      Begin VB.Menu HerrSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MHerraRegDef 
         Caption         =   "Registros por Defecto"
      End
      Begin VB.Menu MHerrConf 
         Caption         =   "Configuración de Comprobantes"
      End
      Begin VB.Menu HerrImporProduc 
         Caption         =   "Importar Productos"
      End
      Begin VB.Menu SepEmp 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEmpresa 
         Caption         =   "Datos Empresa"
      End
      Begin VB.Menu HerrConfigCorreo 
         Caption         =   "Datos Correo Electronico"
      End
   End
   Begin VB.Menu MenuEmpresas 
      Caption         =   "Empresas"
      Begin VB.Menu SubMenu 
         Caption         =   "SubMenu"
         Index           =   1
      End
   End
   Begin VB.Menu Ventana 
      Caption         =   "V&entana"
      WindowList      =   -1  'True
      Begin VB.Menu Horizontal 
         Caption         =   "Mosaico Horizontal"
      End
      Begin VB.Menu Vertical 
         Caption         =   "Mosaico Vertical"
      End
      Begin VB.Menu Cascada 
         Caption         =   "Cascada"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sVersion As String

Private Sub Banco_Click()
  TblBancos.Show
End Sub

Private Sub CajaChequeTer_Click()
  CajaChequesTer.Show
End Sub

Private Sub CierreManual_Click()
   CajaCerrarZ.Show
End Sub

Private Sub Clientes_Click()
  MtrClientes.Show
End Sub

Private Sub CliInformes_Click()
  VentaInfClientes.Show
End Sub

Private Sub CodVenta_Click()
  TblCondVenta.Show
End Sub

Private Sub ComprasAnular_Click()
  ComprasAnularPagos.Show
End Sub

Private Sub ComprasFacturacion_Click()
  ComprasComprobantes.Show
End Sub

Private Sub ComprasNotas_Click()
  ComprasNotaCredito.Show
End Sub
Private Sub ComprasPagos_Click()
  ComprasPago.Show
End Sub

Private Sub ComprasResCta_Click()
  CompraInfResumen.Show
End Sub

Private Sub ComprasVencimientos_Click()
  CompraInfVencimiento.Show
End Sub

Private Sub CompRemitos_Click()
  ComprasRemitos.Show
End Sub

Private Sub Comprobante_Click()
  TblComprobantes.Show
End Sub

Private Sub Consolidados_Click()
   VentaConsolidado.Show
End Sub

Private Sub ConsultoPrecios_Click()
    VentaConsPrecios.Show
End Sub
Private Sub ControlGastos_Click()
  pPlani = 3
  VentaInfFamilias.Show
End Sub
Private Sub Depositos_Click()
  TblDeposito.Show
End Sub
Private Sub Descuento_Click()
  TblDescuentos.Show
End Sub

Private Sub Familia_Click()
  TblFamilias.Show
End Sub

Private Sub HerrIntruc_Click()
  Dim s_script As String
  Dim sArch As String
  
  sArch = App.Path & "\Arch.sql"
  
  Open sArch For Input As #1
  
  s_script = input(LOF(1), #1)
  
  Close #1
  
  Dim sConex As Connection, sCons(0) As String
    
  Set sConex = New Connection
  
  
  sConex.Open sDb
  sConex.BeginTrans
  
  sConex.Execute s_script
    
  sConex.CommitTrans
  sConex.Close
  
  
End Sub

Private Sub FamiliaSub_Click()
  TblFamiliaSub.Show
End Sub

Private Sub HerrConfigCorreo_Click()
  HerraDatosCorreo.Show
End Sub

Private Sub HerrImporProduc_Click()
  HerraImportar.Show
End Sub

Private Sub HerrUserPer_Click()
   HerraUsuarios.Show
End Sub

Private Sub Impuestos_Click()
  TblImpuestos.Show
End Sub
Private Sub Informes_Click()
  VentaInfClientes.Show
End Sub

Private Sub InfPerIB_Click()
  VentasInfPerIb.Show
End Sub

Private Sub InfRetGanancia_Click()
  ComprasInfRetGan.Show
End Sub

Private Sub InfRetIb_Click()
   ComprasInfRetIb.Show
End Sub

Private Sub Kardex_Click()
  StockInfMovimientos.Show
End Sub

Private Sub ListaPrecio_Click()
  TblListaPrecios.Show
End Sub
Private Sub ListProduc_Click()
  StockInfProductos.Show
End Sub

Private Sub Marca_Click()
  TblMarcas.Show
End Sub

Private Sub MCajaRap_Click()
  VentaRapido.Show
End Sub

Private Sub MCitiCompras_Click()
  CitiCompra.Show
End Sub

Private Sub MCitiVentas_Click()
   CitiVenta.Show
End Sub


Private Sub MCobSaldos_Click()
  VentaCobAuto.Show
End Sub

Private Sub MDIForm_Load()

  Dim sVersionSql As ClsSqlVersion
  Set sVersionSql = New ClsSqlVersion

  Dim nVer As Single

  On Error GoTo Errores:
   If App.PrevInstance Then
     MsgBox "La Aplicación ya esta en ejecución ...", vbCritical, "Atención"
     Unload Me
  End If

  Dim cRsl As ClsLectura, i As Byte, sCli As String
  
  sVersion = " - Versión 26.6.2 Turbo"
  
  Set cRsl = New ClsLectura
    
  Call CambiarCR
  CrearDirectorios
  
  sDb = "FILE NAME=" & App.Path & "\Conex.udl"
  
  
  Set db = New ADODB.Connection
  
  nVer = sVersionSql.VerificarVersion()

  db.Open sDb
  
  sVersionSql.ActualizarVersion nVer
  
  db.Open sDb
    
  Dim Rs As New ADODB.Recordset
  Dim sSEnvase As String
  
  Rs.Open "Sistema", db, adOpenKeyset, adLockPessimistic
  pControladora = Rs!Controladora
  bImprimeCtaCte = Rs!ImprimeCtaCte
  nCodeQr = Rs!Qr
  
  sNomFan = Rs!Fantasia
  sDirFan = Rs!Direccion
  sTelFan = Rs!Telefono
  sCiuFan = Rs!Ciudad
  nDigito = Rs!Digitos
  nCodigo = Rs!Cadena
  nRecibo = Rs!Recibo
  nFactura = Rs!IdCheque
  sCuitAfip = Rs!Cuit
  bBuscarActivo = True
  bPagare = Rs!bPagare  ' Busqueda avanzada de Productos si/no por defecto
  bModoSuper = Rs!Pagare 'Modo Precios directo o supermercado
  bRecibo = IIf(Rs!Receta = 1, True, False) ' Modo comanda en los recibos y presupuestos en resto y deliv
  bTransferencia = Rs!bRecibo  ' Permite conciliar las transferencias
  bPermiteStockNega = Rs!bRapido ' Permite Facturar con stock negativo
  nConceptoAfip = Rs!ConceptoAfip
  sWebServiceLogin = Rs!WebServicelogin
  sWebServiceCae = Rs!WebserviceCae
  sCertificado = Rs!certificado
  sClavePrivada = Rs!ClavePrivada
  sArchFartura = Rs!ArchFactura
  sLogoAfip = Rs!LogoAfip
  sAnchoTexto = Rs!AnchoTexto
  bMono = Rs!MONOTRIBUTO
  nSocket = Rs!socket
  CompaniaId = Rs!compania
  CompaniaSuc = Rs!sucursal
  
  nLineasElectronic = Rs!Representante  ' Cantidad de Lineas en Las Facturas Electronicas
  
  nScaner = Rs!Scaner 'si la balanza trabaja con decimales
  
  pClave = IIf(Rs!clave = "0", False, True)
  pClave1 = IIf(Rs!ClaveStock = "0", False, True)
      
  sClave = Rs!clave
  sClave1 = Rs!ClaveStock
  
  nDias = Rs!Dias
  nCadena = Rs!Cadena

  nCantDecimales = Rs!CantDecimales
    
  bAmericana = Rs!FAmericana
  bPercibe = Rs!Percibe
   
  For i = 0 To Rs!CantDecimales
      If i = 0 Then
         nCantDecimales = "0"
      Else
         If i = 1 Then nCantDecimales = nCantDecimales & "."
         nCantDecimales = nCantDecimales & "0"
      End If
  Next i
  
  For i = 0 To Rs!DecimalesCant
      If i = 0 Then
         nDecimalCant = "0"
      Else
         If i = 1 Then nDecimalCant = nDecimalCant & "."
         nDecimalCant = nDecimalCant & "0"
      End If
  Next i
  
  sSEnvase = Rs!Envase
  tEnvase = Rs!TextoEnvase
  fBalanza = Rs!DBalanza
  sVentaCPago = Rs!CondPago

  
  db.Close

  FrmUsuario.Show 1
  
  If FrmUsuario.bCanc = True Then
     Unload Me
  Else
                       
   ' sSEnvase = cRsl.SistemaPorDefecto("Envase")
    
    sEnvase = IIf(Len(sSEnvase) > 1, Left(sSEnvase, Val(Right(sSEnvase, 1)) - 1), "0")
    nEnvase = Len(sEnvase)
   ' tEnvase = cRsl.SistemaPorDefecto("TextoEnvase")
   ' fBalanza = cRsl.SistemaPorDefecto("DBalanza")
                      
    Set rUsu = New ADODB.Recordset
    Set rUsu = cRsl.DamePermisos(FrmUsuario.sUsuario, FrmUsuario.sContra)
    nUsuario = rUsu!Id_Usuario
    sUser = rUsu!Nombre
    StatusBar1.Panels(1).Text = "Summer 2004 - " & Year(Date) & "  Carlos D´Agostino - Usuario :" & sUser
    nPuerto = rUsu!Controladora
    
  '   nScaner = rUsu!Scanner 'Si la balanza trabaja con decimales
    nPcompro = rUsu!Comprobante
    nPcaja = rUsu!Caja
    
    'nUsuario =  cRsl.TraerValorDeUnCampo("Usuarios", "Id_Usuario", "Usuario='" & FrmUsuario.sUsuario & "'")
    Set rRsReg = cRsl.TraerRsCondi("Registros", "Usuario", "Usuario=" & nUsuario & " And Suc = " & nSucursal)
       
    nCondPago = rRsReg!CPago
    nFormaPago = rRsReg!FormaCob
    bAutoNumera = rRsReg!CliProvAut
    bRedondeaLista = rRsReg!Redondea
    bAnulaPres = rRsReg!AnulaPres
    nCaja = rRsReg!Caja '
    bLista = rRsReg!ListaPrecio
    bImprDirecto = rRsReg!ImprimeDirecto
    nVentaFactura = rRsReg!VentaFactura
        
    nMedDetalle = rRsReg!Medida
    nImpDetalle = rRsReg!Pais
    nDepDetalle = rRsReg!Deposito
    nCueDetalle = rRsReg!CuentaRes
    
    nCatUser = rUsu!Categoria
    nControlador = rUsu!Controlador
    bRedondea = rRsReg!NumeraContr
           
    sCli = rRsReg!ClienteConsFinal
    
    sConsFinal = cRsl.TraerValorDeUnCampo("Clientes", "Cliente", "Id=" & sCli)
          
    Permisos
    
    bDescuento = cRsl.TieneDescuento
    ' Limpiar Clientes de Contado
    Menu SubMenu
    If bRedondea = True Then
       Toolbar1.Visible = False
    End If
               
    If cRsl.TraerCantidad("Turnos", "Id", "FechaCierre is null and Caja=" & nCaja) = 0 Then
       db.Open sDb
       db.Execute "INSERT INTO Turnos (Usuario,Fecha,FechaApertura,FechaCierre,Importe,Caja) VALUES ('" & sUser & "','" & Date & "','" & DateTime.Now & "',NULL,0," & nCaja & ")"
       db.Close
    End If
  End If
  
  
  If Not PongActivo() Then
      MsgBox "Problemas de Validación del Sistema. Por favor, comuníquese con el servicio técnico.", _
               vbCritical, "Atención"
        End
  End If
  
  EnviarPresencia
    Timer1.Interval = 60000 ' 1 minuto
    Timer1.Enabled = True
 Exit Sub
Errores:
   MsgBox err.Description, vbCritical, "Atención"
   End
 
End Sub

Private Sub EnviarPresencia()
    On Error Resume Next

    Dim http As Object, cRsl As ClsLectura
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim body As Variant
    Dim pCliente As Integer
    Dim pUsuario As String
    Dim pMaquina As String

    pCliente = 2
    pUsuario = sNomFan             ' usuario actual
    pMaquina = Environ$("COMPUTERNAME")

    body = "{""cliente_id"":" & Chr(34) & pCliente & Chr(34) & _
           ",""usuario"":" & Chr(34) & pUsuario & Chr(34) & _
           ",""maquina"":" & Chr(34) & pMaquina & Chr(34) & "}"
           

    http.SetTimeouts 2000, 2000, 2000, 2000
    
    http.Open "POST", "http://comun.dagsis.net.ar/api/vb6/presence/ping", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body

    ' Si querés loguear:
    ' Debug.Print http.Status, http.responseText
End Sub

Private Function PongActivo() As Boolean
    On Error GoTo errHandler
    
    Dim http As Object
    Dim body As Variant
    Dim pCliente As Integer
    Dim pUsuario As String
    Dim pMaquina As String
    Dim resp As String

    ' Por defecto NO bloqueamos
    PongActivo = True

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    pCliente = 2
    pUsuario = sNomFan             ' mismo usuario que en el ping
    pMaquina = Environ$("COMPUTERNAME")

    body = "{""cliente_id"":" & Chr(34) & pCliente & Chr(34) & _
           ",""usuario"":" & Chr(34) & pUsuario & Chr(34) & _
           ",""maquina"":" & Chr(34) & pMaquina & Chr(34) & "}"

    http.SetTimeouts 2000, 2000, 2000, 2000
    http.Open "POST", "http://comun.dagsis.net.ar/api/vb6/presence/pong", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body

    ' Si no llegó respuesta válida (sin internet, timeout, etc.), no bloqueamos
    If http.Status <> 200 Then
        ' Podrías loguear si querés:
        ' Debug.Print "Error HTTP:", http.Status, http.ResponseText
        PongActivo = True
        Exit Function
    End If

    resp = LCase$(http.ResponseText)
    ' Esperamos algo como: {"activo":true} o {"activo":false}

    If InStr(resp, """activo"":false") > 0 Then
        ' Sólo bloqueamos si EL SERVIDOR responde explícitamente false
        PongActivo = False
    Else
        ' Si dice true o algo raro, dejamos pasar
        PongActivo = True
    End If

    Exit Function

errHandler:
    ' Cualquier error (sin internet, DNS, etc.) ? NO bloquear
    PongActivo = True
End Function


Private Sub ControlarLicencia()
    Call EnviarPresencia

    If Not PongActivo() Then
        MsgBox "Su sistema fue desactivado. Por favor, comuníquese con el servicio técnico.", vbCritical, "Atención"
        End   ' o cerrar formularios, lo que uses
    End If
End Sub

Private Sub Menu(pMenu As Object)
  
  Dim cRsl As ClsLectura, Rs As Recordset, i As Byte, sAux As String, sCli As String

  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRsCondi("Empresa", "Id", "Id<>10")
  
    
  nEmpre(0, 0, 0) = Rs!ID
  nEmpre(0, 1, 0) = Rs!Udl
  nEmpre(0, 1, 1) = Rs!Razon
  
  sUdl = "FILE NAME=" & App.Path & "\" & Rs!Udl
  
 
  
  pMenu(1).Caption = Rs!Razon
  pMenu(1).Checked = True
  i = 2
  Me.BackColor = QBColor(Rs!Color)
  Me.Caption = "Ds Gestión Server - " & Rs!Razon & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal) & sVersion
  
  nMenu = 1
  Rs.MoveNext

  Do While Not Rs.EOF
     Load pMenu(i)
     pMenu(i).Caption = Rs!Razon
     pMenu(i).Checked = False
     i = i + 1
     sAux = "FILE NAME=" & App.Path & "\" & cRsl.TraerValorDeUnCampo("Empresa", "Udl", "Id=" & Rs!ID)
     db.Open sAux
     sCli = cRsl.TraerValorDeUnCampo("Clientes", "Cliente", "Id=" & cRsl.TraerValorDeUnCampo("Registros", "ClienteConsFinal", ""))
     db.Execute "DELETE  FROM CuentaCorriente WHERE Cliente='" & sCli & "'"
     db.Close
     db.Open sUdl
     Rs.MoveNext
  Loop
  If db.State = 1 Then
     db.Close
  End If
  
  
End Sub

Private Sub MenuCajaSaldo_Click()
  SaldoCaja.Show
End Sub

Private Sub MenuCierreTurno_Click()
  On Error GoTo errHandler
  
  Dim Rx1 As Recordset
  Dim cNum As ClsComprobantesL, nMovi As Long
  Dim nCaja As Byte, cLc As ClsLectura, cRx As ClsComprobantesL
  
  Dim nImporte As Double
  
  
  Set cLc = New ClsLectura
  Set cRx = New ClsComprobantesL
  
  If MsgBox("Cierra El Turno ?", 20, "Atención") = 6 Then
        
     nCaja = cLc.RegPorDefecto("Caja")
     nImporte = cRx.TraerCajaAbiertaTurno()

     db.Open sDb
     db.BeginTrans

     db.Execute "UPDATE Caja set Abierta='N' where Caja=" & nCaja
     db.Execute "DELETE FROM Auxiliar "
     db.Execute "DELETE FROM Delivery "
     db.Execute "UPDATE Sistema SET Retencion=0"
     db.Execute "UPDATE Turnos SET Usuario='" & sUser & "', FechaCierre='" & DateTime.Now & "',Importe=" & nImporte & ",Caja=" & nCaja & " WHERE FechaCierre Is Null and Caja=" & nCaja
                                                   
     db.CommitTrans
     db.Close
     MsgBox "Turno Cerrado Con Exito", vbInformation, "Atención"
     End
  End If
Exit Sub
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Private Sub MenuComprasPrecios_Click()
   ComprasAPrecios.Show
End Sub

Private Sub MenuConciliar_Click()
  StockConciliar.Show
End Sub

Private Sub MenuDetCompras_Click()
   CompraInfDetalles.Show
End Sub

Private Sub MenuEmpresa_Click()
  HerraDatosEmpre.Show
End Sub

Private Sub MenuPreciosGrid_Click()
  CambiarPrecios.Show
End Sub

Private Sub MenuReparto_Click()
   VentaReparto.Show
End Sub

Private Sub MenuRepEntregas_Click()
  VentaRepartoEntrega.Show
End Sub

Private Sub MenuRepRemi_Click()
  VentaRepartoRemito.Show
End Sub



Private Sub MHerrPer_Click()
  HerraPermisos.Show
End Sub

Private Sub Moneda_Click()
  TblMoneda.Show
End Sub

Private Sub ResCtaAplic_Click()
  VentaInfResApli.Show
End Sub

Private Sub MVenProy_Click()
  VentasInfProy.Show
End Sub

Private Sub SubMenu_Click(Index As Integer)
    Dim cRsl As ClsLectura
    
    Set cRsl = New ClsLectura
    
  '  db.Close
    
    SubMenu(nMenu).Checked = False
    sDb = sUdl
    db.Open sUdl
    
    sDb = cRsl.TraerValorDeUnCampo("Empresa", "Udl", "Id=" & Index - 1)
    sDb = "FILE NAME=" & App.Path & "\" & sDb

    Me.Caption = "Ds Tango Gestión - " & cRsl.TraerValorDeUnCampo("Empresa", "Razon", "Id=" & Index - 1) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
    nEmpre(0, 0, 0) = cRsl.TraerValorDeUnCampo("Empresa", "Id", "Id=" & Index - 1)
    nEmpre(0, 1, 0) = cRsl.TraerValorDeUnCampo("Empresa", "Udl", "Id=" & Index - 1)
    nEmpre(0, 1, 1) = cRsl.TraerValorDeUnCampo("Empresa", "Razon", "Id=" & Index - 1)
    SubMenu(Index).Checked = True
    Me.BackColor = QBColor(cRsl.TraerValorDeUnCampo("Empresa", "Color", "Id=" & Index - 1))
    nMenu = Index


     db.Close
  '  db.Open sDb
     nPuerto = cRsl.TraerValorDeUnCampo("Usuarios", "Controladora", "Id_Usuario=" & nUsuario)
     nControlador = cRsl.TraerValorDeUnCampo("Usuarios", "Controlador", "Id_Usuario=" & nUsuario)
     nSucursal = cRsl.TraerValorDeUnCampo("Usuarios", "Suc", "Id_Usuario=" & nUsuario)
     nDeposito = cRsl.TraerValorDeUnCampo("Depositos", "Deposito", "Suc=" & nSucursal)

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  End
  Set db = Nothing
End Sub

Private Sub MenProdUnidades_Click()
  MtrProductoUnidad.Show
End Sub

Private Sub MenSalCaja_Click()
  CajaResCaja.Show
End Sub

Private Sub MenuAdeClientes_Click()
  VentasAdelanto.Show
End Sub

Private Sub MenuAdeProv_Click()
  ComprasAdelantos.Show
End Sub

Private Sub MenuCaja_Click()
  TblCajas.Show
End Sub


Private Sub MenuCodBarra_Click()
   pEti = 2
   MtrEtiquetas.Show
End Sub

Private Sub MenuComprobanteCaja_Click()
  CajaComprobantes.Show
End Sub

Private Sub MenuCtaBanc_Click()
  TblCtaCte.Show
End Sub

Private Sub MenuCtaCbles_Click()
   CjaResCtaCble.Show
End Sub

Private Sub MenuDetalles_Click()
  VentaInfDetalles.Show
End Sub

Private Sub MenuDetCob_Click()
  FrmInformeCobranzas.Show
End Sub

Private Sub MenuEtiMovi_Click()
  MtrProdEtiquetas.Show
End Sub

Private Sub MenuEtiquetas_Click()
   pEti = 1
   MtrEtiquetas.Show
End Sub

Private Sub MenuInforMoviCaja_Click()
  CajaInfMovimientos.Show
End Sub
Private Sub MenuInfVenta_Click()
  VentaInfVentas.Show
End Sub

Private Sub MenuIvaCompras_Click()
  ComprasLibroIva.Show
End Sub

Private Sub MenuIvaVenta_Click()
  VentaLibroIva.Show
End Sub

Private Sub MenuPlan_Click()
  TblPlanDeCuenta.Show
End Sub

Private Sub MenuResCueDet_Click()
  VentaInfResumenDet.Show
End Sub
Private Sub MenuResCuenta_Click()
  VentaInfResumen.Show
End Sub

Private Sub MenuSaldoCtaCteBan_Click()
  CajaCtaCteBan.Show
End Sub

Private Sub MenuStockSaldos_Click()
  StockInfSaldos.Show
End Sub

Private Sub MenuTrans_Click()
  TblTransporte.Show
End Sub

Private Sub MenuVencimientos_Click()
  VentaInfVencimiento.Show
End Sub
Private Sub MeVentaAnularCobranza_Click()
  VentaCobranzaAnular.Show
End Sub

Private Sub MHerraRegDef_Click()
  HerraRegDefec.Show
End Sub

Private Sub Movimientos_Click()
  StockInfMovimientos.Show
End Sub

Private Sub MnuChequeCart_Click()
   BancoCheques.Show 1
End Sub

Private Sub MProducProve_Click()
  'MtrProductoProv.Show
End Sub
Private Sub MStockCostos_Click()
  StockInfCostos.Show
End Sub
Private Sub MVentaCobranzas_Click()
  VentaCobranzas.Show 1
End Sub
Private Sub MVentaNotaCredito_Click()
  VentaNotaCredito.Show
End Sub

Private Sub Productos_Click()
  MtrProductos.Show
End Sub
Private Sub Proveedores_Click()
  MtrProveedores.Show
End Sub

Private Sub Ranking_Click()
  VentaRanking.Show
End Sub

Private Sub RemitosPendientes_Click()
  StockInfRemPendientes.Show
End Sub
Private Sub Salir_Click()
  End
End Sub
Private Sub StockComprobante_Click()
  StockComprobantes.Show
End Sub
Private Sub StockFaltantesStock_Click()
  StockInfReposiciones.Show
End Sub
Private Sub StockMovimientos_Click()
  StockInfMovimientos.Show
End Sub

Private Sub TblRecargo_Click()
  TblRecargos.Show
End Sub

Private Sub Timer1_Timer()
   EnviarPresencia
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
 If Button.Visible = True Then
    Select Case Button.key
           Case "Clientes"
                 MtrClientes.Show
           Case "Productos"
                 MtrProductos.Show
           Case "Proveedores"
                MtrProveedores.Show
           Case "Ventas"
                 If cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=35 and Usuario=" & nUsuario) = True Then
                     VentaComprobantes.Show
                 ElseIf cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=85 and Usuario=" & nUsuario) = True Then
                     VentaRapido.Show
                 End If
           Case "Cobranzas"
                 If cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=37 and Usuario=" & nUsuario) = True Then
                    VentaCobranzas.Show 1
                 ElseIf cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=86 and Usuario=" & nUsuario) = True Then
                    VentaCobAuto.Show
                 End If
           Case "Precios"
                 VentaConsPrecios.Show
           Case "PreciosCom"
                 ComprasConsPrecios.Show
           Case "Grillaproduc"
                VentaGrillaProduc.Show
    End Select
  End If
End Sub
Private Sub UniMed_Click()
  TblUniMedida.Show
End Sub
Private Sub Vendedor_Click()
  TblVendedores.Show
End Sub

Private Sub VentaClientes_Click()
  VentaInfClientes.Show
End Sub
Private Sub VentaComprobante_Click()
  VentaComprobantes.Show
End Sub

Private Sub VentaCostoVenta_Click()
  VentaInfCosto.Show
End Sub

Private Sub VentaFactLine_Click()
  VentasFacLine.Show
End Sub

Private Sub VentaInfFami_Click()
  pPlani = 1
  VentaInfFamilias.Show
End Sub
Private Sub VentaListaPrecios_Click()
  VentaInfPrecio.Show
End Sub
Private Sub VentaPrecios_Click()
  VentaActuPrecios.Show
End Sub

Private Sub VentaProductosVendidos_Click()
  VentaProducCosto.Show
End Sub

Private Sub VentaRemitos_Click()
   VentaCompRemitos.Show
End Sub

Private Sub VentasUtilidades_Click()
  pPlani = 2
  VentaInfFamilias.Show
End Sub

Private Sub VerTurnos_Click()
  CajaVerTurnos.Show
End Sub

Private Sub Zona_Click()
  TblZonas.Show
End Sub

Private Sub CerrarCajaX()
   On Error GoTo Errores


  MsgBox "Caja Cerrada con Exito en Modo X", vbInformation, "Atención"
Exit Sub
Errores:
  MsgBox err.Description
End Sub

Private Sub Permisos()
  On Error GoTo Errores:

  Dim cRsl As ClsLectura, RsPermisos As Recordset, P As ClsPrograma
  
  Set cRsl = New ClsLectura
  Set P = New ClsPrograma
    
  MenuHerramientas.Visible = False
  Caja.Visible = False
  MenuVentas.Visible = False
  MenuStock.Visible = False
  Maestros.Visible = False

    
  
  Set RsPermisos = cRsl.TraerRsCondi("UsuariosPermisos", "Id", "Usuario=" & nUsuario)
  P.LoadFromRecordset RsPermisos, nUsuario
  
  pEdit = True ' cRsl.PuedoHacer("Edicion", 1)
  pBorr = True ' cRsl.PuedoHacer("Borrar", 1)
  pInf = True ' cRsl.PuedoHacer("Informes", 1)
  ' Clientes
  pEdit1 = True ' cRsl.PuedoHacer("Edicion", 2)
  pBorr1 = True ' cRsl.PuedoHacer("Borrar", 2)
 ' Clientes.Visible = True ' cRsl.PuedoHacer("Ver", 2)
 
 
  Toolbar1.Buttons(1).Visible = P.HasAny(81) 'cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=81 and Usuario=" & nUsuario)  ' cRsl.PuedoHacer("Ver", 2)
  ' Productos
  pEdit2 = True ' cRsl.PuedoHacer("Edicion", 3)
  pBorr2 = True ' cRsl.PuedoHacer("Borrar", 3)
  Toolbar1.Buttons(2).Visible = P.HasAny(82) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=82 and Usuario=" & nUsuario)  ' cRsl.PuedoHacer("Ver", 3)

  ' Proveedores
  pEdit3 = True ' cRsl.PuedoHacer("Edicion", 4)
  pBorr3 = True ' cRsl.PuedoHacer("Borrar", 4)
  Toolbar1.Buttons(3).Visible = P.HasAny(83) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=83 and Usuario=" & nUsuario)  ' cRsl.PuedoHacer("Ver", 4)
 
  Toolbar1.Buttons(5).Visible = False
  Toolbar1.Buttons(6).Visible = False
  
  
  Toolbar1.Buttons(5).Visible = P.HasAny(83)
  
  
  If P.HasAny(37) = True Then
     Toolbar1.Buttons(6).Visible = True
  ElseIf P.HasAny(86) Then
      Toolbar1.Buttons(6).Visible = True
  End If
    
  Toolbar1.Buttons(8).Visible = P.HasAny(42) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=42 and Usuario=" & nUsuario)  ' cRsl.PuedoHacer("Ver", 4)
  Toolbar1.Buttons(9).Visible = P.HasAny(41) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=41 and Usuario=" & nUsuario)  ' cRsl.PuedoHacer("Ver", 4)
  Toolbar1.Buttons(10).Visible = P.HasAny(41) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=41 and Usuario=" & nUsuario) ' cRsl.PuedoHacer("Ver", 4)
   
    
  Tablas.Visible = P.HasAny(1) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=1 and Usuario=" & nUsuario)
  

  If Tablas.Visible = False Then
     Sep(0).Visible = False
  End If
    
  
  If P.HasAny(81) = False Then
     MClientes.Visible = False
'     Clientes.Visible = False
'     SepCli.Visible = False
  Else
     If P.HasAny(2) = False Then
        MClientes.Visible = False
     Else
        Maestros.Visible = True
        Clientes.Visible = P.HasAny(2) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=2 and Usuario=" & nUsuario)
        CliInformes.Visible = P.HasAny(3) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=3 and Usuario=" & nUsuario)
     End If
  End If
  
  If Clientes.Visible = False And CliInformes.Visible = False Then
     MClientes.Visible = False
  End If
  
  If P.HasAny(82) = False Then
     MProductos.Visible = False
  Else
    Maestros.Visible = True
    Productos.Visible = P.HasAny(4) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=4 and Usuario=" & nUsuario)
  
    MenProdUnidades.Visible = P.HasAny(7) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=7 and Usuario=" & nUsuario)
  
    If Productos.Visible = False And MenProdUnidades.Visible = False Then
       SepProd1.Visible = False
    End If
  
    ListProduc.Visible = P.HasAny(8) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=8 and Usuario=" & nUsuario)
    
    If ListProduc.Visible = False Then
       SepProd.Visible = False
      
    End If
  
    MenuCodBarra.Visible = P.HasAny(11) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=11 and Usuario=" & nUsuario)
  
    MenuEtiquetas.Visible = P.HasAny(10) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=10 and Usuario=" & nUsuario)
    MenuEtiMovi.Visible = P.HasAny(9) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=9 and Usuario=" & nUsuario)
  
  
   If MenuCodBarra.Visible = False Then
      SepProd2.Visible = False
    End If
  
    If SepProd1.Visible = False And SepProd.Visible = False Then
        MProductos.Visible = False
    End If
  
  '  If Productos.Visible = False And MProducProve.Visible = False And MenProdCliente.Visible = False _
  '     And MenProdUnidades.Visible = False And ListProduc.Visible And MenuEtiMovi.Visible = False And _
  '     MenuEtiquetas.Visible And MenuCodBarra.Visible = False Then
  '         MProductos.Visible = False
  '  End If

  End If
'

  If P.HasAny(83) = False Then
     MProveedores.Visible = False
   Else
     Maestros.Visible = True
     Proveedores.Visible = P.HasAny(83) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=83 and Usuario=" & nUsuario) '
     If Proveedores.Visible = False And P.HasAny(13) = False Then
        MProveedores.Visible = False
     End If
  End If
  
  
  MenuStock.Visible = P.HasAny(72) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=72 and Usuario=" & nUsuario)
  If MenuStock.Visible = True Then
     StockComprobante.Visible = P.HasAny(14) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=14 and Usuario=" & nUsuario)
     StockInformes.Visible = P.HasAny(15) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=15 and Usuario=" & nUsuario)
     If StockInformes.Visible = True Then
        RemitosPendientes.Visible = P.HasAny(16) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=16 and Usuario=" & nUsuario)
        If RemitosPendientes.Visible = False Then
           MStockSep1.Visible = False
        End If
        MenuStockSaldos.Visible = P.HasAny(17) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=17 and Usuario=" & nUsuario)
        StockFaltantesStock.Visible = P.HasAny(18) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=18 and Usuario=" & nUsuario)
        Kardex.Visible = P.HasAny(19) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=19 and Usuario=" & nUsuario)
        If MenuStockSaldos.Visible = False And StockFaltantesStock.Visible = False And Kardex.Visible = False Then
           MStockSep1.Visible = False
        End If
        MStockCostos.Visible = P.HasAny(20) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=20 and Usuario=" & nUsuario)
        MenuConciliar.Visible = P.HasAny(94) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=94 and Usuario=" & nUsuario)
        If MStockCostos.Visible = False Then
          MStockSep2.Visible = False
        End If
     End If
  End If

     MenuCompras.Visible = P.HasAny(73) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=73 and Usuario=" & nUsuario)
  If MenuCompras.Visible = True Then
     ComprasFacturacion.Visible = P.HasAny(21) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=21 and Usuario=" & nUsuario)
     CompRemitos.Visible = P.HasAny(22) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=22 and Usuario=" & nUsuario)
     ComprasPagos.Visible = P.HasAny(23) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=23 and Usuario=" & nUsuario)
     MenuAdeProv.Visible = P.HasAny(24) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=24 and Usuario=" & nUsuario)
     ComprasAnular.Visible = P.HasAny(25) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=25 and Usuario=" & nUsuario)
     ComprasPrecios.Visible = P.HasAny(26) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=26 and Usuario=" & nUsuario)
     MenuComprasPrecios.Visible = P.HasAny(27) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=27 and Usuario=" & nUsuario)
     If ComprasPrecios.Visible = False And MenuComprasPrecios.Visible = False Then
        ComprasSep.Visible = False
     End If
     ComprasInformes.Visible = P.HasAny(28) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=28 and Usuario=" & nUsuario)
     If ComprasInformes.Visible = False Then
        ComprasSep2.Visible = False
     End If
     If ComprasInformes.Visible = True Then
        ComprasResCta.Visible = P.HasAny(29) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=29 and Usuario=" & nUsuario)
        ComprasVencimientos.Visible = P.HasAny(30) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=30 and Usuario=" & nUsuario)
        MenuIvaCompras.Visible = P.HasAny(31) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=31 and Usuario=" & nUsuario)
        MenuDetCompras.Visible = P.HasAny(29) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=29 and Usuario=" & nUsuario)
        MCitiCompras.Visible = P.HasAny(32) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=32 and Usuario=" & nUsuario)
        If MenuIvaCompras.Visible = False And MCitiCompras.Visible = False Then
           SepCompras.Visible = False
        End If
        InfRetIb.Visible = P.HasAny(33) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=33 and Usuario=" & nUsuario)
        InfRetGanancia.Visible = P.HasAny(89) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=89 and Usuario=" & nUsuario)
        If InfRetIb.Visible = False Then
           SepIvaCompras.Visible = False
        End If
        MnuChequeCart.Visible = P.HasAny(34) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=34 and Usuario=" & nUsuario)
        If MnuChequeCart.Visible = False Then
           SepChe.Visible = False
        End If
     End If
  End If


  MenuVentas.Visible = P.HasAny(74) 'cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=74 and Usuario=" & nUsuario)
  If MenuVentas.Visible = True Then
     VentaComprobante.Visible = P.HasAny(35) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=35 and Usuario=" & nUsuario)
     VentaRemitos.Visible = P.HasAny(36) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=36 and Usuario=" & nUsuario)
     MVentaCobranzas.Visible = P.HasAny(37) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=37 and Usuario=" & nUsuario)
     MenuAdeClientes.Visible = P.HasAny(38) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=38 and Usuario=" & nUsuario)
     MeVentaAnularCobranza.Visible = P.HasAny(39) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=39 and Usuario=" & nUsuario)
     VentaPrecios.Visible = P.HasAny(40) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=40 and Usuario=" & nUsuario)
     MenuPreciosGrid.Visible = P.HasAny(41) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=41 and Usuario=" & nUsuario)
     ConsultoPrecios.Visible = P.HasAny(42) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=42 and Usuario=" & nUsuario)
     If VentaPrecios.Visible = False And VentaPrecios.Visible = False And MenuPreciosGrid.Visible = False Then
        VentaSep7.Visible = False
     End If
     VentasInformes.Visible = P.HasAny(84) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=84 and Usuario=" & nUsuario)
     If VentasInformes.Visible = False Then
        VentaSep5.Visible = False
     End If
     If VentasInformes.Visible = True Then
        MenuResCuenta.Visible = P.HasAny(43) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=43 and Usuario=" & nUsuario)
        MenuResCueDet.Visible = P.HasAny(44) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=44 and Usuario=" & nUsuario)
        MenResCtaFlia.Visible = P.HasAny(45) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=45 and Usuario=" & nUsuario)
        MenuVencimientos.Visible = P.HasAny(46) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=46 and Usuario=" & nUsuario)
        MVenProy.Visible = P.HasAny(47) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=47 and Usuario=" & nUsuario)
        MenuDetalles.Visible = P.HasAny(48) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=48 and Usuario=" & nUsuario)
        MenuDetCob.Visible = P.HasAny(49) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=49 and Usuario=" & nUsuario)
        MenuInfVenta.Visible = P.HasAny(50) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=50 and Usuario=" & nUsuario)
        VentaInfFami.Visible = P.HasAny(51) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=51 and Usuario=" & nUsuario)
        VentasUtilidades.Visible = P.HasAny(52) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=52 and Usuario=" & nUsuario)
        ControlGastos.Visible = P.HasAny(53) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=53 and Usuario=" & nUsuario)
        Ranking.Visible = P.HasAny(51) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=54 and Usuario=" & nUsuario)
        Consolidados.Visible = P.HasAny(55) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=55 and Usuario=" & nUsuario)
        MenuReparto.Visible = P.HasAny(56) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=56 and Usuario=" & nUsuario)
        MenuRepRemi.Visible = P.HasAny(91) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=91 and Usuario=" & nUsuario)
        MenuRepEntregas.Visible = P.HasAny(92) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=92 and Usuario=" & nUsuario)
        
        If MenuInfVenta.Visible = False And VentaInfFami.Visible = False And VentasUtilidades.Visible = False _
           And ControlGastos.Visible = False And Ranking.Visible = False And Consolidados.Visible = False And MenuReparto.Visible = False Then
           VentaSep1(1).Visible = False
        End If

        VentaFactLine.Visible = P.HasAny(87) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=87 and Usuario=" & nUsuario)
        MenuIvaVenta.Visible = P.HasAny(57) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=57 and Usuario=" & nUsuario)
        MCitiVentas.Visible = P.HasAny(58) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=58 and Usuario=" & nUsuario)
        If MenuIvaVenta.Visible = False And MCitiVentas.Visible = False Then
          SepVent.Visible = False
        End If
        InfPerIB.Visible = P.HasAny(59) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=59 and Usuario=" & nUsuario)
        If InfPerIB.Visible = False Then
           SepIva.Visible = False
        End If
        VentaCosteos.Visible = P.HasAny(60) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=60 and Usuario=" & nUsuario)
        If VentaCosteos.Visible = False Then
           VentaSep.Visible = False
        End If
        VentaListaPrecios.Visible = P.HasAny(61) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=61 and Usuario=" & nUsuario)
        If VentaListaPrecios.Visible = False Then
           VentaSep2.Visible = False
        End If
     End If
  End If

  Caja.Visible = P.HasAny(75) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=75 and Usuario=" & nUsuario)
  If Caja.Visible = True Then

     MCajaRap.Visible = P.HasAny(85) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=85 and Usuario=" & nUsuario)
     MCobSaldos.Visible = P.HasAny(86) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=86 and Usuario=" & nUsuario)
     
     If MCajaRap.Visible = False And MCobSaldos.Visible = False Then
        SepCaja.Visible = False
     End If
           
     MenuComprobanteCaja.Visible = P.HasAny(62) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=62 and Usuario=" & nUsuario)
     MenuInforCaja.Visible = P.HasAny(63) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=63 and Usuario=" & nUsuario)
     If MenuInforCaja.Visible = False Then
        VentaSep.Visible = False
     End If
     If MenuComprobanteCaja.Visible = False Then
       CajaSep2.Visible = False
     End If
     If MenuInforCaja.Visible = True Then
        MenuInforMoviCaja.Visible = P.HasAny(64) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=64 and Usuario=" & nUsuario)
        MenSalCaja.Visible = P.HasAny(65) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=65 and Usuario=" & nUsuario)
        CajaChequeTer.Visible = P.HasAny(66) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=66 and Usuario=" & nUsuario)
        MenuSaldoCtaCteBan.Visible = P.HasAny(67) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=67 and Usuario=" & nUsuario)
        MenuCtaCbles.Visible = P.HasAny(68) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=68 and Usuario=" & nUsuario)
        If CajaChequeTer.Visible = False And MenuSaldoCtaCteBan.Visible = False And MenuCtaCbles.Visible = False Then
           SepCa.Visible = False
        End If
     End If
     If MenuInforCaja.Visible = False Then
      '  CajaSep2.Visible = False
        CajaSep1.Visible = False
        SepCa.Visible = False
     End If

     
     MenuCierreTurno.Visible = P.HasAny(90) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=90 and Usuario=" & nUsuario)
     
     MenuCajaSaldo.Visible = P.HasAny(93) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=93 and Usuario=" & nUsuario)
     CierreManual.Visible = P.HasAny(71) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=71 and Usuario=" & nUsuario)

     If CierreManual.Visible = False Then
        SepCierre.Visible = False
     End If
  End If


  MenuHerramientas.Visible = P.HasAny(76) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=76 and Usuario=" & nUsuario)
  If MenuHerramientas.Visible = True Then
    MHerrPer.Visible = P.HasAny(88) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=88 and Usuario=" & nUsuario)
    MHerraRegDef.Visible = bAdmin
    HerrUserPer.Visible = bAdmin
    MHerrConf.Visible = P.HasAny(79) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=79 and Usuario=" & nUsuario)
    HerrImporProduc.Visible = P.HasAny(80) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=80 and Usuario=" & nUsuario)
    HerrConfigCorreo.Visible = P.HasAny(96) ' cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=96 and Usuario=" & nUsuario)
  End If
  Exit Sub
Errores:
  ' MsgBox Err.Description, vbCritical, "Atención"
   Resume Next
End Sub

Private Sub CrearDirectorios()
  Dim carpeta As String
  
   ' Carpeta de Facturas
  carpeta = Dir(App.Path & "\Facturas", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\Facturas\"
  End If
  
   ' Carpeta de CitiCompras
  carpeta = Dir(App.Path & "\CitiCompras", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\CitiCompras\"
  End If
  
   ' Carpeta de CitiVentas
  carpeta = Dir(App.Path & "\CitiVentas", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\CitiVentas\"
  End If
  ' Carpeta de Cot
  carpeta = Dir(App.Path & "\ArbaCot", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\ArbaCot\"
  End If
  
   ' Carpeta de ArbaPercepcion
  carpeta = Dir(App.Path & "\ArbaPercepcion", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\ArbaPercepcion\"
  End If
  
   ' Carpeta de ArbaRetencion
  carpeta = Dir(App.Path & "\ArbaRetencion", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\ArbaRetencion\"
  End If
  
   ' Carpeta de RetGanacia
  carpeta = Dir(App.Path & "\RetGanacia", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\RetGanacia\"
  End If
  
    ' Carpeta de Scripts
  carpeta = Dir(App.Path & "\Scripts", vbDirectory)
  If carpeta = "" Then
     MkDir App.Path & "\Scripts\"
  End If
  
End Sub

