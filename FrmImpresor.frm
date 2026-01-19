VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmImpresor 
   Caption         =   "o"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3015
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4485
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmImpresor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim nDesde As Double, nHasta As Double
Dim sDesde As String, sHasta As String, pLista As Integer
Dim sPdesde As String, sPhasta As String
Dim fDesde As Date, fHasta As Date
Dim cClases As ClsPrecios, cCanti As ClsProductoL
Dim cCliente As ClsLectura, Rs As ADODB.Recordset
Dim cCaja As ClsComprobantesL

Private Sub Form_Load()
  
  Set cCliente = New ClsLectura
  Set Rs = New ADODB.Recordset
  
  Screen.MousePointer = 11
  Select Case nImpr
         Case 1
              Impre01
         Case 2
              Impre02
         Case 3
              Impre03
         Case 4
              Impre04
         Case 5
              Impre05
         Case 6
              Impre06
         Case 7
              Impre07
         Case 8
              Impre08
         Case 9
              Impre09
         Case 10
              Impre10
         Case 11
              Impre11
         Case 12
              Impre12
         Case 13
              Impre13
         Case 14
              Impre14
         Case 15
              Impre15
         Case 16
              Impre16
         Case 17
              Impre17
         Case 18
              Impre18
         Case 19
              Impre19
         Case 20
              Impre20
         Case 21
              Impre21
         Case 22
              Impre22
         Case 23
              Impre23
         Case 24
              Impre24
         Case 25
              Impre25
         Case 26
              Impre26
         Case 27
              Impre27
         Case 28
              Impre28
         Case 29
              Impre29
         Case 30
              Impre30
         Case 31
              Impre31
         Case 32
              Impre32
         Case 33
              Impre33
         Case 34
              Impre34
         Case 35
              Impre35
         Case 36
              Impre36
         Case 37
              Impre37
         Case 38
              Impre38
         Case 39
              Impre39
         Case 40
              Impre40
         Case 41
              Impre41
         Case 42
              Impre42
         Case 43
              Impre43
         Case 44
              Impre44
         Case 45
              Impre45
         Case 46
              Impre46
         Case 47
              Impre47
         Case 48
              Impre48
         Case 49
              Impre49
         Case 50
              Impre50
         Case 51
              Impre51
         Case 52
              Impre52
         Case 53
              Impre53
         Case 54
              Impre54
         Case 55
              Impre55
         Case 56
              Impre56
         Case 57
              Impre57
         Case 58
              Impre58
         Case 59
              Impre59
         Case 60
              Impre60
         Case 61
              Impre61
         Case 62
              Impre62
         Case 63
              Impre63
         Case 64
              Impre64
         Case 65
              Impre65
         Case 66
              Impre66
         Case 67
              Impre67
         Case 68
              Impre68
         Case 69
              Impre69
  End Select
  
  If Rs.RecordCount <> 0 Then
     CRViewer1.EnablePrintButton = True

     CRViewer1.DisplayGroupTree = False
     CRViewer1.EnableExportButton = True
     CRViewer1.EnableRefreshButton = False
    
     CRViewer1.Zoom (100)
     CRViewer1.ViewReport
  Else
     MsgBox "No hay Registros. Intente con Nuevos Parametros", vbCritical, "Atención"
  End If
  Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
  CRViewer1.Top = 0
  CRViewer1.Left = 0
  CRViewer1.Height = ScaleHeight
  CRViewer1.Width = ScaleWidth
End Sub

Private Sub Impre01()
  Dim report1 As New RptVentaResumen, cRsl As ClsLectura

  Set Rs = VentaInfResumen.Rs
  Set cRsl = New ClsLectura
  
  If Rs.RecordCount <> 0 Then
     report1.ReportTitle = sNomFan
     report1.ParameterFields(1).SetCurrentValue IIf(VentaInfResumen.ChkSaldo.Value = 1, True, False)
     report1.ParameterFields(2).SetCurrentValue VentaInfResumen.CmbZona.text
     report1.ParameterFields(3).SetCurrentValue IIf(VentaInfResumen.ChkTexto.Value = 1, True, False)
     report1.ParameterFields(4).SetCurrentValue IIf(VentaInfResumen.ChkSalto.Value = 1, True, False)
     report1.ParameterFields(5).SetCurrentValue cRsl.TraerValorDeUnCampo("TblTextos", "Texto", "Id=1")
     report1.ParameterFields(6).SetCurrentValue cRsl.TraerValorDeUnCampo("TblTextos", "Texto", "Id=2")
     report1.Database.SetDataSource Rs
     CRViewer1.ReportSource = report1
     CRViewer1.DisplayGroupTree = False
  End If
  Set report1 = Nothing
End Sub

Private Sub Impre02()
  Dim Report2 As New RptVentaVencimiento

  Report2.ReportTitle = "Vencimiento por Zona :" & VentaInfVencimiento.CmbZona.text
  
  Set Rs = VentaInfVencimiento.RsP
  If Rs.RecordCount <> 0 Then
     Report2.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report2
  End If
  Set Report2 = Nothing
End Sub

Private Sub Impre03()
  Dim Report9 As New RptVentaPreciosGrupo, bOrd As Integer
  Dim Report10 As New RptVentaPreciosMarca
  Dim Report3 As New RptVentaPrecios

  Set cClases = New ClsPrecios
          
  pLista = VentaInfPrecio.CmbLista.ItemData(VentaInfPrecio.CmbLista.ListIndex)
  
    
  If VentaInfPrecio.OptCodigo = True Then
     bOrd = 0
     sDesde = VentaInfPrecio.TxtCodigo(0).text
     sHasta = VentaInfPrecio.TxtCodigo(1).text
  Else
    bOrd = 1
     sDesde = VentaInfPrecio.LblDescripcion(0).Caption
     sHasta = VentaInfPrecio.LblDescripcion(1).Caption
  End If
  
  sPdesde = VentaInfPrecio.CmbDesde.text
  sPhasta = VentaInfPrecio.CmbHasta.text
  
  If VentaInfPrecio.OptNinguno = True Then
     Set Rs = cClases.DameListaProductos(pLista, sDesde, sHasta, bOrd)
     If Rs.RecordCount <> 0 Then
        Report3.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.text
        Report3.Database.SetDataSource Rs
        CRViewer1.ReportSource = Report3
     End If
  Else
     If VentaInfPrecio.OptMarcas = True Then
        Set Rs = cClases.DameListaProductosMarca(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report10.ReportTitle = "Listado de Precios Agrupados por Marca"
           Report10.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.text
           Report10.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report10
        End If
     Else
        Set Rs = cClases.DameListaProductosGrupo(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report9.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.text
           Report9.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report9
        End If
    End If
  End If
  Set Report9 = Nothing
  Set Report10 = Nothing
  Set Report3 = Nothing
End Sub
Private Sub Impre04()

  Dim Report9 As New RptVentaPreciosGrupo
  Dim Report10 As New RptVentaPreciosMarca
  Dim Report3 As New RptVentaPrecios, bOrd As Integer

  Set cClases = New ClsPrecios
         
  pLista = VentaInfPrecio.CmbLista.ItemData(VentaInfPrecio.CmbLista.ListIndex)
  
  If VentaInfPrecio.OptCodigo = True Then
     bOrd = 0
     sDesde = VentaInfPrecio.TxtCodigo(0).text
     sHasta = VentaInfPrecio.TxtCodigo(1).text
  Else
    bOrd = 1
     sDesde = VentaInfPrecio.LblDescripcion(0).Caption
     sHasta = VentaInfPrecio.LblDescripcion(1).Caption
  End If
  sPdesde = VentaInfPrecio.CmbDesde.text
  sPhasta = VentaInfPrecio.CmbHasta.text
  If VentaInfPrecio.OptNinguno.Value = True Then
     Set Rs = cClases.DameListaProductos(pLista, sDesde, sHasta, bOrd)
     If Rs.RecordCount <> 0 Then
        Report3.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.text
        Report3.Database.SetDataSource Rs
        CRViewer1.ReportSource = Report3
     End If
  Else
     If VentaInfPrecio.OptMarcas.Value = True Then
        Set Rs = cClases.DameListaProductosMarca(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report10.ReportTitle = "Listado de Precios Agrupados por Marca"
           Report10.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.text
           Report10.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report10
        End If
     Else
        Set Rs = cClases.DameListaProductosGrupo(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report9.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.text
           Report9.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report9
        End If
    End If
  End If
  Set Report9 = Nothing
  Set Report10 = Nothing
  Set Report3 = Nothing
  
End Sub
Private Sub Impre05()
  Dim Report4 As New RptStockSaldo, bSto As Byte, bOrd As Integer

  Set cCanti = New ClsProductoL
              
  pLista = StockInfSaldos.CmbDeposito.ItemData(StockInfSaldos.CmbDeposito.ListIndex)
  If StockInfSaldos.OptCodigo = True Then
     bOrd = 0
     sDesde = StockInfSaldos.TxtCodigo(0).text
     sHasta = StockInfSaldos.TxtCodigo(1).text
  Else
    bOrd = 1
     sDesde = StockInfSaldos.LblDescripcion(0).Caption
     sHasta = StockInfSaldos.LblDescripcion(1).Caption
  End If
  bSto = StockInfSaldos.ChkSaldo.Value
  
  Set Rs = cCanti.DameProductos(pLista, sDesde, sHasta, bOrd, bSto)
  If Rs.RecordCount <> 0 Then
     Report4.ParameterFields(1).SetCurrentValue StockInfSaldos.CmbDeposito.text
     Report4.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report4
  End If
  Set Report4 = Nothing
End Sub

Private Sub Impre06()
  Dim Report4 As New RptStockSaldo, bOrd As Integer
  Dim nStock As Byte, pFamilia As Integer

  Set cCanti = New ClsProductoL
              
  pLista = StockInfSaldos.CmbDeposito.ItemData(StockInfSaldos.CmbDeposito.ListIndex)
  pFamilia = StockInfSaldos.CmbFamilia.ItemData(StockInfSaldos.CmbFamilia.ListIndex)
  
  If StockInfSaldos.OptCodigo = True Then
     bOrd = 0
     sDesde = StockInfSaldos.TxtCodigo(0).text
     sHasta = StockInfSaldos.TxtCodigo(1).text
  Else
     bOrd = 1
     sDesde = StockInfSaldos.LblDescripcion(0).Caption
     sHasta = StockInfSaldos.LblDescripcion(1).Caption
  End If
  
  nStock = StockInfSaldos.ChkSaldo.Value
  
  
  Set Rs = cCanti.DameProductos(pLista, sDesde, sHasta, bOrd, nStock, pFamilia)
  
  If Rs.RecordCount <> 0 Then
     Report4.ParameterFields(1).SetCurrentValue StockInfSaldos.CmbDeposito.text
     Report4.ParameterFields(2).SetCurrentValue StockInfSaldos.CmbFamilia.text
     Report4.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report4
  End If
  Set Report4 = Nothing
End Sub

Private Sub Impre07()

  Dim Report5 As New RptStockReposicion, pFamilias As Integer

  Set cCanti = New ClsProductoL
             
  pFamilias = StockInfSaldos.CmbFamilia.ItemData(StockInfSaldos.CmbFamilia.ListIndex)
             
  sDesde = StockInfReposiciones.TxtCodigo(0)
  sHasta = StockInfReposiciones.TxtCodigo(1)
  sPdesde = StockInfReposiciones.CmbProveedor(0).text
  Set Rs = cCanti.DameReposiciones(sDesde, sHasta, Val(sPdesde), 1)
  If Rs.RecordCount <> 0 Then
     Report5.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report5
  End If
  Set Report5 = Nothing
End Sub

Private Sub Impre08()
  Dim Report5 As New RptStockReposicion

  Set cCanti = New ClsProductoL
             
  sDesde = StockInfReposiciones.LblDescripcion(0)
  sHasta = StockInfReposiciones.LblDescripcion(1)
  sPdesde = StockInfReposiciones.CmbProveedor(0).ItemData(StockInfReposiciones.CmbProveedor(0).ListIndex)
  Set Rs = cCanti.DameReposiciones(sDesde, sHasta, Val(sPdesde), 2)
  If Rs.RecordCount <> 0 Then
     Report5.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report5
  End If
  Set Report5 = Nothing
End Sub

Private Sub Impre09()

End Sub
Private Sub Impre10()
  Dim Report7 As New RptClienteSinZona

  Set Rs = VentaInfClientes.Rs
  If Rs.RecordCount <> 0 Then
     Report7.ReportTitle = "Listado de Clientes"
     Report7.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report7
  End If
  Set Report7 = Nothing
End Sub

Private Sub Impre11()
  Dim Report8 As New RptClientesZona

  Set Rs = VentaInfClientes.Rs
  If Rs.RecordCount <> 0 Then
     Report8.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report8
  End If
  Set Report8 = Nothing
End Sub

Private Sub Impre12()
  Dim Report9 As New RptVentaPreciosGrupo
  Dim Report10 As New RptVentaPreciosMarca
  Dim Report3 As New RptVentaPrecios

  Set cClases = New ClsPrecios
   
  sDesde = StockInfCostos.LblDescripcion(0)
  sHasta = StockInfCostos.LblDescripcion(1)
  sPdesde = StockInfCostos.CmbDesde.text
  sPhasta = StockInfCostos.CmbHasta.text
  If StockInfCostos.OptMarcas = True Then
     Set Rs = cClases.DameCostoProductos(sDesde, sHasta, 3)
     If Rs.RecordCount <> 0 Then
        Report3.ParameterFields(1).SetCurrentValue ""
        Report3.ReportTitle = "Listado de Costos"
        Report3.Database.SetDataSource Rs
        CRViewer1.ReportSource = Report3
     End If
  Else
     If StockInfCostos.OptFamilias = True Then
        Set Rs = cClases.DameCostoProductosMarca(sDesde, sHasta, 3, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report10.ParameterFields(1).SetCurrentValue ""
           Report10.ReportTitle = "Listado de Costos Agrupados por Marca"
           Report10.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report10
        End If
     Else
        Set Rs = cClases.DameCostoProductosGrupo(sDesde, sHasta, 3, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report9.ParameterFields(1).SetCurrentValue ""
           Report9.ReportTitle = "Listado de Costos Agrupados por Familias"
           Report9.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report9
        End If
      End If
   End If
   Set Report3 = Nothing
   Set Report9 = Nothing
   Set Report10 = Nothing
End Sub
Private Sub Impre13()

  Dim Report9 As New RptVentaPreciosGrupo
  Dim Report10 As New RptVentaPreciosMarca
  Dim Report3 As New RptVentaPrecios

  Set cClases = New ClsPrecios
        
  sDesde = StockInfCostos.LblDescripcion(0)
  sHasta = StockInfCostos.LblDescripcion(1)
  sPdesde = StockInfCostos.CmbDesde.text
  sPhasta = StockInfCostos.CmbHasta.text
  If StockInfCostos.OptMarcas.Value = True Then
     Set Rs = cClases.DameCostoProductos(sDesde, sHasta, 4)
     If Rs.RecordCount <> 0 Then
        Report3.ParameterFields(1).SetCurrentValue ""
        Report3.Database.SetDataSource Rs
        CRViewer1.ReportSource = Report3
     End If
  Else
     If StockInfCostos.OptFamilias = True Then
        Set Rs = cClases.DameCostoProductosMarca(sDesde, sHasta, 4, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report10.ParameterFields(1).SetCurrentValue ""
           Report10.ReportTitle = "Listado de Costos Agrupados por Marca"
           Report10.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report10
        End If
     Else
        Set Rs = cClases.DameCostoProductosGrupo(sDesde, sHasta, 4, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
           Report9.ParameterFields(1).SetCurrentValue ""
           Report9.ReportTitle = "Listado de Costos Agrupados por Familias"
           Report9.Database.SetDataSource Rs
           CRViewer1.ReportSource = Report9
        End If
     End If
  End If
   Set Report3 = Nothing
   Set Report9 = Nothing
   Set Report10 = Nothing
  
End Sub

Private Sub Impre14()

  Dim Report11 As New RptVentasFecha

  Set cCaja = New ClsComprobantesL
  fDesde = VentaInfVentas.DTPicker1.Value
  fHasta = VentaInfVentas.DTPicker2.Value
  
  Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
  If Rs.RecordCount <> 0 Then
     Report11.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report11
  End If
  Set Report11 = Nothing
End Sub

Private Sub Impre15()
  Dim Report12 As New RptVentasComprobantes

  Set cCaja = New ClsComprobantesL
  fDesde = VentaInfVentas.DTPicker1.Value
  fHasta = VentaInfVentas.DTPicker2.Value
  
  Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
  If Rs.RecordCount <> 0 Then
     Report12.ReportTitle = "Comprobantes de Ventas"
     Report12.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report12
  End If
  Set Report12 = Nothing
End Sub
Private Sub Impre16()

  Dim Report13 As New RptVentasVendedor

  Set cCaja = New ClsComprobantesL
  fDesde = VentaInfVentas.DTPicker1.Value
  fHasta = VentaInfVentas.DTPicker2.Value
  
  Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
  If Rs.RecordCount <> 0 Then
     Report13.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report13
  End If
  Set Report13 = Nothing
End Sub

Private Sub Impre17()
  Dim Report14 As New RptVentasCajas

  Set cCaja = New ClsComprobantesL
  fDesde = VentaInfVentas.DTPicker1.Value
  fHasta = VentaInfVentas.DTPicker2.Value
  
  Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
  If Rs.RecordCount <> 0 Then
     Report14.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report14
  End If
  Set Report14 = Nothing
End Sub

Private Sub Impre18()
  Dim Report15 As New RptVentasProductos
  
  Set Rs = VentaInfVentas.Rs
  
  If Rs.RecordCount <> 0 Then
     Report15.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report15
  End If
  Set Report15 = Nothing
End Sub

Private Sub Impre19()
  Dim Report16 As New RptDetalleFactura

  Dim cRle As New ClsClienteL, nOrd As Byte, sDesde As String, sHasta As String
  
  If VentaInfDetalles.OptCodigo.Value = True Then
     nOrd = 0
     sDesde = VentaInfDetalles.TxtCliente(0).text
     sHasta = VentaInfDetalles.TxtCliente(1).text
  Else
     sDesde = VentaInfDetalles.LblCliente(0).Caption
     sHasta = VentaInfDetalles.LblCliente(1).Caption
     nOrd = 1
  End If
  
  Set Rs = cRle.TraerDetalleFactura(sDesde, sHasta, VentaInfDetalles.DtpFecha(0).Value, VentaInfDetalles.DtpFecha(1).Value, nOrd)
  If Rs.RecordCount <> 0 Then
     Report16.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report16
  End If
  Set Report16 = Nothing
End Sub


Private Sub Impre33()
  Dim Report16 As New RptDetalleFactura

  Dim cRle As New ClsClienteL, nOrd As Byte, sDesde As String, sHasta As String
  
  If CompraInfDetalles.OptCodigo.Value = True Then
     nOrd = 0
     sDesde = CompraInfDetalles.TxtCliente(0).text
     sHasta = CompraInfDetalles.TxtCliente(1).text
  Else
     sDesde = CompraInfDetalles.LblCliente(0).Caption
     sHasta = CompraInfDetalles.LblCliente(1).Caption
     nOrd = 1
  End If
  
  Report16.ReportTitle = "Proveedores - Detalles de Facturas"
  
  
  Set Rs = cRle.TraerDetalleFacturaP(sDesde, sHasta, CompraInfDetalles.DtpFecha(0).Value, CompraInfDetalles.DtpFecha(1).Value, nOrd)
  
  If Rs.RecordCount <> 0 Then
     Report16.Database.SetDataSource Rs
     Report16.Text4.SetText ("Proveedor")
     CRViewer1.ReportSource = Report16
  End If
  Set Report16 = Nothing
End Sub


Private Sub Impre20()
  Dim Report17 As New RptProductoGene

  Dim crField As CRAXDRT.DatabaseFieldDefinition
  Dim crSection As CRAXDRT.Section
  Dim pGru As Byte, pDesde As String, pHasta As String
  
  Set cCanti = New ClsProductoL
  Dim a As String
  pGru = StockInfProductos.nOrd
  pDesde = StockInfProductos.sDesde
  pHasta = StockInfProductos.sHasta
  
  Set Rs = cCanti.ProductoListadoMtr(StockInfProductos.LblProducto(0).Caption, StockInfProductos.LblProducto(1).Caption, StockInfProductos.ChkAlfabetico.Value, pGru, pDesde, pHasta)
  If pGru = 0 Then
     If StockInfProductos.ChkAlfabetico.Value = 0 Then pGru = 1
     If StockInfProductos.ChkAlfabetico.Value = 1 Then pGru = 2
  End If
  If Rs.RecordCount <> 0 Then
     Set crField = Report17.Database.Tables.Item(1).Fields.Item(pGru)
     Report17.Areas.Item("GH").GroupConditionField = crField
     'Report17.AddGroup 0, crField, crGCAnyValue, crAscendingOrder
     Report17.Database.SetDataSource Rs
     Report17.Areas.Item("GH").Suppress = False
     CRViewer1.ReportSource = Report17
  End If
  Set Report17 = Nothing
End Sub

Private Sub Impre21()
  Dim Report18 As New RptValoCosto, bOrd As Integer

  Dim cCl As ClsProductoL
  Set cCl = New ClsProductoL
  
  If VentaInfCosto.OptCodigo = True Then
     bOrd = 0
  Else
     bOrd = 1
  End If
  
  sDesde = VentaInfCosto.LblDescripcion(0).Caption
  sHasta = VentaInfCosto.LblDescripcion(1).Caption
  
  Set Rs = cCl.TraerValoInventario(sDesde, sHasta, bOrd)
  
  If Rs.RecordCount <> 0 Then
     Report18.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report18
  End If
  Set Report18 = Nothing
End Sub

Private Sub Impre22()
  Dim Report19 As New RptRentabilid

  Dim nOp As Byte, cCl As ClsProductoL
  
  Set cClases = New ClsPrecios
  Set cCl = New ClsProductoL
  
  Set Rs = cCl.TraerRentabilidad(VentaProducCosto.DtpDesde.Value, VentaProducCosto.DtpHasta.Value, VentaProducCosto.LblDescripcion(0).Caption, VentaProducCosto.LblDescripcion(1).Caption)
  
  If Rs.RecordCount <> 0 Then
     Report19.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report19
  End If
  Set Report19 = Nothing
End Sub

Private Sub Impre23()
  Dim Report20 As New RptRemPendientes, rPrin As Recordset, rAux As Recordset
  Dim nCanRem As Single, nCanRemAux As Single

  Dim cRem As New ClsComprobantesL
  Set rPrin = New Recordset
  Set rAux = New Recordset
  
  rPrin.Fields.Append "Comprobante", adVarChar, 30
  rPrin.Fields.Append "Fecha", adDate
  rPrin.Fields.Append "Sucursal", adInteger
  rPrin.Fields.Append "Numero", adInteger
  rPrin.Fields.Append "Cliente", adVarChar, 15
  rPrin.Fields.Append "RazonSocial", adVarChar, 50
  rPrin.Fields.Append "Importe", adCurrency
  rPrin.Open
  
   Dim j As Long, nCantRem As Double, nCantRemAux As Double, RsAux As Recordset
  
      
  Set Rs = cRem.TraerInfRemitosPend(StockInfRemPendientes.LblCliente(0), StockInfRemPendientes.LblCliente(1), StockInfRemPendientes.DtpFecha(0).Value, StockInfRemPendientes.DtpFecha(1).Value)

   Do While Not Rs.EOF
     Set rAux = cRem.TraerDetallesMovimiento(Rs!Movimiento)
     Do While Not rAux.EOF
        nCantRem = Round(cRem.TraerCantidadRemito(rAux!Producto, Rs!Movimiento, rAux!ID), 2)
        nCantRemAux = Round(rAux!Cantidad - Abs(nCantRem), 2)
        If nCantRemAux > 0 Or nCantRem = -1 Then
            rPrin.AddNew
            rPrin!Comprobante = Rs!Comprobante
            rPrin!fecha = Rs!fecha
            rPrin!Sucursal = Rs!Sucursal
            rPrin!Numero = Rs!Numero
            rPrin!Cliente = Rs!Cliente
            rPrin!RazonSocial = Rs!RazonSocial
            rPrin!importe = Abs(IIf(IsNull(Rs!Expr1), 0, Rs!Expr1))
            rPrin.Update
            Exit Do
        End If
        rAux.MoveNext
     Loop
     Rs.MoveNext
  Loop
  
'  Set rs = cRem.TraerInfRemitosPend(StockInfRemPendientes.LblCliente(0), StockInfRemPendientes.LblCliente(1), StockInfRemPendientes.DTPFecha(0).Value, StockInfRemPendientes.DTPFecha(1).Value)
'  Do While Not rs.EOF
'     Set RAux = cRem.TraerDetallesMovimiento(rs!Movimiento)
'     Do While Not RAux.EOF
'        If RAux!Producto <> "" Then
'           nCanRem = Round(cRem.TraerCantidadRemito(RAux!Producto, rs!Movimiento, RAux!Id), 2)
'           nCanRemAux = Round(RAux!Cantidad - nCanRem, 2)
'           Else
'               rPrin.AddNew
'               rPrin!Comprobante = rs!Comprobante
'               rPrin!fecha = rs!fecha
'               rPrin!Sucursal = rs!Sucursal
'               rPrin!Numero = rs!Numero
'               rPrin!Cliente = rs!Cliente
'               rPrin!RazonSocial = rs!RazonSocial
'               rPrin!importe = Abs(IIf(IsNull(rs!Expr1), 0, rs!Expr1))
'               rPrin.Update
'               Exit Do
'        End If
'        If nCanRemAux > 0 Or nCanRem = -1 Then
'           rPrin.AddNew
'           rPrin!Comprobante = rs!Comprobante
'           rPrin!fecha = rs!fecha
'           rPrin!Sucursal = rs!Sucursal
'           rPrin!Numero = rs!Numero
'           rPrin!Cliente = rs!Cliente
'           rPrin!RazonSocial = rs!RazonSocial
'           rPrin!importe = Abs(rs!Expr1)
'           rPrin.Update
'           Exit Do
'        End If
'        RAux.MoveNext
'     Loop
'     rs.MoveNext
'  Loop
  If rPrin.RecordCount <> 0 Then
     Report20.Database.SetDataSource rPrin
     CRViewer1.ReportSource = Report20
  End If
  Set Rs = rPrin
  Set Report20 = Nothing
End Sub

Private Sub Impre24()
  Dim Report21 As New RptRemPendDet, rPrin As Recordset, rAux As Recordset
  Dim nCanRem As Single, nCanRemAux As Single

  Dim cRem As New ClsComprobantesL
  Set rPrin = New Recordset
  Set rAux = New Recordset
  
  rPrin.Fields.Append "Comprobante", adVarChar, 30
  rPrin.Fields.Append "Fecha", adDate
  rPrin.Fields.Append "Sucursal", adInteger
  rPrin.Fields.Append "Numero", adInteger
  rPrin.Fields.Append "Cliente", adVarChar, 15
  rPrin.Fields.Append "RazonSocial", adVarChar, 50
  rPrin.Fields.Append "Producto", adVarChar, 25
  rPrin.Fields.Append "Descripcion", adVarChar, 80
  rPrin.Fields.Append "Cantidad", adSingle
  rPrin.Fields.Append "Importe", adCurrency
  rPrin.Open

  
  Set Rs = cRem.TraerInfRemitosPend(StockInfRemPendientes.LblCliente(0), StockInfRemPendientes.LblCliente(1), StockInfRemPendientes.DtpFecha(0).Value, StockInfRemPendientes.DtpFecha(1).Value)
  Do While Not Rs.EOF
     Set rAux = cRem.TraerDetallesMovimiento(Rs!Movimiento)
     Do While Not rAux.EOF
        If rAux!Producto <> "" Then
           nCanRem = Round(cRem.TraerCantidadRemito(rAux!Producto, Rs!Movimiento, rAux!ID), 2)
           nCanRemAux = Round(rAux!Cantidad - nCanRem, 2)
        Else
           rPrin.AddNew
           rPrin!Comprobante = Rs!Comprobante
           rPrin!fecha = Rs!fecha
           rPrin!Sucursal = Rs!Sucursal
           rPrin!Numero = Rs!Numero
           rPrin!Cliente = Rs!Cliente
           rPrin!RazonSocial = Rs!RazonSocial
           rPrin!Producto = rAux!Producto
           rPrin!Descripcion = rAux!Descripcion
           rPrin!Cantidad = IIf(IsNull(rAux!Cantidad), 0, rAux!Cantidad) - IIf(nCanRem = -1, 0, nCanRem)
           rPrin!importe = rAux!PrecioTotal
           rPrin.Update
           GoTo ir
        End If
        If nCanRemAux > 0 Or nCanRem = -1 Then
           rPrin.AddNew
           rPrin!Comprobante = Rs!Comprobante
           rPrin!fecha = Rs!fecha
           rPrin!Sucursal = Rs!Sucursal
           rPrin!Numero = Rs!Numero
           rPrin!Cliente = Rs!Cliente
           rPrin!RazonSocial = Rs!RazonSocial
           rPrin!Producto = rAux!Producto
           rPrin!Descripcion = rAux!Descripcion
           rPrin!Cantidad = IIf(IsNull(rAux!Cantidad), 0, rAux!Cantidad) - IIf(nCanRem = -1, 0, nCanRem)
           rPrin!importe = IIf(IsNull(rAux!PrecioTotal), 0, rAux!PrecioTotal)
           rPrin.Update
         ' Exit Do
        End If
ir:
        rAux.MoveNext
     Loop
     Rs.MoveNext
  Loop
  
  If rPrin.RecordCount <> 0 Then
     Report21.Database.SetDataSource rPrin
     CRViewer1.ReportSource = Report21
  End If
  Set Rs = rPrin
  Set Report21 = Nothing
End Sub
Private Sub Impre25()
  Dim CrxApp As New CRAXDRT.Application
  Dim CrxRpt As CRAXDRT.Report
  Set CrxRpt = CrxApp.OpenReport(App.Path & "\Factura.rpt")
  Set Rs = VentaComprobantes.rImp
  CrxRpt.Database.SetDataSource Rs
  CRViewer1.ReportSource = CrxRpt
End Sub

Private Sub Impre26()
  Dim CrxApp As New CRAXDRT.Application
  Dim CrxRpt As CRAXDRT.Report
  Set CrxRpt = CrxApp.OpenReport(App.Path & "\Remitos.rpt")
  
  Set Rs = StockComprobantes.rImp
  CrxRpt.Database.SetDataSource Rs
  CRViewer1.ReportSource = CrxRpt
End Sub

Private Sub Impre27()
  Dim Report22 As New RptVentaFamilias

  Set Rs = VentaInfFamilias.Rs
     Report22.Database.SetDataSource Rs
     Report22.ParameterFields(1).SetCurrentValue VentaInfFamilias.DtpDesde.Value
     Report22.ParameterFields(2).SetCurrentValue VentaInfFamilias.DtpHasta.Value
     Report22.ParameterFields(3).SetCurrentValue IIf(VentaInfFamilias.ChkCliente.Value = 1, True, False)
     CRViewer1.ReportSource = Report22
  Set Report22 = Nothing
End Sub

Private Sub Impre28()
  Dim Report23 As New RptPlanillaUtilidades

  Dim cRem As New ClsProductoL
  
  Set Rs = cRem.TraerPlanillaUtilidades(VentaInfFamilias.DtpDesde.Value, VentaInfFamilias.DtpHasta.Value)
  If Rs.RecordCount <> 0 Then
     Report23.Database.SetDataSource Rs
     Report23.ParameterFields(1).SetCurrentValue VentaInfFamilias.DtpDesde.Value
     Report23.ParameterFields(2).SetCurrentValue VentaInfFamilias.DtpHasta.Value
     CRViewer1.ReportSource = Report23
  End If
  Set Report23 = Nothing
End Sub

Private Sub Impre29()
  Dim Report24 As New RptControlGastos

  Dim cRem As New ClsProductoL, cRsl As ClsLectura, cImporte As Single
  
  Set cRsl = New ClsLectura
  Set Rs = cRsl.BuscarTodos("TablaDeGastos", "Descripcion", "Porcentaje")
  If Rs.RecordCount <> 0 Then
     cImporte = cRem.TraerTotalVentas(VentaInfFamilias.DtpDesde.Value, VentaInfFamilias.DtpHasta.Value)
     Report24.ParameterFields(1).SetCurrentValue str(cImporte)
     Report24.ParameterFields(2).SetCurrentValue VentaInfFamilias.DtpDesde.Value
     Report24.ParameterFields(3).SetCurrentValue VentaInfFamilias.DtpHasta.Value
     Report24.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report24
  End If
  Set Report24 = Nothing
End Sub

Private Sub Impre30()
  Dim Report25 As New RptRanking

  Dim cRem As New ClsProductoL, opt As Byte, nBus As Byte, pDesde As String, pHasta As String
  Dim sFami As String
  
  If VentaRanking.Option1(0).Value = True Then nBus = 0
  If VentaRanking.Option1(1).Value = True Then nBus = 1
  If VentaRanking.OptCodigo.Value = True Then opt = 0
  If VentaRanking.OptDescripcion.Value = True Then opt = 1
  
  pDesde = VentaRanking.TxtCodigo(0).text
  pHasta = VentaRanking.TxtCodigo(1).text
  sFami = VentaRanking.Caption
  
  If opt = 1 Then
    pDesde = VentaRanking.LblDescripcion(0).Caption
    pHasta = VentaRanking.LblDescripcion(1).Caption
  End If
  
  Set Rs = cRem.TraerRanking(pDesde, pHasta, VentaRanking.DtpDesde.Value, VentaRanking.DtpHasta.Value, nBus, opt)
  If Rs.RecordCount <> 0 Then
     Report25.ParameterFields(1).SetCurrentValue VentaRanking.DtpDesde.Value
     Report25.ParameterFields(2).SetCurrentValue VentaRanking.DtpHasta.Value
     Report25.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report25
  End If
  Set Report25 = Nothing
End Sub

Private Sub Impre31()
  Dim report1 As New RptVentaResumen

  Set Rs = CompraInfResumen.Rs
  If Rs.RecordCount <> 0 Then
     report1.ReportTitle = "Resumen de Cuenta"
     report1.ParameterFields(1).SetCurrentValue IIf(CompraInfResumen.ChkSaldo.Value = 1, True, False)
     report1.ParameterFields(2).SetCurrentValue ""
     report1.ParameterFields(3).SetCurrentValue False
     report1.ParameterFields(4).SetCurrentValue False
     report1.ParameterFields(5).SetCurrentValue ""
     report1.ParameterFields(6).SetCurrentValue ""
     
     report1.Database.SetDataSource Rs
   '  Report1.Text1.SetText ("Proveedor")
     CRViewer1.ReportSource = report1
     CRViewer1.DisplayGroupTree = False
  End If
  Set report1 = Nothing
End Sub

Private Sub Impre32()
  Dim Report2 As New RptVentaVencimiento

  Report2.ReportTitle = "Vencimientos de Pagos a Proveedores"
  Set Rs = CompraInfVencimiento.RsP
  If Rs.RecordCount <> 0 Then
     Report2.Text1.SetText ("Proveedor")
     Report2.Text13.SetText ("Pagado")
     Report2.Text14.SetText ("a Pagar")
     
     Report2.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report2
  End If
  Set Report2 = Nothing
End Sub



Private Sub Impre34()
  Dim Report32 As New RptResConDetalle
  
  Set Rs = VentaInfResumenDet.Rs
  Report32.ReportTitle = sNomFan
  Report32.ParameterFields(1).SetCurrentValue ((VentaInfResumenDet.DtpFecha(1).Value))
  Report32.ParameterFields(2).SetCurrentValue (IIf(VentaInfResumenDet.ChkCero.Value = 1, True, False))
  Report32.ParameterFields(3).SetCurrentValue VentaInfResumenDet.CmbZona.text
  If Rs.RecordCount <> 0 Then
     Report32.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report32
  End If
  Set Report32 = Nothing
End Sub

Private Sub Impre35()

End Sub

Private Sub Impre36()
  Dim Report27 As New RptResCaja
  
  Dim cRem As New ClsProductoL
  
  Set Rs = CajaResCaja.Rs
  If Rs.RecordCount <> 0 Then
     Report27.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report27
  End If
  Set Report27 = Nothing
End Sub

Private Sub Impre37()
  Dim Report28 As New RptCajaMovi
  
  Dim cRlx As New ClsComprobantesL
  
  Set Rs = cRlx.TraerMovimientoCaja(CajaInfMovimientos.DtpDesde.Value, CajaInfMovimientos.DtpHasta.Value)
  If Rs.RecordCount <> 0 Then
     Report28.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report28
  End If
  Set Report28 = Nothing
End Sub

Private Sub Impre38()
  Dim Report29 As New RptIvaVenta

  Dim cRlx As New ClsComprobantesL, Rx As ADODB.Recordset, cRsl As ClsLectura
  
  
  Set cRsl = New ClsLectura
  Set Rx = cRsl.TraerTodos("Empresa", "*")
  
  Dim sStr As String
  Set Rs = VentaLibroIva.Rs
  If Rs.RecordCount <> 0 Then
     Report29.ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
     Report29.ParameterFields(2).SetCurrentValue (Rx!Direccion)
     Report29.ParameterFields(3).SetCurrentValue (Rx!Cuit)
     Report29.ParameterFields(4).SetCurrentValue "Libro de IVA Ventas " & VentaLibroIva.TxtPeriodo
     Report29.ParameterFields(5).SetCurrentValue (Val(VentaLibroIva.TxtPagina.text))
     Report29.ParameterFields(6).SetCurrentValue IIf(VentaLibroIva.ChkAgrupado.Value = 0, False, True)
    ' Report29.Text15.SetText ("Retención Ing.Brutos") ' Abel diaz
     Report29.Text15.SetText ("Percepción Ing.Brutos")
     Report29.Text16.SetText ("Retención Ganancias")
     
     Report29.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report29
  End If
  Set Report29 = Nothing
End Sub

Private Sub Impre39()
  Dim Report29 As New RptIvaVenta

  Dim cRlx As New ClsComprobantesL, Rx As ADODB.Recordset, cRsl As ClsLectura
  
  
  Set cRsl = New ClsLectura
  Set Rx = cRsl.TraerTodos("Empresa", "*")
  
  Set Rs = ComprasLibroIva.Rs
  If Rs.RecordCount <> 0 Then
     Report29.ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
     Report29.ParameterFields(2).SetCurrentValue (Rx!Direccion)
     Report29.ParameterFields(3).SetCurrentValue (Rx!Cuit)
     Report29.ParameterFields(4).SetCurrentValue "Libro de IVA Compras " & ComprasLibroIva.TxtPeriodo
     Report29.ParameterFields(5).SetCurrentValue (Val(ComprasLibroIva.TxtPagina.text))
     Report29.ParameterFields(6).SetCurrentValue IIf(ComprasLibroIva.ChkAgrupado.Value = 0, False, True)
     
     Report29.Text15.SetText ("Percepción Ing.Brutos")
     Report29.Text16.SetText ("Percepción Iva")
     Report29.Text7.SetText ("Retención Ganancia")
     Report29.Text4.SetText ("Proveedor")
     Report29.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report29
  End If
  Set Report29 = Nothing
End Sub

Private Sub Impre40()
'  Dim cRlx As New ClsComprobantesL, Rx As ADODB.Recordset, cRsl As ClsLectura
'  Dim Rs1 As ADODB.Recordset
'
'  Set cRsl = New ClsLectura
'  Set Rx = cRsl.TraerTodos("Empresa", "*")
'
'  Set Rs = ComprasPago.rPrint2
'
'  If Rs.RecordCount <> 0 Then
'     Report30.ParameterFields(1).SetCurrentValue (Rx!Razon)
'     Report30.ParameterFields(2).SetCurrentValue (Rx!Direccion)
'     Report30.ParameterFields(3).SetCurrentValue (Rx!Localidad)
'     Report30.ParameterFields(4).SetCurrentValue (Rx!Telefono)
'     Report30.ParameterFields(5).SetCurrentValue ("Orden de Pago Número :" & ComprasPago.LblNumero.Caption)
'     Report30.ParameterFields(6).SetCurrentValue (ComprasPago.sCompr)
'     Report30.Database.SetDataSource Rs
'     CRViewer1.ReportSource = Report30
'  End If

End Sub

Private Sub Impre41()
  Dim Report31 As New RptCheques

  Dim cRlx As New ClsComprobantesL
  
  Set Rs = cRlx.TraerChequesTerceros(CajaChequesTer.DtpDesde.Value, CajaChequesTer.DtpHasta.Value, CajaChequesTer.ChkVenc.Value, CajaChequesTer.ChkCartera.Value)
  If Rs.RecordCount <> 0 Then
     Report31.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report31
  End If
  Set Report31 = Nothing
End Sub

Private Sub Impre42()
  Dim report1 As New RptVentaResumen

  Set Rs = CajaCtaCteBan.Rs
  If Rs.RecordCount <> 0 Then
     report1.ReportTitle = "Resumen de Cuenta"
     report1.ParameterFields(1).SetCurrentValue False
     report1.ParameterFields(2).SetCurrentValue ""
     report1.ParameterFields(3).SetCurrentValue False
     report1.ParameterFields(4).SetCurrentValue False
     report1.ParameterFields(5).SetCurrentValue ""
     report1.ParameterFields(6).SetCurrentValue ""
     
     report1.Database.SetDataSource Rs
   '  Report1.Text1.SetText ("Proveedor")
     CRViewer1.ReportSource = report1
     CRViewer1.DisplayGroupTree = False
  End If
  Set report1 = Nothing
End Sub

Private Sub Impre43()
 Dim Report25 As New RptRanking, cRl As ClsLectura, bIva As Boolean
 
 
  Dim cRem As New ClsProductoL, opt As Byte
  Set cRl = New ClsLectura
  If cRl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & VentaConsolidado.CmbComprobante.ItemData(VentaConsolidado.CmbComprobante.ListIndex)) = 1 Then bIva = True
  
  Set Rs = cRem.TraerConsolidado(VentaConsolidado.TxtNum(0).text, VentaConsolidado.TxtNum(1).text, VentaConsolidado.CmbComprobante.ItemData(VentaConsolidado.CmbComprobante.ListIndex), bIva)
  If Rs.RecordCount <> 0 Then
     Report25.ParameterFields(1).SetCurrentValue Date
     Report25.ParameterFields(2).SetCurrentValue Date
     Report25.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report25
  End If
  Set Report25 = Nothing
End Sub

Private Sub Impre44()
  Dim Report33 As New RptPlanCuenta

  Dim cRem As New ClsProductoL, opt As Byte
  
  
  Set Rs = CjaResCtaCble.Rs
  
  If Rs.RecordCount <> 0 Then
     Report33.ParameterFields(1).SetCurrentValue str(CjaResCtaCble.DtpDesde)
     Report33.ParameterFields(2).SetCurrentValue str(CjaResCtaCble.DtpHasta)
     Report33.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report33
  End If
  Set Report33 = Nothing
End Sub

Private Sub Impre45()

  Dim Report34 As New RptStockMovimiento
    
  Set Rs = StockInfMovimientos.Rs
  
  If Rs.RecordCount <> 0 Then
     Report34.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report34
  End If
  Set Report34 = Nothing
End Sub

Private Sub Impre46()
  Dim Report35 As New RptEtiquetas, cRe As ClsProductoL, cRsb As ClsLectura
  Dim X, Y As Byte, rel As Recordset
  
  
  Set cRe = New ClsProductoL
  Set cRsb = New ClsLectura

  Set Rs = cRe.TraerProductosEtiquetas(MtrProdEtiquetas.DtpDesde.Value, MtrProdEtiquetas.DtpHasta.Value)
  ' Generar Archivo Para la Balanza
'  Dim sNom As String, sVariable As String, nPrecio As Double
'  sNom = CurDir() & "\info.jdg"
'  Open sNom For Output As #1
'   Do While Not Rs.EOF
'      sVariable = ""
'      If cRsb.TraerValorDeUnCampo("Productos", "CodCompu", "Producto='" & Rs!Producto & "'") = "Si" And Rs!Balanza <> 0 Then
'         If Mid$(Rs!Producto, 1, 1) <> 0 Then
'            X = 2: Y = 5
'         Else
'            X = 1: Y = 5
'         End If
'         Set rel = cRsb.TraerRsCondi("Balanzas", "Balanza", "Relacion=" & Rs!Balanza)
'         Do While Not rel.EOF
'            sVariable = "C" & Format(rel!Balanza, "00") & "20050" & Mid$(Rs!Producto, X, Y) & Format(Rs!CDepartamento, "000") & Format(Rs!CFamilia, "000") & Rs!Descripcion & Space(52 - Len(Rs!Descripcion)) & Mid$(Rs!Producto, X, Y)
'            sVariable = sVariable & IIf(cRsb.TraerValorDeUnCampo("Productos", "Fijo", "Producto='" & Rs!Producto & "'") = "Si", "N", "P")
'            nPrecio = Rs!precio * 100
'            sVariable = sVariable & "0000000" & Format(nPrecio, "000000") & Format(nPrecio, "000000") & "0000000000000000000000000000010000000010000000"
'
'            Print #1, sVariable
'            rel.MoveNext
'         Loop
'      End If
'      Rs.MoveNext
'   Loop
'   Close #1
  
  If Rs.RecordCount <> 0 Then
     Report35.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report35
  End If
  Set Report35 = Nothing
  
End Sub

Private Sub Impre47()
  Dim Report35 As New RptEtiquetas
  
  
  Set Rs = MtrEtiquetas.rImp
  If Rs.RecordCount <> 0 Then
     Report35.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report35
  End If
  Set Report35 = Nothing
  
End Sub

Private Sub Impre48()
  Dim Report35 As New RptCodigoBarra
  
  
  Set Rs = MtrEtiquetas.rImp
  If Rs.RecordCount <> 0 Then
     Report35.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report35
  End If
  Set Report35 = Nothing
  
End Sub

Private Sub Impre49()
  Dim Report35 As New RptEtiquetasTango
  
  
  Set Rs = ComprasComprobantes.RsC
  
  If Rs.RecordCount <> 0 Then
     Report35.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report35
  End If
  Set Report35 = Nothing

End Sub

Private Sub Impre50()
   Dim Report35 As New RptCuentaFondos, cRsl As ClsLectura
   
  Set cRsl = New ClsLectura
  
   Set Rs = CajaCerrarZ.rPrin
  If Rs.RecordCount <> 0 Then
     Report35.Database.SetDataSource Rs
     Report35.ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
     Report35.ParameterFields(2).SetCurrentValue CajaCerrarZ.DtpFecha.Value
     
     CRViewer1.ReportSource = Report35
  End If
  Set Report35 = Nothing
   
End Sub

Private Sub Impre51()
   Dim Report35 As New RptChequesTarjetas, cRsl As ClsLectura
   
  Set cRsl = New ClsLectura
  
   Set Rs = CajaCerrarZ.rPrin
  If Rs.RecordCount <> 0 Then
     Report35.Database.SetDataSource Rs
     Report35.ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
     Report35.ParameterFields(2).SetCurrentValue CajaCerrarZ.DtpFecha.Value
     CRViewer1.ReportSource = Report35
  End If
  Set Report35 = Nothing

End Sub

Private Sub Impre52()
'   Dim Report35 As New RptCajaDiaria, cRsl As ClsLectura
'
'  Set cRsl = New ClsLectura
'
'   Set Rs = CajaCerrarZ.rPrin
'  If Rs.RecordCount <> 0 Then
'     Report35.Database.SetDataSource Rs
'     Report35.ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
'     Report35.ParameterFields(2).SetCurrentValue str(CajaCerrarZ.lblTotal.Caption)
'     CRViewer1.ReportSource = Report35
'  End If
'  Set Report35 = Nothing

End Sub

Private Sub Impre53()
  Dim Report50 As New RptListaCobranza

  Set Rs = FrmInformeCobranzas.Rs3
  If Rs.RecordCount <> 0 Then
     Report50.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report50
  End If
  Set Report50 = Nothing
End Sub

Private Sub Impre54()
  Dim Report35 As New RptEtiquetasTango
  
  
  Set Rs = MtrProductos.RsC
  
  If Rs.RecordCount <> 0 Then
     Report35.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report35
  End If
  Set Report35 = Nothing

End Sub

Private Sub Impre55()
  Dim report1 As New RptResuSimpli

  Set Rs = VentaInfResumen.Rs
  
  If Rs.RecordCount <> 0 Then
     report1.ReportTitle = sNomFan
     report1.ParameterFields(1).SetCurrentValue IIf(VentaInfResumen.ChkSaldo.Value = 1, True, False)
     report1.ParameterFields(2).SetCurrentValue VentaInfResumen.CmbZona.text
     report1.Database.SetDataSource Rs
     CRViewer1.ReportSource = report1
     CRViewer1.DisplayGroupTree = False
  End If
  Set report1 = Nothing
End Sub

Private Sub Impre56()
  Dim Report29 As New RptIvaVentaRes

  Dim cRlx As New ClsComprobantesL, Rx As ADODB.Recordset, cRsl As ClsLectura
  
  
  Set cRsl = New ClsLectura
  Set Rx = cRsl.TraerTodos("Empresa", "*")
  
  Dim sStr As String
  Set Rs = VentaLibroIva.Rs
  If Rs.RecordCount <> 0 Then
     Report29.ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
     Report29.ParameterFields(2).SetCurrentValue (Rx!Direccion)
     Report29.ParameterFields(3).SetCurrentValue "Resumen Libro de IVA Ventas " & VentaLibroIva.TxtPeriodo
     Report29.ParameterFields(4).SetCurrentValue (Rx!Cuit)
     
     Report29.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report29
  End If
  Set Report29 = Nothing
End Sub

Private Sub Impre57()
  Dim Report29 As New RptIvaVentaRes

  Dim cRlx As New ClsComprobantesL, Rx As ADODB.Recordset, cRsl As ClsLectura
  
  
  Set cRsl = New ClsLectura
  Set Rx = cRsl.TraerTodos("Empresa", "*")
  
  Set Rs = ComprasLibroIva.Rs
  If Rs.RecordCount <> 0 Then
     Report29.ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
     Report29.ParameterFields(2).SetCurrentValue (Rx!Direccion)
     Report29.ParameterFields(3).SetCurrentValue "Libro de IVA Compras " & ComprasLibroIva.TxtPeriodo
     Report29.ParameterFields(4).SetCurrentValue (Rx!Cuit)
     Report29.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report29
  End If
  Set Report29 = Nothing
End Sub

Private Sub Impre58()
 Dim Report25 As New RptReparto, cRl As ClsLectura, bIva As Boolean
 
 
  Dim cRem As New ClsProductoL, opt As Byte
  Set cRl = New ClsLectura
  
  
  Set Rs = cRem.TraerReparto(Val(VentaReparto.TxtReparto.text), VentaReparto.DtpDesde.Value, VentaReparto.DtpHasta.Value)
  If Rs.RecordCount <> 0 Then
     Report25.ParameterFields(1).SetCurrentValue Val(VentaReparto.TxtReparto.text)
     Report25.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report25
  End If
  Set Report25 = Nothing
End Sub

Private Sub Impre59()
  Dim report1 As New RptResuSimpli

  Set Rs = CompraInfResumen.Rs
  
  If Rs.RecordCount <> 0 Then
     report1.ReportTitle = "Resumen de Cuenta Simplificado - Proveedores"
     report1.ParameterFields(1).SetCurrentValue IIf(CompraInfResumen.ChkSaldo.Value = 1, True, False)
     report1.ParameterFields(2).SetCurrentValue "Todas"
     report1.Database.SetDataSource Rs
     CRViewer1.ReportSource = report1
     CRViewer1.DisplayGroupTree = False
  End If
  Set report1 = Nothing
End Sub

Private Sub Impre60()
  Dim Report33 As New RptPlanCtaDet

  Dim cRem As New ClsProductoL, opt As Byte
  
  
  Set Rs = CjaResCtaCble.Rs
  
  If Rs.RecordCount <> 0 Then
     Report33.ParameterFields(1).SetCurrentValue str(CjaResCtaCble.DtpDesde)
     Report33.ParameterFields(2).SetCurrentValue str(CjaResCtaCble.DtpHasta)
     Report33.ParameterFields(3).SetCurrentValue CjaResCtaCble.CmbCuenta.text
     Report33.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report33
  End If
  Set Report33 = Nothing
End Sub

Private Sub Impre61()
  Dim Report33 As New RptCedronar

  Dim cRem As New ClsProductoL, opt As Byte
  
  
  Set Rs = StockInfMovimientos.Rs
  
  If Rs.RecordCount <> 0 Then
     Report33.ParameterFields(1).SetCurrentValue StockInfMovimientos.DtpFecha(0).Value
     Report33.ParameterFields(2).SetCurrentValue StockInfMovimientos.DtpFecha(1).Value
     Report33.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report33
  End If
  Set Report33 = Nothing
End Sub

Private Sub Impre62()
  Dim Report28 As New RptInfPerIB
  
  Dim cRlx As New ClsComprobantesL
  
  Set Rs = cRlx.TraerInPercIB(VentasInfPerIb.DtpDesde.Value, VentasInfPerIb.DtpHasta.Value)
  If Rs.RecordCount <> 0 Then
     Report28.ParameterFields(1).SetCurrentValue str(VentasInfPerIb.DtpDesde.Value)
     Report28.ParameterFields(2).SetCurrentValue str(VentasInfPerIb.DtpHasta.Value)
     Report28.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report28
  End If
  Set Report28 = Nothing
End Sub

Private Sub Impre63()
  Dim Report28 As New RtpInfRetIb
  
  Dim cRlx As New ClsComprobantesL
  
  Set Rs = cRlx.TraerInfRetIB(ComprasInfRetIb.DtpDesde.Value, ComprasInfRetIb.DtpHasta.Value)
  If Rs.RecordCount <> 0 Then
     Report28.ParameterFields(1).SetCurrentValue str(ComprasInfRetIb.DtpDesde.Value)
     Report28.ParameterFields(2).SetCurrentValue str(ComprasInfRetIb.DtpHasta.Value)
     Report28.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report28
  End If
  Set Report28 = Nothing
End Sub

Private Sub Impre64()
  Dim Report33 As New RptR122V2
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Set Rs = ComprasAnularPagos.Rs
  
  If Rs.RecordCount <> 0 Then
     Report33.ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Cuit", "Id=" & nMenu - 1)
     Report33.ParameterFields(2).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "Id=" & nMenu - 1)
     Report33.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report33
  End If
  Set Report33 = Nothing
End Sub

Private Sub Impre65()
  Dim report1 As New RptInfAplica

  Set Rs = VentaInfResumen.Rs
  
  If Rs.RecordCount <> 0 Then
     report1.ReportTitle = sNomFan
     report1.ParameterFields(1).SetCurrentValue str(VentaInfResumen.DtpFecha(0).Value)
     report1.ParameterFields(2).SetCurrentValue str(VentaInfResumen.DtpFecha(1).Value)
     report1.Database.SetDataSource Rs
     CRViewer1.ReportSource = report1
     CRViewer1.DisplayGroupTree = False
  End If
  Set report1 = Nothing
End Sub

Private Sub Impre66()
  Dim report1 As New RptProyeccion

  Set Rs = VentasInfProy.Rs
  
  If Rs.RecordCount <> 0 Then
     report1.ReportTitle = sNomFan
     report1.ParameterFields(1).SetCurrentValue str(VentasInfProy.DtpFecha(0).Value)
     report1.ParameterFields(2).SetCurrentValue str(VentasInfProy.DtpFecha(1).Value)
     report1.ParameterFields(3).SetCurrentValue str(VentasInfProy.DtpFecha(2).Value)
     report1.ParameterFields(4).SetCurrentValue str(VentasInfProy.DtpFecha(3).Value)
     report1.ParameterFields(5).SetCurrentValue str(VentasInfProy.DtpFecha(4).Value)
     
     report1.Database.SetDataSource Rs
     CRViewer1.ReportSource = report1
     CRViewer1.DisplayGroupTree = False
  End If
  Set report1 = Nothing
End Sub

Private Sub Impre67()
  Dim Report33 As New RptCertGanancia
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Set Rs = ComprasAnularPagos.Rs
  
  If Rs.RecordCount <> 0 Then
     
     Report33.ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "Id=" & nMenu - 1)
     Report33.ParameterFields(2).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Direccion", "Id=" & nMenu - 1)
     Report33.ParameterFields(3).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Localidad", "Id=" & nMenu - 1)
     Report33.ParameterFields(4).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Cuit", "Id=" & nMenu - 1)
     
     Report33.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report33
  End If
  Set Report33 = Nothing
  
End Sub

Private Sub Impre68()
  Dim Report4 As New RptStockSaldo, bOrd As Integer
  Dim nStock As Byte

  Set cCanti = New ClsProductoL
              
  pLista = StockInfSaldos.CmbDeposito.ItemData(StockInfSaldos.CmbDeposito.ListIndex)
  
     bOrd = 3
     sDesde = StockInfSaldos.LblDescripcion(0).Caption
     sHasta = StockInfSaldos.LblDescripcion(1).Caption
  
  nStock = StockInfSaldos.ChkSaldo.Value
  
  
  Set Rs = cCanti.DameProductos(pLista, sDesde, sHasta, bOrd, nStock)
  If Rs.RecordCount <> 0 Then
     Report4.ParameterFields(1).SetCurrentValue StockInfSaldos.CmbDeposito.text
     Report4.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report4
  End If
  Set Report4 = Nothing
End Sub

Private Sub Impre69()
  Dim Report4 As New RptRepartoRemi, pZona As String
              
  pZona = VentaRepartoRemito.CmbZona.text
  
  sDesde = VentaRepartoRemito.DtpDesde.Value
  sHasta = VentaRepartoRemito.DtpHasta.Value
  
  Set Rs = VentaRepartoRemito.Rs
  
  If Rs.RecordCount <> 0 Then
     Report4.ParameterFields(1).SetCurrentValue sDesde
     Report4.ParameterFields(2).SetCurrentValue sHasta
     Report4.ParameterFields(3).SetCurrentValue pZona
     Report4.ParameterFields(4).SetCurrentValue VentaRepartoRemito.CmbEstado.text
     
     Report4.Database.SetDataSource Rs
     CRViewer1.ReportSource = Report4
  End If
  Set Report4 = Nothing
End Sub

