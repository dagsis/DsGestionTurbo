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

Dim ReporteActivo As Object

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    On Error GoTo ManejoErrores
    UseDefault = False

    If ReporteActivo Is Nothing Then Exit Sub

    Dim prn As String
    Dim f As Long, t As Long, copies As Long
    Dim collate As Boolean, useRange As Boolean

    If Not ShowWindowsPrintDialog(Me.hwnd, prn, f, t, copies, collate, useRange) Then
       Exit Sub
    End If


    ' Mapear a objeto Printer para obtener Driver/Port confiables
    Dim p As Printer, found As Boolean
    For Each p In Printers
        If StrComp(p.DeviceName, prn, vbTextCompare) = 0 Then
            Set Printer = p
            found = True
            Exit For
        End If
    Next

    ' Crystal: conviene winspool + nombre
    If found Then
        ReporteActivo.SelectPrinter "winspool", Printer.DeviceName, Printer.Port
    Else
        ReporteActivo.SelectPrinter "winspool", prn, ""
    End If

    On Error Resume Next
    ReporteActivo.PrintOptions.NumberOfCopies = copies
    On Error GoTo ManejoErrores

    If useRange And f > 0 And t >= f Then
        ReporteActivo.PrintOut False, , , CInt(f), CInt(t)
    Else
        ReporteActivo.PrintOut False
    End If

    Exit Sub

ManejoErrores:
    If err.Number <> 32755 Then
        MsgBox "Error de impresión: " & err.Description, vbCritical
    End If
    err.Clear
End Sub


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
        ' Configuración normal del visor
        With CRViewer1
            .EnablePrintButton = True
            .DisplayGroupTree = False
            .EnableExportButton = True
            .EnableRefreshButton = False
            .Zoom 100
            .ViewReport
        End With
    Else
        ' No hay registros
        MsgBox "No hay Registros. Intente con Nuevos Parametros", vbCritical, "Atención"
        
        ' 1. Limpiamos puntero
        Screen.MousePointer = vbDefault
        
        ' 2. Descargamos el formulario de forma segura
        ' Usamos un Timer o simplemente Unload Me asegurándonos de que sea la última instrucción real
    '    Unload Me
        Exit Sub
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
  CRViewer1.Top = 0
  CRViewer1.Left = 0
  CRViewer1.Height = ScaleHeight
  CRViewer1.Width = ScaleWidth
End Sub

Private Sub Impre01()
  Dim cRsl As New ClsLectura
  Set Rs = VentaInfResumen.Rs
    
  If Rs.RecordCount <> 0 Then
        ' Usamos la variable global
        Set ReporteActivo = New RptVentaResumen
        
        ReporteActivo.ReportTitle = sNomFan
        ReporteActivo.ParameterFields(1).SetCurrentValue IIf(VentaInfResumen.ChkSaldo.Value = 1, True, False)
        ReporteActivo.ParameterFields(2).SetCurrentValue VentaInfResumen.CmbZona.Text
        ReporteActivo.ParameterFields(3).SetCurrentValue IIf(VentaInfResumen.ChkTexto.Value = 1, True, False)
        ReporteActivo.ParameterFields(4).SetCurrentValue IIf(VentaInfResumen.ChkSalto.Value = 1, True, False)
        ReporteActivo.ParameterFields(5).SetCurrentValue cRsl.TraerValorDeUnCampo("TblTextos", "Texto", "Id=1")
        ReporteActivo.ParameterFields(6).SetCurrentValue cRsl.TraerValorDeUnCampo("TblTextos", "Texto", "Id=2")
             
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre02()
  If VentaInfVencimiento.RsP.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentaVencimiento
        ReporteActivo.ReportTitle = "Vencimiento por Zona :" & VentaInfVencimiento.CmbZona.Text
        
        Set Rs = VentaInfVencimiento.RsP
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre03()
    ' Eliminamos los "As New" y usamos la variable global ReporteActivo
    ' bOrd se queda como variable local
    Dim bOrd As Integer

    Set cClases = New ClsPrecios
    pLista = VentaInfPrecio.CmbLista.ItemData(VentaInfPrecio.CmbLista.ListIndex)

    ' Lógica de ordenamiento
    If VentaInfPrecio.OptCodigo = True Then
        bOrd = 0
        sDesde = VentaInfPrecio.TxtCodigo(0).Text
        sHasta = VentaInfPrecio.TxtCodigo(1).Text
    Else
        bOrd = 1
        sDesde = VentaInfPrecio.LblDescripcion(0).Caption
        sHasta = VentaInfPrecio.LblDescripcion(1).Caption
    End If

    sPdesde = VentaInfPrecio.CmbDesde.Text
    sPhasta = VentaInfPrecio.CmbHasta.Text

    ' Selección del tipo de Reporte
    If VentaInfPrecio.OptNinguno = True Then
        ' 1. Reporte Estándar
        Set Rs = cClases.DameListaProductos(pLista, sDesde, sHasta, bOrd)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPrecios ' <--- Asignamos a la global
            ReporteActivo.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.Text
        End If
        
    ElseIf VentaInfPrecio.OptMarcas = True Then
        ' 2. Reporte por Marcas
        Set Rs = cClases.DameListaProductosMarca(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPreciosMarca ' <--- Asignamos a la global
            ReporteActivo.ReportTitle = "Listado de Precios Agrupados por Marca"
            ReporteActivo.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.Text
        End If
        
    Else
        ' 3. Reporte por Grupos
        Set Rs = cClases.DameListaProductosGrupo(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPreciosGrupo ' <--- Asignamos a la global
            ReporteActivo.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.Text
        End If
    End If

    ' Si se asignó un reporte y hay datos, lo mostramos
    If Not ReporteActivo Is Nothing Then
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If

    ' IMPORTANTE: Borramos los "Set ReportX = Nothing" de aquí.
    ' El reporte debe quedar vivo para que el visor pueda imprimirlo luego.
End Sub

Private Sub Impre04()
    ' Eliminamos los Dims de reportes internos y el bOrd lo dejamos local
    Dim bOrd As Integer

    Set cClases = New ClsPrecios
    pLista = VentaInfPrecio.CmbLista.ItemData(VentaInfPrecio.CmbLista.ListIndex)
  
    ' Lógica de filtrado
    If VentaInfPrecio.OptCodigo = True Then
        bOrd = 0
        sDesde = VentaInfPrecio.TxtCodigo(0).Text
        sHasta = VentaInfPrecio.TxtCodigo(1).Text
    Else
        bOrd = 1
        sDesde = VentaInfPrecio.LblDescripcion(0).Caption
        sHasta = VentaInfPrecio.LblDescripcion(1).Caption
    End If
    
    sPdesde = VentaInfPrecio.CmbDesde.Text
    sPhasta = VentaInfPrecio.CmbHasta.Text
    
    ' Selección de Reporte y Carga de Datos
    If VentaInfPrecio.OptNinguno.Value = True Then
        Set Rs = cClases.DameListaProductos(pLista, sDesde, sHasta, bOrd)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPrecios ' <--- Variable Global
            ReporteActivo.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.Text
        End If
    Else
        If VentaInfPrecio.OptMarcas.Value = True Then
            Set Rs = cClases.DameListaProductosMarca(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
            If Rs.RecordCount <> 0 Then
                Set ReporteActivo = New RptVentaPreciosMarca ' <--- Variable Global
                ReporteActivo.ReportTitle = "Listado de Precios Agrupados por Marca"
                ReporteActivo.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.Text
            End If
        Else
            Set Rs = cClases.DameListaProductosGrupo(pLista, sDesde, sHasta, bOrd, sPdesde, sPhasta)
            If Rs.RecordCount <> 0 Then
                Set ReporteActivo = New RptVentaPreciosGrupo ' <--- Variable Global
                ReporteActivo.ParameterFields(1).SetCurrentValue VentaInfPrecio.CmbLista.Text
            End If
        End If
    End If

    ' Asignación al Visor
    If Not ReporteActivo Is Nothing Then
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If

    ' Se eliminaron los Set ... = Nothing para que el reporte no muera
End Sub

Private Sub Impre05()
    ' Eliminamos el Dim Report4 As New...
    Dim bSto As Byte, bOrd As Integer

    Set cCanti = New ClsProductoL
    
    ' Capturamos parámetros de la pantalla
    pLista = StockInfSaldos.CmbDeposito.ItemData(StockInfSaldos.CmbDeposito.ListIndex)
    
    If StockInfSaldos.OptCodigo = True Then
        bOrd = 0
        sDesde = StockInfSaldos.TxtCodigo(0).Text
        sHasta = StockInfSaldos.TxtCodigo(1).Text
    Else
        bOrd = 1
        sDesde = StockInfSaldos.LblDescripcion(0).Caption
        sHasta = StockInfSaldos.LblDescripcion(1).Caption
    End If
    
    bSto = StockInfSaldos.ChkSaldo.Value
    
    ' Obtenemos los datos
    Set Rs = cCanti.DameProductos(pLista, sDesde, sHasta, bOrd, bSto)
    
    ' Si hay datos, creamos el reporte en la variable global
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptStockSaldo
        ReporteActivo.ParameterFields(1).SetCurrentValue StockInfSaldos.CmbDeposito.Text
        ReporteActivo.Database.SetDataSource Rs
        
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' No ponemos el Set ReporteActivo = Nothing para que no se borre de la vista
End Sub

Private Sub Impre06()
    ' Variables locales de configuración
    Dim bOrd As Integer
    Dim nStock As Byte, pFamilia As Integer

    Set cCanti = New ClsProductoL
    
    ' Obtenemos IDs de los combos
    pLista = StockInfSaldos.CmbDeposito.ItemData(StockInfSaldos.CmbDeposito.ListIndex)
    pFamilia = StockInfSaldos.CmbFamilia.ItemData(StockInfSaldos.CmbFamilia.ListIndex)
    
    ' Lógica de orden (Código vs Descripción)
    If StockInfSaldos.OptCodigo = True Then
        bOrd = 0
        sDesde = StockInfSaldos.TxtCodigo(0).Text
        sHasta = StockInfSaldos.TxtCodigo(1).Text
    Else
        bOrd = 1
        sDesde = StockInfSaldos.LblDescripcion(0).Caption
        sHasta = StockInfSaldos.LblDescripcion(1).Caption
    End If
    
    nStock = StockInfSaldos.ChkSaldo.Value
    
    ' Ejecutamos la consulta a la base de datos
    Set Rs = cCanti.DameProductos(pLista, sDesde, sHasta, bOrd, nStock, pFamilia)
    
    ' Si hay datos, cargamos el reporte global
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptStockSaldo ' Usamos el mismo diseño que el 05
        
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue StockInfSaldos.CmbDeposito.Text
            .ParameterFields(2).SetCurrentValue StockInfSaldos.CmbFamilia.Text
            .Database.SetDataSource Rs
        End With
        
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' Recordá: No destruir ReporteActivo aquí
End Sub

Private Sub Impre07()
    Dim pFamilias As Integer
    Set cCanti = New ClsProductoL
    
    pFamilias = StockInfSaldos.CmbFamilia.ItemData(StockInfSaldos.CmbFamilia.ListIndex)
    
    sDesde = StockInfReposiciones.TxtCodigo(0)
    sHasta = StockInfReposiciones.TxtCodigo(1)
    sPdesde = StockInfReposiciones.CmbProveedor(0).Text
    
    ' Ejecutamos la consulta
    Set Rs = cCanti.DameReposiciones(sDesde, sHasta, Val(sPdesde), 1)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptStockReposicion
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre08()
    Set cCanti = New ClsProductoL
    
    ' Cambiamos a etiquetas de descripción
    sDesde = StockInfReposiciones.LblDescripcion(0)
    sHasta = StockInfReposiciones.LblDescripcion(1)
    sPdesde = StockInfReposiciones.CmbProveedor(0).ItemData(StockInfReposiciones.CmbProveedor(0).ListIndex)
    
    ' Tipo de búsqueda 2 (Descripción)
    Set Rs = cCanti.DameReposiciones(sDesde, sHasta, Val(sPdesde), 2)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptStockReposicion
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre09()

End Sub

Private Sub Impre10()
    Set Rs = VentaInfClientes.Rs
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptClienteSinZona
        With ReporteActivo
            .ReportTitle = "Listado de Clientes"
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre11()
    Set Rs = VentaInfClientes.Rs
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptClientesZona
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre12()
    Set cClases = New ClsPrecios
    
    sDesde = StockInfCostos.LblDescripcion(0)
    sHasta = StockInfCostos.LblDescripcion(1)
    sPdesde = StockInfCostos.CmbDesde.Text
    sPhasta = StockInfCostos.CmbHasta.Text
    
    If StockInfCostos.OptMarcas = True Then
        Set Rs = cClases.DameCostoProductos(sDesde, sHasta, 3)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPrecios
            ReporteActivo.ReportTitle = "Listado de Costos"
            ReporteActivo.ParameterFields(1).SetCurrentValue ""
        End If
    ElseIf StockInfCostos.OptFamilias = True Then
        Set Rs = cClases.DameCostoProductosMarca(sDesde, sHasta, 3, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPreciosMarca
            ReporteActivo.ReportTitle = "Listado de Costos Agrupados por Marca"
            ReporteActivo.ParameterFields(1).SetCurrentValue ""
        End If
    Else
        Set Rs = cClases.DameCostoProductosGrupo(sDesde, sHasta, 3, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPreciosGrupo
            ReporteActivo.ReportTitle = "Listado de Costos Agrupados por Familias"
            ReporteActivo.ParameterFields(1).SetCurrentValue ""
        End If
    End If

    ' Si se creó algún reporte, le pasamos los datos
    If Not ReporteActivo Is Nothing Then
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre13()
    Set cClases = New ClsPrecios
    
    sDesde = StockInfCostos.LblDescripcion(0)
    sHasta = StockInfCostos.LblDescripcion(1)
    sPdesde = StockInfCostos.CmbDesde.Text
    sPhasta = StockInfCostos.CmbHasta.Text
    
    If StockInfCostos.OptMarcas.Value = True Then
        Set Rs = cClases.DameCostoProductos(sDesde, sHasta, 4)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPrecios
            ReporteActivo.ParameterFields(1).SetCurrentValue ""
        End If
    ElseIf StockInfCostos.OptFamilias = True Then
        Set Rs = cClases.DameCostoProductosMarca(sDesde, sHasta, 4, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPreciosMarca
            ReporteActivo.ReportTitle = "Listado de Costos Agrupados por Marca"
            ReporteActivo.ParameterFields(1).SetCurrentValue ""
        End If
    Else
        Set Rs = cClases.DameCostoProductosGrupo(sDesde, sHasta, 4, sPdesde, sPhasta)
        If Rs.RecordCount <> 0 Then
            Set ReporteActivo = New RptVentaPreciosGrupo
            ReporteActivo.ReportTitle = "Listado de Costos Agrupados por Familias"
            ReporteActivo.ParameterFields(1).SetCurrentValue ""
        End If
    End If

    If Not ReporteActivo Is Nothing Then
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre14()
    Set cCaja = New ClsComprobantesL
    
    fDesde = VentaInfVentas.DTPicker1.Value
    fHasta = VentaInfVentas.DTPicker2.Value
    
    Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentasFecha
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre15()
    Set cCaja = New ClsComprobantesL
    fDesde = VentaInfVentas.DTPicker1.Value
    fHasta = VentaInfVentas.DTPicker2.Value
    
    Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentasComprobantes
        With ReporteActivo
            .ReportTitle = "Comprobantes de Ventas"
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre16()
    Set cCaja = New ClsComprobantesL
    fDesde = VentaInfVentas.DTPicker1.Value
    fHasta = VentaInfVentas.DTPicker2.Value
    
    Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentasVendedor
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre17()
    Set cCaja = New ClsComprobantesL
    fDesde = VentaInfVentas.DTPicker1.Value
    fHasta = VentaInfVentas.DTPicker2.Value
    
    Set Rs = cCaja.TraerMoviVentas(fDesde, fHasta)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentasCajas
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre18()
    ' Este reporte toma el recordset directamente del formulario de origen
    Set Rs = VentaInfVentas.Rs
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentasProductos
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre19()
    Dim cRle As New ClsClienteL
    Dim nOrd As Byte
    
    ' Definimos el orden y los rangos según los controles del formulario
    If VentaInfDetalles.OptCodigo.Value = True Then
        nOrd = 0
        sDesde = VentaInfDetalles.TxtCliente(0).Text
        sHasta = VentaInfDetalles.TxtCliente(1).Text
    Else
        nOrd = 1
        sDesde = VentaInfDetalles.LblCliente(0).Caption
        sHasta = VentaInfDetalles.LblCliente(1).Caption
    End If
    
    ' Traemos los datos pasando fechas y orden
    Set Rs = cRle.TraerDetalleFactura(sDesde, sHasta, VentaInfDetalles.DtpFecha(0).Value, VentaInfDetalles.DtpFecha(1).Value, nOrd)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptDetalleFactura
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre20()
    Dim crField As CRAXDRT.DatabaseFieldDefinition
    Dim pGru As Byte, pDesde As String, pHasta As String
    
    Set cCanti = New ClsProductoL
    
    ' Tomamos los valores del formulario de stock
    pGru = StockInfProductos.nOrd
    pDesde = StockInfProductos.sDesde
    pHasta = StockInfProductos.sHasta
    
    Set Rs = cCanti.ProductoListadoMtr(StockInfProductos.LblProducto(0).Caption, _
                StockInfProductos.LblProducto(1).Caption, _
                StockInfProductos.ChkAlfabetico.Value, pGru, pDesde, pHasta)
                
    If pGru = 0 Then
        If StockInfProductos.ChkAlfabetico.Value = 0 Then pGru = 1
        If StockInfProductos.ChkAlfabetico.Value = 1 Then pGru = 2
    End If
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos el reporte dinámico
        Set ReporteActivo = New RptProductoGene
        
        ' Buscamos el campo por el cual se va a agrupar
        Set crField = ReporteActivo.Database.Tables.Item(1).Fields.Item(pGru)
        
        ' Asignamos el campo a la sección del Grupo (GH = Group Header)
        ReporteActivo.Areas.Item("GH").GroupConditionField = crField
        ReporteActivo.Database.SetDataSource Rs
        ReporteActivo.Areas.Item("GH").Suppress = False
        
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre21()
    Dim bOrd As Integer
    Dim cCl As New ClsProductoL
    
    If VentaInfCosto.OptCodigo = True Then
        bOrd = 0
    Else
        bOrd = 1
    End If
    
    sDesde = VentaInfCosto.LblDescripcion(0).Caption
    sHasta = VentaInfCosto.LblDescripcion(1).Caption
    
    Set Rs = cCl.TraerValoInventario(sDesde, sHasta, bOrd)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptValoCosto
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre22()
    Dim cCl As New ClsProductoL
    
    ' Ejecutamos la consulta de rentabilidad pasando fechas y descripciones
    Set Rs = cCl.TraerRentabilidad(VentaProducCosto.DtpDesde.Value, _
                                   VentaProducCosto.DtpHasta.Value, _
                                   VentaProducCosto.LblDescripcion(0).Caption, _
                                   VentaProducCosto.LblDescripcion(1).Caption)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptRentabilid
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre23()
    Dim rPrin As New ADODB.Recordset
    Dim rAux As ADODB.Recordset
    Dim nCantRem As Double, nCantRemAux As Double
    Dim cRem As New ClsComprobantesL

    ' 1. Creamos el Recordset temporal en memoria
    rPrin.Fields.Append "Comprobante", adVarChar, 30
    rPrin.Fields.Append "Fecha", adDate
    rPrin.Fields.Append "Sucursal", adInteger
    rPrin.Fields.Append "Numero", adInteger
    rPrin.Fields.Append "Cliente", adVarChar, 15
    rPrin.Fields.Append "RazonSocial", adVarChar, 50
    rPrin.Fields.Append "Importe", adCurrency
    rPrin.Open
    
    ' 2. Traemos la información base
    Set Rs = cRem.TraerInfRemitosPend(StockInfRemPendientes.LblCliente(0), _
                                     StockInfRemPendientes.LblCliente(1), _
                                     StockInfRemPendientes.DtpFecha(0).Value, _
                                     StockInfRemPendientes.DtpFecha(1).Value)

    ' 3. Bucle de procesamiento de saldos
    Do While Not Rs.EOF
        Set rAux = cRem.TraerDetallesMovimiento(Rs!Movimiento)
        Do While Not rAux.EOF
            nCantRem = Round(cRem.TraerCantidadRemito(rAux!Producto, Rs!Movimiento, rAux!ID), 2)
            nCantRemAux = Round(rAux!Cantidad - Abs(nCantRem), 2)
            
            If nCantRemAux > 0 Or nCantRem = -1 Then
                rPrin.AddNew
                rPrin!Comprobante = Rs!Comprobante
                rPrin!fecha = Rs!fecha
                rPrin!sucursal = Rs!sucursal
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
    
    ' 4. Cargamos el reporte con los datos calculados
    If rPrin.RecordCount <> 0 Then
        Set ReporteActivo = New RptRemPendientes
        ReporteActivo.Database.SetDataSource rPrin
        CRViewer1.ReportSource = ReporteActivo
        Set Rs = rPrin ' Asignamos a la global para que Form_Load lo vea
    End If
End Sub


Private Sub Impre24()
    Dim rPrin As New ADODB.Recordset
    Dim rAux As ADODB.Recordset
    Dim nCanRem As Single, nCanRemAux As Single
    Dim cRem As New ClsComprobantesL

    ' 1. Estructura del Recordset temporal
    With rPrin.Fields
        .Append "Comprobante", adVarChar, 30
        .Append "Fecha", adDate
        .Append "Sucursal", adInteger
        .Append "Numero", adInteger
        .Append "Cliente", adVarChar, 15
        .Append "RazonSocial", adVarChar, 50
        .Append "Producto", adVarChar, 25
        .Append "Descripcion", adVarChar, 80
        .Append "Cantidad", adSingle
        .Append "Importe", adCurrency
    End With
    rPrin.Open

    ' 2. Traemos datos base
    Set Rs = cRem.TraerInfRemitosPend(StockInfRemPendientes.LblCliente(0), _
                                     StockInfRemPendientes.LblCliente(1), _
                                     StockInfRemPendientes.DtpFecha(0).Value, _
                                     StockInfRemPendientes.DtpFecha(1).Value)

    ' 3. Procesamiento detallado
    Do While Not Rs.EOF
        Set rAux = cRem.TraerDetallesMovimiento(Rs!Movimiento)
        Do While Not rAux.EOF
            If rAux!Producto <> "" Then
                nCanRem = Round(cRem.TraerCantidadRemito(rAux!Producto, Rs!Movimiento, rAux!ID), 2)
                nCanRemAux = Round(rAux!Cantidad - nCanRem, 2)
            Else
                ' Si no tiene producto (ej. un concepto), lo agregamos directo
                GoTo AgregarRegistro
            End If

            If nCanRemAux > 0 Or nCanRem = -1 Then
AgregarRegistro:
                rPrin.AddNew
                rPrin!Comprobante = Rs!Comprobante
                rPrin!fecha = Rs!fecha
                rPrin!sucursal = Rs!sucursal
                rPrin!Numero = Rs!Numero
                rPrin!Cliente = Rs!Cliente
                rPrin!RazonSocial = Rs!RazonSocial
                rPrin!Producto = rAux!Producto
                rPrin!Descripcion = rAux!Descripcion
                rPrin!Cantidad = IIf(IsNull(rAux!Cantidad), 0, rAux!Cantidad) - IIf(nCanRem = -1, 0, nCanRem)
                rPrin!importe = IIf(IsNull(rAux!PrecioTotal), 0, rAux!PrecioTotal)
                rPrin.Update
            End If
            rAux.MoveNext
        Loop
        Rs.MoveNext
    Loop
    
    ' 4. Carga final
    If rPrin.RecordCount <> 0 Then
        Set ReporteActivo = New RptRemPendDet
        ReporteActivo.Database.SetDataSource rPrin
        CRViewer1.ReportSource = ReporteActivo
        Set Rs = rPrin
    End If
End Sub

Private Sub Impre25()
    Dim CrxApp As New CRAXDRT.Application
    
    ' Abrimos el archivo .rpt desde la carpeta del programa
    Set ReporteActivo = CrxApp.OpenReport(App.Path & "\Factura.rpt")
    
    Set Rs = VentaComprobantes.rImp
    
    If Not Rs Is Nothing Then
        If Rs.RecordCount <> 0 Then
            ReporteActivo.Database.SetDataSource Rs
            CRViewer1.ReportSource = ReporteActivo
        End If
    End If
End Sub

Private Sub Impre26()
    Dim CrxApp As New CRAXDRT.Application
    
    ' Abrimos el archivo .rpt desde la carpeta del programa
    Set ReporteActivo = CrxApp.OpenReport(App.Path & "\Remitos.rpt")
    
    Set Rs = StockComprobantes.rImp
    
    If Not Rs Is Nothing Then
        If Rs.RecordCount <> 0 Then
            ReporteActivo.Database.SetDataSource Rs
            CRViewer1.ReportSource = ReporteActivo
        End If
    End If
End Sub

Private Sub Impre27()
    Set Rs = VentaInfFamilias.Rs
    If Rs Is Nothing Then Exit Sub
    
    Set ReporteActivo = New RptVentaFamilias
    With ReporteActivo
        .Database.SetDataSource Rs
        .ParameterFields(1).SetCurrentValue VentaInfFamilias.DtpDesde.Value
        .ParameterFields(2).SetCurrentValue VentaInfFamilias.DtpHasta.Value
        .ParameterFields(3).SetCurrentValue IIf(VentaInfFamilias.ChkCliente.Value = 1, True, False)
    End With
    CRViewer1.ReportSource = ReporteActivo
End Sub

Private Sub Impre28()
    Dim cRem As New ClsProductoL
    Set Rs = cRem.TraerPlanillaUtilidades(VentaInfFamilias.DtpDesde.Value, VentaInfFamilias.DtpHasta.Value)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptPlanillaUtilidades
        With ReporteActivo
            .Database.SetDataSource Rs
            .ParameterFields(1).SetCurrentValue VentaInfFamilias.DtpDesde.Value
            .ParameterFields(2).SetCurrentValue VentaInfFamilias.DtpHasta.Value
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre29()
    Dim cRem As New ClsProductoL, cRsl As New ClsLectura, cImporte As Single
    Set Rs = cRsl.BuscarTodos("TablaDeGastos", "Descripcion", "Porcentaje")
    
    If Rs.RecordCount <> 0 Then
        cImporte = cRem.TraerTotalVentas(VentaInfFamilias.DtpDesde.Value, VentaInfFamilias.DtpHasta.Value)
        Set ReporteActivo = New RptControlGastos
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue str(cImporte)
            .ParameterFields(2).SetCurrentValue VentaInfFamilias.DtpDesde.Value
            .ParameterFields(3).SetCurrentValue VentaInfFamilias.DtpHasta.Value
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre30()
    Dim cRem As New ClsProductoL, opt As Byte, nBus As Byte
    Dim pDesde As String, pHasta As String
    
    ' Determinamos criterios de búsqueda
    If VentaRanking.Option1(0).Value = True Then nBus = 0 Else nBus = 1
    If VentaRanking.OptCodigo.Value = True Then opt = 0 Else opt = 1
    
    If opt = 0 Then
        pDesde = VentaRanking.TxtCodigo(0).Text
        pHasta = VentaRanking.TxtCodigo(1).Text
    Else
        pDesde = VentaRanking.LblDescripcion(0).Caption
        pHasta = VentaRanking.LblDescripcion(1).Caption
    End If
    
    Set Rs = cRem.TraerRanking(pDesde, pHasta, VentaRanking.DtpDesde.Value, VentaRanking.DtpHasta.Value, nBus, opt)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptRanking
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue VentaRanking.DtpDesde.Value
            .ParameterFields(2).SetCurrentValue VentaRanking.DtpHasta.Value
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre31()
    Set Rs = CompraInfResumen.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentaResumen ' Reutilizamos el diseño de resumen
        With ReporteActivo
            .ReportTitle = "Resumen de Cuenta - Proveedores"
            .ParameterFields(1).SetCurrentValue IIf(CompraInfResumen.ChkSaldo.Value = 1, True, False)
            .ParameterFields(2).SetCurrentValue ""
            .ParameterFields(3).SetCurrentValue False
            .ParameterFields(4).SetCurrentValue False
            .ParameterFields(5).SetCurrentValue ""
            .ParameterFields(6).SetCurrentValue ""
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre32()
    Set Rs = CompraInfVencimiento.RsP
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptVentaVencimiento
        With ReporteActivo
            .ReportTitle = "Vencimientos de Pagos a Proveedores"
            .Text1.SetText "Proveedor"
            .Text13.SetText "Pagado"
            .Text14.SetText "a Pagar"
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub


Private Sub Impre33()
    Dim cRle As New ClsClienteL, nOrd As Byte
    
    If CompraInfDetalles.OptCodigo.Value = True Then
        nOrd = 0
        sDesde = CompraInfDetalles.TxtCliente(0).Text
        sHasta = CompraInfDetalles.TxtCliente(1).Text
    Else
        nOrd = 1
        sDesde = CompraInfDetalles.LblCliente(0).Caption
        sHasta = CompraInfDetalles.LblCliente(1).Caption
    End If
    
    Set Rs = cRle.TraerDetalleFacturaP(sDesde, sHasta, CompraInfDetalles.DtpFecha(0).Value, CompraInfDetalles.DtpFecha(1).Value, nOrd)
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptDetalleFactura
        With ReporteActivo
            .ReportTitle = "Proveedores - Detalles de Facturas"
            .Text4.SetText "Proveedor"
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre34()
    Set Rs = VentaInfResumenDet.Rs
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptResConDetalle
        With ReporteActivo
            .ReportTitle = sNomFan
            ' Pasamos la fecha desde el formulario
            .ParameterFields(1).SetCurrentValue VentaInfResumenDet.DtpFecha(1).Value
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre35()

End Sub

Private Sub Impre36()
    ' Quitamos el Dim local y el As New
    Dim cRem As New ClsProductoL
    
    Set Rs = CajaResCaja.Rs
    
    If Rs.RecordCount <> 0 Then
        ' Asignamos a la variable global para que el botón Imprimir funcione
        Set ReporteActivo = New RptResCaja
        
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
    ' Eliminamos el Set Report27 = Nothing para que no se pierda la referencia
End Sub

Private Sub Impre37()
    ' Suponiendo que el formulario es CajaMovimientos
    Set Rs = CajaMovimientos.Rs
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptCajaDetalle
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre38()
    ' Eliminamos el Dim local de Report29
    Dim cRlx As New ClsComprobantesL, Rx As ADODB.Recordset, cRsl As ClsLectura
    Dim sStr As String
    
    Set cRsl = New ClsLectura
    Set Rx = cRsl.TraerTodos("Empresa", "*")
    
    Set Rs = VentaLibroIva.Rs
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos la variable global
        Set ReporteActivo = New RptIvaVenta
        
        With ReporteActivo
            ' Parámetros de encabezado
            .ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
            .ParameterFields(2).SetCurrentValue (Rx!Direccion)
            .ParameterFields(3).SetCurrentValue (Rx!Cuit)
            .ParameterFields(4).SetCurrentValue "Libro de IVA Ventas " & VentaLibroIva.TxtPeriodo
            .ParameterFields(5).SetCurrentValue (Val(VentaLibroIva.TxtPagina.Text))
            .ParameterFields(6).SetCurrentValue IIf(VentaLibroIva.ChkAgrupado.Value = 0, False, True)
            
            ' Cambio dinámico de etiquetas (TextObjects)
            .Text15.SetText ("Percepción Ing.Brutos")
            .Text16.SetText ("Retención Ganancias")
            
            ' Asignación de datos
            .Database.SetDataSource Rs
        End With
        
        ' Mostrar en el visor
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' Recordá: No poner Set ReporteActivo = Nothing aquí
    Set cRsl = Nothing
    Set Rx = Nothing
End Sub

Private Sub Impre39()
    ' Eliminamos el Dim local de Report29 para usar ReporteActivo
    Dim cRlx As New ClsComprobantesL, Rx As ADODB.Recordset, cRsl As ClsLectura
    
    Set cRsl = New ClsLectura
    Set Rx = cRsl.TraerTodos("Empresa", "*")
    
    Set Rs = ComprasLibroIva.Rs
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos la variable global con el reporte de IVA
        Set ReporteActivo = New RptIvaVenta
        
        With ReporteActivo
            ' Configuración de parámetros de encabezado
            .ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
            .ParameterFields(2).SetCurrentValue (Rx!Direccion)
            .ParameterFields(3).SetCurrentValue (Rx!Cuit)
            .ParameterFields(4).SetCurrentValue "Libro de IVA Compras " & ComprasLibroIva.TxtPeriodo
            .ParameterFields(5).SetCurrentValue (Val(ComprasLibroIva.TxtPagina.Text))
            .ParameterFields(6).SetCurrentValue IIf(ComprasLibroIva.ChkAgrupado.Value = 0, False, True)
            
            ' Ajuste dinámico de etiquetas específicas para Compras
            .Text15.SetText ("Percepción Ing.Brutos")
            .Text16.SetText ("Percepción Iva")
            .Text7.SetText ("Retención Ganancia")
            .Text4.SetText ("Proveedor")
            
            ' Vinculación de datos
            .Database.SetDataSource Rs
        End With
        
        ' Enviamos al visor
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' Limpieza de objetos de datos, pero mantenemos vivo ReporteActivo
    Set cRsl = Nothing
    Set Rx = Nothing
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
    ' Eliminamos el Dim local de report1 para usar la variable global ReporteActivo
    
    Set Rs = CajaCtaCteBan.Rs
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos el objeto global con el diseño de resumen
        Set ReporteActivo = New RptVentaResumen
        
        With ReporteActivo
            .ReportTitle = "Resumen de Cuenta"
            
            ' Configuración de parámetros (CUIDADO: mantener el orden exacto de los índices)
            .ParameterFields(1).SetCurrentValue False
            .ParameterFields(2).SetCurrentValue ""
            .ParameterFields(3).SetCurrentValue False
            .ParameterFields(4).SetCurrentValue False
            .ParameterFields(5).SetCurrentValue ""
            .ParameterFields(6).SetCurrentValue ""
            
            ' Asignación del Recordset
            .Database.SetDataSource Rs
        End With
        
        ' Configuración del visor
        CRViewer1.ReportSource = ReporteActivo
        CRViewer1.DisplayGroupTree = False
    End If
    
    ' Recordá: NO pongas Set ReporteActivo = Nothing para que el botón de imprimir funcione
End Sub

Private Sub Impre43()
    ' Eliminamos el Dim local de Report25 para usar ReporteActivo
    Dim cRl As ClsLectura, bIva As Boolean
    Dim cRem As New ClsProductoL, opt As Byte
    
    Set cRl = New ClsLectura
    
    ' Lógica para determinar si el comprobante lleva costo/IVA
    If cRl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & _
       VentaConsolidado.CmbComprobante.ItemData(VentaConsolidado.CmbComprobante.ListIndex)) = 1 Then
        bIva = True
    End If
    
    ' Obtenemos los datos del consolidado
    Set Rs = cRem.TraerConsolidado(VentaConsolidado.TxtNum(0).Text, _
                                   VentaConsolidado.TxtNum(1).Text, _
                                   VentaConsolidado.CmbComprobante.ItemData(VentaConsolidado.CmbComprobante.ListIndex), _
                                   bIva)
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos el objeto global con el reporte de Ranking (usado para consolidado)
        Set ReporteActivo = New RptRanking
        
        With ReporteActivo
            ' Seteamos parámetros (usando la fecha actual según tu código original)
            .ParameterFields(1).SetCurrentValue Date
            .ParameterFields(2).SetCurrentValue Date
            
            ' Vinculamos el Recordset
            .Database.SetDataSource Rs
        End With
        
        ' Asignamos al visor
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' Limpiamos objetos auxiliares pero dejamos vivo ReporteActivo
    Set cRl = Nothing
End Sub

Private Sub Impre44()
    ' Eliminamos el Dim local de Report33 para usar la variable global ReporteActivo
    Dim cRem As New ClsProductoL, opt As Byte
    
    Set Rs = CjaResCtaCble.Rs
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos el objeto global
        Set ReporteActivo = New RptPlanCuenta
        
        With ReporteActivo
            ' Pasamos los parámetros de fecha convirtiéndolos a String como en el original
            .ParameterFields(1).SetCurrentValue str(CjaResCtaCble.DtpDesde)
            .ParameterFields(2).SetCurrentValue str(CjaResCtaCble.DtpHasta)
            
            ' Vinculamos los datos
            .Database.SetDataSource Rs
        End With
        
        ' Cargamos el visor
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' Recordá: NO pongas Set ReporteActivo = Nothing para que el botón de imprimir funcione
End Sub

Private Sub Impre45()
    ' Eliminamos el Dim local de Report34 para usar ReporteActivo
    
    ' El recordset ya viene cargado desde el formulario de movimientos
    Set Rs = StockInfMovimientos.Rs
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos el objeto global con el reporte de movimientos de stock
        Set ReporteActivo = New RptStockMovimiento
        
        ' Vinculamos los datos directamente
        ReporteActivo.Database.SetDataSource Rs
        
        ' Cargamos el visor de Crystal con la variable global
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' IMPORTANTE: No liberamos el objeto (Set ReporteActivo = Nothing)
    ' para que el botón de imprimir del visor pueda acceder a él.
End Sub

Private Sub Impre46()
    ' Eliminamos el Dim local de Report35 para usar ReporteActivo
    Dim cRe As ClsProductoL, cRsb As ClsLectura
    Dim X, Y As Byte, rel As Recordset
    
    Set cRe = New ClsProductoL
    Set cRsb = New ClsLectura

    ' Obtenemos los productos para las etiquetas
    Set Rs = cRe.TraerProductosEtiquetas(MtrProdEtiquetas.DtpDesde.Value, MtrProdEtiquetas.DtpHasta.Value)
    
    ' (La lógica de generación de archivo para Balanza se mantiene comentada como en el original)
    
    If Rs.RecordCount <> 0 Then
        ' Inicializamos la variable global con el reporte de etiquetas
        Set ReporteActivo = New RptEtiquetas
        
        ' Vinculamos los datos
        ReporteActivo.Database.SetDataSource Rs
        
        ' Cargamos el visor
        CRViewer1.ReportSource = ReporteActivo
    End If
    
    ' IMPORTANTE: No liberamos ReporteActivo para que el selector de impresoras
    ' de Windows pueda manejar el cambio a la impresora de etiquetas o PDF.
    Set cRe = Nothing
    Set cRsb = Nothing
End Sub

Private Sub Impre47()
    Set Rs = MtrEtiquetas.rImp
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptEtiquetas
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre48()
    Set Rs = MtrEtiquetas.rImp
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptCodigoBarra
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre49()
    Set Rs = ComprasComprobantes.RsC
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptEtiquetasTango
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre50()
    Dim cRsl As New ClsLectura
    Set Rs = CajaCerrarZ.rPrin
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptCuentaFondos
        With ReporteActivo
            .Database.SetDataSource Rs
            .ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
            .ParameterFields(2).SetCurrentValue CajaCerrarZ.DtpFecha.Value
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre51()
    Dim cRsl As New ClsLectura
    Set Rs = CajaCerrarZ.rPrin
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptChequesTarjetas
        With ReporteActivo
            .Database.SetDataSource Rs
            .ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
            .ParameterFields(2).SetCurrentValue CajaCerrarZ.DtpFecha.Value
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
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
    Set Rs = FrmInformeCobranzas.Rs3
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptListaCobranza
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre54()
    Set Rs = MtrProductos.RsC
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptEtiquetasTango
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre55()
    Set Rs = VentaInfResumen.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptResuSimpli
        With ReporteActivo
            .ReportTitle = sNomFan
            .ParameterFields(1).SetCurrentValue IIf(VentaInfResumen.ChkSaldo.Value = 1, True, False)
            .ParameterFields(2).SetCurrentValue VentaInfResumen.CmbZona.Text
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
        CRViewer1.DisplayGroupTree = False
    End If
End Sub

Private Sub Impre56()
    Dim cRsl As New ClsLectura, Rx As ADODB.Recordset
    Set Rx = cRsl.TraerTodos("Empresa", "*")
    Set Rs = VentaLibroIva.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptIvaVentaRes
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
            .ParameterFields(2).SetCurrentValue (Rx!Direccion)
            .ParameterFields(3).SetCurrentValue "Resumen Libro de IVA Ventas " & VentaLibroIva.TxtPeriodo
            .ParameterFields(4).SetCurrentValue (Rx!Cuit)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre57()
    Dim cRsl As New ClsLectura, Rx As ADODB.Recordset
    Set Rx = cRsl.TraerTodos("Empresa", "*")
    Set Rs = ComprasLibroIva.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptIvaVentaRes
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue (Rx!Razon) & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
            .ParameterFields(2).SetCurrentValue (Rx!Direccion)
            .ParameterFields(3).SetCurrentValue "Libro de IVA Compras " & ComprasLibroIva.TxtPeriodo
            .ParameterFields(4).SetCurrentValue (Rx!Cuit)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre58()
    Dim cRem As New ClsProductoL
    Set Rs = cRem.TraerReparto(Val(VentaReparto.TxtReparto.Text), VentaReparto.DtpDesde.Value, VentaReparto.DtpHasta.Value)
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptReparto
        ReporteActivo.ParameterFields(1).SetCurrentValue Val(VentaReparto.TxtReparto.Text)
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre59()
    Set Rs = CompraInfResumen.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptResuSimpli
        With ReporteActivo
            .ReportTitle = "Resumen de Cuenta Simplificado - Proveedores"
            .ParameterFields(1).SetCurrentValue IIf(CompraInfResumen.ChkSaldo.Value = 1, True, False)
            .ParameterFields(2).SetCurrentValue "Todas"
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
        CRViewer1.DisplayGroupTree = False
    End If
End Sub

Private Sub Impre60()
    Set Rs = CjaResCtaCble.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptPlanCtaDet
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue str(CjaResCtaCble.DtpDesde)
            .ParameterFields(2).SetCurrentValue str(CjaResCtaCble.DtpHasta)
            .ParameterFields(3).SetCurrentValue CjaResCtaCble.CmbCuenta.Text
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre61()
    Set Rs = StockInfMovimientos.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptCedronar
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue StockInfMovimientos.DtpFecha(0).Value
            .ParameterFields(2).SetCurrentValue StockInfMovimientos.DtpFecha(1).Value
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre62()
    Dim cRlx As New ClsComprobantesL
    Set Rs = cRlx.TraerInPercIB(VentasInfPerIb.DtpDesde.Value, VentasInfPerIb.DtpHasta.Value)
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptInfPerIB
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue str(VentasInfPerIb.DtpDesde.Value)
            .ParameterFields(2).SetCurrentValue str(VentasInfPerIb.DtpHasta.Value)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre63()
    Dim cRlx As New ClsComprobantesL
    Set Rs = cRlx.TraerInfRetIB(ComprasInfRetIb.DtpDesde.Value, ComprasInfRetIb.DtpHasta.Value)
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RtpInfRetIb
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue str(ComprasInfRetIb.DtpDesde.Value)
            .ParameterFields(2).SetCurrentValue str(ComprasInfRetIb.DtpHasta.Value)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre64()
    Dim cRsl As New ClsLectura
    Set Rs = ComprasAnularPagos.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptR122V2
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Cuit", "Id=" & nMenu - 1)
            .ParameterFields(2).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "Id=" & nMenu - 1)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre65()
    Set Rs = VentaInfResumen.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptInfAplica
        With ReporteActivo
            .ReportTitle = sNomFan
            .ParameterFields(1).SetCurrentValue str(VentaInfResumen.DtpFecha(0).Value)
            .ParameterFields(2).SetCurrentValue str(VentaInfResumen.DtpFecha(1).Value)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
        CRViewer1.DisplayGroupTree = False
    End If
End Sub

Private Sub Impre66()
    Set Rs = VentasInfProy.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptProyeccion
        With ReporteActivo
            .ReportTitle = sNomFan
            .ParameterFields(1).SetCurrentValue str(VentasInfProy.DtpFecha(0).Value)
            .ParameterFields(2).SetCurrentValue str(VentasInfProy.DtpFecha(1).Value)
            .ParameterFields(3).SetCurrentValue str(VentasInfProy.DtpFecha(2).Value)
            .ParameterFields(4).SetCurrentValue str(VentasInfProy.DtpFecha(3).Value)
            .ParameterFields(5).SetCurrentValue str(VentasInfProy.DtpFecha(4).Value)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
        CRViewer1.DisplayGroupTree = False
    End If
End Sub

Private Sub Impre67()
    Dim cRsl As New ClsLectura
    Set Rs = ComprasAnularPagos.Rs
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptCertGanancia
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Razon", "Id=" & nMenu - 1)
            .ParameterFields(2).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Direccion", "Id=" & nMenu - 1)
            .ParameterFields(3).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Localidad", "Id=" & nMenu - 1)
            .ParameterFields(4).SetCurrentValue cRsl.TraerValorDeUnCampo("Empresa", "Cuit", "Id=" & nMenu - 1)
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre68()
    Dim bOrd As Integer, nStock As Byte
    Set cCanti = New ClsProductoL
    pLista = StockInfSaldos.CmbDeposito.ItemData(StockInfSaldos.CmbDeposito.ListIndex)
    bOrd = 3
    sDesde = StockInfSaldos.LblDescripcion(0).Caption
    sHasta = StockInfSaldos.LblDescripcion(1).Caption
    nStock = StockInfSaldos.ChkSaldo.Value
    
    Set Rs = cCanti.DameProductos(pLista, sDesde, sHasta, bOrd, nStock)
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptStockSaldo
        ReporteActivo.ParameterFields(1).SetCurrentValue StockInfSaldos.CmbDeposito.Text
        ReporteActivo.Database.SetDataSource Rs
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Impre69()
    Dim pZona As String
    pZona = VentaRepartoRemito.CmbZona.Text
    sDesde = VentaRepartoRemito.DtpDesde.Value
    sHasta = VentaRepartoRemito.DtpHasta.Value
    Set Rs = VentaRepartoRemito.Rs
    
    If Rs.RecordCount <> 0 Then
        Set ReporteActivo = New RptRepartoRemi
        With ReporteActivo
            .ParameterFields(1).SetCurrentValue sDesde
            .ParameterFields(2).SetCurrentValue sHasta
            .ParameterFields(3).SetCurrentValue pZona
            .ParameterFields(4).SetCurrentValue VentaRepartoRemito.CmbEstado.Text
            .Database.SetDataSource Rs
        End With
        CRViewer1.ReportSource = ReporteActivo
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Liberamos el reporte activo para no dejar basura en RAM
    Set ReporteActivo = Nothing
    ' Cerramos el Recordset global si está abierto
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
        Set Rs = Nothing
    End If
End Sub
