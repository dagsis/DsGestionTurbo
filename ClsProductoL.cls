VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProductoL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mConn As ADODB.Connection
Private mConnStr As String

Private Sub Class_Initialize()
    mConnStr = "FILE NAME=" & App.Path & "\Conex.udl"
    Set mConn = New ADODB.Connection
    mConn.ConnectionTimeout = 5
    mConn.CommandTimeout = 30
    mConn.Open mConnStr
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not mConn Is Nothing Then
        If mConn.State <> adStateClosed Then mConn.Close
    End If
    Set mConn = Nothing
End Sub

Public Sub EnsureConn()
    If mConn Is Nothing Then
        Set mConn = New ADODB.Connection
        sDb = "FILE NAME=" & App.Path & "\Conex.udl"
        mConn.Open sDb
    ElseIf mConn.State = adStateClosed Then
        sDb = "FILE NAME=" & App.Path & "\Conex.udl"
        mConn.Open sDb
    End If
End Sub

Private Function DetachRS(ByVal Rs As ADODB.Recordset) As ADODB.Recordset
    Rs.ActiveConnection = Nothing   ' <-- SIN Set
    Set DetachRS = Rs
End Function

Public Function TraerProducto(sProd As String, Optional pDat As Long) As ADODB.Recordset
  On Error GoTo errHandler
  EnsureConn

  Dim Rs As ADODB.Recordset
  Set Rs = New ADODB.Recordset

  Dim sSql As String
    
  sSql = "SELECT Productos.Id, Productos.Producto, Productos.Moneda, Productos.Descripcion, Productos.UMVenta, Productos.Deposito, Productos.TipoInterno, Productos.ValorInterno,"
  sSql = sSql & "Impuestos.Impuesto ,Impuestos.Porcentaje, Precios.Lista, Precios.Precio "
  sSql = sSql & "FROM Impuestos RIGHT OUTER JOIN "
  sSql = sSql & "ProductoImpuesto ON Impuestos.Impuesto = ProductoImpuesto.Impuesto RIGHT OUTER JOIN "
  sSql = sSql & "Productos INNER JOIN "
  sSql = sSql & "Precios ON Productos.Producto = Precios.Producto ON ProductoImpuesto.Producto = Productos.Producto "
  If pDat = 0 Then
     sSql = sSql & "WHERE Productos.Producto='" & sProd & "'"
  Else
     sSql = sSql & "WHERE Productos.Id=" & pDat
  End If
    
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
  Rs.ActiveConnection = Nothing
         
  Set TraerProducto = Rs
  Exit Function
   
errHandler:
   err.Raise err.Number
End Function

Public Function TieneImpuesto(sProd As String, nImpu As Integer) As Boolean
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Chk As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Chk = New ADODB.Recordset
   
   db.Open sDb
   sSql = "SELECT  * FROM ProductoImpuesto WHERE Producto='" & sProd & "' and Impuesto=" & nImpu
 
   Rs_Chk.Open sSql, db, adOpenStatic
   If Rs_Chk.BOF = False Or Rs_Chk.EOF = False Then
      TieneImpuesto = True
   End If
   Rs_Chk.Close
   
   db.Close
   Set Rs_Chk = Nothing
   Exit Function
   
errHandler:
   err.Raise err.Number
End Function
Public Function BuscarNombreProducto(sCod As String) As String
  Dim Rs As ADODB.Recordset, sSql As String
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Descripcion FROM Productos WHERE Producto='" & sCod & "' and Activo=1"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarNombreProducto = "" & Rs!Descripcion
  End If
  db.Close
Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function DameProductoProvee(sPdesde As Integer) As Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Productos.Producto, Productos.Descripcion, Proveedores.RazonSocial "
  sSql = sSql & "FROM Proveedores RIGHT JOIN Productos ON Proveedores.Id = Productos.Proveedor "
  If sPdesde <> 0 Then
     sSql = sSql & "WHERE Proveedores.Id>=" & sPdesde & " "
  End If
  sSql = sSql & "ORDER BY Productos.Descripcion"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set DameProductoProvee = Rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function DameCantidad(sCod As String, Optional nDepo As Integer) As Long
 Dim Rs As ADODB.Recordset, sSql As String
 
 On Error GoTo errHandler
 Set Rs = New ADODB.Recordset
 db.Open sDb

 If sCod <> "" Then
    If nDepo <> 0 Then
       sSql = "SELECT (Sum(CantidadDebe)-Sum(CantidadHaber)) as Canti FROM Cantidades WHERE Producto='" & sCod & "'  and Deposito=" & nDepo
    Else
       sSql = "SELECT (Sum(CantidadDebe)-Sum(CantidadHaber)) as Canti FROM Cantidades WHERE Producto='" & sCod & "'"
    End If
    sSql = sSql & " AND Estado=1 "
    
    Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
    DameCantidad = 0
    If Rs!Canti <> 0 Then
       DameCantidad = Rs!Canti
    End If
    db.Close
End If
Exit Function

errHandler:
  err.Raise err.Number
 
End Function

Public Function TraerPrecio(pLista As Integer, sProduc As String) As Single
  Dim sSql As String, Rs As ADODB.Recordset
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Precio FROM Precios WHERE Lista=" & pLista & " and Producto='" & sProduc & "'"
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     TraerPrecio = Rs!precio
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerPrecioCompra(sProduc As String) As Double
  Dim sSql As String, Rs As ADODB.Recordset
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT PrecioCompra FROM Productos WHERE Producto='" & sProduc & "' and Activo=1"
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     TraerPrecioCompra = Rs!PrecioCompra
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerProductoTalle(pDepo As Integer, pDesde As String, pHasta As String, pOrd As Integer) As ADODB.Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 db.Open sDb

 Set Rs = New ADODB.Recordset
 If pDepo = 0 Then
    sSql = "SELECT Cantidades.Producto, Productos.Descripcion, Talles.Descripcion, (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber)) AS Cantidad "
    sSql = sSql & "FROM ((Cantidades LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto) LEFT JOIN Talles ON Cantidades.Lote = Talles.Id) LEFT JOIN Depositos ON Cantidades.Deposito = Depositos.Deposito "
    sSql = sSql & "Where Cantidades.Anulado = 0 and Cantidades.Estado=1 "
    sSql = sSql & "GROUP BY Cantidades.Producto, Productos.Descripcion, Talles.Descripcion, Cantidades.Anulado "
    sSql = sSql & "HAVING (((Cantidades.Producto)>='" & pDesde & "' And (Cantidades.Producto)<='" & pHasta & "')) "
    If pOrd = 0 Then
        sSql = sSql & "ORDER BY Cantidades.Producto,Talles.Descripcion"
    Else
        sSql = sSql & "ORDER BY Productos.Descripcion,Talles.Descripcion"
    End If
 Else
    sSql = "SELECT Cantidades.Producto, Productos.Descripcion, Talles.Descripcion, (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber)) AS Cantidad "
    sSql = sSql & "FROM ((Cantidades LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto) LEFT JOIN Talles ON Cantidades.Lote = Talles.Id) LEFT JOIN Depositos ON Cantidades.Deposito = Depositos.Deposito "
    sSql = sSql & "Where (((Depositos.Deposito) =" & pDepo & ") And ((Cantidades.Anulado) = 0)) "
    sSql = sSql & "GROUP BY Cantidades.Producto, Productos.Descripcion, Talles.Descripcion, Cantidades.Anulado "
    sSql = sSql & "HAVING (((Cantidades.Producto)>='" & pDesde & "' And (Cantidades.Producto)<='" & pHasta & "')) "
    If pOrd = 0 Then
        sSql = sSql & "ORDER BY Cantidades.Producto,Talles.Descripcion"
    Else
        sSql = sSql & "ORDER BY Productos.Descripcion,Talles.Descripcion"
    End If
 End If
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

 Set TraerProductoTalle = Rs
 db.Close
End Function

Public Function TraerTotalesDeDepositos() As ADODB.Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
  db.Open sDb
  
  sSql = "SELECT Depositos.Descripcion, (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber)) AS Total "
  sSql = sSql & "FROM Cantidades LEFT JOIN Depositos ON Cantidades.Deposito = Depositos.Deposito "
  sSql = sSql & "WHERE Cantidades.Estado=1 "
  sSql = sSql & "GROUP BY Depositos.Descripcion "
  
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerTotalesDeDepositos = Rs
  db.Close
  
End Function

Public Function DameProductos(pDepo As Integer, pDesde As String, pHasta As String, pOrd As Integer, Optional pStock As Byte, Optional pFamilia As Integer) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
 db.Open sDb

 If pDepo = 0 Then
    sSql = "SELECT Cantidades.Producto, Productos.Descripcion, (Sum(Cantidades.CantidadDebe) - Sum(CantidadHaber)) AS SumaDeCantidad "
    sSql = sSql & "FROM Cantidades LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto "
    sSql = sSql & "WHERE  Cantidades.Estado=1 AND Productos.Activo=1 " ' And Cantidades.Suc=" & nSucursal & " "
    If pFamilia <> 0 Then
       sSql = sSql & "AND Productos.Familia=" & pFamilia & " "
    End If
    sSql = sSql & "GROUP BY Cantidades.Producto, Productos.Descripcion, Cantidades.Producto "
    If pDesde <> "" Then
       If pOrd = 0 Then
          sSql = sSql & "HAVING (((Cantidades.Producto)>='" & pDesde & "' And (Cantidades.Producto)<='" & pHasta & "')) "
       Else
           sSql = sSql & "HAVING (((Productos.Descripcion)>='" & pDesde & "' And (Productos.Descripcion)<='" & pHasta & "')) "
       End If
       If pStock = 1 Then
          sSql = sSql & " AND Sum(Cantidades.CantidadDebe) - Sum(CantidadHaber)<>0 "
       End If
    End If
    If pOrd = 2 Then
       sSql = sSql & " ORDER BY Productos.Descripcion"
    Else
       sSql = sSql & " ORDER BY Cantidades.Producto"
    End If
 Else
    sSql = "SELECT Cantidades.Producto, Productos.Descripcion, (Sum(Cantidades.CantidadDebe) - Sum(CantidadHaber)) AS SumaDeCantidad ,Cantidades.Deposito "
    sSql = sSql & "FROM Cantidades LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto "
    sSql = sSql & "WHERE  Cantidades.Deposito=" & pDepo & " AND Productos.Activo=1 "
    If pFamilia <> 0 Then
       sSql = sSql & "AND Productos.Familia=" & pFamilia & " "
    End If
    If pOrd = 3 Then
       sSql = sSql & "And Cantidades.Estado=2 "
    Else
       sSql = sSql & "And Cantidades.Estado=1 "
    End If
    sSql = sSql & " GROUP BY Cantidades.Producto, Productos.Descripcion, Cantidades.Producto,Cantidades.Deposito "
    If pDesde <> "" Then
        If pOrd = 0 Then
          sSql = sSql & "HAVING (((Cantidades.Producto)>='" & pDesde & "' And (Cantidades.Producto)<='" & pHasta & "')) "
       Else
           sSql = sSql & "HAVING (((Productos.Descripcion)>='" & pDesde & "' And (Productos.Descripcion)<='" & pHasta & "')) "
       End If
       If pStock = 1 Then
          sSql = sSql & " AND Sum(Cantidades.CantidadDebe) - Sum(CantidadHaber)<>0 "
       End If
    End If
    If pOrd = 1 Then
       sSql = sSql & " ORDER BY Productos.Descripcion"
    Else
       sSql = sSql & " ORDER BY Cantidades.Producto"
    End If
 End If
 Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
 Rs.Close
 Rs.CursorType = adOpenStatic
 Rs.CursorLocation = adUseClient
 Rs.LockType = adLockBatchOptimistic
 Rs.Open
 Rs.ActiveConnection = Nothing

 Set DameProductos = Rs
 db.Close
 
End Function

Public Function DameReposiciones(sDesde As String, sHasta As String, sPdesde As Integer, pOrd As Integer) As Recordset
  Dim sSql As String, Rs As ADODB.Recordset

  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.CantMinima, (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber)) AS SumaDeCantidad, Proveedores.RazonSocial "
  sSql = sSql & "FROM Proveedores RIGHT JOIN (Productos LEFT JOIN Cantidades ON Productos.Producto = Cantidades.Producto) ON Proveedores.Id = Productos.Proveedor "
  sSql = sSql & "WHERE Productos.Activo=1 and Cantidades.Anulado<>1 And Cantidades.Suc=" & nSucursal & " "
   If sPdesde <> 0 Then
     sSql = sSql & "AND Proveedores.Id=" & sPdesde & " "
  End If
  sSql = sSql & "GROUP BY Productos.Producto, Productos.Descripcion, Productos.CantMinima, Proveedores.RazonSocial "
  If pOrd = 1 Then
    sSql = sSql & "HAVING Productos.Producto>='" & sDesde & "' And Productos.Producto<='" & sHasta & "' "
  Else
    sSql = sSql & "HAVING Productos.Descripcion>='" & sDesde & "' And Productos.Descripcion<='" & sHasta & "' "
  End If
  
  sSql = sSql & "AND ((Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber))<=[Productos].[CantMinima]) "
 
  If pOrd = 1 Then
     sSql = sSql & " ORDER BY Productos.Producto"
  Else
     sSql = sSql & " ORDER BY Productos.Descripcion"
  End If
   
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set DameReposiciones = Rs
  db.Close
  
End Function

Public Function DameListaProducto(sPdesde As String, sPhasta As String, sProDes As String, sProHas As String, nOrd As Byte, sTip As String, sDatDes As String, sDatHas As String) As Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Productos.Producto, Productos.Descripcion, Familias.Descripcion, Marcas.Descripcion, Proveedores.RazonSocial, UnidadMedida.Descripcion, UnidadMedida_1.Descripcion, Productos.CantMinima, Productos.CantMaxima "
  sSql = sSql & "FROM (UnidadMedida RIGHT JOIN (Proveedores RIGHT JOIN (Marcas RIGHT JOIN (Familias RIGHT JOIN Productos ON Familias.Familia = Productos.Familia) ON Marcas.Marca = Productos.Marca) ON Proveedores.Id = Productos.Proveedor) ON UnidadMedida.Unidad = Productos.UMCompra) LEFT JOIN UnidadMedida AS UnidadMedida_1 ON Productos.UMVenta = UnidadMedida_1.Unidad "
  sSql = sSql & " WHERE Productos.Producto'" >= sPdesde & "' and Productos.Producto<='" & sPhasta & "' and Productos.Activo=1 "
  sSql = sSql & " and Proveedores.RazonSocial>='" & sProDes & "' and Proveedores.RazonSocial<='" & sProHas & "' "
  If sDatDes <> "" Then
     If sTip = "F" Then
        sSql = sSql & "and Familias.Descripcion>='" & sDatDes & "' and Familias.Descripcion<='" & sDatHas & "'"
     Else
        sSql = sSql & "and Marcas.Descripcion>='" & sDatDes & "' and Marcas.Descripcion<='" & sDatHas & "'"
     End If
  End If
  If nOrd = 0 Then
     sSql = "ORDER BY Productos.Producto"
  Else
     sSql = "ORDER BY Productos.Descripcion"
  End If
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set DameListaProducto = Rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function MueveStock(sProduc As String) As Boolean
  Dim sSql As String, Rs As ADODB.Recordset
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT ActuaStock FROM Productos WHERE Producto='" & sProduc & "'"
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     If Rs!ActuaStock = "SI" Then
        MueveStock = True
     End If
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function ProductoCompuesto(sProduc As String) As Boolean
  Dim sSql As String, Rs As ADODB.Recordset
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CodCompu FROM Productos WHERE Producto='" & sProduc & "'"
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     If Rs!CodCompu = "Si" Then
        ProductoCompuesto = True
     End If
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function ProductoDigito(sProduc As String) As Byte
  Dim sSql As String, Rs As ADODB.Recordset
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CodCompu FROM Productos WHERE Producto='" & sProduc & "'"
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     ProductoDigito = Rs!CodCompu
  End If
  
  db.Close
  
End Function

Public Function ProductoListadoMtr(pDesde As String, pHasta As String, pOrd As Byte, pGru As Byte, _
  Optional sDesde As String, Optional sHasta As String) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Productos.Producto, Productos.Descripcion, Marcas.Descripcion, UnidadMedida.Descripcion, UnidadMedida_1.Descripcion, Depositos.Descripcion, Familias.Descripcion, Proveedores.RazonSocial "
  sSql = sSql & "FROM Proveedores RIGHT JOIN (UnidadMedida RIGHT JOIN (Marcas RIGHT JOIN (Depositos RIGHT JOIN ((Productos LEFT JOIN UnidadMedida AS UnidadMedida_1 ON Productos.UMVenta = UnidadMedida_1.Unidad) LEFT JOIN Familias ON Productos.Familia = Familias.Familia) ON Depositos.Deposito = Productos.Deposito) ON Marcas.Marca = Productos.Marca) ON UnidadMedida.Unidad = Productos.UMCompra) ON Proveedores.Id = Productos.Proveedor "
  sSql = sSql & "WHERE Productos.Descripcion>='" & pDesde & "' And Productos.Descripcion<='" & pHasta & "' and Productos.Activo=1 "
  Select Case pGru
         Case 0   ' Ninguno
              sSql = sSql & ""
         Case 3   ' Marca
              sSql = sSql & "AND Marcas.Descripcion>='" & sDesde & "' AND Marcas.Descripcion<='" & sHasta & "' "
         Case 7   ' Familia
              sSql = sSql & "AND Familias.Descripcion>='" & sDesde & "' AND Familias.Descripcion<='" & sHasta & "' "
         Case 8   ' Proveedores
              sSql = sSql & "AND Proveedores.RazonSocial>='" & sDesde & "' AND Proveedores.RazonSocial<='" & sHasta & "' "
  End Select
  If pOrd = 0 Then
     sSql = sSql & "ORDER BY Productos.Producto "
  Else
     sSql = sSql & "ORDER BY Productos.Descripcion "
  End If
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set ProductoListadoMtr = Rs
  db.Close
  
End Function

Public Function TraerProductoProveedor(pProve As String) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  db.Open sDb

  Set Rs = New ADODB.Recordset
  sSql = "SELECT ProductoProveedor.Producto, ProductoProveedor.CodProvee, Productos.Descripcion "
  sSql = sSql & "FROM Productos RIGHT JOIN ProductoProveedor ON Productos.Producto = ProductoProveedor.Producto "
  sSql = sSql & "WHERE ProductoProveedor.Proveedor='" & pProve & "'"
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerProductoProveedor = Rs
  db.Close
  
End Function

Public Function EstaElCodigoProv(pProve As String, pProdu As String) As Boolean
  Dim Rs As ADODB.Recordset, sSql As String

  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CodProvee FROM ProductoProveedor WHERE Proveedor='" & pProve & "' "
  sSql = sSql & "AND CodProvee='" & pProdu & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.RecordCount <> 0 Then
     EstaElCodigoProv = True
  End If
  db.Close
  
End Function

Public Function EstaElProdProv(pProve As String, pProdu As String) As Boolean
  Dim Rs As ADODB.Recordset, sSql As String

  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT CodProvee FROM ProductoProveedor WHERE Proveedor='" & pProve & "' "
  sSql = sSql & "AND Producto='" & pProdu & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.RecordCount <> 0 Then
     EstaElProdProv = True
  End If
  db.Close
  
End Function

Public Function TraerUnProductoProveedor(pProve As String, pProd) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT ProductoProveedor.Producto, ProductoProveedor.CodProvee, Productos.Descripcion "
  sSql = sSql & "FROM Productos RIGHT JOIN ProductoProveedor ON Productos.Producto = ProductoProveedor.Producto "
  sSql = sSql & "WHERE ProductoProveedor.Proveedor='" & pProve & "' AND ProductoProveedor.CodProvee='" & pProd & "'"
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerUnProductoProveedor = Rs
  db.Close
  
End Function
Public Function TraerVentaFamilia(pDesde As Date, pHasta As Date, pCliDesde As String, pCliHasta As String, pCliente As Byte, pResumen As Byte) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset
  db.Open sDb

  If pCliente = 0 Or pResumen = 1 Then
     sSql = "SELECT Familias.Descripcion,CabComprobantes.Comprobante,"
     sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN Sum(DetallesComprobantes.PrecioTotal) "
     sSql = sSql & "ELSE -Sum(DetallesComprobantes.PrecioTotal) END  AS SumaDePrecioTotal, "
     sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN Sum(DetallesComprobantes.Impuesto) "
     sSql = sSql & "ELSE -Sum(DetallesComprobantes.Impuesto) END as Impuesto "
     sSql = sSql & "FROM CabComprobantes LEFT JOIN (Familias RIGHT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON Familias.Familia = Productos.Familia) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
     sSql = sSql & "WHERE (CabComprobantes.Tipo=1 OR CabComprobantes.Tipo=5 ) and CabComprobantes.Fecha >= '" & FechaAmericana(pDesde) & "' And CabComprobantes.Fecha <= '" & FechaAmericana(pHasta) & "' and Familias.Descripcion is not null And CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "
     If pResumen = 1 Then
        sSql = sSql & "AND CabComprobantes.Cliente>='" & pCliDesde & "' AND CabComprobantes.Cliente<='" & pCliHasta & "'"
     End If
     sSql = sSql & "GROUP BY Familias.Descripcion,CabComprobantes.Comprobante,CabComprobantes.Tipo "
     sSql = sSql & "ORDER BY Familias.Descripcion"
  Else
     sSql = "SELECT Familias.Descripcion, CabComprobantes.Comprobante, Clientes.Cliente, Clientes.RazonSocial,"
     sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN Sum(DetallesComprobantes.PrecioTotal) "
     sSql = sSql & "ELSE -Sum(DetallesComprobantes.PrecioTotal) END  AS SumaDePrecioTotal, "
     sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN Sum(DetallesComprobantes.Impuesto) "
     sSql = sSql & "ELSE -Sum(DetallesComprobantes.Impuesto) END as Impuesto "
     sSql = sSql & "FROM (Familias RIGHT JOIN (CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) ON Familias.Familia = Productos.Familia) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente "
     sSql = sSql & "WHERE (CabComprobantes.Tipo=1 OR CabComprobantes.Tipo=5 ) and CabComprobantes.Fecha >= '" & FechaAmericana(pDesde) & "' And CabComprobantes.Fecha <= '" & FechaAmericana(pHasta) & "' and Familias.Descripcion is not null And CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "
     sSql = sSql & "AND Clientes.Cliente>='" & pCliDesde & "' AND Clientes.Cliente<='" & pCliHasta & "' "
     sSql = sSql & "GROUP BY Familias.Descripcion, CabComprobantes.Comprobante, Clientes.Cliente, Clientes.RazonSocial,CabComprobantes.Tipo "
     sSql = sSql & "order by Clientes.Cliente,Familias.Descripcion"
  End If
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerVentaFamilia = Rs
  db.Close
  
End Function

Public Function TraerPlanillaUtilidades(pDesde As Date, pHasta As Date) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset

  db.Open sDb

'  sSql = "SELECT Familias.Descripcion,"
'  sSql = sSql & "CASE "
'  sSql = sSql & "WHEN TipoMovimiento.Descripcion='Resta' "
'  sSql = sSql & "THEN -Sum(DetallesComprobantes.PrecioTotal) "
'  sSql = sSql & "ELSE Sum (DetallesComprobantes.PrecioTotal) "
'  sSql = sSql & "END AS Expr1 ,TipoMovimiento.Descripcion,CabComprobantes.Movimiento "
'
'  Sum(IIf(CabComprobantes.Tipo=1,DetallesComprobantes.PrecioTotal+DetallesComprobantes.Impuesto*CabComprobantes.Cotizacion,-DetallesComprobantes.PrecioTotal+DetallesComprobantes.Impuesto*CabComprobantes.Cotizacion)) AS SumaDePrecioTotal, Familias.Porcentaje "
'  sSql = sSql & "FROM Familias RIGHT JOIN (CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) ON Familias.Familia = Productos.Familia "
'  sSql = sSql & "Where (((CabComprobantes.Tipo) = 1 Or (CabComprobantes.Tipo) = 5) And ((CabComprobantes.Fecha) >= '" & FechaAmericana(pDesde) & "' And (CabComprobantes.Fecha) <= '" & FechaAmericana(pHasta) & "')) and CabComprobantes.Suc=" & nSucursal & " "
'  sSql = sSql & "GROUP BY Familias.Descripcion, Familias.Porcentaje "
'  sSql = sSql & "HAVING (((Familias.Descripcion) Is Not Null)) "
    
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerPlanillaUtilidades = Rs
  db.Close
  
End Function

Public Function TraerTotalVentas(pDesde As Date, pHasta As Date) As Single
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Sum((CabComprobantes.Debe + CabComprobantes.Iva1)  * CabComprobantes.Cotizacion) AS SumaDeDebe "
  sSql = sSql & "From CabComprobantes "
  sSql = sSql & "WHERE CabComprobantes.Fecha >= '" & FechaAmericana(pDesde) & "' And CabComprobantes.Fecha <= '" & FechaAmericana(pHasta) & "' "
  sSql = sSql & "GROUP BY CabComprobantes.Tipo "
  sSql = sSql & "HAVING (((CabComprobantes.Tipo)=1))"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  
  TraerTotalVentas = Rs!SumaDeDebe
  db.Close
  
End Function

Public Function TraerRanking(pPDesde As String, pPHasta As String, pDesde As Date, pHasta As Date, pOrd As Byte, pBus As Byte) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset
  db.Open sDb


'  sSql = "SELECT DetallesComprobantes.Producto, Productos.Descripcion,"
'  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN Sum(DetallesComprobantes.Cantidad) "
'  sSql = sSql & "ELSE -Sum(DetallesComprobantes.Cantidad) END  AS SumaDeCantidad, "
'  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN  Sum((DetallesComprobantes.PrecioTotal + DetallesComprobantes.Impuesto)  * CabComprobantes.Cotizacion) "
'  sSql = sSql & "ELSE - Sum((DetallesComprobantes.PrecioTotal + DetallesComprobantes.Impuesto)  * CabComprobantes.Cotizacion) END as SumaDePrecioTotal "
'  sSql = sSql & "FROM CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
'  sSql = sSql & "Where (((CabComprobantes.Tipo) = 1 or (CabComprobantes.Tipo) = 5) And ((CabComprobantes.Fecha) >= '" & FechaAmericana(pDesde) & "' And (CabComprobantes.Fecha) <= '" & FechaAmericana(pHasta) & "')) And CabComprobantes.Suc=" & nSucursal & " "
  
  sSql = "SELECT DetallesComprobantes.Producto, Productos.Descripcion,"
  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN Sum(DetallesComprobantes.Cantidad) "
  sSql = sSql & "ELSE -Sum(DetallesComprobantes.Cantidad) END  AS SumaDeCantidad, "
  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN  Sum((DetallesComprobantes.Precio )  * CabComprobantes.Cotizacion) "
  sSql = sSql & "ELSE - Sum((DetallesComprobantes.Precio )  * CabComprobantes.Cotizacion) END as SumaDePrecioTotal "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "Where (((CabComprobantes.Tipo) = 1 or (CabComprobantes.Tipo) = 5) And ((CabComprobantes.Fecha) >= '" & FechaAmericana(pDesde) & "' And (CabComprobantes.Fecha) <= '" & FechaAmericana(pHasta) & "')) And CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "

  If pBus = 0 Then
     sSql = sSql & "AND DetallesComprobantes.Producto>='" & pPDesde & "' and DetallesComprobantes.Producto<='" & pPHasta & "' "
  Else
     sSql = sSql & "AND DetallesComprobantes.Descripcion>='" & pPDesde & "' and DetallesComprobantes.Descripcion<='" & pPHasta & "' "
  End If
  sSql = sSql & "GROUP BY DetallesComprobantes.Producto, Productos.Descripcion,CabComprobantes.Tipo "
  sSql = sSql & "Having (((Productos.Descripcion) Is Not Null)) "
  If pOrd = 0 Then
     sSql = sSql & "ORDER BY Sum(DetallesComprobantes.Cantidad) DESC"
  Else
     sSql = sSql & "ORDER BY Sum(DetallesComprobantes.PrecioTotal) DESC"
  End If
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  
  Set TraerRanking = Rs
  db.Close
  
End Function
Public Function ProductoImpuesto(sProduc As String) As Single
  Dim Rs As ADODB.Recordset, sSql As String
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Impuestos.Porcentaje "
  sSql = sSql & "FROM Impuestos RIGHT JOIN (Productos LEFT JOIN ProductoImpuesto ON Productos.Producto = ProductoImpuesto.Producto) ON Impuestos.Impuesto = ProductoImpuesto.Impuesto "
  sSql = sSql & "WHERE Productos.Producto='" & sProduc & "'"
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     ProductoImpuesto = Rs!Porcentaje
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function ProductoImpuestoDesc(sProduc As String) As String
  Dim Rs As ADODB.Recordset, sSql As String
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  db.Open sDb

   sSql = "SELECT Impuestos.Descripcion "
   sSql = sSql & "FROM ProductoImpuesto LEFT JOIN Impuestos ON ProductoImpuesto.Impuesto = Impuestos.Impuesto "
   sSql = sSql & "WHERE ProductoImpuesto.Producto='" & sProduc & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  If Rs.EOF = False Or Rs.BOF = False Then
     ProductoImpuestoDesc = Rs!Descripcion
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function


Public Function TraerProductoCliente(pProve As String) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  db.Open sDb

  Set Rs = New ADODB.Recordset
  sSql = "SELECT ProductoCliente.Producto, ProductoCliente.CodCliente, Productos.Descripcion "
  sSql = sSql & "FROM Productos RIGHT JOIN ProductoCliente ON Productos.Producto = ProductoCliente.Producto "
  sSql = sSql & "WHERE ProductoCliente.Cliente='" & pProve & "'"
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerProductoCliente = Rs
  db.Close
  
End Function

Public Function EstaElCodigoCliente(pProve As String, pProdu As String) As Boolean
  Dim Rs As ADODB.Recordset, sSql As String

  Set Rs = New ADODB.Recordset
  db.Open sDb
 
  sSql = "SELECT CodCliente FROM ProductoCliente WHERE Cliente='" & pProve & "' "
  sSql = sSql & "AND CodCliente='" & pProdu & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.RecordCount <> 0 Then
     EstaElCodigoCliente = True
  End If
  db.Close
  
End Function

Public Function EstaElProdCliente(pProve As String, pProdu As String) As Boolean
  Dim Rs As ADODB.Recordset, sSql As String

  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CodCliente FROM ProductoCliente WHERE Cliente='" & pProve & "' "
  sSql = sSql & "AND Producto='" & pProdu & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.RecordCount <> 0 Then
     EstaElProdCliente = True
  End If
  db.Close
  
End Function

Public Function TraerUnProductoCliente(pProve As String, pProd) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT ProductoCliente.Producto, ProductoCliente.CodCliente, Productos.Descripcion "
  sSql = sSql & "FROM Productos RIGHT JOIN ProductoCliente ON Productos.Producto = ProductoCliente.Producto "
  sSql = sSql & "WHERE ProductoCliente.Cliente='" & pProve & "' AND ProductoCliente.CodCliente='" & pProd & "'"
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerUnProductoCliente = Rs
  db.Close
  
End Function

Public Function TraerValoInventario(pDesde As String, pHasta As String, pOrd As Integer) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String

  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Cantidades.Producto, Productos.Descripcion, (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber)) AS Cantidad,( Productos.PrecioCompra) as Unitario,(Productos.PrecioCompra * (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber))) as Total "
  sSql = sSql & "FROM (Cantidades LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto) "
  sSql = sSql & "Where Cantidades.Estado=1 "
  sSql = sSql & "GROUP BY Cantidades.Producto, Productos.Descripcion, Productos.PrecioCompra "
  sSql = sSql & "HAVING (((Productos.Descripcion)>='" & pDesde & " ' And (Productos.Descripcion)<='" & pHasta & "')) "
  If pOrd = 1 Then
     sSql = sSql & "ORDER BY Productos.Descripcion "
  Else
     sSql = sSql & "ORDER BY Cantidades.Producto "
  End If
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerValoInventario = Rs
  db.Close
  
End Function

Public Function TraerRentabilidad(pDesde As Date, pHasta As Date, pProDesde As String, pProHasta As String) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String

  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT DetallesComprobantes.Producto, Productos.Descripcion," 'Sum(DetallesComprobantes.Cantidad) AS SumaDeCantidad, Sum(DetallesComprobantes.PrecioTotal) AS SumaDePrecioTotal, Productos.PrecioCompra "
  sSql = sSql & "CASE "
  sSql = sSql & "WHEN CabComprobantes.Tipo=1 "
  sSql = sSql & "THEN Sum(abs(DetallesComprobantes.Cantidad)) "
  sSql = sSql & "ELSE -SUM(abs(DetallesComprobantes.Cantidad)) "
  sSql = sSql & "END AS SumaDeCantidad,"
  sSql = sSql & "CASE "
  sSql = sSql & "WHEN CabComprobantes.Tipo=1 "
  sSql = sSql & "THEN Sum(DetallesComprobantes.PrecioTotal) "
  sSql = sSql & "ELSE -SUM(DetallesComprobantes.PrecioTotal) "
  sSql = sSql & "END AS SumaDePrecioTotal,"
  sSql = sSql & "CASE "
  sSql = sSql & "WHEN CabComprobantes.Tipo=1 "
  sSql = sSql & "THEN Productos.PrecioCompra "
  sSql = sSql & "ELSE Productos.PrecioCompra END as PrecioCompra "
  sSql = sSql & "FROM (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) LEFT JOIN CabComprobantes ON DetallesComprobantes.Movimiento = CabComprobantes.Movimiento "
  sSql = sSql & "WHERE (CabComprobantes.Tipo=1 or CabComprobantes.Tipo=5) and CabComprobantes.Fecha >='" & FechaAmericana(pDesde) & "' AND  CabComprobantes.Fecha <='" & FechaAmericana(pHasta) & "' AND CabComprobantes.Anulado=0 AND CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "
  sSql = sSql & "GROUP BY CabComprobantes.Tipo, DetallesComprobantes.Producto, Productos.Descripcion, Productos.PrecioCompra,CabComprobantes.Tipo "
  sSql = sSql & "HAVING Productos.Descripcion>='" & pProDesde & "' AND Productos.Descripcion<='" & pProHasta & "' "
  sSql = sSql & "AND Sum(DetallesComprobantes.PrecioTotal)<>0 AND Productos.PrecioCompra<>0"
  
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing
  
  Set TraerRentabilidad = Rs
  db.Close
  
End Function

Public Function TraerConsolidado(pDesde As Long, pHasta As Long, pComp As Integer, bIva As Boolean) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset
  db.Open sDb

  If bIva = False Then
     sSql = "SELECT DetallesComprobantes.Producto, Productos.Descripcion, Sum(DetallesComprobantes.Cantidad) AS SumaDeCantidad, Sum(DetallesComprobantes.PrecioTotal) AS SumaDePrecioTotal "
     sSql = sSql & "FROM CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
     sSql = sSql & " Where CabComprobantes.Anulado=0 AND CabComprobantes.Tipo = 1 And CabComprobantes.Numero >=" & pDesde & " And CabComprobantes.Numero <=" & pHasta & " And CabComprobantes.Comprobante =" & pComp & " and CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "
     sSql = sSql & "GROUP BY DetallesComprobantes.Producto, Productos.Descripcion "
     sSql = sSql & "Having (((Productos.Descripcion) Is Not Null)) "
  Else
     sSql = "SELECT DetallesComprobantes.Producto, Productos.Descripcion, Sum(DetallesComprobantes.Cantidad) AS SumaDeCantidad,(Sum(DetallesComprobantes.Impuesto)+ Sum(DetallesComprobantes.PrecioTotal)) AS SumaDePrecioTotal "
     sSql = sSql & "FROM CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
     sSql = sSql & " Where CabComprobantes.Anulado=0 AND CabComprobantes.Tipo = 1 And CabComprobantes.Numero >=" & pDesde & " And CabComprobantes.Numero <=" & pHasta & " And CabComprobantes.Comprobante =" & pComp & " and CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "
     sSql = sSql & "GROUP BY DetallesComprobantes.Producto, Productos.Descripcion "
     sSql = sSql & "Having (((Productos.Descripcion) Is Not Null)) "
  End If
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set TraerConsolidado = Rs
  db.Close
  
End Function

Public Function TraerReparto(pReparto As Long, pDesde As Date, pHasta As Date) As ADODB.Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT DetallesComprobantes.Producto, Productos.Descripcion, Sum(DetallesComprobantes.Cantidad) AS Cantidad, Sum(DetallesComprobantes.PrecioTotal) AS PrecioTotal "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "WHERE CabComprobantes.Tipo=1 and CabComprobantes.Fecha >='" & FechaAmericana(pDesde) & "' AND  CabComprobantes.Fecha <='" & FechaAmericana(pHasta) & "' AND CabComprobantes.Anulado=0 "
  sSql = sSql & "AND CabComprobantes.Hasta=" & pReparto & " and CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "
  sSql = sSql & "GROUP BY DetallesComprobantes.Producto, Productos.Descripcion "
  sSql = sSql & "Having (((Productos.Descripcion) Is Not Null)) "
  sSql = sSql & "ORDER BY DetallesComprobantes.Producto"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  
  Set TraerReparto = Rs
  db.Close
  
  
End Function

Public Sub AgegrarPreciosCajas()
  Dim sSql As String
  
  db.Open sDb
  
  db.Execute "DELETE FROM PreciosCaja"

  sSql = "INSERT INTO PreciosCaja (Producto,Precio) "
  sSql = sSql & "SELECT Producto,Precio FROM Precios "
  sSql = sSql & "ORDER BY Id"
  
  db.Execute sSql
  
  db.Close

End Sub

Public Sub ActualizarPreciosCajaMovi(pProducto As String, pPrecio As Double)
  Dim sSql As String
  
  db.Open sDb
  
  sSql = "UPDATE PreciosCaja SET PreciosCaja.Precio=" & pPrecio & " "
  sSql = sSql & "WHERE Producto='" & pProducto & "'"
  
  db.Execute sSql
  
  db.Close
End Sub


Public Sub ActualizarPreciosCajas(pDesde As String, pHasta As String)
  Dim sSql As String
  
  db.Open sDb

  sSql = "UPDATE PreciosCaja SET PreciosCaja.Precio = Precios.Precio "
  sSql = sSql & "FROM Precios LEFT OUTER JOIN "
  sSql = sSql & "Productos ON Precios.Producto = Productos.Producto LEFT OUTER JOIN "
  sSql = sSql & "PreciosCaja ON Precios.Producto = PreciosCaja.Producto "
  sSql = sSql & "WHERE Productos.FechaAlta>='" & pDesde & "' AND Productos.FechaAlta<='" & pHasta & "'"
  
  db.Execute sSql
  
  db.Close
    
End Sub

Public Function TraerProductosEtiquetas(pDesde As Date, pHasta As Date) As Recordset
  Dim Rs As ADODB.Recordset, sSql As String, nListaVenta As Integer
  Dim cRsl As ClsLectura
  
  Set Rs = New ADODB.Recordset
  Set cRsl = New ClsLectura
  
  nListaVenta = cRsl.RegPorDefecto("ListaRes")
 
  db.Open sDb
  

  sSql = "SELECT Productos.Producto,Productos.Peso,Productos.Descripcion,Productos.Descripcion1, Precios.Precio,Productos.Balanza,Productos.CDepartamento,Productos.CFamilia "
  sSql = sSql & "FROM Productos LEFT JOIN Precios ON Productos.Producto = Precios.Producto "
  sSql = sSql & "WHERE Productos.FechaAlta>='" & FechaAmericana(pDesde) & "' AND Productos.FechaAlta<='" & FechaAmericana(pHasta) & "' AND Productos.FechaBaja Is Not NULL "
  sSql = sSql & "AND Precios.Lista=" & nListaVenta & " AND Productos.Activo=1 "
  sSql = sSql & "ORDER BY Productos.Descripcion"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set TraerProductosEtiquetas = Rs
  db.Close
  
End Function

Public Function TraerListaProduc(pDesde As String, X, Optional pSucursal = 0, Optional pAvan As Boolean = False, Optional pActivo As Boolean = True) As Recordset
  Dim Rs As ADODB.Recordset, sSql As String, cRsl As ClsLectura
  
  
  Set cRsl = New ClsLectura
  
  Set Rs = New ADODB.Recordset
  If db.State = 0 Then
     db.Open sDb
  End If
  sSql = "SELECT TOP 200 Productos.Producto, Productos.Descripcion, Productos.Descripcion1, (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber)) AS Cantidad "
  sSql = sSql & "FROM (Productos LEFT JOIN Cantidades ON Productos.Producto = Cantidades.Producto) LEFT JOIN Depositos ON Cantidades.Deposito = Depositos.Deposito "
  Select Case X
         Case 1
              sSql = sSql & "WHERE Productos.Producto ='" & pDesde & "' AND Productos.Tipo= 3 "
         Case 2
              If pAvan = True Then
                 sSql = sSql & "WHERE Productos.Descripcion like '%" & pDesde & "%' "
              Else
                 sSql = sSql & "WHERE Productos.Descripcion like '" & pDesde & "%' "
              End If
         Case 3
              If pAvan = True Then
                 sSql = sSql & "WHERE Productos.Descripcion1 like '%" & pDesde & "%' "
              Else
                 sSql = sSql & "WHERE Productos.Descripcion1 like '" & pDesde & "%' "
              End If
  End Select
  If pActivo = True Then
     sSql = sSql & "AND Productos.Activo=1 "
  End If
 ' sSql = sSql & "AND Depositos.Deposito=" & cRsl.RegPorDefecto("Deposito") & " "
  If pSucursal = 0 Then
     sSql = sSql & "AND Cantidades.Suc=" & nSucursal & " AND Cantidades.Estado=1 "
  Else
       sSql = sSql & "AND Cantidades.Deposito=" & pSucursal & " AND Cantidades.Estado=1 "
  End If
  
  sSql = sSql & "GROUP BY Productos.Producto, Productos.Descripcion, Productos.Descripcion1 "
  Select Case X
         Case 1
              sSql = sSql & "ORDER BY  Productos.Producto"
         Case 2
              sSql = sSql & "ORDER BY  Productos.Descripcion"
         Case 3
              sSql = sSql & "ORDER BY  Productos.Descripcion1"
  End Select
    If db.State = 0 Then
     db.Open sDb
  End If
  Rs.Open sSql, db, adOpenStatic, adLockReadOnly
  
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set TraerListaProduc = Rs

 db.Close
End Function

Public Function TraerListaCliente(pDesde As String, X, Optional pAvan As Boolean = False) As Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset

   db.Open sDb
   sSql = "SELECT TOP 100 Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio,Clientes.Observaciones "
   sSql = sSql & "FROM Clientes  "
   Select Case X
         Case 1
              If pAvan = False Then
                 sSql = sSql & "WHERE Clientes.Cliente like '" & pDesde & "%'"
              Else
                 sSql = sSql & "WHERE Clientes.Cliente like '%" & pDesde & "%'"
              End If
              sSql = sSql & "ORDER BY Clientes.Cliente"
         Case 2
              If pAvan = False Then
                 sSql = sSql & "WHERE Clientes.RazonSocial LIKE '" & pDesde & "%'"
              Else
                 sSql = sSql & "WHERE Clientes.RazonSocial LIKE '%" & pDesde & "%'"
              End If
              sSql = sSql & "ORDER BY Clientes.RazonSocial"
  End Select
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set TraerListaCliente = Rs
  db.Close
  
End Function

Public Function TraerListaProveedor(pDesde As String, X, Optional pAvan As Boolean = False) As Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset

   db.Open sDb
   sSql = "SELECT TOP 200 Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio,Proveedores.Observaciones "
   sSql = sSql & "FROM Proveedores  "
   Select Case X
         Case 1
              If pAvan = False Then
                 sSql = sSql & "WHERE Proveedores.Proveedor like '" & pDesde & "%' "
              Else
                 sSql = sSql & "WHERE Proveedores.Proveedor like '%" & pDesde & "%' "
              End If
              sSql = sSql & "ORDER BY Proveedores.Proveedor"
         Case 2
              If pAvan = False Then
                 sSql = sSql & "WHERE Proveedores.RazonSocial like '" & pDesde & "%' "
              Else
                 sSql = sSql & "WHERE Proveedores.RazonSocial like '%" & pDesde & "%' "
              End If
              sSql = sSql & "ORDER BY Proveedores.RazonSocial"
  End Select
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set TraerListaProveedor = Rs
  db.Close
  
End Function

Public Function TraerUltimoSistema(pCampo As String) As Long
   Dim Conexion As ADODB.Connection, sSql As String
   Dim Rs As ADODB.Recordset
 
   Dim Com As ADODB.Command
   Set Com = New ADODB.Command

   Set Conexion = New ADODB.Connection
   Set Rs = New ADODB.Recordset
   
   Conexion.Open "FILE NAME=" & App.Path & "\Conex.udl"
  ' db.Open sDb
   Conexion.Execute "UPDATE Sistema SET " & pCampo & "=" & pCampo & "+1"

   sSql = "SELECT " & pCampo & " as Valor FROM Sistema "
   Rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
   TraerUltimoSistema = Rs!Valor
   Conexion.Close
End Function

Public Function TraerPrecioProducto(pProducto As String, pLista As String) As Recordset
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT Productos.Producto, Productos.Descripcion AS Descripcion,Productos.FechaCompra as Fecha, Monedas.Descripcion as Moneda, Monedas.Cotizacion, Precios.Precio, Precios.Lista,(Monedas.Cotizacion * Precios.Precio) as Importe "
  sSql = sSql & "FROM Productos LEFT OUTER JOIN "
  sSql = sSql & "Precios ON Productos.Producto = Precios.Producto LEFT OUTER JOIN "
  sSql = sSql & "Monedas ON Productos.Moneda = Monedas.Id "
  sSql = sSql & "WHERE Productos.Producto='" & pProducto & "' AND Productos.Activo=1 AND Precios.Lista=" & pLista & " "
  sSql = sSql & "ORDER BY   Productos.Producto "
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.Close
  Rs.CursorType = adOpenStatic
  Rs.CursorLocation = adUseClient
  Rs.LockType = adLockBatchOptimistic
  Rs.Open
  Rs.ActiveConnection = Nothing

  Set TraerPrecioProducto = Rs
  db.Close
End Function

Public Function TraerPrecioImp(pProducto As String, pLista As String) As Currency
  Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT (Monedas.Cotizacion * Precios.Precio) as Importe "
  sSql = sSql & "FROM Productos LEFT OUTER JOIN "
  sSql = sSql & "Precios ON Productos.Producto = Precios.Producto LEFT OUTER JOIN "
  sSql = sSql & "Monedas ON Productos.Moneda = Monedas.Id "
  sSql = sSql & "WHERE Productos.Producto='" & pProducto & "' AND Precios.Lista=" & pLista & " "
  sSql = sSql & "ORDER BY   Productos.Producto "
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  If Rs.RecordCount <> 0 Then
     TraerPrecioImp = Rs!importe
  End If
  db.Close
End Function

Public Function TraerCotizacion(pProducto As String) As Double
Dim Rs As ADODB.Recordset, sSql As String
  
  Set Rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT Monedas.Cotizacion  as Valor "
  sSql = sSql & "FROM Productos LEFT OUTER JOIN "
  sSql = sSql & "Precios ON Productos.Producto = Precios.Producto LEFT OUTER JOIN "
  sSql = sSql & "Monedas ON Productos.Moneda = Monedas.Id "
  sSql = sSql & "WHERE Productos.Producto='" & pProducto & "' "
  
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  If Rs.RecordCount <> 0 Then
     TraerCotizacion = Rs!Valor
  End If
  db.Close
End Function

