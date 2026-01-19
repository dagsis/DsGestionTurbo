VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DameProductoCosto(pFami As Integer, pProv As Long) As Recordset
  Dim sSql As String, Rs As ADODB.Recordset
 
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Productos.Producto, Productos.Descripcion,Productos.Descripcion,Productos.Descripcion1 ,Productos.PrecioCompra as Precio,FechaCompra as Fecha "
  sSql = sSql & "From Productos "
  If pFami <> 0 Then
     sSql = sSql & "WHERE Productos.Familia=" & pFami & " AND Productos.Activo=1 "
  End If
  If pProv <> 0 Then
     sSql = sSql & "WHERE Productos.Proveedor=" & pProv & " AND Productos.Activo=1 "
  End If
        
  sSql = sSql & "ORDER BY Productos.Producto "
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenStatic, adLockPessimistic
  Rs.ActiveConnection = Nothing
  Set DameProductoCosto = Rs
  db.Close
  
End Function

Public Function DameProductos(pLista As Integer, pDesde As String, pHasta As String, Optional pFami As Integer, Optional pOder As Boolean, Optional pProv As Long, Optional pMarca As Long) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
  db.Open sDb

 If pDesde = "" Then
    sSql = "SELECT Precios.Producto, Productos.Descripcion,Productos.Descripcion1,Precios.Porcentaje,Precios.Precio,Productos.FechaAlta as Fecha "
    sSql = sSql & "FROM Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto "
    If pLista <> 0 Then
       sSql = sSql & "WHERE lista=" & pLista & " AND Precios.Precio<> 0 AND Productos.Activo=1 "
    End If
    If pFami <> 0 Then
       sSql = sSql & " AND Productos.Familia=" & pFami
    End If
    If pMarca <> 0 Then
       sSql = sSql & " AND Productos.Marca=" & pMarca
    End If
    If pProv <> 0 Then
       sSql = sSql & "AND Productos.Proveedor=" & pProv
    End If
    If pOder = False Then
       sSql = sSql & " ORDER BY Productos.Descripcion"
    Else
       sSql = sSql & " ORDER BY Precios.Producto"
    End If
 Else
    sSql = "SELECT Precios.Producto, Productos.Descripcion,Productos.Descripcion,Precios.Porcentaje ,Precios.Precio "
    sSql = sSql & "FROM Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto "
    sSql = sSql & "WHERE lista=" & pLista & " and Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta & "' AND Productos.Activo=1 "
    If pOder = False Then
       sSql = sSql & " ORDER BY Productos.Descripcion"
    Else
       sSql = sSql & " ORDER BY Precios.Producto"
    End If
 End If
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.ActiveConnection = Nothing
 Set DameProductos = Rs
 db.Close
End Function
Public Function DameListaProductos(pLista As Integer, pDesde As String, pHasta As String, pOrden As Integer) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
 db.Open sDb

' If pOrden = 3 Then
    sSql = "SELECT Precios.Producto, Productos.Descripcion,Productos.FechaCompra,Productos.Moneda ,Precios.Precio "
    sSql = sSql & "FROM Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto "
    sSql = sSql & "WHERE lista=" & pLista
    If pOrden = 1 Then
       sSql = sSql & " and Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta
    Else
       sSql = sSql & " and Productos.Producto>='" & pDesde & "' and Productos.Producto <='" & pHasta
    End If
    sSql = sSql & "' AND Productos.Activo=1 "
    If pOrden = 0 Then
       sSql = sSql & "ORDER BY Precios.Producto"
    Else
        sSql = sSql & "ORDER BY Productos.Descripcion"
    End If
' Else
'    sSql = "SELECT Precios.Producto, Productos.Descripcion,Productos.Moneda, Precios.Precio "
'    sSql = sSql & "FROM Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto "
'    sSql = sSql & "WHERE lista=" & pLista & " and Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta & "' AND Productos.Activo=1 "
'    sSql = sSql & "ORDER BY Productos.Descripcion"
' End If
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Rs.ActiveConnection = Nothing
 
 Set DameListaProductos = Rs
 db.Close
End Function

Public Function DameListaProductosGrupo(pLista As Integer, pDesde As String, pHasta As String, pOrden As Integer, pFaDes As String, pFaHas As String) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
  db.Open sDb

' If pOrden = 3 Then
    sSql = "SELECT DISTINCT  Precios.Producto ,Productos.Descripcion, Precios.Precio, Familias.Descripcion,Productos.Moneda "
    sSql = sSql & "FROM Familias RIGHT JOIN (Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto) ON Familias.Familia = Productos.Familia "
    sSql = sSql & "WHERE Precios.lista=" & pLista
    If pOrden = 1 Then
       sSql = sSql & " and Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta
    Else
       sSql = sSql & " and Productos.Producto>='" & pDesde & "' and Productos.Producto <='" & pHasta
    End If
    sSql = sSql & "' and Familias.Descripcion>='" & pFaDes & "' and Familias.Descripcion<='" & pFaHas & "' AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Familias.Descripcion,Productos.Descripcion"
' Else
'    sSql = "SELECT Precios.Producto, Productos.Descripcion, Precios.Precio, Familias.Descripcion,Productos.Moneda "
'    sSql = sSql & "FROM Familias RIGHT JOIN (Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto) ON Familias.Familia = Productos.Familia "
'    sSql = sSql & "WHERE lista=" & pLista & " and Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta & "' "
'    sSql = sSql & " and Familias.Descripcion>='" & pFaDes & "' and Familias.Descripcion<='" & pFaHas & "' AND Productos.Activo=1 "
'    sSql = sSql & "ORDER BY Familias.Descripcion,Productos.Descripcion"
' End If
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Rs.ActiveConnection = Nothing
 Set DameListaProductosGrupo = Rs
 db.Close
End Function
Public Function DameListaProductosMarca(pLista As Integer, pDesde As String, pHasta As String, pOrden As Integer, pMaDes As String, pMaHas As String) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
 db.Open sDb

' If pOrden = 3 Then
    sSql = "SELECT DISTINCT  Precios.Producto, Productos.Descripcion, Precios.Precio, Marcas.Descripcion,Productos.Moneda "
    sSql = sSql & "FROM Marcas RIGHT JOIN (Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto) ON Marcas.Marca = Productos.Marca "
    sSql = sSql & "WHERE lista=" & pLista
    If pOrden = 1 Then
       sSql = sSql & " and Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta & "' "
    Else
       sSql = sSql & " and Productos.Producto>='" & pDesde & "' and Productos.Producto <='" & pHasta & "' "
    End If
    sSql = sSql & " and Marcas.Descripcion>='" & pMaDes & "' and Marcas.Descripcion<='" & pMaHas & "' AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Marcas.Descripcion,Productos.Descripcion"
' Else
'    sSql = "SELECT Precios.Producto, Productos.Descripcion, Precios.Precio, Marcas.Descripcion,Productos.Moneda "
'    sSql = sSql & "FROM Marcas RIGHT JOIN (Productos RIGHT JOIN Precios ON Productos.Producto = Precios.Producto) ON Marcas.Marca = Productos.Marca "
'    sSql = sSql & "WHERE lista=" & pLista & " and Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta & "' "
'    sSql = sSql & " and Marcas.Descripcion>='" & pMaDes & "' and Marcas.Descripcion<='" & pMaHas & "' AND Productos.Activo=1 "
'    sSql = sSql & "ORDER BY Marcas.Descripcion,Productos.Descripcion"
' End If
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Rs.ActiveConnection = Nothing

 Set DameListaProductosMarca = Rs
 db.Close
End Function


Public Sub BorrarPrecios(pLista As Integer)
  Dim sSql As String
  On Error GoTo Errores
  db.Open sDb

  sSql = "DELETE FROM Precios WHERE lista=" & pLista
  db.BeginTrans
  db.Execute sSql
  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
End Sub

Public Sub ActulizarPrecios(pLista As Integer, pProduc As String, pPrecio As Double, Optional pRep As Long = 0)
  Dim sSql As String, cRsl As ClsLectura
  
  On Error GoTo Errores
  Dim sConex As Connection
  
  Set sConex = New Connection
  
  Set cRsl = New ClsLectura
  sConex.Open sDb
  
  sConex.BeginTrans

  Dim cPrec As Single

  cPrec = pPrecio
  
  sSql = "Update Precios SET Precio=" & cPrec & " WHERE Lista=" & pLista & " and Producto='" & pProduc & "'"
  
  sConex.Execute sSql
  
  sSql = "Update Productos SET Actualiza=0,"
  sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
  sSql = sSql & "WHERE Producto='" & pProduc & "' OR Productos.Id=" & pRep & " "
  sConex.Execute sSql
  
  If pRep <> 0 Then
     
    sSql = "UPDATE Precios set Precios.Precio=" & cPrec & " "
    sSql = sSql & "FROM Precios LEFT OUTER JOIN  Productos ON Precios.Producto = Productos.Producto "
    sSql = sSql & "Where Productos.Representante =" & pRep & " OR Productos.Id=" & pRep & " "
    sConex.Execute sSql
    
    sSql = "Update Productos SET PrecioCompra=" & pPrecio & ",Actualiza=0, "
    sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
    sSql = sSql & "Where Representante =" & pRep & " "
    sConex.Execute sSql
   
  End If
  
  sConex.CommitTrans
  sConex.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
  sConex.Close

End Sub

Public Sub ActulizarPreciosCosto(pListaPrec As Integer, pProduc As String, pPrecio As Double, pPorcentaje As Double, pVenta As Double, Optional pRep As Long = 0)
 Dim sSql As String, sProv As String
  On Error GoTo Errores
  db.Open sDb

  db.BeginTrans
  
  sSql = "Update Productos SET PrecioCompra=" & pPrecio & ",Actualiza=0, "
  sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
  sSql = sSql & " WHERE Producto='" & pProduc & "'"
  db.Execute sSql

    
  Dim cPrec As Single

  cPrec = pVenta
  db.Execute "UPDATE Precios SET Precio=" & cPrec & ",Porcentaje=" & pPorcentaje & " WHERE Producto='" & pProduc & "' and Lista=" & pListaPrec
    
  If pRep <> 0 Then
     
    sSql = "UPDATE Precios set Precios.Precio=" & cPrec & ",Porcentaje=" & pPorcentaje & " "
    sSql = sSql & "FROM Precios LEFT OUTER JOIN  Productos ON Precios.Producto = Productos.Producto "
    sSql = sSql & "Where Productos.Representante =" & pRep
    db.Execute sSql

    sSql = "Update Productos SET PrecioCompra=" & pPrecio & ",Actualiza=0, "
    sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
    sSql = sSql & "Where Representante =" & pRep
    db.Execute sSql
     
  End If
  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
  db.Close
End Sub

Public Sub ActulizarPreciosCostoLista(pProduc As String, pPrecio As Double, pSubTotal As Double, pValor As Double, pCondi As String)
 Dim sSql As String, sProv As String, cRsl As ClsLectura
  On Error GoTo Errores
  db.Open sDb

  db.BeginTrans
  
  If pValor = 0 Then
      If pCondi <> "Todos" Then
         sSql = "Update Productos SET PrecioCompra=" & pPrecio & ",Actualiza=1,"
      Else
         sSql = "Update Productos SET PrecioCompra=PrecioCompra,Actualiza=1,"
      End If
  Else
      sSql = "Update Productos SET PrecioCompra=PrecioCompra * " & ((pValor / 100) + 1) & ",Actualiza=1,"
  End If
  sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
  If pCondi = "" Then
     sSql = sSql & " WHERE Producto='" & pProduc & "'"
  Else
     If pCondi <> "Todos" Then
        sSql = sSql & "WHERE " & pCondi
     End If
  End If
  
  db.Execute sSql
  
  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
  db.Close
End Sub


Public Sub ActulizarPreciosCostoMen(pProduc As String, pPrecio As Currency)
  Dim sSql As String, sProv As String
  On Error GoTo Errores
  db.Open sDb

  sSql = "Update Productos SET PrecioCompra=" & pPrecio & ","
  sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "' "
  sSql = sSql & " WHERE Producto='" & pProduc & "'"
  db.BeginTrans
  
  db.Execute sSql
  
  sSql = "Update Productos SET PrecioCompra=" & pPrecio & ","
  sSql = sSql & "FechaBaja='" & FechaAmericana(Date) & "' "
  sSql = sSql & " WHERE Producto='" & pProduc & "'"
  db.Execute sSql

  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
  db.Close
End Sub

Public Sub ActulizarImporte(pLista As Integer, pProduc As String, pPrecio As Currency)
  Dim sSql As String
  On Error GoTo Errores
  db.Open sDb

  sSql = "Update Precios SET Precio=Precio +" & pPrecio & " WHERE Lista=" & pLista & " and Producto='" & pProduc & "'"
  db.BeginTrans
  db.Execute sSql
  sSql = "UPDATE Productos SET FechaAlta='" & FechaAmericana(Date) & "' WHERE Producto='" & pProduc & "'"
  db.Execute sSql
  
  sSql = "UPDATE Productos SET FechaBaja='" & FechaAmericana(Date) & "' WHERE Producto='" & pProduc & "'"
  db.Execute sSql
  
  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
  db.Close
End Sub

Public Sub ActulizarImporteCosto(pProduc As String, pPrecio As Currency)
  Dim sSql As String
  On Error GoTo Errores
  db.Open sDb

  sSql = "Update Productos SET PrecioCompra=PrecioCompra +" & pPrecio & ","
  sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "' "
  sSql = sSql & " WHERE  Producto='" & pProduc & "'"
  db.BeginTrans
  db.Execute sSql
  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
End Sub

Public Sub ActulizarPorcentaje(pLista As Integer, pProduc As String, pPrecio As Double)
  Dim sSql As String
  On Error GoTo Errores
  db.Open sDb

  sSql = "Update Precios SET Precio=" & pPrecio & " WHERE Lista=" & pLista & " and Producto='" & pProduc & "'"
  
  db.BeginTrans
  db.Execute sSql
  sSql = "UPDATE Productos SET FechaAlta='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "'  WHERE Producto='" & pProduc & "' AND Productos.Activo=1"
  db.Execute sSql
  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
End Sub

Public Sub ActulizarPorcentajeCosto(pProduc As String, pPrecio As Currency)
  Dim sSql As String
  On Error GoTo Errores
  db.Open sDb

  sSql = "Update Productos SET PrecioCosto=PrecioCosto +(PrecioCosto * " & pPrecio / 100 & "), "
  sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "' "
  sSql = sSql & "WHERE  Producto='" & pProduc & "' AND Productos.Activo=1"
  db.BeginTrans
  db.Execute sSql
  db.CommitTrans
  db.Close
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
End Sub

Public Sub CopiarPrecios(pRs As Recordset, pLista As Integer)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo Errores
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  If pRs.RecordCount > 0 Then
     pRs.MoveFirst
     Do While Not pRs.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "PreciosCopiar"
        Com.Parameters.Append Com.CreateParameter("Lista", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Porcentaje", adDouble)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble)
        Com.Parameters("Lista").Value = pLista
        Com.Parameters("Producto").Value = pRs!codigo
        Com.Parameters("Porcentaje").Value = 0
        Com.Parameters("Precio").Value = pRs!precio
        Com.Execute
        pRs.MoveNext
     Loop
  End If
  db.CommitTrans
  db.Close
  Set Com = Nothing
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
  db.Close
End Sub

Public Function DameCostoProductos(pDesde As String, pHasta As String, pOrden As Integer) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
 db.Open sDb

 If pOrden = 3 Then
    sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.PrecioCompra "
    sSql = sSql & "FROM Productos "
    sSql = sSql & "WHERE Productos.Producto>='" & pDesde & "' and Productos.Producto <='" & pHasta & "' AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Productos.Producto"
 Else
    sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.PrecioCompra "
    sSql = sSql & "FROM Productos "
    sSql = sSql & "WHERE Productos.Producto>='" & pDesde & "' and Productos.Producto <='" & pHasta & "' AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Productos.Descripcion"
 End If
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.ActiveConnection = Nothing
 Set DameCostoProductos = Rs
 db.Close
End Function

Public Function DameCostoProductosMarca(pDesde As String, pHasta As String, pOrden As Integer, pMaDes As String, pMaHas As String) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
 db.Open sDb

 If pOrden = 3 Then
    sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.PrecioCompra, Marcas.Descripcion "
    sSql = sSql & "FROM Marcas RIGHT JOIN Productos ON Marcas.Marca = Productos.Marca "
    sSql = sSql & "WHERE  Productos.Producto>='" & pDesde & "' and Productos.Producto <='" & pHasta & "' "
    sSql = sSql & " and Marcas.Descripcion>='" & pMaDes & "' and Marcas.Descripcion<='" & pMaHas & "' AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Marcas.Descripcion,Productos.Producto"
 Else
    sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.PrecioCompra, Marcas.Descripcion "
    sSql = sSql & "FROM Marcas RIGHT JOIN Productos ON Marcas.Marca = Productos.Marca "
    sSql = sSql & "WHERE  Productos.Producto>='" & pDesde & "' and Productos.Producto <='" & pHasta & "' "
    sSql = sSql & " and Marcas.Descripcion>='" & pMaDes & "' and Marcas.Descripcion<='" & pMaHas & "' AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Marcas.Descripcion,Productos.Descripcion"
 End If
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.ActiveConnection = Nothing
 Set DameCostoProductosMarca = Rs
 db.Close
End Function

Public Function DameCostoProductosGrupo(pDesde As String, pHasta As String, pOrden As Integer, pFaDes As String, pFaHas As String) As Recordset
 Dim sSql As String, Rs As ADODB.Recordset
 
 Set Rs = New ADODB.Recordset
 db.Open sDb

 If pOrden = 3 Then
    sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.PrecioCompra, Familias.Descripcion "
    sSql = sSql & "FROM Familias RIGHT JOIN Productos ON Familias.Familia = Productos.Familia "
    sSql = sSql & "WHERE  Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta & "' "
    sSql = sSql & " and Familias.Descripcion>='" & pFaDes & "' and Familias.Descripcion<='" & pFaHas & "' AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Familias.Descripcion,Productos.Producto"
 Else
    sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.PrecioCompra, Familias.Descripcion "
    sSql = sSql & "FROM Familias RIGHT JOIN Productos ON Familias.Familia = Productos.Familia "
    sSql = sSql & "WHERE  Productos.Descripcion>='" & pDesde & "' and Productos.Descripcion <='" & pHasta & "' "
    sSql = sSql & " and Familias.Descripcion>='" & pFaDes & "' and Familias.Descripcion<='" & pFaHas & "'  AND Productos.Activo=1 "
    sSql = sSql & "ORDER BY Familias.Descripcion,Productos.Descripcion"
 End If
  Rs.CursorLocation = adUseClient
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs.ActiveConnection = Nothing
 Set DameCostoProductosGrupo = Rs
 db.Close
End Function

Public Function DameCostoLista(pProDes As String, pProHas As String, pLis As Integer, pOrd As Byte)
 Dim sSql As String, Rs As ADODB.Recordset
 
  Set Rs = New ADODB.Recordset
  db.Open sDb
  
  sSql = "SELECT Productos.Producto, Productos.Descripcion, Precios.Precio, Productos.PrecioCompra "
  sSql = sSql & "FROM Productos LEFT JOIN Precios ON Productos.Producto = Precios.Producto "
  sSql = sSql & "WHERE  Productos.Activo=1 AND  Productos.Descripcion>='" & pProDes & "' and Productos.Descripcion <='" & pProHas & "' and  Precios.Lista = " & pLis
  If pOrd = 1 Then
     sSql = sSql & " ORDER BY Productos.Descripcion"
  Else
     sSql = sSql & " ORDER BY Productos.Producto"
  End If
  
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Set DameCostoLista = Rs
  db.Close
End Function

Public Function DameVentaCosto(pProDes As String, pProHas As String, pDes As Date, pHas As Date, pOrd As Byte)
  Dim sSql As String, Rs As ADODB.Recordset
 
  Set Rs = New ADODB.Recordset

  pDes = FechaAmericana(pDes)
  pHas = FechaAmericana(pHas)
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Fecha, DetallesComprobantes.Producto, Productos.Descripcion, Sum(DetallesComprobantes.Cantidad) AS SumaDeCantidad, DetallesComprobantes.PrecioUnitario, Productos.PrecioCompra "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Productos ON DetallesComprobantes.Producto = Productos.Producto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "WHERE  Productos.Descripcion>='" & pProDes & "' and Productos.Descripcion <='" & pProHas & "' "
  sSql = sSql & "AND CabComprobantes.Fecha>='" & pDes & "' and CabComprobantes.Fecha<='" & pHas & "' AND Productos.Activo=1 "
  sSql = sSql & "GROUP BY CabComprobantes.Fecha,  DetallesComprobantes.Producto, Productos.Descripcion,DetallesComprobantes.PrecioUnitario ,Productos.PrecioCompra "
  Select Case pOrd
         Case 0
              sSql = sSql & " ORDER BY  DetallesComprobantes.Producto"
         Case 1
              sSql = sSql & " ORDER BY Productos.Descripcion"
  End Select
  
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  Set DameVentaCosto = Rs
  db.Close
End Function

