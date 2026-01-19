VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProductoE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Agregar(pRsProducto As ADODB.Recordset, pRsImpuestos As ADODB.Recordset, pRsPrecios As Recordset, Optional pRep As Long = 0)
 Dim iCampos As Integer, sSql As String
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  
  db.BeginTrans
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "ProductoAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Clientes
  
  iCampos = pRsProducto.Fields.Count - 1
  For i = 0 To iCampos
      If pRsProducto(i).Name <> "Id" Then
         If pRsProducto(i).Type = adDouble Or pRsProducto(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsProducto(i).Name, adDouble, adParamInput, pRsProducto(i).DefinedSize, pRsProducto(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsProducto(i).Name, pRsProducto(i).Type, adParamInput, pRsProducto(i).DefinedSize, pRsProducto(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  ' graba los impuestos asociados
  If pRsImpuestos.RecordCount > 0 Then
     pRsImpuestos.MoveFirst
     Do While Not pRsImpuestos.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "ProductoImpuestoAgregar"
        Com.Parameters.Append Com.CreateParameter("Produ", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Impue", adInteger, adParamInput)
        Com.Parameters("Produ").Value = Trim(pRsImpuestos!Producto)
        Com.Parameters("Impue").Value = pRsImpuestos!Impuesto
        Com.Execute
    
        pRsImpuestos.MoveNext
     Loop
     Set Com = Nothing
  End If

  ' Agregar en registros en la lista de Precios
  If pRsPrecios.RecordCount > 0 Then
     pRsPrecios.MoveFirst
     Do While Not pRsPrecios.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "PreciosAgregar"
        Com.Parameters.Append Com.CreateParameter("Lista", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Porcentaje", adDouble)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble)
        Com.Parameters("Lista").Value = pRsPrecios!ListaPrecio
        Com.Parameters("Producto").Value = pRsProducto!Producto
        Com.Parameters("Porcentaje").Value = 0
        Com.Parameters("Precio").Value = 0
        Com.Execute
        pRsPrecios.MoveNext
     Loop
     Set Com = Nothing
  End If

  Set Com = New ADODB.Command

  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CantidadesAgregar"
  Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput)
  Com.Parameters.Append Com.CreateParameter("Movimiento", adInteger, adParamInput)
  Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 1)
  Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
  Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput)
  Com.Parameters.Append Com.CreateParameter("CantDebe", adDouble)
  Com.Parameters.Append Com.CreateParameter("CantHaber", adDouble)
  Com.Parameters.Append Com.CreateParameter("Anulado", adBoolean, adParamInput, , 0)
  Com.Parameters("Suc").Value = nSucursal
  Com.Parameters("Movimiento").Value = 0
  Com.Parameters("Producto").Value = pRsProducto!Producto
  Com.Parameters("Deposito").Value = pRsProducto!Deposito
  Com.Parameters("CantDebe").Value = 0
  Com.Parameters("CantHaber").Value = 0

  Com.Execute
    
  sSql = "INSERT INTO ProductoUnidades ( Producto, UnidadA, CantidadA, UnidadB, CantidadB ) "
  sSql = sSql & "VALUES ('" & pRsProducto!Producto & "'," & pRsProducto!UMCompra & "," & pRsProducto!CantDespacho & ","
  sSql = sSql & pRsProducto!UMVenta & "," & pRsProducto!CantVenta & ")"
  
  db.Execute sSql
  
  If pRep <> 0 Then
     db.Execute "UPDATE PRODUCTOS SET Representante=" & pRep & " WHERE Id=" & pRep
  End If
  
  Set Com = Nothing
  
  db.CommitTrans
  
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Public Sub Modificar(pRsProducto As ADODB.Recordset, pRsImpuestos As ADODB.Recordset, Cod As String, Optional pRep As Long = 0)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "ProductoModif"
    
  
  
 ' Com.Parameters.Append Com.CreateParameter("Param", adVarChar, adParamInput, Len(Cod), Cod)
' recorre la coleccion fields y carga los parametros de la tabla de cliente

  iCampos = pRsProducto.Fields.Count - 1
  For i = 0 To iCampos
      If pRsProducto(i).Name <> "Id" Then
         If pRsProducto(i).Type = adDouble Or pRsProducto(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsProducto(i).Name, adDouble, adParamInput, pRsProducto(i).DefinedSize, pRsProducto(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsProducto(i).Name, pRsProducto(i).Type, adParamInput, pRsProducto(i).DefinedSize, pRsProducto(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  ' primero borra los impuestos
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "ProductoImpuestoBorrar"
  
  Com.Parameters.Append Com.CreateParameter("Cod", adChar, adParamInput, 20, pRsProducto!Producto)
  Com.Execute
  Set Com = Nothing
  
  ' Graba los impuestos asociados
  
  If pRsImpuestos.RecordCount > 0 Then
     pRsImpuestos.MoveFirst
     Do While Not pRsImpuestos.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "ProductoImpuestoAgregar"
        Com.Parameters.Append Com.CreateParameter("Produ", adChar, adParamInput, 20)
        Com.Parameters.Append Com.CreateParameter("Imp", adSmallInt, adParamInput)
        Com.Parameters("Produ").Value = pRsImpuestos!Producto
        Com.Parameters("Imp").Value = pRsImpuestos!Impuesto
        Com.Execute
        pRsImpuestos.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  Dim sSql As String
  
  sSql = "UPDATE ProductoUnidades SET UnidadA=" & pRsProducto!UMCompra & ","
  sSql = sSql & "CantidadA=" & pRsProducto!CantDespacho & ","
  sSql = sSql & "UnidadB=" & pRsProducto!UMVenta & ","
  sSql = sSql & "CantidadB=" & pRsProducto!CantVenta & " "
  sSql = sSql & "WHERE Producto='" & pRsProducto!Producto & "'"
    
  db.Execute sSql

  If pRep <> 0 Then
     db.Execute "UPDATE PRODUCTOS SET Representante=" & pRep & " WHERE Id=" & pRep
  End If

  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
    
End Sub

Public Sub Borrar(pStoreProc As String, pProdu As String, Cod As Variant, pDepo As Integer, Optional pRsParam As Recordset)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
  
  db.Open sDb
  db.BeginTrans
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = pStoreProc
    
    ' recorre la coleccion fields y carga los parametros
  If Cod <> 0 Then
     If IsNumeric(Cod) Then
        Com.Parameters.Append Com.CreateParameter("Param", adInteger, adParamInput, , Cod)
     Else
        Com.Parameters.Append Com.CreateParameter("Param", adVarChar, adParamInput, Len(Cod), Cod)
     End If
  End If
  If Not pRsParam Is Nothing Then
     iCampos = pRsParam.Fields.Count - 1
     For i = 0 To iCampos
         If pRsParam(i).Name <> "Id" Then
         If pRsParam(i).Type = adDouble Or pRsParam(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsParam(i).Name, adDouble, adParamInput, pRsParam(i).DefinedSize, pRsParam(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsParam(i).Name, pRsParam(i).Type, adParamInput, pRsParam(i).DefinedSize, pRsParam(i).Value)
         End If
      End If
     Next
  End If
  Com.Execute
  
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "ProductoImpuestoBorrar"
  
  Com.Parameters.Append Com.CreateParameter("Cod", adChar, adParamInput, 15, Cod)
  Com.Execute
  Set Com = Nothing
  
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CantidadesBorrar"
  
  Com.Parameters.Append Com.CreateParameter("Cod", adChar, adParamInput, 15, pProdu)
  Com.Parameters.Append Com.CreateParameter("Dep", adInteger, adParamInput, , pDepo)
  Com.Execute
  
  db.Execute "DELETE FROM Precios WHERE Producto='" & pProdu & "'"
  
  db.Execute "DELETE FROM ProductoUnidades WHERE Producto='" & pProdu & "'"
  
  db.Execute "DELETE FROM PreciosCaja WHERE Producto='" & pProdu & "'"
    
  Set Com = Nothing
 
  db.CommitTrans
  db.Close
 Exit Sub
    
errHandler:
    err.Raise err.Number
    db.RollbackTrans
    db.Close
End Sub

Public Sub GrabarProductoProveedor(pProvee As String, pProdu As String, pProdProv As String)
  Dim sSql As String
  
   db.Open sDb

  sSql = "INSERT INTO ProductoProveedor (Producto,Proveedor,CodProvee) VALUES('" & pProdu & "',"
  sSql = sSql & "'" & pProvee & "',"
  sSql = sSql & "'" & pProdProv & "')"
  
  db.Execute sSql
  db.Close
End Sub

Public Sub ModificarProductoProveedor(pProvee As String, pProdu As String, pProdProv As String)
  Dim sSql As String
  
  db.Open sDb

  sSql = "UPDATE ProductoProveedor SET CodProvee='" & pProdProv & "' "
  sSql = sSql & "WHERE Proveedor='" & pProvee & "' "
  sSql = sSql & "AND Producto='" & pProdu & "'"
  
  db.Execute sSql
  db.Close
End Sub

Public Sub GrabarProductoCliente(pProvee As String, pProdu As String, pProdProv As String)
  Dim sSql As String
  
  db.Open sDb

  sSql = "INSERT INTO ProductoCliente (Producto,Cliente,CodCliente) VALUES('" & pProdu & "',"
  sSql = sSql & "'" & pProvee & "',"
  sSql = sSql & "'" & pProdProv & "')"
  
  db.Execute sSql
  db.Close
End Sub

Public Sub ModificarProductoCliente(pProvee As String, pProdu As String, pProdProv As String)
  Dim sSql As String
  
  sSql = "UPDATE ProductoCliente SET CodCliente='" & pProdProv & "' "
  sSql = sSql & "WHERE Cliente='" & pProvee & "' "
  sSql = sSql & "AND Producto='" & pProdu & "'"
  
  db.Execute sSql
  
End Sub

Public Function AgregarImportacion(pRsProducto As ADODB.Recordset, pRsImpuestos As ADODB.Recordset, pRsPrecios As Recordset) As Boolean
  Dim iCampos As Integer, sSql As String
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
    
  db.Open sDb
  
  If pRsProducto.RecordCount > 0 Then

    db.BeginTrans
     
    pRsProducto.MoveFirst
    Do While Not pRsProducto.EOF
       Set Com = New ADODB.Command
       Com.ActiveConnection = db
       Com.CommandType = adCmdStoredProc
       Com.CommandText = "ProductoAgregar"
      
    ' recorre la coleccion fields y carga los parametros de la tabla de Clientes
  
      iCampos = pRsProducto.Fields.Count - 1
      For i = 0 To iCampos
        If pRsProducto(i).Name <> "Id" Then
           If pRsProducto(i).Type = adDouble Or pRsProducto(i).Type = adNumeric Then
              Com.Parameters.Append Com.CreateParameter(pRsProducto(i).Name, adDouble, adParamInput, pRsProducto(i).DefinedSize, pRsProducto(i).Value)
           Else
              Com.Parameters.Append Com.CreateParameter(pRsProducto(i).Name, pRsProducto(i).Type, adParamInput, pRsProducto(i).DefinedSize, pRsProducto(i).Value)
           End If
        End If
      Next
     Com.Execute
     Set Com = Nothing
     
     Set Com = New ADODB.Command
  
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "CantidadesAgregar"
     Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Movimiento", adInteger, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 1)
     Com.Parameters.Append Com.CreateParameter("Produ", adVarChar, adParamInput, 25)
     Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput)
     Com.Parameters.Append Com.CreateParameter("CantDebe", adSingle)
     Com.Parameters.Append Com.CreateParameter("CantHaber", adSingle)
     Com.Parameters.Append Com.CreateParameter("Anulado", adInteger)
     Com.Parameters("Suc").Value = nSucursal
     Com.Parameters("Movimiento").Value = 0
     Com.Parameters("Produ").Value = pRsProducto!Producto
     Com.Parameters("Deposito").Value = pRsProducto!Deposito
     Com.Parameters("CantDebe").Value = pRsProducto!CantMaxima
     Com.Parameters("CantHaber").Value = 0
     Com.Parameters("Anulado").Value = 0
     Com.Execute
      
     sSql = "INSERT INTO ProductoUnidades ( Producto, UnidadA, CantidadA, UnidadB, CantidadB ) "
     sSql = sSql & "VALUES ('" & pRsProducto!Producto & "'," & pRsProducto!UMCompra & "," & pRsProducto!CantDespacho & ","
     sSql = sSql & pRsProducto!UMVenta & "," & pRsProducto!CantVenta & ")"
    
     db.Execute sSql
    Set Com = Nothing
    
     pRsProducto.MoveNext
    Loop
  ' graba los impuestos asociados
    If pRsImpuestos.RecordCount > 0 Then
       pRsImpuestos.MoveFirst
       Do While Not pRsImpuestos.EOF
          Set Com = New ADODB.Command
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "ProductoImpuestoAgregar"
          Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
          Com.Parameters.Append Com.CreateParameter("Impuesto", adInteger, adParamInput)
          Com.Parameters("Producto").Value = Trim(pRsImpuestos!Producto)
          Com.Parameters("Impuesto").Value = pRsImpuestos!Impuesto
          Com.Execute
      
          pRsImpuestos.MoveNext
       Loop
       Set Com = Nothing
    End If
  
    ' Agregar en registros en la lista de Precios
    If pRsPrecios.RecordCount > 0 Then
       pRsPrecios.MoveFirst
       Do While Not pRsPrecios.EOF
          Set Com = New ADODB.Command
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "PreciosAgregar"
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "PreciosAgregar"
          Com.Parameters.Append Com.CreateParameter("Lista", adInteger, adParamInput)
          Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
          Com.Parameters.Append Com.CreateParameter("Porcentaje", adDouble)
          Com.Parameters.Append Com.CreateParameter("Precio", adDouble)
          Com.Parameters("Lista").Value = pRsPrecios!Lista
          Com.Parameters("Producto").Value = pRsPrecios!Producto
          Com.Parameters("Porcentaje").Value = 0
          Com.Parameters("Precio").Value = pRsPrecios!precio
          Com.Execute
          pRsPrecios.MoveNext
       Loop
       Set Com = Nothing
    End If
    
    db.CommitTrans
  End If
  db.Close
  AgregarImportacion = True
  Exit Function
    
errHandler:
    db.RollbackTrans
    db.Close
    MsgBox err.Description, vbCritical, "" ' pRsProducto!Producto
    AgregarImportacion = False
End Function

