VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsClienteL"
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

Public Function TraerDatosCliente(sClie As String, Optional pId As Long = 0) As Recordset
  On Error GoTo errHandler
  EnsureConn

  Dim Rs As ADODB.Recordset
  Set Rs = New ADODB.Recordset

    Dim sSql As String
    
    sSql = "SELECT Clientes.*,"
    sSql = sSql & "ListaPrecio,ListaPrecio,ListadePrecio.Descripcion AS ListaDesc, ListadePrecio.Porcentaje as ListaPorc, ListadePrecio.PrecioIva, Descuentos.Porcentaje AS DescPorc, Zonas.Descripcion AS ZonaDesc,"
    sSql = sSql & "Descuentos.Descripcion AS DescDesc, Transporte.Descripcion AS TransDesc, TipoIva.Descripcion AS TIvaDesc,"
    sSql = sSql & "Vendedores.Descripcion as VendDesc,Vendedores.Comision as VendComision "
    sSql = sSql & "FROM Clientes LEFT OUTER JOIN "
    sSql = sSql & "Vendedores ON Clientes.Vendedor = Vendedores.Vendedor LEFT OUTER JOIN "
    sSql = sSql & "Descuentos ON Clientes.Descuento = Descuentos.Descuento LEFT OUTER JOIN "
    sSql = sSql & "Zonas ON Clientes.Zona = Zonas.Zona LEFT OUTER JOIN "
    sSql = sSql & "Recargos ON Clientes.Recargo = Recargos.Recargo LEFT OUTER JOIN "
    sSql = sSql & "Transporte ON Clientes.Id = Transporte.Id LEFT OUTER JOIN "
    sSql = sSql & "ListadePrecio ON Clientes.Precio = ListadePrecio.ListaPrecio LEFT OUTER JOIN "
    sSql = sSql & "TipoIva ON Clientes.TipoIva = TipoIva.TipoIva LEFT OUTER JOIN "
    sSql = sSql & "Provincias ON Clientes.Provincia = Provincias.Provincia "
    If pId = 0 Then
       sSql = sSql & "WHERE Clientes.Cliente='" & sClie & "'"
    Else
       sSql = sSql & "WHERE Clientes.Id=" & pId
    End If
    
    Rs.CursorLocation = adUseClient
    Rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
    Rs.ActiveConnection = Nothing
         
    Set TraerDatosCliente = Rs
    Exit Function
   
errHandler:
   err.Raise err.Number
End Function

Public Function ClientesAlfabetico(nZona As Integer, Optional bOrd As Boolean) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
   
  db.Open sDb
 
  If bOrd = False Then
     If nZona = 0 Then
        sSql = "SELECT Clientes.Id, Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Clientes.Cuit "
        sSql = sSql & "FROM TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Localidades RIGHT JOIN Clientes ON Localidades.Localidad = Clientes.Localidad) ON Provincias.Provincia = Clientes.Provincia) ON TipoIva.TipoIva = Clientes.TipoIva "
        sSql = sSql & "ORDER BY RazonSocial"
     Else
        sSql = "SELECT Clientes.Id, Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Clientes.Cuit "
        sSql = sSql & "FROM TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Localidades RIGHT JOIN Clientes ON Localidades.Localidad = Clientes.Localidad) ON Provincias.Provincia = Clientes.Provincia) ON TipoIva.TipoIva = Clientes.TipoIva "
        sSql = sSql & "WHERE Zona=" & nZona
        sSql = sSql & " ORDER BY RazonSocial"
    End If
  Else
      sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Contacto, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion as Provincia, Clientes.Telefono, TipoIva.Descripcion as TipoIva, Clientes.Cuit, Clientes.IngBruto, Zonas.Descripcion, ListadePrecio.Descripcion "
      sSql = sSql & "FROM ListadePrecio RIGHT JOIN (Zonas RIGHT JOIN (Localidades RIGHT JOIN (Provincias RIGHT JOIN (TipoIva LEFT JOIN Clientes ON TipoIva.TipoIva = Clientes.TipoIva) ON Provincias.Provincia = Clientes.Provincia) ON Localidades.Localidad = Clientes.Localidad) ON Zonas.Zona = Clientes.Zona) ON ListadePrecio.ListaPrecio = Clientes.Precio "
      sSql = sSql & "ORDER BY Clientes.Cliente"
  End If
  Rs_Chk.Open sSql, db, adOpenStatic
  
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing

  Set ClientesAlfabetico = Rs_Chk
    
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function ClientesListados(bOrd As Boolean, sDes As String, sHas As String, Optional sZodes As String, Optional sZoHas As String) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
   
  db.Open sDb
 
  If bOrd = True Then
       sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Contacto, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion AS Provincia,"
       sSql = sSql & "Clientes.Telefono, TipoIva.Descripcion AS TipoIva, Clientes.Cuit, Transporte.Descripcion, Zonas.Descripcion, ListadePrecio.Descripcion "
       sSql = sSql & "FROM Zonas RIGHT OUTER JOIN "
       sSql = sSql & "Localidades RIGHT OUTER JOIN "
       sSql = sSql & "Provincias RIGHT OUTER JOIN "
       sSql = sSql & "TipoIva RIGHT OUTER JOIN "
       sSql = sSql & "Transporte RIGHT OUTER JOIN "
       sSql = sSql & "Clientes ON Transporte.Id = Clientes.Transporte ON TipoIva.TipoIva = Clientes.TipoIva ON Provincias.Provincia = Clientes.Provincia ON "
       sSql = sSql & "Localidades.Localidad = Clientes.Localidad ON Zonas.Zona = Clientes.Zona LEFT OUTER JOIN "
       sSql = sSql & "ListadePrecio ON Clientes.Precio = ListadePrecio.ListaPrecio "
      sSql = sSql & "WHERE Clientes.RazonSocial >='" & sDes & "' and Clientes.RazonSocial<='" & sHas & "' "
      If sZodes <> "" Then
         If sZodes <> "TODAS" Then
            sSql = sSql & " and Zonas.Descripcion>='" & sZodes & "' and Zonas.Descripcion<='" & sZoHas & "'"
         End If
      End If
      sSql = sSql & " ORDER BY Clientes.RazonSocial,Zonas.Descripcion"
  Else
       sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Contacto, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion AS Provincia,"
       sSql = sSql & "Clientes.Telefono, TipoIva.Descripcion AS TipoIva, Clientes.Cuit, Transporte.Descripcion, Zonas.Descripcion, ListadePrecio.Descripcion "
       sSql = sSql & "FROM Zonas RIGHT OUTER JOIN "
       sSql = sSql & "Localidades RIGHT OUTER JOIN "
       sSql = sSql & "Provincias RIGHT OUTER JOIN "
       sSql = sSql & "TipoIva RIGHT OUTER JOIN "
       sSql = sSql & "Transporte RIGHT OUTER JOIN "
       sSql = sSql & "Clientes ON Transporte.Id = Clientes.Transporte ON TipoIva.TipoIva = Clientes.TipoIva ON Provincias.Provincia = Clientes.Provincia ON "
       sSql = sSql & "Localidades.Localidad = Clientes.Localidad ON Zonas.Zona = Clientes.Zona LEFT OUTER JOIN "
       sSql = sSql & "ListadePrecio ON Clientes.Precio = ListadePrecio.ListaPrecio "
      
      sSql = sSql & "WHERE Clientes.RazonSocial >='" & sDes & "' and Clientes.RazonSocial<='" & sHas & "' "
      If sZodes <> "" Then
         If sZodes <> "TODAS" Then
            sSql = sSql & " and Zonas.Descripcion>='" & sZodes & "' and Zonas.Descripcion<='" & sZoHas & "'"
         End If
      End If
      sSql = sSql & " ORDER BY Clientes.Cliente"
  End If
  Rs_Chk.Open sSql, db, adOpenStatic
  
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing

  Set ClientesListados = Rs_Chk
  
  Set Rs_Chk = Nothing
  db.Close
  
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function TieneImpuesto(sClie As String, nImpu As Integer) As Boolean
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Chk As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Chk = New ADODB.Recordset
   db.Open sDb
 
   sSql = "SELECT  * FROM ClienteImpuesto WHERE Cliente='" & sClie & "' and Impuesto=" & nImpu
 
   Rs_Chk.Open sSql, db, adOpenStatic
   If Rs_Chk.BOF = False And Rs_Chk.EOF = False Then
      TieneImpuesto = True
   End If
   Rs_Chk.Close
   Set Rs_Chk = Nothing
   
   db.Close
   
   Exit Function
   
errHandler:
   err.Raise err.Number
End Function
Public Function TieneImpuestoProv(sProve As String, nImpu As Integer) As Boolean
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Chk As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Chk = New ADODB.Recordset
   
   db.Open sDb

   sSql = "SELECT  * FROM ProveedorImpuesto WHERE Proveedores='" & sProve & "' and Impuesto=" & nImpu
 
   Rs_Chk.Open sSql, db, adOpenStatic
   If Rs_Chk.BOF = False And Rs_Chk.EOF = False Then
      TieneImpuestoProv = True
   End If
   Rs_Chk.Close
   Set Rs_Chk = Nothing
   
   db.Close
   
   Exit Function
   
errHandler:
   err.Raise err.Number
End Function
Public Function TieneDescuento(sClie As String, nDesc As Integer) As Boolean
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Chk As ADODB.Recordset
 
   db.Open sDb

   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Chk = New ADODB.Recordset
   
   sSql = "SELECT  * FROM ClienteDescuento WHERE Cliente='" & sClie & "' and Descuento=" & nDesc
 
   Rs_Chk.Open sSql, db, adOpenStatic
   If Rs_Chk.BOF = False And Rs_Chk.EOF = False Then
      TieneDescuento = True
   End If
   Rs_Chk.Close
   Set Rs_Chk = Nothing
   
   db.Close
   
   Exit Function
   
errHandler:
   err.Raise err.Number
End Function

Public Function BuscarNombreCliente(sCod As String) As String
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT RazonSocial FROM Clientes WHERE Cliente='" & sCod & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarNombreCliente = Rs!RazonSocial
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function BuscarNombreProveedor(sCod As String) As String
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT RazonSocial FROM Proveedores WHERE Proveedor='" & sCod & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarNombreProveedor = Rs!RazonSocial
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function BuscarTipoIva(sCod As String) As String
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT TipoIva.Descripcion "
  sSql = sSql & "FROM Clientes LEFT JOIN TipoIva ON Clientes.TipoIva = TipoIva.TipoIva "
  sSql = sSql & "WHERE (((Clientes.Cliente)='" & sCod & "'))"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarTipoIva = Rs!Descripcion
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function BuscarTipoIvaReceptor(sCod As String) As Integer
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT TipoIva.DocControlador "
  sSql = sSql & "FROM Clientes LEFT JOIN TipoIva ON Clientes.TipoIva = TipoIva.TipoIva "
  sSql = sSql & "WHERE (((Clientes.Cliente)='" & sCod & "'))"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarTipoIvaReceptor = Rs!DocControlador
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function BuscarTipoIvaProv(sCod As String) As String
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT TipoIva.Descripcion "
  sSql = sSql & "FROM Proveedores LEFT JOIN TipoIva ON Proveedores.TipoIva = TipoIva.TipoIva "
  sSql = sSql & "WHERE (((Proveedores.Proveedor)='" & sCod & "'))"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarTipoIvaProv = Rs!Descripcion
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function BuscarLista(sCod As String) As Integer
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Precio "
  sSql = sSql & "FROM Clientes "
  sSql = sSql & "WHERE Cliente='" & sCod & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarLista = Rs!Precio
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function BuscarDepositos(sCod As String) As Integer
'  Dim sSql As String
'  Dim Rs As ADODB.Recordset
'
'  On Error GoTo errHandler
'  Set Rs = New ADODB.Recordset
'
'  sSql = "SELECT Precio "
'  sSql = sSql & "FROM Clientes "
'  sSql = sSql & "WHERE Cliente='" & sCod & "'"
'  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
'  If Rs.EOF = False Or Rs.BOF = False Then
'     BuscarLista = Rs!Precio
'  End If
'  Rs.Close
'
'Exit Function
'errHandler:
'  Err.Raise Err.Number
  
End Function

Public Function BuscarVendedor(sCod As String) As Integer
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Vendedor "
  sSql = sSql & "FROM Clientes "
  sSql = sSql & "WHERE Cliente='" & sCod & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarVendedor = Rs!Vendedor
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function BuscarCuit(sCliente As String) As String
  Dim sSql As String
  Dim Rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set Rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Cuit "
  sSql = sSql & "FROM Clientes "
  sSql = sSql & "WHERE Cliente='" & sCliente & "'"
  Rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If Rs.EOF = False Or Rs.BOF = False Then
     BuscarCuit = Rs!Cuit
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerDetalleFactura(sCliDes As String, sCliHas As String, fDesde As Date, fHasta As Date, nOrd As Byte) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

'  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono, Comprobantes.Descripcion, CabComprobantes.Fecha, CabComprobantes.Numero, CabComprobantes.Sucursal, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad, DetallesComprobantes.Precio,"
'  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 then  DetallesComprobantes.PrecioTotal  "
'  sSql = sSql & "ELSE -(DetallesComprobantes.PrecioTotal) END "
  
  
  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono, Comprobantes.Descripcion, CabComprobantes.Fecha, CabComprobantes.Numero, CabComprobantes.Sucursal, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad,"
  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 "
  sSql = sSql & "          then Case when Clientes.TipoIva=1 then DetallesComprobantes.PrecioUnitario  "
  sSql = sSql & "          Else DetallesComprobantes.Precio * DetallesComprobantes.Cantidad "
  sSql = sSql & "          End "
  sSql = sSql & "ELSE case when Clientes.TipoIva =1 then -(DetallesComprobantes.PrecioUnitario) "
  sSql = sSql & "          else   -(DetallesComprobantes.Precio * DetallesComprobantes.Cantidad) "
  sSql = sSql & "          end "
  sSql = sSql & "END, "
  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 "
  sSql = sSql & "     then Case when Clientes.TipoIva=1 then DetallesComprobantes.PrecioTotal "
  sSql = sSql & "          else (DetallesComprobantes.Impuesto + DetallesComprobantes.PrecioTotal) "
  sSql = sSql & "          end "
  sSql = sSql & "ELSE case when Clientes.TipoIva =1 then -(DetallesComprobantes.PrecioTotal) "
  sSql = sSql & "     ELSE   -(DetallesComprobantes.Impuesto + DetallesComprobantes.PrecioTotal) "
  sSql = sSql & "     End "
  sSql = sSql & "END "
  sSql = sSql & "FROM ((Clientes LEFT JOIN CabComprobantes ON Clientes.Cliente = CabComprobantes.Cliente) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "Where CabComprobantes.Numero Is Not Null and CabComprobantes.Anulado<>1 and Clientes.Tipo=1 And CabComprobantes.Suc=" & nSucursal & " and (CabComprobantes.Tipo=1  or cabcomprobantes.tipo=5) "
  If nOrd = 1 Then
     sSql = sSql & " and Clientes.RazonSocial>='" & sCliDes & "' and Clientes.RazonSocial<='" & sCliHas & "'"
  Else
     sSql = sSql & " and Clientes.Cliente>='" & sCliDes & "' and Clientes.Cliente<='" & sCliHas & "'"
  End If
  sSql = sSql & " and CabComprobantes.Fecha>='" & FechaAmericana(fDesde) & "' and CabComprobantes.Fecha<='" & FechaAmericana(fHasta) & "' AND DetallesComprobantes.Descripcion<>'COBRANZA'"
  If nOrd = 1 Then
     sSql = sSql & " ORDER BY Clientes.RazonSocial, CabComprobantes.Numero "
  Else
     sSql = sSql & " ORDER BY Clientes.Cliente,CabComprobantes.Fecha,CabComprobantes.Numero "
  End If
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set TraerDetalleFactura = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

End Function

Public Function TraerDetalleFacturaP(sCliDes As String, sCliHas As String, fDesde As Date, fHasta As Date, nOrd As Byte) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.Telefono, Comprobantes.Descripcion, CabComprobantes.Fecha, CabComprobantes.Numero, CabComprobantes.Sucursal, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad, DetallesComprobantes.Precio,"
  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=8 then  DetallesComprobantes.PrecioTotal + DetallesComprobantes.Impuesto "
  sSql = sSql & "ELSE -(DetallesComprobantes.PrecioTotal+ DetallesComprobantes.Impuesto) END "
  sSql = sSql & "FROM ((Proveedores LEFT JOIN CabComprobantes ON Proveedores.Proveedor = CabComprobantes.Cliente) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "Where CabComprobantes.Numero Is Not Null and CabComprobantes.Anulado<>1 And CabComprobantes.Suc=" & nSucursal & " and (CabComprobantes.Tipo=8  or cabcomprobantes.tipo=12) "
  If nOrd = 1 Then
     sSql = sSql & " and Proveedores.RazonSocial>='" & sCliDes & "' and Proveedores.RazonSocial<='" & sCliHas & "'"
  Else
     sSql = sSql & " and Proveedores.Proveedor>='" & sCliDes & "' and Proveedores.Proveedor<='" & sCliHas & "'"
  End If
  sSql = sSql & " and CabComprobantes.Fecha>='" & FechaAmericana(fDesde) & "' and CabComprobantes.Fecha<='" & FechaAmericana(fHasta) & "' AND DetallesComprobantes.Descripcion<>'COBRANZA'"
  If nOrd = 1 Then
     sSql = sSql & " ORDER BY CabComprobantes.Fecha"
  Else
     sSql = sSql & " ORDER BY Proveedores.Proveedor,CabComprobantes.Fecha,CabComprobantes.Numero "
  End If
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set TraerDetalleFacturaP = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

End Function


Public Function ResumenSaldoAnteriorBus(sCliDes As String) As Currency
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT Clientes.Cliente,Sum(CuentaCorriente.Debe*CabComprobantes.Cotizacion) AS SumaDeDebe, Sum(CuentaCorriente.Haber*CabComprobantes.Cotizacion) AS SumaDeHaber "
  sSql = sSql & "FROM Localidades RIGHT JOIN (CabComprobantes RIGHT JOIN (TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Clientes LEFT JOIN (CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) ON Clientes.Cliente = CuentaCorriente.Cliente) ON Provincias.Provincia = Clientes.Provincia) ON TipoIva.TipoIva = Clientes.TipoIva) ON CabComprobantes.Movimiento = CuentaCorriente.Venta) ON Localidades.Localidad = Clientes.Localidad "
  sSql = sSql & "WHERE Clientes.Cliente='" & sCliDes & "'  "
  sSql = sSql & "GROUP BY Clientes.Cliente "

  Rs_Chk.Open sSql, db, adOpenStatic
  If Not IsNull(Rs_Chk!SumaDeDebe) And Not IsNull(Rs_Chk!SumaDeHaber) And Rs_Chk.RecordCount <> 0 Then
      ResumenSaldoAnteriorBus = Round(Rs_Chk!SumaDeDebe, 2) - Round(Rs_Chk!SumaDeHaber, 2)
  End If
  Set Rs_Chk = Nothing
  
  db.Close
  
Exit Function
   
errHandler:
   err.Raise err.Number

End Function

Public Function ResumenSaldoAnteriorBusP(sCliDes As String) As Currency
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT Proveedores.Proveedor,Sum(CuentaCorriente.Debe*CabComprobantes.Cotizacion) AS SumaDeDebe, Sum(CuentaCorriente.Haber*CabComprobantes.Cotizacion) AS SumaDeHaber "
  sSql = sSql & "FROM ((Proveedores LEFT JOIN CuentaCorriente ON Proveedores.Proveedor = CuentaCorriente.Cliente) LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) LEFT JOIN CabComprobantes ON CuentaCorriente.Venta = CabComprobantes.Movimiento "
  sSql = sSql & "WHERE Proveedores.Proveedor='" & sCliDes & "' "
  sSql = sSql & "GROUP BY Proveedores.Proveedor "

  Rs_Chk.Open sSql, db, adOpenStatic
  If Not IsNull(Rs_Chk!SumaDeDebe) And Not IsNull(Rs_Chk!SumaDeHaber) And Rs_Chk.RecordCount <> 0 Then
      ResumenSaldoAnteriorBusP = Round(Rs_Chk!SumaDeHaber, 2) - Round(Rs_Chk!SumaDeDebe, 2)
  End If
  Set Rs_Chk = Nothing
  
  db.Close
  
Exit Function
   
errHandler:
   err.Raise err.Number

End Function

Public Function ResumenSaldoAnterior(sCliDes As String, sCliHas As String, pFecha As Date, Optional sZodes As Integer = 0) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
  
   db.Open sDb

  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset


  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Zona,Clientes.Ciudad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Clientes.Cuit, Clientes.Telefono, Sum(CuentaCorriente.Debe*CabComprobantes.Cotizacion) AS SumaDeDebe, Sum(CuentaCorriente.Haber*CabComprobantes.Cotizacion) AS SumaDeHaber "
  sSql = sSql & "FROM Localidades RIGHT JOIN (CabComprobantes RIGHT JOIN (TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Clientes LEFT JOIN (CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) ON Clientes.Cliente = CuentaCorriente.Cliente) ON Provincias.Provincia = Clientes.Provincia) ON TipoIva.TipoIva = Clientes.TipoIva) ON CabComprobantes.Movimiento = CuentaCorriente.Venta) ON Localidades.Localidad = Clientes.Localidad "
  sSql = sSql & "WHERE Clientes.RazonSocial >='" & sCliDes & "'  And Clientes.RazonSocial <= '" & sCliHas & "' And CuentaCorriente.Fecha <'" & FechaAmericana(pFecha) & "' "
  If sZodes <> 0 Then
     sSql = sSql & " and Clientes.Zona=" & sZodes & " "
  End If
  sSql = sSql & "AND Clientes.Cliente<> '.' "

  sSql = sSql & "GROUP BY Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.CodigoPostal,Clientes.Zona, Clientes.Ciudad, Provincias.Descripcion, TipoIva.Descripcion, Clientes.Cuit, Clientes.Telefono "
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs_Chk.ActiveConnection = Nothing
  
    
  Set ResumenSaldoAnterior = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

End Function
Public Function TraerInterPropiosAnte(pBanco As Integer, pFecha As Date) As ADODB.Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb

'  sSql = "SELECT DISTINCT ctabanco.Debe, ctabanco.Haber, CuentasBancarias.Banco,ctabanco.Banco "
'  sSql = sSql & "FROM ctabanco LEFT JOIN CuentasBancarias ON ctabanco.Banco = CuentasBancarias.Id "
'  sSql = sSql & "WHERE ((ctabanco.FechaAcreditacion)<#" & FechaAmericana(pFecha) & "') "
'  sSql = sSql & "GROUP BY ctabanco.Banco, ctabanco.Anulado, CuentasBancarias.Banco,ctabanco.Debe, ctabanco.Haber "
'  sSql = sSql & "HAVING ctabanco.Anulado=0 AND CuentasBancarias.Banco=" & pBanco

   sSql = "SELECT DISTINCT ctabanco.Debe, ctabanco.Haber, CuentasBancarias.Banco, ctabanco.Banco, ctabanco.NumeroCheque, ctabanco.FechaAcreditacion "
   sSql = sSql & "FROM ctabanco LEFT JOIN CuentasBancarias ON ctabanco.Banco = CuentasBancarias.Id "
   sSql = sSql & "WHERE ((ctabanco.FechaAcreditacion)<'" & FechaAmericana(pFecha) & "') "
   sSql = sSql & "GROUP BY ctabanco.Debe, ctabanco.Haber, CuentasBancarias.Banco, ctabanco.Banco, ctabanco.Anulado, ctabanco.NumeroCheque, ctabanco.FechaAcreditacion "
   sSql = sSql & "HAVING ctabanco.Anulado=0 AND CuentasBancarias.Banco=" & pBanco

  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing

  Set TraerInterPropiosAnte = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

End Function

Public Function TraerCodigoCuenta(pProv As String) As Integer
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Proveedores.Proveedor, CuentasBancarias.Banco "
  sSql = sSql & "FROM Proveedores LEFT JOIN CuentasBancarias ON Proveedores.Id = CuentasBancarias.Banco "
  sSql = sSql & "WHERE Proveedores.Proveedor='" & pProv & "'"
  
  
  Rs_Chk.Open sSql, db, adOpenStatic
  
  TraerCodigoCuenta = Rs_Chk!Banco
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

  
End Function
Public Function TraerInterPropios(pBanco As Integer, pDesde As Date, pHasta As Date) As ADODB.Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT DISTINCT ctabanco.Banco, ctabanco.FormaPago,ctabanco.Nombre, ctabanco.NumeroCheque, ctabanco.FechaAcreditacion, ctabanco.Debe, ctabanco.Haber "
  sSql = sSql & "From ctabanco "
 ' sSql = sSql & "WHERE (ctabanco.FormaPago='INTERDEPOSITOS' Or ctabanco.FormaPago='CHEQUES PROPIOS' Or ctabanco.FormaPago='DEBITO AUTOMATICO')"
  sSql = sSql & "WHERE (ctabanco.FormaPago<>'EFECTIVO') "
  sSql = sSql & "AND ctabanco.Banco=" & pBanco & " AND Anulado=0 "
  sSql = sSql & "AND Fechaacreditacion>='" & FechaAmericana(pDesde) & "' AND Fechaacreditacion<='" & FechaAmericana(pHasta) & "' "
 ' sSql = sSql & "ORDER BY ctabanco.Fecha"
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
    
  Set TraerInterPropios = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function TraerMovideBancos(pBanco As Integer, pDesde As Date, pHasta As Date)
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CtaBanco.Fecha, Comprobantes.Descripcion, CtaBanco.Sucursal, CtaBanco.Numero, Sum(CtaBanco.Debe) AS SumaDeDebe, Sum(CtaBanco.Haber) AS SumaDeHaber "
  sSql = sSql & "FROM CabComprobantes RIGHT JOIN (CtaBanco LEFT JOIN Comprobantes ON CtaBanco.Comprobante = Comprobantes.Id) ON CabComprobantes.Movimiento = CtaBanco.Movimiento "
  sSql = sSql & "WHERE CtaBanco.Anulado=0 "
  sSql = sSql & "GROUP BY CtaBanco.Fecha, Comprobantes.Descripcion, CtaBanco.Sucursal, CtaBanco.Numero, Comprobantes.TipoOperacion, CabComprobantes.Vendedor "
  sSql = sSql & "having Comprobantes.TipoOperacion=14 AND CabComprobantes.Vendedor=" & pBanco & "  "
  sSql = sSql & "AND CtaBanco.Fecha>='" & FechaAmericana(pDesde) & "' AND CtaBanco.Fecha<='" & FechaAmericana(pHasta) & "' "
  sSql = sSql & "ORDER BY CtaBanco.Fecha"
  
  
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set TraerMovideBancos = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function TraerMovideBancosAnt(pBanco As Integer, pDesde As Date)
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Sum(ctabanco.Debe) AS SumaDeDebe, Sum(ctabanco.Haber) AS SumaDeHaber "
  sSql = sSql & "FROM CabComprobantes RIGHT JOIN (ctabanco LEFT JOIN Comprobantes ON ctabanco.Comprobante = Comprobantes.Id) ON CabComprobantes.Movimiento = ctabanco.Movimiento "
  sSql = sSql & "Where CabComprobantes.Fecha < '" & FechaAmericana(pDesde) & "' AND ctabanco.Anulado=0 "
  sSql = sSql & "GROUP BY Comprobantes.TipoOperacion, CabComprobantes.Vendedor "
  sSql = sSql & "HAVING (((Comprobantes.TipoOperacion)=14) AND ((CabComprobantes.Vendedor)=" & pBanco & "));"
  
   
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set TraerMovideBancosAnt = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function ResumenSaldoAnteriorP(sCliDes As String, sCliHas As String, pFecha As Date) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.CodigoPostal, Proveedores.Localidad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Proveedores.Cuit, Proveedores.Telefono, Sum(CuentaCorriente.Debe) AS SumaDeDebe, Sum(CuentaCorriente.Haber) AS SumaDeHaber "
  sSql = sSql & "FROM TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Proveedores LEFT JOIN (CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) ON Proveedores.Proveedor = CuentaCorriente.Cliente) ON Provincias.Provincia = Proveedores.Provincia) ON TipoIva.TipoIva = Proveedores.TipoIva "
  sSql = sSql & "Where Proveedores.RazonSocial >='" & sCliDes & "'  And Proveedores.RazonSocial <= '" & sCliHas & "' And CuentaCorriente.Fecha <'" & FechaAmericana(pFecha) & "' "
  sSql = sSql & "GROUP BY Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.CodigoPostal, Proveedores.Localidad, Provincias.Descripcion, TipoIva.Descripcion, Proveedores.Cuit, Proveedores.Telefono "
  
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set ResumenSaldoAnteriorP = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

End Function

Public Function ResumenDeCuenta(sCliDes As String, sCliHas As String, sFeDes As Date, sFeHas As Date, Optional pFecha As Boolean, Optional sZodes As Integer = 0) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Zona,Clientes.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Clientes.Cuit, Clientes.Telefono, CuentaCorriente.Fecha, Comprobantes.Descripcion as Comprobante, CuentaCorriente.Numero, CuentaCorriente.Sucursal, CuentaCorriente.Cuota, CuentaCorriente.Debe*CabComprobantes.Cotizacion AS Debe, CuentaCorriente.Haber *CabComprobantes.Cotizacion as Haber,CabComprobantes.Motivo,CabComprobantes.Tipo "
  sSql = sSql & "FROM Localidades RIGHT JOIN ((TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Clientes LEFT JOIN (CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) ON Clientes.Cliente = CuentaCorriente.Cliente) ON Provincias.Provincia = Clientes.Provincia) ON TipoIva.TipoIva = Clientes.TipoIva) LEFT JOIN CabComprobantes ON CuentaCorriente.Venta = CabComprobantes.Movimiento) ON Localidades.Localidad = Clientes.Localidad "
  sSql = sSql & "Where Clientes.RazonSocial >='" & sCliDes & "' And Clientes.RazonSocial <='" & sCliHas & "' "
  sSql = sSql & "and CuentaCorriente.Fecha>='" & FechaAmericana(sFeDes) & "' "
  If pFecha = True Then
     sSql = sSql & "and CuentaCorriente.Fecha<='" & FechaAmericana(sFeHas) & "' "
  End If
  If sZodes <> 0 Then
     sSql = sSql & " and Clientes.Zona=" & sZodes & " "
  End If

  sSql = sSql & "ORDER BY CuentaCorriente.Fecha,CuentaCorriente.Id, Clientes.Cliente "

  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing

  
  
  Set ResumenDeCuenta = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
End Function

Public Function TraerSaldoCtaCte(pCliente As String) As Double
 Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb
  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono, (Sum(CuentaCorriente.Debe)-Sum(CuentaCorriente.Haber)) AS Saldo "
  sSql = sSql & "FROM Clientes LEFT JOIN CuentaCorriente ON Clientes.Cliente = CuentaCorriente.Cliente "
  sSql = sSql & "Where Clientes.Cliente='" & pCliente & "' "
  sSql = sSql & "GROUP BY Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono "
  sSql = sSql & "Having (((Sum(CuentaCorriente.Debe)) Is Not Null)) "
   
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  If Rs_Chk.RecordCount <> 0 Then
     TraerSaldoCtaCte = Rs_Chk!saldo
  End If
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
   
End Function

Public Function ResumenDeCuentaSimpli(sCliDes As String, sCliHas As String, sFeDes As Date, sFeHas As Date, Optional pFecha As Boolean, Optional sZodes As Integer) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono, (Sum(CuentaCorriente.Debe)-Sum(CuentaCorriente.Haber)) AS Saldo "
  sSql = sSql & "FROM Clientes LEFT JOIN CuentaCorriente ON Clientes.Cliente = CuentaCorriente.Cliente "
  sSql = sSql & "Where Clientes.RazonSocial >='" & sCliDes & "' And Clientes.RazonSocial <='" & sCliHas & "' "
  If pFecha = True Then
     sSql = sSql & "and CuentaCorriente.Fecha<='" & FechaAmericana(sFeHas) & "' "
  End If
  
  If sZodes <> 0 Then
    sSql = sSql & " and Clientes.Zona=" & sZodes & " "
  End If

  sSql = sSql & "GROUP BY Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono "
  sSql = sSql & "Having (((Sum(CuentaCorriente.Debe)) Is Not Null)) "
  sSql = sSql & "ORDER BY Clientes.Cliente "

  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  
  Set ResumenDeCuentaSimpli = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

End Function

Public Function ResumenDeCuentaSimpliCompra(sCliDes As String, sCliHas As String, sFeDes As Date, sFeHas As Date, Optional pFecha As Boolean, Optional sZodes As Integer) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.Telefono, (Sum(CuentaCorriente.Debe)-Sum(CuentaCorriente.Haber)) AS Saldo "
  sSql = sSql & "FROM Proveedores LEFT JOIN CuentaCorriente ON Proveedores.Proveedor = CuentaCorriente.Cliente "
  sSql = sSql & "Where Proveedores.RazonSocial >='" & sCliDes & "' And Proveedores.RazonSocial <='" & sCliHas & "' "
  
  sSql = sSql & "GROUP BY Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.Telefono "
  sSql = sSql & "Having (((Sum(CuentaCorriente.Debe)) Is Not Null)) "
  sSql = sSql & "ORDER BY Proveedores.Proveedor "

  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  
  Set ResumenDeCuentaSimpliCompra = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number

End Function

Public Function BuscarDatosCli(pClien As String) As ADODB.Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva "
  sSql = sSql & "FROM (Clientes LEFT JOIN Provincias ON Clientes.Provincia = Provincias.Provincia) LEFT JOIN TipoIva ON Clientes.TipoIva = TipoIva.TipoIva "
  sSql = sSql & "WHERE Cliente='" & pClien & "'"
 
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set BuscarDatosCli = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function


Public Function ProveedorAlfabetico() As ADODB.Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
   
  db.Open sDb
 
  sSql = "SELECT Proveedores.Id, Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.CodigoPostal, Proveedores.Localidad, Provincias.Descripcion as Provincia, TipoIva.Descripcion, Proveedores.Cuit "
  sSql = sSql & "FROM TipoIva RIGHT JOIN (Provincias RIGHT JOIN Proveedores ON Provincias.Provincia = Proveedores.Provincia) ON TipoIva.TipoIva = Proveedores.TipoIva "
  sSql = sSql & "ORDER BY Proveedores.RazonSocial; "
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  Set ProveedorAlfabetico = Rs_Chk
    
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function ResumenDeCuentaP(sCliDes As String, sCliHas As String, sFeDes As Date, sFeHas As Date) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.CodigoPostal, Proveedores.Localidad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Proveedores.Cuit, Proveedores.Telefono, CuentaCorriente.Fecha, Comprobantes.Descripcion as Comprobante, CuentaCorriente.Numero, CuentaCorriente.Sucursal, CuentaCorriente.Cuota, CuentaCorriente.Debe*CabComprobantes.Cotizacion AS Debe, CuentaCorriente.Haber*CabComprobantes.Cotizacion AS Haber "
  sSql = sSql & "FROM CabComprobantes RIGHT JOIN (TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Proveedores LEFT JOIN (CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) ON Proveedores.Proveedor = CuentaCorriente.Cliente) ON Provincias.Provincia = Proveedores.Provincia) ON TipoIva.TipoIva = Proveedores.TipoIva) ON CabComprobantes.Movimiento = CuentaCorriente.Venta "
  sSql = sSql & "Where Proveedores.RazonSocial >='" & sCliDes & "' And Proveedores.RazonSocial <='" & sCliHas & "'  "
 ' sSql = sSql & "AND CuentaCorriente.Debe - CuentaCorriente.Haber<>0 "
  sSql = sSql & "and CuentaCorriente.Fecha>='" & FechaAmericana(sFeDes) & "' and CuentaCorriente.Fecha<='" & FechaAmericana(sFeHas) & "' and Proveedores.CtaCte=0 "
  sSql = sSql & "ORDER BY CuentaCorriente.Fecha,CuentaCorriente.Id, Proveedores.Proveedor "


  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set ResumenDeCuentaP = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
End Function

Public Function ResumenDeCuentaB(sCliDes As String, sCliHas As String, sFeDes As Date, sFeHas As Date) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.Localidad as ciudad, Proveedores.CodigoPostal, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Proveedores.Cuit, Proveedores.Telefono, CuentaCorriente.Fecha, Comprobantes.Descripcion, CuentaCorriente.Numero, CuentaCorriente.Sucursal, CuentaCorriente.Cuota, CuentaCorriente.Debe, CuentaCorriente.Haber "
  sSql = sSql & "FROM TipoIva RIGHT JOIN (Provincias RIGHT JOIN (Proveedores LEFT JOIN (CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) ON Proveedores.Proveedor = CuentaCorriente.Cliente)  ON Provincias.Provincia = Proveedores.Provincia) ON TipoIva.TipoIva = Proveedores.TipoIva "
  sSql = sSql & "Where Proveedores.RazonSocial >='" & sCliDes & "' And Proveedores.RazonSocial <='" & sCliHas & "'  "
  sSql = sSql & "and CuentaCorriente.Fecha>='" & FechaAmericana(sFeDes) & "' and CuentaCorriente.Fecha<='" & FechaAmericana(sFeHas) & "' "
  sSql = sSql & "ORDER BY CuentaCorriente.Fecha,CuentaCorriente.Id, Proveedores.Proveedor "

  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set ResumenDeCuentaB = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
End Function


Public Function ProveedorListados(bOrd As Boolean, sDes As String, sHas As String, Optional sZodes As String) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  db.Open sDb
  
  If bOrd = True Then
      sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Contacto, Proveedores.Domicilio, Proveedores.CodigoPostal, Proveedores.Localidad, Provincias.Descripcion as Provincia, Proveedores.Telefono, TipoIva.Descripcion as TipoIva, Proveedores.Cuit, Proveedores.IngBrutos "
      sSql = sSql & "FROM Provincias RIGHT JOIN (TipoIva RIGHT JOIN Proveedores ON TipoIva.TipoIva = Proveedores.TipoIva) ON Provincias.Provincia = Proveedores.Provincia "
      sSql = sSql & "WHERE Proveedores.RazonSocial >='" & sDes & "' and Proveedores.RazonSocial<='" & sHas & "' "
      sSql = sSql & " ORDER BY Proveedores.RazonSocial"
  Else
      sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Contacto, Proveedores.Domicilio, Proveedores.CodigoPostal, Proveedores.Localidad, Provincias.Descripcion as Provincia, Proveedores.Telefono, TipoIva.Descripcion as TipoIva, Proveedores.Cuit, Proveedores.IngBrutos "
      sSql = sSql & "FROM Provincias RIGHT JOIN (TipoIva RIGHT JOIN Proveedores ON TipoIva.TipoIva = Proveedores.TipoIva) ON Provincias.Provincia = Proveedores.Provincia "
      sSql = sSql & "WHERE Proveedores.RazonSocial >='" & sDes & "' and Proveedores.RazonSocial<='" & sHas & "' "
      sSql = sSql & " ORDER BY Proveedores.Proveedor"
  End If
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  Set ProveedorListados = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

'Public Function TraerDetalleFacturaP(sCliDes As String, sCliHas As String, fDesde As Date, fHasta As Date, nOrd As Byte) As Recordset
'  Dim sSql As String
'  Dim Conexion As ADODB.Connection
'  Dim Rs_Chk As ADODB.Recordset
'
'  On Error GoTo errHandler
'  Set Conexion = New ADODB.Connection
'  Set Rs_Chk = New ADODB.Recordset
'
'   db.Open sDb
'
'  sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.Telefono, Comprobantes.Descripcion, CabComprobantes.Fecha, CabComprobantes.Numero, CabComprobantes.Sucursal, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad, DetallesComprobantes.PrecioUnitario, DetallesComprobantes.PrecioTotal "
'  sSql = sSql & "FROM ((Proveedores LEFT JOIN CabComprobantes ON Proveedores.Proveedor = CabComprobantes.Cliente) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
'  sSql = sSql & "Where CabComprobantes.Numero Is Not Null and CabComprobantes.Anulado<>1 and CabComprobantes.Tipo=8"
'  sSql = sSql & " and Proveedores.RazonSocial>='" & sCliDes & "' and Proveedores.RazonSocial<='" & sCliHas & "'"
'  sSql = sSql & " and CabComprobantes.Fecha>='" & FechaAmericana(fDesde) & "' and CabComprobantes.Fecha<='" & FechaAmericana(fHasta) & "'"
'  If nOrd = 1 Then
'     sSql = sSql & " ORDER BY Proveedores.RazonSocial, CabComprobantes.Numero "
'  Else
'     sSql = sSql & " ORDER BY Proveedores.Proveedor,CabComprobantes.Numero "
'  End If
'  Rs_Chk.Open sSql, db, adOpenStatic
'  Rs_Chk.Close
'  Rs_Chk.CursorType = adOpenStatic
'  Rs_Chk.CursorLocation = adUseClient
'  Rs_Chk.LockType = adLockBatchOptimistic
'  Rs_Chk.Open
'  Rs_Chk.ActiveConnection = Nothing
'
'  Set TraerDetalleFacturaP = Rs_Chk
'  db.Close
'
'  Set Rs_Chk = Nothing
'Exit Function
'
'errHandler:
'   err.Raise err.Number
'
'End Function

Public Function ResumenCuentaDetalles(sCliDes As String, sCliHas As String, sFeDes As Date, sFeHas As Date, Optional sZodes As Integer = 0) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb



  sSql = "SELECT CabComprobantes.Fecha,CabComprobantes.Movimiento,CabComprobantes.Lista,Comprobantes.Comprobante ,Comprobantes.Descripcion, CabComprobantes.Numero, CabComprobantes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad,DetallesComprobantes.Descripcion as Detalle,sum(DetallesComprobantes.Cantidad)as Cantidad, DetallesComprobantes.PrecioUnitario*CabComprobantes.Cotizacion AS PRecioUnitario, Sum(DetallesComprobantes.Impuesto*CabComprobantes.Cotizacion) AS Impuesto,Sum(DetallesComprobantes.Descuento*CabComprobantes.Cotizacion) as Descuento ,Sum(DetallesComprobantes.PrecioTotal*CabComprobantes.Cotizacion) AS PrecioTotal,CabComprobantes.Motivo "
  sSql = sSql & "FROM ((CabComprobantes LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE CabComprobantes.Anulado<>1 and CabComprobantes.CondVenta=0 and CabComprobantes.Tipo<>6 and  Comprobantes.Comprobante<>'FAPP'  AND (Comprobantes.CtaCte<>'No' Or  CabComprobantes.Tipo=4) "
  If sZodes <> 0 Then
     sSql = sSql & " and Clientes.Zona=" & sZodes & " "
  End If
  sSql = sSql & "AND Clientes.Cliente<> '.' "
  sSql = sSql & "GROUP BY CabComprobantes.Fecha,CabComprobantes.Movimiento,CabComprobantes.Lista,Comprobantes.Comprobante, Comprobantes.Descripcion, CabComprobantes.Numero, CabComprobantes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.CodigoPostal, Clientes.Ciudad,DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad,CabComprobantes.Motivo,DetallesComprobantes.PrecioUnitario,CabComprobantes.Cotizacion "
  sSql = sSql & "HAVING CabComprobantes.Fecha>='" & FechaAmericana(sFeDes) & "' And CabComprobantes.Fecha<='" & FechaAmericana(sFeHas) & "' AND Clientes.RazonSocial>='" & sCliDes & "' And Clientes.RazonSocial<='" & sCliHas & "' "
  sSql = sSql & "ORDER BY CabComprobantes.Fecha,Comprobantes.Comprobante,CabComprobantes.Numero"
  
  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing

  Set ResumenCuentaDetalles = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function ResumenCuentaDetallesP(sCliDes As String, sCliHas As String, sFeDes As Date, sFeHas As Date) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT DISTINCT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.CodigoPostal, Clientes.Ciudad, Provincias.Descripcion as Provincia, TipoIva.Descripcion as TipoIva, Proveedores.Cuit, Proveedores.Telefono, CuentaCorriente.Fecha, Comprobantes.Descripcion, CuentaCorriente.Numero, CuentaCorriente.Sucursal, CuentaCorriente.Cuota, CuentaCorriente.Debe, CuentaCorriente.Haber, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad, DetallesComprobantes.PrecioUnitario,DetallesComprobantes.PrecioTotal,Comprobantes.Comprobante "
  sSql = sSql & "FROM Provincias RIGHT JOIN (Localidades RIGHT JOIN ((TipoIva RIGHT JOIN ((Proveedores LEFT JOIN (CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id) ON Proveedores.Proveedor = CuentaCorriente.Cliente) LEFT JOIN CabComprobantes ON (CuentaCorriente.Comprobante = CabComprobantes.Comprobante) AND (CuentaCorriente.Venta = CabComprobantes.Movimiento)) ON TipoIva.TipoIva = Proveedores.TipoIva) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) ON Localidades.Localidad = Proveedores.Localidad) ON Provincias.Provincia = Proveedores.Provincia "
  sSql = sSql & "Where Proveedores.RazonSocial >='" & sCliDes & "' And Proveedores.RazonSocial <='" & sCliHas & "' "
  sSql = sSql & "and CuentaCorriente.Fecha>='" & FechaAmericana(sFeDes) & "' and CuentaCorriente.Fecha<='" & FechaAmericana(sFeHas) & "' "
  sSql = sSql & "ORDER BY CuentaCorriente.Fecha, Proveedores.Proveedor "

  Rs_Chk.Open sSql, db, adOpenStatic
  Rs_Chk.Close
  Rs_Chk.CursorType = adOpenStatic
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.LockType = adLockBatchOptimistic
  Rs_Chk.Open
  Rs_Chk.ActiveConnection = Nothing
  
  Set ResumenCuentaDetallesP = Rs_Chk
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function ClienteDiscrimina(sClien As String) As Boolean
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
   
  db.Open sDb
 
  sSql = sSql & "SELECT TipoIva.Discrimina "
  sSql = sSql & "FROM Clientes LEFT JOIN TipoIva ON Clientes.TipoIva = TipoIva.TipoIva "
  sSql = sSql & "WHERE Clientes.Cliente='" & sClien & "'"
  
  Rs_Chk.Open sSql, db, adOpenStatic
  If Rs_Chk!Discrimina = True Then
     ClienteDiscrimina = True
  End If
  
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
  
End Function

Public Function TraerUltimoPago(pCliente As String) As Date

Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  On Error GoTo errHandler
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset
   
  db.Open sDb
  
  sSql = "SELECT TOP (1) Fecha "
  sSql = sSql & "From CuentaCorriente "
  sSql = sSql & "WHERE (Cliente = '" & pCliente & "') AND (Comprobante = 8) "
  sSql = sSql & "ORDER BY Fecha DESC "
  
  Rs_Chk.Open sSql, db, adOpenStatic
  
  If Rs_Chk.RecordCount <> 0 Then
    TraerUltimoPago = Rs_Chk!fecha
  Else
    TraerUltimoPago = Date
  End If
  
  db.Close
  
  Set Rs_Chk = Nothing
Exit Function
   
errHandler:
   err.Raise err.Number
   
End Function
