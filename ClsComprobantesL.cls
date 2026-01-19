VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsComprobantesL"
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

Public Function TraerGananciaResumen(pDesde As Date, pHasta As Date) As Recordset
 Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  db.Open sDb

  Set rs = New ADODB.Recordset
  sSql = "SELECT CabComprobantes.Fecha, SUM(CabGanancia.Base) AS Neto, CabGanancia.Importe,"
  sSql = sSql & "CabComprobantes.Numero AS NCompro, Proveedores.Cuit,CabGanancia.Numero "
  sSql = sSql & "FROM Proveedores RIGHT OUTER JOIN "
  sSql = sSql & "CabComprobantes ON Proveedores.Proveedor = CabComprobantes.Cliente RIGHT OUTER JOIN "
  sSql = sSql & "CabGanancia ON CabComprobantes.Movimiento = CabGanancia.MoviPag "
  sSql = sSql & "WHERE CabComprobantes.Fecha>='" & pDesde & "' AND CabComprobantes.Fecha<='" & pHasta & "' "
  sSql = sSql & "GROUP BY CabGanancia.MoviPag, CabGanancia.Numero, CabGanancia.Importe, CabComprobantes.Movimiento, CabComprobantes.Fecha, CabComprobantes.Numero, Proveedores.Cuit "
  sSql = sSql & "Having (CabGanancia.Numero <> 0) "
  sSql = sSql & "ORDER BY CabGanancia.Numero"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerGananciaResumen = rs
  
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function TraerCategoriaGanancia(pCategoria As Byte, pNeto As Single) As Recordset
 Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  db.Open sDb

  Set rs = New ADODB.Recordset
  sSql = "SELECT Id_Ganancia, Importe, Desde, Hasta, Monto, Porcentaje, SobreExedente, Minimo, Codigo "
  sSql = sSql & "From TblGananciaItems "
  If pCategoria = 3 Then
     sSql = sSql & "Where Id_ganancia = " & pCategoria & " And (Desde <= " & pNeto & " - importe And hasta >= " & pNeto & " - importe) "
  Else
      sSql = sSql & "Where Id_ganancia = " & pCategoria
  End If
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerCategoriaGanancia = rs
  
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerCabGaganancia(pNumero As Long) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
   On Error GoTo errHandler
   db.Open sDb
   
   Set rs = New ADODB.Recordset
    
   sSql = "SELECT DISTINCT MoviPag, Movimiento, Provincia, FechaRec, Fecha, Numero, OPago, Proveedor, Base, Alicuota, Importe "
   sSql = sSql & "From CabGanancia "
   sSql = sSql & "WHERE Numero=" & pNumero & " "
   sSql = sSql & "ORDER BY OPago "
   
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerCabGaganancia = rs
  
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
     
End Function
Public Function TraerRsComprobante(pId As Long) As Recordset
  On Error GoTo errHandler
  EnsureConn
  
  Dim sSql As String
  Dim rs As ADODB.Recordset
    
  Set rs = New ADODB.Recordset

  sSql = "SELECT  CabComprobantes.*,Clientes.RazonSocial as RazonSocial, Clientes.TipoIva,Clientes.Precio as Precio "
  sSql = sSql & "FROM CabComprobantes LEFT OUTER JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente "
  sSql = sSql & "WHERE (CabComprobantes.Id = " & pId & ") "

  rs.CursorLocation = adUseClient
  rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerRsComprobante = rs
    
  Exit Function
errHandler:
  err.Raise err.Number
End Function


Public Function TraerComprobantesRemitoVentas() As Recordset
  On Error GoTo errHandler
  EnsureConn
  
  Dim sSql As String
  Dim rs As ADODB.Recordset
    
  Set rs = New ADODB.Recordset

  sSql = "SELECT TOP 50 Id,Movimiento "
  sSql = sSql & "FROM CabComprobantes "
  sSql = sSql & "WHERE (CabComprobantes.Suc = " & nSucursal & ") "
  sSql = sSql & "AND (CabComprobantes.Tipo = 6 ) "
  sSql = sSql & "ORDER BY Id Desc"

  rs.CursorLocation = adUseClient
  rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerComprobantesRemitoVentas = rs
    
  Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerComprobantesRemitoCompras() As Recordset
  
 Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  db.Open sDb

  Set rs = New ADODB.Recordset

  sSql = "SELECT TOP 200 CabComprobantes.* "
  sSql = sSql & "FROM  CabComprobantes "
  sSql = sSql & "WHERE (CabComprobantes.Suc = " & nSucursal & ") "
  sSql = sSql & "AND (CabComprobantes.Tipo = 7) "
  sSql = sSql & "ORDER BY Id Desc"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerComprobantesRemitoCompras = rs
  
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerComprobantesCompras() As Recordset
  
 Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  db.Open sDb

  Set rs = New ADODB.Recordset

  sSql = "SELECT TOP 200 CabComprobantes.* "
  sSql = sSql & "FROM  CabComprobantes "
  sSql = sSql & "WHERE (CabComprobantes.Suc = " & nSucursal & ") "
  sSql = sSql & "AND (CabComprobantes.Tipo = 8 or  CabComprobantes.Tipo = 12) "
  sSql = sSql & "ORDER BY Id Desc"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerComprobantesCompras = rs
  
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
End Function


Public Function TraerComprobantesStock() As Recordset
  
 Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  db.Open sDb

  Set rs = New ADODB.Recordset

  sSql = "SELECT TOP 200 CabComprobantes.* "
  sSql = sSql & "FROM  CabComprobantes "
  sSql = sSql & "WHERE (CabComprobantes.Suc = " & nSucursal & ") "
  sSql = sSql & "AND CabComprobantes.Tipo = 2 "
  sSql = sSql & "ORDER BY Id Desc"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerComprobantesStock = rs
  
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
End Function


Public Function TraerComprobantesVentas() As Recordset
  
 Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  db.Open sDb

  Set rs = New ADODB.Recordset

  sSql = "SELECT TOP 100 CabComprobantes.* "
  sSql = sSql & "FROM  CabComprobantes "
  sSql = sSql & "WHERE (CabComprobantes.Suc = " & nSucursal & ")   AND (CabComprobantes.Tipo = 1 OR "
  sSql = sSql & "CabComprobantes.Tipo = 5) "
  sSql = sSql & "ORDER BY Id Desc"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerComprobantesVentas = rs
  
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerTodos(pTabla As String, pCampo As String, pTipo As Byte) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  db.Open sDb

  Set rs = New ADODB.Recordset
  sSql = "SELECT " & pCampo & " FROM " & pTabla
  If pTipo <> 2 Then
     If pTipo <> 1 Then
        If pTipo = 8 Then
           sSql = sSql & " WHERE Tipo=8 or Tipo=12 and Suc=" & nSucursal & " "
        Else
           sSql = sSql & " WHERE tipo=" & pTipo & " and Suc=" & nSucursal & " "
        End If
     Else
        sSql = sSql & " WHERE (tipo=1 or tipo=5 or Tipo=13) and Suc=" & nSucursal & " "
     End If
  Else
     sSql = sSql & " WHERE tipo=2 and Suc=" & nSucursal & " "
  End If
  sSql = sSql & " ORDER BY " & pCampo

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerTodos = rs
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerTodosSin(pTabla As String, pCampo As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb

  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT " & pCampo & " FROM " & pTabla
  sSql = sSql & " ORDER BY " & pCampo
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerTodosSin = rs
  db.Close
  
  Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function TraerCajaAbiertaTurno() As Double
  Dim sSql As String, cRsl As ClsLectura, nCaja As Integer, bPermiso As Boolean
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  Set cRsl = New ClsLectura
  
  nCaja = cRsl.RegPorDefecto("Caja")
  
  db.Open sDb
  
  bPermiso = cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=99 and Usuario=" & nUsuario)

  
  If bPermiso = True Then
     sSql = "SELECT (SUM(Debe)- SUM(Haber)) as Total "
     sSql = sSql & "From Caja "
     sSql = sSql & "WHERE Abierta='S'"
  Else
    sSql = "SELECT (SUM(Debe)- SUM(Haber)) as Total "
    sSql = sSql & "From Caja "
    sSql = sSql & "WHERE Abierta='S' AND Caja=" & nCaja
  End If
  
 
  
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerCajaAbiertaTurno = IIf(IsNull(rs!Total), 0, rs!Total)
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerUltimoNumero(sComp As Integer) As Long
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb
  db.Execute "UPDATE Comprobantes SET Numero=Numero+1 WHERE Id=" & sComp

  sSql = "SELECT Numero FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerUltimoNumero = rs!Numero
  End If
  db.Close
  
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerUltimoNumeroVenta(sComp As Integer) As Long
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Numero FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerUltimoNumeroVenta = rs!Numero
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function ImpuestoProducto(pProducto As String) As Double
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT ProductoImpuesto.Producto, Impuestos.Porcentaje "
  sSql = sSql & "FROM Impuestos RIGHT JOIN ProductoImpuesto ON Impuestos.Impuesto = ProductoImpuesto.Impuesto "
  sSql = sSql & "WHERE ProductoImpuesto.Producto='" & pProducto & "'"
  
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     ImpuestoProducto = rs!Porcentaje
  End If
  db.Close
  
End Function
Public Function TraerSucursalNumero(sComp As Integer) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Sucursal FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerSucursalNumero = rs!sucursal
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerTipoCondicion(nCond As Integer) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Tipo FROM CondVenta WHERE CondVta=" & nCond
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerTipoCondicion = rs!tipo
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerFinanciacion(nCond As Integer) As Single
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Financiacion FROM CondVenta WHERE CondVta=" & nCond
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerFinanciacion = rs!Financiacion
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerUltimoMovimiento() As Long
   Dim Conexion As ADODB.Connection, sSql As String
   Dim rs As ADODB.Recordset
 
   Dim Com As ADODB.Command
   Set Com = New ADODB.Command

   Set Conexion = New ADODB.Connection
   Set rs = New ADODB.Recordset
   
   Conexion.Open sDb
  ' db.Open sDb
   Conexion.Execute "UPDATE Sistema SET Movimiento=Movimiento+1"

   sSql = "SELECT Movimiento FROM Sistema "
   rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
   TraerUltimoMovimiento = rs!Movimiento
   Conexion.Close
End Function

Public Function TraerTipoMovimiento(sComp As Integer) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT TipoMovimiento FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerTipoMovimiento = rs!TipoMovimiento
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerTipoOperacion(sComp As Integer) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT TipoOperacion FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerTipoOperacion = rs!TipoOperacion
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function TraerTipoRecibo(sComp As Integer) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Cobranza FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerTipoRecibo = rs!Cobranza
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function ActualizaIva(sComp As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Iva FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  ActualizaIva = False
  If rs.EOF = False Or rs.BOF = False Then
     If rs!Iva = 1 Then ActualizaIva = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function ActualizaCtaCte(sComp As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT CtaCte FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  ActualizaCtaCte = False
  If rs.EOF = False Or rs.BOF = False Then
     If rs!ctacte <> "No" Then ActualizaCtaCte = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function ActualizaCaja(sComp As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Caja FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  ActualizaCaja = False
  If rs.EOF = False Or rs.BOF = False Then
     If rs!Caja <> "No" Then ActualizaCaja = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function SumaCaja(sComp As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Caja FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     If rs!Caja = "Suma" Then SumaCaja = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function ActualizaStock(sComp As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT MoviStock FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  ActualizaStock = False
  If rs.EOF = False Or rs.BOF = False Then
     If rs!MoviStock <> "No" Then ActualizaStock = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function QueMoviStock(sComp As Integer) As String
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT MoviStock FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  QueMoviStock = rs!MoviStock
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function ActualizaCosto(sComp As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Costo FROM Comprobantes WHERE Id=" & sComp
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     If rs!Costo = 1 Then ActualizaCosto = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function CantCuotas(sCond As Integer) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Cuotas FROM CondVenta WHERE condvta=" & sCond
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     CantCuotas = rs!Cuotas
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function CantDias(sCond As Integer) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Dias FROM CondVenta WHERE condvta=" & sCond
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     CantDias = rs!Dias
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function DepositoVenta() As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Deposito FROM Depositos WHERE Venta=1"
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  DepositoVenta = 0
  If rs.BOF = False Or rs.EOF = False Then
     DepositoVenta = rs!deposito
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

' Comprobantes de Cobro
Public Function DameCuentaCorriente(sCliDesde As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CuentaCorriente.* "
  sSql = sSql & "FROM CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE Comprobantes.TipoOperacion=1 "
  sSql = sSql & "AND Cliente='" & sCliDesde & "' and Anulado = 0 and Debe<>0 or (Cliente='" & sCliDesde & "' and Debe=0 and CuentaCorriente.Numero=0 and CuentaCorriente.sucursal=0 and CuentaCorriente.Cuota=0) "
  
  sSql = sSql & "ORDER BY CuentaCorriente.ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorriente = rs
  db.Close
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function DameCuentaCorrienteH(sCli As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT * FROM CuentaCorriente "
  sSql = sSql & "WHERE Cliente='" & sCli & "'  "
  sSql = sSql & "AND Anulado<> 1 and Haber<>0 and Cuota=1000 "
  sSql = sSql & "ORDER BY ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorrienteH = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function TraerKardexAnterior(pDesde As String, pHasta As String, pFdesde As Date) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Cantidades.Producto, Productos.Descripcion, Sum(Cantidades.CantidadDebe) AS SumaDeCantidadDebe, Sum(Cantidades.CantidadHaber) AS SumaDeCantidadHaber, (Sum(Cantidades.CantidadDebe)-Sum(Cantidades.CantidadHaber)) as Saldo "
  sSql = sSql & "FROM (((CabComprobantes RIGHT JOIN Cantidades ON CabComprobantes.Movimiento = Cantidades.Movimiento) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto) LEFT JOIN Depositos ON Cantidades.Deposito = Depositos.Deposito "
  sSql = sSql & "Where Productos.Descripcion >= '" & pDesde & "' and Productos.Descripcion<='" & pHasta & "' And CabComprobantes.Suc=" & nSucursal & " "
  sSql = sSql & "AND CabComprobantes.Fecha<'" & FechaAmericana(pFdesde) & "' And CabComprobantes.Suc=" & nSucursal & " "
  sSql = sSql & "GROUP BY Cantidades.Producto, Productos.Descripcion "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerKardexAnterior = rs
  db.Close
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function TraerKardex(pDesde As String, pHasta As String, pFdesde As Date, pFhasta As Date) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Tipo, CabComprobantes.Cliente, Depositos.Descripcion AS Deposito, CabComprobantes.Fecha, Cantidades.Producto, Productos.Descripcion AS ProdDesc, (Cantidades.CantidadDebe-Cantidades.CantidadHaber) AS Saldo, Comprobantes.Descripcion as ComprDesc, CabComprobantes.Sucursal, CabComprobantes.Numero,CabComprobantes.Motivo "
  sSql = sSql & "FROM (CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) RIGHT JOIN ((Cantidades LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto) LEFT JOIN Depositos ON Cantidades.Deposito = Depositos.Deposito) ON CabComprobantes.Movimiento = Cantidades.Movimiento "
  sSql = sSql & "Where Productos.Descripcion >= '" & pDesde & "' and Productos.Descripcion<='" & pHasta & "'"
  sSql = sSql & "AND CabComprobantes.Fecha>='" & FechaAmericana(pFdesde) & "' AND CabComprobantes.Fecha<='" & FechaAmericana(pFhasta) & "' "
  sSql = sSql & "ORDER BY CabComprobantes.Fecha, Cantidades.Id; "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerKardex = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerCedronar(pDesde As String, pHasta As String, pFdesde As Date, pFhasta As Date, pTipo As Byte) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  If pTipo = 0 Then
  
    sSql = "SELECT CabComprobantes.Cliente, Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Cuit, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, Sum(DetallesComprobantes.Cantidad) AS Cantidad "
    sSql = sSql & "FROM ((CabComprobantes LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN Proveedores ON CabComprobantes.Cliente = Proveedores.Proveedor) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
    sSql = sSql & "Where (((CabComprobantes.Fecha) >='" & FechaAmericana(pFdesde) & "' And (CabComprobantes.Fecha) <= '" & FechaAmericana(pFhasta) & "') And ((Comprobantes.Iva) = 1)) And CabComprobantes.Suc=" & nSucursal & " "
    sSql = sSql & "GROUP BY CabComprobantes.Cliente, Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Cuit, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, CabComprobantes.Tipo "
    sSql = sSql & "HAVING (((DetallesComprobantes.Producto)='" & pDesde & "') AND ((CabComprobantes.Tipo)=8)) "
    sSql = sSql & "ORDER BY Proveedores.RazonSocial "
  Else
    sSql = "SELECT CabComprobantes.Cliente, Clientes.RazonSocial, Clientes.Cuit, Cantidades.Producto, Productos.Descripcion, (Sum(Cantidades.CantidadHaber)-Sum(Cantidades.CantidadDebe)) AS Cantidad "
    sSql = sSql & "FROM ((CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) RIGHT JOIN ((Cantidades LEFT JOIN Productos ON Cantidades.Producto = Productos.Producto) LEFT JOIN Depositos ON Cantidades.Deposito = Depositos.Deposito) ON CabComprobantes.Movimiento = Cantidades.Movimiento) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente "
    sSql = sSql & "WHERE CabComprobantes.Tipo=1 AND CabComprobantes.Fecha>='" & FechaAmericana(pFdesde) & "' And CabComprobantes.Fecha<='" & FechaAmericana(pFhasta) & "' AND CabComprobantes.Anulado=0 And CabComprobantes.Suc=" & nSucursal & " "
    sSql = sSql & "GROUP BY CabComprobantes.Cliente, Clientes.RazonSocial, Clientes.Cuit, Cantidades.Producto,Productos.Producto, Productos.Descripcion "
    sSql = sSql & "HAVING Productos.Producto>='" & pDesde & "' And Productos.Producto<='" & pHasta & "' "
    sSql = sSql & "ORDER BY Clientes.RazonSocial "
   End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerCedronar = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function DameCuentaCorrienteP(sCliDesde As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT CuentaCorriente.* "
  sSql = sSql & "FROM CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE Comprobantes.TipoOperacion=8 "
  sSql = sSql & "AND CuentaCorriente.Cliente='" & sCliDesde & "' and Anulado<> 1  and Haber<>0 "
  sSql = sSql & "ORDER BY CuentaCorriente.ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorrienteP = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function DameCuentaCorrienteHP(sCli As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT * FROM CuentaCorriente WHERE Cliente='" & sCli & "'  and Anulado<> 1 and Debe<>0 and Cuota=1000  ORDER BY ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorrienteHP = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function DameValores(sVent As Long) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT FormaPago,Importe FROM Valores WHERE Venta=" & sVent & " and PorVenta='X'  ORDER BY ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameValores = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function DameCodComprobante(nCod As Integer) As String
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Descripcion FROM Comprobantes WHERE id=" & nCod
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     DameCodComprobante = rs!Descripcion
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
 
End Function

Public Function DamePagos(pCuota As Integer, pVenta As Long, Optional Auxi As Boolean) As Currency
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
'  db.Open sDb

  If Auxi = False Then
     sSql = "SELECT SUM(Aplicaciones.Importe) AS total "
     sSql = sSql & "FROM Aplicaciones "
     sSql = sSql & "WHERE Aplicaciones.Venta=" & pVenta & " AND Aplicaciones.Cuota=" & pCuota & " AND Aplicaciones.Anulado<>1"
  Else
     sSql = "SELECT SUM(AplicacionAuxiliar.Importe) AS total "
     sSql = sSql & "FROM AplicacionAuxiliar "
     sSql = sSql & "WHERE AplicacionAuxiliar.Venta=" & pVenta & " AND AplicacionAuxiliar.Cuota=" & pCuota & " AND AplicacionAuxiliar.Anulado<>1"
  End If
          
  rs.Open sSql, db1, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     If Not IsNull(rs!Total) Then
        DamePagos = rs!Total
     End If
  End If
 ' db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function


Public Function ListaConIva(pLista As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
    
  db.Open sDb
  
  sSql = "SELECT PrecioIva "
  sSql = sSql & "FROM ListadePrecio WHERE ListaPrecio=" & pLista
  
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If rs.EOF = False Or rs.BOF = False Then
     If rs!PrecioIva = 1 Then
        ListaConIva = True
     End If
  End If
  db.Close
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function DiscriminaIva(sCliente As String) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT TipoIva.Descripcion, TipoIva.Discrimina "
  sSql = sSql & "FROM Clientes LEFT JOIN TipoIva ON Clientes.TipoIva = TipoIva.TipoIva "
  sSql = sSql & "WHERE Clientes.Cliente='" & sCliente & "'"
 
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing

  If rs.EOF = False Or rs.BOF = False Then
     Set DiscriminaIva = rs
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerIdComprobante(nComp As Integer, nSuc As Long, nNum As Long) As Long
  Dim sSql As String, rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT id FROM Cabcomprobantes WHERE Comprobante=" & nComp & " and Numero=" & nNum & " and Sucursal=" & nSuc
 ' sSql = sSql & " AND Anulado=0"
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerIdComprobante = rs!ID
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerCantidad(sArt As String, nDep As Integer) As Single
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT (Sum(CantidadDebe) - Sum(CantidadHaber)) as Cantidad FROM Cantidades WHERE Anulado<>1 and Producto='" & sArt & "' and Deposito=" & nDep
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If rs.EOF = False Or rs.BOF = False Then
     TraerCantidad = IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
  End If
  db.Close
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerMovimientoStock(pPDesde As String, pPHasta As String, pFdesde As Date, pFhasta As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Id, CabComprobantes.Tipo, TipoMovimiento.Descripcion, Comprobantes.Descripcion, CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Numero, CabComprobantes.Sucursal, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, Depositos.Descripcion, Depositos_1.Descripcion, DetallesComprobantes.Cantidad "
  sSql = sSql & "FROM (((Comprobantes RIGHT JOIN (CabComprobantes LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) ON Comprobantes.Id = CabComprobantes.Comprobante) LEFT JOIN Depositos ON CabComprobantes.Desde = Depositos.Deposito) LEFT JOIN Depositos AS Depositos_1 ON CabComprobantes.Hasta = Depositos_1.Deposito) LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento "
  sSql = sSql & "Where (((CabComprobantes.Tipo) <= 2) And ((DetallesComprobantes.Producto) <> "" and DetallesComprobantes.Descripcion>='" & pPDesde & "' or  DetallesComprobantes.Descripcion<='" & pPHasta & "' and CabComprobantes.Fecha>='" & pFdesde & "' or CabComprobantes.Fecha<='" & pFhasta & "' )) "
  sSql = sSql & "ORDER BY CabComprobantes.Id"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  If rs.EOF = False Or rs.BOF = False Then
     Set TraerMovimientoStock = rs
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerRecibosCobros(pCliente As String) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CuentaCorriente.Venta, CuentaCorriente.Fecha,  CuentaCorriente.Comprobante, Comprobantes.Descripcion,CuentaCorriente.Sucursal, CuentaCorriente.Numero, CuentaCorriente.Haber "
  sSql = sSql & "FROM CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE Comprobantes.TipoMovimiento=10 AND ((CuentaCorriente.Cliente='" & pCliente & "'))"
  sSql = sSql & "ORDER BY  CuentaCorriente.Fecha,Comprobantes.Descripcion,CuentaCorriente.Sucursal ,CuentaCorriente.Numero DESC"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerRecibosCobros = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function TraerOrdenPagos(pCliente As String) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CuentaCorriente.Venta, CuentaCorriente.Fecha,  CuentaCorriente.Comprobante, Comprobantes.Descripcion,CuentaCorriente.Sucursal, CuentaCorriente.Numero, CuentaCorriente.Debe "
  sSql = sSql & "FROM CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE Comprobantes.TipoMovimiento=15 AND CuentaCorriente.Cliente='" & pCliente & "' "
  sSql = sSql & "ORDER BY  CuentaCorriente.Fecha,Comprobantes.Descripcion,CuentaCorriente.Sucursal ,CuentaCorriente.Numero ASC"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerOrdenPagos = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerCajaAbierta(pCaja As Byte) As Recordset
  On Error GoTo errHandler
  Dim conex As ADODB.Connection, sSql As String
  Dim rs As ADODB.Recordset

  Set conex = New ADODB.Connection
  conex.ConnectionString = db
  conex.Open
  
  Set rs = New ADODB.Recordset
  sSql = "SELECT Comprobantes.Descripcion, Caja.Sucursal, Caja.Numero, Caja.Debe, Caja.Haber "
  sSql = sSql & "FROM (((((Caja LEFT JOIN CabComprobantes ON Caja.Movimiento = CabComprobantes.Movimiento) LEFT JOIN Cajas ON Caja.Caja = Cajas.Caja) LEFT JOIN Comprobantes ON Caja.Comprobante = Comprobantes.Id) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN Vendedores ON CabComprobantes.Vendedor = Vendedores.Vendedor) LEFT JOIN Comprobantes AS Comprobantes_1 ON CabComprobantes.Comprobante = Comprobantes_1.Id "
  sSql = sSql & "Where (((Caja.Abierta) = 'S') And ((Caja.Anulado) <> 1) and Caja.Caja =" & pCaja & " ) "
  sSql = sSql & "ORDER BY Caja.Id"
  
  rs.ActiveConnection = conex
  rs.Properties("IRowsetIdentity") = True
  
  rs.Open sSql, , adOpenKeyset, adLockPessimistic
  
  Set TraerCajaAbierta = rs
    
  conex.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerCajaSaldos(pCaja As Byte) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Sum(Caja.Debe) as TotalDebe, Sum(Caja.Haber) as TotalHaber, (Sum(Caja.Debe) -Sum(Caja.Haber)) as Saldo "
  sSql = sSql & "From Caja "
  sSql = sSql & "WHERE (((Caja.Abierta)='S') AND ((Caja.Anulado)<>1) and Caja.Caja=" & pCaja & ")"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerCajaSaldos = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerMoviCaja(sDesde As Date, sHasta As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Caja.Fecha, Comprobantes.Descripcion, Caja.Sucursal, Caja.Numero, Clientes.RazonSocial, Vendedores.Descripcion, Comprobantes_1.Descripcion, CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, Cajas.Descripcion, Caja.Hora, Caja.FormaPago, Caja.Nombre, Bancos.Descripcion, Caja.NumeroCheque, Caja.FechaAcreditacion, Caja.NumeroTarjeta, Caja.FechaVencimiento, Caja.Autorizacion, Caja.Cupon, Caja.Debe, Caja.Haber "
  sSql = sSql & "FROM ((((((Caja LEFT JOIN CabComprobantes ON Caja.Movimiento = CabComprobantes.Movimiento) LEFT JOIN Cajas ON Caja.Caja = Cajas.Caja) LEFT JOIN Comprobantes ON Caja.Comprobante = Comprobantes.Id) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN Vendedores ON CabComprobantes.Vendedor = Vendedores.Vendedor) LEFT JOIN Comprobantes AS Comprobantes_1 ON CabComprobantes.Comprobante = Comprobantes_1.Id) LEFT JOIN Bancos ON Caja.Banco = Bancos.Banco "
  sSql = sSql & "Where (((Caja.Anulado) <> 1)) and Caja.Fecha>='" & FechaAmericana(sDesde) & "' and Caja.Fecha<='" & FechaAmericana(sHasta) & "'"
  sSql = sSql & "ORDER BY Caja.Id"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerMoviCaja = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerMoviVentas(sDes As Date, sHas As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Numero, CabComprobantes.Sucursal,"
  sSql = sSql & "Clientes.RazonSocial,Vendedores.Descripcion, Cajas.Descripcion,"
  sSql = sSql & "(CabComprobantes.Debe - CabComprobantes.Haber * CabComprobantes.Cotizacion) as Importe,"
  sSql = sSql & "CabComprobantes.IBrutosCompras,"
  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN ((CabComprobantes.Neto - CabComprobantes.Descuentos)  * CabComprobantes.Cotizacion)* CabComprobantes.IBrutosCompras / 100 "
  sSql = sSql & "ELSE -(((CabComprobantes.Neto - CabComprobantes.Descuentos)  * CabComprobantes.Cotizacion)* CabComprobantes.IBrutosCompras / 100)END,"
  sSql = sSql & "Zonas.Descripcion AS Zona "
  sSql = sSql & "FROM CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "Clientes ON CabComprobantes.Cliente = Clientes.Cliente LEFT OUTER JOIN "
  sSql = sSql & "Vendedores ON CabComprobantes.Vendedor = Vendedores.Vendedor LEFT OUTER JOIN "
  sSql = sSql & "Cajas ON CabComprobantes.Caja = Cajas.Caja LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id LEFT OUTER JOIN "
  sSql = sSql & "Zonas ON Clientes.Zona = Zonas.Zona "
  sSql = sSql & "WHERE (((CabComprobantes.Anulado)<>1) AND (CabComprobantes.Tipo=1 or CabComprobantes.Tipo=5)) And CabComprobantes.Suc=" & nSucursal & " "
  sSql = sSql & " and CabComprobantes.Fecha>='" & FechaAmericana(sDes) & "' and CabComprobantes.Fecha<='" & FechaAmericana(sHas) & "' and CabComprobantes.CondVenta<>2 "
  sSql = sSql & "ORDER BY CabComprobantes.Sucursal, CabComprobantes.Numero"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerMoviVentas = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerMoviVentasProduc(sDes As Date, sHas As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Tipo,Vendedores.Descripcion as Vendedor, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion as Detalle, CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Numero, CabComprobantes.Sucursal, Clientes.RazonSocial, Vendedores.Descripcion, Cajas.Descripcion as Caja, DetallesComprobantes.Cantidad ,DetallesComprobantes.Descuento,DetallesComprobantes.PrecioUnitario,DetallesComprobantes.PrecioTotal,DetallesComprobantes.Impuesto,Comprobantes.Comprobante "
  sSql = sSql & "FROM ((((CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN Vendedores ON CabComprobantes.Vendedor = Vendedores.Vendedor) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN Cajas ON CabComprobantes.Caja = Cajas.Caja) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "WHERE (((CabComprobantes.Anulado)<>1) AND ((CabComprobantes.Tipo)=1 or CabComprobantes.Tipo=5)) And CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.CondVenta<>2 "
  sSql = sSql & " and CabComprobantes.Fecha>='" & FechaAmericana(sDes) & "' and CabComprobantes.Fecha<='" & FechaAmericana(sHas) & "' "
  sSql = sSql & " ORDER BY CabComprobantes.Fecha,DetallesComprobantes.Descripcion, CabComprobantes.Sucursal,CabComprobantes.Numero"
    
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerMoviVentasProduc = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerDetalles(nMovi As Long) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT CabComprobantes.Fecha, Comprobantes.Descripcion as CompCompro, CabComprobantes.Numero, DetallesComprobantes.Producto, DetallesComprobantes.Cantidad,DetallesComprobantes.Descripcion as DesCompro "
  sSql = sSql & "FROM (CabComprobantes LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE (((CabComprobantes.Movimiento)=" & nMovi & "))"
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerDetalles = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function ComprobanteAnulado(nMovi As Long) As Boolean
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Anulado FROM CabComprobantes WHERE Movimiento=" & nMovi
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  
  If rs!Anulado = 1 Then
     ComprobanteAnulado = True
  End If
  db.Close
  Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function ProductoImpuesto(sProdu As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Impuestos.Porcentaje, Impuestos.Interno "
  sSql = sSql & "FROM Impuestos RIGHT JOIN ProductoImpuesto ON Impuestos.Impuesto = ProductoImpuesto.Impuesto "
  sSql = sSql & "WHERE ProductoImpuesto.Producto='" & sProdu & "'"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set ProductoImpuesto = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerClientesRemitos(sClien As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb
            
  sSql = "SELECT CabComprobantes.Movimiento,CabComprobantes.Cliente, CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE (((CabComprobantes.Tipo)=6) AND CabComprobantes.Factura=0 AND ((CabComprobantes.Anulado)=0)) AND (Comprobantes.TipoMovimiento=17 or Comprobantes.TipoMovimiento=33) And CabComprobantes.Suc=" & nSucursal & " "
  sSql = sSql & "AND CabComprobantes.Cliente='" & sClien & "' and CabComprobantes.Motivo='' "
        
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerClientesRemitos = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerProveedorRemitos(sClien As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Movimiento,CabComprobantes.Cliente, CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE (((CabComprobantes.Factura)=0) AND ((CabComprobantes.Tipo)=7) AND ((CabComprobantes.Anulado)=0)) AND Comprobantes.TipoMovimiento=30 "
  sSql = sSql & "AND CabComprobantes.Cliente='" & sClien & "' and CabComprobantes.Suc=" & nSucursal
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerProveedorRemitos = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerClientesRemitosDev(sClien As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Movimiento,CabComprobantes.Cliente, CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE CabComprobantes.Factura=0 AND CabComprobantes.Tipo=6 AND CabComprobantes.Anulado=0  "
  sSql = sSql & "AND CabComprobantes.Cliente='" & sClien & "'"
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerClientesRemitosDev = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function
Public Function TraerClientesPedidos(sClien As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Movimiento,CabComprobantes.Cliente, CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE (((CabComprobantes.Factura)=0) AND ((CabComprobantes.Tipo)=10) AND ((CabComprobantes.Anulado)=0) AND ((Comprobantes.TipoMovimiento)=20)) "
  sSql = sSql & "AND CabComprobantes.Cliente='" & sClien & "'"
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerClientesPedidos = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerClientesPedidosDev(sClien As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Movimiento,CabComprobantes.Cliente, CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE (((CabComprobantes.Factura)=0) AND ((CabComprobantes.Tipo)=10) AND ((CabComprobantes.Anulado)=0) AND ((Comprobantes.TipoMovimiento)=20)) "
  sSql = sSql & "AND CabComprobantes.Cliente='" & sClien & "'"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerClientesPedidosDev = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerDetallesMovimientoRemi(pMovi As Long) As Single
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT   SUM(DetallesComprobantes.Cantidad) AS Cantidad "
  sSql = sSql & "FROM  CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "Where (CabComprobantes.Movimiento = " & pMovi & ") "

  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  
  rs.Close
  rs.CursorType = adOpenStatic
  rs.CursorLocation = adUseClient
  rs.LockType = adLockBatchOptimistic
  rs.Open
  rs.ActiveConnection = Nothing

  If rs.RecordCount <> 0 Then
     TraerDetallesMovimientoRemi = rs!Cantidad
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerDetallesMovimiento(pMovi As Long) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  
  db.Open sDb

  sSql = "SELECT DetallesComprobantes.Id,CabComprobantes.Comprobante,DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad,DetallesComprobantes.Tasa,DetallesComprobantes.Deposito,DetallesComprobantes.Medida,DetallesComprobantes.Impuesto,DetallesComprobantes.Precio,DetallesComprobantes.PrecioUnitario,DetallesComprobantes.PrecioTotal "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "WHERE CabComprobantes.Movimiento=" & pMovi
  

  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  
  rs.Close
  rs.CursorType = adOpenStatic
  rs.CursorLocation = adUseClient
  rs.LockType = adLockBatchOptimistic
  rs.Open
  rs.ActiveConnection = Nothing

  
  Set TraerDetallesMovimiento = rs

  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function


Public Function SumaCantidad(pCompr As Integer) As Boolean
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT TipoMovimiento.Descripcion "
  sSql = sSql & "FROM Comprobantes LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento "
  sSql = sSql & "WHERE Comprobantes.Id=" & pCompr
  
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If rs!Descripcion = "Resta" Then
        SumaCantidad = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function SumaCantidadProv(pCompr As Integer) As Boolean
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT TipoMovimiento.Descripcion "
  sSql = sSql & "FROM Comprobantes LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento "
  sSql = sSql & "WHERE Comprobantes.Id=" & pCompr
  
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If rs!Descripcion = "Suma" Then
        SumaCantidadProv = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerInfRemitosPend(sDesde As String, sHasta As String, pDesde As Date, pHasta As Date, Optional sZodes As Integer = 0) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb



  sSql = "SELECT Comprobantes.Descripcion As Comprobante,CabComprobantes.Movimiento ,CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente, Clientes.RazonSocial, "
  sSql = sSql & "CASE "
  sSql = sSql & "WHEN TipoMovimiento.Descripcion='Resta' "
  sSql = sSql & "THEN -Sum(DetallesComprobantes.PrecioTotal * CASE  WHEN Monedas.Cotizacion IS NULL THEN 1 ELSE Monedas.Cotizacion END) "
  sSql = sSql & "ELSE Sum (DetallesComprobantes.PrecioTotal * CASE WHEN Monedas.Cotizacion IS NULL  THEN 1 ELSE Monedas.Cotizacion END)"
  sSql = sSql & "END AS Expr1 ,TipoMovimiento.Descripcion,CabComprobantes.Movimiento "
  sSql = sSql & "FROM  CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id LEFT OUTER JOIN "
  sSql = sSql & "Clientes ON CabComprobantes.Cliente = Clientes.Cliente LEFT OUTER JOIN "
  sSql = sSql & "TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento LEFT OUTER JOIN "
  sSql = sSql & "DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento LEFT OUTER JOIN "
  sSql = sSql & "Productos ON DetallesComprobantes.Producto = Productos.Producto LEFT OUTER JOIN "
  sSql = sSql & "Monedas ON Productos.Moneda = Monedas.Id "
  sSql = sSql & "WHERE CabComprobantes.Suc=" & nSucursal & " "

  If sZodes <> 0 Then
     sSql = sSql & "AND Clientes.Zona=" & sZodes & " "
  End If
  ' ,Monedas.Cotizacion
  
  sSql = sSql & "GROUP BY Comprobantes.Descripcion,CabComprobantes.Movimiento ,CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente, Clientes.RazonSocial, TipoMovimiento.Descripcion, CabComprobantes.Id, CabComprobantes.Tipo, CabComprobantes.Factura, CabComprobantes.Anulado,CabComprobantes.Movimiento "
  sSql = sSql & "HAVING CabComprobantes.Tipo=6 AND CabComprobantes.Factura=0 AND CabComprobantes.Anulado=0  "
  sSql = sSql & "AND Clientes.RazonSocial>='" & sDesde & "' AND Clientes.RazonSocial<='" & sHasta & "' "
  sSql = sSql & "AND CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' AND CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "' "
  sSql = sSql & "ORDER BY  Clientes.RazonSocial,Cabcomprobantes.Id"
   
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  rs.Close
  rs.CursorType = adOpenStatic
  rs.CursorLocation = adUseClient
  rs.LockType = adLockBatchOptimistic
  rs.Open
  rs.ActiveConnection = Nothing

 Set TraerInfRemitosPend = rs
 db.Close
 
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerInfRemitosPendDet(sDesde As String, sHasta As String, pDesde As Date, pHasta As Date) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb
 
  sSql = "SELECT Comprobantes.Descripcion as Comprobante, CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente, Clientes.RazonSocial, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad, "
  sSql = sSql & "case "
  sSql = sSql & "WHEN TipoMovimiento.Descripcion='Resta' "
  sSql = sSql & "then - DetallesComprobantes.PrecioTotal "
  sSql = sSql & "Else DetallesComprobantes.PrecioTotal "
  sSql = sSql & "end as  Importe,CabComprobantes.Movimiento,CabComprobantes.Motivo "
  sSql = sSql & "FROM (((CabComprobantes LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento "
  sSql = sSql & "WHERE (((CabComprobantes.Tipo)=6) AND ((CabComprobantes.Factura)=0) AND ((CabComprobantes.Anulado)=0)) "
  sSql = sSql & "AND Clientes.RazonSocial>='" & sDesde & "' AND Clientes.RazonSocial<='" & sHasta & "' "
  sSql = sSql & "AND CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' AND CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "'and CabComprobantes.Suc=" & nSucursal & " "
  sSql = sSql & "AND CabComprobantes.Motivo Not like 'Factura Nº:%' "
  
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  rs.Close
  rs.CursorType = adOpenStatic
  rs.CursorLocation = adUseClient
  rs.LockType = adLockBatchOptimistic
  rs.Open
  rs.ActiveConnection = Nothing

  Set TraerInfRemitosPendDet = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerTodosRemitosPend(sDesde As String, sHasta As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

' sSql = "SELECT DISTINCT Comprobantes.Descripcion, CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente, Clientes.RazonSocial, IIf(TipoMovimiento.Descripcion='Resta',-Sum(DetallesComprobantes.PrecioTotal),Sum(DetallesComprobantes.PrecioTotal)) AS Neto, TipoMovimiento.Descripcion "
' sSql = sSql & "FROM (((CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
' sSql = sSql & "GROUP BY Comprobantes.Descripcion, CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente, Clientes.RazonSocial, TipoMovimiento.Descripcion, CabComprobantes.Id, CabComprobantes.Tipo, CabComprobantes.Factura, CabComprobantes.Anulado "
' sSql = sSql & "HAVING (((CabComprobantes.Tipo)=6) AND ((CabComprobantes.Factura)=0) AND ((CabComprobantes.Anulado)=0)) "
' sSql = sSql & "AND Clientes.RazonSocial>='" & sDesde & "' AND Clientes.RazonSocial<='" & sHasta & "' "
'' sSql = sSql & "ORDER BY  Clientes.RazonSocial,Cabcomprobantes.Id"
' sSql = sSql & "ORDER BY  Clientes.RazonSocial"
  sSql = "SELECT DISTINCT CabComprobantes.Cliente, Clientes.RazonSocial "
  sSql = sSql & "FROM (((CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
  sSql = sSql & "GROUP BY CabComprobantes.Cliente, Clientes.RazonSocial, CabComprobantes.Id, CabComprobantes.Tipo, CabComprobantes.Factura, CabComprobantes.Anulado "
 sSql = sSql & "HAVING (((CabComprobantes.Tipo)=6) AND ((CabComprobantes.Factura)=0) AND ((CabComprobantes.Anulado)=0)) "
 sSql = sSql & "AND Clientes.RazonSocial>='" & sDesde & "' AND Clientes.RazonSocial<='" & sHasta & "' "
'' sSql = sSql & "ORDER BY  Clientes.RazonSocial,Cabcomprobantes.Id"
 sSql = sSql & "ORDER BY  Clientes.RazonSocial"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
 Set TraerTodosRemitosPend = rs
 db.Close
 
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function SumoOresto(pCompro As String) As Boolean
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Comprobantes.Descripcion, TipoMovimiento.Descripcion as Tipo "
  sSql = sSql & "FROM Comprobantes LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento "
  sSql = sSql & "WHERE Comprobantes.Comprobante='" & pCompro & "'"
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  
  If rs!tipo = "Credito" Then
    SumoOresto = True
  End If
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerMateriaPrima(pProduc As String) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

   db.Open sDb

  sSql = "SELECT DetalleReceta.Producto, DetalleReceta.Cantidad "
  sSql = sSql & "FROM CabReceta LEFT JOIN DetalleReceta ON CabReceta.Movimiento = DetalleReceta.Movimiento "
  sSql = sSql & "WHERE CabReceta.Producto='" & pProduc & "'"
  

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerMateriaPrima = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerOtrasCajas() As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  
  db.Open sDb

  sSql = "SELECT DISTINCT Caja.FormaPago, (Sum(Caja.Debe)-Sum(Caja.Haber)) AS Total "
'  sSql = sSql & "FROM ((CabComprobantes LEFT JOIN Caja ON CabComprobantes.Movimiento = Caja.Movimiento) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) "
  sSql = sSql & "FROM Caja "
  sSql = sSql & "Where (((Caja.Anulado) = 0) ) And Caja.Suc=" & nSucursal & " "
  sSql = sSql & "GROUP BY Caja.FormaPago "
  sSql = sSql & "HAVING ((([Caja].[FormaPago])<>'CHEQUES DE TERCEROS' And ([Caja].[FormaPago])<>'INTERDEPOSITOS')) "
  
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
 Set TraerOtrasCajas = rs
 db.Close
 
Exit Function
errHandler:
  err.Raise err.Number


End Function

Public Function TraerCajacheque() As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT DISTINCT Caja.FormaPago, Bancos.Descripcion, Cheques.Numero, Cheques.Titular, Cheques.FechaAcre, Cheques.Importe "
  sSql = sSql & "FROM (Caja RIGHT JOIN Cheques ON (Caja.NumeroCheque = Cheques.Numero) AND (Caja.Banco = Cheques.Banco)) LEFT JOIN Bancos ON Cheques.Banco = Bancos.Banco "
  sSql = sSql & "WHERE Cheques.FechaSalida Is Null AND Cheques.Anulado=0 and Caja.Suc=" & nSucursal & " "
  sSql = sSql & "ORDER BY Cheques.FechaAcre"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerCajacheque = rs
  db.Close
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerMovimientoCaja(pDesde As Date, pHasta As Date) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT Caja.Fecha, Comprobantes.Descripcion, Caja.FormaPago, Caja.Nombre, CabComprobantes.Motivo, Caja.Debe, Caja.Haber "
  sSql = sSql & "FROM CabComprobantes RIGHT JOIN (Caja LEFT JOIN Comprobantes ON Caja.Comprobante = Comprobantes.Id) ON CabComprobantes.Movimiento = Caja.Movimiento "
  sSql = sSql & "WHERE Caja.Anulado=0 AND Caja.Fecha>='" & FechaAmericana(pDesde) & "' AND Caja.Fecha<='" & FechaAmericana(pHasta) & "' AND Caja.Suc=" & nSucursal & " "
  sSql = sSql & "ORDER BY Caja.Fecha"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerMovimientoCaja = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerIvaVenta(pDesde As Date, pHasta As Date, pAgru As Boolean) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Movimiento, CabComprobantes.Fecha,Comprobantes.Id as Compr,Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero, Clientes.RazonSocial, Clientes.Cuit, CabComprobantes.Neto, CabComprobantes.Iva1, CabComprobantes.NoGravados,CabComprobantes.PIvaCompras, CabComprobantes.Debe , CabComprobantes.Haber, CabComprobantes.Anulado "
  sSql = sSql & "FROM (CabComprobantes  LEFT JOIN AfipCaes ON CabComprobantes.Movimiento = AfipCaes.Movimiento LEFT OUTER JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente "
  sSql = sSql & "WHERE (CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' AND CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "' And Anulado=0  and AfipCaes.Movimiento is not null)  OR (Comprobantes.Afip=80 and CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' AND CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "') "
  sSql = sSql & "GROUP BY CabComprobantes.Movimiento,CabComprobantes.Fecha,CabComprobantes.Id,Comprobantes.Id ,Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero, Clientes.RazonSocial, Clientes.Cuit, CabComprobantes.Neto, CabComprobantes.Iva1, CabComprobantes.NoGravados,CabComprobantes.PIvaCompras, CabComprobantes.Debe, CabComprobantes.Haber, Comprobantes.Iva, CabComprobantes.Tipo, CabComprobantes.Anulado,AfipCaes.Movimiento "
  sSql = sSql & "Having (((Comprobantes.Iva) = 1) And ((CabComprobantes.Tipo) = 1 Or (CabComprobantes.Tipo) = 5 Or (CabComprobantes.Tipo) = 13)) "
  If pAgru = False Then
   '  sSql = sSql & "ORDER BY CabComprobantes.Fecha,CabComprobantes.Id "
     sSql = sSql & "ORDER BY CabComprobantes.Fecha,Comprobantes.Descripcion,CabComprobantes.Sucursal,CabComprobantes.Numero "
  Else
       sSql = sSql & "ORDER BY Comprobantes.Descripcion,CabComprobantes.Fecha,CabComprobantes.Sucursal,CabComprobantes.Numero "
  End If

  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  rs.Close
  rs.CursorType = adOpenStatic
  rs.CursorLocation = adUseClient
  rs.LockType = adLockBatchOptimistic
  rs.Open
  rs.ActiveConnection = Nothing
  
  Set TraerIvaVenta = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerDetIvaVenta(pMovimiento) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  

  sSql = "SELECT DetallesComprobantes.Tasa, Sum(DetallesComprobantes.Impuesto ) AS Impuesto,Sum(DetallesComprobantes.Descuento ) AS Descuento ,Sum(DetallesComprobantes.PrecioTotal) AS PrecioTotal "
  sSql = sSql & "From DetallesComprobantes "
  sSql = sSql & "Where (((DetallesComprobantes.Movimiento) =" & pMovimiento & ")) " ' and DetallesComprobantes.Tasa<>21 "
  sSql = sSql & "GROUP BY DetallesComprobantes.Tasa "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db1, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing

  Set TraerDetIvaVenta = rs
  
  Exit Function
End Function

Public Function TraerDetIvaCompra(pMovimiento) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  

  sSql = "SELECT DetallesComprobantes.Tasa, Sum(DetallesComprobantes.Impuesto) AS Impuesto, Sum(DetallesComprobantes.PrecioTotal) AS PrecioTotal "
  sSql = sSql & "From DetallesComprobantes "
  sSql = sSql & "Where (((DetallesComprobantes.Movimiento) =" & pMovimiento & ")) "
  sSql = sSql & "GROUP BY DetallesComprobantes.Tasa "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db1, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing

  Set TraerDetIvaCompra = rs
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerIvaCompra(pDesde As Date, pHasta As Date, pAgru As Boolean) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT CabComprobantes.Movimiento, CabComprobantes.FechaIva, Comprobantes.Id as Compr,Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero, Proveedores.RazonSocial, Proveedores.Cuit, CabComprobantes.Neto, CabComprobantes.Iva1, CabComprobantes.NoGravados, CabComprobantes.PIvaCompras, CabComprobantes.IBrutosCompras, CabComprobantes.Debe, CabComprobantes.Haber "
  sSql = sSql & "FROM (CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN Proveedores ON CabComprobantes.Cliente = Proveedores.Proveedor "
 ' sSql = sSql & "WHERE convert(datetime,CabComprobantes.Motivo,101)>='" & pDesde & "' AND convert(datetime,CabComprobantes.Motivo,101)<='" & pHasta & "' and CabComprobantes.Suc=" & nSucursal & " "
'  sSql = sSql & "WHERE CabComprobantes.Motivo>='" & FechaAmericanaIva(pDesde) & "' AND CabComprobantes.Motivo<='" & FechaAmericanaIva(pHasta) & "' and CabComprobantes.Suc=" & nSucursal & " "
  sSql = sSql & "WHERE CabComprobantes.FEchaIva>='" & pDesde & "' AND CabComprobantes.FechaIva<='" & pHasta & "' and CabComprobantes.Suc=" & nSucursal & " "
  
  sSql = sSql & "GROUP BY CabComprobantes.Movimiento, CabComprobantes.FechaIva, Comprobantes.Id,Comprobantes.Descripcion, CabComprobantes.Sucursal, CabComprobantes.Numero, Proveedores.RazonSocial, Proveedores.Cuit, CabComprobantes.Neto, CabComprobantes.Iva1, CabComprobantes.NoGravados, CabComprobantes.PIvaCompras, CabComprobantes.IBrutosCompras, CabComprobantes.Debe, CabComprobantes.Haber, CabComprobantes.Id, Comprobantes.Iva, CabComprobantes.Tipo, CabComprobantes.Anulado "
  sSql = sSql & "Having (((Comprobantes.Iva) = 1) And ((CabComprobantes.Tipo) = 8 Or (CabComprobantes.Tipo) = 12) And ((CabComprobantes.Anulado) = 0)) "
'  sSql = sSql & "ORDER BY convert(datetime,CabComprobantes.Motivo,101), CabComprobantes.Id "
  If pAgru = False Then
     sSql = sSql & "ORDER BY CabComprobantes.FechaIva, CabComprobantes.Id "
  Else
       sSql = sSql & "ORDER BY Comprobantes.Descripcion,CabComprobantes.FechaIva, CabComprobantes.Sucursal,CabComprobantes.Numero "
  End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerIvaCompra = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerChequesTerceros(pDesde As Date, pHasta As Date, pOrd As Boolean, pCartera As Boolean) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

   db.Open sDb

  sSql = "SELECT Cheques.Fecha, Cheques.Id, Bancos.Descripcion, Cheques.Numero, Cheques.Titular, Proveedores.RazonSocial, Cheques.FechaAcre, Cheques.Importe, Cheques.FechaSalida "
  sSql = sSql & "FROM (Cheques LEFT JOIN Bancos ON Cheques.Banco = Bancos.Banco) LEFT JOIN Proveedores ON Cheques.Proveedor = Proveedores.Proveedor "
  If pOrd = False Then
     sSql = sSql & "WHERE Cheques.Fecha>='" & FechaAmericana(pDesde) & "' And Cheques.Fecha<='" & FechaAmericana(pHasta) & "' AND  Cheques.Anulado=0 and Bancos.Descripcion<>'' and Cheques.Suc=" & nSucursal & " "
     If pCartera = True Then
        sSql = sSql & "AND Cheques.FechaSalida is NULL "
     End If
     sSql = sSql & "ORDER BY Cheques.Fecha,Cheques.Id"
  Else
     sSql = sSql & "WHERE Cheques.FechaAcre>='" & FechaAmericana(pDesde) & "' And Cheques.FechaAcre<='" & FechaAmericana(pHasta) & "' AND  Cheques.Anulado=0 and Bancos.Descripcion<>'' and Cheques.Suc=" & nSucursal & " "
     If pCartera = True Then
        sSql = sSql & "AND Cheques.FechaSalida is NULL "
     End If
     sSql = sSql & "ORDER BY Cheques.FechaAcre,Cheques.Id"
  End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerChequesTerceros = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TieneAplicacion(pCompr As Integer, pNumero As Long, pDat As String) As Currency
  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
   db.Open sDb

  sSql = "SELECT SUM(Aplicaciones.Importe) AS SumaDeImporte, CuentaCorriente.Cliente "
  sSql = sSql & "FROM Aplicaciones LEFT OUTER JOIN "
  sSql = sSql & "CuentaCorriente ON Aplicaciones.MoviCta = CuentaCorriente.Venta "
  sSql = sSql & " GROUP BY Aplicaciones.Comprobante, Aplicaciones.Numero, CuentaCorriente.Cliente "
  sSql = sSql & "HAVING (Aplicaciones.Comprobante = " & pCompr & ") AND (Aplicaciones.Numero = " & pNumero & ") and CuentaCorriente.Cliente='" & pDat & "'"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  If rs.RecordCount <> 0 Then
     TieneAplicacion = rs!SumaDeImporte
  End If
  db.Close
  
End Function

Public Function TraerTodosCon(pTabla As String, pCampo As String, pCondi As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT " & pCampo & " FROM " & pTabla & " WHERE " & pCondi
  sSql = sSql & " ORDER BY " & pCampo
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerTodosCon = rs
  db.Close
  
  Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerRetenciones(pDesde As Date, pHasta As Date) As ADODB.Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  

  sSql = "SELECT Retenciones.Fecha, Retenciones.Tipo, Clientes.RazonSocial, Clientes.Cuit, TipoValor.Descripcion, Retenciones.Numero, Retenciones.Importe,Retenciones.Movimiento "
  sSql = sSql & "FROM (Retenciones LEFT JOIN TipoValor ON Retenciones.Tipo = TipoValor.Tipo) LEFT JOIN Clientes ON Retenciones.Cliente = Clientes.Cliente "
  sSql = sSql & "WHERE Retenciones.Fecha>='" & FechaAmericana(pDesde) & "' AND Retenciones.Fecha<='" & FechaAmericana(pHasta) & "' and Retenciones.Suc=" & nSucursal & " "
  sSql = sSql & "ORDER BY Retenciones.Fecha "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db1, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerRetenciones = rs
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerResPlanCta(pDesde As Date, pHasta As Date, pTodos As Byte) As ADODB.Recordset
  Dim rs As ADODB.Recordset, sSql As String
  
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT TipoGrupo.Descripcion AS Grupo, PlanDeCuenta.Id, PlanDeCuenta.Descripcion AS Cuenta, Sum(DetallesComprobantes.PrecioTotal) AS SumaDePrecioTotal, TipoMovimiento.Descripcion AS Movi "
  sSql = sSql & "FROM ((CabComprobantes LEFT JOIN ((DetallesComprobantes LEFT JOIN PlanDeCuenta ON DetallesComprobantes.Cuenta = PlanDeCuenta.Id) LEFT JOIN TipoGrupo ON PlanDeCuenta.Grupo = TipoGrupo.Id) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento "
  sSql = sSql & "Where CabComprobantes.FechaIva >= '" & FechaAmericana(pDesde) & "' AND  CabComprobantes.FechaIva <= '" & FechaAmericana(pHasta) & "' "
  If pTodos = 0 Then
    sSql = sSql & "AND ((Comprobantes.Iva)=1) "
  End If
  sSql = sSql & "GROUP BY TipoGrupo.Descripcion, PlanDeCuenta.Id, PlanDeCuenta.Descripcion, DetallesComprobantes.Cuenta, CabComprobantes.Anulado, TipoMovimiento.Descripcion "
  sSql = sSql & "HAVING (((Sum(DetallesComprobantes.PrecioTotal)) Is Not Null) AND ((DetallesComprobantes.Cuenta)<>0) AND ((CabComprobantes.Anulado)=0) ) "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerResPlanCta = rs
  db.Close
  
End Function

Public Function TraerResPlanCtaDet(pDesde As Date, pHasta As Date, pCuenta As Integer) As ADODB.Recordset
  Dim rs As ADODB.Recordset, sSql As String
  
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT TipoGrupo.Descripcion AS Grupo, CabComprobantes.FechaIva, CabComprobantes.Cliente, CabComprobantes.Sucursal, CabComprobantes.Numero, Comprobantes.Descripcion As DesCompr, PlanDeCuenta.Descripcion AS Cuenta, DetallesComprobantes.Producto, DetallesComprobantes.Cantidad, DetallesComprobantes.Descripcion as DesDet, DetallesComprobantes.PrecioTotal "
  sSql = sSql & "FROM ((CabComprobantes LEFT JOIN ((DetallesComprobantes LEFT JOIN PlanDeCuenta ON DetallesComprobantes.Cuenta = PlanDeCuenta.Id) LEFT JOIN TipoGrupo ON PlanDeCuenta.Grupo = TipoGrupo.Id) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN TipoMovimiento ON Comprobantes.TipoMovimiento = TipoMovimiento.TipoMovimiento "
  sSql = sSql & "Where CabComprobantes.Fecha >= '" & FechaAmericana(pDesde) & "' AND  CabComprobantes.Fecha <= '" & FechaAmericana(pHasta) & "' "
  sSql = sSql & " AND DetallesComprobantes.PrecioTotal Is Not Null AND DetallesComprobantes.Cuenta<>0 AND CabComprobantes.Anulado=0 "
  If pCuenta <> -1 Then
     sSql = sSql & " AND PlanDeCuenta.Id=" & pCuenta & " "
  End If
  sSql = sSql & "ORDER BY CabComprobantes.FechaIva"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerResPlanCtaDet = rs
  db.Close

End Function

Public Function TraerResPlanCtaCaja(pDesde As Date, pHasta As Date, pCuenta) As Currency
  Dim rs As ADODB.Recordset, sSql As String
  
  Set rs = New ADODB.Recordset
   
  db.Open sDb

  sSql = "SELECT PlanDeCuenta.Grupo, CabComprobantes.Lista as id, PlanDeCuenta.Descripcion as Cuenta, Sum(CabComprobantes.Neto) AS SumaDePrecioTotal "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN PlanDeCuenta ON CabComprobantes.Lista = PlanDeCuenta.Id "
  sSql = sSql & "Where (((CabComprobantes.Tipo) = 3) And ((CabComprobantes.Anulado) = 0)) and CabComprobantes.Lista=" & pCuenta & " "
  sSql = sSql & "and CabComprobantes.Fecha >= '" & FechaAmericana(pDesde) & "' AND  CabComprobantes.Fecha <= '" & FechaAmericana(pHasta) & "' "
  sSql = sSql & "GROUP BY PlanDeCuenta.Grupo, CabComprobantes.Lista, PlanDeCuenta.Descripcion; "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
   
   If rs.RecordCount <> 0 Then
      TraerResPlanCtaCaja = rs!SumaDePrecioTotal
   End If
   db.Close
   
End Function

Public Function TraerCantidadRemito(pProd As String, pMovi As Long, pId As Long) As Double
  Dim rs As ADODB.Recordset, sSql As String
  
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT DetallesComprobantes.Producto, Sum(DetallesComprobantes.Cantidad) AS SumaDeCantidad "
  sSql = sSql & "From DetallesComprobantes "
'  sSql = sSql & "Where DetallesComprobantes.MoviRemi =" & pMovi & " AND DetallesComprobantes.Producto='" & pProd & "' AND Cast(DetallesComprobantes.Leyenda as Int)=" & pId & " "
  sSql = sSql & "Where DetRemi=" & pId & " "  ' & "  or Cast(DetallesComprobantes.Leyenda as Int)=" & pId & " "
  sSql = sSql & "GROUP BY DetallesComprobantes.Producto; "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
   
   TraerCantidadRemito = -1
   If rs.RecordCount <> 0 Then
      TraerCantidadRemito = rs!SumaDeCantidad
   End If
   db.Close
   
End Function

Public Function TraerCantidadFactura(pProd As String, pMovi As Long, pSaldo As Byte) As Double
  Dim rs As ADODB.Recordset, sSql As String
  
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Cantidad FROM DetallesComprobantes "
  sSql = sSql & "Where DetallesComprobantes.Movimiento =" & pMovi & " AND DetallesComprobantes.Producto='" & pProd & "' "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
   If rs.RecordCount <> 0 Then
      TraerCantidadFactura = rs!Cantidad
      If TraerCantidadFactura <> 0 Then
         sSql = "SELECT  Sum(DetallesComprobantes.Cantidad) AS SumaDeCantidad "
         sSql = sSql & "From DetallesComprobantes "
         sSql = sSql & "Where DetallesComprobantes.MoviRemi =" & pMovi & " AND DetallesComprobantes.Producto='" & pProd & "' "
         rs.Close
         rs.CursorLocation = adUseClient
         rs.Open sSql, db, adOpenStatic, adLockOptimistic
         rs.ActiveConnection = Nothing
   
         If rs.RecordCount <> 0 Then
            If Not IsNull(rs!SumaDeCantidad) Then
               If pSaldo = 2 Then
                  TraerCantidadFactura = rs!SumaDeCantidad
               Else
                   TraerCantidadFactura = TraerCantidadFactura - rs!SumaDeCantidad
               End If
            Else
               If pSaldo = 2 Then
                  TraerCantidadFactura = 0
               End If
            End If
         End If
     
      End If
   End If
  

 
   db.Close
   
End Function

Public Function DameCuentaCorrienteHA(sCli As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT * FROM CuentaCorriente "
  sSql = sSql & "WHERE Cliente='" & sCli & "'  "
  sSql = sSql & "AND Anulado<> 1 and Haber<>0 and Cuota=0 "
  sSql = sSql & "ORDER BY ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorrienteHA = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function DameCuentaCorrienteHAConFecha(sCli As String, pDesde As Date) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
   db.Open sDb
 
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT * "
  sSql = sSql & "From CuentaCorriente "
  sSql = sSql & "WHERE (((CuentaCorriente.Fecha)>='" & FechaAmericana(pDesde) & "') AND ((CuentaCorriente.Cliente)='" & sCli & "') AND ((CuentaCorriente.Haber)<>0) AND ((CuentaCorriente.Cuota)=0)) "
  sSql = sSql & "ORDER BY CuentaCorriente.Id "
 
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorrienteHAConFecha = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function DamePagosA(pVenta As Long) As Currency
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT SUM(Aplicaciones.Importe) AS total "
  sSql = sSql & "FROM Aplicaciones "
  sSql = sSql & "WHERE Aplicaciones.MoviCta=" & pVenta & " AND Aplicaciones.Anulado<>1"
          
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  If rs.EOF = False Or rs.BOF = False Then
     If Not IsNull(rs!Total) Then
        DamePagosA = rs!Total
     End If
  End If
  db.Close
  
End Function

Public Function DameCuentaCorrienteDA(sCli As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT * FROM CuentaCorriente "
  sSql = sSql & "WHERE Cliente='" & sCli & "'  "
  sSql = sSql & "AND Anulado<> 1 and Debe<>0 and Cuota=0 "
  sSql = sSql & "ORDER BY ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorrienteDA = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function MovimientosDeCaja(pFecha As Date, pCaja As Integer, pTurno As Byte) As Recordset
 On Error GoTo errHandler
  EnsureConn
  
  Dim sSql As String
  Dim rs As ADODB.Recordset
    
  Set rs = New ADODB.Recordset
   
   If pCaja <> 0 Then
      sSql = "SELECT CabComprobantes.Movimiento,CabComprobantes.Tipo,CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Debe, CabComprobantes.Haber "
      sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
      sSql = sSql & "Where CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.Caja=" & pCaja & " and CabComprobantes.Factura=0  and ((CabComprobantes.Tipo = 1 And CabComprobantes.CondVenta = 1) or (CabComprobantes.Tipo = 4 And CabComprobantes.CondVenta = 0) or (CabComprobantes.Tipo = 5 And CabComprobantes.CondVenta = 1)) "
      sSql = sSql & "GROUP BY CabComprobantes.Movimiento,CabComprobantes.Tipo, CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.CondVenta, CabComprobantes.Debe, CabComprobantes.Haber "
      If pTurno = 0 Then
         sSql = sSql & "HAVING (((CabComprobantes.Fecha)='" & FechaAmericana(pFecha) & "')) "
      Else
         sSql = sSql & "HAVING (((CabComprobantes.Fecha)>='" & FechaAmericana(pFecha) & "')) "
      End If
      sSql = sSql & "ORDER BY CabComprobantes.Tipo, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero "
   Else
      sSql = "SELECT CabComprobantes.Movimiento,CabComprobantes.Tipo,CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Debe, CabComprobantes.Haber "
      sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
      sSql = sSql & "Where CabComprobantes.Suc=" & nSucursal & " and CabComprobantes.Factura=0  and ((CabComprobantes.Tipo = 1 And CabComprobantes.CondVenta = 1) or (CabComprobantes.Tipo = 4 And CabComprobantes.CondVenta = 0) or (CabComprobantes.Tipo = 5 And CabComprobantes.CondVenta = 1)) "
      sSql = sSql & "GROUP BY CabComprobantes.Movimiento,CabComprobantes.Tipo, CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.CondVenta, CabComprobantes.Debe, CabComprobantes.Haber "
       If pTurno = 0 Then
         sSql = sSql & "HAVING (((CabComprobantes.Fecha)='" & FechaAmericana(pFecha) & "')) "
      Else
         sSql = sSql & "HAVING (((CabComprobantes.Fecha)>='" & FechaAmericana(pFecha) & "')) "
      End If
      sSql = sSql & "ORDER BY CabComprobantes.Tipo, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero "
   End If


   rs.CursorLocation = adUseClient
   rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
   rs.ActiveConnection = Nothing
  
  Set MovimientosDeCaja = rs
    
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function IngresoEgresoCaja(pFecha As Date, pCaja As Integer) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Comprobantes.Comprobante,Caja.Movimiento ,Caja.Sucursal, Caja.Numero, Caja.Debe, Caja.Haber "
  sSql = sSql & "FROM Caja LEFT JOIN Comprobantes ON Caja.Comprobante = Comprobantes.Id "
  If pCaja = 0 Then
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND   Comprobantes.TipoOperacion=3 AND Caja.Abierta='S' "
  Else
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND Caja.Caja=" & pCaja & " AND Caja.Abierta='S' AND (Comprobantes.TipoOperacion=3)"
  End If
  sSql = sSql & "ORDER BY Caja.Id "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set IngresoEgresoCaja = rs
  db.Close
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function IngresoCajaRes(pFecha As Date, pCaja As Integer) As Recordset
 On Error GoTo errHandler
  EnsureConn
  
  Dim sSql As String
  Dim rs As ADODB.Recordset
    
  Set rs = New ADODB.Recordset

  sSql = "SELECT Comprobantes.Comprobante, Caja.Hora ,Caja.Sucursal,Caja.Movimiento ,Caja.Numero, Caja.Debe, Caja.Haber "
  sSql = sSql & "FROM Caja LEFT JOIN Comprobantes ON Caja.Comprobante = Comprobantes.Id "
  If pCaja = 0 Then
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND   Comprobantes.TipoOperacion=3 AND Caja.Abierta='S' AND Comprobantes.TipoMovimiento=8 "
  Else
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND Caja.Caja=" & pCaja & " AND Caja.Abierta='S' AND (Comprobantes.TipoOperacion=3) AND Comprobantes.TipoMovimiento=8 "
  End If
  sSql = sSql & "ORDER BY Caja.Id "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set IngresoCajaRes = rs

Exit Function
errHandler:
  
  err.Raise err.Number
  
End Function

Public Function EgresoCajaRes(pFecha As Date, pCaja As Integer) As Recordset
  On Error GoTo errHandler
  EnsureConn
  
  Dim sSql As String
  Dim rs As ADODB.Recordset
    
  Set rs = New ADODB.Recordset

  sSql = "SELECT Comprobantes.Comprobante,Caja.Movimiento ,Caja.Sucursal, Caja.Numero, Caja.Debe, Caja.Haber "
  sSql = sSql & "FROM Caja LEFT JOIN Comprobantes ON Caja.Comprobante = Comprobantes.Id "
  If pCaja = 0 Then
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND   Comprobantes.TipoOperacion=3 AND Caja.Abierta='S' AND Comprobantes.TipoMovimiento=9 "
  Else
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND Caja.Caja=" & pCaja & " AND Caja.Abierta='S' AND (Comprobantes.TipoOperacion=3) AND Comprobantes.TipoMovimiento=9 "
  End If
  sSql = sSql & "ORDER BY Caja.Id "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set EgresoCajaRes = rs
  
Exit Function
errHandler:
  
  err.Raise err.Number
  
End Function

Public Function EgresoCajaPagos(pFecha As Date, pCaja As Integer) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Comprobantes.Comprobante, Caja.Nombre,Caja.Sucursal, Caja.Numero, Caja.Debe, Caja.Haber "
  sSql = sSql & "FROM Caja LEFT JOIN Comprobantes ON Caja.Comprobante = Comprobantes.Id "
  If pCaja = 0 Then
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND   Comprobantes.TipoOperacion=9 AND Caja.Abierta='S' "
  Else
     sSql = sSql & "WHERE Caja.Fecha='" & FechaAmericana(pFecha) & "' and Caja.Suc=" & nSucursal & " AND Caja.Caja=" & pCaja & " AND Caja.Abierta='S' AND (Comprobantes.TipoOperacion=9) "
  End If
  sSql = sSql & "ORDER BY Caja.Id "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set EgresoCajaPagos = rs
  db.Close
Exit Function
errHandler:
  db.Close
  err.Raise err.Number
  
End Function



Public Function TraerIngresosEgresos(pFecha As Date, pCaja As Integer) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Debe, CabComprobantes.Haber, CabComprobantes.Motivo "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "Where CabComprobantes.Tipo = 3 and CabComprobantes.Caja=" & pCaja & " "
  sSql = sSql & "GROUP BY CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Debe, CabComprobantes.Haber, CabComprobantes.Motivo "
  sSql = sSql & "HAVING (((CabComprobantes.Fecha)='" & FechaAmericana(pFecha) & "'));"


  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerIngresosEgresos = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerChequeTarjetas(pFecha As Date, pCaja As Integer, pTurno As Byte) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  

'  sSql = "SELECT Valores.*, Bancos.Descripcion "
'  sSql = sSql & "FROM (CabComprobantes RIGHT JOIN Valores ON CabComprobantes.Movimiento = Valores.Venta) LEFT JOIN Bancos ON Valores.Banco = Bancos.Banco "
  sSql = "SELECT Valores.Id, Valores.Venta, Valores.CodPago, Valores.FormaPago, Valores.Nombre, Valores.Banco, Valores.IdCheque, Valores.NumeroCheque, Valores.FechaAcreditacion, Valores.NumeroTarjeta, Valores.FechaVencimiento,"
  sSql = sSql & "Valores.Autorizacion, Valores.Cupon, Valores.Importe, Valores.PorVenta, Bancos.Descripcion, CabComprobantes.Numero, CabComprobantes.Sucursal, Comprobantes.Descripcion AS Comprobante "
  sSql = sSql & "FROM CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id RIGHT OUTER JOIN "
  sSql = sSql & "Valores ON CabComprobantes.Movimiento = Valores.Venta LEFT OUTER JOIN "
  sSql = sSql & "Bancos ON Valores.Banco = Bancos.Banco "
  If pCaja <> 0 Then
     If pTurno = 0 Then
        sSql = sSql & "WHERE CabComprobantes.Caja=" & pCaja
        sSql = sSql & " AND Valores.CodPago<>1 AND CabComprobantes.Fecha='" & FechaAmericana(pFecha) & "' and CabComprobantes.Suc=" & nSucursal & " "
     Else
        sSql = sSql & "WHERE CabComprobantes.Caja=" & pCaja
        sSql = sSql & " AND Valores.CodPago<>1 AND CabComprobantes.Fecha>='" & FechaAmericana(pFecha) & "' and CabComprobantes.Suc=" & nSucursal & " "
     End If
  Else
       If pTurno = 0 Then
          sSql = sSql & " WHERE Valores.CodPago<>1 AND CabComprobantes.Fecha='" & FechaAmericana(pFecha) & "' and CabComprobantes.Suc=" & nSucursal & " "
       Else
          sSql = sSql & " WHERE Valores.CodPago<>1 AND CabComprobantes.Fecha>='" & FechaAmericana(pFecha) & "' and CabComprobantes.Suc=" & nSucursal & " "
       End If
  End If
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerChequeTarjetas = rs
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerCobranzaLista(pDesde As Date, pHasta As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT CabComprobantes.Movimiento, CabComprobantes.Fecha, CabComprobantes.Numero, CabComprobantes.Sucursal, CabComprobantes.Cliente "
  sSql = sSql & "From CabComprobantes "
  sSql = sSql & "WHERE (((CabComprobantes.Tipo)=4) AND ((CabComprobantes.Fecha)>='" & FechaAmericana(pDesde) & "' And (CabComprobantes.Fecha)<='" & FechaAmericana(pHasta) & "') AND ((CabComprobantes.Anulado)=0)) and CabComprobantes.Suc=" & nSucursal & " "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerCobranzaLista = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerResCliSimpli(pDesde As String, pHasta As String) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  db.Open sDb

  sSql = "SELECT Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono, (Sum(CuentaCorriente.Debe) -Sum(CuentaCorriente.Haber)) as Saldo "
  sSql = sSql & "FROM Clientes LEFT JOIN CuentaCorriente ON Clientes.Cliente = CuentaCorriente.Cliente "
  sSql = sSql & "GROUP BY Clientes.Cliente, Clientes.RazonSocial, Clientes.Domicilio, Clientes.Telefono "
  sSql = sSql & "HAVING (((Clientes.Cliente)<>'.') AND ((Sum(CuentaCorriente.Debe)) Is Not Null)) "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerResCliSimpli = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TraerResProvSimpli(pDesde As String, pHasta As String) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.Telefono, (Sum(CuentaCorriente.Debe) -Sum(CuentaCorriente.Haber)) as Saldo "
  sSql = sSql & "FROM Proveedores LEFT JOIN CuentaCorriente ON Proveedores.Proveedor = CuentaCorriente.Cliente "
  sSql = sSql & "GROUP BY Proveedores.Proveedor, Proveedores.RazonSocial, Proveedores.Domicilio, Proveedores.Telefono "
  sSql = sSql & "HAVING Sum(CuentaCorriente.Debe Is Not Null "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerResProvSimpli = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function TraerNetoGrabado(pDesde As Date, pHasta As Date, pTipo As Byte) As Recordset
   Dim sSql As String, rs As ADODB.Recordset
   On Error GoTo errHandler
   Set rs = New ADODB.Recordset

   db.Open sDb

   sSql = "SELECT (Sum(DetallesComprobantes.PrecioTotal)-Sum(DetallesComprobantes.Impuesto)) AS NetoTotal "
   sSql = sSql & "FROM CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Impuestos ON DetallesComprobantes.Tasa = Impuestos.Impuesto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
   sSql = sSql & "Where CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' And CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "' "
   sSql = sSql & "GROUP BY CabComprobantes.Tipo "
   sSql = sSql & "HAVING (((CabComprobantes.Tipo)=" & pTipo & ")) "
   
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerNetoGrabado = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
   
End Function

Public Function TraerResumenIvaC(pDesde As Date, pHasta As Date, pTipo As Byte) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

   db.Open sDb

   sSql = "SELECT Impuestos.Descripcion,"
   sSql = sSql & "case "
   sSql = sSql & "when CabComprobantes.Tipo=8 then "
   sSql = sSql & "Sum (DetallesComprobantes.Impuesto) "
   sSql = sSql & "Else "
   sSql = sSql & "-Sum(DetallesComprobantes.Impuesto) "
   sSql = sSql & "end as Total,"
   sSql = sSql & "case "
   sSql = sSql & "when CabComprobantes.Tipo=8 then "
   sSql = sSql & "Sum (DetallesComprobantes.PrecioTotal) "
   sSql = sSql & "Else "
   sSql = sSql & "-Sum(DetallesComprobantes.PrecioTotal) "
   sSql = sSql & "end As NetoTotal "
   sSql = sSql & "FROM (CabComprobantes LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN Impuestos ON DetallesComprobantes.Tasa = Impuestos.Impuesto "
   sSql = sSql & "Where CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' And CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "' and CabComprobantes.Suc=" & nSucursal & " "
   sSql = sSql & "GROUP BY CabComprobantes.Tipo, Impuestos.Descripcion "
   sSql = sSql & "HAVING CabComprobantes.Tipo=" & pTipo & " OR CabComprobantes.Tipo=12 "
   sSql = sSql & "ORDER BY Impuestos.Descripcion"
'
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerResumenIvaC = rs
  db.Close

End Function

Public Function TraerResumemIva(pDesde As Date, pHasta As Date, pTipo As Byte) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
   db.Open sDb
   
   sSql = "SELECT Impuestos.Descripcion,"
   sSql = sSql & "case "
   sSql = sSql & "when CabComprobantes.Tipo=1 then "
   sSql = sSql & "Sum (DetallesComprobantes.Impuesto) "
   sSql = sSql & "Else "
   sSql = sSql & "-Sum(DetallesComprobantes.Impuesto) "
   sSql = sSql & "end as Total,"
   sSql = sSql & "case "
   sSql = sSql & "when CabComprobantes.Tipo=1 then "
   sSql = sSql & "Sum (DetallesComprobantes.PrecioTotal) "
   sSql = sSql & "Else "
   sSql = sSql & "-Sum(DetallesComprobantes.PrecioTotal) "
   sSql = sSql & "end As NetoTotal "

   sSql = sSql & "FROM (CabComprobantes LEFT JOIN (DetallesComprobantes LEFT JOIN Impuestos ON DetallesComprobantes.Tasa = Impuestos.Impuesto) ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
   sSql = sSql & "Where CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' And CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "' And CabComprobantes.Suc=" & nSucursal & " AND Impuestos.Descripcion is not Null "
   sSql = sSql & "GROUP BY CabComprobantes.Tipo, Impuestos.Descripcion "
   sSql = sSql & "HAVING CabComprobantes.Tipo=" & pTipo & " OR CabComprobantes.Tipo=5 "
   sSql = sSql & "ORDER BY Impuestos.Descripcion"
      
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerResumemIva = rs
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function


Public Function TraerInPercIB(pDesde As Date, pHasta As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

   db.Open sDb
   
  sSql = "SELECT CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, Clientes.Cuit, Clientes.RazonSocial, CabComprobantes.Neto,"
  sSql = sSql & "case "
  sSql = sSql & "when CabComprobantes.Tipo=1 then "
  sSql = sSql & "CabComprobantes.PIvaCompras  "
  sSql = sSql & "Else "
  sSql = sSql & "-CabComprobantes.PIvaCompras  "
  sSql = sSql & "end as PIvaCompras,"
  sSql = sSql & "Comprobantes.Comprobante "
  sSql = sSql & "FROM  CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id LEFT OUTER JOIN "
  sSql = sSql & "Clientes ON CabComprobantes.Cliente = Clientes.Cliente "
  sSql = sSql & "WHERE ((CabComprobantes.Fecha)>='" & FechaAmericana(pDesde) & "' And (CabComprobantes.Fecha)<='" & FechaAmericana(pHasta) & "' AND CabComprobantes.PIvaCompras<>0 and (CabComprobantes.Tipo=1 or CabComprobantes.Tipo=5) ) "
  sSql = sSql & "ORDER BY CabComprobantes.Fecha"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerInPercIB = rs
  db.Close

End Function

Public Function TraerInfRetIB(pDesde As Date, pHasta As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

   db.Open sDb
   
    sSql = "SELECT CabRetencion.FechaRec,CabRetencion.Fecha,CabRetencion.Numero ,CabRetencion.OPago, Proveedores.Cuit, Proveedores.RazonSocial, CabRetencion.Base, CabRetencion.Alicuota, CabRetencion.Importe "
    sSql = sSql & "FROM CabRetencion LEFT JOIN Proveedores ON CabRetencion.Proveedor = Proveedores.Proveedor "
    sSql = sSql & "WHERE (((CabRetencion.FechaRec)>='" & FechaAmericana(pDesde) & "' And (CabRetencion.FechaRec)<='" & FechaAmericana(pHasta) & "')) "
    sSql = sSql & "ORDER BY CabRetencion.Numero,CabRetencion.FechaRec"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerInfRetIB = rs
  db.Close

End Function

Public Function TraerInfRetIBex(pDesde As Date, pHasta As Date) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

   db.Open sDb
   
   sSql = "SELECT CabRetencion.FechaRec, CabRetencion.Fecha, CabRetencion.OPago, Proveedores.Cuit, Proveedores.RazonSocial, CabRetencion.Base, CabRetencion.Alicuota,"
   sSql = sSql & "CabRetencion.Importe, CabComprobantes.Sucursal, CabComprobantes.Numero AS Numero "
   sSql = sSql & "FROM CabRetencion LEFT OUTER JOIN "
   sSql = sSql & "CabComprobantes ON CabRetencion.Movimiento = CabComprobantes.Movimiento LEFT OUTER JOIN "
   sSql = sSql & "Proveedores ON CabRetencion.Proveedor = Proveedores.Proveedor "
   sSql = sSql & "WHERE (((CabRetencion.FechaRec)>='" & FechaAmericana(pDesde) & "' And (CabRetencion.FechaRec)<='" & FechaAmericana(pHasta) & "')) "
   sSql = sSql & "ORDER BY CabRetencion.Numero,CabRetencion.FechaRec"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerInfRetIBex = rs
  db.Close

End Function

Public Function TraerUltimoSistema(pCampo As String) As Long
   Dim Conexion As ADODB.Connection, sSql As String
   Dim rs As ADODB.Recordset
 
   Dim Com As ADODB.Command
   Set Com = New ADODB.Command

   Set Conexion = New ADODB.Connection
   Set rs = New ADODB.Recordset
   
   Conexion.Open "FILE NAME=" & App.Path & "\Conex.udl"
  ' db.Open sDb
   Conexion.Execute "UPDATE Sistema SET " & pCampo & "=" & pCampo & "+1"

   sSql = "SELECT " & pCampo & " as Valor FROM Sistema "
   rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
   TraerUltimoSistema = rs!Valor
   Conexion.Close
End Function

Public Function TraerFacturasAplicadas(pDesde As Date, pHasta As Date, pCliente As String) As Recordset
  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT CuentaCorriente.Cliente, CuentaCorriente.Fecha AS FechaFactura, Comprobantes_1.Descripcion, CuentaCorriente.Numero,  CuentaCorriente.Debe,"
  sSql = sSql & "CuentaCorriente.Haber, CuentaCorriente_1.Fecha AS FechaPago, Comprobantes.Descripcion AS Cobranza,CuentaCorriente_1.Numero AS NumRec, Aplicaciones.Importe "
  sSql = sSql & "FROM  Comprobantes INNER JOIN "
  sSql = sSql & "CuentaCorriente AS CuentaCorriente_1 ON Comprobantes.Id = CuentaCorriente_1.Comprobante RIGHT OUTER JOIN "
  sSql = sSql & "CuentaCorriente LEFT OUTER JOIN "
  sSql = sSql & "Aplicaciones ON CuentaCorriente.Venta = Aplicaciones.Venta LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes AS Comprobantes_1 ON CuentaCorriente.Comprobante = Comprobantes_1.Id ON CuentaCorriente_1.Venta = Aplicaciones.MoviCta "
  sSql = sSql & "WHERE (CuentaCorriente.Cliente = '" & pCliente & "') AND (CuentaCorriente.Fecha >= '" & pDesde & "') AND (CuentaCorriente.Fecha <= '" & pHasta & "') "
  sSql = sSql & "AND (Comprobantes_1.TipoOperacion = 1) "
  sSql = sSql & "ORDER BY FechaFactura "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerFacturasAplicadas = rs
  db.Close
  
End Function

Public Function TraerSaldoAplica(pDesde As Date, pHasta As Date, pCliente As String) As Double
  Dim sSql As String, rs As ADODB.Recordset, nSaldo As Double
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT CuentaCorriente.Cliente, CuentaCorriente.Debe, SUM(Aplicaciones.Importe) AS Cancelado, Comprobantes.TipoOperacion "
  sSql = sSql & "FROM Aplicaciones RIGHT OUTER JOIN "
  sSql = sSql & "CuentaCorriente ON Aplicaciones.Venta = CuentaCorriente.Venta LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id "
  sSql = sSql & "GROUP BY CuentaCorriente.Fecha, CuentaCorriente.Cliente, CuentaCorriente.Debe, Comprobantes.TipoOperacion "
  sSql = sSql & "HAVING (CuentaCorriente.Fecha >='" & pDesde & "') AND (CuentaCorriente.Fecha <='" & pHasta & "') "
  sSql = sSql & "AND (CuentaCorriente.Cliente ='" & pCliente & "') AND (Comprobantes.TipoOperacion = 1) "
  sSql = sSql & "ORDER BY CuentaCorriente.Fecha "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  nSaldo = 0
  If rs.RecordCount <> 0 Then
     Do While Not rs.EOF
        nSaldo = nSaldo + (rs!Debe - IIf(IsNull(rs!Cancelado), 0, rs!Cancelado))
        rs.MoveNext
     Loop
  End If
  
  TraerSaldoAplica = nSaldo
  db.Close
  
End Function

Public Function TraerCabezaCtiVenta(pDesde As Date, pHasta As Date) As Recordset
 Dim db As Connection
 Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
    
  sSql = "SELECT CabComprobantes.Fecha, CabComprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente, CabComprobantes.Neto, CabComprobantes.Iva1,"
  sSql = sSql & "CabComprobantes.Descuentos, CabComprobantes.NoGravados, CabComprobantes.PIvaCompras, CabComprobantes.IBrutosCompras, CabComprobantes.Debe, CabComprobantes.Haber,"
  sSql = sSql & "CabComprobantes.Movimiento, Comprobantes.Iva, CabComprobantes.Tipo, AfipCaes.Movimiento AS Expr1 "
  sSql = sSql & "FROM CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "AfipCaes ON CabComprobantes.Movimiento = AfipCaes.Movimiento LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "Where (CabComprobantes.Suc = 1) And (Comprobantes.Afip <> 0) And AfipCaes.Movimiento Is Not Null "
  sSql = sSql & "GROUP BY CabComprobantes.Fecha, CabComprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente, CabComprobantes.Neto, CabComprobantes.Iva1,"
  sSql = sSql & "CabComprobantes.Descuentos, CabComprobantes.NoGravados, CabComprobantes.PIvaCompras, CabComprobantes.IBrutosCompras, CabComprobantes.Debe, CabComprobantes.Haber,"
  sSql = sSql & "CabComprobantes.Movimiento , Comprobantes.Iva, CabComprobantes.tipo, AfipCaes.Movimiento "
  sSql = sSql & "HAVING (CabComprobantes.Fecha >= '" & pDesde & "') AND (CabComprobantes.Fecha <= '" & pHasta & "') AND (Comprobantes.Iva = 1) AND ((CabComprobantes.Tipo = 1) OR  "
  sSql = sSql & "(CabComprobantes.Tipo = 5) OR (CabComprobantes.Tipo = 13)) "
  sSql = sSql & "ORDER BY CabComprobantes.Fecha, CabComprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerCabezaCtiVenta = rs
  db.Close
  
End Function

Public Function TraerCabezaCtiCompra(pDesde As Date, pHasta As Date) As Recordset
 Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT CabComprobantes.FechaIva, CabComprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente,"
  sSql = sSql & "CabComprobantes.Neto,CabComprobantes.Iva1, CabComprobantes.NoGravados, CabComprobantes.PIvaCompras, CabComprobantes.IBrutosCompras, CabComprobantes.Debe,"
  sSql = sSql & "CabComprobantes.Haber , CabComprobantes.Movimiento, Comprobantes.Iva, CabComprobantes.tipo "
  sSql = sSql & "FROM CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "Where (CabComprobantes.Suc = 1) AND Comprobantes.Afip<>0 "
  sSql = sSql & "GROUP BY CabComprobantes.FechaIva, CabComprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Cliente,"
  sSql = sSql & "CabComprobantes.Neto,CabComprobantes.Iva1, CabComprobantes.NoGravados, CabComprobantes.PIvaCompras, CabComprobantes.IBrutosCompras, CabComprobantes.Debe,"
  sSql = sSql & "CabComprobantes.Haber , CabComprobantes.Movimiento, Comprobantes.Iva, CabComprobantes.tipo "
  sSql = sSql & "HAVING (CabComprobantes.FechaIva >= '" & pDesde & "') AND (CabComprobantes.FechaIva <= '" & pHasta & "') AND (Comprobantes.Iva = 1) AND ((CabComprobantes.Tipo = 8) OR  "
  sSql = sSql & "(CabComprobantes.Tipo = 12)) "
  sSql = sSql & "ORDER BY CabComprobantes.FechaIva,CabComprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerCabezaCtiCompra = rs
  db.Close
  
End Function

Public Function TraerDetalleCitiVenta(pMovimiento As Long) As Recordset
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT DetallesComprobantes.Movimiento, DetallesComprobantes.Tasa, SUM(DetallesComprobantes.Impuesto) AS SumaDeImpuesto,"
  sSql = sSql & "(SUM(DetallesComprobantes.PrecioTotal)-SUM(DetallesComprobantes.Descuento)) AS SumaDePrecioTotal, Impuestos.CodAfip "
  sSql = sSql & "FROM DetallesComprobantes LEFT OUTER JOIN "
  sSql = sSql & "Impuestos ON DetallesComprobantes.Tasa = Impuestos.Impuesto "
  sSql = sSql & "WHERE DetallesComprobantes.Movimiento=" & pMovimiento & " and DetallesComprobantes.Tasa<>21 "
  sSql = sSql & " GROUP BY DetallesComprobantes.Movimiento, DetallesComprobantes.Tasa, Impuestos.CodAfip "
   
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerDetalleCitiVenta = rs
  db.Close
     
     
End Function

Public Function TraerComprobantesProyeccion(pDesde As String, pHasta As String, pCliente As String) As Double
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
  sSql = "SELECT SUM(Debe - dbo.traeraplicacion(Venta,Cuota)) as Importe "
  sSql = sSql & "From CuentaCorriente "
  sSql = sSql & "WHERE (Debe <> 0) AND Fecha >'" & pDesde & "'  and Fecha <='" & pHasta & "' " ' (Debe <> 0) AND
  sSql = sSql & "AND Cliente='" & pCliente & "'"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  If rs.RecordCount <> 0 Then
     TraerComprobantesProyeccion = IIf(IsNull(rs!importe), 0, rs!importe)
  End If
  db.Close
   
End Function

Public Function TraerFacturasOnLine(pUser As Integer) As Recordset
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Comprobantes.Id, Comprobantes.Descripcion,Comprobantes.Afip "
  sSql = sSql & "FROM Comprobantes LEFT OUTER JOIN "
  sSql = sSql & "UsuariosComprobantesVentas ON Comprobantes.Id = UsuariosComprobantesVentas.Comprobante "
  sSql = sSql & "Where (UsuariosComprobantesVentas.Usuario = " & pUser & ") And (UsuariosComprobantesVentas.Valor = 1) And (Comprobantes.Afip <> 0) "
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerFacturasOnLine = rs
  db.Close
  
End Function

Public Function TraerMoviCae(pCompro As Integer, pNumero As Long, pSuc As Integer, pTipo As Byte) As Long
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb

  If pTipo = 0 Then
     sSql = "SELECT AfipCaes.Movimiento "
     sSql = sSql & "FROM CabComprobantes LEFT OUTER JOIN "
     sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id LEFT OUTER JOIN "
     sSql = sSql & "AfipCaes ON CabComprobantes.Movimiento = AfipCaes.Movimiento "
     sSql = sSql & " Where CabComprobantes.Numero = " & pNumero & " And CabComprobantes.Sucursal =" & pSuc & " "
     sSql = sSql & "AND Comprobantes.Afip=" & pCompro
  Else
    sSql = "SELECT CabComprobantes.Movimiento "
    sSql = sSql & "FROM  CabComprobantes LEFT OUTER JOIN "
    sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
    sSql = sSql & "Where (CabComprobantes.Numero = " & pNumero & ") And (CabComprobantes.Sucursal =" & pSuc & ") And (Comprobantes.Afip =" & pCompro & " )"
  End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  If rs.RecordCount <> 0 Then
     TraerMoviCae = IIf(IsNull(rs!Movimiento), 0, rs!Movimiento)
  End If
  db.Close
  
  
End Function

Public Function TraerPercExp(pDesde As String, pHasta As String) As Recordset
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Comprobantes.Afip,Comprobantes.Descripcion, Clientes.Cuit, CabComprobantes.Fecha, CabComprobantes.Sucursal, CabComprobantes.Numero, CabComprobantes.Neto,"
  sSql = sSql & "CabComprobantes.PIvaCompras "
  sSql = sSql & "FROM CabComprobantes LEFT OUTER JOIN "
  sSql = sSql & "Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id LEFT OUTER JOIN "
  sSql = sSql & "Clientes ON CabComprobantes.Cliente = Clientes.Cliente "
  sSql = sSql & "WHERE (Comprobantes.TipoOperacion = 1 or Comprobantes.TipoOperacion = 5)  and PIvaCompras<>0 and fecha>='" & pDesde & "' and Fecha<='" & pHasta & "'  AND (Comprobantes.Iva = 1) "
  sSql = sSql & "ORDER BY Comprobantes.Descripcion,CabComprobantes.Sucursal,CabComprobantes.Numero"

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerPercExp = rs
  db.Close

End Function

Public Function TraerMoviDesp(pDesde As Integer, pHasta As Long) As Long
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT top 1 Cantidades.Movimiento, Cantidades_1.Movimiento AS Expr1, Cantidades_1.Estado, Cantidades_1.Deposito "
  sSql = sSql & "FROM Cantidades LEFT OUTER JOIN "
  sSql = sSql & "Cantidades AS Cantidades_1 ON Cantidades.Movimiento = Cantidades_1.Movimiento "
  sSql = sSql & "Where ((Cantidades.Estado = 2) And (Cantidades.Deposito =" & pDesde & ")) And (Cantidades_1.Deposito =" & pHasta & ")"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  TraerMoviDesp = rs.RecordCount
  
  db.Close
End Function

Public Function TraerRepartoRemi(pDesde As String, pHasta As String, pEstado As Byte, pZona As String) As Recordset
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT RegTransporte.Fecha, RegTransporte.Hora, RegTransporte.Destinatario, RegTransporte.Direccion,RegTransporte.Telefono,RegTransporte.Obsrevaciones, RegTransporte.Estado, CabComprobantes.Numero,CabComprobantes.Sucursal,"
  sSql = sSql & "DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, DetallesComprobantes.Cantidad,CabComprobantes.Vendedor ,CabComprobantes.Motivo, Zonas.Descripcion AS Zona "
  sSql = sSql & "FROM Zonas INNER JOIN "
  sSql = sSql & "Clientes ON Zonas.Zona = Clientes.Zona RIGHT OUTER JOIN "
  sSql = sSql & "RegTransporte INNER JOIN "
  sSql = sSql & "CabComprobantes ON RegTransporte.Movimiento = CabComprobantes.Movimiento INNER JOIN "
  sSql = sSql & "DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento ON Clientes.Cliente = CabComprobantes.Cliente "
  sSql = sSql & "WHERE (RegTransporte.Estado = " & pEstado & ") AND RegTransporte.Fecha>='" & pDesde & "' AND RegTransporte.Fecha<='" & pHasta & "' "
  If pZona <> "TODAS" Then
     sSql = sSql & "AND Zonas.Descripcion='" & pZona & "' "
  End If
  sSql = sSql & "ORDER BY RegTransporte.Destinatario"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerRepartoRemi = rs
  
  db.Close
End Function

Public Function TraerRemitoEntrega(pSucursal As Long, pNumero As Long) As Recordset
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
  
  sSql = "SELECT  CabComprobantes.Movimiento,CabComprobantes.Fecha, CabComprobantes.Numero, CabComprobantes.Sucursal, RegTransporte.Destinatario, CabComprobantes.Motivo "
  sSql = sSql & "FROM CabComprobantes INNER JOIN "
  sSql = sSql & "RegTransporte ON CabComprobantes.Movimiento = RegTransporte.Movimiento "
  sSql = sSql & "WHERE RegTransporte.Estado = 2 And CabComprobantes.Numero =" & pNumero & " And CabComprobantes.Sucursal =" & pSucursal
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerRemitoEntrega = rs
  
  db.Close
  
End Function

Public Function TraerVentas(pDesde As Date, pHasta As Date) As Recordset
 Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT Count(Comprobantes.Descripcion) as Cantidad ,Comprobantes.Descripcion, Sum(CabComprobantes.Debe)- Sum(CabComprobantes.Haber) AS SumaDePTotal "
  sSql = sSql & "FROM CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "Where  CabComprobantes.Anulado=0 and CabComprobantes.Tipo=1 "
  sSql = sSql & "AND CabComprobantes.Fecha >='" & FechaAmericana(pDesde) & "' and CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "' "
  sSql = sSql & "GROUP BY Comprobantes.Descripcion; "

  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerVentas = rs
  db.Close

End Function

Public Function TraerChequesPropios() As Recordset
  Dim db As Connection
  Set db = New Connection

  Dim sSql As String, rs As ADODB.Recordset
  Set rs = New ADODB.Recordset

  db.Open sDb
  sSql = "SELECT CONVERT(char, CtaBanco.Fecha,103) as FechaEmision, CtaBanco.FormaPago, CtaBanco.Nombre, CtaBanco.NumeroCheque, convert(char,CtaBanco.FechaAcreditacion,103) as FechaAcreditacion, CtaBanco.Haber,"
  sSql = sSql & "CuentasBancarias.Descripcion "
  sSql = sSql & "FROM CtaBanco LEFT OUTER JOIN "
  sSql = sSql & "CuentasBancarias ON CtaBanco.Banco = CuentasBancarias.Id "
  sSql = sSql & "WHERE (CtaBanco.FormaPago = 'CHEQUES PROPIOS') and Fecha>='01/06/2015' and Fecha<='31/07/2016' "
  sSql = sSql & "ORDER BY CtaBanco.Fecha "
   
rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerChequesPropios = rs
  db.Close
  
End Function

Public Function TraerIvaPorFamilia(pFamilia As Integer, pTasa As Boolean, pDesde As String, pHasta As String) As Recordset
  Dim rs As ADODB.Recordset, sSql As String
  
  Set rs = New ADODB.Recordset

  db.Open sDb
  sSql = "SELECT  Familias.Descripcion, "
  sSql = sSql & "CASE WHEN CabComprobantes.Tipo=1 THEN Sum(DetallesComprobantes.PrecioTotal) "
  sSql = sSql & "ELSE -Sum(DetallesComprobantes.PrecioTotal) END  AS SumaDePrecioTotal "
  sSql = sSql & "FROM Impuestos RIGHT OUTER JOIN "
  sSql = sSql & "DetallesComprobantes INNER JOIN "
  sSql = sSql & "CabComprobantes ON DetallesComprobantes.Movimiento = CabComprobantes.Movimiento ON Impuestos.Impuesto = DetallesComprobantes.Tasa LEFT OUTER JOIN "
  sSql = sSql & "Familias RIGHT OUTER JOIN "
  sSql = sSql & "Productos ON Familias.Familia = Productos.Familia ON DetallesComprobantes.Producto = Productos.Producto "
  sSql = sSql & "Where (Productos.Familia = " & pFamilia & "  And (CabComprobantes.tipo = 1 Or CabComprobantes.tipo = 5) "
  If pTasa = False Then
     sSql = sSql & "And DetallesComprobantes.Impuesto = 0) "
  Else
   sSql = sSql & "And DetallesComprobantes.Impuesto <> 0) "
  End If
  
  sSql = sSql & "AND CabComprobantes.Fecha >= '" & pDesde & "' AND  CabComprobantes.Fecha<='" & pHasta & "' "
  sSql = sSql & "GROUP BY Familias.Descripcion, CabComprobantes.Tipo "
  
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.Close
  rs.CursorType = adOpenStatic
  rs.CursorLocation = adUseClient
  rs.LockType = adLockBatchOptimistic
  rs.Open
  rs.ActiveConnection = Nothing

  Set TraerIvaPorFamilia = rs
  db.Close
End Function


