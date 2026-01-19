VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsComprobantesE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Version(pCons As String)
  On Error GoTo errHandler

  Dim sConex As Connection, sCons(0) As String, Rx As Recordset, cRsl As ClsLectura
  Dim nVersion As Integer
  
  Set sConex = New Connection
  Set cRsl = New ClsLectura
  
  Set Rx = cRsl.TraerTodos("VersionSQL", "*")
  
  nVersion = cRsl.TraerValorDeUnCampo("Sistema", "VersionPrinc", "")
  
  sConex.Open sDb
  sConex.BeginTrans
  Do While Not Rx.EOF
     If Rx!Version > nVersion Then
        sConex.Execute Replace(Rx!Instruccion, "''", "'")
     End If
     Rx.MoveNext
  Loop
  
  sConex.CommitTrans
  sConex.Close
Exit Sub

errHandler:
    nVersion = 2
    Resume Next
End Sub

Public Sub Agregar(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, _
  pRsRecibo As Recordset, pRsValor As Recordset, pRsAplica As Recordset, pRsCantidad As Recordset, pRsCaja As Recordset, Optional bRem As Boolean, Optional rRem As Recordset)
  Dim iCampos As Integer, sProv As String
  Dim Com As ADODB.Command
  Dim i As Integer
  Dim cRsl As ClsLectura
  
  Dim sConex As Connection
  Dim cRse As ClsEscritura
  
  Set sConex = New Connection
  Set cRse = New ClsEscritura
  
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  sConex.Open sDb
  sConex.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = sConex
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Name = "Cliente" Then sProv = pRsVenta(i).Value
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next i
  Com.Execute
  Set Com = Nothing
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13)
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("MoviRemi", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = IIf(pRsDetalle!Descripcion = "", " ", pRsDetalle!Descripcion)
        Com.Parameters("Leyenda").Value = IIf(pRsDetalle!Leyenda = "", " ", pRsDetalle!Leyenda)
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
        Com.Parameters("PDesc").Value = IIf(IsNull(pRsDetalle!Pdesc), 0, pRsDetalle!Pdesc)
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = pRsDetalle!Impuesto
        Com.Parameters("Precio").Value = IIf(IsNull(pRsDetalle!precio), pRsDetalle!PrecioUnitario, pRsDetalle!precio)
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = pRsDetalle!PrecioTotal
        Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
        Com.Parameters("DetRemi").Value = IIf(IsNull(pRsDetalle!DetRemi), 0, pRsDetalle!DetRemi)
        Com.Execute
        
        If pRsVenta!tipo = 8 Then
           Dim nTotal As Double
           
           If pRsDetalle!MoviRemi = 1 Then
              nTotal = pRsDetalle!PrecioUnitario
              sConex.Execute "UPDATE Productos SET PrecioCompra=" & nTotal & ",FechaCompra='" & FechaAmericana(Date) & "',Actualiza=1  WHERE Producto='" & pRsDetalle!Producto & "'"
           Else
              nTotal = pRsDetalle!PrecioUnitario
              sConex.Execute "UPDATE Productos SET PrecioCompra=" & nTotal & ",FechaCompra='" & FechaAmericana(Date) & "',Actualiza=0  WHERE Producto='" & pRsDetalle!Producto & "'"
           End If
        End If
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  
  ' Actualizo el Ultimo Recibo
  If pRsCaja.RecordCount <> 0 Then
     Set Com = New ADODB.Command
     Com.ActiveConnection = sConex
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "ComprobanteActualizarNumero"
     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsCaja!Comprobante)
     Com.Execute
     Set Com = Nothing
  End If
  
  ' Finaliza la Operacion de Venta
  
  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
  If pRsRecibo.RecordCount <> 0 Then
     pRsRecibo.MoveFirst
     Do While Not pRsRecibo.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
    
       iCampos = pRsRecibo.Fields.Count - 1
       For i = 0 To iCampos
           If pRsRecibo(i).Name <> "Id" Then
              If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              Else
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              End If
           End If
       Next
       Com.Execute
       pRsRecibo.MoveNext
    Loop
  End If
  Set Com = Nothing
  If pRsValor.RecordCount <> 0 Then
  pRsValor.MoveFirst
  Do While Not pRsValor.EOF
     Set Com = New ADODB.Command
     Com.ActiveConnection = sConex
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "ValorAgregar"
     Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput, , pRsValor!Venta)
     Com.Parameters.Append Com.CreateParameter("CodPago", adInteger, adParamInput, , pRsValor!CodPago)
     Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsValor!formapago)
     Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre)
     Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , IIf(IsNull(pRsValor!Banco), 0, pRsValor!Banco))
     Com.Parameters.Append Com.CreateParameter("IdCheque", adInteger, adParamInput, , IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque))
     Com.Parameters.Append Com.CreateParameter("NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
     Com.Parameters.Append Com.CreateParameter("FAcredi", adVarChar, adParamInput, 10, pRsValor!FechaAcreditacion)
     Com.Parameters.Append Com.CreateParameter("NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta)
     Com.Parameters.Append Com.CreateParameter("FVencim", adVarChar, adParamInput, 10, pRsValor!fechavencimiento)
     Com.Parameters.Append Com.CreateParameter("Autoriz", adVarChar, adParamInput, 30, "")
     Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, "")
     Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
     Com.Parameters.Append Com.CreateParameter("PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta)
     Com.Execute
     pRsValor.MoveNext
  Loop
  End If
  Set Com = Nothing
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Venta
  If bRem = False Then
     If pRsCantidad.RecordCount <> 0 Then
        pRsCantidad.MoveFirst
        Do While Not pRsCantidad.EOF
           If pRsCantidad!Producto <> "" Then
              Set Com = New ADODB.Command
              Com.ActiveConnection = sConex
              Com.CommandType = adCmdStoredProc
              Com.CommandText = "CantidadesAgregar"
              Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
              Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
              Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
              Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
              Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
              Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
              Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
              Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
              Com.Execute
           End If
           pRsCantidad.MoveNext
        Loop
        Set Com = Nothing
     End If
  Else
    If Not rRem Is Nothing Then
     rRem.MoveFirst
       Do While Not rRem.EOF
          Dim sSql As String
          sSql = "UPDATE CabComprobantes SET Factura=" & pRsVenta!Movimiento
          sSql = sSql & " WHERE Movimiento=" & rRem!Movimiento
          rRem.MoveNext
          sConex.Execute sSql
       Loop
     End If
  End If
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , IIf(IsNull(pRsCaja!fechavencimiento), Date, pRsCaja!fechavencimiento))
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  Set cRsl = New ClsLectura
  
  If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & pRsVenta!Comprobante & " AND Usuario=" & nUsuario) = False Then
     Set Com = New ADODB.Command
     Com.ActiveConnection = sConex
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "ComprobanteActualizarNumero"
     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
     Com.Execute
  End If
  Set Com = Nothing

  Call cRse.GuardarEventoCompleto(sConex, "venta", "insert", pRsVenta, pRsDetalle, pRsCaja, pRsRecibo)

  sConex.CommitTrans
    
  sConex.Close
  Exit Sub
    
errHandler:
    sConex.RollbackTrans
    sConex.Close
    
    err.Raise err.Number
End Sub


Public Sub ModificarVenta(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, _
  pRsRecibo As Recordset, pRsValor As Recordset, pRsAplica As Recordset, pRsCantidad As Recordset, pRsCaja As Recordset, Optional bRem As Boolean, Optional rRem As Recordset, Optional Cod As Long)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
  
  Dim cRse As ClsEscritura
  Set cRse = New ClsEscritura
    
  db.Open sDb
  db.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaModif"
'
'  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
'
  If Cod <> 0 Then
     If VarType(Cod) Then
        Com.Parameters.Append Com.CreateParameter("Param", adInteger, adParamInput, , Cod)
     Else
        Com.Parameters.Append Com.CreateParameter("Param", adVarChar, adParamInput, Len(Cod), Cod)
     End If
  End If
  If Not pRsVenta Is Nothing Then
     iCampos = pRsVenta.Fields.Count - 1
     For i = 0 To iCampos
         If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
     Next
  End If
  Com.Execute
  Set Com = Nothing
  
   db.Execute "DELETE  FROM DetallesComprobantes WHERE Movimiento=" & pRsVenta!Movimiento
   db.Execute "DELETE  FROM CuentaCorriente WHERE Venta=" & pRsVenta!Movimiento
   db.Execute "DELETE  FROM Aplicaciones WHERE Venta=" & pRsVenta!Movimiento
   db.Execute "DELETE  FROM Cantidades WHERE Movimiento=" & pRsVenta!Movimiento
    
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13, "")
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("MoviRemi", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = IIf(pRsDetalle!Descripcion = "", " ", pRsDetalle!Descripcion)
        Com.Parameters("Leyenda").Value = IIf(pRsDetalle!Leyenda = "", " ", pRsDetalle!Leyenda)
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
        Com.Parameters("PDesc").Value = pRsDetalle!Pdesc
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = pRsDetalle!Impuesto
        Com.Parameters("Precio").Value = pRsDetalle!precio
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = pRsDetalle!PrecioTotal
        Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
        Com.Parameters("DetRemi").Value = IIf(IsNull(pRsDetalle!DetRemi), 0, pRsDetalle!DetRemi)
        Com.Execute
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  If pRsVenta!tipo = 8 Then
     If pRsDetalle.RecordCount > 0 Then
        pRsDetalle.MoveFirst
        Do While Not pRsDetalle.EOF
           db.Execute "UPDATE Productos SET PrecioCompra=" & pRsDetalle!PrecioUnitario & " WHERE Producto='" & pRsDetalle!Producto & "'"
           pRsDetalle.MoveNext
        Loop
     End If
  End If
  
 
  If pRsRecibo.RecordCount <> 0 Then
     pRsRecibo.MoveFirst
     Do While Not pRsRecibo.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
    
       iCampos = pRsRecibo.Fields.Count - 1
       For i = 0 To iCampos
           If pRsRecibo(i).Name <> "Id" Then
              If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              Else
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              End If
           End If
       Next
       Com.Execute
       pRsRecibo.MoveNext
    Loop
  End If
  Set Com = Nothing
  If pRsValor.RecordCount <> 0 Then
  pRsValor.MoveFirst
  Do While Not pRsValor.EOF
     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "ValorAgregar"
     Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput, , pRsValor!Venta)
     Com.Parameters.Append Com.CreateParameter("CodPago", adInteger, adParamInput, , pRsValor!CodPago)
     Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsValor!formapago)
     Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre)
     Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
     Com.Parameters.Append Com.CreateParameter("NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
     Com.Parameters.Append Com.CreateParameter("FAcredi", adDate, adParamInput, , pRsValor!FechaAcreditacion)
     Com.Parameters.Append Com.CreateParameter("NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta)
     Com.Parameters.Append Com.CreateParameter("FVencim", adDate, adParamInput, , pRsValor!fechavencimiento)
     Com.Parameters.Append Com.CreateParameter("Autoriz", adVarChar, adParamInput, 30, pRsValor!Autorizacion)
     Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsValor!Cupon)
     Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
     Com.Parameters.Append Com.CreateParameter("PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta)
     Com.Execute
     pRsValor.MoveNext
  Loop
  End If
  Set Com = Nothing
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Venta
  If bRem = False Then
     If pRsCantidad.RecordCount <> 0 Then
        pRsCantidad.MoveFirst
        Do While Not pRsCantidad.EOF
           If pRsCantidad!Producto <> "" Then
              Set Com = New ADODB.Command
              Com.ActiveConnection = db
              Com.CommandType = adCmdStoredProc
              Com.CommandText = "CantidadesAgregar"
              Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
              Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
              Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
              Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
              Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
              Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
              Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
              Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
              Com.Execute
           End If
           pRsCantidad.MoveNext
        Loop
        Set Com = Nothing
     End If
  End If
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  Call cRse.GuardarEventoCompleto(db, "venta", "insert", pRsVenta, pRsDetalle, pRsCaja, pRsRecibo)

  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Public Sub AgregarCobranzas(pRsVenta As Recordset, pRsRecibo As Recordset, pRsValor As Recordset, pRsAplica As Recordset, pRsCaja As Recordset, pRsBanco As Recordset, pRet As Recordset)
  Dim iCampos As Integer, ID As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans

  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
  If pRsRecibo.RecordCount <> 0 Then
  

     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
     iCampos = pRsVenta.Fields.Count - 1
     For i = 0 To iCampos
         If pRsVenta(i).Name <> "Id" Then
            If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
               Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
            Else
               Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
            End If
         End If
     Next i
     Com.Execute
     Set Com = Nothing
  
  
     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
     
     iCampos = pRsRecibo.Fields.Count - 1
     For i = 0 To iCampos
         If pRsRecibo(i).Name <> "Id" Then
            If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
               Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
            Else
               Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
            End If
         End If
     Next
     Com.Execute
     Set Com = Nothing
    
     pRsValor.MoveFirst
     Do While Not pRsValor.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "ValorAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput, , pRsValor!Venta)
        Com.Parameters.Append Com.CreateParameter("CodPago", adInteger, adParamInput, , pRsValor!CodPago)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsValor!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
        Com.Parameters.Append Com.CreateParameter("IdCheque", adInteger, adParamInput, , IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque))
        Com.Parameters.Append Com.CreateParameter("NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FAcredi", adDate, adParamInput, , pRsValor!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FVencim", adDate, adParamInput, , pRsValor!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autoriz", adVarChar, adParamInput, 30, pRsValor!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsValor!Cupon)
        Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
        Com.Parameters.Append Com.CreateParameter("PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta)
        Com.Execute
        pRsValor.MoveNext
     Loop
     Set Com = Nothing
     
     
     pRsValor.MoveFirst
     Do While Not pRsValor.EOF
        If pRsValor!IdCheque = 1 Then
           Set Com = New ADODB.Command
           Com.ActiveConnection = db
           Com.CommandType = adCmdStoredProc
           Com.CommandText = "ChequesAgregar"
           Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
           Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsRecibo!fecha)
           Com.Parameters.Append Com.CreateParameter("Movimiento", adInteger, adParamInput, , pRsValor!Venta)
           Com.Parameters.Append Com.CreateParameter("MoviProv", adInteger, adParamInput, , 0)
           Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
           Com.Parameters.Append Com.CreateParameter("Numero", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
           Com.Parameters.Append Com.CreateParameter("Titular", adVarChar, adParamInput, 50, pRsValor!Nombre & " " & pRsValor!NumeroTarjeta)
           Com.Parameters.Append Com.CreateParameter("Cliente", adVarChar, adParamInput, 15, pRsRecibo!Cliente)
           Com.Parameters.Append Com.CreateParameter("Proveedor", adVarChar, adParamInput, 15, "")
           Com.Parameters.Append Com.CreateParameter("FAcre", adDate, adParamInput, , pRsValor!FechaAcreditacion)
           Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
           Com.Execute
       End If
       pRsValor.MoveNext
     Loop
     Set Com = Nothing
          
     Dim cRsl As New ClsLectura
     Set cRsl = New ClsLectura
     If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & pRsRecibo!Comprobante & " AND Usuario=" & nUsuario) = False Then
         Set Com = New ADODB.Command
         Com.ActiveConnection = db
         Com.CommandType = adCmdStoredProc
         Com.CommandText = "ComprobanteActualizarNumero"
         Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsRecibo!Comprobante)
         Com.Execute
         Set Com = Nothing
     End If
  End If
    Set Com = Nothing
'
          
     pRsAplica.MoveFirst
     Do While Not pRsAplica.EOF
        ' Graba Aplicaciones
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "AplicacionAgregar"

    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente

       iCampos = pRsAplica.Fields.Count - 1
       For i = 0 To iCampos
           If pRsAplica(i).Name <> "Id" Then
              If pRsAplica(i).Type = adDouble Or pRsAplica(i).Type = adNumeric Then
                 Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
              Else
                 Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, pRsAplica(i).Type, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
              End If
          End If
      Next
      Com.Execute
      db.Execute "INSERT INTO DetallesComprobantes (Movimiento,Descripcion,Tasa,PrecioTotal) VALUES(" & pRsVenta!Movimiento & ",'COBRANZA'," & pRsAplica!Venta & "," & pRsAplica!importe & ")"
      pRsAplica.MoveNext
    Loop
        
    Set Com = Nothing
  ' Grabo el Ultimo Comprobante de Pago
    
  ' Grabo el Ultimo Cheque Id

  
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  
  If pRsBanco.BOF = False And pRsBanco.EOF = False Then
     pRsBanco.MoveFirst
     Do While Not pRsBanco.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoBancoAgregar"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsBanco!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsBanco!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsBanco!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsBanco!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsBanco!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsBanco!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsBanco!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsBanco!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsBanco!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsBanco!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsBanco!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsBanco!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsBanco!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsBanco!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsBanco!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsBanco!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsBanco!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsBanco!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsBanco!Abierta)
        Com.Execute
        pRsBanco.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  If pRet.RecordCount <> 0 Then
     pRet.MoveFirst
     Do While Not pRet.EOF
     
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "RetencionesAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
  
        iCampos = pRet.Fields.Count - 1
        For i = 0 To iCampos
            If pRet(i).Name <> "Id" Then
               If pRet(i).Type = adDouble Or pRet(i).Type = adNumeric Then
                  Com.Parameters.Append Com.CreateParameter(pRet(i).Name, adDouble, adParamInput, pRet(i).DefinedSize, pRet(i).Value)
               Else
                  Com.Parameters.Append Com.CreateParameter(pRet(i).Name, pRet(i).Type, adParamInput, pRet(i).DefinedSize, pRet(i).Value)
               End If
            End If
         Next
         Com.Execute
         Set Com = Nothing
         pRet.MoveNext
     Loop
  End If
    
'  Set Com = New ADODB.Command
'  Com.ActiveConnection = db
'  Com.CommandType = adCmdStoredProc
'  Com.CommandText = "MovimientoActualizarNumero"
'  Com.Execute
'  Set Com = Nothing

  bContado = False
  
  
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    bContado = False
    err.Raise err.Number

End Sub
Public Sub AgregarCobranzaVenta(pRsVenta As Recordset, pRsRecibo As Recordset, pRsValor As Recordset, pRsAplica As Recordset, pRsCaja As Recordset, pRsBanco As Recordset, pRet As Recordset)
  Dim iCampos As Integer, ID As Long
  Dim cRse As ClsEscritura
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  
  Set Com = New ADODB.Command
  Set cRse = New ClsEscritura
    
  db.Open sDb
  db.BeginTrans

  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
     If bContado = False Then
         Com.ActiveConnection = db
         Com.CommandType = adCmdStoredProc
         Com.CommandText = "CabezaAgregar"
        
      ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
      
         iCampos = pRsVenta.Fields.Count - 1
         For i = 0 To iCampos
             If pRsVenta(i).Name <> "Id" Then
                If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
                   Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
                Else
                   Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
                End If
             End If
         Next i
         Com.Execute
         Set Com = Nothing
     End If
     pRsValor.MoveFirst
     Do While Not pRsValor.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "ValorAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput, , pRsValor!Venta)
        Com.Parameters.Append Com.CreateParameter("CodPago", adInteger, adParamInput, , pRsValor!CodPago)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsValor!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
        Com.Parameters.Append Com.CreateParameter("IdCheque", adInteger, adParamInput, , IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque))
        Com.Parameters.Append Com.CreateParameter("NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FAcredi", adDate, adParamInput, , pRsValor!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FVencim", adDate, adParamInput, , pRsValor!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autoriz", adVarChar, adParamInput, 30, pRsValor!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsValor!Cupon)
        Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
        Com.Parameters.Append Com.CreateParameter("PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta)
        Com.Execute
        pRsValor.MoveNext
     Loop
     Set Com = Nothing
     
               
'     Set Com = New ADODB.Command
'     Com.ActiveConnection = db
'     Com.CommandType = adCmdStoredProc
'     Com.CommandText = "ComprobanteActualizarNumero"
'     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
'     Com.Execute
'
'     Set Com = Nothing
     
    Set Com = New ADODB.Command
     
    If pRsCaja.BOF = False And pRsCaja.EOF = False Then
       pRsCaja.MoveFirst
       Do While Not pRsCaja.EOF
          Set Com = New ADODB.Command
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "MovimientoCajaAgregar"
          Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
          Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
          Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
          Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
          Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
          Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
          Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
          Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
          Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
          Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
          Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
          Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
          Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
          Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
          Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
          Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
          Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
          Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
          Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
          Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
          Com.Execute
          pRsCaja.MoveNext
       Loop
       Set Com = Nothing
    End If
  
    pRsValor.MoveFirst
    Do While Not pRsValor.EOF
       If pRsValor!IdCheque = 1 Then
           Set Com = New ADODB.Command
           Com.ActiveConnection = db
           Com.CommandType = adCmdStoredProc
           Com.CommandText = "ChequesAgregar"
           Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
           Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsVenta!fecha)
           Com.Parameters.Append Com.CreateParameter("Movimiento", adInteger, adParamInput, , pRsValor!Venta)
           Com.Parameters.Append Com.CreateParameter("MoviProv", adInteger, adParamInput, , 0)
           Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
           Com.Parameters.Append Com.CreateParameter("Numero", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
           Com.Parameters.Append Com.CreateParameter("Titular", adVarChar, adParamInput, 50, pRsValor!Nombre)
           Com.Parameters.Append Com.CreateParameter("Cliente", adVarChar, adParamInput, 15, pRsVenta!Cliente)
           Com.Parameters.Append Com.CreateParameter("Proveedor", adVarChar, adParamInput, 15, "")
           Com.Parameters.Append Com.CreateParameter("FAcre", adDate, adParamInput, , pRsValor!FechaAcreditacion)
           Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
           Com.Execute
      End If
      pRsValor.MoveNext
    Loop
    Set Com = Nothing
      
    If pRsBanco.BOF = False And pRsBanco.EOF = False Then
       pRsBanco.MoveFirst
       Do While Not pRsBanco.EOF
          Set Com = New ADODB.Command
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "MovimientoBancoAgregar"
          Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsBanco!Movimiento)
          Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsBanco!Comprobante)
          Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsBanco!Numero)
          Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsBanco!sucursal)
          Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsBanco!fecha)
          Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsBanco!Hora)
          Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsBanco!Caja)
          Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsBanco!formapago)
          Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsBanco!Nombre)
          Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsBanco!Banco)
          Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsBanco!NumeroCheque)
          Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsBanco!FechaAcreditacion)
          Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsBanco!NumeroTarjeta)
          Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsBanco!fechavencimiento)
          Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsBanco!Autorizacion)
          Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsBanco!Cupon)
          Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsBanco!Debe)
          Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsBanco!Haber)
          Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsBanco!Abierta)
          Com.Execute
          pRsBanco.MoveNext
       Loop
       Set Com = Nothing
    End If
    
    If pRet.RecordCount <> 0 Then
       pRet.MoveFirst
       Do While Not pRet.EOF
       
          Set Com = New ADODB.Command
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "RetencionesAgregar"
      
      ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
    
          iCampos = pRet.Fields.Count - 1
          For i = 0 To iCampos
              If pRet(i).Name <> "Id" Then
                 If pRet(i).Type = adDouble Or pRet(i).Type = adNumeric Then
                    Com.Parameters.Append Com.CreateParameter(pRet(i).Name, adDouble, adParamInput, pRet(i).DefinedSize, pRet(i).Value)
                 Else
                    Com.Parameters.Append Com.CreateParameter(pRet(i).Name, pRet(i).Type, adParamInput, pRet(i).DefinedSize, pRet(i).Value)
                 End If
              End If
           Next
           Com.Execute
           Set Com = Nothing
           pRet.MoveNext
       Loop
    End If
      
  bContado = False
  
 ' Call cRse.GuardarEventoCaja(db, "venta", "insert", pRsCaja, pRsRecibo)
      
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    bContado = False
    db.RollbackTrans
    db.Close
    err.Raise err.Number

End Sub


Public Sub AgregarPagos(pRsVenta As Recordset, pRsRecibo As Recordset, pRsValor As Recordset, pRsAplica As Recordset, pRsCaja As Recordset, pRsBanco As Recordset, pRsGanancia As Recordset)
  Dim iCampos As Integer, ID As Long
  Dim Com As ADODB.Command
  Dim i As Long
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  Dim sCuit As String, nAlicuota As Double, nRetenido As Double, nNum As Integer, nMovi As Integer, sSql As String, nRet As Long, nGan As Long
  Dim cNum As ClsComprobantesL, cRsl As ClsLectura
'
  Set cNum = New ClsComprobantesL
  Set cRsl = New ClsLectura
'
   sCuit = cRsl.TraerValorDeUnCampo("Proveedores", "Cuit", "Proveedor='" & pRsVenta!Cliente & "'")
   nAlicuota = cRsl.TraerValorDeUnCampo("Proveedores", "Retencion", "Cuit='" & sCuit & "'")
               
   db.Open sDb
   db.BeginTrans

  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
  If pRsRecibo.RecordCount <> 0 Then
  
  
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
     iCampos = pRsVenta.Fields.Count - 1
     For i = 0 To iCampos
         If pRsVenta(i).Name <> "Id" Then
            If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
               Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
            Else
               Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
            End If
         End If
     Next i
     Com.Execute
     Set Com = Nothing

     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
  
    iCampos = pRsRecibo.Fields.Count - 1
    For i = 0 To iCampos
        If pRsRecibo(i).Name <> "Id" Then
           If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
              Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
           Else
              Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
           End If
        End If
    Next
    Com.Execute
    Set Com = Nothing
    
    pRsValor.MoveFirst
    Do While Not pRsValor.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "ValorAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput, , pRsValor!Venta)
        Com.Parameters.Append Com.CreateParameter("CodPago", adInteger, adParamInput, , pRsValor!CodPago)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsValor!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
        Com.Parameters.Append Com.CreateParameter("IdCheque", adInteger, adParamInput, , pRsValor!IdCheque)
        Com.Parameters.Append Com.CreateParameter("NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FAcredi", adDate, adParamInput, , pRsValor!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FVencim", adDate, adParamInput, , pRsValor!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autoriz", adVarChar, adParamInput, 30, pRsValor!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsValor!Cupon)
        Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
        Com.Parameters.Append Com.CreateParameter("PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta)
        Com.Execute
        pRsValor.MoveNext
     Loop
     Set Com = Nothing
     
    pRsValor.MoveFirst
    ID = IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque)

    Do While Not pRsValor.EOF
       If Not IsNull(pRsValor!IdCheque) And pRsValor!IdCheque <> 0 Then
          Set Com = New ADODB.Command
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "ChequesModif"
          Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque))
          Com.Parameters.Append Com.CreateParameter("FechaSalida", adDate, adParamInput, , pRsRecibo!fecha)
          Com.Parameters.Append Com.CreateParameter("Proveedor", adVarChar, adParamInput, 15, pRsRecibo!Cliente)
          Com.Parameters.Append Com.CreateParameter("MoviProv", adInteger, adParamInput, , pRsRecibo!Venta)
          Com.Execute
        End If
        pRsValor.MoveNext
     Loop
     Set Com = Nothing
'
          
      If nValorRetencion <> 0 Then
         nGan = cNum.TraerUltimoSistema("Ganancia")
      End If
          
     pRsAplica.MoveFirst
     Do While Not pRsAplica.EOF
     
        If nAlicuota <> 0 Then
         If cRsl.TraerValorDeUnCampoSin("Comprobantes", "Iva", "Id=" & pRsAplica!Comprobante) = 1 Then
            If cRsl.TraerCantidadSin("CabRetencion", "Id", "Movimiento=" & pRsAplica!Venta) = 0 Then
               If nRet = 0 Then
                  nRet = cNum.TraerUltimoSistema("Retencion")
               End If
               nRetenido = cRsl.TraerValorDeUnCampoSin("CabComprobantes", "Neto", "Movimiento=" & pRsAplica!Venta) * nAlicuota / 100
               db.Execute "INSERT INTO CabRetencion (MoviPag,Movimiento,Provincia,FechaRec,Fecha,Numero,OPago,Proveedor,Base,Alicuota,Importe) VALUES(" & pRsVenta!Movimiento & "," & pRsAplica!Venta & ",1,'" & FechaAmericana(pRsVenta!fecha) & "','" & FechaAmericana(cRsl.TraerValorDeUnCampoSin("CabComprobantes", "Fecha", "Movimiento=" & pRsAplica!Venta)) & "'," & nRet & ",'" & Format(pRsAplica!sucursal, "0000") & "-" & Format(pRsAplica!Numero, "00000000") & "','" & pRsVenta!Cliente & "'," & cRsl.TraerValorDeUnCampoSin("CabComprobantes", "Neto", "Movimiento=" & pRsAplica!Venta) & "," & nAlicuota & "," & Math.Round(nRetenido, 2) & ")"
            End If
         End If
      End If
                 
        ' Graba Aplicaciones
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "AplicacionAgregar"

    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente

       iCampos = pRsAplica.Fields.Count - 1
       For i = 0 To iCampos
           If pRsAplica(i).Name <> "Id" Then
              If pRsAplica(i).Name = "Retencion" Then
                If nRet <> 0 Then
                   Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, nRet)
                Else
                   Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, 0)
                End If
              Else
                  If pRsAplica(i).Name <> "Ganancia" Then
                     If pRsAplica(i).Type = adDouble Or pRsAplica(i).Type = adNumeric Then
                        Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
                     Else
                        Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, pRsAplica(i).Type, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
                     End If
                  Else
                     If nGan <> 0 Then
                        Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, nGan)
                     Else
                        Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, 0)
                     End If
                  End If
              End If
           End If
       Next
      
       Com.Execute
      
      db.Execute "INSERT INTO DetallesComprobantes (Movimiento,Descripcion,Tasa,PrecioTotal) VALUES(" & pRsVenta!Movimiento & ",'PAGOS REALIZADOS'," & pRsAplica!Venta & "," & pRsAplica!importe & ")"
      If cRsl.TraerValorDeUnCampo("Proveedores", "CatGanancia", "Proveedor='" & pRsVenta!Cliente & "'") <> 5 Then
         db.Execute "INSERT INTO CabGanancia (MoviPag,Movimiento,Provincia,FechaRec,Fecha,Numero,OPago,Proveedor,Base,Alicuota,Importe) VALUES(" & pRsVenta!Movimiento & "," & pRsAplica!Venta & ",1,'" & FechaAmericana(pRsVenta!fecha) & "','" & FechaAmericana(cRsl.TraerValorDeUnCampoSin("CabComprobantes", "Fecha", "Movimiento=" & pRsAplica!Venta)) & "'," & nGan & ",'" & Format(pRsAplica!sucursal, "0000") & "-" & Format(pRsAplica!Numero, "00000000") & "','" & pRsVenta!Cliente & "'," & cRsl.TraerValorDeUnCampoSin("CabComprobantes", "Neto", "Movimiento=" & pRsAplica!Venta) & ",0," & nValorRetencion & ")"
      End If
      
      
      pRsAplica.MoveNext
    Loop
    Set Com = Nothing
    
  End If
  
  Dim sFec1 As String, sFec2 As String
'
   sFec1 = ObtenerPrimerUltimoDia(pRsVenta!fecha)(0)
   sFec2 = ObtenerPrimerUltimoDia(pRsVenta!fecha)(1)

  If pRsGanancia.RecordCount <> 0 Then
     Dim nAnu As Byte
     nAnu = 0
     Do While Not pRsGanancia.EOF
        If pRsGanancia!Retenido <> 0 Then
           nAnu = 1
           db.Execute "UPDATE TblGananciaAcum SET Anulado=1 WHERE Proveedor='" & pRsVenta!Cliente & "' AND Fecha>='" & sFec1 & "' AND Fecha<='" & sFec2 & "'"
           db.Execute "UPDATE CabGanancia SET MoviPag=" & pRsVenta!Movimiento & ",Numero=" & nGan & ",Importe=" & pRsGanancia!Retenido & " WHERE Proveedor='" & pRsVenta!Cliente & "' AND FechaRec>='" & sFec1 & "' AND FechaRec<='" & sFec2 & "' AND Numero=0"
        End If
        
        db.Execute "INSERT INTO TblGananciaAcum (Movimiento,Proveedor,Fecha,Importe,Retenido,Anulado) VALUES(" & pRsGanancia!Movimiento & ",'" & pRsVenta!Cliente & "','" & pRsVenta!fecha & "'," & pRsGanancia!importe & "," & pRsGanancia!Retenido & "," & nAnu & ")"
        pRsGanancia.MoveNext
     Loop
  End If

  
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
  End If
  
    If pRsBanco.BOF = False And pRsBanco.EOF = False Then
     pRsBanco.MoveFirst
     Do While Not pRsBanco.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoBancoAgregar"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsBanco!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsBanco!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsBanco!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsBanco!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsBanco!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsBanco!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsBanco!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsBanco!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsBanco!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsBanco!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsBanco!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsBanco!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsBanco!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsBanco!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsBanco!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsBanco!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsBanco!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsBanco!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsBanco!Abierta)
        Com.Execute
        pRsBanco.MoveNext
     Loop
     Set Com = Nothing
  End If

  nValorRetencion = 0

'  If pRsVenta.RecordCount <> 0 Then
'     Set Com = New ADODB.Command
'     Com.ActiveConnection = db
'     Com.CommandType = adCmdStoredProc
'     Com.CommandText = "ComprobanteActualizarNumero"
'     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
'     Com.Execute
'     Set Com = Nothing
'  End If

   
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number

End Sub

Public Sub GrabarStock(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, pRsCantidad As Recordset, pTipo As Integer)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  Dim cRsl As ClsLectura, sSql As String
  Set cRsl = New ClsLectura

  Dim rLis As Recordset, nImpu As Single, cImpor As Double
  
  Set rLis = cRsl.TraerTodos("ListaDePrecio", "*")
  
  Dim nCanReg As Integer
    
    
  db.Open sDb
  db.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  
  nCanReg = rLis.RecordCount - 1
  
  Dim nLista() As Integer
  Dim nPorcen() As Double
  
  ReDim nLista(nCanReg)
  ReDim nPorcen(nCanReg)
  
  
  Dim P, h As Byte
  
  P = 0
  
  Do While Not rLis.EOF
     nLista(P) = rLis!ListaPrecio
     nPorcen(P) = rLis!Porcentaje
     P = P + 1
     rLis.MoveNext
  Loop
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13, "")
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("MoviRemi", adInteger, adParamInput, , 0)
        Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput, , 0)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = pRsDetalle!Descripcion
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
        Com.Parameters("PDesc").Value = pRsDetalle!Pdesc
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = IIf(IsNull(pRsDetalle!Impuesto), 0, pRsDetalle!Impuesto)
        Com.Parameters("Precio").Value = pRsDetalle!precio
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = IIf(IsNull(pRsDetalle!PrecioTotal), 0, pRsDetalle!PrecioTotal)
        Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
        Com.Execute
        
        
        If pRsDetalle!MoviRemi = 1 Then
           db.Execute "UPDATE Productos SET PrecioCompra=" & pRsDetalle!PrecioUnitario & ",FechaCompra='" & FechaAmericana(Date) & "',Actualiza=" & pRsDetalle!MoviRemi & " WHERE Producto='" & pRsDetalle!Producto & "'"
           If pRsDetalle!Tasa = 1 Then
              sSql = "Update Precios SET Precio=" & pRsDetalle!precio & " WHERE Lista=" & nListaStock & " and Producto='" & pRsDetalle!Producto & "'"
              db.Execute sSql
      
              sSql = "Update Productos SET PrecioCompra=" & pRsDetalle!PrecioUnitario & ",Actualiza=0, "
              sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
              sSql = sSql & " WHERE Producto='" & pRsDetalle!Producto & "'"
              db.Execute sSql
    
              If pRsDetalle!Cuenta <> 0 Then
                 sSql = "UPDATE Precios set Precios.Precio=" & pRsDetalle!precio & " "
                 sSql = sSql & "FROM Precios LEFT OUTER JOIN  Productos ON Precios.Producto = Productos.Producto "
                 sSql = sSql & "Where Productos.Representante =" & pRsDetalle!Cuenta & " OR Productos.Id=" & pRsDetalle!Cuenta & " "
                 db.Execute sSql
      
                 sSql = "Update Productos SET PrecioCompra=" & pRsDetalle!PrecioUnitario & ",Actualiza=0, "
                 sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
                 sSql = sSql & "Where Representante =" & pRsDetalle!Cuenta
                 db.Execute sSql
              End If
           End If
        End If
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Stock
  If pRsCantidad.RecordCount <> 0 Then
     Select Case pTipo
            Case 5
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 1)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Desde)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                    End If
                    Com.Execute
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       If bTransferencia = False Then
                          Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 1) ' CORREGIR TEMA ESTADOS
                       Else
                           Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 2) ' CORREGIR TEMA ESTADOS
                       End If
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Hasta)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
            Case 6, 7, 17, 18, 30, 33, 34
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
     End Select
  End If
  

  
 ' If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & pRsVenta!Comprobante & " AND Usuario=" & nUsuario) = False Then
     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "ComprobanteActualizarNumero"
     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
     Com.Execute
 ' End If
  
  Set cRsl = Nothing
  Set Com = Nothing
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub
Public Sub ModificarStock(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, pRsCantidad As Recordset, pTipo As Integer)
  
  Dim iCampos As Integer, sSql As String
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura

  Dim rLis As Recordset, nImpu As Single, cImpor As Double
  
  Set rLis = cRsl.TraerTodos("ListaDePrecio", "*")
  
  Dim nCanReg As Integer
    
    
  db.Open sDb
  db.BeginTrans
  
  sSql = "UPDATE CabComprobantes SET  Cliente='" & pRsVenta!Cliente & "',Neto=" & pRsVenta!Neto
  sSql = sSql & ",Motivo='" & pRsVenta!Motivo & "'"
  sSql = sSql & " WHERE Movimiento=" & pRsVenta!Movimiento
  db.Execute sSql
    
  ' Comienza la Operación de Venta
  db.Execute "DELETE  FROM DetallesComprobantes WHERE Movimiento=" & pRsVenta!Movimiento
  db.Execute "DELETE  FROM Cantidades WHERE Movimiento=" & pRsVenta!Movimiento
  
  
  nCanReg = rLis.RecordCount - 1
  
  Dim nLista() As Integer
  Dim nPorcen() As Double
  
  ReDim nLista(nCanReg)
  ReDim nPorcen(nCanReg)
  
  
  Dim P, h As Byte
  
  P = 0
  
  Do While Not rLis.EOF
     nLista(P) = rLis!ListaPrecio
     nPorcen(P) = rLis!Porcentaje
     P = P + 1
     rLis.MoveNext
  Loop

  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13, "")
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("MoviRemi", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput, , 0)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = pRsDetalle!Descripcion
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
        Com.Parameters("PDesc").Value = pRsDetalle!Pdesc
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = pRsDetalle!Impuesto
        Com.Parameters("Precio").Value = pRsDetalle!precio
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = pRsDetalle!PrecioTotal
        Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
        Com.Execute
        
        
        
         If pRsDetalle!MoviRemi = 1 Then
           db.Execute "UPDATE Productos SET PrecioCompra=" & pRsDetalle!PrecioUnitario & ",FechaCompra='" & FechaAmericana(Date) & "',Actualiza=" & pRsDetalle!MoviRemi & " WHERE Producto='" & pRsDetalle!Producto & "'"
           If pRsDetalle!Tasa = 1 Then
              sSql = "Update Precios SET Precio=" & pRsDetalle!precio & " WHERE Lista=" & nListaStock & " and Producto='" & pRsDetalle!Producto & "'"
              db.Execute sSql
      
              sSql = "Update Productos SET PrecioCompra=" & pRsDetalle!PrecioUnitario & ",Actualiza=0, "
              sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
              sSql = sSql & " WHERE Producto='" & pRsDetalle!Producto & "'"
              db.Execute sSql
    
              If pRsDetalle!Cuenta <> 0 Then
                 sSql = "UPDATE Precios set Precios.Precio=" & pRsDetalle!precio & " "
                 sSql = sSql & "FROM Precios LEFT OUTER JOIN  Productos ON Precios.Producto = Productos.Producto "
                 sSql = sSql & "Where Productos.Representante =" & pRsDetalle!Cuenta
                 db.Execute sSql
      
                 sSql = "Update Productos SET PrecioCompra=" & pRsDetalle!PrecioUnitario & ",Actualiza=0, "
                 sSql = sSql & "FechaAlta='" & FechaAmericana(Date) & "',FechaCompra='" & FechaAmericana(Date) & "',FechaBaja='" & FechaAmericana(Date) & "' "
                 sSql = sSql & "Where Representante =" & pRsDetalle!Cuenta
                 db.Execute sSql
              End If
           End If
        End If
       pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  ' Grabo el Ultimo Comprobante de Stock
  
  ' Finaliza la Operacion de Stock
  
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Stock
  If pRsCantidad.RecordCount <> 0 Then
     Select Case pTipo
            Case 5
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 1)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Desde)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                    End If
                    Com.Execute
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       If bTransferencia = False Then
                          Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 1) ' CORREGIR TEMA ESTADOS
                       Else
                           Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , 2) ' CORREGIR TEMA ESTADOS
                       End If
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Hasta)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
            Case 6, 7, 17, 18, 33, 34
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
     End Select
  End If
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
  
End Sub

Public Function SumoProducto(nComp As Integer) As Boolean
  Dim sSql As String, Rs As ADODB.Recordset
  
  Set Rs = New ADODB.Recordset
  On Error GoTo errHandler
  
  db.Open sDb

  sSql = "Select MoviStock FROM Comprobantes WHERE id=" & nComp
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs.RecordCount <> 0 Then
     If Rs!MoviStock = "Suma" Then
        SumoProducto = True
     End If
  End If
  
  db.Close
 Exit Function
 
errHandler:
    err.Raise err.Number
  
End Function
Public Function ProductoEsta(sPro As String, nDep As Integer) As Boolean
  Dim sSql As String, Rs As ADODB.Recordset
  
  Set Rs = New ADODB.Recordset
  On Error GoTo errHandler
  
  db.Open sDb
  
  sSql = "Select Producto FROM Cantidades WHERE Producto='" & sPro & "' and Deposito=" & nDep
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs.RecordCount <> 0 Then
     ProductoEsta = True
  End If
  db.Close
  
 Exit Function
 
errHandler:
    err.Raise err.Number
End Function

Public Function DepositoEsta(nDep As Integer) As Boolean
  Dim sSql As String, Rs As ADODB.Recordset
  
  Set Rs = New ADODB.Recordset
  On Error GoTo errHandler
  db.Open sDb

  sSql = "Select SUM(cantidadDebe - CantidadHAber) as Canti FROM Cantidades WHERE Deposito=" & nDep
  Rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If Rs!Canti <> 0 Then
     DepositoEsta = True
  End If
  db.Close
 Exit Function
 
errHandler:
    err.Raise err.Number
  
End Function

Public Sub GrabarCaja(pRsVenta As ADODB.Recordset, pRsCaja As ADODB.Recordset, pRsBanco As ADODB.Recordset, pRsValor As ADODB.Recordset, Optional bCierr As Boolean)
  Dim iCampos As Integer, sSql As String, ID As Long, chAux As Integer
  Dim Com As ADODB.Command
  Dim i As Integer, nCaj As Byte
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  
  ' Comienza la Operación de Caja
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next
  Com.Execute
  
  
  If pRsVenta.RecordCount > 0 Then
     pRsVenta.MoveFirst
     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "DetallesAgregar"
     Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
     Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
     Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13, "")
     Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
     Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , 0)
     Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , 0)
     Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
     Com.Parameters.Append Com.CreateParameter("MoviRemi", adInteger, adParamInput)
     Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput, , 0)
     Com.Parameters("Venta").Value = pRsVenta!Movimiento
     Com.Parameters("Producto").Value = ""
     Com.Parameters("Descripcion").Value = pRsVenta!Motivo
     Com.Parameters("Cantidad").Value = 1
     Com.Parameters("Cuenta").Value = pRsVenta!Lista
     Com.Parameters("PDesc").Value = 0
     Com.Parameters("Descuento").Value = 0
     Com.Parameters("Tasa").Value = 0
     Com.Parameters("Impuesto").Value = 0
     Com.Parameters("Precio").Value = 0
     Com.Parameters("PrecioUnitario").Value = IIf(pRsVenta!Debe = 0, pRsVenta!Haber, pRsVenta!Debe)
     Com.Parameters("PrecioTotal").Value = IIf(pRsVenta!Debe = 0, pRsVenta!Haber, pRsVenta!Debe)
     Com.Parameters("MoviRemi").Value = 0
     Com.Execute
     Set Com = Nothing
  End If
  
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     nCaj = pRsCaja!Caja
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        If pRsCaja!formapago = "CHEQUES PROPIOS" Then
           db.Execute "UPDATE CuentasBancarias SET Numero=" & pRsCaja!NumeroCheque + 1 & " WHERE Id=" & pRsCaja!Banco
        End If
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
     
  End If
  
  If pRsBanco.BOF = False And pRsBanco.EOF = False Then
     pRsBanco.MoveFirst
     nCaj = pRsBanco!Caja
     Do While Not pRsBanco.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoBancoAgregar"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsBanco!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsBanco!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsBanco!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsBanco!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsBanco!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsBanco!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsBanco!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsBanco!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsBanco!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsBanco!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsBanco!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsBanco!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsBanco!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsBanco!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsBanco!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsBanco!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsBanco!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsBanco!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsBanco!Abierta)
        Com.Execute
        pRsBanco.MoveNext
     Loop
     Set Com = Nothing
     
  End If
  
  If bCierr = True Then
     If pRsValor.RecordCount <> 0 Then
        pRsValor.MoveFirst
        Do While Not pRsValor.EOF
          If pRsValor!IdCheque <> 0 Then
              Set Com = New ADODB.Command
              Com.ActiveConnection = db
              Com.CommandType = adCmdStoredProc
              Com.CommandText = "ChequesAgregar"
              Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
              Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsVenta!fecha)
              Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
              Com.Parameters.Append Com.CreateParameter("MoviProv", adDouble, adParamInput, , 0)
              Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
              Com.Parameters.Append Com.CreateParameter("Numero", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
              Com.Parameters.Append Com.CreateParameter("Titular", adVarChar, adParamInput, 50, pRsValor!Nombre)
              Com.Parameters.Append Com.CreateParameter("Cliente", adVarChar, adParamInput, 15, "")
              Com.Parameters.Append Com.CreateParameter("Proveedor", adVarChar, adParamInput, 15, "CAXXX")
              Com.Parameters.Append Com.CreateParameter("FAcre", adDate, adParamInput, , pRsValor!FechaAcreditacion)
              Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
              Com.Execute
           End If
           pRsValor.MoveNext
        Loop
        Set Com = Nothing
     End If
  Else
     pRsValor.MoveFirst
     ID = IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque)
     Dim sProv As String
     
     If pRsVenta!Cliente = "" Then
        sProv = "CAXXX"
     Else
       sProv = pRsVenta!Cliente
     End If

     Do While Not pRsValor.EOF
        If Not IsNull(pRsValor!IdCheque) Then
           Set Com = New ADODB.Command
           Com.ActiveConnection = db
           Com.CommandType = adCmdStoredProc
           Com.CommandText = "ChequesModif"
           Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque))
           Com.Parameters.Append Com.CreateParameter("FechaSalida", adDate, adParamInput, , pRsVenta!fecha)
           Com.Parameters.Append Com.CreateParameter("Proveedor", adVarChar, adParamInput, 15, sProv)
           Com.Parameters.Append Com.CreateParameter("MoviProv", adDouble, adParamInput, , pRsVenta!Movimiento)
           Com.Execute
        End If
       pRsValor.MoveNext
     Loop
     Set Com = Nothing

  End If
  
  
  ' Grabo el Ultimo Comprobante de Caja
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "ComprobanteActualizarNumero"
  Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
  Com.Execute
  Set Com = Nothing
      
'  If bCierr = True Then
'     sSql = "UPDATE Caja SET Abierta='N' WHERE Abierta='S' and Caja=" & nCaj
'     db.Execute sSql
'  End If
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    MsgBox err.Description
End Sub

Public Sub AnularComprobanteCantidades(pMovimiento As Long, Optional pCob As Boolean, Optional pComprobante As Integer = 0, Optional pNumero As Long = 0, Optional pSucursal As Integer = 0)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  sSql = "UPDATE CabComprobantes SET Factura=0 WHERE Factura=" & pMovimiento
  db.Execute sSql

  ' Anular Cabeza Comprobante
  sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
  db.Execute sSql
  
  ' Anular Cantidades
  
  sSql = "DELETE  FROM  Cantidades  WHERE Movimiento=" & pMovimiento
  db.Execute sSql
  
  ' Anula Cuenta Corriente
  sSql = "DELETE  FROM CuentaCorriente  WHERE Venta=" & pMovimiento
  db.Execute sSql
  
  ' Anular Aplicaciones
  sSql = "DELETE  FROM  Aplicaciones  WHERE Venta=" & pMovimiento
  db.Execute sSql
  
  ' Anular Caja
  sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
  db.Execute sSql
    
  If pCob = False Then
     sSql = "UPDATE Cheques SET Anulado=1 WHERE Movimiento=" & pMovimiento
  Else
     sSql = "UPDATE Cheques SET Proveedor=Null,FechaSalida=Null WHERE MoviProv=" & pMovimiento
  End If
  db.Execute sSql
  
  If pComprobante <> 0 Then
     sSql = "DELETE  FROM Aplicaciones WHERE Comprobante=" & pComprobante & " AND Numero=" & pNumero & " AND Sucursal=" & pSucursal
     db.Execute sSql
  End If
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
  
End Sub
Public Sub AnulaAplicacion(pMovimiento As Long)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  sSql = "DELETE FROM Aplicaciones WHERE MoviCta=" & pMovimiento
  db.Execute sSql

  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub
Public Sub AnularComprobantes(pMovimiento As Long, pOpe As Integer, pTipo As Integer, Optional pFecha As String, Optional pProv As String = "", Optional pRsd As Recordset)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  Select Case pOpe
     Case 1  ' Venta
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM CabComprobantes  WHERE Lista=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  CuentaCorriente WHERE Venta=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  Cantidades WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                      
           sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                      
           sSql = "DELETE  FROM  Valores WHERE Venta=" & pMovimiento
           db.Execute sSql
           
           pRsd.MoveFirst
           Do While Not pRsd.EOF
              db.Execute "UPDATE CabComprobantes set Anulado=0 WHERE Movimiento=" & IIf(IsNull(pRsd!MoviRemi), 0, pRsd!MoviRemi)
              pRsd.MoveNext
           Loop
           
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "UPDATE  CabComprobantes set factura=0 WHERE factura=" & pMovimiento
           db.Execute sSql


     Case 2  ' Stock
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM  Cantidades WHERE Movimiento=" & pMovimiento
           db.Execute sSql

     Case 3  ' Caja
         Select Case pTipo
              Case 8  ' Suma
                   sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
                   db.Execute sSql
                   
                   sSql = "DELETE FROM  DetallesComprobantes WHERE Movimiento=" & pMovimiento
                   db.Execute sSql
                   
                   sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
                   db.Execute sSql

                   sSql = "DELETE FROM  Cheques WHERE Movimiento=" & pMovimiento
                   db.Execute sSql

                   sSql = "DELETE FROM  CtaBanco WHERE Movimiento=" & pMovimiento
                   db.Execute sSql

              Case 9  ' Resta
                   sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
                   db.Execute sSql
                   
                   sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
                   db.Execute sSql
                   
                   sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
                   db.Execute sSql

                   sSql = "UPDATE Cheques SET FechaSalida=Null,MoviProv=0 WHERE MoviProv=" & pMovimiento
                   db.Execute sSql
              
                   sSql = "DELETE FROM  CtaBanco WHERE Movimiento=" & pMovimiento
                   db.Execute sSql
         End Select
     Case 4  ' Cobranza
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  CuentaCorriente WHERE Venta=" & pMovimiento
           db.Execute sSql
  
           sSql = "DELETE  FROM  Valores WHERE Venta=" & pMovimiento
           db.Execute sSql
           
           sSql = "UPDATE CuentaCorriente set Anulado= 0 where CuentaCorriente.Venta  IN "
           sSql = sSql & "(SELECT Aplicaciones.Venta "
           sSql = sSql & "From Aplicaciones "
           sSql = sSql & "WHERE Aplicaciones.MoviCta = " & pMovimiento & ")"
           db.Execute sSql
    
           sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovimiento
           db.Execute sSql
  
           sSql = "DELETE  FROM Caja WHERE Movimiento=" & pMovimiento
           db.Execute sSql
  
           sSql = "DELETE FROM  Cheques WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE FROM  CtaBanco WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE FROM Retenciones WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
          
     Case 5  ' Nota de D/C Venta
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  CuentaCorriente WHERE Venta=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  Cantidades WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovimiento
           db.Execute sSql
     
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
     Case 6  ' Remitos Ventas
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  Cantidades WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
     
     Case 7  ' Remitos Compras
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  Cantidades WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
     Case 8  ' Compras
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  CuentaCorriente WHERE Venta=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  Cantidades WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
          sSql = "UPDATE CabComprobantes SET Factura=0 WHERE Factura=" & pMovimiento
          db.Execute sSql
          
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql

     Case 9  ' Pagos
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  CuentaCorriente WHERE Venta=" & pMovimiento
           db.Execute sSql
  
           sSql = "DELETE  FROM  Valores WHERE Venta=" & pMovimiento
           db.Execute sSql
  
           sSql = "DELETE  FROM CtaBanco WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "UPDATE CuentaCorriente set Anulado= 0 where CuentaCorriente.Venta  IN "
           sSql = sSql & "(SELECT     Aplicaciones.Venta "
           sSql = sSql & "From Aplicaciones "
           sSql = sSql & "WHERE Aplicaciones.MoviCta = " & pMovimiento & ") "
           db.Execute sSql
  
           sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovimiento
           db.Execute sSql
  
           sSql = "DELETE  FROM Caja WHERE Movimiento=" & pMovimiento
           db.Execute sSql
  
           sSql = "UPDATE Cheques SET FechaSalida=Null,MoviProv=0 WHERE MoviProv=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM Caja WHERE Movimiento=" & pMovimiento
           db.Execute sSql
  
           sSql = "DELETE  FROM CabRetencion WHERE Movipag=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM CabGanancia WHERE Movipag=" & pMovimiento
           db.Execute sSql
           
           Dim sFec1 As String, sFec2 As String

           sFec1 = ObtenerPrimerUltimoDia(pFecha)(0)
           sFec2 = ObtenerPrimerUltimoDia(pFecha)(1)
        
           sSql = "UPDATE TblGananciaAcum SET Anulado=0 WHERE Proveedor='" & pProv & "' AND Fecha>='" & sFec1 & "' AND Fecha<='" & sFec2 & "'"
           db.Execute sSql
         
           sSql = "DELETE FROM TblGananciaAcum WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "UPDATE CabGanancia  SET Numero=0,Importe=0 WHERE Proveedor='" & pProv & "' AND Fecha>='" & sFec1 & "' AND Fecha<='" & sFec2 & "'"
           db.Execute sSql

           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
     Case 10 ' Pedidos
     Case 11 ' Produccion
     Case 12 ' N de D/C Compras
           sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  CuentaCorriente WHERE Venta=" & pMovimiento
           db.Execute sSql
                   
           sSql = "DELETE  FROM  Cantidades WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovimiento
           db.Execute sSql
           
           sSql = "DELETE  FROM DetallesComprobantes  WHERE Movimiento=" & pMovimiento
           db.Execute sSql
           
     Case 13 ' Sin Usar
     Case 14 ' Banco
          Select Case pTipo
                  Case 26  ' Debito
                       sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
                       db.Execute sSql
                   
                       sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
                       db.Execute sSql

                       sSql = "UPDATE Cheques SET FechaSalida=Null,MoviProv=0 WHERE MoviProv=" & pMovimiento
                       db.Execute sSql
              
                       sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
                       db.Execute sSql
                 Case 27 ' Credito
                 
                       sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
                       db.Execute sSql
                   
                       sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
                       db.Execute sSql

                       sSql = "UPDATE Cheques SET FechaSalida=Null,MoviProv=0 WHERE MoviProv=" & pMovimiento
                       db.Execute sSql
              
                       sSql = "DELETE  FROM  CtaBanco WHERE Movimiento=" & pMovimiento
                       db.Execute sSql
          End Select
  End Select
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
  
End Sub


Public Sub AnularCaja(pMovimiento As Long)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  
  ' Anular Cabeza Comprobante
  sSql = "DELETE  FROM CabComprobantes  WHERE Movimiento=" & pMovimiento
  db.Execute sSql
    
  ' Anular Caja
  sSql = "DELETE  FROM  Caja WHERE Movimiento=" & pMovimiento
  db.Execute sSql
    
  
  sSql = "UPDATE Cheques SET Proveedor=Null,FechaSalida=Null WHERE MoviProv=" & pMovimiento
  db.Execute sSql
  
  sSql = "DELETE FROM  CtaBanco WHERE Movimiento=" & pMovimiento
  db.Execute sSql
  
  sSql = "DELETE FROM  Cheques WHERE Movimiento=" & pMovimiento
  db.Execute sSql

  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
  
End Sub

Public Sub DesImputarComprobante(pMovi As Long)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  
  sSql = "UPDATE CabComprobantes SET Motivo='ADE' WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  sSql = "UPDATE CuentaCorriente set Anulado= 0 where CuentaCorriente.Venta  IN "
  sSql = sSql & "(SELECT     Aplicaciones.Venta "
  sSql = sSql & "From Aplicaciones "
  sSql = sSql & "WHERE Aplicaciones.MoviCta = " & pMovi & ") "
  db.Execute sSql

  
  ' Anular Aplicaciones
  sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovi
  db.Execute sSql
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    err.Raise err.Number

End Sub

Public Sub AnularAdeCliente(pComprobante As Integer, pSucursal As Integer, pNumero As Long, pMovi As Long, Optional pComp As Boolean)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  
  ' Anula Cuenta Corriente
  sSql = "DELETE  FROM CabComprobantes WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  
  sSql = "DELETE  FROM CuentaCorriente  WHERE Venta=" & pMovi
  db.Execute sSql
  
  ' Anular Aplicaciones
  sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovi
  db.Execute sSql
  
  ' Anular Caja
  sSql = "DELETE  FROM Caja WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  sSql = "DELETE Cheques WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  sSql = "DELETE  FROM CtaBanco WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    err.Raise err.Number
  
End Sub

Public Sub AnularAdeProveedores(pComprobante As Integer, pSucursal As Integer, pNumero As Long, pMovi As Long, Optional pComp As Boolean)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  
  ' Anula Cuenta Corriente
  sSql = "DELETE  FROM CabComprobantes WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  
  sSql = "DELETE  FROM CuentaCorriente  WHERE Venta=" & pMovi
  db.Execute sSql
  
  ' Anular Aplicaciones
  sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovi
  db.Execute sSql
  
  ' Anular Caja
  sSql = "DELETE  FROM Caja WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  sSql = "UPDATE Cheques SET MoviProv=0,FechaSalida=null,Proveedor='' WHERE MoviProv=" & pMovi
  db.Execute sSql
  
  sSql = "DELETE  FROM CtaBanco WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    err.Raise err.Number
  
End Sub


Public Sub AnularComprobante(pComprobante As Integer, pSucursal As Integer, pNumero As Long, pMovi As Long, Optional pComp As Boolean)
  Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
  
  ' Anula Cuenta Corriente
  sSql = "DELETE  FROM CabComprobantes WHERE Comprobante=" & pComprobante
  sSql = sSql & " AND Sucursal=" & pSucursal & " and Numero=" & pNumero
  db.Execute sSql
  
  
  sSql = "DELETE  FROM CuentaCorriente  WHERE Comprobante=" & pComprobante
  sSql = sSql & " AND Sucursal=" & pSucursal & " and Numero=" & pNumero
  db.Execute sSql
  
  ' Anular Aplicaciones
  sSql = "DELETE  FROM  Aplicaciones WHERE  MoviCta=" & pMovi
  db.Execute sSql
  
  ' Anular Caja
  sSql = "DELETE  FROM Caja WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  sSql = "UPDATE Cheques SET FechaSalida=null WHERE MoviProv=" & pMovi
  db.Execute sSql
  
  sSql = "DELETE  FROM  Valores WHERE Venta=" & pMovi
  db.Execute sSql
  
  sSql = "DELETE  FROM CtaBanco WHERE Movimiento=" & pMovi
  db.Execute sSql
  
  If pComp = True Then
     sSql = "UPDATE Cheques SET MoviProv=0,FechaSalida=null,Proveedor='' WHERE MoviProv=" & pMovi
     db.Execute sSql
  Else
     sSql = "DELETE  FROM  Cheques  WHERE Movimiento=" & pMovi
     db.Execute sSql
  End If
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    err.Raise err.Number
  
End Sub

Public Sub ActualizarFechasCompras(pFecha As Date, pFIva As Date, pMovi As Long)
    Dim sSql As String
    On Error GoTo errHandler
  
    db.Open sDb
    db.BeginTrans
    
     sSql = "UPDATE CabComprobantes SET Fecha='" & pFecha & "',FechaIva='" & pFIva & "'  WHERE Movimiento=" & pMovi
     db.Execute sSql
    
    db.CommitTrans
    db.Close
Exit Sub
    
errHandler:
    db.RollbackTrans
    err.Raise err.Number
End Sub

Public Sub GrabarAplicaciones(pRsAplica As Recordset)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  
  ' Comienza la Operación de Venta
  
  pRsAplica.MoveFirst
  Do While Not pRsAplica.EOF
      ' Graba Aplicaciones
     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "AplicaAuxiliarAgregar"

    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente

       iCampos = pRsAplica.Fields.Count - 1
     For i = 0 To iCampos
         If pRsAplica(i).Name <> "Id" Then
            If pRsAplica(i).Type = adDouble Or pRsAplica(i).Type = adNumeric Then
               Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
            Else
                Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, pRsAplica(i).Type, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
            End If
         End If
     Next
     Com.Execute
     pRsAplica.MoveNext
   Loop
   
   db.CommitTrans
   db.Close
   Set Com = Nothing
   Exit Sub
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number

End Sub

Public Sub AgregarNotaCD(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, _
  pRsRecibo As Recordset, pRsCantidad As Recordset)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13, "")
        Com.Parameters.Append Com.CreateParameter("Cantidad", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = pRsDetalle!Descripcion
        Com.Parameters("Corte").Value = pRsDetalle!Corte
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = pRsDetalle!Impuesto
        Com.Parameters("Precio").Value = pRsDetalle!precio
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = pRsDetalle!PrecioTotal
        Com.Execute
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  ' Grabo el Ultimo Comprobante de Venta
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "ComprobanteActualizarNumero"
  Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
  Com.Execute
  Set Com = Nothing
  ' Actualizo La ultima Venta
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "MovimientoActualizarNumero"
  Com.Execute
  Set Com = Nothing
  
  ' Finaliza la Operacion de Venta
  
  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
  If pRsRecibo.RecordCount <> 0 Then
     pRsRecibo.MoveFirst
     Do While Not pRsRecibo.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
    
       iCampos = pRsRecibo.Fields.Count - 1
       For i = 0 To iCampos
           If pRsRecibo(i).Name <> "Id" Then
              If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              Else
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              End If
           End If
       Next
       Com.Execute
       pRsRecibo.MoveNext
    Loop
  End If
  Set Com = Nothing
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Venta
  If pRsCantidad.RecordCount <> 0 Then
     pRsCantidad.MoveFirst
     Do While Not pRsCantidad.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "CantidadesAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
        Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
        Com.Parameters.Append Com.CreateParameter("Lote", adInteger, adParamInput, , pRsCantidad!Lote)
        Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
        Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
        Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
        Com.Execute
        pRsCantidad.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Public Sub GrabarReceta(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabRecetaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetalleRecetaAgregar"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Costo", adSingle, adParamInput)
        Com.Parameters("Movimiento").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = pRsDetalle!Descripcion
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Costo").Value = pRsDetalle!Costo
        Com.Execute
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  ' Actualizo La ultima Venta
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "MovimientoActualizarNumero"
  Com.Execute
  Set Com = Nothing
  
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub
Public Sub ModificarReceta(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset)
  
  Dim iCampos As Integer, sSql As String
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
      
  ' Comienza la Operación de Venta
  db.Execute "DELETE  FROM DetalleReceta WHERE Movimiento=" & pRsVenta!Movimiento
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetalleRecetaAgregar"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Costo", adSingle, adParamInput)
        Com.Parameters("Movimiento").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = pRsDetalle!Descripcion
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Costo").Value = pRsDetalle!Costo
        Com.Execute
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
  
End Sub

Public Sub GrabarTransporte(pRsDetalle As ADODB.Recordset, pMovi As Long)
  
  Dim iCampos As Integer, sSql As String
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
      
  db.Execute "DELETE FROM RegTransporte WHERE Movimiento=" & pMovi
  
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "RegTransporte_A"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsDetalle!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsDetalle!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adVarChar, adParamInput, 50, pRsDetalle!Hora)
        Com.Parameters.Append Com.CreateParameter("Destinatario", adVarChar, adParamInput, 50, pRsDetalle!Destinatario)
        Com.Parameters.Append Com.CreateParameter("Direccion", adVarChar, adParamInput, 50, pRsDetalle!Direccion)
        Com.Parameters.Append Com.CreateParameter("Localidad", adVarChar, adParamInput, 50, pRsDetalle!Localidad)
        Com.Parameters.Append Com.CreateParameter("Telefono", adVarChar, adParamInput, 50, pRsDetalle!Telefono)
        Com.Parameters.Append Com.CreateParameter("Cuit", adVarChar, adParamInput, 13, pRsDetalle!Cuit)
        Com.Parameters.Append Com.CreateParameter("Transporte", adVarChar, adParamInput, 50, pRsDetalle!Transporte)
        Com.Parameters.Append Com.CreateParameter("Domicilio", adVarChar, adParamInput, 50, pRsDetalle!Domicilio)
        Com.Parameters.Append Com.CreateParameter("CuitTrans", adVarChar, adParamInput, 13, pRsDetalle!CuitTrans)
        Com.Parameters.Append Com.CreateParameter("PatenteCamion", adVarChar, adParamInput, 10, pRsDetalle!PatenteCamion)
        Com.Parameters.Append Com.CreateParameter("PatenteAcoplado", adVarChar, adParamInput, 10, pRsDetalle!PatenteAcoplado)
        Com.Parameters.Append Com.CreateParameter("Obsrevaciones", adVarChar, adParamInput, 50, pRsDetalle!Obsrevaciones)
        Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsDetalle!Estado)
        Com.Execute
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number

End Sub

Public Sub GrabarTexto(pRsDetalle As ADODB.Recordset, pMovi As Long)
  
  Dim iCampos As Integer, sSql As String
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
      
  db.Execute "DELETE FROM CabTexto WHERE Movimiento=" & pMovi
  
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "CabTextoAgregar"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsDetalle!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Texto1", adVarChar, adParamInput, 60, pRsDetalle!Texto1)
        Com.Parameters.Append Com.CreateParameter("Texto2", adVarChar, adParamInput, 60, pRsDetalle!Texto2)
        Com.Parameters.Append Com.CreateParameter("Texto3", adVarChar, adParamInput, 60, pRsDetalle!Texto3)
        Com.Parameters.Append Com.CreateParameter("Texto4", adVarChar, adParamInput, 60, pRsDetalle!Texto4)
        Com.Parameters.Append Com.CreateParameter("TextoFiscal", adVarChar, adParamInput, 2000, pRsDetalle!TextoFiscal)
        Com.Execute
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number

End Sub


Public Sub AgregarAplicacion(pRsAplica As Recordset, Optional nLlama As Byte = 0)
  Dim iCampos As Integer, cRp As ClsLectura
  
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  
  Set Com = New ADODB.Command
  Set cRp = New ClsLectura
    
  db.Open sDb
  db.BeginTrans

  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
 If pRsAplica.RecordCount <> 0 Then
    pRsAplica.MoveFirst
    Do While Not pRsAplica.EOF
        ' Graba Aplicaciones
       Set Com = New ADODB.Command
       Com.ActiveConnection = db
       Com.CommandType = adCmdStoredProc
       Com.CommandText = "AplicacionAgregar"

    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente

       iCampos = pRsAplica.Fields.Count - 1
       For i = 0 To iCampos
           If pRsAplica(i).Name <> "Id" Then
              If pRsAplica(i).Type = adDouble Or pRsAplica(i).Type = adNumeric Then
                 Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, adDouble, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
              Else
                 Com.Parameters.Append Com.CreateParameter(pRsAplica(i).Name, pRsAplica(i).Type, adParamInput, pRsAplica(i).DefinedSize, pRsAplica(i).Value)
              End If
          End If
      Next
      Com.Execute
      
      If nLlama = 0 Then
         db.Execute "INSERT INTO TblGananciaAcum (Movimiento,Proveedor,Fecha,Importe,Retenido,Anulado) VALUES(" & pRsAplica!MoviCta & ",'" & cRp.TraerValorDeUnCampoSin("CabComprobantes", "Cliente", "Movimiento=" & pRsAplica!Venta) & "','" & Date & "'," & -pRsAplica!importe & ",0,0)"
         db.Execute "INSERT INTO CabGanancia (MoviPag,Movimiento,Provincia,FechaRec,Fecha,Numero,OPago,Proveedor,Base,Alicuota,Importe) VALUES(" & pRsAplica!Venta & "," & pRsAplica!MoviCta & ",1,'" & Date & "','" & FechaAmericana(cRp.TraerValorDeUnCampoSin("CabComprobantes", "Fecha", "Movimiento=" & pRsAplica!Venta)) & "',0,'" & Format(pRsAplica!sucursal, "0000") & "-" & Format(pRsAplica!Numero, "00000000") & "','" & cRp.TraerValorDeUnCampoSin("CabComprobantes", "Cliente", "Movimiento=" & pRsAplica!Venta) & "'," & -pRsAplica!importe & ",0,0)"
      End If

      pRsAplica.MoveNext
    Loop
    Set Com = Nothing
 End If
      
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
       
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number

End Sub


Public Sub AgregarAdelantos(pRsVenta As Recordset, pRsRecibo As Recordset, pRsValor As Recordset, pRsCaja As Recordset, pRsBanco As Recordset, bCierr As Boolean)
  Dim iCampos As Integer, ID As Long
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans

  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
    Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next i
  Com.Execute
  Set Com = Nothing

  
  If pRsRecibo.RecordCount <> 0 Then
     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
  
    iCampos = pRsRecibo.Fields.Count - 1
    For i = 0 To iCampos
        If pRsRecibo(i).Name <> "Id" Then
           If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
              Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
           Else
              Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
           End If
        End If
    Next
    Com.Execute
    Set Com = Nothing
 End If
     
  If bCierr = True Then
     If pRsValor.RecordCount <> 0 Then
        pRsValor.MoveFirst
        Do While Not pRsValor.EOF
          If pRsValor!IdCheque <> 0 Then
              Set Com = New ADODB.Command
              Com.ActiveConnection = db
              Com.CommandType = adCmdStoredProc
              Com.CommandText = "ChequesAgregar"
              Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
              Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsVenta!fecha)
              Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
              Com.Parameters.Append Com.CreateParameter("MoviProv", adDouble, adParamInput, , 0)
              Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
              Com.Parameters.Append Com.CreateParameter("Numero", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
              Com.Parameters.Append Com.CreateParameter("Titular", adVarChar, adParamInput, 50, pRsValor!Nombre)
              Com.Parameters.Append Com.CreateParameter("Cliente", adVarChar, adParamInput, 15, "")
              Com.Parameters.Append Com.CreateParameter("Proveedor", adVarChar, adParamInput, 15, "CAXXX")
              Com.Parameters.Append Com.CreateParameter("FAcre", adDate, adParamInput, , pRsValor!FechaAcreditacion)
              Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
              Com.Execute
           End If
           pRsValor.MoveNext
        Loop
        Set Com = Nothing
     End If
  Else
     pRsValor.MoveFirst
     ID = IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque)
     Do While Not pRsValor.EOF
        If Not IsNull(pRsValor!IdCheque) Then
           Set Com = New ADODB.Command
           Com.ActiveConnection = db
           Com.CommandType = adCmdStoredProc
           Com.CommandText = "ChequesModif"
           Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque))
           Com.Parameters.Append Com.CreateParameter("FechaSalida", adDate, adParamInput, , pRsVenta!fecha)
           Com.Parameters.Append Com.CreateParameter("Proveedor", adVarChar, adParamInput, 15, "CAXXX")
           Com.Parameters.Append Com.CreateParameter("MoviProv", adDouble, adParamInput, , pRsVenta!Movimiento)
           Com.Execute
        End If
       pRsValor.MoveNext
     Loop
     Set Com = Nothing

  End If
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  If pRsBanco.BOF = False And pRsBanco.EOF = False Then
     pRsBanco.MoveFirst
     Do While Not pRsBanco.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoBancoAgregar"
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsBanco!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsBanco!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsBanco!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsBanco!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsBanco!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsBanco!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsBanco!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsBanco!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsBanco!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsBanco!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsBanco!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsBanco!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsBanco!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsBanco!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsBanco!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsBanco!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsBanco!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsBanco!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsBanco!Abierta)
        Com.Execute
        pRsBanco.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  ' Actualizo el Ultimo Movimiento
  Set Com = New ADODB.Command
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "MovimientoActualizarNumero"
  Com.Execute
  Set Com = Nothing

    Set Com = New ADODB.Command
    Com.ActiveConnection = db
    Com.CommandType = adCmdStoredProc
    Com.CommandText = "ComprobanteActualizarNumero"
    Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
    Com.Execute
    Set Com = Nothing


  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number

End Sub

Public Sub ActualizarMovimiento(pCompro As Integer)
   Dim Conexion As ADODB.Connection
   Dim Rs_Text As ADODB.Recordset
 
   Dim Com As ADODB.Command
   Set Com = New ADODB.Command

   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   
   Conexion.Open sDb

   Conexion.BeginTrans

   Set Com = New ADODB.Command
   Com.ActiveConnection = Conexion
   Com.CommandType = adCmdStoredProc
   Com.CommandText = "ComprobanteActualizarNumero"
   Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pCompro)
   Com.Execute
   Set Com = Nothing
  ' Actualizo La ultima Venta
   Set Com = New ADODB.Command
   Com.ActiveConnection = Conexion
   Com.CommandType = adCmdStoredProc
   Com.CommandText = "MovimientoActualizarNumero"
   Com.Execute
   Conexion.CommitTrans
   Conexion.Close
  
  Set Com = Nothing
 Exit Sub
    
errHandler:
    err.Raise err.Number
    Conexion.RollbackTrans
  
End Sub

Public Function GrabarSaldosCaja(pRsCaja As Recordset, pFecha As Date, pSaldo As Currency) As Boolean
  Dim iCampos As Integer, sSql As String, ID As Integer, chAux As Integer
  Dim Com As ADODB.Command
  Dim i As Integer, nCaj As Byte
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     
     db.Execute "INSERT INTO SaldoInicial (Suc,Fecha,Saldo) VALUES(" & nSucursal & ",'" & FechaAmericana(pFecha) & "'," & pSaldo & ")"
     
     db.CommitTrans
     db.Close
  End If
 Exit Function
    
errHandler:
    db.RollbackTrans
    db.Close
    MsgBox err.Description

End Function

Public Sub GrabarRemito(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, pRsCantidad As Recordset, pTipo As Integer)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura

  Dim rLis As Recordset, nImpu As Single, cImpor As Double
  
  Set rLis = cRsl.TraerTodos("ListaDePrecio", "*")
  
  Dim nCanReg As Integer
    
    
  db.Open sDb
  db.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  
  nCanReg = rLis.RecordCount - 1
  
  Dim nLista() As Integer
  Dim nPorcen() As Double
  
  ReDim nLista(nCanReg)
  ReDim nPorcen(nCanReg)
  
  
  Dim P, h As Byte
  
  P = 0
  
  Do While Not rLis.EOF
     nLista(P) = rLis!ListaPrecio
     nPorcen(P) = rLis!Porcentaje
     P = P + 1
     rLis.MoveNext
  Loop
  
  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        If pRsDetalle!Descripcion <> "" Then
          Set Com = New ADODB.Command
          Com.ActiveConnection = db
          Com.CommandType = adCmdStoredProc
          Com.CommandText = "DetallesAgregar"
          Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
          Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
          Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13, "")
          Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
          Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
          Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
          Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
          Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
          Com.Parameters.Append Com.CreateParameter("MoviRemi", adInteger, adParamInput, , 0)
          Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput, , 0)
          Com.Parameters("Venta").Value = pRsVenta!Movimiento
          Com.Parameters("Producto").Value = pRsDetalle!Producto
          Com.Parameters("Descripcion").Value = pRsDetalle!Descripcion
          Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
          Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
          Com.Parameters("PDesc").Value = pRsDetalle!Pdesc
          Com.Parameters("Descuento").Value = pRsDetalle!Descuento
          Com.Parameters("Tasa").Value = pRsDetalle!Tasa
          Com.Parameters("Impuesto").Value = IIf(IsNull(pRsDetalle!Impuesto), 0, pRsDetalle!Impuesto)
          Com.Parameters("Precio").Value = 0
          Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
          Com.Parameters("PrecioTotal").Value = IIf(IsNull(pRsDetalle!PrecioTotal), 0, pRsDetalle!PrecioTotal)
          Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
          Com.Execute
       End If
       pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Stock
  If pRsCantidad.RecordCount <> 0 Then
     Select Case pTipo
            Case 5
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Desde)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                    End If
                    Com.Execute
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Hasta)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
            Case 6, 7, 17, 18, 30, 33, 34
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
     End Select
  End If
  

  
  If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & pRsVenta!Comprobante & " AND Usuario=" & nUsuario) = False Then
     Set Com = New ADODB.Command
     Com.ActiveConnection = db
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "ComprobanteActualizarNumero"
     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
     Com.Execute
  End If
  
  Set cRsl = Nothing
  Set Com = Nothing
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub
Public Sub ModificarRemito(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, pRsCantidad As Recordset, pTipo As Integer)
  
  Dim iCampos As Integer, sSql As String
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura

  Dim rLis As Recordset, nImpu As Single, cImpor As Double
  
  Set rLis = cRsl.TraerTodos("ListaDePrecio", "*")
  
  Dim nCanReg As Integer
    
    
  db.Open sDb
  db.BeginTrans
  
  sSql = "UPDATE CabComprobantes SET  Cliente='" & pRsVenta!Cliente & "',Neto=" & pRsVenta!Neto
  sSql = sSql & ",Motivo='" & pRsVenta!Motivo & "'"
  sSql = sSql & " WHERE Movimiento=" & pRsVenta!Movimiento
  db.Execute sSql
    
  ' Comienza la Operación de Venta
  db.Execute "DELETE  FROM DetallesComprobantes WHERE Movimiento=" & pRsVenta!Movimiento
  db.Execute "DELETE  FROM Cantidades WHERE Movimiento=" & pRsVenta!Movimiento
  
  
  nCanReg = rLis.RecordCount - 1
  
  Dim nLista() As Integer
  Dim nPorcen() As Double
  
  ReDim nLista(nCanReg)
  ReDim nPorcen(nCanReg)
  
  
  Dim P, h As Byte
  
  P = 0
  
  Do While Not rLis.EOF
     nLista(P) = rLis!ListaPrecio
     nPorcen(P) = rLis!Porcentaje
     P = P + 1
     rLis.MoveNext
  Loop

  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13, "")
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("MoviRemi", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput, , 0)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = pRsDetalle!Descripcion
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
        Com.Parameters("PDesc").Value = pRsDetalle!Pdesc
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = pRsDetalle!Impuesto
        Com.Parameters("Precio").Value = pRsDetalle!precio
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = pRsDetalle!PrecioTotal
        Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
        Com.Execute
        
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
    
  ' Grabo el Ultimo Comprobante de Stock
  
  ' Finaliza la Operacion de Stock
  
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Stock
  If pRsCantidad.RecordCount <> 0 Then
     Select Case pTipo
            Case 5
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Desde)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                    End If
                    Com.Execute
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsVenta!Hasta)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , 0)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
            Case 6, 7, 17, 18, 33, 34
                 pRsCantidad.MoveFirst
                 Do While Not pRsCantidad.EOF
                    Set Com = New ADODB.Command
                    Com.ActiveConnection = db
                    Com.CommandType = adCmdStoredProc
                    Com.CommandText = "CantidadesAgregar"
                    If pRsCantidad!Producto <> "" Then
                       Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
                       Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
                       Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
                       Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
                       Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
                       Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
                       Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
                       Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
                       Com.Execute
                    End If
                    pRsCantidad.MoveNext
                 Loop
                 Set Com = Nothing
     End Select
  End If
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
  
End Sub

Public Sub AgregarConReciboResto(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, _
  pRsRecibo As Recordset, pRsValor As Recordset, pRsAplica As Recordset, pRsCantidad As Recordset, pRsCaja As Recordset, pRecibo As Recordset, Optional bRem As Boolean, Optional rRem As Recordset)
  Dim iCampos As Integer, sProv As String
  Dim Com As ADODB.Command
  Dim i As Integer
  Dim cRsl As ClsLectura
  
  Dim sConex As Connection
  
  Set sConex = New Connection
  
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  sConex.Open sDb
  sConex.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = sConex
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
      
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Name = "Cliente" Then sProv = pRsVenta(i).Value
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next i
  Com.Execute
  Set Com = Nothing

  If pRecibo.RecordCount <> 0 Then
      Set Com = New ADODB.Command
    
      Com.ActiveConnection = sConex
      Com.CommandType = adCmdStoredProc
      Com.CommandText = "CabezaAgregar"
      ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
    
      iCampos = pRecibo.Fields.Count - 1
      For i = 0 To iCampos
          If pRecibo(i).Name <> "Id" Then
             If pRecibo(i).Name = "Cliente" Then sProv = pRecibo(i).Value
             If pRecibo(i).Type = adDouble Or pRecibo(i).Type = adNumeric Then
                Com.Parameters.Append Com.CreateParameter(pRecibo(i).Name, adDouble, adParamInput, pRecibo(i).DefinedSize, pRecibo(i).Value)
             Else
                Com.Parameters.Append Com.CreateParameter(pRecibo(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRecibo(i).Value)
             End If
          End If
      Next i
      Com.Execute
      Set Com = Nothing
  End If

  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13)
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("MoviRemi", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput, , 0)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = IIf(pRsDetalle!Descripcion = "", " ", pRsDetalle!Descripcion)
        Com.Parameters("Leyenda").Value = IIf(pRsDetalle!Leyenda = "", " ", pRsDetalle!Leyenda)
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
        Com.Parameters("PDesc").Value = IIf(IsNull(pRsDetalle!Pdesc), 0, pRsDetalle!Pdesc)
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = pRsDetalle!Impuesto
        Com.Parameters("Precio").Value = IIf(IsNull(pRsDetalle!precio), pRsDetalle!PrecioUnitario, pRsDetalle!precio)
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = pRsDetalle!PrecioTotal
        Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
        Com.Execute
        
        If pRsVenta!tipo = 8 Then
           Dim nTotal As Double
           
           If pRsDetalle!MoviRemi = 1 Then
              nTotal = pRsDetalle!PrecioUnitario
              sConex.Execute "UPDATE Productos SET PrecioCompra=" & nTotal & ",FechaCompra='" & FechaAmericana(Date) & "',Actualiza=" & pRsDetalle!MoviRemi & " WHERE Producto='" & pRsDetalle!Producto & "'"
           End If
        End If
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  
  ' Actualizo el Ultimo Recibo
  If pRsCaja.RecordCount <> 0 Then
     Set Com = New ADODB.Command
     Com.ActiveConnection = sConex
     Com.CommandType = adCmdStoredProc
     Com.CommandText = "ComprobanteActualizarNumero"
     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsCaja!Comprobante)
     Com.Execute
     Set Com = Nothing
  End If
  
  ' Finaliza la Operacion de Venta
  
  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
  If pRsRecibo.RecordCount <> 0 Then
     pRsRecibo.MoveFirst
     Do While Not pRsRecibo.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
    
       iCampos = pRsRecibo.Fields.Count - 1
       For i = 0 To iCampos
           If pRsRecibo(i).Name <> "Id" Then
              If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              Else
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              End If
           End If
       Next
       Com.Execute
       pRsRecibo.MoveNext
    Loop
  End If
  Set Com = Nothing
  If pRsValor.RecordCount <> 0 Then
      pRsValor.MoveFirst
      Do While Not pRsValor.EOF
         Set Com = New ADODB.Command
         Com.ActiveConnection = sConex
         Com.CommandType = adCmdStoredProc
         Com.CommandText = "ValorAgregar"
         Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput, , pRsValor!Venta)
         Com.Parameters.Append Com.CreateParameter("CodPago", adInteger, adParamInput, , pRsValor!CodPago)
         Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsValor!formapago)
         Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre)
         Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
         Com.Parameters.Append Com.CreateParameter("IdCheque", adInteger, adParamInput, , 0)
         Com.Parameters.Append Com.CreateParameter("NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
         Com.Parameters.Append Com.CreateParameter("FAcredi", adDate, adParamInput, , pRsValor!FechaAcreditacion)
         Com.Parameters.Append Com.CreateParameter("NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta)
         Com.Parameters.Append Com.CreateParameter("FVencim", adDate, adParamInput, , pRsValor!fechavencimiento)
         Com.Parameters.Append Com.CreateParameter("Autoriz", adVarChar, adParamInput, 30, pRsValor!Autorizacion)
         Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsValor!Cupon)
         Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
         Com.Parameters.Append Com.CreateParameter("PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta)
         Com.Execute
         pRsValor.MoveNext
  Loop
  End If
  Set Com = Nothing
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Venta
  If bRem = False Then
     If pRsCantidad.RecordCount <> 0 Then
        pRsCantidad.MoveFirst
        Do While Not pRsCantidad.EOF
           If pRsCantidad!Producto <> "" Then
              Set Com = New ADODB.Command
              Com.ActiveConnection = sConex
              Com.CommandType = adCmdStoredProc
              Com.CommandText = "CantidadesAgregar"
              Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
              Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
              Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
              Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
              Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
              Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
              Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
              Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
              Com.Execute
           End If
           pRsCantidad.MoveNext
        Loop
        Set Com = Nothing
     End If
  Else
'     rRem.MoveFirst
'     Do While Not rRem.EOF
'        Dim sSql As String
'        sSql = "UPDATE CabComprobantes SET Factura=" & pRsVenta!Movimiento
'        sSql = sSql & " WHERE Movimiento=" & rRem!Movimiento
'        rRem.MoveNext
'        sConex.Execute sSql
'     Loop
  End If
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  Set cRsl = New ClsLectura
  
  If bDelivery = False Then
     sConex.Execute "DELETE  FROM auxiliar WHERE a_mesa=" & nMesa & " AND a_mensaje='ABIERTA'"
  Else
     sConex.Execute "DELETE  FROM auxiliar WHERE a_mesa=" & nMesa & " AND a_mensaje='DELIVERY'"
  End If

  sConex.CommitTrans
  sConex.Close
  Exit Sub
    
errHandler:
    sConex.RollbackTrans
    sConex.Close
    
    err.Raise err.Number
End Sub


Public Sub AgregarConRecibo(pRsVenta As ADODB.Recordset, pRsDetalle As ADODB.Recordset, _
  pRsRecibo As Recordset, pRsValor As Recordset, pRsAplica As Recordset, pRsCantidad As Recordset, pRsCaja As Recordset, pRecibo As Recordset, Optional bRem As Boolean, Optional rRem As Recordset)
  Dim iCampos As Integer, sProv As String
  Dim Com As ADODB.Command
  Dim i As Integer
  Dim cRsl As ClsLectura
  Dim cRse As ClsEscritura
  
  Dim sConex As Connection
  
  Set sConex = New Connection
  Set cRse = New ClsEscritura
  
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  sConex.Open sDb
  sConex.BeginTrans
  
  ' Comienza la Operación de Venta
  
  Com.ActiveConnection = sConex
  Com.CommandType = adCmdStoredProc
  Com.CommandText = "CabezaAgregar"
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
  
      
  iCampos = pRsVenta.Fields.Count - 1
  For i = 0 To iCampos
      If pRsVenta(i).Name <> "Id" Then
         If pRsVenta(i).Name = "Cliente" Then sProv = pRsVenta(i).Value
         If pRsVenta(i).Type = adDouble Or pRsVenta(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, adDouble, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsVenta(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRsVenta(i).Value)
         End If
      End If
  Next i
  Com.Execute
  Set Com = Nothing

  If pRecibo.RecordCount <> 0 Then
      Set Com = New ADODB.Command
    
      Com.ActiveConnection = sConex
      Com.CommandType = adCmdStoredProc
      Com.CommandText = "CabezaAgregar"
      ' recorre la coleccion fields y carga los parametros de la tabla de Ventas
    
      iCampos = pRecibo.Fields.Count - 1
      For i = 0 To iCampos
          If pRecibo(i).Name <> "Id" Then
             If pRecibo(i).Name = "Cliente" Then sProv = pRecibo(i).Value
             If pRecibo(i).Type = adDouble Or pRecibo(i).Type = adNumeric Then
                Com.Parameters.Append Com.CreateParameter(pRecibo(i).Name, adDouble, adParamInput, pRecibo(i).DefinedSize, pRecibo(i).Value)
             Else
                Com.Parameters.Append Com.CreateParameter(pRecibo(i).Name, pRsVenta(i).Type, adParamInput, pRsVenta(i).DefinedSize, pRecibo(i).Value)
             End If
          End If
      Next i
      Com.Execute
      Set Com = Nothing
  End If

  ' graba los Detalles
  If pRsDetalle.RecordCount > 0 Then
     pRsDetalle.MoveFirst
     Do While Not pRsDetalle.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "DetallesAgregar"
        Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25)
        Com.Parameters.Append Com.CreateParameter("Descripcion", adVarChar, adParamInput, 80)
        Com.Parameters.Append Com.CreateParameter("Leyenda", adVarChar, adParamInput, 13)
        Com.Parameters.Append Com.CreateParameter("Cantidad", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Cuenta", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PDesc", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Descuento", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Tasa", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Deposito", adDouble, adParamInput, , pRsDetalle!Deposito)
        Com.Parameters.Append Com.CreateParameter("Medida", adDouble, adParamInput, , pRsDetalle!Medida)
        Com.Parameters.Append Com.CreateParameter("Impuesto", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioUnitario", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("PrecioTotal", adDouble, adParamInput)
        Com.Parameters.Append Com.CreateParameter("MoviRemi", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("DetRemi", adInteger, adParamInput, , 0)
        Com.Parameters("Venta").Value = pRsVenta!Movimiento
        Com.Parameters("Producto").Value = pRsDetalle!Producto
        Com.Parameters("Descripcion").Value = IIf(pRsDetalle!Descripcion = "", " ", pRsDetalle!Descripcion)
        Com.Parameters("Leyenda").Value = IIf(pRsDetalle!Leyenda = "", " ", pRsDetalle!Leyenda)
        Com.Parameters("Cantidad").Value = pRsDetalle!Cantidad
        Com.Parameters("Cuenta").Value = pRsDetalle!Cuenta
        Com.Parameters("PDesc").Value = IIf(IsNull(pRsDetalle!Pdesc), 0, pRsDetalle!Pdesc)
        Com.Parameters("Descuento").Value = pRsDetalle!Descuento
        Com.Parameters("Tasa").Value = pRsDetalle!Tasa
        Com.Parameters("Impuesto").Value = pRsDetalle!Impuesto
        Com.Parameters("Precio").Value = pRsDetalle!PrecioTotal + pRsDetalle!Impuesto
        Com.Parameters("PrecioUnitario").Value = pRsDetalle!PrecioUnitario
        Com.Parameters("PrecioTotal").Value = pRsDetalle!PrecioTotal
        Com.Parameters("MoviRemi").Value = IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
        Com.Execute
        
        If pRsVenta!tipo = 8 Then
           Dim nTotal As Double
           
           If pRsDetalle!MoviRemi = 1 Then
              nTotal = pRsDetalle!PrecioUnitario
              sConex.Execute "UPDATE Productos SET PrecioCompra=" & nTotal & ",FechaCompra='" & FechaAmericana(Date) & "',Actualiza=" & pRsDetalle!MoviRemi & " WHERE Producto='" & pRsDetalle!Producto & "'"
           End If
        End If
        pRsDetalle.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  
  ' Actualizo el Ultimo Recibo
   If pRsCaja.RecordCount <> 0 Then
      Set Com = New ADODB.Command
      Com.ActiveConnection = sConex
      Com.CommandType = adCmdStoredProc
      Com.CommandText = "ComprobanteActualizarNumero"
      Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsCaja!Comprobante)
      Com.Execute
      Set Com = Nothing
   Else
      Set Com = New ADODB.Command
      Com.ActiveConnection = sConex
      Com.CommandType = adCmdStoredProc
      Com.CommandText = "ComprobanteActualizarNumero"
      Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsRecibo!Comprobante)
      Com.Execute
   End If
  
  ' Finaliza la Operacion de Venta
  
  ' Graba Recibo de Pago
  ' Comienza la Operacion de Cobranza y Valores de Pago
  
  If pRsRecibo.RecordCount <> 0 Then
     pRsRecibo.MoveFirst
     Do While Not pRsRecibo.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "CtaCteAgregar"
    
    ' recorre la coleccion fields y carga los parametros de la tabla de CuentaCoriente
    
       iCampos = pRsRecibo.Fields.Count - 1
       For i = 0 To iCampos
           If pRsRecibo(i).Name <> "Id" Then
              If pRsRecibo(i).Type = adDouble Or pRsRecibo(i).Type = adNumeric Then
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, adDouble, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              Else
                 Com.Parameters.Append Com.CreateParameter(pRsRecibo(i).Name, pRsRecibo(i).Type, adParamInput, pRsRecibo(i).DefinedSize, pRsRecibo(i).Value)
              End If
           End If
       Next
       Com.Execute
       pRsRecibo.MoveNext
    Loop
  End If
  Set Com = Nothing
  If pRsValor.RecordCount <> 0 Then
      pRsValor.MoveFirst
      Do While Not pRsValor.EOF
         Set Com = New ADODB.Command
         Com.ActiveConnection = sConex
         Com.CommandType = adCmdStoredProc
         Com.CommandText = "ValorAgregar"
         Com.Parameters.Append Com.CreateParameter("Venta", adDouble, adParamInput, , pRsValor!Venta)
         Com.Parameters.Append Com.CreateParameter("CodPago", adInteger, adParamInput, , pRsValor!CodPago)
         Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 100, pRsValor!formapago)
         Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre)
         Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsValor!Banco)
         Com.Parameters.Append Com.CreateParameter("IdCheque", adInteger, adParamInput, , 0)
         Com.Parameters.Append Com.CreateParameter("NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque)
         Com.Parameters.Append Com.CreateParameter("FAcredi", adDate, adParamInput, , pRsValor!FechaAcreditacion)
         Com.Parameters.Append Com.CreateParameter("NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta)
         Com.Parameters.Append Com.CreateParameter("FVencim", adDate, adParamInput, , pRsValor!fechavencimiento)
         Com.Parameters.Append Com.CreateParameter("Autoriz", adVarChar, adParamInput, 30, pRsValor!Autorizacion)
         Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsValor!Cupon)
         Com.Parameters.Append Com.CreateParameter("Importe", adDouble, adParamInput, , pRsValor!importe)
         Com.Parameters.Append Com.CreateParameter("PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta)
         Com.Execute
         pRsValor.MoveNext
  Loop
  End If
  Set Com = Nothing
  ' Actualizo las Cantidades Si esta habilitada la opcion Stock en el Comprobante de Venta
  If bRem = False Then
     If pRsCantidad.RecordCount <> 0 Then
        pRsCantidad.MoveFirst
        Do While Not pRsCantidad.EOF
           If pRsCantidad!Producto <> "" Then
              Set Com = New ADODB.Command
              Com.ActiveConnection = sConex
              Com.CommandType = adCmdStoredProc
              Com.CommandText = "CantidadesAgregar"
              Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
              Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsVenta!Movimiento)
              Com.Parameters.Append Com.CreateParameter("Estado", adTinyInt, adParamInput, , pRsCantidad!Estado)
              Com.Parameters.Append Com.CreateParameter("Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto)
              Com.Parameters.Append Com.CreateParameter("Deposito", adInteger, adParamInput, , pRsCantidad!Deposito)
              Com.Parameters.Append Com.CreateParameter("CantidadDebe", adDouble, adParamInput, , pRsCantidad!CantidadDebe)
              Com.Parameters.Append Com.CreateParameter("CantidadHaber", adDouble, adParamInput, , pRsCantidad!CantidadHaber)
              Com.Parameters.Append Com.CreateParameter("Anulado", adInteger, adParamInput, , 0)
              Com.Execute
           End If
           pRsCantidad.MoveNext
        Loop
        Set Com = Nothing
     End If
  Else
'     rRem.MoveFirst
'     Do While Not rRem.EOF
'        Dim sSql As String
'        sSql = "UPDATE CabComprobantes SET Factura=" & pRsVenta!Movimiento
'        sSql = sSql & " WHERE Movimiento=" & rRem!Movimiento
'        rRem.MoveNext
'        sConex.Execute sSql
'     Loop
  End If
  ' Actualizo la Caja si esta habilitada lo opcion Caja en el Comprobante de Venta
  Set Com = New ADODB.Command
     
  If pRsCaja.BOF = False And pRsCaja.EOF = False Then
     pRsCaja.MoveFirst
     Do While Not pRsCaja.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = sConex
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "MovimientoCajaAgregar"
        Com.Parameters.Append Com.CreateParameter("Suc", adTinyInt, adParamInput, , nSucursal)
        Com.Parameters.Append Com.CreateParameter("Movimiento", adDouble, adParamInput, , pRsCaja!Movimiento)
        Com.Parameters.Append Com.CreateParameter("Comprobante", adInteger, adParamInput, , pRsCaja!Comprobante)
        Com.Parameters.Append Com.CreateParameter("Numero", adDouble, adParamInput, , pRsCaja!Numero)
        Com.Parameters.Append Com.CreateParameter("Sucursal", adInteger, adParamInput, , pRsCaja!sucursal)
        Com.Parameters.Append Com.CreateParameter("Fecha", adDate, adParamInput, , pRsCaja!fecha)
        Com.Parameters.Append Com.CreateParameter("Hora", adDate, adParamInput, , pRsCaja!Hora)
        Com.Parameters.Append Com.CreateParameter("Caja", adInteger, adParamInput, , pRsCaja!Caja)
        Com.Parameters.Append Com.CreateParameter("FormaPago", adVarChar, adParamInput, 100, pRsCaja!formapago)
        Com.Parameters.Append Com.CreateParameter("Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre)
        Com.Parameters.Append Com.CreateParameter("Banco", adInteger, adParamInput, , pRsCaja!Banco)
        Com.Parameters.Append Com.CreateParameter("NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque)
        Com.Parameters.Append Com.CreateParameter("FechaAcreditacion", adDate, adParamInput, , pRsCaja!FechaAcreditacion)
        Com.Parameters.Append Com.CreateParameter("NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta)
        Com.Parameters.Append Com.CreateParameter("FechaVencimiento", adDate, adParamInput, , pRsCaja!fechavencimiento)
        Com.Parameters.Append Com.CreateParameter("Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion)
        Com.Parameters.Append Com.CreateParameter("Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon)
        Com.Parameters.Append Com.CreateParameter("Debe", adDouble, adParamInput, , pRsCaja!Debe)
        Com.Parameters.Append Com.CreateParameter("Haber", adDouble, adParamInput, , pRsCaja!Haber)
        Com.Parameters.Append Com.CreateParameter("Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta)
        Com.Execute
        pRsCaja.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  Set cRsl = New ClsLectura
  
'  If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & pRsVenta!Comprobante & " AND Usuario=" & nUsuario) = False Then
'     Set Com = New ADODB.Command
'     Com.ActiveConnection = sConex
'     Com.CommandType = adCmdStoredProc
'     Com.CommandText = "ComprobanteActualizarNumero"
'     Com.Parameters.Append Com.CreateParameter("Id", adInteger, adParamInput, , pRsVenta!Comprobante)
'     Com.Execute
'  End If
'  Set Com = Nothing


 ' Call cRse.GuardarEventoConDetalle(sConex, "Ventas", "Insert", pRsVenta, pRsDetalle)
 ' Call cRse.GuardarEventoConDetalle(sConex, "Ctacte", "Insert", pRsRecibo)
 ' Call cRse.GuardarEventoConDetalle(sConex, "Caja", "Insert", pRsCaja)
  
  Call cRse.GuardarEventoCompleto(sConex, "venta", "insert", pRsVenta, pRsDetalle, pRsCaja, pRsRecibo)

  sConex.CommitTrans
  sConex.Close
  Exit Sub
    
errHandler:
    sConex.RollbackTrans
    sConex.Close
    
    err.Raise err.Number
End Sub

