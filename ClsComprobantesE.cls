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

Private mConn As ADODB.Connection
Private mConnStr As String

'========================
'  CICLO DE VIDA
'========================
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
    If mConnStr = "" Then mConnStr = "FILE NAME=" & App.Path & "\Conex.udl"
    If mConn Is Nothing Then Set mConn = New ADODB.Connection
    If mConn.State = adStateClosed Then
        mConn.ConnectionTimeout = 5
        mConn.CommandTimeout = 30
        mConn.Open mConnStr
    End If
End Sub

'========================
'  TX
'========================
Private Sub BeginTx()
    EnsureConn
    mConn.BeginTrans
End Sub

Private Sub CommitTx()
    mConn.CommitTrans
End Sub

Private Sub RollbackTx()
    On Error Resume Next
    mConn.RollbackTrans
End Sub

'========================
'  COMMAND HELPERS
'========================
Private Function NewCmd(ByVal spName As String) As ADODB.Command
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = mConn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = spName
    Set NewCmd = cmd
End Function

Private Sub AddParam(ByRef cmd As ADODB.Command, _
                     ByVal pName As String, _
                     ByVal pType As DataTypeEnum, _
                     ByVal pDir As ParameterDirectionEnum, _
                     Optional ByVal pSize As Long = 0, _
                     Optional ByVal pValue As Variant)

    Dim p As ADODB.Parameter
    If pSize > 0 Then
        Set p = cmd.CreateParameter(pName, pType, pDir, pSize, pValue)
    Else
        Set p = cmd.CreateParameter(pName, pType, pDir, , pValue)
    End If
    cmd.Parameters.Append p
End Sub

'========================
'  RS => PARAMS (GENÉRICO)
'========================
' allowList: Dictionary con keys = nombres de campo permitidos (opcional)
' Si allowList = Nothing => mete todos (menos "Id")
Private Sub ExecRsFields(ByRef cmd As ADODB.Command, _
                         ByRef Rs As ADODB.Recordset, _
                         Optional ByVal allowList As Object)

    Dim i As Long
    Dim f As ADODB.Field
    Dim nm As String
    Dim v As Variant
    Dim t As DataTypeEnum
    Dim sz As Long

    For i = 0 To Rs.Fields.Count - 1
        Set f = Rs.Fields(i)
        nm = f.Name

        If UCase$(nm) <> "ID" Then
            If Not allowList Is Nothing Then
                If Not allowList.Exists(nm) Then GoTo nextField
            End If

            v = f.Value
            t = f.Type
            sz = f.DefinedSize

            ' Tu regla legacy: Double/Numeric -> adDouble
            If t = adDouble Or t = adNumeric Then t = adDouble

            ' Si es Null y es texto, mandá "" o " " según tu estándar (acá "")
            If IsNull(v) Then
                ' ojo: si algún SP no acepta NULL en ciertos campos, acá lo podés normalizar
                ' yo lo dejo tal cual Null por defecto
            End If

            If sz > 0 Then
                AddParam cmd, nm, t, adParamInput, sz, v
            Else
                AddParam cmd, nm, t, adParamInput, 0, v
            End If
        End If
nextField:
    Next i

    cmd.Execute
End Sub

' Helper rápido para allowlist
Private Function NewAllowList(ParamArray names()) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = LBound(names) To UBound(names)
        d(CStr(names(i))) = True
    Next
    Set NewAllowList = d
End Function

Public Sub Agregar( _
    pRsVenta As ADODB.Recordset, _
    pRsDetalle As ADODB.Recordset, _
    pRsRecibo As ADODB.Recordset, _
    pRsValor As ADODB.Recordset, _
    pRsCantidad As ADODB.Recordset, _
    pRsCaja As ADODB.Recordset)

    Dim cmd As ADODB.Command
    Dim okTx As Boolean
    okTx = False

    On Error GoTo errHandler

    BeginTx

    '=========================
    ' CABECERA
    '=========================
    Set cmd = NewCmd("CabezaAgregar")

    ' Si querés “blindarlo” para que NUNCA meta campos nuevos:
    'Dim wl As Object
    'Set wl = NewAllowList("Suc","Tipo","Movimiento","Comprobante","Fecha","FechaIva","Numero","Sucursal","Cliente","Lista","Vendedor","Caja","CondVenta","Neto","Iva1","Iva2","Descuentos","Financiacion","NoGravados","PIvaCompras","IBrutosCompras","Debe","Haber","Desde","Hasta","Factura","Motivo","Cotizacion","Anulado")
    'ExecRsFields cmd, pRsVenta, wl

    ' Si no querés allowlist por ahora:
    ExecRsFields cmd, pRsVenta, Nothing
    Set cmd = Nothing

    '=========================
    ' DETALLES (fijo, como tu legacy)
    '=========================
    If Not pRsDetalle Is Nothing Then
        If pRsDetalle.RecordCount > 0 Then
            pRsDetalle.MoveFirst
            Do While Not pRsDetalle.EOF
                Set cmd = NewCmd("DetallesAgregar")

                AddParam cmd, "Venta", adInteger, adParamInput, 0, pRsVenta!Movimiento
                AddParam cmd, "Producto", adVarChar, adParamInput, 25, pRsDetalle!Producto
                AddParam cmd, "Descripcion", adVarChar, adParamInput, 80, IIf(pRsDetalle!Descripcion = "", " ", pRsDetalle!Descripcion)
                AddParam cmd, "Leyenda", adVarChar, adParamInput, 13, IIf(IsNull(pRsDetalle!Leyenda) Or pRsDetalle!Leyenda = "", " ", pRsDetalle!Leyenda)
                AddParam cmd, "Cantidad", adSingle, adParamInput, 0, pRsDetalle!Cantidad
                AddParam cmd, "Cuenta", adInteger, adParamInput, 0, pRsDetalle!Cuenta
                AddParam cmd, "PDesc", adDouble, adParamInput, 0, IIf(IsNull(pRsDetalle!PDesc), 0, pRsDetalle!PDesc)
                AddParam cmd, "Descuento", adDouble, adParamInput, 0, pRsDetalle!Descuento
                AddParam cmd, "Tasa", adDouble, adParamInput, 0, pRsDetalle!tasa          'legacy OK
                AddParam cmd, "Deposito", adDouble, adParamInput, 0, pRsDetalle!Deposito
                AddParam cmd, "Medida", adDouble, adParamInput, 0, pRsDetalle!Medida
                AddParam cmd, "Impuesto", adDouble, adParamInput, 0, pRsDetalle!Impuesto
                AddParam cmd, "Precio", adDouble, adParamInput, 0, IIf(IsNull(pRsDetalle!precio), pRsDetalle!PrecioUnitario, pRsDetalle!precio)
                AddParam cmd, "PrecioUnitario", adDouble, adParamInput, 0, pRsDetalle!PrecioUnitario
                AddParam cmd, "PrecioTotal", adDouble, adParamInput, 0, pRsDetalle!PrecioTotal
                AddParam cmd, "MoviRemi", adInteger, adParamInput, 0, IIf(IsNull(pRsDetalle!MoviRemi), 0, pRsDetalle!MoviRemi)
                AddParam cmd, "DetRemi", adInteger, adParamInput, 0, IIf(IsNull(pRsDetalle!DetRemi), 0, pRsDetalle!DetRemi)

                cmd.Execute
                Set cmd = Nothing

                pRsDetalle.MoveNext
            Loop
        End If
    End If

    '=========================
    ' CTA CTE (GENÉRICO como tu legacy)
    '=========================
    If Not pRsRecibo Is Nothing Then
        If pRsRecibo.RecordCount <> 0 Then
            pRsRecibo.MoveFirst
            Do While Not pRsRecibo.EOF
                Set cmd = NewCmd("CtaCteAgregar")
                ExecRsFields cmd, pRsRecibo, Nothing
                Set cmd = Nothing
                pRsRecibo.MoveNext
            Loop
        End If
    End If

    '=========================
    ' VALORES (fijo como tu legacy)
    '=========================
    If Not pRsValor Is Nothing Then
        If pRsValor.RecordCount <> 0 Then
            pRsValor.MoveFirst
            Do While Not pRsValor.EOF
                Set cmd = NewCmd("ValorAgregar")

                AddParam cmd, "Venta", adDouble, adParamInput, 0, pRsValor!Venta
                AddParam cmd, "CodPago", adInteger, adParamInput, 0, pRsValor!CodPago
                AddParam cmd, "FormaPago", adVarChar, adParamInput, 30, pRsValor!formapago
                AddParam cmd, "Nombre", adVarChar, adParamInput, 50, pRsValor!Nombre
                AddParam cmd, "Banco", adInteger, adParamInput, 0, IIf(IsNull(pRsValor!Banco), 0, pRsValor!Banco)
                AddParam cmd, "IdCheque", adInteger, adParamInput, 0, IIf(IsNull(pRsValor!IdCheque), 0, pRsValor!IdCheque)
                AddParam cmd, "NCheque", adVarChar, adParamInput, 30, pRsValor!NumeroCheque
                AddParam cmd, "FAcredi", adVarChar, adParamInput, 10, pRsValor!FechaAcreditacion
                AddParam cmd, "NTarjet", adVarChar, adParamInput, 30, pRsValor!NumeroTarjeta
                AddParam cmd, "FVencim", adVarChar, adParamInput, 10, pRsValor!FechaVencimiento
                AddParam cmd, "Autoriz", adVarChar, adParamInput, 30, ""
                AddParam cmd, "Cupon", adVarChar, adParamInput, 30, ""
                AddParam cmd, "Importe", adDouble, adParamInput, 0, pRsValor!importe
                AddParam cmd, "PorVenta", adVarChar, adParamInput, 1, pRsValor!PorVenta

                cmd.Execute
                Set cmd = Nothing

                pRsValor.MoveNext
            Loop
        End If
    End If

    '=========================
    ' CANTIDADES (fijo)
    '=========================
    If Not pRsCantidad Is Nothing Then
        If pRsCantidad.RecordCount <> 0 Then
            pRsCantidad.MoveFirst
            Do While Not pRsCantidad.EOF
                If pRsCantidad!Producto <> "" Then
                    Set cmd = NewCmd("CantidadesAgregar")

                    AddParam cmd, "Suc", adTinyInt, adParamInput, 0, nSucursal
                    AddParam cmd, "Movimiento", adDouble, adParamInput, 0, pRsVenta!Movimiento
                    AddParam cmd, "Estado", adTinyInt, adParamInput, 0, pRsCantidad!Estado
                    AddParam cmd, "Producto", adVarChar, adParamInput, 25, pRsCantidad!Producto
                    AddParam cmd, "Deposito", adInteger, adParamInput, 0, pRsCantidad!Deposito
                    AddParam cmd, "CantidadDebe", adDouble, adParamInput, 0, pRsCantidad!CantidadDebe
                    AddParam cmd, "CantidadHaber", adDouble, adParamInput, 0, pRsCantidad!CantidadHaber
                    AddParam cmd, "Anulado", adInteger, adParamInput, 0, 0

                    cmd.Execute
                    Set cmd = Nothing
                End If
                pRsCantidad.MoveNext
            Loop
        End If
    End If

    '=========================
    ' CAJA (fijo)
    '=========================
    If Not pRsCaja Is Nothing Then
        If (pRsCaja.BOF = False And pRsCaja.EOF = False) Then
            pRsCaja.MoveFirst
            Do While Not pRsCaja.EOF
                Set cmd = NewCmd("MovimientoCajaAgregar")

                AddParam cmd, "Suc", adTinyInt, adParamInput, 0, nSucursal
                AddParam cmd, "Movimiento", adDouble, adParamInput, 0, pRsCaja!Movimiento
                AddParam cmd, "Comprobante", adInteger, adParamInput, 0, pRsCaja!Comprobante
                AddParam cmd, "Numero", adDouble, adParamInput, 0, pRsCaja!Numero
                AddParam cmd, "Sucursal", adInteger, adParamInput, 0, pRsCaja!sucursal
                AddParam cmd, "Fecha", adDate, adParamInput, 0, pRsCaja!fecha
                AddParam cmd, "Hora", adDate, adParamInput, 0, pRsCaja!Hora
                AddParam cmd, "Caja", adInteger, adParamInput, 0, pRsCaja!Caja
                AddParam cmd, "FormaPago", adVarChar, adParamInput, 30, pRsCaja!formapago
                AddParam cmd, "Nombre", adVarChar, adParamInput, 50, pRsCaja!Nombre
                AddParam cmd, "Banco", adInteger, adParamInput, 0, pRsCaja!Banco
                AddParam cmd, "NumeroCheque", adVarChar, adParamInput, 30, pRsCaja!NumeroCheque
                AddParam cmd, "FechaAcreditacion", adDate, adParamInput, 0, pRsCaja!FechaAcreditacion
                AddParam cmd, "NumeroTarjeta", adVarChar, adParamInput, 30, pRsCaja!NumeroTarjeta
                AddParam cmd, "FechaVencimiento", adDate, adParamInput, 0, IIf(IsNull(pRsCaja!FechaVencimiento), Date, pRsCaja!FechaVencimiento)
                AddParam cmd, "Autorizacion", adVarChar, adParamInput, 30, pRsCaja!Autorizacion
                AddParam cmd, "Cupon", adVarChar, adParamInput, 30, pRsCaja!Cupon
                AddParam cmd, "Debe", adDouble, adParamInput, 0, pRsCaja!Debe
                AddParam cmd, "Haber", adDouble, adParamInput, 0, pRsCaja!Haber
                AddParam cmd, "Abierta", adVarChar, adParamInput, 1, pRsCaja!Abierta

                cmd.Execute
                Set cmd = Nothing

                pRsCaja.MoveNext
            Loop
        End If
    End If

    CommitTx
    okTx = True
    Exit Sub

errHandler:
    On Error Resume Next
    If Not cmd Is Nothing Then Set cmd = Nothing
    If okTx = False Then RollbackTx
    err.Raise err.Number, "ClsComprobantesE.Agregar", err.Description
End Sub


Public Sub AnularComprobantes( _
    ByVal pMovimiento As Long, _
    ByVal pOpe As Integer, _
    ByVal pTipo As Integer, _
    Optional ByVal pFecha As String, _
    Optional ByVal pProv As String = "", _
    Optional ByVal pRsd As ADODB.Recordset)

    On Error GoTo errHandler
    EnsureConn

    mConn.BeginTrans

    Dim sSql As String

    Select Case pOpe

        Case 1 ' Venta
            mConn.Execute "DELETE FROM CabComprobantes WHERE Movimiento=" & pMovimiento
            'OJO: esta línea revisala (Lista=Movimiento huele mal)
            'mConn.Execute "DELETE FROM CabComprobantes WHERE Lista=" & pMovimiento

            mConn.Execute "DELETE FROM CuentaCorriente WHERE Venta=" & pMovimiento
            mConn.Execute "DELETE FROM Cantidades WHERE Movimiento=" & pMovimiento
            mConn.Execute "DELETE FROM Caja WHERE Movimiento=" & pMovimiento
            mConn.Execute "DELETE FROM Valores WHERE Venta=" & pMovimiento

            If Not pRsd Is Nothing Then
                If pRsd.State = adStateOpen And pRsd.RecordCount > 0 Then
                    pRsd.MoveFirst
                    Do While Not pRsd.EOF
                        mConn.Execute "UPDATE CabComprobantes SET Anulado=0 WHERE Movimiento=" & _
                                      CLng(IIf(IsNull(pRsd!MoviRemi), 0, pRsd!MoviRemi))
                        pRsd.MoveNext
                    Loop
                End If
            End If

            mConn.Execute "DELETE FROM DetallesComprobantes WHERE Movimiento=" & pMovimiento
            mConn.Execute "UPDATE CabComprobantes SET Factura=0 WHERE Factura=" & pMovimiento

        Case 2 ' Stock
            mConn.Execute "DELETE FROM CabComprobantes WHERE Movimiento=" & pMovimiento
            mConn.Execute "DELETE FROM DetallesComprobantes WHERE Movimiento=" & pMovimiento
            mConn.Execute "DELETE FROM Cantidades WHERE Movimiento=" & pMovimiento

        '... y así con el resto igual (Caja, Cobranza, etc.)

    End Select

    mConn.CommitTrans
    Exit Sub

errHandler:
    On Error Resume Next
    mConn.RollbackTrans
    err.Raise err.Number, "AnularComprobantes", err.Description
End Sub


