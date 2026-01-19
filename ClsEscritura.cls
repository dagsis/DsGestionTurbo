VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEscritura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Actualizar(pStoreProc As String, Optional pRsParam As ADODB.Recordset, Optional Cod As Variant = 0, Optional sCamp As String)
  Dim iCampos As Long, sSql As String
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
     If VarType(Cod) Then
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
  db.CommitTrans
  db.Close
 Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    
    MsgBox err.Description
    If err.Number = -2147467259 Then
       MsgBox "Este Registro no se puede Borrar", vbCritical, "Integridad Referencial"
       Exit Sub
    End If
    err.Raise err.Number

End Sub

Public Sub ActualizarVen(pStoreProc As String, Optional pRsParam As ADODB.Recordset, Optional Cod As Variant = 0, Optional sCamp As String)
  Dim iCampos As Integer, sSql As String
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
  If VarType(Cod) Then
      Com.Parameters.Append Com.CreateParameter("Param", adInteger, adParamInput, , Cod)
  Else
     Com.Parameters.Append Com.CreateParameter("Param", adVarChar, adParamInput, Len(Cod), Cod)
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
  db.CommitTrans
  db.Close
 Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    
    MsgBox err.Description
    If err.Number = -2147467259 Then
       MsgBox "Este Registro no se puede Borrar", vbCritical, "Integridad Referencial"
       Exit Sub
    End If
    err.Raise err.Number

End Sub

Public Sub ActualizarPrecios(pLista As Integer, pRsPrecios As Recordset)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
  
  db.Open sDb
  db.BeginTrans

  ' Agregar en registros en la lista de Precios
  If pRsPrecios.RecordCount > 0 Then
     pRsPrecios.MoveFirst
     Do While Not pRsPrecios.EOF
        Set Com = New ADODB.Command
        Com.ActiveConnection = db
        Com.CommandType = adCmdStoredProc
        Com.CommandText = "PreciosAgregar"
        Com.Parameters.Append Com.CreateParameter("Lista", adInteger, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Producto", adChar, adParamInput, 15)
        Com.Parameters.Append Com.CreateParameter("Porcentaje", adSingle, adParamInput)
        Com.Parameters.Append Com.CreateParameter("Precio", adSingle, adParamInput)
        Com.Parameters("Lista").Value = pLista
        Com.Parameters("Producto").Value = pRsPrecios!Producto
        Com.Parameters("Porcentaje").Value = 0
        Com.Parameters("Precio").Value = 0
        Com.Execute
        pRsPrecios.MoveNext
     Loop
     Set Com = Nothing
  End If
  
  db.CommitTrans
  db.Close
 Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    MsgBox err.Description
    If err.Number = -2147467259 Then
       MsgBox "Este Registro no se puede Borrar", vbCritical, "Integridad Referencial"
       Exit Sub
    End If
    err.Raise err.Number

End Sub

Public Sub ActualizarRegistro(pUser As Long, pTabla() As Long, pByte() As Byte)
 
  Dim sSql As String
  On Error GoTo errHandler
  db.Open sDb
  db.BeginTrans
  
  sSql = "UPDATE Registros SET Banco=" & pTabla(0) & ","
  sSql = sSql & "Marca=" & pTabla(1) & ","
  sSql = sSql & "Medida=" & pTabla(2) & ","
  sSql = sSql & "Caja=" & pTabla(3) & ","
  sSql = sSql & "Provincia=" & pTabla(4) & ","
  sSql = sSql & "Vendedor=" & pTabla(5) & ","
  sSql = sSql & "Descuento=" & pTabla(6) & ","
  sSql = sSql & "Pais=" & pTabla(7) & ","
  sSql = sSql & "Zona=" & pTabla(8) & ","
  sSql = sSql & "Deposito=" & pTabla(9) & ","
  sSql = sSql & "Familia=" & pTabla(10) & ","
  sSql = sSql & "Impuesto=" & pTabla(11) & ","
  sSql = sSql & "ListaPrecio=" & pTabla(12) & ","
  sSql = sSql & "Localidad=" & pTabla(13) & ","
  sSql = sSql & "Cobro=" & pTabla(14) & ","
  sSql = sSql & "VentaConsFinalN=" & pTabla(23) & ","
  sSql = sSql & "VentaFactura=" & pTabla(15) & ","
  sSql = sSql & "VentaConsFinal=" & pTabla(16) & ","
  sSql = sSql & "ClienteConsFinal=" & pTabla(17) & ","
  sSql = sSql & "VentaResponsable=" & pTabla(24) & ","
  sSql = sSql & "VentaOtraCondicion=" & pTabla(25) & ","
  sSql = sSql & "TipoVenta=" & pByte(9) & ","
  sSql = sSql & "Proveedor=" & pTabla(18) & ","
  sSql = sSql & "Moneda=" & pTabla(19) & ","
  sSql = sSql & "Precio=" & pByte(0) & ","
  sSql = sSql & "Venta=" & pByte(1) & ","
  sSql = sSql & "NumeraContr=" & pByte(2) & ","
  sSql = sSql & "FormaCob=" & pTabla(20) & ","
  sSql = sSql & "Cobranza=" & pTabla(21) & ","
  sSql = sSql & "CPago=" & pTabla(22) & ","
  sSql = sSql & "ImprimeDirecto=" & pByte(3) & ","
  sSql = sSql & "ListaRes=" & pTabla(26) & ","
  sSql = sSql & "ListaDel=" & pTabla(27) & ","
  sSql = sSql & "CuentaRes=" & pTabla(28) & ","
  sSql = sSql & "CuentaDel=" & pTabla(29) & ","
  sSql = sSql & "CliProvAut=" & pByte(12) & ","
  sSql = sSql & "DescCaja=" & pByte(13) & ","
  sSql = sSql & "Redondea=" & pByte(14) & ","
  sSql = sSql & "AnulaPres=" & pByte(15) & " "

  sSql = sSql & "WHERE Usuario=" & pUser
  
  db.Execute sSql
  
  sSql = "UPDATE SISTEMA SET Pagare=" & pByte(4) & ",bPagare=" & pByte(5) & ",bRecibo=" & pByte(6) & ",bRapido=" & pByte(7) & ",Receta=" & pByte(8) & ","
  sSql = sSql & "MueveStock=" & pByte(10) & ",MueveCaja=" & pByte(11)
  db.Execute sSql
  
 If pByte(10) = 0 Then
    sSql = "DELETE FROM CANTIDADES WHERE Movimiento<>0"
    db.Execute sSql
    
    sSql = "UPDATE Comprobantes SET MoviStock='No' WHERE TipoOperacion=8 or TipoOperacion=12 Or tipoOperacion=1 Or TipoOperacion=5"
    db.Execute sSql
    
 Else
    sSql = "UPDATE Comprobantes SET MoviStock='Suma' WHERE TipoOperacion=8 or TipoOperacion=5"
    db.Execute sSql
 
    sSql = "UPDATE Comprobantes SET MoviStock='Resta' WHERE TipoOperacion=12 or TipoOperacion=1"
    db.Execute sSql

 End If
  
  If pByte(11) = 0 Then
     sSql = "UPDATE Comprobantes SET Caja='No' WHERE TipoOperacion=8 or TipoOperacion=12 Or tipoOperacion=1 Or TipoOperacion=5"
     db.Execute sSql

  Else
     sSql = "UPDATE Comprobantes SET Caja='Resta' WHERE TipoOperacion=8 or TipoOperacion=5"
     db.Execute sSql
 
     sSql = "UPDATE Comprobantes SET Caja='Suma' WHERE TipoOperacion=12 or TipoOperacion=1"
     db.Execute sSql

  End If
  
  
  db.CommitTrans
  db.Close
  Exit Sub
errHandler:
    MsgBox err.Description
    db.RollbackTrans
    db.Close
End Sub

Public Sub GrabarValordeUnCampo(pTabla As String, pCampo As String, pValor As String, pCondi As String)
  Dim db As Connection, sSql As String
  
  Set db = New Connection
  
  db.Open sDb
  If VarType(pValor) <> vbString Then
     sSql = "UPDATE " & pTabla & " SET " & pCampo & "='" & pValor & "' "
     If pCondi <> "" Then
        sSql = sSql & "WHERE " & pCondi
     End If
  Else
     sSql = "UPDATE " & pTabla & " SET " & pCampo & "=" & pValor
     If pCondi <> "" Then
        sSql = sSql & " WHERE " & pCondi
     End If
  End If
  db.Execute sSql
  db.Close
End Sub

Public Sub GuardarEventoConDetalle(con As Connection, tipo As String, operacion As String, rsPrincipal As ADODB.Recordset, Optional rsDetalle As Recordset)
    Dim sSql As String

    Dim compania As Integer
    Dim sucursal As Integer

    
    compania = CompaniaId
    sucursal = CompaniaSuc

    If nSocket = 0 Then
        Exit Sub
    End If

    If rsPrincipal.RecordCount = 0 Then
        Exit Sub
    End If

    Dim dictData As New Scripting.Dictionary
    Dim Campo As Field

    rsPrincipal.MoveFirst

 ' Agregale estos dos campos a la cabecera también, si querés
    dictData("Compania") = compania
    dictData("CompaniaSucursal") = sucursal

    ' Copiar campos de la cabecera
    For Each Campo In rsPrincipal.Fields
        dictData(Campo.Name) = ValorSeguro(Campo.Value)
    Next
   
    ' Si hay detalle, lo serializa como array de objetos JSON
    Dim detalles As String
    Dim jsonFinal As String
    If Not rsDetalle Is Nothing Then
        detalles = "["
        rsDetalle.MoveFirst
        Do While Not rsDetalle.EOF
            Dim dictDet As New Scripting.Dictionary
            For Each Campo In rsDetalle.Fields
                dictDet(Campo.Name) = ValorSeguro(Campo.Value)
            Next
            ' Agregar compañía y sucursal a cada detalle:
            dictDet("Compania") = compania
            dictDet("CompaniaSucursal") = sucursal
            detalles = detalles & EncodeJSON(dictDet) & ","
            rsDetalle.MoveNext
        Loop
        If Right(detalles, 1) = "," Then detalles = Left(detalles, Len(detalles) - 1)
        detalles = detalles & "]"
    End If

    ' Y el armado del JSON final:
    If detalles <> "" Then
        jsonFinal = "{""tipo"":""" & tipo & """,""operacion"":""" & operacion & """,""compania"":" & compania & ",""sucursal"":" & sucursal & ",""data"":" & EncodeJSON(dictData) & ",""detalle"":" & detalles & "}"
    Else
        jsonFinal = "{""tipo"":""" & tipo & """,""operacion"":""" & operacion & """,""compania"":" & compania & ",""sucursal"":" & sucursal & ",""data"":" & EncodeJSON(dictData) & "}"
    End If

    ' Insertar en la base
    Dim sql As String
    sql = "INSERT INTO eventos (Evento, Procesado) VALUES ('" & Replace(jsonFinal, "'", "''") & "', 0)"
    con.Execute sql

End Sub


Public Sub GuardarEventoCompleto(con As Connection, tipo As String, operacion As String, _
    rsPrincipal As ADODB.Recordset, _
    Optional rsDetalle As Recordset, _
    Optional RsCaja As Recordset, _
    Optional RsCtaCte As Recordset)

    Dim compania As Integer
    Dim sucursal As Integer
    
    compania = CompaniaId
    sucursal = CompaniaSuc
        
    If nSocket = 0 Then Exit Sub
    If rsPrincipal.RecordCount = 0 Then Exit Sub

    Dim dictCab As New Scripting.Dictionary
    Dim Campo As Field
    rsPrincipal.MoveFirst
    dictCab("Compania") = compania
    dictCab("CompaniaSucursal") = sucursal
    For Each Campo In rsPrincipal.Fields
        dictCab(Campo.Name) = LimpiarControlChars(Campo.Value)
    Next

    ' Serializar detalle, caja y ctacte
    Dim detallesJson As String: detallesJson = ""
    If Not rsDetalle Is Nothing Then
        If rsDetalle.RecordCount <> 0 Then
            detallesJson = JsonArrayFromRecordset(rsDetalle, compania, sucursal)
        End If
    End If

    Dim Caja As String: Caja = ""
    If Not RsCaja Is Nothing Then
        If RsCaja.RecordCount <> 0 Then
            Caja = JsonArrayFromRecordset(RsCaja, compania, sucursal)
        End If
    End If

    Dim ctacte As String: ctacte = ""
    If Not RsCtaCte Is Nothing Then
        If RsCtaCte.RecordCount <> 0 Then
            ctacte = JsonArrayFromRecordset(RsCtaCte, compania, sucursal)
        End If
    End If

    ' Construir el JSON final
    Dim jsonFinal As String
    jsonFinal = "{""tipo"":""" & tipo & """,""operacion"":""" & operacion & """,""compania"":" & compania & ",""sucursal"":" & sucursal & _
        ",""cabComprobante"":" & EncodeJSON(dictCab)
    If detallesJson <> "" Then jsonFinal = jsonFinal & ",""detalle"":" & detallesJson
    If Caja <> "" Then jsonFinal = jsonFinal & ",""caja"":" & Caja
    If ctacte <> "" Then jsonFinal = jsonFinal & ",""ctacte"":" & ctacte
    jsonFinal = jsonFinal & "}"

    ' Reemplazo para que los arrays no queden como string en el JSON
    If detallesJson <> "" Then jsonFinal = Replace(jsonFinal, """detalle"":""" & detallesJson & """", """detalle"":" & detallesJson)
    If Caja <> "" Then jsonFinal = Replace(jsonFinal, """caja"":""" & Caja & """", """caja"":" & Caja)
    If ctacte <> "" Then jsonFinal = Replace(jsonFinal, """ctacte"":""" & ctacte & """", """ctacte"":" & ctacte)

    ' Guardar en la base
    Dim sql As String
    sql = "INSERT INTO eventos (Evento, Procesado) VALUES ('" & Replace(jsonFinal, "'", "''") & "', 0)"
    con.Execute sql
End Sub


Public Function JsonArrayFromRecordset(Rs As Recordset, compania As Integer, sucursal As Integer) As String
    Dim arr As String
    arr = "["
    Rs.MoveFirst
    Do While Not Rs.EOF
        Dim dictDet As New Scripting.Dictionary
        Dim Campo As Field
        For Each Campo In Rs.Fields
            Dim Valor As String
            Valor = LimpiarControlChars(Campo.Value)
            Valor = Replace(Valor, """", "\""") ' Escapar comillas dobles
            dictDet(Campo.Name) = Valor
        Next
        dictDet("Compania") = compania
        dictDet("CompaniaSucursal") = sucursal
        arr = arr & EncodeJSON(dictDet) & ","
        Rs.MoveNext
    Loop
    If Right(arr, 1) = "," Then arr = Left(arr, Len(arr) - 1)
    arr = arr & "]"
    JsonArrayFromRecordset = arr
End Function

Public Function LimpiarControlChars(Valor As Variant) As String
    Dim i As Integer
    Dim Resultado As String

    If IsNull(Valor) Then
        LimpiarControlChars = ""
        Exit Function
    End If

    Resultado = ""
    For i = 1 To Len(CStr(Valor))
        Dim ascii As Integer
        ascii = Asc(Mid(CStr(Valor), i, 1))
        If ascii >= 32 Then
            Resultado = Resultado & Mid(CStr(Valor), i, 1)
        End If
    Next i
    LimpiarControlChars = Resultado
End Function

