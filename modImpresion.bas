Attribute VB_Name = "modImpresion"
Option Explicit

Private mPrinting As Boolean

' ========= PUBLICO =========

Public Sub ImprimirNuevo( _
    ByVal Frm As Form, _
    ByRef RsCli As ADODB.Recordset, _
    ByRef RsComp As ADODB.Recordset, _
    ByRef Rsd As ADODB.Recordset, _
    ByVal nUsuario As Long, _
    ByVal sComputadora As String, _
    ByVal nMoviAfip As Long)

    On Error GoTo EH
    If mPrinting Then Exit Sub
    mPrinting = True

    If EsComprobanteAfip_Rs(RsComp) Then
        ' BLANCO: siempre Agiles/PDF
        ImprimirBlanco_Flow Frm, nUsuario, sComputadora, nMoviAfip, RsCli, RsComp, Rsd
    Else
        ' NEGRO: A4 o Comandera según Comanda
        ImprimirNegro_Flow Frm, nUsuario, sComputadora, RsComp
    End If

FIN:
    mPrinting = False
    Exit Sub
EH:
    mPrinting = False
    ManejaErrorImpresion "ImprimirNuevo", err
    Resume FIN
End Sub

' ========= FLUJO NEGRO =========

Private Sub ImprimirNegro_Flow( _
    ByVal Frm As Form, _
    ByVal nUsuario As Long, _
    ByVal sComputadora As String, _
    ByRef RsComp As ADODB.Recordset)

    On Error GoTo EH

    Dim bComanda As Boolean
    bComanda = CBool(RSVal(RsComp, "Comanda", False))

    ' Tu regla: Afip=0 y Comanda=True => comandera 80mm (si NO fuerza A4)
    If bComanda = True And Frm.ChkA4.Value = 0 Then
        ' >>>> ENCHUFAR TU RUTINA REAL
        ' Call ImprimirComandera80mm(Frm, nUsuario, sComputadora)
    Else
        ' >>>> ENCHUFAR TU RUTINA REAL
        ' Call ImprimirModeloA4(Frm, nUsuario, sComputadora)
    End If

    Exit Sub
EH:
    ManejaErrorImpresion "ImprimirNegro_Flow", err
End Sub

' ========= FLUJO BLANCO (CAE + PDF) =========

Private Sub ImprimirBlanco_Flow( _
    ByVal Frm As Form, _
    ByVal nUsuario As Long, _
    ByVal sComputadora As String, _
    ByVal nMoviAfip As Long, _
    ByRef RsCli As ADODB.Recordset, _
    ByRef RsComp As ADODB.Recordset, _
    ByRef Rsd As ADODB.Recordset)

    On Error GoTo EH

    ' 0) Datos del comprobante desde RsComp (NO DB)
    Dim nComproAfip As Integer, ok As Boolean
    nComproAfip = CInt(RSVal(RsComp, "Afip", 0))

    Dim nCosto As Byte
    nCosto = CByte(RSVal(RsComp, "Costo", 0))

    ' Punto de venta fiscal: tu regla:
    ' - Si es Afip>0: usar RsComp!Sucursal (punto de venta)
    ' - Si no: no aplica (pero igual lo usamos para numeración/armado de nombre)
    Dim pv As Integer
    pv = CInt(RSVal(RsComp, "Sucursal", Val(Frm.TxtSucursal.text)))
    If pv <= 0 Then pv = Val(Frm.TxtSucursal.text)

    ' 1) Documento receptor (desde RsCli)
    Dim tipoDocRec As Byte, nroDocRec As String
    PrepararDocumentoReceptor_Rs Frm, RsCli, tipoDocRec, nroDocRec

    ' 2) Calcular neto/iva UNA sola vez
    Dim nNeto As Double, nIva As Double
    CalcularNetoIva_Rsd Frm, Rsd, nComproAfip, nCosto, nNeto, nIva

    ' 3) CAE (idempotente) -> acá podés seguir consultando AfipCaes (está bien)
    Dim cRsl As ClsLectura
    Set cRsl = New ClsLectura

    If Not ExisteCAE(cRsl, nMoviAfip) Then
        Frm.CmdBotones(7).Enabled = False
             
        ok = Frm.Autorizar( _
                CByte(nComproAfip), nConcepto, tipoDocRec, nroDocRec, _
                CStr(Frm.LblTotal.Caption), CStr(Frm.LblNeto.Caption), CStr(Frm.LblIva1.text), _
                pv, nMoviAfip, nCosto)

        Frm.CmdBotones(7).Enabled = True
        If Not ok Then Exit Sub
    End If

    ' 4) PDF (siempre Agiles)
    Dim pdfPath As String
    pdfPath = BuildPdfPath_FromRs(Frm, RsComp, pv, nMoviAfip)

    EnsureFolderExists GetFolderName(pdfPath)

    Dim vIva(3) As String, vPorcIva(3) As Byte
    vIva(0) = 0: vIva(1) = 0: vIva(2) = 0: vIva(3) = 0
    vPorcIva(0) = 0: vPorcIva(1) = 4: vPorcIva(2) = 5: vPorcIva(3) = 6

    ' Localidad / Provincia / Domicilio salen de RsCli (sin DB)
    Dim sDomi As String, sLoc As String, sProv As String
    sDomi = CStr(RSVal(RsCli, "Domicilio", ""))
    sLoc = CStr(RSVal(RsCli, "CodigoPostal", "")) & " - " & _
           NzText(RSVal(RsCli, "Ciudad", ""), RSVal(RsCli, "Localidad", ""))

    ' soporta ambos alias
    sProv = NzText(RSVal(RsCli, "ProvinciaDesc", ""), RSVal(RsCli, "ProvDesc", ""))
    If sProv = "" Then sProv = CStr(RSVal(RsCli, "Provincia", ""))

    ' Llamada final (misma firma que ya tenías)
    Call Frm.ImprimirPdf( _
        CByte(nComproAfip), CByte(nConcepto), tipoDocRec, nroDocRec, _
        CStr(Frm.LblTotal.Caption), CStr(Frm.LblNeto.Caption), CStr(Frm.LblIva1.text), _
        vIva(), vPorcIva(), pv, nMoviAfip, _
        sDomi, sLoc, sProv, _
        nCosto)

    Exit Sub
EH:
    ManejaErrorImpresion "ImprimirBlanco_Flow", err
End Sub

' ========= HELPERS (RsComp/RsCli) =========

Private Function EsComprobanteAfip_Rs(ByRef RsComp As ADODB.Recordset) As Boolean
    EsComprobanteAfip_Rs = (CLng(RSVal(RsComp, "Afip", 0)) <> 0)
End Function

Private Sub PrepararDocumentoReceptor_Rs( _
    ByVal Frm As Form, _
    ByRef RsCli As ADODB.Recordset, _
    ByRef tipoDocRec As Byte, _
    ByRef nroDocRec As String)

    Dim tipoIva As Byte
    tipoIva = CByte(RSVal(RsCli, "TipoIva", 0))

    ' 80 CUIT / 96 DNI según tu regla
    tipoDocRec = 80
    If Frm.TxtCliente.text = "." Or tipoIva = 1 Then tipoDocRec = 96

    nroDocRec = Trim$(CStr(RSVal(RsCli, "Cuit", RSVal(RsCli, "CUIT", ""))))

    ' Si es "." y supera umbral, pedís doc
    If Frm.TxtCliente.text = "." And Val(Frm.LblTotal.Caption) > 10000000 Then
        FrmDocumento.Show 1
        If Frm.nDocAfip = 0 Then err.Raise vbObjectError + 500, "Documento", "Documento requerido"
        nroDocRec = CStr(Frm.nDocAfip)
        Exit Sub
    End If

    If nroDocRec = "" And Frm.TxtCliente.text <> "." Then
        err.Raise vbObjectError + 501, "Documento", "Contribuyente sin Número de CUIT"
    End If

    ' Consumidor final / DNI
    If tipoDocRec <> 80 Then
        tipoDocRec = 99
        nroDocRec = "000000000"
    End If
End Sub

Private Sub CalcularNetoIva_Rsd( _
    ByVal Frm As Form, _
    ByRef Rsd As ADODB.Recordset, _
    ByVal nComproAfip As Integer, _
    ByVal nCosto As Byte, _
    ByRef nNeto As Double, _
    ByRef nIva As Double)

    nIva = 0

    If Not (Rsd Is Nothing) Then
        If Not (Rsd.BOF And Rsd.EOF) Then
            Rsd.MoveFirst
            Do While Not Rsd.EOF
                nIva = nIva + CDbl(Rsd!Impuesto)
                Rsd.MoveNext
            Loop
        End If
    End If

    nNeto = CDbl(Frm.LblNeto.Caption) - CDbl(Frm.LblBonificacion.Caption)

    If nComproAfip >= 6 And nComproAfip <= 8 And nIva <> 0 Then
        nNeto = CDbl(Frm.LblTotal.Caption) - CDbl(Frm.TxtNoGrav.text) - CDbl(Frm.LblPercepcion.Caption) - nIva
    End If

    If nComproAfip >= 11 And nComproAfip <= 15 Then
        nNeto = CDbl(Frm.LblTotal.Caption)
        nIva = 0
    End If

    If nCosto = 0 Then
        nIva = 0
        nNeto = CDbl(Frm.LblTotal.Caption)
    End If
End Sub

Private Function ExisteCAE(ByVal cRsl As ClsLectura, ByVal movi As Long) As Boolean
    ExisteCAE = (CLng(cRsl.TraerCantidad("AfipCaes", "Id", "Movimiento=" & movi)) > 0)
End Function

Private Function BuildPdfPath_FromRs(ByVal Frm As Form, ByRef RsComp As ADODB.Recordset, ByVal pv As Integer, ByVal nMoviAfip As Long) As String
    Dim basePath As String
    basePath = App.Path & "\Facturas\"

    Dim sCob As String
    sCob = CStr(RSVal(RsComp, "Comprobante", "CBTE"))

    Dim nro As String
    nro = Format$(Val(Frm.TxtNumero.text), "00000000")

    Dim fileName As String
    fileName = SanitizeFileName(sCob & "-" & Format$(pv, "0000") & "-" & nro & "-M" & nMoviAfip & ".pdf")

    BuildPdfPath_FromRs = basePath & fileName
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, badChars(i), "_")
    Next
    If Len(s) > 180 Then s = Left$(s, 180)
    SanitizeFileName = s
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Function GetFolderName(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetFolderName = Left$(fullPath, p - 1)
    Else
        GetFolderName = App.Path
    End If
End Function

Private Sub ManejaErrorImpresion(ByVal origen As String, ByVal e As ErrObject)
    MsgBox origen & ": " & e.Description, vbCritical, "Atención"
End Sub

Private Function RSVal(ByRef rs As ADODB.Recordset, ByVal fieldName As String, Optional ByVal defValue As Variant) As Variant
    On Error GoTo EH
    If rs Is Nothing Then RSVal = defValue: Exit Function
    If (rs.BOF And rs.EOF) Then RSVal = defValue: Exit Function
    If rs.Fields.Count = 0 Then RSVal = defValue: Exit Function

    If IsNull(rs.Fields(fieldName).Value) Then
        RSVal = defValue
    Else
        RSVal = rs.Fields(fieldName).Value
    End If
    Exit Function
EH:
    RSVal = defValue
End Function

Private Function NzText(ByVal a As Variant, ByVal b As Variant) As String
    Dim sa As String, sb As String
    sa = Trim$(CStr(IIf(IsNull(a), "", a)))
    sb = Trim$(CStr(IIf(IsNull(b), "", b)))
    If sa <> "" Then
        NzText = sa
    Else
        NzText = sb
    End If
End Function

Private Function NzNum(ByVal v As Variant, ByVal def As Double) As Double
    On Error GoTo EH
    If IsNull(v) Or v = "" Then NzNum = def Else NzNum = CDbl(v)
    Exit Function
EH:
    NzNum = def
End Function


