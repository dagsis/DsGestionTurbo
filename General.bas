Attribute VB_Name = "General"
Option Explicit


Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
 
Type NUMBERFMT
NumDigits As Long ' número de dígitos decimales

LeadingZero As Long ' si hay ceros iniciales en los campos decimales

Grouping As Long ' tamaño del grupo a la izquierda del decimal

lpDecimalSep As String ' puntero a la cadena del separador de decimales

lpThousandSep As String ' puntero a la cadena del separador de miles

NegativeOrder As Long ' orden de números negativos

End Type
 
 
Declare Function GetNumberFormat Lib "kernel32" Alias "GetNumberFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As NUMBERFMT, ByVal lpNumberStr As String, ByVal cchNumber As Long) As Long
 
Public Const LOCAL_DEFAULT = &H2C0A
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_IDIGITS = &H11
Public Const LOCALE_STIMEFORMAT = &H1003
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SLONGDATE = &H20
Public Const LOCALE_SCURRENCY = &H14
Public Const LOCALE_SMONDECIMALSEP = &H16
Public Const LOCALE_SMONTHOUSANDSEP = &H17
 
Public Const FMT_FECHA_CORTA As String = "dd/MM/yyyy"
Public Const FMT_FECHA_LARGA As String = "dddd, d' de 'MMMM' de 'yyyy"
Public Const FMT_HORA As String = "HH:mm:ss"
Public Const SIMB_MONEDA As String = "$"
Public Const SEP_DEC As String = "."
Public Const SEP_MILES As String = ","

Public bPercibe As Boolean
Public bImprDirecto As Boolean
Public nCaja As Integer
Public nCondPago As Integer
Public nCons As Integer
Public nVentaFactura As Integer
Public sVentaCPago As String
Public sConsFinal As String
Public nBanco As Integer
Public nConcepto As Byte
Public sTa As String

Public nMedDetalle As Integer
Public nImpDetalle As Integer
Public nDepDetalle As Integer
Public nCueDetalle As Integer

Public pPase As String

Public db As ADODB.Connection
Public rUsu As ADODB.Recordset
Public rRsReg As Recordset

Public Us As String
Public id_datos As Byte
Public sCuitAfip As String
Public nConceptoAfip As Byte
Public sWebServiceLogin As String
Public sWebServiceCae As String
Public sCertificado As String
Public sClavePrivada As String
Public sTexFiscal() As String
Public sArchFartura As String
Public sLogoAfip As String
Public sAnchoTexto As Long
Public nDocAfip As Long
Public bMono As Boolean

Public nDeposito As Integer
Public nSucursal As Byte
Public nUsuario As Integer
Public sUser As String
Public bAdmin As Boolean
Public bProg As Boolean
Public sEmail As Byte
Public nPuerto As Integer
Public nScaner As Integer
Public sDb As String
Public db1 As Connection
Public nBuscar As Integer
Public bBuscarActivo As Boolean
Public nLlamo As Byte, fDesde As Date, fHasta As Date
Public nDat As Long
Public bVer As Boolean
Public nImpr As Integer
Public nLlama As Integer ' 1 Venta Comprobante,  2 Stock,3 Caja,4 Cobranzas 5 C.Rapida
Public pControladora As Byte ' Con este compatibilizo remitos anteriores Sistema campo Controladora
Public nRecibo As Byte
Public nFactura As Byte
Public nEmpre(0, 1, 1) As String
Public nMenu As Byte
Public sUdl As String
' Variables de Actualizacion de Combobox
Public bLocalidad As Boolean
Public bProvincia As Boolean
Public bPais As Boolean
Public bZona As Boolean
Public bLista As Boolean
Public nLista As Byte
Public nListaStock As Byte

Public bVendedor As Boolean
Public bFamilia As Boolean
Public bMarca As Boolean
Public bProveedores As Boolean
Public bUMCompra As Boolean
Public bUMVenta As Boolean
Public bDeposito As Boolean
Public bBanco As Boolean
Public bAbiertoVenta As Boolean
Public bTransferencia As Boolean
Public bPermiteStockNega As Boolean
Public pEdit As Boolean, pBorr As Boolean, pInf As Boolean, pVer As Boolean     ' Tablas
Public pEdit1 As Boolean, pBorr1 As Boolean, pInf1 As Boolean, pVer1 As Boolean ' Clientes
Public pEdit2 As Boolean, pBorr2 As Boolean, pInf2 As Boolean, pVer2 As Boolean ' Productos
Public pEdit3 As Boolean, pBorr3 As Boolean, pInf3 As Boolean, pVer3 As Boolean ' Proveedores
Public pVenta As String
Public sNomFan, sDirFan, sTelFan, sCiuFan As String, sPagare As String, bPagare As Boolean, bRecibo As Boolean
Public nDigito As Byte
Public pPlani As Byte
Public pEti As Byte
Public pClave As Boolean, bClave As Boolean, pClave1 As Boolean, bClave1 As Boolean
Public sClave As String, sClave1 As String, nDias As Byte
Public nPcaja As Integer, nPcompro As Integer, nMoviAde As Long, nCadena As Byte, nCodigo As Byte
Public nCantDecimales As String, nImpDecimales As String, nDecimalCant As String
Public bAgrupar As Boolean, bAgruRemito As Boolean, nFac As Long
Public bImprimeCtaCte As Boolean
Public nCatUser As Byte  ' 0 Programador 1 Administrador 2 Operador
Public nFormaPago As Byte ' 0 Aplica Automaticamente 2 Pide Forma de Pago
Public bContado As Boolean ' Es una operacion de Contado
Public bModoSuper As Boolean
Public sComputadora As String, nControlador As Byte, bRedondea As Boolean
Public bDescuento As Boolean, dPorcDescuento As Double, sFormaPago As String  ' Promociones
Public nValorRetencion As Double
Public nLllamaApli As Byte
Public nLineasElectronic As Byte
Public bOtroAbierto As Boolean
Public bAutoNumera As Boolean
Public bRedondeaLista As Boolean
Public CantiPuntos As Byte
Public bAnulaPres As Boolean
Public sEnvase As String
Public nEnvase As Byte
Public tEnvase As String
Public fBalanza As Byte
Public nSocket As Byte

' Variables Solamente del Resto
Public nMesa As Integer, bAbri As Boolean, nPedido As Long, bDelivery As Boolean
Public bCerrar As Boolean, bFormaNueva As Boolean, bCajaSalida As Boolean
Public bAnswer As Boolean, sCmd As String, sCmdExt As String, bNotaAsoc As Boolean, bAmericana As Boolean
Public bDecimalBalanza As Boolean
Public nCodeQr As Integer
Public CompaniaId As Integer
Public CompaniaSuc As Integer

'-- Funcion auxiliar que devuelve el nombre del equipo llamando al API
Public Function ComputerName() As String
  ComputerName = Space$(260)

 GetComputerName ComputerName, Len(ComputerName)
 ComputerName = Left$(ComputerName, InStr(ComputerName, vbNullChar) - 1)

End Function


Public Function Redondeo(ByVal Numero, ByVal Decimales)
      Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
End Function

Public Function Redondear(Numero As Double) As Double
  Dim a As Double
  a = Numero
   If a > Int(a) Then
       a = a + 1
       a = Int(a)
   End If
   Redondear = a
End Function

Public Function ConvierteFecha(pFecha As Date) As String
   Dim pDia As String
   Dim pMes As String
   Dim pYear As String

   pDia = CStr(Day(pFecha))
   If Len(pDia) = 1 Then
      pDia = "0" & pDia
   End If
   pMes = CStr(Month(pFecha))
   If Len(pMes) = 1 Then
      pMes = "0" & pMes
   End If
   pYear = CStr(Year(pFecha))

   ConvierteFecha = pYear & pMes & pDia
End Function

Public Function SoloNumero(key As Integer) As Boolean
   Select Case (key)
   Case 48 To 57, 8, 45, 32
      SoloNumero = True
   Case Else
      Beep
      SoloNumero = False
   End Select
End Function
 
  Public Function SoloNumeroDecimalFinal(key As Integer, pCaja As TextBox) As Boolean
   Dim i As Double, h As Byte, bPunto As Boolean
   Dim deci As String
    'me fijo si la conf. reg es punto decimal o coma
    i = 1.1
    
    If InStr(1, CDbl(i), ".") > 0 Then
        'el simbolo decimal es punto
        If key = 44 Then
            'SendKeys "."
            key = 46
        End If
        
    Else
        'el simbolo decimal es Coma
        If key = 46 Then
            'SendKeys ","
            key = 44
        End If
    End If
    If key = 46 Or key = 44 Then
       For h = 1 To Len(pCaja) + 1
           If (h > Len(pCaja.Text) And key = 46) Then
               If bPunto = True Then
                  SoloNumeroDecimalFinal = False
                  Exit Function
               Else
                  Exit For
               End If
           Else
             If IsNumeric(pCaja) Then
                If CDbl(pCaja.Text) = 0 Then
                   Exit For
                End If
             Else
             
             End If
           End If
           If Asc(Mid(pCaja.Text, h, 1)) = key Then
              If bPunto = True Then
                 SoloNumeroDecimalFinal = False
                 Exit Function
              Else
                 bPunto = True
              End If
           End If
       Next h
    End If
    Select Case (key)
    Case 48 To 57, 8, 45, 32, 46, 44
       SoloNumeroDecimalFinal = True
    Case Else
       Beep
       SoloNumeroDecimalFinal = False
    End Select
End Function
 
Public Function ValidaCuit(pCuit As String) As Boolean
   ' controla el digito verificador del cuit
   Dim v_cuit(1 To 11) As String
   Dim sCuit As String
   Dim i As Integer
   Dim j As Integer
   Dim lSuma As Long
   Dim iMult As Integer
   Dim iMod As Integer
   
   'sCuit = Mid(pCuit, 1, 2) & Mid(pCuit, 4, 8) & Mid(pCuit, 13, 1)
   sCuit = pCuit
   
   j = 1
   For i = 11 To 1 Step -1
      v_cuit(j) = Val(Mid(sCuit, i, 1))
      j = j + 1
   Next i
   
   For i = 1 To 11
      If i < 8 Then
         iMult = i
      Else
         iMult = i - 6
      End If
      lSuma = lSuma + v_cuit(i) * iMult
   Next i
   
   iMod = lSuma Mod 11
   If iMod <> 0 Then
      ValidaCuit = False
   Else
      ValidaCuit = True
   End If
   
End Function
 
Public Function FormatCurrencyUS(ByVal Monto As Currency) As String
    On Error Resume Next
    
    Dim sResult As String, MontoFix As Currency
    Dim sDecimals As String
    Dim sZeroFiller As String
    Dim sAux As String
    Dim lIndex As Long
    Dim lZeroCount As Long
    Dim sDecSep As String
    
    MontoFix = Fix(Monto)
    sDecimals = CStr(Abs((MontoFix - Monto) * (10 ^ (Len(CStr(Monto - MontoFix)) - 2))))
    If Val(sDecimals) > 0 Then
        sResult = CStr(MontoFix) & "." & String((Len(CStr(Monto - MontoFix)) - 2) - Len(sDecimals), "0") & sDecimals
    Else
        sResult = CStr(MontoFix) & ".00"
    End If
    FormatCurrencyUS = sResult
    
End Function
 
Public Sub ManejaErrores()
   Beep
   Screen.MousePointer = vbDefault
   MsgBox "Error : " & err.Number & Chr(10) & "Descripción : " & err.Description & Chr(10) & "Fuente: " & err.Source, vbInformation, "Errores"
End Sub


Public Function Upper(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 123 Then
    KeyAscii = KeyAscii - 32
 End If
 If KeyAscii = 39 Or KeyAscii = 34 Then Exit Function
 Upper = KeyAscii
End Function
Public Function FechaAmericana(fecha As Date) As String
   If bAmericana = False Then
       FechaAmericana = Day(fecha) & "/" & Month(fecha) & "/" & Year(fecha)
    Else
       FechaAmericana = Format(Month(fecha), "00") & "/" & Format(Day(fecha), "00") & "/" & Year(fecha)
    End If
End Function

Public Function FechaAmericanaIva(fecha As Date) As String
 '  FechaAmericanaIva = Format(Month(fecha), "00") & "/" & Format(Day(fecha), "00") & "/" & Year(fecha)

   FechaAmericanaIva = Month(fecha) & "/" & Day(fecha) & "/" & Year(fecha)

End Function

Public Function ConfiguracionArgentina() As Boolean
    Dim pCant As Double
    
    pCant = 1.1
    If InStr(1, CDbl(pCant), ".") > 0 Then
        'el simbolo decimal es el punto
        ConfiguracionArgentina = False
    Else
       ConfiguracionArgentina = True
    End If

End Function

 Public Function TraerCultura(ByVal pDato As String) As Double
    pDato = Replace(pDato, "$", "")
    pDato = Replace(pDato, ",", "")
    TraerCultura = CDbl(pDato)
   ' .Replace(".", "")
  '  TraerCultura = Convert.ToDecimal(pDato, MiCultura)
  End Function

Public Function ImprimirXLinea(qControl As TextBox) As String()

  Dim Lineas() As String

  Dim arrTem() As String
  Dim i As Byte

  Lineas = Split(qControl, vbNewLine)
  
  If UBound(Lineas) <> -1 Then
     ReDim arrTem(UBound(Lineas))
    
     If qControl = vbNullString Then Return
 
     For i = 0 To UBound(Lineas)
         arrTem(i) = Lineas(i)
     Next
     ImprimirXLinea = arrTem
     Erase arrTem
  End If
 
End Function

Public Function ImprimirXLineaMovi(qControl As String) As String()

  Dim Lineas() As String

  Dim arrTem() As String
  Dim i As Byte

  Lineas = Split(qControl, vbNewLine)
  
  If UBound(Lineas) <> -1 Then
     ReDim arrTem(UBound(Lineas))
    
     If qControl = vbNullString Then Return
 
     For i = 0 To UBound(Lineas)
         arrTem(i) = Lineas(i)
     Next
     ImprimirXLineaMovi = arrTem
     Erase arrTem
  End If
 
End Function

Public Sub LimpiarRecorset(pRecorset As Recordset)
  If pRecorset.RecordCount <> 0 Then
     pRecorset.MoveFirst
     While Not pRecorset.EOF
        pRecorset.Delete
        pRecorset.MoveNext
     Wend
     If pRecorset.RecordCount = 1 Then
        pRecorset.Delete
     End If
  End If
End Sub
 
Public Function CambiarCR(Optional strError As String) As Boolean
    Dim lngResu As Long
    Dim Buffer As String * 255
 
    On Error GoTo Errores
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SSHORTDATE, FMT_FECHA_CORTA)
    If lngResu = 0 Then strError = "Error al setear fecha corta."
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SLONGDATE, FMT_FECHA_LARGA)
    If lngResu = 0 Then strError = "Error al setear fecha larga."
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SDECIMAL, SEP_DEC)
    If lngResu = 0 Then strError = "Error al setear separador de decimales."
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_STHOUSAND, SEP_MILES)
    If lngResu = 0 Then strError = "Error al setear separador de miles."
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_STIMEFORMAT, FMT_HORA)
    If lngResu = 0 Then strError = "Error al setear formato de hora."
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONDECIMALSEP, SEP_DEC)
    If lngResu = 0 Then strError = "Error al setear separador de decimales de moneda."
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONTHOUSANDSEP, SEP_MILES)
    If lngResu = 0 Then strError = "Error al setear separador de miles de moneda."
 
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SCURRENCY, SIMB_MONEDA)
    If lngResu = 0 Then strError = "Error al setear símbolo de moneda."
 
    lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_SDECIMAL, Buffer, Len(Buffer))
    If Left$(Buffer, 1) = SEP_DEC Then
    lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONDECIMALSEP, Buffer, Len(Buffer))
    If Left$(Buffer, 1) = SEP_DEC Then
    lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_STHOUSAND, Buffer, Len(Buffer))
    If Left$(Buffer, 1) = SEP_MILES Then
    lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONTHOUSANDSEP, Buffer, Len(Buffer))
    If Left$(Buffer, 1) = SEP_MILES Then
    CambiarCR = (strError = vbNullString)
    End If
    End If
    End If
    End If
 
    Exit Function
Errores:
    CambiarCR = False
End Function
   
Public Function ObtenerPrimerUltimoDia(ByVal fecha As Date) As String()
    Dim sFec(1) As String

    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
    sFec(0) = DateSerial(Year(fecha), Month(fecha) + 0, 1)
    sFec(1) = DateSerial(Year(fecha), Month(fecha) + 1, 0)

   ObtenerPrimerUltimoDia = sFec()

  End Function


Public Function Enviar_Mail_CDO(SerVidor_SMTP As String, _
                             Para As String, _
                             De As String, _
                             Asunto As String, _
                             Mensaje As String, _
                             Optional Path_Adjunto As String, _
                             Optional Puerto As String = "25", _
                             Optional Usuario As String, _
                             Optional Password As String, _
                             Optional Usar_Autentificacion As Boolean = True, _
                             Optional Usar_SSL As Boolean = True) As String
           
   On Error GoTo Errores
    ' Variable de objeto Cdo.Message
    Dim Obj_Email As CDO.Message


    ' Crea un Nuevo objeto CDO.Message
    Set Obj_Email = New CDO.Message
      
    ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
      del servidor o su dirección IP )
    Obj_Email.Configuration.Fields(cdoSMTPServer) = SerVidor_SMTP
      
    Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2
      
    ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
      465 o  el puerto 587 ( este último me dio error )
      
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(Puerto)
  
      
    ' Indica el tipo de autentificación con el servidor de correo _
     El valor 0 no requiere autentificarse, el valor 1 es con autentificación
    Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
                "configuration/smtpauthenticate") = Abs(Usar_Autentificacion)
      
      
      
        ' Tiempo máximo de espera en segundos para la conexión
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 100
  
      
    ' Configura las opciones para el login en el SMTP
    If Usar_Autentificacion Then
  
    ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
     mas el @gmail.com )
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Usuario
  
    ' Password de la cuenta
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password
  
    ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Usar_SSL
      
    End If
      
  
    ' *********************************************************************************
    ' Estructura del mail
    '**********************************************************************************
      
    ' Dirección del Destinatario
    Obj_Email.To = Para
      
    ' Dirección del remitente
    Obj_Email.From = De
      
    ' Asunto del mensaje
    Obj_Email.Subject = Asunto
      
    ' Cuerpo del mensaje
    Obj_Email.TextBody = Mensaje
      
    'Ruta del archivo adjunto
      
    If Path_Adjunto <> vbNullString Then
        Obj_Email.AddAttachment (Path_Adjunto)
    End If
      
    ' Actualiza los datos antes de enviar
    Obj_Email.Configuration.Fields.Update
      
    On Error Resume Next
    ' Envía el email
    Obj_Email.Send
      
      
    If err.Number = 0 Then
       Enviar_Mail_CDO = "OK"
    Else
       MsgBox err.Description, vbCritical, " Error al enviar el amil "
    End If
      
    ' Descarga la referencia
    If Not Obj_Email Is Nothing Then
        Set Obj_Email = Nothing
    End If
      
 Exit Function
Errores:
 
    Enviar_Mail_CDO = err.Description
  
End Function

Public Function RedondearPrecios(pDato As Double) As Double
    RedondearPrecios = IIf(bRedondeaLista = True, Redondear(pDato), pDato)
End Function
  
Public Function EnviarBinarioForm(sUrl As String, sFileName As String, sUsuario As String, sPassword As String, sDisplay, sEnvia As String, Optional ByVal bAsync As Boolean) As String
    Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    Dim sPostData       As String
    Dim EnviarBinario   As Variant
    Dim body As String
    Dim cRsl As ClsLectura
    
    Set cRsl = New ClsLectura
    
    body = cRsl.TraerValorDeUnCampo("EmailDatos", "Asunto", "Id=1")
 
    '--- read file
    nFile = FreeFile
    Open sFileName For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
        sPostData = StrConv(baBuffer, vbUnicode)
    End If
    
    Close nFile
    
     '--- prepare body
      sPostData = "--" & STR_BOUNDARY & vbCrLf & _
      "Content-Disposition: form-data; name=""uploadfile""; filename=""" & Mid$(sFileName, InStrRev(sFileName, "\") + 1) & """" & vbCrLf & _
      "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
      sPostData & vbCrLf
       
     sPostData = sPostData & "--" & STR_BOUNDARY & vbCrLf
     sPostData = sPostData & "Content-Disposition: form-data; name=""envia""" & vbCrLf
     sPostData = sPostData & "Content-Type: text/plain" & vbCrLf & vbCrLf
     sPostData = sPostData & sEnvia & vbCrLf
     
     sPostData = sPostData & "--" & STR_BOUNDARY & vbCrLf
     sPostData = sPostData & "Content-Disposition: form-data; name=""usuario""" & vbCrLf
     sPostData = sPostData & "Content-Type: text/plain" & vbCrLf & vbCrLf
     sPostData = sPostData & sUsuario & vbCrLf
             
     sPostData = sPostData & "--" & STR_BOUNDARY & vbCrLf
     sPostData = sPostData & "Content-Disposition: form-data; name=""displayName""" & vbCrLf
     sPostData = sPostData & "Content-Type: text/plain" & vbCrLf & vbCrLf
     sPostData = sPostData & sDisplay & vbCrLf
     
      sPostData = sPostData & "--" & STR_BOUNDARY & vbCrLf
     sPostData = sPostData & "Content-Disposition: form-data; name=""password""" & vbCrLf
     sPostData = sPostData & "Content-Type: text/plain" & vbCrLf & vbCrLf
     sPostData = sPostData & sPassword & vbCrLf
     
     sPostData = sPostData & "--" & STR_BOUNDARY & vbCrLf
     sPostData = sPostData & "Content-Disposition: form-data; name=""asunto""" & vbCrLf
     sPostData = sPostData & "Content-Type: text/plain" & vbCrLf & vbCrLf
     sPostData = sPostData & "- Factura Electronica" & vbCrLf
    
     sPostData = sPostData & "--" & STR_BOUNDARY & vbCrLf
     sPostData = sPostData & "Content-Disposition: form-data; name=""htmlMessage""" & vbCrLf
     sPostData = sPostData & "Content-Type: text/plain" & vbCrLf & vbCrLf
     sPostData = sPostData & body & vbCrLf & _
     "--" & STR_BOUNDARY & "--"
    '--- post
      
       
       With CreateObject("Microsoft.XMLHTTP")
        .Open "POST", sUrl, bAsync
        .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        .Send pvToByteArray(sPostData)
        If Not bAsync Then
            EnviarBinario = .ResponseText
            If .Status = 200 Then
               EnviarBinarioForm = "OK"
            Else
               EnviarBinarioForm = EnviarBinario
            End If
        End If
    End With
End Function

Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function

Public Function EncodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML.DOMDocument
    Dim objNode As MSXML.IXMLDOMElement

    ' help from MSXML
    Set objXML = New MSXML.DOMDocument

    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text



    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing

End Function

Public Function EncodeJSON(dict As Scripting.Dictionary) As String
    Dim key As Variant, result As String
    result = "{"
    For Each key In dict.Keys
        result = result & """" & key & """" & ":" & """" & dict(key) & """" & ","
    Next
    If Right(result, 1) = "," Then result = Left(result, Len(result) - 1)
    result = result & "}"
    EncodeJSON = result
End Function

Public Function DecodeJSON(json As String) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    Dim clean As String, pairs() As String, i As Integer

    clean = Replace(json, "{", "")
    clean = Replace(clean, "}", "")
    clean = Replace(clean, """", "")
    pairs = Split(clean, ",")

    For i = LBound(pairs) To UBound(pairs)
        Dim pos As Integer
        pos = InStr(pairs(i), ":")
        If pos > 0 Then
            dict(Trim(Left(pairs(i), pos - 1))) = Trim(Mid(pairs(i), pos + 1))
        End If
    Next

    Set DecodeJSON = dict
End Function

Function ValorSeguro(valor As Variant) As String
    If IsNull(valor) Or IsEmpty(valor) Then
        ValorSeguro = "0"
    Else
        ValorSeguro = Replace(CStr(valor), """", "'")
    End If
End Function
