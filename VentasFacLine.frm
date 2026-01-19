VERSION 5.00
Begin VB.Form VentasFacLine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas On Line"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   7560
   Begin VB.CommandButton CmdEstado 
      Caption         =   "Consultar Estado del Servidor de la Afip"
      Height          =   390
      Left            =   135
      TabIndex        =   25
      Top             =   4110
      Width           =   3435
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   150
      TabIndex        =   6
      Top             =   1905
      Width           =   7245
      Begin VB.CommandButton CmdReparar 
         Caption         =   "Reparar"
         Height          =   270
         Left            =   3780
         TabIndex        =   4
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label LblVencimiento 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3255
         TabIndex        =   24
         Top             =   285
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label LblResultado 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1935
         TabIndex        =   23
         Top             =   1770
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "Resultado :"
         Height          =   255
         Left            =   195
         TabIndex        =   22
         Top             =   1770
         Width           =   1065
      End
      Begin VB.Label Label10 
         Caption         =   "Razón Social :"
         Height          =   255
         Left            =   225
         TabIndex        =   21
         Top             =   1398
         Width           =   1185
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1950
         TabIndex        =   20
         Top             =   1398
         Width           =   5175
      End
      Begin VB.Label LblNumDoc 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20170550189"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5850
         TabIndex        =   19
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label LblTipoDoc 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5850
         TabIndex        =   18
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label8 
         Caption         =   "N.Documento :"
         Height          =   270
         Left            =   4620
         TabIndex        =   17
         Top             =   675
         Width           =   1260
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Doc. :"
         Height          =   270
         Left            =   4620
         TabIndex        =   16
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label LblImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1950
         TabIndex        =   15
         Top             =   1027
         Width           =   1245
      End
      Begin VB.Label LblCae 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1950
         TabIndex        =   14
         Top             =   675
         Width           =   1620
      End
      Begin VB.Label LblFechaCompro 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1950
         TabIndex        =   13
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label6 
         Caption         =   "Importe Total :"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   1027
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Cae :"
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   675
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Comprobante :"
         Height          =   210
         Left            =   225
         TabIndex        =   10
         Top             =   285
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   7245
      Begin VB.TextBox TxtNumero 
         Height          =   300
         Left            =   1290
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "00000000"
         Top             =   1110
         Width           =   870
      End
      Begin VB.TextBox TxtSucursal 
         Height          =   315
         Left            =   1305
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "0000"
         Top             =   720
         Width           =   600
      End
      Begin VB.ComboBox CmbComprobante 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   2700
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "Consultar"
         Height          =   450
         Left            =   5685
         TabIndex        =   3
         Top             =   1275
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   "Número :"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   1110
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal :"
         Height          =   240
         Left            =   225
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobante :"
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   330
         Width           =   1155
      End
   End
End
Attribute VB_Name = "VentasFacLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdEstado_Click()
 Dim cRsl As ClsLectura
 Dim cRse As ClsEscritura

  Dim sCli As String
  
  On Error GoTo Errores

    Dim WSAA As Object, WSFEv1 As Object
    Dim ttl As Long, tra As String, Path As String, certificado As String, ClavePrivada As String
    Dim cms As String, cacert As String, wsdl As String
    Dim cache As String, wrapper As String, proxy As String, ok As Boolean, ta As String
    Dim tipo_cbte As Byte, punto_vta As Integer, cbte_nro As String, v As Variant, fecha As String
    Dim CAE As String, cae2 As String, token As String, Sign As String
    
    Set cRsl = New ClsLectura
    Set cRse = New ClsEscritura
      
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    If WSAA.Version < "2.02c" Then
       MsgBox "Debe instalar una versión más actualizada de PyAfipWs WSAA!"
       End
    End If
      
      
   ta = cRsl.TraerValorDeUnCampo("Sistema", "Ta", "")
   If ta <> "" Then
      ok = WSAA.AnalizarXml(ta)
       If Not WSAA.Expirado() Then
           token = WSAA.ObtenerTagXml("token")
           Sign = WSAA.ObtenerTagXml("sign")
       End If
   End If
            
   If token = "" Or Sign = "" Then
      
          ' deshabilito errores no manejados (version 2.04 o superior)
       WSAA.LanzarExcepciones = False
              
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
      ttl = 43200 ' tiempo de vida = 12hs hasta expiración
      tra = WSAA.CreateTRA("wsfe", ttl)
      ControlarExcepcion WSAA
      Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    
      Path = CurDir() + "\"
 
      certificado = sCertificado '"dagsis.crt" ' certificado de prueba
      ClavePrivada = sClavePrivada '"dagsis.key" ' clave privada de prueba
        
    ' Generar el mensaje firmado (CMS)
      cms = WSAA.SignTRA(tra, Path + certificado, Path + ClavePrivada)
      ControlarExcepcion WSAA
      Debug.Print cms
    
    ' Conectarse con el webservice de autenticación:
      cache = ""
      proxy = "" '"usuario:clave@localhost:8000"
      wrapper = ""
      cacert = ""
          
      wsdl = sWebServiceLogin
   
      ok = WSAA.Conectar(cache, wsdl, proxy, wrapper, cacert) ' Homologación
      ControlarExcepcion WSAA
    
    ' Llamar al web service para autenticar:
      ta = WSAA.LoginCMS(cms)
      ControlarExcepcion WSAA

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
      Debug.Print ta
      Debug.Print "Token:", WSAA.token
      Debug.Print "Sign:", WSAA.Sign
    
      If ta <> "" Then
           cRse.GrabarValordeUnCampo "Sistema", "Ta", "'" & ta & "'", ""
      End If
      token = WSAA.token
      Sign = WSAA.Sign
    End If
        
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    If WSAA.Version < "1.12" Then
        MsgBox "Debe instalar una versión mas actualizada de PyAfipWs WSFEv1!"
        End
    End If
    'Debug.Print WSFEv1.InstallDir

    ' Setear tocken y sing de autorización (pasos previos)
    WSFEv1.token = token
    WSFEv1.Sign = Sign

    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = sCuitAfip
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False

    wsdl = Trim(sWebServiceCae)

    ok = WSFEv1.Conectar(cache, wsdl, proxy, wrapper, cacert) ' homologación
    Debug.Print WSFEv1.Version
    ControlarExcepcion WSFEv1

    ' mostrar bitácora de depuración:
    Debug.Print WSFEv1.DebugLog

    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    
    WSFEv1.Dummy
    ControlarExcepcion WSFEv1
    
'    Dim x, ctz
'    For Each x In WSFEv1.ParamGetPtosVenta()
'        MsgBox x
'    Next
'
'    ctz = WSFEv1.ParamGetCotizacion("DOL")
'    MsgBox "Cotización Dólar: " & ctz
    
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

    MsgBox "Estado Servidor de Aplicación :" & WSFEv1.AppServerStatus & _
    Chr(10) + Chr(13) & "Estado Servidor Db : " & WSFEv1.DbServerStatus & _
    Chr(10) + Chr(13) & "Servidor de Autorización :" & WSFEv1.AuthServerStatus, vbInformation, "Atención"

    ' Buscar la factura
 Exit Sub
Errores:
   MsgBox "Espere Un Momento y Vuelva a Intentar..." & _
   Chr(10) + Chr(13) & err.Description, vbInformation, "Atención"
   Resume Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
     KeyAscii = Upper(KeyAscii)
End Sub

Private Sub CmdConsultar_Click()
  Dim cRsl As ClsLectura, cRse As ClsEscritura
  Dim sCli As String
  
  On Error GoTo Errores

  If CmbComprobante.ListIndex <> -1 Then

    Dim WSAA As Object, WSFEv1 As Object
    Dim ttl As Long, tra As String, Path As String, certificado As String, ClavePrivada As String
    Dim cms As String, cacert As String, wsdl As String
    Dim cache As String, wrapper As String, proxy As String, ok As Boolean, ta As String
    Dim tipo_cbte As Byte, punto_vta As Long, cbte_nro As String, v As Variant, fecha As String
    Dim CAE As String, cae2 As String, token As String, Sign As String
    
    LimpiarResultado
    
    Set cRsl = New ClsLectura
    Set cRse = New ClsEscritura
      
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    If WSAA.Version < "2.02c" Then
       MsgBox "Debe instalar una versión más actualizada de PyAfipWs WSAA!"
       End
    End If
      
      
   ta = cRsl.TraerValorDeUnCampo("Sistema", "Ta", "")
   If ta <> "" Then
      ok = WSAA.AnalizarXml(ta)
       If Not WSAA.Expirado() Then
           token = WSAA.ObtenerTagXml("token")
           Sign = WSAA.ObtenerTagXml("sign")
       End If
   End If
      
   If token = "" Or Sign = "" Then
          ' deshabilito errores no manejados (version 2.04 o superior)
       WSAA.LanzarExcepciones = False
              
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
      ttl = 43200 ' tiempo de vida = 12hs hasta expiración
      tra = WSAA.CreateTRA("wsfe", ttl)
      ControlarExcepcion WSAA
      Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    
      Path = CurDir() + "\"
 
      certificado = sCertificado '"dagsis.crt" ' certificado de prueba
      ClavePrivada = sClavePrivada '"dagsis.key" ' clave privada de prueba
        
    ' Generar el mensaje firmado (CMS)
      cms = WSAA.SignTRA(tra, Path + certificado, Path + ClavePrivada)
      ControlarExcepcion WSAA
      Debug.Print cms
    
    ' Conectarse con el webservice de autenticación:
      cache = ""
      proxy = "" '"usuario:clave@localhost:8000"
      wrapper = ""
      cacert = ""
          
      wsdl = sWebServiceLogin
   
      ok = WSAA.Conectar(cache, wsdl, proxy, wrapper, cacert) ' Homologación
      ControlarExcepcion WSAA
    
    ' Llamar al web service para autenticar:
      ta = WSAA.LoginCMS(cms)
      ControlarExcepcion WSAA

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
      Debug.Print ta
      Debug.Print "Token:", WSAA.token
      Debug.Print "Sign:", WSAA.Sign
    
      If ta <> "" Then
           cRse.GrabarValordeUnCampo "Sistema", "Ta", "'" & ta & "'", ""
      End If
      token = WSAA.token
      Sign = WSAA.Sign
      
    End If
        
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    If WSAA.Version < "1.12" Then
        MsgBox "Debe instalar una versión mas actualizada de PyAfipWs WSFEv1!"
        End
    End If
    'Debug.Print WSFEv1.InstallDir

    ' Setear tocken y sing de autorización (pasos previos)
    WSFEv1.token = token ' WSAA.Token
    WSFEv1.Sign = Sign ' WSAA.Sign

    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = sCuitAfip
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False

    ' Conectar al Servicio Web de Facturación
    proxy = "" ' "usuario:clave@localhost:8000"
    proxy = ""

    wsdl = Trim(sWebServiceCae)

   ' wsdl = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL"



    cache = "" 'Path
    wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    cacert = "" ' WSAA.InstallDir & "\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)"

    ok = WSFEv1.Conectar(cache, wsdl, proxy, wrapper, cacert) ' homologación
    Debug.Print WSFEv1.Version
    ControlarExcepcion WSFEv1

    ' mostrar bitácora de depuración:
    Debug.Print WSFEv1.DebugLog

    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    ControlarExcepcion WSFEv1
    
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

  

    ' Buscar la factura
    cae2 = WSFEv1.CompConsultar(CmbComprobante.ItemData(CmbComprobante.ListIndex), TxtSucursal.text, TxtNumero.text)
    ControlarExcepcion WSFEv1
   
    If WSFEv1.CAE <> "" Then
        LblFechaCompro.Caption = Right(WSFEv1.FechaCbte, 2) & "/" & Mid(WSFEv1.FechaCbte, 5, 2) & "/" & Left(WSFEv1.FechaCbte, 4)
        LblVencimiento.Caption = WSFEv1.Vencimiento
        LblImporte.Caption = Format(WSFEv1.ImpTotal, "0.00")
        LblCae.Caption = WSFEv1.CAE
        Select Case WSFEv1.Resultado
               Case "A"
                     LblResultado.Caption = "Autorizado"
               Case "O"
                     LblResultado.Caption = "Observado"
               Case "R"
                    LblResultado.Caption = "Rechazado"
        End Select
         
        Debug.Print "Fecha Comprobante:", WSFEv1.FechaCbte
        Debug.Print "Fecha Vencimiento CAE", WSFEv1.Vencimiento
        Debug.Print "Importe Total:", WSFEv1.ImpTotal
        Debug.Print "Resultado:", WSFEv1.Resultado
        Debug.Print "Cae:", WSFEv1.CAE
    
        If WSFEv1.Version >= "1.12a" Then
            ok = WSFEv1.AnalizarXml("XmlResponse")
            If ok Then
                Debug.Print "CAE:", WSFEv1.ObtenerTagXml("CodAutorizacion"), WSFEv1.CAE
                Debug.Print "CbteFch:", WSFEv1.ObtenerTagXml("CbteFch"), WSFEv1.FechaCbte
                Debug.Print "Moneda:", WSFEv1.ObtenerTagXml("MonId")
                Debug.Print "Cotizacion:", WSFEv1.ObtenerTagXml("MonCotiz")
                Debug.Print "DocTIpo:", WSFEv1.ObtenerTagXml("DocTipo")
                Debug.Print "DocNro:", WSFEv1.ObtenerTagXml("DocNro")
                LblTipoDoc.Caption = IIf(WSFEv1.ObtenerTagXml("DocTipo") = 96, "DNI", "CUIT")
                LblNumDoc.Caption = WSFEv1.ObtenerTagXml("DocNro")
                sCli = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cuit='" & LblNumDoc.Caption & "'")
                LblCliente.Caption = IIf(sCli = "0", "Consumidor Final", sCli)
            Else
                ' hubo error, muestro mensaje
                MsgBox WSFEv1.Excepcion, vbCritical, "Atención"
            End If
        End If
    Else
       MsgBox "No Existe El Comprobante Solicitado", vbCritical, "Atención"
    End If
  End If
 Exit Sub
Errores:
   MsgBox "Espere Un Momento y Vuelva a Intentar...", vbInformation, "Atención"
   Resume Next
End Sub

Private Sub CmdReparar_Click()
  Dim nMovi As Long, cRsl As ClsLectura, cRco As ClsComprobantesL, cRle As ClsEscritura
  Dim sStore As String

  Set cRsl = New ClsLectura
  Set cRco = New ClsComprobantesL
  
  nMovi = cRco.TraerMoviCae(CmbComprobante.ItemData(CmbComprobante.ListIndex), TxtNumero.text, TxtSucursal.text, 0)
  If nMovi = 0 Then
     If cRco.TraerMoviCae(CmbComprobante.ItemData(CmbComprobante.ListIndex), TxtNumero.text, TxtSucursal.text, 1) = 0 Then
        MsgBox "Factura Autorizada Pero Borrada del Sistema..." & Chr(10) & Chr(13) & "Debe Volver a Crearla Sin Autorizarla y Repetir Está Operación.", vbCritical, "Atención"
     Else
        Dim rCae As ADODB.Recordset
        Set rCae = New ADODB.Recordset
        
        Set cRle = New ClsEscritura
                
        rCae.Fields.Append "Movimiento", adInteger
        rCae.Fields.Append "Fecha", adVarChar, 10
        rCae.Fields.Append "CAE", adVarChar, 50
        rCae.Fields.Append "Documento", adVarChar, 20
        rCae.Fields.Append "Comprobante", adInteger
        rCae.Fields.Append "Numero", adInteger
        rCae.Open
        
        rCae.AddNew
        rCae!Movimiento = cRco.TraerMoviCae(CmbComprobante.ItemData(CmbComprobante.ListIndex), TxtNumero.text, TxtSucursal.text, 1)
        rCae!fecha = LblVencimiento.Caption
        rCae!CAE = LblCae.Caption
        rCae!Documento = LblNumDoc.Caption
        rCae!Comprobante = cRsl.TraerValorDeUnCampo("Comprobantes", "Id", "Afip=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
        rCae!Numero = TxtNumero.text
        rCae.Update
                
        cRle.Actualizar "AfipCaes_A", rCae
        MsgBox "Factura Reparada Con Exito", vbInformation, "Atención"
     End If
     Exit Sub
  Else
     MsgBox "Factura Autorizada y Correctamente Procesada en el Sistema", vbInformation, "Atención"
  End If
  
End Sub

Private Sub Form_Load()
  Limpiar
  LimpiarResultado
  CargarCombos
End Sub

Private Sub CargarCombos()
  Dim cRco As ClsComprobantesL, Rx As Recordset
  
  Set cRco = New ClsComprobantesL
  
  Set Rx = cRco.TraerFacturasOnLine(nUsuario)
  If Rx.RecordCount <> 0 Then
     Do While Not Rx.EOF
        CmbComprobante.AddItem Rx!Descripcion
        CmbComprobante.ItemData(CmbComprobante.NewIndex) = Rx!Afip
        Rx.MoveNext
     Loop
  End If
  
End Sub

Private Sub Limpiar()
  TxtSucursal.text = "0000"
  TxtNumero.text = "00000000"
End Sub

Private Sub LimpiarResultado()
  LblFechaCompro.Caption = ""
  LblCae.Caption = ""
  LblTipoDoc.Caption = ""
  LblNumDoc.Caption = ""
  LblImporte.Caption = "0.00"
  LblCliente.Caption = ""
  LblResultado.Caption = ""
End Sub

Private Sub TxtSucursal_GotFocus()
  TxtSucursal.SelStart = 0
  TxtSucursal.SelLength = Len(TxtSucursal.text)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSucursal_LostFocus()
  If TxtSucursal.text = "" Then TxtSucursal.text = 1
  TxtSucursal.text = Format(TxtSucursal.text, "0000")
End Sub

Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtNumero_LostFocus()
  If TxtNumero.text = "" Then TxtNumero.text = 0
  TxtNumero.text = Format(TxtNumero.text, "00000000")
End Sub

Sub ControlarExcepcion(obj As Object)
    Dim fd As String

    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0
    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
        fd = FreeFile
        Open "c:\excepcion.txt" For Append As fd
        Print #fd, obj.Excepcion
        Print #fd, obj.Traceback
        Print #fd, obj.XmlRequest
        Print #fd, obj.XmlResponse
        Close fd
        MsgBox "Servidor de AFIP Caido. Intentelo Mas tarde....", vbCritical, "Atención"
        End
    End If
End Sub
