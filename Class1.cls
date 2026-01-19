VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DsPadronA4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const cCuit = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
              "<ConsultarCuit xmlns=""http://www.padrona5.storage.com.ar"">" & _
              "<pCuit>1</pCuit>" & _
            "</ConsultarCuit>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"
    
Private Const cAlicuota = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
              "<TraerAlicuotas xmlns=""http://www.padrona5.storage.com.ar"">" & _
              "<pCuit>1</pCuit>" & _
              "<pUser>1</pUser>" & _
              "<pPassword>1</pPassword>" & _
            "</TraerAlicuotas>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"


Private UrlStr As String
 
Private m_objDOMPeople As DOMDocument
Dim temp As String

Private pCuit As String
Private pNombre As String
Private pDireccion As String
Private pTipoPersona As String
Private pCodPostal As String
Private pProvincia As String
Private pLocalidad As String
Private pActividad As String
Private pCategoriaIva As String
Private pEstado As String
Private pInicio As String
Private pPercepcion As String
Private pRetencion As String
Private pResultado As String
Private pMensaje As String
   
Public Property Get Cuit() As String
    Cuit = pCuit
End Property

Public Property Let Cuit(ByVal vNewValue As String)
    pCuit = vNewValue
End Property
Public Property Get Nombre() As String
    Nombre = pNombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    pNombre = vNewValue
End Property

Public Property Get Direccion() As String
    Direccion = pDireccion
End Property

Public Property Let Direccion(ByVal vNewValue As String)
    pDireccion = vNewValue
End Property

Public Property Get TipoPersona() As String
    TipoPersona = pTipoPersona
End Property

Public Property Let TipoPersona(ByVal vNewValue As String)
    pTipoPersona = vNewValue
End Property

Public Property Get CodPostal() As String
    CodPostal = pCodPostal
End Property

Public Property Let CodPostal(ByVal vNewValue As String)
    pCodPostal = vNewValue
End Property

Public Property Get Provincia() As String
    Provincia = pProvincia
End Property

Public Property Let Provincia(ByVal vNewValue As String)
    pProvincia = vNewValue
End Property

Public Property Get Localidad() As String
    Localidad = pLocalidad
End Property

Public Property Let Localidad(ByVal vNewValue As String)
    pLocalidad = vNewValue
End Property

Public Property Get Actividad() As String
    Actividad = pActividad
End Property

Public Property Let Actividad(ByVal vNewValue As String)
    pActividad = vNewValue
End Property
  
Public Property Get CategoriaIva() As String
    CategoriaIva = pCategoriaIva
End Property

Public Property Let CategoriaIva(ByVal vNewValue As String)
    pCategoriaIva = vNewValue
End Property

Public Property Get Estado() As String
    Estado = pEstado
End Property

Public Property Let Estado(ByVal vNewValue As String)
    pEstado = vNewValue
End Property

Public Property Get Inicio() As String
    Inicio = pInicio
End Property

Public Property Let Inicio(ByVal vNewValue As String)
    pInicio = vNewValue
End Property

Public Property Get Resultado() As String
    Resultado = pResultado
End Property

Public Property Let Resultado(ByVal vNewValue As String)
    pResultado = vNewValue
End Property

Public Property Get Mensaje() As String
    Mensaje = pMensaje
End Property

Public Property Let Mensaje(ByVal vNewValue As String)
    pMensaje = vNewValue
End Property

Public Property Get Percepcion() As String
    Percepcion = pPercepcion
End Property

Public Property Let Percepcion(ByVal vNewValue As String)
    pPercepcion = vNewValue
End Property

Public Property Get Retencion() As String
    Retencion = pRetencion
End Property

Public Property Let Retencion(ByVal vNewValue As String)
    pRetencion = vNewValue
End Property
    
    
Public Function SolicitarCuit(pCuit As String) As Boolean
  On Error GoTo err
  Dim parser As DOMDocument
  Set parser = New DOMDocument
    ' cargar el código SOAP para FaC
  parser.loadXML cCuit
    '
    ' Indicar el parámetro a enviar
  parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuit/pCuit").text = pCuit
   
  UrlStr = "http://www.padrona5.storage.com.ar/padrona5.asmx"
     
   '  UrlStr = "http://localhost:1779/PadronA4.asmx"
 
    ' usar XMLHTTPRequest para enviar la información al servicio Web
  Dim oHttReq As XMLHTTPRequest
  Set oHttReq = New XMLHTTPRequest
    '
    ' Enviar el comando de forma síncrona (se espera a que se reciba la respuesta)
  oHttReq.Open "POST /padrona5.asmx HTTP/1.1", UrlStr, False
    ' las cabeceras a enviar al servicio Web
    ' (no incluir los dos puntos en el nombre de la cabecera)
  oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
  oHttReq.setRequestHeader "SOAPAction", "http://www.padrona5.storage.com.ar/ConsultarCuit"
    
    ' enviar el comando
  oHttReq.Send parser.xml
    '
    ' este será el texto recibido del servicio Web
   
  parser.loadXML oHttReq.responseText
  
  Debug.Print oHttReq.responseText
        
  Cuit = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Cuit").text
  Nombre = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Nombre").text
  Direccion = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Direccion").text
  TipoPersona = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/TipoPersona").text
  Actividad = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Actividad").text
  Estado = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Estado").text
  CategoriaIva = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/CategoriaIva").text
  CodPostal = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/CodPostal").text
  Inicio = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Inicio").text
  Localidad = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Localidad").text
  Provincia = parser.selectSingleNode("/soap:Envelope/soap:Body/ConsultarCuitResponse/ConsultarCuitResult/Provincia").text
  
  Resultado = "OK"
  SolicitarCuit = True
  
Exit Function
err:
  SolicitarCuit = False
  Resultado = err.Description
  
End Function

 Public Function SolicitarIB(pCuit As String) As Boolean
  On Error GoTo err
  Dim parser As DOMDocument
  Set parser = New DOMDocument
    ' cargar el código SOAP para FaC
  parser.loadXML cAlicuota
    '
    ' Indicar el parámetro a enviar

  parser.selectSingleNode("/soap:Envelope/soap:Body/TraerAlicuotas/pCuit").text = pCuit
  parser.selectSingleNode("/soap:Envelope/soap:Body/TraerAlicuotas/pUser").text = "30592132539"
  parser.selectSingleNode("/soap:Envelope/soap:Body/TraerAlicuotas/pPassword").text = "623581"

   
  UrlStr = "http://www.padrona5.storage.com.ar/ArbaIb.asmx"
     
   '  UrlStr = "http://localhost:1779/PadronA4.asmx"
 
    ' usar XMLHTTPRequest para enviar la información al servicio Web
  Dim oHttReq As XMLHTTPRequest
  Set oHttReq = New XMLHTTPRequest
    '
    ' Enviar el comando de forma síncrona (se espera a que se reciba la respuesta)
  oHttReq.Open "POST /ArbaIb.asmx HTTP/1.1", UrlStr, False
    ' las cabeceras a enviar al servicio Web
    ' (no incluir los dos puntos en el nombre de la cabecera)
  oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
  oHttReq.setRequestHeader "SOAPAction", "http://www.padrona5.storage.com.ar/TraerAlicuotas"
    
    ' enviar el comando
  oHttReq.Send parser.xml
    '
    ' este será el texto recibido del servicio Web
   
  parser.loadXML oHttReq.responseText
  
  Debug.Print oHttReq.responseText
        
  Percepcion = parser.selectSingleNode("/soap:Envelope/soap:Body/TraerAlicuotasResponse/TraerAlicuotasResult/Perception").text
  Retencion = parser.selectSingleNode("/soap:Envelope/soap:Body/TraerAlicuotasResponse/TraerAlicuotasResult/Retencion").text
  Resultado = parser.selectSingleNode("/soap:Envelope/soap:Body/TraerAlicuotasResponse/TraerAlicuotasResult/Resultado").text
  Mensaje = parser.selectSingleNode("/soap:Envelope/soap:Body/TraerAlicuotasResponse/TraerAlicuotasResult/Mensaje").text
  
  SolicitarIB = True

Exit Function
err:
  SolicitarIB = False
  Resultado = "Error"
  Mensaje = err.Description
  
End Function

