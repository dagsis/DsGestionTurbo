VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CajaCtaCteBan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Cuenta Bancario"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   LinkTopic       =   "form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   5490
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   4425
      TabIndex        =   7
      Top             =   2955
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen de Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   165
      TabIndex        =   0
      Top             =   90
      Width           =   5205
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   330
         Left            =   1515
         TabIndex        =   6
         Top             =   1665
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53477377
         CurrentDate     =   38687
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   330
         Left            =   1515
         TabIndex        =   5
         Top             =   1125
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53477377
         CurrentDate     =   38687
      End
      Begin VB.ComboBox CmbCuenta 
         Height          =   315
         Left            =   375
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4605
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Fecha :"
         Height          =   270
         Left            =   390
         TabIndex        =   4
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Desde Fecha :"
         Height          =   255
         Left            =   390
         TabIndex        =   3
         Top             =   1155
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   390
         TabIndex        =   2
         Top             =   315
         Width           =   810
      End
   End
End
Attribute VB_Name = "CajaCtaCteBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim cRsl As ClsClienteL
Public Rs As ADODB.Recordset


Private Sub CmdAplicar_Click()
  If CmbCuenta.ListIndex <> -1 Then
     CrearRs
     TraerAnterior
     TraerSiguiente
     nImpr = 42
     FrmImpresor.Show
  End If
End Sub

Private Sub Form_Load()
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  CargarCombos
End Sub

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbCuenta, "CuentasBancarias", "Banco", "Descripcion", ""
  Set cRsl = New ClsLectura
End Sub

Private Sub CrearRs()
  
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Proveedor", adVarChar, 15
  Rs.Fields.Append "Razon", adVarChar, 50
  Rs.Fields.Append "Domicilio", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Postal", adVarChar, 8, adFldIsNullable
  Rs.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Iva", adVarChar, 25, adFldIsNullable
  Rs.Fields.Append "Cuit", adVarChar, 13, adFldIsNullable
  Rs.Fields.Append "Telefono", adVarChar, 20, adFldIsNullable
  Rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs.Fields.Append "Comprobante", adVarChar, 100, adFldIsNullable
  Rs.Fields.Append "Numero", adVarChar, 20, adFldIsNullable
  Rs.Fields.Append "Sucursal", adInteger, , adFldIsNullable
  Rs.Fields.Append "Cuota", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Debe", adCurrency
  Rs.Fields.Append "Haber", adCurrency
  Rs.Open
End Sub

Private Sub TraerAnterior()
  Dim RsAux As ADODB.Recordset, nBan As Integer, cRa As ClsLectura, RsAux2 As ADODB.Recordset, nBan1 As Integer
  Dim sDebe As Currency, sHaber As Currency, sBan As String
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsClienteL
  Set cRa = New ClsLectura
  
  sBan = cRa.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & CmbCuenta.ItemData(CmbCuenta.ListIndex))
  
  Set RsAux = cRsl.ResumenSaldoAnteriorP(sBan, sBan, DtpDesde.Value)
  
  Do While Not RsAux.EOF
     sDebe = 0
     sHaber = 0
     Rs.AddNew
     Rs!Proveedor = RsAux!Proveedor
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = RsAux!CodigoPostal
     Rs!Localidad = ""  ' RsAux!Ciudad
     Rs!Provincia = RsAux!Provincia
     Rs!Iva = RsAux!TipoIva
     Rs!Cuit = RsAux!Cuit
     Rs!Telefono = RsAux!Telefono
     Rs!fecha = DtpDesde.Value
     Rs!Comprobante = "SALDO ANTERIOR"
     Rs!Numero = Null
     Rs!Sucursal = Null
     Rs!Cuota = Null
     If cRa.TraerValorDeUnCampo("Proveedores", "CtaCte", "Proveedor='" & RsAux!Proveedor & "'") = True Then
        nBan = cRsl.TraerCodigoCuenta(RsAux!Proveedor)
        If nBan <> nBan1 Then
            nBan1 = nBan
            Set RsAux2 = cRsl.TraerInterPropiosAnte(nBan, DtpDesde.Value)
            If RsAux2.RecordCount <> 0 Then
               Do While Not RsAux2.EOF
               sDebe = sDebe + RsAux2!Debe
               sHaber = sHaber + RsAux2!Haber
               RsAux2.MoveNext
               Loop
            End If
        End If
     End If
     nBan = CmbCuenta.ItemData(CmbCuenta.ListIndex)
     Set RsAux2 = cRsl.TraerMovideBancosAnt(nBan, DtpDesde.Value)
     If RsAux2.RecordCount <> 0 Then
        sDebe = sDebe + RsAux2!SumaDeDebe
        sHaber = sHaber + RsAux2!SumaDeHaber
     End If
     
     Rs!Debe = RsAux!SumaDeDebe + sDebe
     Rs!Haber = RsAux!SumaDeHaber + sHaber
     Rs.Update
     RsAux.MoveNext
  Loop
    
  If RsAux.RecordCount = 0 Then
     nBan = CmbCuenta.ItemData(CmbCuenta.ListIndex)
     Set RsAux2 = cRsl.TraerInterPropiosAnte(nBan, DtpDesde.Value)
     If RsAux2.RecordCount <> 0 Then
'        sDebe = RsAux2!SumaDeDebe
'        sHaber = RsAux2!SumaDeHaber
        Rs.AddNew
        Rs!Proveedor = cRa.TraerValorDeUnCampo("Proveedores", "Proveedor", "Id=" & nBan)
        Rs!Razon = cRa.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & nBan)
        Rs!Domicilio = cRa.TraerValorDeUnCampo("Proveedores", "Domicilio", "Id=" & nBan)
        Rs!Postal = cRa.TraerValorDeUnCampo("Proveedores", "CodigoPostal", "Id=" & nBan)
        Rs!Localidad = Rs!Domicilio = cRa.TraerValorDeUnCampo("Proveedores", "Localidad", "Id=" & nBan)
        Rs!Provincia = cRa.TraerValorDeUnCampo("Provincias", "Descripcion", "Provincia=" & cRa.TraerValorDeUnCampo("Proveedores", "Provincia", "Id=" & nBan))
        Rs!Iva = cRa.TraerValorDeUnCampo("TipoIva", "Descripcion", "TipoIva=" & cRa.TraerValorDeUnCampo("Proveedores", "TipoIva", "Id=" & nBan))
        Rs!Cuit = cRa.TraerValorDeUnCampo("Proveedores", "Cuit", "Id=" & nBan)
        Rs!Telefono = cRa.TraerValorDeUnCampo("Proveedores", "Telefono", "Id=" & nBan)
        Rs!fecha = DtpDesde.Value
        Rs!Comprobante = "SALDO ANTERIOR"
        Rs!Numero = Null
        Rs!Sucursal = Null
        Rs!Cuota = Null
        nBan = CmbCuenta.ItemData(CmbCuenta.ListIndex)
'        Set RsAux2 = cRsl.TraerMovideBancosAnt(nBan, DTPDesde.Value)
        If RsAux2.RecordCount <> 0 Then
           sDebe = sDebe + RsAux2!Debe
           sHaber = sHaber + RsAux2!Haber
        End If
        Rs!Debe = sDebe
        Rs!Haber = sHaber
        Rs.Update
     End If
  End If
  Set cRa = Nothing

End Sub

Private Sub TraerSiguiente()
  Dim RsAux As ADODB.Recordset, nBan As Integer, cRa As ClsLectura, RsAux2 As ADODB.Recordset, nBan1 As Integer, sBan As String
  Set RsAux = New ADODB.Recordset
  
  Set cRa = New ClsLectura

  sBan = cRa.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & CmbCuenta.ItemData(CmbCuenta.ListIndex))

  Set RsAux = cRsl.ResumenDeCuentaB(sBan, sBan, DtpDesde.Value, DtpHasta.Value)
   
  nBan1 = 0
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Proveedor = RsAux!Proveedor
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = RsAux!CodigoPostal
     Rs!Localidad = RsAux!Ciudad
     Rs!Provincia = RsAux!Provincia
     Rs!Iva = RsAux!TipoIva
     Rs!Cuit = RsAux!Cuit
     Rs!Telefono = RsAux!Telefono
     Rs!fecha = RsAux!fecha
     Rs!Comprobante = RsAux!Descripcion
     Rs!Numero = RsAux!Numero
     Rs!Sucursal = Null
     Rs!Cuota = Null
     Rs!Debe = RsAux!Debe
     Rs!Haber = RsAux!Haber
     Rs.Update
     RsAux.MoveNext
  Loop
  
  sBan = cRa.TraerValorDeUnCampo("Proveedores", "Proveedor", "Id=" & CmbCuenta.ItemData(CmbCuenta.ListIndex))
'  nBan = cRa.TraerValorDeUnCampo("CuentasBancarias", "Id", "Banco=" & cRsl.TraerCodigoCuenta(sBan))
  nBan = cRa.TraerValorDeUnCampo("CuentasBancarias", "Id", "Banco=" & CmbCuenta.ItemData(CmbCuenta.ListIndex))
  
  
  Set RsAux2 = cRsl.TraerInterPropios(nBan, DtpDesde.Value, DtpHasta.Value)
  nBan = CmbCuenta.ItemData(CmbCuenta.ListIndex)


  Do While Not RsAux2.EOF
     Rs.AddNew
     Rs!Proveedor = cRa.TraerValorDeUnCampo("Proveedores", "Proveedor", "Id=" & nBan)
     Rs!Razon = cRa.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & nBan)
     Rs!Domicilio = cRa.TraerValorDeUnCampo("Proveedores", "Domicilio", "Id=" & nBan)
     Rs!Postal = cRa.TraerValorDeUnCampo("Proveedores", "CodigoPostal", "Id=" & nBan)
     Rs!Localidad = cRa.TraerValorDeUnCampo("Proveedores", "Localidad", "Id=" & nBan)
     Rs!Provincia = cRa.TraerValorDeUnCampo("Provincias", "Descripcion", "Provincia=" & cRa.TraerValorDeUnCampo("Proveedores", "Provincia", "Id=" & nBan))
     Rs!Iva = cRa.TraerValorDeUnCampo("TipoIva", "Descripcion", "TipoIva=" & cRa.TraerValorDeUnCampo("Proveedores", "TipoIva", "Id=" & nBan))
     Rs!Cuit = cRa.TraerValorDeUnCampo("Proveedores", "Cuit", "Id=" & nBan)
     Rs!Telefono = cRa.TraerValorDeUnCampo("Proveedores", "Telefono", "Id=" & nBan)
     Rs!fecha = RsAux2!FechaAcreditacion
     Select Case RsAux2!formapago
            Case "CHEQUES PROPIOS"
                  Rs!Comprobante = RsAux2!Nombre
            Case "INTERDEPOSITOS"
                  Rs!Comprobante = "INTER." & " - " & RsAux2!Nombre
            Case "CHEQUES PROPIOS"
                  Rs!Comprobante = RsAux2!Nombre
'            Case "DEBITO AUTOMATICO"
'                  Rs!Comprobante = "DEBITO AUTOMATICO"
            Case Else
                 Rs!Comprobante = RsAux2!formapago
     End Select
     Rs!Numero = IIf(RsAux2!NumeroCheque = "", 0, RsAux2!NumeroCheque)
     Rs!Sucursal = Null
     Rs!Cuota = Null
     Rs!Debe = RsAux2!Debe
     Rs!Haber = RsAux2!Haber
     Rs.Update
     RsAux2.MoveNext
  Loop
  
  Set RsAux2 = cRsl.TraerMovideBancos(nBan, DtpDesde.Value, DtpHasta.Value)
  Do While Not RsAux2.EOF
     Rs.AddNew
     Rs!Proveedor = cRa.TraerValorDeUnCampo("Proveedores", "Proveedor", "Id=" & nBan)
     Rs!Razon = cRa.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & nBan)
     Rs!Domicilio = cRa.TraerValorDeUnCampo("Proveedores", "Domicilio", "Id=" & nBan)
     Rs!Postal = cRa.TraerValorDeUnCampo("Proveedores", "CodigoPostal", "Id=" & nBan)
     Rs!Localidad = Rs!Domicilio = cRa.TraerValorDeUnCampo("Proveedores", "Localidad", "Id=" & nBan)
     Rs!Provincia = cRa.TraerValorDeUnCampo("Provincias", "Descripcion", "Provincia=" & cRa.TraerValorDeUnCampo("Proveedores", "Provincia", "Id=" & nBan))
     Rs!Iva = cRa.TraerValorDeUnCampo("TipoIva", "Descripcion", "TipoIva=" & cRa.TraerValorDeUnCampo("Proveedores", "TipoIva", "Id=" & nBan))
     Rs!Cuit = cRa.TraerValorDeUnCampo("Proveedores", "Cuit", "Id=" & nBan)
     Rs!Telefono = cRa.TraerValorDeUnCampo("Proveedores", "Telefono", "Id=" & nBan)
     Rs!fecha = RsAux2!fecha
     Rs!Comprobante = RsAux2!Descripcion
     Rs!Numero = RsAux2!Numero
     Rs!Sucursal = Null
     Rs!Cuota = Null
     Rs!Debe = RsAux2!SumaDeDebe
     Rs!Haber = RsAux2!SumaDeHaber
     Rs.Update
     RsAux2.MoveNext
  Loop

  If Rs.RecordCount <> 0 Then
     Rs.Sort = "Fecha"
  End If
  Set cRa = Nothing

End Sub



