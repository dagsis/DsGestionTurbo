VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CompraInfResumen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumén de Cuenta"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   6405
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5175
      TabIndex        =   4
      Top             =   2775
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   165
      TabIndex        =   5
      Top             =   165
      Width           =   5970
      Begin VB.CheckBox ChkSimpli 
         Caption         =   "Simplificado"
         Height          =   225
         Left            =   3210
         TabIndex        =   16
         Top             =   2265
         Width           =   1215
      End
      Begin VB.CheckBox ChkSaldo 
         Caption         =   "Sin Saldos Cero"
         Height          =   195
         Left            =   1485
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox ChkAbc 
         Caption         =   "Alfabetico"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4980
         TabIndex        =   7
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   255
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "TxtCliente"
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4980
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   255
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "TxtCliente"
         Top             =   600
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   0
         Left            =   975
         TabIndex        =   2
         Top             =   1815
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   708837377
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   1
         Left            =   3195
         TabIndex        =   3
         Top             =   1815
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   708837377
         CurrentDate     =   37021
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Proveedor :"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Proveedor :"
         Height          =   195
         Left            =   255
         TabIndex        =   12
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   165
         Left            =   2640
         TabIndex        =   11
         Top             =   1815
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   345
         TabIndex        =   10
         Top             =   1815
         Width           =   615
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1155
         TabIndex        =   9
         Top             =   1260
         Width           =   3705
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1155
         TabIndex        =   8
         Top             =   600
         Width           =   3705
      End
   End
End
Attribute VB_Name = "CompraInfResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsClienteL
Public Rs As ADODB.Recordset

Private Sub CmdAplicar_Click()
  If ChkSimpli.Value = 0 Then
     CrearRs
     TraerAnterior
     TraerSiguiente
     nImpr = 31
     FrmImpresor.Show
  Else
     CrearRsSimp
     TraerResumen
     nImpr = 59
     FrmImpresor.Show
  End If
End Sub

Private Sub CrearRsSimp()
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Cliente", adVarChar, 15
  Rs.Fields.Append "Razon", adVarChar, 50
  Rs.Fields.Append "Domicilio", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Telefono", adVarChar, 80, adFldIsNullable
  Rs.Fields.Append "Saldo", adCurrency
  Rs.Open

End Sub

Private Sub TraerResumen()
  Dim RsAux As ADODB.Recordset, cRsl As ClsClienteL
  
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsClienteL

  
  Set RsAux = cRsl.ResumenDeCuentaSimpliCompra(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value, False, 0)
   
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Cliente = RsAux!Proveedor
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Telefono = RsAux!Telefono
     Rs!saldo = RsAux!saldo
     Rs.Update
     RsAux.MoveNext
  Loop
  If ChkAbc.Value = 1 Then
     Rs.Sort = "Razon"
  End If

  Set cRsl = Nothing

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim crs As ClsLectura
  Set crs = New ClsLectura
  nDat = 0
  nBuscar = 2
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Proveedores", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ProveedorTraerUno", nDat, True)
     LinkearTexto Index
     If Index = 0 Then
        TxtCliente(1).text = Rs!Proveedor
        LblCliente(1).Caption = Rs!RazonSocial
     End If
  End If
  Set cRsl = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  TxtCliente(Index).text = Rs!Proveedor
  LblCliente(Index).Caption = Rs!RazonSocial
End Sub

Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  DtpFecha(0).Value = Date
  DtpFecha(1).Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub


Private Sub DtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
  Set cRsl = New ClsClienteL
  Set Rs = New ADODB.Recordset

  Limpiar
  CargarEtiquetas
  
End Sub

Private Sub CargarEtiquetas()
  Set Rs = cRsl.ProveedorAlfabetico
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Proveedor
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Proveedor
  LblCliente(1).Caption = Rs!RazonSocial
End Sub
Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
  CmdAplicar.Enabled = False
  LblCliente(Index).Caption = ""
  If TxtCliente(Index).text <> "" Then
     If cRsl.BuscarNombreProveedor(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cRsl.BuscarNombreProveedor(TxtCliente(Index).text)
        CmdAplicar.Enabled = True
     End If
  End If
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
  Rs.Fields.Append "Cuit", adVarChar, 15, adFldIsNullable
  Rs.Fields.Append "Telefono", adVarChar, 80, adFldIsNullable
  Rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs.Fields.Append "Comprobante", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Numero", adDouble, , adFldIsNullable
  Rs.Fields.Append "Sucursal", adInteger, , adFldIsNullable
  Rs.Fields.Append "Cuota", adInteger, , adFldIsNullable
  Rs.Fields.Append "Debe", adCurrency
  Rs.Fields.Append "Haber", adCurrency
  Rs.Open
End Sub

Private Sub TraerAnterior()
  Dim RsAux As ADODB.Recordset, nBan As Integer, cRa As ClsLectura, RsAux2 As ADODB.Recordset, nBan1 As Integer
  Dim sDebe As Currency, sHaber As Currency
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsClienteL
  Set cRa = New ClsLectura
  
  Set RsAux = cRsl.ResumenSaldoAnteriorP(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value)
  
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Proveedor = RsAux!Proveedor
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = RsAux!CodigoPostal
     Rs!Localidad = RsAux!Localidad
     Rs!Provincia = "" ' RsAux![Provincias.Descripcion]
     Rs!Iva = "" ' RsAux![TipoIva.Descripcion]
     Rs!Cuit = RsAux!Cuit
     Rs!Telefono = RsAux!Telefono
     Rs!fecha = DtpFecha(0).Value
     Rs!Comprobante = "SALDO ANTERIOR"
     Rs!Numero = Null
     Rs!Sucursal = Null
     Rs!Cuota = Null
     Rs!Debe = RsAux!SumaDeDebe
     Rs!Haber = RsAux!SumaDeHaber
     Rs.Update
     RsAux.MoveNext
  Loop
  Set cRa = Nothing
End Sub

Private Sub TraerSiguiente()
  Dim RsAux As ADODB.Recordset, nBan As Integer, cRa As ClsLectura, RsAux2 As ADODB.Recordset, nBan1 As Integer
  Set RsAux = New ADODB.Recordset
  
  Set cRa = New ClsLectura

  Set RsAux = cRsl.ResumenDeCuentaP(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value)
   
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Proveedor = RsAux!Proveedor
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = RsAux!CodigoPostal
     Rs!Localidad = RsAux!Localidad
     Rs!Provincia = RsAux!Provincia
     Rs!Iva = RsAux!TipoIva
     Rs!Cuit = RsAux!Cuit
     Rs!Telefono = RsAux!Telefono
     Rs!fecha = RsAux!fecha
     Rs!Comprobante = RsAux!Comprobante
     Rs!Numero = RsAux!Numero
     Rs!Sucursal = RsAux!Sucursal
     Rs!Cuota = RsAux!Cuota
     Rs!Debe = IIf(IsNull(RsAux!Debe), 0, RsAux!Debe)
     Rs!Haber = IIf(IsNull(RsAux!Haber), 0, RsAux!Haber)
     Rs.Update
     RsAux.MoveNext
  Loop
  
  If ChkAbc.Value = 1 Then
     Rs.Sort = "Razon"
  End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
     KeyAscii = Upper(KeyAscii)
End Sub
