VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form StockInfRemPendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos Pendientes"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   6210
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5130
      TabIndex        =   4
      Top             =   3045
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   105
      TabIndex        =   5
      Top             =   150
      Width           =   5970
      Begin VB.CheckBox ChkDetalle 
         Caption         =   "Remito Con Detalle"
         Height          =   270
         Left            =   255
         TabIndex        =   14
         Top             =   2415
         Width           =   2025
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
         Format          =   708968449
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
         Format          =   708968449
         CurrentDate     =   37021
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
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
Attribute VB_Name = "StockInfRemPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsClienteL
Public Rs As ADODB.Recordset, sw As Boolean
Private Sub CmdAplicar_Click()
  nImpr = 23
  If ChkDetalle.Value = 1 Then
     nImpr = 24
  End If
  FrmImpresor.Show
End Sub
Private Sub CmdBuscar_Click(Index As Integer)
  Dim crs As ClsLectura
  Set crs = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Clientes", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ClienteTraerUno", nDat, True)
     LinkearTexto Index
     If Index = 0 Then
        TxtCliente(1).text = Rs!Cliente
        LblCliente(1).Caption = Rs!RazonSocial
     End If
  End If
  Set cRsl = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  TxtCliente(Index).text = Rs!Cliente
  LblCliente(Index).Caption = Rs!RazonSocial
End Sub
Private Sub DtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Activate()
  CmdAplicar.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
   KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  DtpFecha(0).Value = Date - 30
  DtpFecha(1).Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub

Private Sub Form_Load()
  Set cRsl = New ClsClienteL
  Set Rs = New ADODB.Recordset

  Limpiar
  sw = False
  CargarEtiquetas
  
End Sub

Private Sub CargarEtiquetas()
  Set Rs = cRsl.ClientesAlfabetico(0, False)
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
End Sub

Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
  Set cRsl = New ClsClienteL
  CmdAplicar.Enabled = False
  LblCliente(Index).Caption = ""
  If TxtCliente(Index).text <> "" Then
     If cRsl.BuscarNombreCliente(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cRsl.BuscarNombreCliente(TxtCliente(Index).text)
        If sw = True And Index = 0 Then
           TxtCliente(1).text = TxtCliente(Index).text
           LblCliente(1).Caption = LblCliente(Index).Caption
        Else
           sw = True
        End If
        CmdAplicar.Enabled = True
     End If
  End If
End Sub

Private Sub CrearRs()
  
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Cliente", adVarChar, 15
  Rs.Fields.Append "Razon", adVarChar, 50
  Rs.Fields.Append "Domicilio", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Postal", adBoolean, , adFldIsNullable
  Rs.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Iva", adVarChar, 25, adFldIsNullable
  Rs.Fields.Append "Cuit", adVarChar, 13, adFldIsNullable
  Rs.Fields.Append "Telefono", adVarChar, 20, adFldIsNullable
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
  Dim RsAux As ADODB.Recordset
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsClienteL
  
  Set RsAux = cRsl.ResumenSaldoAnterior(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value)
  
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Cliente = RsAux!Cliente
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = RsAux!CodigoPostal
     Rs!Localidad = RsAux!Ciudad
     Rs!Provincia = RsAux!Provincia
     Rs!Iva = RsAux![TipoIva.Descripcion]
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
  
End Sub

Private Sub TraerSiguiente()
  Dim RsAux As ADODB.Recordset
  Set RsAux = New ADODB.Recordset
  
  Set RsAux = cRsl.ResumenDeCuenta(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value)
   
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Cliente = RsAux!Cliente
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = RsAux!CodigoPostal
     Rs!Localidad = RsAux!Ciudad
     Rs!Provincia = RsAux!Provincia
     Rs!Iva = RsAux!TipoIva
     Rs!Cuit = RsAux!Cuit
     Rs!Telefono = RsAux!Telefono
     Rs!fecha = RsAux!fecha
     Rs!Comprobante = RsAux!Comprobante
     Rs!Numero = RsAux!Numero
     Rs!Sucursal = RsAux!Sucursal
     Rs!Cuota = RsAux!Cuota
     Rs!Debe = RsAux!Debe
     Rs!Haber = RsAux!Haber
     Rs.Update
     RsAux.MoveNext
  Loop

End Sub

