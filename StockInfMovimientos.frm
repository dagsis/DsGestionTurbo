VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form StockInfMovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Stock"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   7080
   Begin VB.Frame Frame2 
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
      Height          =   2715
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   6810
      Begin VB.CheckBox ChkCedro 
         Caption         =   "Cedronar"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   2400
         Width           =   1110
      End
      Begin VB.OptionButton Option2 
         Caption         =   "C.Venta"
         Height          =   210
         Left            =   2685
         TabIndex        =   15
         Top             =   2400
         Width           =   900
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Index           =   1
         Left            =   5700
         TabIndex        =   8
         Top             =   1185
         Width           =   765
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Index           =   1
         Left            =   405
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "TxtCodigo"
         Top             =   1185
         Width           =   1470
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Index           =   0
         Left            =   5700
         TabIndex        =   4
         Top             =   630
         Width           =   765
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Index           =   0
         Left            =   420
         MaxLength       =   25
         TabIndex        =   2
         Text            =   "TxtCodigo"
         Top             =   630
         Width           =   1470
      End
      Begin VB.OptionButton Option1 
         Caption         =   "C.Compra"
         Height          =   210
         Left            =   1470
         TabIndex        =   1
         Top             =   2400
         Width           =   1080
      End
      Begin MSComCtl2.DTPicker DTPfecha 
         Height          =   315
         Index           =   0
         Left            =   1035
         TabIndex        =   5
         Top             =   1815
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709427201
         CurrentDate     =   36935
      End
      Begin MSComCtl2.DTPicker DTPfecha 
         Height          =   315
         Index           =   1
         Left            =   3195
         TabIndex        =   6
         Top             =   1815
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709427201
         CurrentDate     =   36935
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   225
         Index           =   1
         Left            =   390
         TabIndex        =   14
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   165
         Index           =   1
         Left            =   2550
         TabIndex        =   13
         Top             =   1815
         Width           =   645
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescripcion"
         Height          =   285
         Index           =   1
         Left            =   1950
         TabIndex        =   12
         Top             =   1185
         Width           =   3705
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescripcion"
         Height          =   285
         Index           =   0
         Left            =   1950
         TabIndex        =   11
         Top             =   630
         Width           =   3705
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   420
         TabIndex        =   10
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Desde :"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   9
         Top             =   405
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   5820
      TabIndex        =   7
      Top             =   2940
      Width           =   1050
   End
End
Attribute VB_Name = "StockInfMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As ADODB.Recordset
Dim cRsl As ClsLectura, cPro As ClsProductoL
Dim cDes As ClsProductoL

Private Sub CmdAplicar_Click()
  If ChkCedro.Value = 0 Then
     CrearRs
     TraerAnterior
     TraerSiguiente
     nImpr = 45
  Else
     CrearRs1
     TraerSiguiente1
     nImpr = 61
  End If
  FrmImpresor.Show

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtCodigo(Index) = RsP!Producto
     LblDescripcion(Index).Caption = RsP!Descripcion
     If Index = 0 Then
        TxtCodigo(1).text = RsP!Producto
        LblDescripcion(1).Caption = RsP!Descripcion
     End If
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub TxtCodigo_Change(Index As Integer)
  TodoBien (Index)
End Sub

Private Sub TxtCodigo_GotFocus(Index As Integer)
  TxtCodigo(Index).SelStart = 0
  TxtCodigo(Index).SelLength = Len(TxtCodigo(Index))
End Sub

Private Sub TxtCodigo_LostFocus(Index As Integer)
  Dim cDes As ClsProductoL
  
  
  Set cDes = New ClsProductoL
  
  LblDescripcion(Index).Caption = cDes.BuscarNombreProducto(TxtCodigo(Index).text)
  
  If LblDescripcion(Index).Caption = "" Then
     TxtCodigo(Index).text = ""
  Else
     TxtCodigo(1).text = TxtCodigo(Index).text
     LblDescripcion(1).Caption = LblDescripcion(Index).Caption
  End If
  
  Set cDes = Nothing
  
End Sub

Private Sub TodoBien(Index As Integer)
  CmdAplicar.Enabled = False
  If Len(TxtCodigo(Index)) <> 0 Then
     CmdAplicar.Enabled = True
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
  '   KeyAscii = Upper(KeyAscii)
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
  
  LblDescripcion(0).Caption = ""
  LblDescripcion(1).Caption = ""
End Sub

Private Sub Form_Load()
  Limpiar
End Sub

Private Sub CrearRs()
  
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs.Fields.Append "Comprobante", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Sucursal", adInteger, , adFldIsNullable
  Rs.Fields.Append "Numero", adDouble, , adFldIsNullable
  Rs.Fields.Append "Deposito", adVarChar, 20, adFldIsNullable
  Rs.Fields.Append "Producto", adVarChar, 25, adFldIsNullable
  Rs.Fields.Append "Descripcion", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Saldo", adCurrency
  Rs.Fields.Append "Razon", adVarChar, 100, adFldIsNullable
  Rs.Open
End Sub

Private Sub TraerAnterior()
  Dim cPr As ClsComprobantesL
  
  Dim RsAux As ADODB.Recordset
  Set RsAux = New ADODB.Recordset
  Set cPr = New ClsComprobantesL
  
  Set RsAux = cPr.TraerKardexAnterior(LblDescripcion(0).Caption, LblDescripcion(1).Caption, DtpFecha(0).Value)
  
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!fecha = Null
     Rs!Comprobante = "SALDO ANTERIOR"
     Rs!Sucursal = Null
     Rs!Numero = Null
     Rs!Deposito = Null
     Rs!Producto = RsAux!Producto
     Rs!Descripcion = RsAux!Descripcion
     Rs!saldo = RsAux!saldo
     Rs!Razon = Null
     Rs.Update
     RsAux.MoveNext
  Loop
  
End Sub

Private Sub TraerSiguiente()
  Dim cPr As ClsComprobantesL, cRx As ClsLectura
  
  
  Dim RsAux As ADODB.Recordset
  Set RsAux = New ADODB.Recordset
  Set cPr = New ClsComprobantesL
  Set cRx = New ClsLectura
  
  Set RsAux = cPr.TraerKardex(LblDescripcion(0).Caption, LblDescripcion(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value)
   
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!fecha = RsAux!fecha
     Rs!Comprobante = RsAux!ComprDesc
     Rs!Sucursal = RsAux!Sucursal
     Rs!Numero = RsAux!Numero
     Rs!Deposito = RsAux!Deposito
     Rs!Producto = RsAux!Producto
     Rs!Descripcion = RsAux!ProdDesc
     Rs!saldo = RsAux!saldo
     Rs!Razon = RsAux!Motivo
     Select Case RsAux!tipo
            Case 1
                 Rs!Razon = cRx.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & RsAux!Cliente & "'")
            Case 6
                 Rs!Razon = cRx.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & RsAux!Cliente & "'")
            Case 7
                 Rs!Razon = cRx.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Proveedor='" & RsAux!Cliente & "'")
            Case 8
                 Rs!Razon = cRx.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Proveedor='" & RsAux!Cliente & "'")
     End Select
     Rs.Update
     RsAux.MoveNext
  Loop
  
End Sub


Private Sub CrearRs1()
  
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Cliente", adVarChar, 15, adFldIsNullable
  Rs.Fields.Append "Razon", adVarChar, 100, adFldIsNullable
  Rs.Fields.Append "Cuit", adVarChar, 15, adFldIsNullable
  Rs.Fields.Append "Producto", adVarChar, 25, adFldIsNullable
  Rs.Fields.Append "Descripcion", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Cantidad", adSingle
  Rs.Open
End Sub

Private Sub TraerSiguiente1()
  Dim cPr As ClsComprobantesL, cRx As ClsLectura, nTipo As Byte
  
  
  Dim RsAux As ADODB.Recordset
  Set RsAux = New ADODB.Recordset
  Set cPr = New ClsComprobantesL
  Set cRx = New ClsLectura
  If Option1.Value = True Then nTipo = 0
  If Option2.Value = True Then nTipo = 1
  
  Set RsAux = cPr.TraerCedronar(TxtCodigo(0).text, TxtCodigo(1).text, DtpFecha(0).Value, DtpFecha(1).Value, nTipo)
   
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Cliente = RsAux!Cliente
     Rs!Razon = RsAux!RazonSocial
     Rs!Cuit = RsAux!Cuit
     Rs!Producto = RsAux!Producto
     Rs!Descripcion = RsAux!Descripcion
     Rs!Cantidad = RsAux!Cantidad
     Rs.Update
     RsAux.MoveNext
  Loop
  
End Sub
