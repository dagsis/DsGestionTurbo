VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VentaInfVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Ventas"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   5685
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   180
      TabIndex        =   13
      Top             =   2985
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   330
      Left            =   4290
      TabIndex        =   8
      Top             =   3015
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ventas"
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
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5370
      Begin VB.Frame Frame3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   2175
         TabIndex        =   7
         Top             =   210
         Width           =   3030
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   1050
            TabIndex        =   12
            Top             =   1050
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   159055873
            CurrentDate     =   36996
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1050
            TabIndex        =   11
            Top             =   405
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   159055873
            CurrentDate     =   36996
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta :"
            Height          =   180
            Left            =   300
            TabIndex        =   10
            Top             =   1050
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "Desde :"
            Height          =   240
            Left            =   300
            TabIndex        =   9
            Top             =   405
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Grupos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   135
         TabIndex        =   1
         Top             =   210
         Width           =   1920
         Begin VB.OptionButton Option5 
            Caption         =   "Vendedores"
            Height          =   195
            Left            =   255
            TabIndex        =   6
            Top             =   1485
            Width           =   1230
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Caja"
            Height          =   195
            Left            =   255
            TabIndex        =   5
            Top             =   1095
            Width           =   1230
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Productos"
            Height          =   195
            Left            =   255
            TabIndex        =   4
            Top             =   1875
            Width           =   1230
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Comprobantes"
            Height          =   195
            Left            =   255
            TabIndex        =   3
            Top             =   705
            Width           =   1440
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   255
            TabIndex        =   2
            Top             =   315
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "VentaInfVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As Recordset

Private Sub Command1_Click()
  Command1.Enabled = False
  If nImpr = 18 Then
'     nLlamo = 1: fDesde = DTPicker1.Value: fHasta = DTPicker2.Value
'     FrmGrilla.Show 1

     Set Rs = New Recordset
     Rs.Fields.Append "Campo01", adVarChar, 50
     Rs.Fields.Append "Producto", adVarChar, 25
     Rs.Fields.Append "Descripcion", adVarChar, 100, adFldIsNullable
     Rs.Fields.Append "Fecha", adDate
     Rs.Fields.Append "Comprobante", adVarChar, 30
     Rs.Fields.Append "Numero", adInteger
     Rs.Fields.Append "Sucursal", adInteger
     Rs.Fields.Append "Cliente", adVarChar, 80
     Rs.Fields.Append "Vendedor", adVarChar, 30
     Rs.Fields.Append "Caja", adVarChar, 30
     Rs.Fields.Append "Cantidad", adDouble
     Rs.Fields.Append "PrecioUnitario", adCurrency
     Rs.Fields.Append "PrecioTotal", adCurrency
     Rs.Fields.Append "Impuesto", adCurrency
     Rs.Fields.Append "Codigo", adVarChar, 20
     Rs.Open
     Dim cCaja As ClsComprobantesL, Rx As Recordset, cRsl As ClsLectura, nIva As Integer, nCosto As Integer
     Dim j As Long
     
     Set cCaja = New ClsComprobantesL
     Set cRsl = New ClsLectura
     fDesde = VentaInfVentas.DTPicker1.Value
     fHasta = VentaInfVentas.DTPicker2.Value

     Set Rx = cCaja.TraerMoviVentasProduc(fDesde, fHasta)
     ProgressBar1.Min = 0
     ProgressBar1.Max = IIf(Rx.RecordCount = 0, 1, Rx.RecordCount)
     ProgressBar1.Value = 0
     Do While Not Rx.EOF
        If Not IsNull(Rx!Producto) Then
        Rs.AddNew
        Rs!Campo01 = Rx!Vendedor
        Rs!Producto = "" & Rx!Producto
        Rs!Descripcion = Left(Rx!Detalle, 50)
        Rs!fecha = Rx!fecha
        Rs!Comprobante = Rx!Comprobante
        Rs!Numero = Rx!Numero
        Rs!sucursal = Rx!sucursal
        Rs!Cliente = "" & Rx!RazonSocial
        Rs!Vendedor = Rx!Vendedor
        Rs!Caja = Rx!Caja
        Rs!Cantidad = Rx!Cantidad
        nIva = cRsl.TraerValorDeUnCampo("Comprobantes", "Iva", "Comprobante='" & Rx!Comprobante & "'")
        nCosto = cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Comprobante='" & Rx!Comprobante & "'")
        If nIva = 0 Or (nIva = 1 And nCosto = 0) Then
           Rs!Impuesto = Rx!Impuesto
           If Rx!PrecioTotal <> 0 Then
              Rs!PrecioUnitario = Rx!PrecioTotal / Rx!Cantidad
              Rs!PrecioUnitario = Rs!PrecioUnitario - (Rx!Descuento / Rx!Cantidad)
              Rs!PrecioTotal = IIf(Rx!tipo = 1, Rx!PrecioTotal - Rx!Descuento, -(Rx!PrecioTotal - Rx!Descuento))
           Else
              Rs!PrecioUnitario = 0
              Rs!PrecioTotal = 0
           End If
        Else
           Rs!Impuesto = Rx!Impuesto
           Rs!PrecioUnitario = (Rx!PrecioUnitario + (Rx!Impuesto / Rx!Cantidad)) '* Rx!Cantidad
           Rs!PrecioUnitario = Rs!PrecioUnitario - (Rx!Descuento / Rx!Cantidad)
           Rs!PrecioTotal = IIf(Rx!tipo = 1, Rx!PrecioTotal + Rx!Impuesto - Rx!Descuento, -(Rx!PrecioTotal + Rx!Impuesto - Rx!Descuento))
        End If
        Rs!codigo = Rx!Comprobante
        Rs.Update
        End If
        ProgressBar1.Value = j
        j = j + 1
        Rx.MoveNext
        DoEvents
     Loop
  End If
  Command1.Enabled = True
  
  If Not (Rs Is Nothing) Then
    If Not (Rs.EOF And Rs.BOF) Then
        FrmImpresor.Show
    Else
        MsgBox "No hay Registros. Intente con Nuevos Parametros", vbCritical, "Atención"
    End If
   Else
        MsgBox "No hay Registros. Intente con Nuevos Parametros", vbCritical, "Atención"
   End If
End Sub
Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub Form_Load()
  DTPicker1.Value = Date
  DTPicker2.Value = Date
End Sub
Private Sub Option1_Click()
  nImpr = 14
End Sub
Private Sub Option2_Click()
  nImpr = 15
End Sub
Private Sub Option3_Click()
  nImpr = 18
End Sub
Private Sub Option4_Click()
  nImpr = 17
End Sub
Private Sub Option5_Click()
  nImpr = 16
End Sub
