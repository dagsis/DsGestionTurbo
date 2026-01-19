VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form HerraExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   6060
   Begin VB.CommandButton CmdProducto 
      Caption         =   "Productos/Precios"
      Height          =   435
      Left            =   90
      TabIndex        =   2
      Top             =   1260
      Width           =   1650
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   5895
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   300
         Left            =   4335
         TabIndex        =   4
         Top             =   195
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         Format          =   17235969
         CurrentDate     =   39332
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   645
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   300
         Left            =   2115
         TabIndex        =   5
         Top             =   210
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         Format          =   17235969
         CurrentDate     =   39332
      End
      Begin VB.Label Label2 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   1440
         TabIndex        =   6
         Top             =   210
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta :"
         Height          =   210
         Left            =   3675
         TabIndex        =   3
         Top             =   195
         Width           =   600
      End
   End
End
Attribute VB_Name = "HerraExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdProducto_Click()
  Dim cRsl As ClsLectura, rs As Recordset, x As Integer
  Dim Rs1 As Recordset, Rs2 As Recordset, Rs3 As Recordset
  Dim cRu As ClsHerramientas
   
  Set cRu = New ClsHerramientas
  Set cRsl = New ClsLectura
  
  cRu.BorrarLasTablas
  
  Set Rs1 = cRu.RsVacio("Productos", "Id", "N")
  Set Rs2 = cRu.RsVacio("ProductoImpuesto", "Id", "N")
  Set Rs3 = cRu.RsVacio("Precios", "Id", "N")
  
  CmdProducto.Enabled = False
  
  Set rs = cRsl.TraerRsCondi("Productos", "Producto", "FechaAlta>='" & FechaAmericana(DtpDesde.Value) & "' And FechaAlta<='" & FechaAmericana(DTPHasta.Value) & "'")
  x = 1
  If rs.RecordCount <> 0 Then
     ProgressBar1.Max = rs.RecordCount
     Do While Not rs.EOF
        Rs1.AddNew
   '     Rs1!Id = 0
        Rs1!Producto = rs!Producto
        Rs1!Descripcion = rs!Descripcion
        Rs1!Descripcion1 = rs!Descripcion1
        Rs1!Tipo = rs!Tipo
        Rs1!Familia = rs!Familia
        Rs1!Marca = rs!Marca
        Rs1!FechaAlta = rs!FechaAlta
        Rs1!FechaBaja = rs!FechaBaja
        Rs1!Proveedor = rs!Proveedor
        Rs1!UMCompra = rs!UMCompra
        Rs1!UMVenta = rs!UMVenta
        Rs1!Deposito = rs!Deposito
        Rs1!Origen = rs!Origen
        Rs1!CantDespacho = rs!CantDespacho
        Rs1!CantVenta = rs!CantVenta
        Rs1!FechaCompra = rs!FechaCompra
        Rs1!PrecioCompra = rs!PrecioCompra
        Rs1!CantMinima = rs!CantMinima
        Rs1!CantMaxima = rs!CantMaxima
        Rs1!ActuaStock = rs!ActuaStock
        Rs1!CodCompu = rs!CodCompu
        Rs1!Fijo = rs!Fijo
        Rs1!TipoInterno = rs!TipoInterno
        Rs1!ValorInterno = rs!ValorInterno
        Rs1!Calcular = rs!Calcular
        Rs1.Update
        
        Rs2.AddNew
        Rs2!Producto = rs!Producto
        Rs2!Impuesto = cRsl.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & rs!Producto & "'")
        Rs2.Update

        Rs3.AddNew
        Rs3!Lista = cRsl.TraerValorDeUnCampo("Registros", "ListaPrecio", "")
        Rs3!Producto = rs!Producto
        Rs3!Precio = cRsl.TraerValorDeUnCampo("Precios", "Precio", "Lista=" & Rs3!Lista & " AND Producto='" & rs!Producto & "'")
        Rs3.Update

        ProgressBar1.Value = x
        x = x + 1
        rs.MoveNext
     Loop
  End If
  CmdProducto.Enabled = True

  Set cRsl = Nothing
End Sub

Private Sub Form_Load()
  On Error GoTo errores
  
  Us = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Sucursal.mdb"
  DtpDesde.Value = Date
  DTPHasta.Value = Date
  Exit Sub
errores:
  MsgBox Err.Description, vbCritical, "Atención"
  Unload Me
End Sub
