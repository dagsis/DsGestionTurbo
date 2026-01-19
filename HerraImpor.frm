VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form HerraImpor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1740
   ScaleWidth      =   6120
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   75
      TabIndex        =   1
      Top             =   45
      Width           =   5895
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   300
         Left            =   4335
         TabIndex        =   2
         Top             =   195
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         Format          =   55640065
         CurrentDate     =   39332
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   645
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   210
         Left            =   3555
         TabIndex        =   4
         Top             =   195
         Width           =   675
      End
   End
   Begin VB.CommandButton CmdProducto 
      Caption         =   "Productos/Precios"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1650
   End
End
Attribute VB_Name = "HerraImpor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdProducto_Click()
  Dim cRsl As ClsLectura, rs As Recordset, X As Integer
  Dim Rs1 As Recordset, Rs2 As Recordset, Rs3 As Recordset, Rs4 As Recordset
  Dim cRu As ClsHerramientas, nLista As Integer, nPrecio As Currency
   
  Set cRu = New ClsHerramientas
  Set cRsl = New ClsLectura
  
  
  Set Rs1 = cRsl.RsVacioD("Productos", "Id", "N")
  Set Rs2 = cRsl.RsVacioD("ProductoImpuesto", "Id", "N")
  Set Rs3 = cRsl.RsVacioD("Precios", "Id", "N")
  Set Rs4 = cRsl.RsVacioD("ProductoUnidades", "Id", "N")
  
  CmdProducto.Enabled = False
  
  Set rs = cRu.TraerTodos("Productos", "*")
  X = 1
  Dim nProv As Integer, nMedi As Integer, nDep As Integer, nOri As Integer, nFam As Integer, nMar As Integer
  
  nLista = cRsl.TraerValorDeUnCampo("Registros", "ListaPrecio", "")
  nProv = cRsl.TraerValorDeUnCampo("Registros", "Proveedor", "")
  nMedi = cRsl.TraerValorDeUnCampo("Registros", "Medida", "")
  nDep = cRsl.TraerValorDeUnCampo("Registros", "Deposito", "")
  nOri = cRsl.TraerValorDeUnCampo("Registros", "Pais", "")
  nFam = cRsl.TraerValorDeUnCampo("Registros", "Familia", "")
  nMar = cRsl.TraerValorDeUnCampo("Registros", "Marca", "")
  
  If rs.RecordCount <> 0 Then
     ProgressBar1.Max = rs.RecordCount
     Do While Not rs.EOF
        If cRsl.TraerCantidad("Productos", "Id", "Producto='" & rs!Producto & "'") = 0 Then
           Rs1.AddNew
           Rs1!Producto = rs!Producto
           Rs1!Descripcion = rs!Descripcion
           Rs1!Descripcion1 = rs!Descripcion1
           Rs1!Tipo = rs!Tipo
           Rs1!Familia = nFam
           Rs1!Marca = nMar
           Rs1!FechaAlta = rs!FechaAlta
           Rs1!FechaBaja = rs!FechaBaja
           Rs1!Proveedor = nProv
           Rs1!UMCompra = nMedi
           Rs1!UMVenta = nMedi
           Rs1!Deposito = nDep
           Rs1!Origen = nOri
           Rs1!CantDespacho = rs!CantDespacho
           Rs1!CantVenta = rs!CantVenta
           Rs1!FechaCompra = rs!FechaCompra
           Rs1!Preciocompra = rs!Preciocompra
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
           Rs2!Impuesto = cRu.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & rs!Producto & "'")
           Rs2.Update

           Rs3.AddNew
           Rs3!Lista = nLista
           Rs3!Producto = rs!Producto
           Rs3!Precio = cRu.TraerValorDeUnCampo("Precios", "Precio", "Lista=" & nLista & " AND Producto='" & rs!Producto & "'")
           Rs3.Update
           
           Rs4.AddNew
           Rs4!Producto = rs!Producto
           Rs4!CantidadA = rs!CantDespacho
           Rs4!UnidadA = rs!UMCompra
           Rs4!CantidadB = rs!CantVenta
           Rs4!Unidadb = rs!UMVenta
           Rs4.Update
        Else
            nPrecio = cRu.TraerValorDeUnCampo("Precios", "Precio", "Lista=" & nLista & " AND Producto='" & rs!Producto & "'")
            db.Execute "UPDATE Precios SET Precio=" & nPrecio & " WHERE Producto='" & rs!Producto & "' AND Lista=" & nLista
        End If
        ProgressBar1.Value = X
        X = X + 1
        rs.MoveNext
     Loop
  End If
  
  MsgBox "Importación Realizada con Exito", vbCritical, "Atención"
  CmdProducto.Enabled = True

  Set cRsl = Nothing

End Sub

Private Sub Form_Load()
  Us = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Sucursal.mdb"
  DtpFecha.Value = Date
End Sub
