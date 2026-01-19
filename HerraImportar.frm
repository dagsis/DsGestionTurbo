VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form HerraImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Productos"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   5385
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   4710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   465
      Left            =   4245
      TabIndex        =   9
      Top             =   4560
      Width           =   1005
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
      Height          =   4455
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   5175
      Begin VB.CheckBox ChkBarra 
         Caption         =   "B.Cdigo Barra"
         Height          =   270
         Left            =   2385
         TabIndex        =   24
         Top             =   3525
         Width           =   1350
      End
      Begin VB.CommandButton CmdProveedor 
         Caption         =   "Proveedores"
         Height          =   390
         Left            =   3900
         TabIndex        =   23
         Top             =   3735
         Width           =   1125
      End
      Begin VB.CommandButton CmdCliente 
         Caption         =   "Clientes"
         Height          =   390
         Left            =   3900
         TabIndex        =   22
         Top             =   3240
         Width           =   1125
      End
      Begin VB.CheckBox ChkExiste 
         Caption         =   "No Agregar si no Existe"
         Height          =   270
         Left            =   165
         TabIndex        =   21
         Top             =   3510
         Width           =   1995
      End
      Begin VB.Frame Frame2 
         Caption         =   "Origen"
         Height          =   570
         Left            =   120
         TabIndex        =   18
         Top             =   3810
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "Excel"
            Height          =   225
            Left            =   270
            TabIndex        =   20
            Top             =   255
            Width           =   945
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Access"
            Height          =   240
            Left            =   1590
            TabIndex        =   19
            Top             =   255
            Width           =   915
         End
      End
      Begin VB.CheckBox ChkActualiza 
         Caption         =   "Actualiza Precios de Ventas por Compras"
         Height          =   225
         Left            =   165
         TabIndex        =   17
         Top             =   3240
         Width           =   4020
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   345
         Left            =   3270
         TabIndex        =   15
         Top             =   390
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         _Version        =   393216
         Format          =   154206209
         CurrentDate     =   42323
      End
      Begin VB.ComboBox CmbUnidad 
         Height          =   315
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2415
         Width           =   2970
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2790
         Width           =   2970
      End
      Begin VB.ComboBox CmbMarca 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2025
         Width           =   2970
      End
      Begin VB.ComboBox CmbFamilia 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1650
         Width           =   2970
      End
      Begin VB.ComboBox CmbProveedor 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1275
         Width           =   2970
      End
      Begin VB.ComboBox CmbLista 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   900
         Width           =   2970
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   390
         TabIndex        =   14
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "U. Medida :"
         Height          =   210
         Left            =   390
         TabIndex        =   13
         Top             =   2415
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Moneda :"
         Height          =   240
         Left            =   390
         TabIndex        =   12
         Top             =   2790
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Marca :"
         Height          =   240
         Left            =   390
         TabIndex        =   4
         Top             =   2025
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Familia :"
         Height          =   210
         Left            =   390
         TabIndex        =   3
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor :"
         Height          =   285
         Left            =   390
         TabIndex        =   2
         Top             =   1275
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Lista de Precio :"
         Height          =   225
         Left            =   390
         TabIndex        =   1
         Top             =   900
         Width           =   1290
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   135
      TabIndex        =   16
      Top             =   4575
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "HerraImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbLista, "ListaDePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbFamilia, "Familias", "Familia", "Descripcion", ""
  cRsl.CargaCombo CmbMarca, "Marcas", "Marca", "Descripcion", ""
  cRsl.CargaCombo CmbProveedor, "Proveedores", "Id", "RazonSocial", ""
  cRsl.CargaCombo CmbUnidad, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbMoneda, "Monedas", "Id", "Descripcion", ""

  CmbMoneda.text = cRsl.TraerValorDeUnCampo("Monedas", "Descripcion", "Id=" & cRsl.TraerValorDeUnCampo("Registros", "Moneda", "Usuario=" & nUsuario))
  Set cRsl = Nothing
End Sub

Private Sub CmdAplicar_Click()
 On Error GoTo Errores
 Dim cRsl As ClsLectura, Rs As Recordset, X As Long, cRle As ClsProductoE
  Dim Rs1 As Recordset, Rs2 As Recordset, Rs3 As Recordset
  Dim cRu As ClsHerramientas, nLista As Long, nPrecio As Double, nPorce As Double, nRow As Long, nMone As Byte
   
  Dim nTotal As Long, nNuevo As Long, nModi As Long
   
  Set cRu = New ClsHerramientas
  Set cRsl = New ClsLectura
  Set cRle = New ClsProductoE
  
  Set Rs1 = cRsl.RsVacio("Productos", "Id", "N")
  Set Rs2 = cRsl.RsVacio("ProductoImpuesto", "Id", "N")
  Set Rs3 = cRsl.RsVacio("Precios", "Id", "N")
  
  CmdAplicar.Enabled = False
  If Option2.Value = True Then
      Set Rs = cRu.TraerTodos("Productos", "*")
  Else
      Dim Libro As String
      Dim Hoja As String
      Dim rango As String
'
      CommonDialog1.InitDir = App.Path
      CommonDialog1.DialogTitle = "Seleccione un archivo a Importar"
      CommonDialog1.Filter = "Archivos de Excel|*.xls"
      CommonDialog1.ShowOpen

      If CommonDialog1.FileName <> "" Then

         Libro = CommonDialog1.FileName
               
         Hoja = "Hoja1$"
         Dim Conexion As ADODB.Connection
         Set Conexion = New Connection
      
         Conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Libro & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
         Set Rs = New ADODB.Recordset
         With Rs
             .CursorLocation = adUseClient
             .CursorType = adOpenStatic
             .LockType = adLockOptimistic
        End With

        Rs.Open "SELECT * FROM [" & Hoja & "]", Conexion, , , adCmdText
        ' Rutina para importar planillas varias
'        If Rs.RecordCount <> 0 Then
'           db.Open sDb
'           Do While Not Rs.EOF
'           db.Execute "UPDATE Productos SET Origen=2 WHERE Producto='" & Rs!Producto & "'"
'           Rs.MoveNext
'           Loop
'           db.Close
'           End
'        End If
      Else
         Exit Sub
      End If
  
  End If
  
  
  
  X = 1
  Dim nProv As Long, nMedi As Long, nDep As Long, nOri As Long, nFam As Long, nMar As Long
  Dim sCodAux As String, sCodAuxTbl As String
  
  nLista = CmbLista.ItemData(CmbLista.ListIndex)
  nProv = CmbProveedor.ItemData(CmbProveedor.ListIndex)
  nMedi = CmbUnidad.ItemData(CmbUnidad.ListIndex)
  nDep = cRsl.TraerValorDeUnCampo("Registros", "Deposito", "")
  nOri = 1
  nFam = CmbFamilia.ItemData(CmbFamilia.ListIndex)
  nMar = CmbMarca.ItemData(CmbMarca.ListIndex)
  nPorce = cRsl.TraerValorDeUnCampo("ListaDePrecio", "Porcentaje", "ListaPrecio=" & CmbLista.ItemData(CmbLista.ListIndex))
  nMone = CmbMoneda.ItemData(CmbMoneda.ListIndex)
  
  If Rs.RecordCount <> 0 Then
     nRow = 1
     ProgressBar1.Max = Rs.RecordCount
     nTotal = Rs.RecordCount
     
     Do While Not Rs.EOF
         If cRsl.TraerCantidad("Productos", "Producto", "Cbarra='" & Rs!Producto & "'") <> 0 Then
            sCodAux = cRsl.TraerValorDeUnCampo("Productos", "Producto", "Cbarra='" & Rs!Producto & "'")
         Else
            sCodAux = Rs!Producto
         End If
        If ChkExiste.Value = 0 Then
           If cRsl.TraerCantidad("Productos", "Id", "Producto='" & Trim(sCodAux & "'")) = 0 Then
              If Not IsNull(Trim(Rs!Producto)) Then
                 If Rs!Proveedor <> 0 Then
                    nProv = cRsl.TraerValorDeUnCampo("Proveedores", "Id", "Proveedor='" & Rs!Proveedor & "'")
                 End If
                 Rs1.AddNew
                 Rs1!Activo = True
                 Rs1!Producto = Trim(Rs!Producto)
                 Rs1!Moneda = IIf(Rs!Moneda = 0, CmbMoneda.ItemData(CmbMoneda.ListIndex), Rs!Moneda)
                 If Not IsNull(Rs!Descripcion) Then
                 Rs1!Descripcion = Left(Replace(Rs!Descripcion, "'", " "), 50)
                 Else
                   Rs1!Descripcion = "REVISAR"
                 End If
                 Rs1!Descripcion1 = Replace(IIf(IsNull(Rs!Adicional), "", Left(Rs!Adicional, 50)), "'", " ")
                 Rs1!tipo = 3
                 Rs1!Familia = IIf(Rs!Familia = 0, CmbFamilia.ItemData(CmbFamilia.ListIndex), Rs!Familia)
                 Rs1!FamiliaSub = Null
                 Rs1!Representante = 0
                 Rs1!Peso = 0
                 Rs1!Marca = IIf(Rs!Marca = 0, CmbMarca.ItemData(CmbMarca.ListIndex), Rs!Marca)
                 Rs1!FechaAlta = DtpFecha.Value
                 Rs1!FechaBaja = DtpFecha.Value
                 Rs1!Proveedor = IIf(Rs!Proveedor = 0, CmbProveedor.ItemData(CmbProveedor.ListIndex), Rs!Proveedor)
                 Rs1!UMCompra = nMedi
                 Rs1!UMVenta = nMedi
                 Rs1!Deposito = nDep
                 Rs1!Origen = nOri
                 Rs1!CantDespacho = 1
                 Rs1!CantVenta = 1
                 Rs1!FechaCompra = DtpFecha.Value
                 Rs1!PrecioCompra = IIf(IsNull(Rs!Compra), 0, Rs!Compra)
                 Rs1!Actualiza = False
                 Rs1!CantMinima = 0
                 Rs1!CantMaxima = Rs!Cantidad
                 Rs1!ActuaStock = "Si"
                 Rs1!CodCompu = "No"
                 Rs1!Fijo = "Si"
                 Rs1!TipoInterno = "No"
                 Rs1!ValorInterno = 0
                 Rs1!Calcular = "No"
                 Rs1!CBarra = "" & Rs!CodBarra
                 Rs1!ValorEnvase = 0
                 Rs1!Balanza = 0
                 Rs1!CDepartamento = 1
                 Rs1!CFamilia = 1
                 Rs1.Update
                
                 Rs2.AddNew
                 Rs2!Producto = Trim(Rs!Producto)
                 If Rs!Iva <> 0 Then
                    Rs2!Impuesto = IIf(Rs!Iva = 21, 15, 9)
                 Else
                   Rs2!Impuesto = 10
                 End If
                 Rs2.Update
      
                 Rs3.AddNew
                 Rs3!Lista = nLista
                 Rs3!Producto = Trim(Rs!Producto)
                 Rs3!Porcentaje = 0
                 Rs3!precio = IIf(IsNull(Rs!Venta), 0, RedondearPrecios(IIf(IsNull(Rs!Venta), 0, Rs!Venta)))
                 If ChkActualiza.Value = 1 Then
                     Dim sImp As Double
                     sImp = Rs!Compra * (((Rs!Iva) / 100) + 1)
                     sImp = Round(sImp, 2)
                     Rs3!precio = RedondearPrecios(sImp * (((nPorce) / 100) + 1))
                 End If
                 Rs3.Update
                 
                 nNuevo = nNuevo + 1
              Else
                 GoTo seguir
              End If
           Else
              db.Open sDb
              If ChkActualiza.Value = 0 Then
                 db.Execute "UPDATE Productos SET FechaAlta='" & DtpFecha.Value & "', FechaBaja='" & DtpFecha.Value & "',FechaCompra='" & DtpFecha.Value & "', Familia=" & IIf(Rs!Familia = 0, CmbFamilia.ItemData(CmbFamilia.ListIndex), Rs!Familia) & ",Marca=" & IIf(Rs!Marca = 0, CmbMarca.ItemData(CmbMarca.ListIndex), Rs!Marca) & ",Proveedor=" & nProv & ",UmCompra=" & nMedi & ",UMVenta=" & nMedi & ",Actualiza=0,Moneda=" & nMone & ",PrecioCompra=" & IIf(IsNull(Rs!Compra), 0, Rs!Compra) & " WHERE Producto='" & Trim(sCodAux) & "'"
                 db.Execute "UPDATE Precios SET Precio=" & RedondearPrecios(IIf(IsNull(Rs!Venta), 0, Rs!Venta)) & " WHERE Producto='" & Trim(sCodAux) & "' AND Lista=" & nLista
              Else
                 db.Execute "UPDATE Productos SET FechaCompra='" & DtpFecha.Value & "', Familia=" & nFam & ",Marca=" & nMar & ",Proveedor=" & nProv & ",UmCompra=" & nMedi & ",UMVenta=" & nMedi & ",Actualiza=1,PrecioCompra=" & Rs!Compra & " WHERE Producto='" & Trim(sCodAux) & "'"
              End If
              db.Close
              nModi = nModi + 1
           End If
        Else
           If Not IsNull(Trim(Rs!Producto)) Then
            db.Open sDb
            Dim sSql As String
            sSql = "UPDATE Productos SET Descripcion='" & Left(Replace(Rs!Descripcion, "'", " "), 50) & "',FechaAlta='" & DtpFecha.Value & "',Familia=" & Rs!Familia & ",Marca=" & Rs!Marca & ",FechaBaja='" & DtpFecha.Value & "', FechaCompra='" & DtpFecha.Value & "',Actualiza=0,Moneda=" & Rs!Moneda & ",PrecioCompra=" & IIf(IsNull(Rs!Compra), 0, Rs!Compra) & " WHERE Producto='" & Trim(sCodAux) & "'"
            db.Execute sSql
            db.Execute "UPDATE Precios SET Precio=" & IIf(IsNull(Rs!Venta), 0, Rs!Venta) & " WHERE Producto='" & Trim(sCodAux) & "' AND Lista=" & nLista
            db.Close
            nModi = nModi + 1
           End If
        End If
seguir:
        ProgressBar1.Value = X
        X = X + 1
      
        Rs.MoveNext
        nRow = nRow + 1
     Loop
     If ChkExiste.Value = 0 Then
        If cRle.AgregarImportacion(Rs1, Rs2, Rs3) = True Then
                MsgBox "Importación Realizada con Exito" & Chr(10) & Chr(13) & "Total Importado:" & nTotal _
                & Chr(10) & Chr(13) & "Total Agregados:" & nNuevo _
                & Chr(10) & Chr(13) & "Total Modificados:" & nModi, vbCritical, "Atención"
        Else
           MsgBox "No se Proceso ningun Registro....", vbCritical, "Atención"
        End If
     Else
         MsgBox "Importación Realizada con Exito" & Chr(10) & Chr(13) & "Total Importado:" & nTotal _
         & Chr(10) & Chr(13) & "Total Agregados:" & nNuevo _
         & Chr(10) & Chr(13) & "Total Modificados:" & nModi, vbCritical, "Atención"
     End If
  End If

  CmdAplicar.Enabled = True

  Set cRsl = Nothing
  Exit Sub
Errores:
   MsgBox err.Description, vbCritical, nModi
 '  If nModi <> 0 Then
     GoTo seguir
'   Else
'      MsgBox "Problemas en el Archivo de Importación", vbCritical, "Atención"
'   End If
End Sub

Private Sub CmdCliente_Click()
 On Error GoTo Errores
 Dim cRsl As ClsLectura, Rs As Recordset, X As Long, cRse As ClsClienteE
  Dim Rs1 As Recordset
  Dim cRu As ClsHerramientas, nLista As Long, nPrecio As Double, nPorce As Double
   
  Dim nTotal As Long, nNuevo As Long, nModi As Long
   
  Set cRu = New ClsHerramientas
  Set cRsl = New ClsLectura
  Set cRse = New ClsClienteE
  
  Set Rs1 = cRsl.RsVacio("Clientes", "Id", "N")
  
  CmdCliente.Enabled = False

  Dim Libro As String
  Dim Hoja As String
  Dim rango As String
'
  CommonDialog1.InitDir = App.Path
  CommonDialog1.DialogTitle = "Seleccione un archivo a Importar"
  CommonDialog1.Filter = "Archivos de Excel|*.xls"
  CommonDialog1.ShowOpen

  If CommonDialog1.FileName <> "" Then

     Libro = CommonDialog1.FileName
           
     Hoja = "Hoja1$"
     Dim Conexion As ADODB.Connection
     Set Conexion = New Connection
  
     Conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & Libro & _
              ";Extended Properties=""Excel 8.0;HDR=Yes;"""
     Set Rs = New ADODB.Recordset
     With Rs
         .CursorLocation = adUseClient
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
    End With
    
    Rs.Open "SELECT * FROM [" & Hoja & "]", Conexion, , , adCmdText

    End If
  
    X = 1
 
    If Rs.RecordCount <> 0 Then
     ProgressBar1.Max = Rs.RecordCount
     nTotal = Rs.RecordCount
     
     Do While Not Rs.EOF
           If cRsl.TraerCantidad("Clientes", "Id", "Cliente='" & Trim(Rs!Cliente & "'")) = 0 Then
              If Not IsNull(Trim(Rs!Cliente)) Then
                 Rs1.AddNew
                 Rs1!tipo = 1
                 Rs1!Cliente = Trim(Rs!Cliente)
                 Rs1!RazonSocial = Trim(Left(Rs!RazonSocial, 45))
                 Rs1!Contacto = "" & Trim(Left(Rs!Contacto, 45))
                 Rs1!Observaciones = "" & Trim(Rs!Observaciones)
                 Rs1!FechaAlta = Date
                 Rs1!FechaBaja = Null
                 Rs1!Domicilio = "" & Trim(Left(Rs!Domicilio, 45))
                 Rs1!CodigoPostal = "" & Trim(Rs!CodigoPostal)
                 Rs1!Localidad = 1
                 Rs1!Ciudad = "" & Trim(Left(Rs!Ciudad, 30))
                 Rs1!Provincia = Rs!Provincia
                 Rs1!Pais = 1
                 Rs1!Telefono = "" & Trim(Rs!Telefono)
                 Rs1!Fax = ""
                 Rs1!Transporte = 1
                 Rs1!TipoIva = Rs!TipoIva
                 Rs1!Cuit = "" & Trim(Rs!Cuit)
                 Rs1!IngBruto = ""
                 Rs1!Zona = 1
                 Rs1!precio = 1
                 Rs1!Vendedor = 1
                 Rs1!email = "" & Trim(Rs!email)
                 Rs1!Descuento = 1
                 Rs1!Limite = 0
                 Rs1!Recargo = 1
                 Rs1!Retencion = 0
                 Rs1!Percepcion = 0
                 Rs1.Update
                                 
                 nNuevo = nNuevo + 1
              End If
           End If
    
seguir:
        ProgressBar1.Value = X
        X = X + 1
      
        Rs.MoveNext
     Loop
     If ChkExiste.Value = 0 Then
        If Rs1.RecordCount <> 0 Then
                Rs1.MoveFirst
                Do While Not Rs1.EOF
                   cRse.Agregar Rs1, 0
                   Rs1.MoveNext
                Loop
                MsgBox "Importación Realizada con Exito" & Chr(10) & Chr(13) & "Total Importado:" & nTotal _
                & Chr(10) & Chr(13) & "Total Agregados:" & nNuevo _
                & Chr(10) & Chr(13) & "Total Modificados:" & nModi, vbCritical, "Atención"
        Else
           MsgBox "No se Proceso ningun Registro....", vbCritical, "Atención"
        End If
     Else
         MsgBox "Importación Realizada con Exito" & Chr(10) & Chr(13) & "Total Importado:" & nTotal _
         & Chr(10) & Chr(13) & "Total Agregados:" & nNuevo _
         & Chr(10) & Chr(13) & "Total Modificados:" & nModi, vbCritical, "Atención"
     End If
  End If

  CmdCliente.Enabled = True

  Set cRsl = Nothing
  Exit Sub
Errores:
   MsgBox err.Description, vbCritical, nModi
   If nModi <> 0 Then
     GoTo seguir
   Else
      MsgBox "Problemas en el Archivo de Importación", vbCritical, "Atención"
   End If
End Sub

Private Sub Form_Load()
  CargarCombos
  LinkearTextos
  
  Us = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Importador.mdb"
  DtpFecha.Value = Date
  Option1.Value = True

End Sub

Private Sub LinkearTextos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  CmbLista.text = cRsl.DatoCombo("ListaDePrecio", "ListaPrecio", "Descripcion", cRsl.RegPorDefecto("ListaPrecio"))
  CmbFamilia.text = cRsl.DatoCombo("Familias", "Familia", "Descripcion", cRsl.RegPorDefecto("Familia"))
  CmbMarca.text = cRsl.DatoCombo("Marcas", "Marca", "Descripcion", cRsl.RegPorDefecto("Marca"))
  CmbProveedor.text = cRsl.DatoCombo("Proveedores", "Id", "RazonSocial", cRsl.RegPorDefecto("Proveedor"))
  CmbUnidad.text = cRsl.DatoCombo("UnidadMedida", "Unidad", "Descripcion", cRsl.RegPorDefecto("Medida"))
  CmbMoneda.text = cRsl.DatoCombo("Monedas", "Id", "Descripcion", cRsl.RegPorDefecto("Moneda"))

  Set cRsl = Nothing
End Sub

