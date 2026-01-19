VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscarEx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkAvance 
      Caption         =   "B.Avanzada"
      Height          =   210
      Left            =   6900
      TabIndex        =   6
      Top             =   5415
      Width           =   1740
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Descripción"
      Height          =   225
      Index           =   2
      Left            =   3285
      TabIndex        =   5
      Top             =   5400
      Width           =   2370
   End
   Begin VB.TextBox TxtBuscar 
      Height          =   285
      Left            =   825
      TabIndex        =   2
      Text            =   "TxtBuscar"
      Top             =   75
      Width           =   10245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Código"
      Height          =   225
      Index           =   0
      Left            =   165
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Descripción"
      Height          =   225
      Index           =   1
      Left            =   1530
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   4740
      Left            =   165
      TabIndex        =   3
      Top             =   510
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   8361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Criterio :"
      Height          =   225
      Left            =   195
      TabIndex        =   4
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "FrmBuscarEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Rsd As Recordset

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
  
Private Const WM_NEXTDLGCTL As Long = &H28

Private mInEnter As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        nDat = 0
        Unload Me
        Exit Sub
    End If

    If KeyCode = vbKeyReturn Or KeyCode = 13 Then

        ' Si el foco está en TxtBuscar, no tabules (porque Enter ahí ejecuta búsqueda)
        If Not ActiveControl Is Nothing Then
            If ActiveControl.Name = "TxtBuscar" Then Exit Sub
        End If

        ' Si es un TextBox multiline, dejá Enter normal
        If TypeOf ActiveControl Is TextBox Then
            If ActiveControl.MultiLine Then Exit Sub
        End If

        KeyCode = 0
        SendMessage Me.hwnd, WM_NEXTDLGCTL, 0, 0
        Exit Sub
    End If

End Sub


Private Sub ChkAvance_Click()
  TxtBuscar.text = ""
  If ChkAvance.Value = 0 Then
     TxtBuscar.SetFocus
  End If
End Sub

Private Sub LlenarGrillaProducto(sBus As String, X As Byte)
  Dim Rx As Recordset, cRx As ClsProductoL, cRsl As ClsLectura
  
  Set cRx = New ClsProductoL
  Set cRsl = New ClsLectura
  
  Dim nAvan As Boolean
  
  nAvan = False
  If ChkAvance.Value = 1 Then nAvan = True
  
  Set Rx = cRx.TraerListaProduc(sBus, X, 0, nAvan, bBuscarActivo)
 ' CabGrid
   If Rsd.RecordCount <> 0 Then
      Rsd.MoveFirst
      Do While Not Rsd.EOF
         Rsd.Delete
         Rsd.MoveNext
      Loop
   End If
  If Rx.RecordCount <> 0 Then
     Do While Not Rx.EOF
        Rsd.AddNew
        Rsd!Producto = Rx!Producto
        Rsd!Descripcion = Rx!Descripcion
        Rsd!Adicional = Rx!Descripcion1
        Rsd!Cantidad = Rx!Cantidad
        If nBuscar = 4 Then
           Rsd!PrecioCompra = cRsl.TraerValorDeUnCampo("Productos", "PrecioCompra", "Producto='" & Rx!Producto & "'")
        End If
        If nLista = 0 Then nLista = 1
        If nLlama = 5 And bModoSuper = True Then
           Rsd!precio = cRsl.TraerValorDeUnCampo("PreciosCaja", "Precio", "Producto='" & Rx!Producto & "'")
        Else
           Rsd!precio = cRsl.TraerValorDeUnCampo("Precios", "Precio", "Producto='" & Rx!Producto & "' And Lista=" & nLista)
        End If
        Rsd.Update
        Rx.MoveNext
     Loop
     Rsd.MoveFirst
  End If
End Sub

Private Sub Form_Activate()
  TxtBuscar.SetFocus
  Option1(1).Value = True
End Sub

Private Sub Form_Load()
  TxtBuscar.text = ""
  
  Select Case nBuscar
         Case 1
           Option1(2).Visible = False
           CabGridCliente
         Case 2
           Option1(2).Visible = False
           CabGridProveedor
         Case 3
           Option1(2).Caption = "Descripción Adicional"
           CabGridProducto
           ChkAvance.Value = IIf(bPagare = False, 0, 1)
         Case 4
           Option1(2).Caption = "Descripción Adicional"
           CabGridProductoCompra
           ChkAvance.Value = IIf(bPagare = False, 0, 1)

  End Select
End Sub

Private Sub CabGridProducto()
  Set Rsd = New Recordset

  Rsd.Fields.Append "Producto", adVarChar, 25
  Rsd.Fields.Append "Descripcion", adVarChar, 50
  Rsd.Fields.Append "Adicional", adVarChar, 50, adFldIsNullable
  Rsd.Fields.Append "Cantidad", adSingle
  Rsd.Fields.Append "Precio", adCurrency
  Rsd.Open
  
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Produc."
      .Columns(0).Width = 1100
      .Columns(1).Caption = "Descripción"
      .Columns(1).Width = 5000
      .Columns(2).Caption = "Desc. Adicional"
      .Columns(2).Width = 2700
      .Columns(3).Caption = "Cant."
      .Columns(3).Width = 700
      .Columns(3).Alignment = dbgRight
      .Columns(3).NumberFormat = "#0.000"
      .Columns(4).Caption = "P.Uni."
      .Columns(4).Width = 1000
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(4).Alignment = dbgRight
 End With

End Sub

Private Sub CabGridProductoCompra()
  Set Rsd = New Recordset

  Rsd.Fields.Append "Producto", adVarChar, 25
  Rsd.Fields.Append "Descripcion", adVarChar, 50
  Rsd.Fields.Append "Adicional", adVarChar, 50, adFldIsNullable
  Rsd.Fields.Append "Cantidad", adSingle
  Rsd.Fields.Append "PrecioCompra", adCurrency
  Rsd.Fields.Append "Precio", adCurrency
  Rsd.Open
  
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Produc."
      .Columns(0).Width = 1100
      .Columns(1).Caption = "Descripción"
      .Columns(1).Width = 4000
      .Columns(2).Caption = "Desc. Adicional"
      .Columns(2).Width = 2700
      .Columns(3).Caption = "Cant."
      .Columns(3).Width = 700
      .Columns(3).Alignment = dbgRight
      .Columns(3).NumberFormat = "#0.000"
      .Columns(4).Caption = "P.Compr"
      .Columns(4).Width = 1000
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(4).Alignment = dbgRight
      .Columns(5).Caption = "P.Venta"
      .Columns(5).Width = 1000
      .Columns(5).NumberFormat = nCantDecimales
      .Columns(5).Alignment = dbgRight
      
 End With

End Sub

Private Sub Grid1_DblClick()
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura
  
  If Rsd.RecordCount <> 0 Then
     Select Case nBuscar
            Case 1
                nDat = cRsl.TraerValorDeUnCampo("Clientes", "Id", "Cliente='" & Rsd.Fields(0).Value & "'")
                Unload Me
            Case 2
                nDat = cRsl.TraerValorDeUnCampo("Proveedores", "Id", "Proveedor='" & Rsd.Fields(0).Value & "'")
                Unload Me
            Case 3, 4
                nDat = cRsl.TraerValorDeUnCampo("Productos", "Id", "Producto='" & Rsd.Fields(0).Value & "'")
                Unload Me
      End Select
  End If
  Set cRsl = Nothing
End Sub

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
  Rsd.Sort = Rsd.Fields(ColIndex).Name & " Asc"
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
 
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura
  If KeyCode = 13 Or KeyCode = 144 Then
  If Rsd.RecordCount <> 0 Then
      Select Case nBuscar
             Case 1
                nDat = cRsl.TraerValorDeUnCampo("Clientes", "Id", "Cliente='" & Rsd.Fields(0).Value & "'")
                Unload Me
             Case 2
                nDat = cRsl.TraerValorDeUnCampo("Proveedores", "Id", "Proveedor='" & Rsd.Fields(0).Value & "'")
                Unload Me
             Case 3, 4
                nDat = cRsl.TraerValorDeUnCampo("Productos", "Id", "Producto='" & Rsd.Fields(0).Value & "'")
                Unload Me
       End Select
     End If
  End If
  Set cRsl = Nothing

End Sub

Private Sub Option1_Click(Index As Integer)
  TxtBuscar.SetFocus
  TxtBuscar.text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        TxtBuscar_LostFocus
        Grid1.SetFocus
        Exit Sub
    End If
End Sub


Private Sub TxtBuscar_LostFocus()
  Dim X As Byte
  
  If TxtBuscar.text <> "" Then
     Select Case nBuscar
            Case 1
               If Option1(0).Value = True Then X = 1
               If Option1(1).Value = True Then X = 2
               LlenarGrillaCliente TxtBuscar.text, X
            Case 2
               If Option1(0).Value = True Then X = 1
               If Option1(1).Value = True Then X = 2
               LlenarGrillaProveedor TxtBuscar.text, X
            Case 3, 4
               If Option1(0).Value = True Then X = 1
               If Option1(1).Value = True Then X = 2
               If Option1(2).Value = True Then X = 3
               LlenarGrillaProducto TxtBuscar.text, X
     End Select
  End If
End Sub

Private Sub CabGridCliente()
  Set Rsd = New Recordset

  Rsd.Fields.Append "Cliente", adVarChar, 15
  Rsd.Fields.Append "Razon", adVarChar, 50
  Rsd.Fields.Append "Domicilio", adVarChar, 50
  Rsd.Fields.Append "Observacion", adVarChar, 50
  Rsd.Fields.Append "Saldo", adCurrency
  Rsd.Open
  
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Cliente"
      .Columns(0).Width = 800
      .Columns(1).Caption = "R.Social"
      .Columns(1).Width = 3500
      .Columns(2).Caption = "Domicilio"
      .Columns(2).Width = 2300
      .Columns(3).Caption = "Observación"
      .Columns(3).Width = 2500
      .Columns(4).Caption = "Saldo"
      .Columns(4).Width = 950
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(4).Alignment = dbgRight
 End With

End Sub

Private Sub CabGridProveedor()
  Set Rsd = New Recordset

  Rsd.Fields.Append "Cliente", adVarChar, 15
  Rsd.Fields.Append "Razon", adVarChar, 50
  Rsd.Fields.Append "Domicilio", adVarChar, 50
  Rsd.Fields.Append "Observacion", adVarChar, 50
  Rsd.Fields.Append "Saldo", adCurrency
  Rsd.Open
  
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Proveedor"
      .Columns(0).Width = 800
      .Columns(1).Caption = "R.Social"
      .Columns(1).Width = 3500
      .Columns(2).Caption = "Domicilio"
      .Columns(2).Width = 2300
      .Columns(3).Caption = "Observación"
      .Columns(3).Width = 2500
      .Columns(4).Caption = "Saldo"
      .Columns(4).Width = 950
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(4).Alignment = dbgRight
 End With

End Sub
Private Sub LlenarGrillaCliente(sBus As String, X As Byte)
  Dim Rx As Recordset, cRx As ClsProductoL, cRe As ClsClienteL
  
  
  Dim nAvan As Boolean
  
  nAvan = False
  If ChkAvance.Value = 1 Then nAvan = True
  
  
  Set cRx = New ClsProductoL
  Set Rx = cRx.TraerListaCliente(sBus, X, nAvan)
  Set cRe = New ClsClienteL
  
 ' CabGrid
   If Rsd.RecordCount <> 0 Then
      Rsd.MoveFirst
      Do While Not Rsd.EOF
         Rsd.Delete
         Rsd.MoveNext
      Loop
   End If
  If Rx.RecordCount <> 0 Then
     Do While Not Rx.EOF
        Rsd.AddNew
        Rsd!Cliente = Rx!Cliente
        Rsd!Razon = "" & Rx!RazonSocial
        Rsd!Domicilio = "" & Rx!Domicilio
        Rsd!Observacion = "" & Left(Rx!Observaciones, 50)
        Rsd!saldo = cRe.ResumenSaldoAnteriorBus(Rx!Cliente) ' IIf(IsNull(Rx!Saldo), 0, Rx!Saldo)
        Rsd.Update
        Rx.MoveNext
     Loop
     Rsd.MoveFirst
  End If
End Sub

Private Sub LlenarGrillaProveedor(sBus As String, X As Byte)
  Dim Rx As Recordset, cRx As ClsProductoL, cRe As ClsClienteL
  
  
  Dim nAvan As Boolean
  
  nAvan = False
  If ChkAvance.Value = 1 Then nAvan = True
  
  Set cRx = New ClsProductoL
  Set Rx = cRx.TraerListaProveedor(sBus, X, nAvan)
  Set cRe = New ClsClienteL
  
 ' CabGrid
   If Rsd.RecordCount <> 0 Then
      Rsd.MoveFirst
      Do While Not Rsd.EOF
         Rsd.Delete
         Rsd.MoveNext
      Loop
   End If
  If Rx.RecordCount <> 0 Then
     Do While Not Rx.EOF
        Rsd.AddNew
        Rsd!Cliente = Rx!proveedor
        Rsd!Razon = "" & Rx!RazonSocial
        Rsd!Domicilio = "" & Rx!Domicilio
        Rsd!Observacion = "" & Left(Rx!Observaciones, 50)
        Rsd!saldo = cRe.ResumenSaldoAnteriorBusP(Rx!proveedor)
        Rsd.Update
        Rx.MoveNext
     Loop
     Rsd.MoveFirst
  End If
End Sub

