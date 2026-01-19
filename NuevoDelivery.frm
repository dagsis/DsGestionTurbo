VERSION 5.00
Begin VB.Form NuevoDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Pedido"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtObservacion 
      Height          =   810
      Left            =   3630
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "NuevoDelivery.frx":0000
      Top             =   2025
      Width           =   3960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   3585
      TabIndex        =   13
      Top             =   90
      Width           =   4005
      Begin VB.TextBox txtNTelefono 
         Height          =   285
         Left            =   930
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "txtNTelefono"
         Top             =   1290
         Width           =   2970
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   195
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "txtDireccion"
         Top             =   855
         Width           =   3705
      End
      Begin VB.TextBox txtNombre 
         Height          =   300
         Left            =   870
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "txtNombre"
         Top             =   300
         Width           =   3060
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono :"
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Nombre :"
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdPlatos 
      Caption         =   "&Platos"
      Height          =   450
      Left            =   1605
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2985
      Width           =   4695
   End
   Begin VB.TextBox txtTelefono 
      Height          =   285
      Left            =   210
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "txtTelefono"
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label lblCliente 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2310
      TabIndex        =   19
      Top             =   615
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "Nº de Cliente :"
      Height          =   225
      Left            =   1035
      TabIndex        =   18
      Top             =   585
      Width           =   1110
   End
   Begin VB.Label Label11 
      Caption         =   "Observación :"
      Height          =   225
      Left            =   3630
      TabIndex        =   17
      Top             =   1770
      Width           =   1215
   End
   Begin VB.Label lblPedido 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblPedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2310
      TabIndex        =   12
      Top             =   195
      Width           =   1140
   End
   Begin VB.Label Label6 
      Caption         =   "Nº de Pedido :"
      Height          =   225
      Left            =   1080
      TabIndex        =   11
      Top             =   195
      Width           =   1110
   End
   Begin VB.Line Line1 
      X1              =   225
      X2              =   3435
      Y1              =   2805
      Y2              =   2805
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   915
      Width           =   750
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblNombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   9
      Top             =   1725
      Width           =   3225
   End
   Begin VB.Label lblDireccion 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDireccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   8
      Top             =   2445
      Width           =   3225
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre :"
      Height          =   225
      Left            =   225
      TabIndex        =   7
      Top             =   1485
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Dirección :"
      Height          =   225
      Left            =   225
      TabIndex        =   6
      Top             =   2205
      Width           =   1215
   End
End
Attribute VB_Name = "NuevoDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bBusDel As Boolean

Private Sub CmdPlatos_Click()
  Dim sSql As String, cRsl As ClsLectura, nCli As Long, nZona As Integer
  Dim rAux As Recordset, cRle As ClsEscritura
    
  Set rAux = New Recordset
  Set cRle = New ClsEscritura

  Set cRsl = New ClsLectura
  
  On Error GoTo errHandler
  
  nCli = lblCliente.Caption
  nZona = cRsl.RegPorDefecto("Zona")
  nPedido = cRsl.ActualizarDeliveryTraer
    
  lblPedido.Caption = Format(nPedido, "00000000")
    
     
  If Frame1.Enabled = True Then
     db.Open sDb
     db.BeginTrans

     sSql = "INSERT INTO Clientes ([Tipo],[Cliente],[RazonSocial],[Contacto],[Observaciones],[FechaAlta],[FechaBaja]"
     sSql = sSql & ",[Domicilio],[CodigoPostal],[Localidad],[Ciudad],[Provincia],[Pais],[Telefono],[Fax],[Transporte],[TipoIva],[Cuit],[IngBruto]"
     sSql = sSql & ",[Zona],[Precio],[Vendedor],[Email],[Descuento],[Limite],[Recargo],[Retencion],[Percepcion]) "
     sSql = sSql & "Values "
     sSql = sSql & "(1,'" & nCli & "','" & lblNombre.Caption & "','','','01/01/2017',NULL,'" & lblDireccion.Caption & "',0,'','',"
     sSql = sSql & "1,1,'" & txtTelefono.text & "','',1,1,'',''," & nZona & ",1,1,'',1,0,1,0,0)"

     db.Execute sSql

     db.CommitTrans
     db.Close
 End If
    bDelivery = True
    Set rAux = cRsl.RsVacio("Auxiliar", "Id", "N")
    
    rAux.AddNew
    rAux![a_codart] = "02-0099"
    rAux![A_DESCRI] = "SERVICIO DE DELIVERY"
    rAux![a_cant] = 1
    rAux![a_Mesa] = nPedido
    rAux![a_mozo] = cRsl.RegPorDefecto("Vendedor")
    rAux![a_impcom] = "N"
    rAux![a_hora] = Time
    rAux![a_fecha] = Date
    rAux![A_MENSAJE] = "DELIVERY"
    rAux![a_lista] = cRsl.RegPorDefecto("ListaDel")
    rAux![a_precio] = cRsl.TraerValorDeUnCampo("Precios", "Precio", "Producto='" & rAux![a_codart] & "' AND Lista=" & rAux![a_lista])
    rAux![a_puerto] = "D"
    rAux![a_puerto2] = cRsl.TraerValorDeUnCampo("Clientes", "Cliente", "Cliente='" & lblCliente.Caption & "'")
    rAux![a_puerto3] = "00:00:00"
    rAux![a_total] = 0
    rAux![a_Observ] = txtObservacion.text
    rAux.Update
  
    cRle.Actualizar "Auxiliar_A", rAux
  
    nLista = cRsl.RegPorDefecto("Precio")
  ' nMozo = cRsl.RegPorDefecto("Vendedor")
    nMesa = nPedido
  
    Unload Me
    
    Delivery.LlenarGrilla
    
    frmMesa.Show 1
  
Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    Err.Raise Err.Number
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim cRsl As ClsLectura, Rx As Recordset
 
 Set cRsl = New ClsLectura
 
 If KeyCode = vbKeyF1 Then
    bBusDel = True
    txtTelefono.text = ""
    nDat = 0
    nBuscar = 1
    FrmBuscarEx.Show 1
    If nDat <> 0 Then
       Set Rx = cRsl.TraerRsCondi("Clientes", "Id", "Id=" & nDat)
       bBusDel = False
       Frame1.Enabled = False
       txtTelefono.text = Rx!Cliente
       lblNombre.Caption = Rx!RazonSocial
       lblDireccion.Caption = IIf(Rx!Domicilio = "", "S/E", Rx!Domicilio)
       CmdPlatos.SetFocus
    End If
 End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
     End If
'     KeyAscii = Upper(KeyAscii)
End Sub

Private Sub Form_Load()
   Limpiar
   Habilitar False, False, False
   bBusDel = False
End Sub

Private Sub Limpiar()
  lblPedido.Caption = "00000000"
  lblCliente.Caption = ""
  txtTelefono.text = ""
  lblNombre.Caption = ""
  lblDireccion.Caption = ""
  txtNombre.text = ""
  txtDireccion.text = ""
  txtNTelefono.text = ""
  txtObservacion.text = ""
End Sub

Private Sub Habilitar(pHabi As Boolean, pNuevo As Boolean, pPlato As Boolean)
  Frame1.Enabled = pHabi
  CmdPlatos.Enabled = pPlato
End Sub

Private Sub lblDireccion_Change()
  TodoBien
End Sub

Private Sub lblNombre_Change()
  TodoBien
End Sub

Private Sub txtDireccion_LostFocus()
  lblDireccion.Caption = txtDireccion.text
End Sub

Private Sub txtNombre_LostFocus()
  lblNombre.Caption = txtNombre.text
End Sub

Private Sub txtNTelefono_Change()
  TodoBien
End Sub

Private Sub txtNTelefono_LostFocus()
  txtTelefono.text = txtNTelefono.text
End Sub

Private Sub txtTelefono_LostFocus()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  lblNombre.Caption = ""
  lblCliente.Caption = ""
  lblDireccion.Caption = ""
  
  If bBusDel = False Then
    If txtTelefono.text <> "" Then
       If cRsl.TraerCantidad("Clientes", "Cliente", "Cliente='" & txtTelefono.text & "'") <> 0 Then
          lblCliente = txtTelefono.text
          lblNombre.Caption = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & txtTelefono.text & "'")
          lblDireccion.Caption = cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & txtTelefono.text & "'")
          If lblDireccion.Caption = "" Then lblDireccion.Caption = "S/E"
          Frame1.Enabled = False
          txtObservacion.SetFocus
       Else
           lblCliente.Caption = cRsl.ActualizarCampoTraer("Cliente")
           txtTelefono.text = lblCliente.Caption
           Frame1.Enabled = True
           txtNombre.SetFocus
       End If
    Else
       lblCliente.Caption = cRsl.ActualizarCampoTraer("Cliente")
       txtTelefono.text = lblCliente.Caption
       Frame1.Enabled = True
       txtNombre.SetFocus
    End If
  End If
End Sub

Private Sub TodoBien()
  CmdPlatos.Enabled = False
  If Len(txtTelefono.text) <> 0 And Len(lblNombre.Caption) <> 0 And Len(lblDireccion.Caption) <> 0 Then
     CmdPlatos.Enabled = True
  End If
End Sub
