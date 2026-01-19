VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form VentaRapido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas "
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   8250
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   6945
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   8055
      Begin VB.TextBox TxtCPrecio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5055
         MaxLength       =   10
         TabIndex        =   45
         Text            =   "TxtCa"
         Top             =   5775
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Height          =   4980
         Left            =   390
         TabIndex        =   20
         Top             =   300
         Width           =   7155
         Begin VB.Frame Frame3 
            Caption         =   "Su pago"
            Height          =   4230
            Left            =   195
            TabIndex        =   22
            Top             =   255
            Width           =   6690
            Begin VB.CommandButton Command1 
               Caption         =   "F"
               Height          =   360
               Left            =   6120
               TabIndex        =   51
               Top             =   2550
               Width           =   375
            End
            Begin VB.TextBox TxtCliente 
               Height          =   300
               Left            =   1380
               MaxLength       =   15
               TabIndex        =   47
               Text            =   "TxtCliente"
               Top             =   270
               Width           =   795
            End
            Begin VB.CommandButton BtnBuscar 
               Caption         =   "..."
               Height          =   240
               Left            =   6210
               TabIndex        =   21
               Top             =   285
               Width           =   315
            End
            Begin VB.ComboBox CmbCond 
               Height          =   315
               Left            =   3465
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   2100
               Width           =   3015
            End
            Begin VB.CommandButton CmdVuelto 
               Caption         =   "Aceptar"
               Height          =   375
               Index           =   0
               Left            =   5070
               TabIndex        =   30
               Top             =   3705
               Width           =   1470
            End
            Begin VB.TextBox TxtPago 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   5280
               TabIndex        =   23
               Text            =   "TxtPago"
               Top             =   1185
               Width           =   1215
            End
            Begin VB.ComboBox CmbCompro 
               Height          =   315
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   1665
               Width           =   3015
            End
            Begin VB.ComboBox CmbForma 
               Height          =   315
               Left            =   3465
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   2550
               Width           =   2625
            End
            Begin VB.Label LblCliente 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblCliente"
               Height          =   330
               Left            =   2190
               TabIndex        =   48
               Top             =   270
               Width           =   3960
            End
            Begin VB.Label Label18 
               Caption         =   "F5"
               Height          =   270
               Left            =   135
               TabIndex        =   46
               Top             =   315
               Width           =   345
            End
            Begin VB.Label Label8 
               Caption         =   "F.Pago :"
               Height          =   255
               Left            =   645
               TabIndex        =   43
               Top             =   2115
               Width           =   1095
            End
            Begin VB.Label lblImpre 
               Height          =   255
               Left            =   75
               TabIndex        =   42
               Top             =   3045
               Width           =   495
            End
            Begin VB.Label Label7 
               Caption         =   "Cliente :"
               Height          =   255
               Left            =   630
               TabIndex        =   41
               Top             =   300
               Width           =   780
            End
            Begin VB.Label LblVuelto 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblVuelto"
               Height          =   330
               Left            =   5280
               TabIndex        =   40
               Top             =   3015
               Width           =   1215
            End
            Begin VB.Label LblCompra 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblCompra"
               Height          =   330
               Left            =   5265
               TabIndex        =   39
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Su Vuelto :"
               Height          =   255
               Left            =   3480
               TabIndex        =   38
               Top             =   3015
               Width           =   1155
            End
            Begin VB.Label Label5 
               Caption         =   "Su Pago :"
               Height          =   255
               Left            =   675
               TabIndex        =   37
               Top             =   1185
               Width           =   1455
            End
            Begin VB.Label Label4 
               Caption         =   "Total de su Compra :"
               Height          =   255
               Left            =   675
               TabIndex        =   36
               Top             =   720
               Width           =   2415
            End
            Begin VB.Label Label9 
               Caption         =   "Compr. :"
               Height          =   255
               Left            =   690
               TabIndex        =   35
               Top             =   1650
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "F9 - Regresa"
               Height          =   210
               Left            =   135
               TabIndex        =   34
               Top             =   3795
               Width           =   1200
            End
            Begin VB.Label Label19 
               Caption         =   "F11"
               Height          =   270
               Left            =   135
               TabIndex        =   33
               Top             =   2115
               Width           =   390
            End
            Begin VB.Label Label17 
               Caption         =   "F7"
               Height          =   270
               Left            =   135
               TabIndex        =   32
               Top             =   1185
               Width           =   345
            End
            Begin VB.Label Label15 
               Caption         =   "F6"
               Height          =   270
               Left            =   135
               TabIndex        =   29
               Top             =   1650
               Width           =   345
            End
            Begin VB.Label Label20 
               Caption         =   "F12"
               Height          =   270
               Left            =   135
               TabIndex        =   27
               Top             =   2535
               Width           =   390
            End
            Begin VB.Label Label21 
               Caption         =   "C.Pago :"
               Height          =   255
               Left            =   630
               TabIndex        =   25
               Top             =   2550
               Width           =   1095
            End
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   180
            Left            =   210
            TabIndex        =   31
            Top             =   4605
            Width           =   6810
            _ExtentX        =   12012
            _ExtentY        =   318
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin VB.TextBox TxtId 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "TxtCa"
         Top             =   6525
         Width           =   495
      End
      Begin VB.TextBox TxtProducto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1365
         MaxLength       =   25
         TabIndex        =   0
         Top             =   5445
         Width           =   1290
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1635
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "TxtCa"
         Top             =   5820
         Width           =   930
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "TxtPrecio"
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5790
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   5250
         Left            =   180
         TabIndex        =   5
         Top             =   150
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   9260
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
      Begin VB.PictureBox Picture1 
         Height          =   150
         Left            =   7200
         ScaleHeight     =   90
         ScaleWidth      =   90
         TabIndex        =   49
         Top             =   180
         Width           =   150
      End
      Begin VB.Label Label22 
         Caption         =   "F8 - % Desc. :"
         Height          =   210
         Left            =   3690
         TabIndex        =   44
         Top             =   5775
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "F7 - Borrar x Id"
         Height          =   210
         Left            =   1425
         TabIndex        =   15
         Top             =   6525
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Supr - Borrar Ultimo"
         Height          =   210
         Left            =   6060
         TabIndex        =   14
         Top             =   5790
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "F5 - Cerrar"
         Height          =   210
         Left            =   195
         TabIndex        =   13
         Top             =   6525
         Width           =   1200
      End
      Begin VB.Label TxtDetalle 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2700
         TabIndex        =   12
         Top             =   5460
         Width           =   3525
      End
      Begin VB.Label Label12 
         Caption         =   "F9 - Codigo :"
         Height          =   210
         Left            =   165
         TabIndex        =   11
         Top             =   5475
         Width           =   1200
      End
      Begin VB.Label Label13 
         Caption         =   "F11 - Cantidad :"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   5820
         Width           =   1410
      End
      Begin VB.Label Label14 
         Caption         =   "F12 - Precio :"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   6135
         Width           =   1290
      End
      Begin VB.Label Label10 
         Caption         =   "Neto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6315
         TabIndex        =   8
         Top             =   5475
         Width           =   525
      End
      Begin VB.Label Label16 
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3960
         TabIndex        =   7
         Top             =   6195
         Width           =   660
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   4665
         TabIndex        =   6
         Top             =   6195
         Width           =   3345
      End
   End
   Begin VB.Label LblDescuento 
      Caption         =   "LblDescuento"
      Height          =   240
      Left            =   120
      TabIndex        =   50
      Top             =   7905
      Width           =   1680
   End
   Begin VB.Label LblTotalIva 
      Caption         =   "LblTotalIva"
      Height          =   195
      Left            =   3720
      TabIndex        =   19
      Top             =   7545
      Width           =   1485
   End
   Begin VB.Label LblTotalNoGrabado 
      Caption         =   "LblTotalNoGrabado"
      Height          =   225
      Left            =   1770
      TabIndex        =   18
      Top             =   7530
      Width           =   1695
   End
   Begin VB.Label LblTotalGrabado 
      Caption         =   "LblTotalGrabado"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7530
      Width           =   1545
   End
End
Attribute VB_Name = "VentaRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rsd As ADODB.Recordset, sw As Boolean, nId As Integer, nListaVenta As Integer
Dim bCod As Boolean, bBoto As Boolean, bBte As String, bPrec As String * 2
Dim mPrec As Single, Cadena As String, bCompro As Boolean, vCliente As Long
'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private KeyRetroceso As Boolean
Dim sConsFinal As String, sNomFinal As String, nMoviAfip As Long, sCondPago As String, dTotalAux As Double
Dim bVerificar As Boolean
Dim nUnitario() As Double
Dim bF11 As Boolean, bCajaDesc As Boolean
Dim bPercibe As Boolean, nPrecioEnvase As Double
Dim bEnvase As Boolean, bDejo As Boolean
Dim sCae As String, sVencimiento As String, sNumDoc As String, nNumCompro As Long
Dim cQrCode As ClsQrCode, bError As Boolean
Dim imp_iva_fact As Double, imp_tributo_fact As Double



Private Sub BtnBuscar_Click()
  Dim RsC As ADODB.Recordset, cRsl As ClsLectura
  On Error GoTo errHandler
  
  Dim dTotal As Double
  
  dTotal = dTotalAux
  
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     TxtCliente.text = RsC!Cliente
     LblCliente.Caption = RsC!RazonSocial
     CmbCompro.text = cRsl.TraerComproCliente(TxtCliente.text, Frame2.Caption)
     If TxtCliente.text = "." Then
        CmbCond.text = "Contado"
     Else
        CmbCond.text = sCondPago
     End If
     
     dTotal = dTotalAux - CalcularDescuento(dTotal)
     
     LblTotal.Caption = Format(dTotal, "#0.00")
     LblCompra.Caption = LblTotal.Caption
     TxtPago.SetFocus
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
End Sub

Private Sub CmbCompro_Click()
  Dim dTotal As Double
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura
  If Frame1.Visible = True Then
     If CmbCompro.ListIndex <> -1 Then
'        If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = 5 And cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 0 Then
'           AplicarFactura.Show 1
'        End If
        If cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) <> 0 Then
            Frame2.Caption = "P"
        Else
           If Frame2.Caption = "P" Then
              Frame2.Caption = "N"
           End If
        End If
    End If
  End If
 Set cRsl = Nothing
End Sub

Private Sub CmbCond_Click()
  If VerificarCuenta(TxtCliente.text) = False Then
       bVerificar = False
       Frame2.Visible = False
       HabilitarDetalles bCod, False
       CmdDetalle(0).Enabled = bBoto
       TxtCliente.text = sConsFinal
       LblCliente.Caption = sNomFinal
       CmbCond.text = sCondPago
       CmbForma.text = "EFECTIVO"
       LblTotal.Caption = Format(dTotalAux, "#0.00")
       Grid1.Enabled = True
    Else
       bVerificar = True
       Dim dTotal As Double, cRsl As ClsLectura
  
       dTotal = dTotalAux
  
       Set cRsl = New ClsLectura
       
       If TxtCliente <> "." Then
          If CmbCond.text <> "Contado" Then
             Recalcular cRsl.TraerRecargoCliente(TxtCliente.text)
           Else
              Recalcular 0
           End If
           If CmbCond.text = "Contado" Then
              Recalcular -cRsl.TraerDescuentoCliente(TxtCliente.text)
           End If
       Else
          Recalcular 0
       End If
       dTotal = dTotalAux - CalcularDescuento(dTotal)
       LblTotal.Caption = Format(dTotal, "#0.00")
       LblCompra.Caption = LblTotal.Caption
       TxtPago_LostFocus
    End If

End Sub

Private Sub CmbForma_Click()
    Dim dTotal As Double, cRsl As ClsLectura
    
    Set cRsl = New ClsLectura
    dTotal = dTotalAux
  
    dTotal = dTotalAux - CalcularDescuento(dTotal)
    LblTotal.Caption = Format(dTotal, "#0.00")
    
    TxtPago_LostFocus
  
    Recalcular cRsl.TraerValorDeUnCampo("CondVenta", "RgTarjeta", "CondVta=" & CmbForma.ItemData(CmbForma.ListIndex))

    LblCompra.Caption = LblTotal.Caption

End Sub

Private Sub CmdDetalle_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 3
              Salir
  End Select
End Sub

Private Sub Salir()

End Sub

Private Sub Nuevo()
  If sw = True Then
     sw = False
     HabilitarDetalles True, False
     LimpiarDetalles
     BotonDetalles False, True
     CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(0).Caption = "Grabar"
     TxtProducto.SetFocus
  Else
     If Grabar = True Then
        bF11 = False
        CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
        CmdDetalle(0).Caption = "Agregar"
        BotonDetalles True, False
        sw = True
        HabilitarDetalles False, False
     End If
  End If
End Sub

Private Function Grabar() As Boolean
  On Error GoTo Errores
  If TxtPrecio.text <> 0 Then
     Rsd.AddNew
     Rsd!ID = nId
     Rsd!Producto = TxtProducto.text
     Rsd!Descripcion = TxtDetalle.Caption
     
     If TxtProducto.text = sEnvase Then
        TxtCantidad.text = -(TxtCantidad.text)
     End If
     
     If bPrec = "Si" Then
       TxtCantidad.text = Format(CCur(TxtPrecio.text) / mPrec, "0.000")
       Rsd!Cantidad = CDbl(TxtCantidad.text)
       Rsd!PUnitario = mPrec
       Rsd!Total = TxtCantidad.text * mPrec 'CCur(TxtPrecio.Text)
     Else
        Rsd!Cantidad = Round(CSng(TxtCantidad.text), 4)
        Rsd!PUnitario = Round(CDbl(TxtPrecio.text), 4)
        Rsd!Total = Round(CDbl(TxtCantidad.text) * TxtPrecio.text, 4)
     End If
     Rsd.Update
     nId = nId + 1
     CalcularTotales Rsd, True

     If nPrecioEnvase <> 0 Then
        Rsd.AddNew
        Rsd!ID = nId
        Rsd!Producto = "ENV"
        Rsd!Descripcion = tEnvase
        Rsd!Cantidad = Round(CSng(TxtCantidad.text), 4)
        Rsd!PUnitario = nPrecioEnvase
        Rsd!Total = Round(CDbl(TxtCantidad.text) * nPrecioEnvase)
        Rsd.Update
        nId = nId + 1
        nPrecioEnvase = 0
        CalcularTotales Rsd, True
     End If
     
     GrabarTxtRsd Rsd
     
     Grabar = True
     bDejo = True
  End If
Exit Function
Errores:
  If Rsd.RecordCount <> 0 Then
     Rsd.CancelUpdate
  End If
   Grabar = False
End Function

Private Sub LimpiarDetalles()
  TxtId.text = ""
  TxtCPrecio.text = ""
  TxtProducto.text = ""
  TxtDetalle.Caption = ""
  TxtCantidad.text = "1.00"
  TxtPrecio.text = "0.000"
  If ConfiguracionArgentina = True Then
     TxtCantidad.text = "1,00"
     TxtPrecio.text = "0,000"
  End If
End Sub

Private Sub ComprobarPercepcion(pCuit As String)
 On Error GoTo errHandler
   Dim cRx As ClsLectura, cPer As DsPadronA4, bResul As Boolean
   Set cRx = New ClsLectura
   Set cPer = New DsPadronA4
   
   bResul = cPer.SolicitarIB(pCuit)
   If bResul = True Then
      db.Open sDb
      db.BeginTrans
      db.Execute "UPDATE Clientes Set Percepcion =" & Replace(cPer.Percepcion, ",", ".") & " WHERE Cuit='" & pCuit & "'"
      db.CommitTrans
      db.Close
   End If
Exit Sub
errHandler:
        db.RollbackTrans
        db.Close
End Sub

Private Sub CmdVuelto_Click(Index As Integer)

 On Error GoTo Errores


  CmdVuelto(0).Enabled = False
  
  
  sCae = ""
  sVencimiento = ""
  sNumDoc = ""
  nNumCompro = 0
  
  imp_iva_fact = 0
  imp_tributo_fact = 0
  
  Dim cRsl As ClsLectura, cRle As ClsEscritura, cRv As ADODB.Recordset, nMovimiento As Long
  Dim nInter As Single, sInter As String, nTasa As Single, nInterAux As Single
  Dim cCuit As String, cDire As String, nComp As Long
  Dim nNumero As Long, bCierreControlador As Boolean, bGraboTick As Boolean
  Dim nTotal, nGrabado, nIvas, nNoGrabado As Currency
  Dim cNum As ClsComprobantesL
  Dim nAlicuota As Single, nPercepcion As Currency, sPercep As String
  Dim cPe As ClsComprobantesE
  Dim pInterPorc As Double, vInterno As Double
  
  Dim X As Integer, Y As Single
  Dim P As Printer, sImpre As String

  nGrabado = 0
  Set cNum = New ClsComprobantesL
  Set cPe = New ClsComprobantesE
  
  nAlicuota = 0
  nPercepcion = 0
  
  bCierreControlador = False
  bGraboTick = False
  
  Set cRsl = New ClsLectura
  

  
  
  If bPercibe = True And Frame2.Caption = "P" And TxtCliente.text <> "." Then
    cCuit = cRsl.TraerValorDeUnCampo("Clientes", "Cuit", "Cliente='" & TxtCliente.text & "'")
    If cCuit <> "" Then
        ComprobarPercepcion (cCuit)
        nAlicuota = cRsl.TraerValorDeUnCampo("Clientes", "Percepcion", "Cliente='" & TxtCliente.text & "'")
    Else
        MsgBox "El Cliente no tiene el Cuit Cargado...", vbCritical, "Atencion"
         Frame2.Caption = "N"
         CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinalN"))
         CmbCompro.Enabled = False
         TxtCliente.text = sConsFinal
         LblCliente.Caption = sNomFinal
         If TxtPago.Visible = True Then
            TxtCliente.SetFocus
            TxtCliente.SelStart = 0
            TxtCliente.SelLength = Len(TxtCliente.text)
         End If
         Exit Sub
    End If
  End If
  
  nPercepcion = CCur(LblTotalNoGrabado.Caption) * nAlicuota / 100
     
  Dim FS As String
  Dim sComando As String

  Dim pParam(1) As String, pRs(0) As ADODB.Recordset
  
  

  Set cRsl = New ClsLectura
  Set cRle = New ClsEscritura
  
  Set cRv = cRsl.RsVacio("Ventas", "Id", "N")
  
  FS = Chr$(28)
      
  BtnBuscar.Enabled = False
  TxtCliente.Enabled = False
  CmdVuelto(0).Enabled = False
  CmbCond.Enabled = False
     
  nMovimiento = cRsl.ActualizarMovimientoTraer
  nMovimiento = nMovimiento + (cRsl.RegPorDefecto("Caja") * 10000000)
  nMoviAfip = nMovimiento

  
  If Rsd.RecordCount <> 0 Then
  
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        cRv.AddNew
        cRv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
        cRv!Movimiento = nMovimiento
        cRv!fecha = Date
        cRv!Cliente = IIf(TxtCliente.text = "", ".", TxtCliente.text)
        cRv!Lista = nListaVenta
        cRv!Vendedor = cRsl.RegPorDefecto("Vendedor")
        cRv!Caja = cRsl.RegPorDefecto("Caja")
        cRv!formapago = CmbForma.ItemData(CmbForma.ListIndex)
        cRv!CondVenta = CmbCond.ItemData(CmbCond.ListIndex)
        If Frame2.Caption = "N" Or Frame2.Caption = "J" Then
           nNumero = cRsl.TraerValorDeUnCampo("Comprobantes", "Numero", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
           bCompro = True
           cRv!Comprobante = CmbCompro.ItemData(CmbCompro.ListIndex)
           cRv!Numero = nNumero
           cRv!sucursal = cRsl.TraerValorDeUnCampo("Comprobantes", "Sucursal", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
           cRv!TotalGrabado = CCur(LblTotal.Caption)
           cRv!TotalNoGrabado = 0
           cRv!TotalIva = 0
           cRv!Grabado = 0
           cRv!Impuestos = 0
           cRv!NoGrabado = 0
           cRv!Percepcion = 0
           cRv!TotalGeneral = CCur(LblTotal.Caption)
           nComp = CmbCompro.ItemData(CmbCompro.ListIndex)
         Else
           bCompro = False
           cRv!Comprobante = CmbCompro.ItemData(CmbCompro.ListIndex)
           cRv!Numero = 0
           cRv!sucursal = cRsl.TraerValorDeUnCampo("Comprobantes", "Sucursal", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
         
           cRv!TotalNoGrabado = CCur(LblTotalNoGrabado.Caption)
           cRv!TotalIva = CCur(LblTotalIva.Caption)
           cRv!TotalGeneral = CCur(LblTotal.Caption)
           sInter = cRsl.TraerValorDeUnCampo("Productos", "TipoInterno", "Producto='" & Rsd!Producto & "'")
           Dim sInter1 As String
           If Rsd!Producto = "ENV" Then
              nTasa = 21
              cRv!Tasa = 15
              sInter = "No"
           Else
              nTasa = cRsl.TraerTasaProducto(Rsd!Producto)
              cRv!Tasa = cRsl.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & Rsd!Producto & "'")
           End If
           pInterPorc = 0: vInterno = 0
           Select Case sInter
                  Case "No"
                        nInter = 0
                        sInter = "0"
                  Case "Porcentaje"
                        sInter = ""
                        sInter = "0"
                        pInterPorc = cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & Rsd!Producto & "'") '  / 100) + 1
                        vInterno = (Rsd!PUnitario * pInterPorc) / 100
                        nInter = vInterno
                        nInterAux = nInterAux + (nInter * Rsd!Cantidad)
                  Case "Fijo"
                       nInter = cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & Rsd!Producto & "'")
                       sInter = Format(Int(nInter), "0000000")
                       sInter1 = sInter & Replace(nInter, "0.", "") & Mid("00000000", 1, Len(Replace(nInter, "0.", "")))
           End Select
           Rsd!Total = Rsd!Total - CalcularDescuento(Rsd!Total)
           cRv!Impuestos = (Rsd!Total - (Rsd!Cantidad * nInter)) - (Rsd!Total - (Rsd!Cantidad * nInter)) / ((nTasa / 100) + 1)
          
           cRv!Grabado = (Rsd!PUnitario - nInter) / ((nTasa / 100) + 1) * Rsd!Cantidad
          nGrabado = nGrabado + cRv!Grabado
           cRv!Impuestos = cRv!Grabado * nTasa / 100
           cRv!TotalGrabado = nGrabado
           cRv!NoGrabado = nNoGrabado + (Rsd!Cantidad * nInter)
           nComp = CmbCompro.ItemData(CmbCompro.ListIndex)
        End If
        cRv!Producto = Rsd!Producto
        cRv!Descripcion = Rsd!Descripcion
        cRv!Cantidad = Rsd!Cantidad
        If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = "0" Then
           cRv!PUnitario = Rsd!PUnitario - CalcularDescuento(Rsd!PUnitario)
        Else
          cRv!PUnitario = (Rsd!PUnitario - nInter) / ((nTasa / 100) + 1)
        End If
        
        cRv!Percepcion = nPercepcion
        cRv!pTotal = Rsd!Total + nPercepcion
        cRv!Anulado = 0
        cRv.Update
        Rsd.MoveNext
     Loop
     
     Dim sCondiIva As String, nFaNo As String, pCuit As String
     Dim sCant, sPrecio

     If Frame2.Caption = "P" Then
     
'         Dim i As Byte
     
        Dim ver As String, fecha As String, Cuit As Double, ptoVta As Integer, tipoComp As Integer, nroCmp As Long, importeAfip As Double, Moneda As String
        Dim ctz As Double, tipoDocRec As Integer, nroDocRec As Double, tipoCodAut As String, codAut As Double
         
        ' Factura electronica
        Dim sNum As String, sRazon As String, sDire As String, sLocal As String, sTel As String
        Dim sCuit As String, sIva As String, sCondi As String, sProv As String
        Dim nPoc As String, nCosto As Byte, nComproAfip As Byte
  
      
        sNum = cRsl.TraerValorDeUnCampo("Comprobantes", "Sucursal", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
        sRazon = LblCliente.Caption
        sDire = cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & TxtCliente.text & "'")
        sCuit = cRsl.TraerValorDeUnCampo("Clientes", "CUIT", "Cliente='" & TxtCliente.text & "'")
        sLocal = cRsl.TraerValorDeUnCampo("Clientes", "CodigoPostal", "Cliente='" & TxtCliente.text & "'") & " - " & cRsl.TraerValorDeUnCampo("Clientes", "Ciudad", "Cliente='" & TxtCliente.text & "'")
        sProv = cRsl.TraerValorDeUnCampo("Provincias", "Descripcion", "Provincia=" & cRsl.TraerValorDeUnCampo("Clientes", "Provincia", "Cliente='" & TxtCliente.text & "'"))
        
        Dim tIva As Byte
              
        tIva = cRsl.TraerValorDeUnCampo("Clientes", "TipoIva", "Cliente='" & TxtCliente.text & "'")

        sIva = cRsl.TraerValorDeUnCampo("TipoIva", "Descripcion", "TipoIva=" & tIva)

       ' C Consumidor final E Exento M Monotributo
        Dim sCondicion As String
  
       nCosto = cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
       nComproAfip = cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
       
       tipoComp = nComproAfip
       
        Select Case nComproAfip
       
           Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 51, 52, 53, 201
                Dim Respuesta As Boolean
                
                Dim vIva(3) As String, vPorcIva(3) As Byte, nNeto As Single, nIva As Single, nTDoc As Byte
                
                
                nTDoc = 80
                If TxtCliente.text = "." Or tIva = 1 Then
                  nTDoc = 96
                End If
                
                                     
                vIva(0) = 0: vIva(1) = 0: vIva(2) = 0: vIva(3) = 0
                vPorcIva(0) = 0: vPorcIva(1) = 4: vPorcIva(2) = 5: vPorcIva(3) = 6
              
                cRv.MoveFirst
                
             
                    
                 nNeto = nGrabado ' cRv!TotalGrabado    '  LblDescuento.Caption
                
                 If nComproAfip >= 6 And nComproAfip <= 8 And nIva <> 0 Then
                   nNeto = cRv!TotalGeneral - cRv!NoGrabado - cRv!Percepcion - cRv!TotalIva
                   nIva = nIva
                End If

                If nComproAfip >= 11 And nComproAfip <= 15 Then
                   nNeto = cRv!TotalGeneral
                   nIva = 0
                End If
'
                Do While Not cRv.EOF
                    nIva = nIva + cRv!Impuestos
                    cRv.MoveNext
                Loop
                
               
'
                If cRsl.TraerCantidad("AfipCaes", "Id", "Movimiento=" & nMoviAfip) = 0 Then
                   If TxtCliente.text = "." And LblTotal.Caption > 10000000 Then
                       FrmDocumento.Show 1
                       If nDocAfip = 0 Then
                          Exit Sub
                       End If
                       sCuit = nDocAfip
                   Else
                       If sCuit = "" And TxtCliente.text <> "." And tIva <> 1 Then
                          MsgBox "Cliente sin Número de CUIT", vbCritical, "Atención"
                          Exit Sub
                       Else
                          If nTDoc <> 80 Then
                             nTDoc = 99
                             sCuit = "000000000000"
                          End If
                       End If
                   End If
                   
                   tipoDocRec = nTDoc
                   nroDocRec = sCuit
                   ptoVta = sNum
                   Cuit = sCuitAfip
                   
                   Respuesta = Autorizar(nComproAfip, nConceptoAfip, nTDoc, sCuit, LblTotal.Caption, str(nNeto), str(nIva), Val(sNum), nMoviAfip, nCosto, cRv, str(nInterAux), "0")
                   If bError = True Then
                      bError = False
                      Exit Sub
                   End If
                Else

                End If
                If nCosto = 0 Then
                   nIva = 0
                   nNeto = LblTotal.Caption
                End If
                
                 Dim rDatos As ADODB.Recordset
    
                 Set rDatos = New ADODB.Recordset
                  
                 Set rDatos = cRsl.TraerRsCondi("AfipDatos", "Id_Datos", "Id_Datos=" & Val(sNum))
                  
                  If rDatos.RecordCount = 0 Then
                     MsgBox "Debe Definir las Sucursales en AFIP Datos", vbInformation, "Atención"
                     Exit Sub
                  End If
                  
                    
                sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & CmbCompro.ItemData(CmbCompro.ListIndex) & " AND Usuario=" & nUsuario)
        
                 If sImpre <> "Ninguna" And sImpre <> "0" And bAnulaPres <> True Then
                                            
                    For Each P In Printers
                        If P.DeviceName = sImpre Then
                           Set Printer = P
                              Printer.FontName = "Tahoma"
                         '       Printer.FontName = "Courier New"
                           Printer.ScaleMode = 7
                           Printer.FontSize = 9
                           Exit For
                        End If
                    Next
                 End If
                 
                 Dim espaciado As Single
                 
                 espaciado = 0.35
                                   
                 Printer.FontSize = 12
                 Y = 1
                 X = 0
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.FontBold = True
                 Printer.Print sNomFan
                 Printer.FontBold = False
              
                 Printer.FontSize = 9
                 
                 Y = Y + 1
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print Trim(rDatos!Empresa)
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "C.U.I.T. Nro.: "; Trim(rDatos!Cuit)
                 
                  Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "INGRESOS BRUTOS: "; Trim(rDatos!IIBB)
                 
                  Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print Trim(rDatos!Direccion)
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print Trim(rDatos!Localidad)
                 
                  Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "INICIO DE ACTIVIDADES: " & Trim(rDatos!Inicio)
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print Trim(rDatos!Iva)
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "---------------------------------------------------------------"
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = 4 - (Len(CmbCompro.text) / 5)
                 Printer.FontBold = True
                 Printer.Print CmbCompro.text
                 Printer.FontBold = False
                 
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = 1.7
                 Printer.Print "ORIGINAL (Cod." & Format(nComproAfip, "000") & ")"
                 
                 
                  Y = Y + 0.5
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "Nro.: " & Format(Val(sNum), "0000") & "-" & Format(nNumCompro, "00000000") & "       Fecha:" & Date
                  
      
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "---------------------------------------------------------------"
                 
                
                 
                  Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print sRazon
                 
                  Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print sDire
                 
                  Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print sCuit
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print sIva
                 
                  Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print sLocal
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print sProv
                 
                 Y = Y + espaciado
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "---------------------------------------------------------------"
                                  
                 
                 
                 cRv.MoveFirst
                 
                 Printer.FontName = "Courier New"
                 Printer.FontBold = True
                 Do While Not cRv.EOF
                 
                    Dim nTotDet As Double
                    Dim cantCanti As Byte
                    Dim cantTotdet As Byte
                    Dim nAdi As Byte
                    

                    nTotDet = cRv!PUnitario * cRv!Cantidad

                    cantTotdet = Len(Format(nTotDet, "#0.00"))
                    cantCanti = Len(cRv!Cantidad & " x " & Format(cRv!PUnitario, "#0.00"))
                    
                    Y = Y + espaciado
                    If Y >= 28 Then
                        Y = 0
                        Printer.NewPage
                    End If
                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print Left(cRv!Descripcion, 25)
                    
                    Y = Y + espaciado
                    If Y >= 28 Then
                        Y = 0
                        Printer.NewPage
                    End If
                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print cRv!Cantidad & " x " & Format(cRv!PUnitario, "#0.00") & Space(19 - cantCanti) & Space(12 - cantTotdet) & Format(nTotDet, "$0.00")
                    Y = Y + espaciado
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If
                    cRv.MoveNext
                 Loop
                 
               
     
                 Printer.FontBold = False
                                            
                 Printer.FontName = "Tahoma"
                 
                 Y = Y + 0.7
                 If Y >= 28 Then
                    Y = 0
                    Printer.NewPage
                 End If

                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "---------------------------------------------------------------"
                                                                 
                 cRv.MoveFirst
                 
                 If nCosto = 0 Then
                    Y = Y + espaciado
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If
                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print "       TOTAL: " & Space(40 - Len(Format(cRv!TotalGeneral, "$0.00"))) & Format(cRv!TotalGeneral, "$0.00")
                    
                    Y = Y + espaciado
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If
                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print "---------------------------------------------------------------"
                 Else
                      Dim nTasaI(3) As Double, ID As Double, base_imp As Double, importe As Double
                      Dim sBaseCamI(3) As Currency, sImporteI(3) As Currency, h As Byte

                      Printer.FontName = "Courier New"
                      Printer.FontBold = True

                      Y = Y + espaciado
                      If Y >= 28 Then
                         Y = 0
                         Printer.NewPage
                      End If
                      Printer.CurrentY = Y
                      Printer.CurrentX = X
                      Printer.Print "     SUBTOTAL: " & Space(17 - Len(Format(nGrabado, "$0.00"))) & Format(nGrabado, "$0.00")

                      Do While Not cRv.EOF
                          Select Case cRv!Tasa
                              Case 9
                                 nTasaI(0) = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & cRv!Tasa)
                                 sBaseCamI(0) = sBaseCamI(0) + (cRv!PUnitario)  ' - Rsd!Descuento
                                 sImporteI(0) = sImporteI(0) + cRv!Impuestos
                              Case 10
                                 nTasaI(1) = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & cRv!Tasa)
                                 sBaseCamI(1) = sBaseCamI(1) + (cRv!PUnitario)  ' - Rsd!Descuento
                                 sImporteI(1) = sImporteI(1) + cRv!Impuestos
                              Case 15
                                 nTasaI(2) = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & cRv!Tasa)
                                 sBaseCamI(2) = sBaseCamI(2) + (cRv!PUnitario)  '- Rsd!Descuento
                                 sImporteI(2) = sImporteI(2) + cRv!Impuestos
                              Case 16
                                 nTasaI(3) = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & cRv!Tasa)
                                 sBaseCamI(3) = sBaseCamI(3) + (cRv!PUnitario) ' - Rsd!Descuento
                                 sImporteI(3) = sImporteI(3) + cRv!Impuestos
                         End Select
                         cRv.MoveNext
                     Loop
                     For h = 0 To 3
                        If sImporteI(h) <> 0 Or nTasaI(h) = 3 Then
                           ID = nTasaI(h)
                           base_imp = Round(sBaseCamI(h), 2) - IIf(nCosto = 0, sImporteI(h), 0)
                           base_imp = Round(base_imp, 2)
                           importe = Format(sImporteI(h), "#0.00")
                           

                           Y = Y + espaciado
                           If Y >= 28 Then
                              Y = 0
                              Printer.NewPage
                           End If
                           Printer.CurrentY = Y
                           Printer.CurrentX = X
                           Printer.Print "     IVA " & Format(ID, "#0.00") & "% : " & Space(14 - Len(Format(importe, "$0.00"))) & Format(importe, "$0.00")
 
                        End If
                    Next
                    
                    If nInterAux <> 0 Then
                        Y = Y + espaciado
                        If Y >= 28 Then
                           Y = 0
                          Printer.NewPage
                        End If
                        Printer.CurrentY = Y
                        Printer.CurrentX = X
                        Printer.Print "     IMP. INTERNOS : " & Space(11 - Len(Format(nInterAux, "$0.00"))) & Format(nInterAux, "$0.00")
                    End If
                    
                    Printer.FontBold = False
                                              
                    Printer.FontName = "Tahoma"


                    cRv.MoveFirst
                    
                    Y = Y + 1
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If
                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print "     TOTAL: " & Space(42 - Len(Format(cRv!TotalGeneral, "$0.00"))) & Format(cRv!TotalGeneral, "$0.00")

                    Y = Y + espaciado
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If
                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print "---------------------------------------------------------------"

                 End If
                 
                 Y = Y + espaciado
                 
                 If nCosto = 0 And bMono = False Then
                     
                     Printer.CurrentY = Y
                     Printer.CurrentX = X
                     Printer.Print "Régimen de Transp. Fiscal al Cons.(Ley 27743)"
                     Y = Y + espaciado
                     
                     Printer.CurrentY = Y
                     Printer.CurrentX = X
                     Printer.Print "IVA Contenido :" & Format(imp_iva_fact, "#0.00")
                     Y = Y + espaciado
                     
                     Printer.CurrentY = Y
                     Printer.CurrentX = X
                     Printer.Print "Otros Imp. Nacionales Indirectos :" & Format(imp_tributo_fact, "#0.00")
                     Y = Y + espaciado
                     Y = Y + espaciado
                 End If
                 
                 If Y >= 28 Then
                    Y = 0
                    Printer.NewPage
                 End If
                 
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "REFERENCIA ELEC. DEL COMPROBANTE"
                 
                 Y = Y + espaciado
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If
                 
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "C.A.E. Nro. " & sCae & "  Vto. " & Right(sVencimiento, 2) & "/" & Mid$(sVencimiento, 5, 2) & "/" & Mid$(sVencimiento, 1, 4)

                 If nCosto = 0 Then
                   
                    Y = Y + espaciado
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If

                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print "Orientación al Consumidor Prov.Bs.As"
                    Y = Y + espaciado
                    
                    Printer.CurrentY = Y
                    Printer.CurrentX = X
                    Printer.Print "0800-222-9042"
                    Y = Y + espaciado
                                        
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If
                    
                   
                 End If

                 Y = Y + 0.5
 
                    If Y >= 28 Then
                       Y = 0
                       Printer.NewPage
                    End If

                 X = 0.5
                 
                 Dim sDate As String
                 sDate = Format(Date, "yyyy-mm-dd")
    
                 
                  Dim afipUrl As String, valorAfip As String, valor64 As String, params As String
                  ver = 1
                  fecha = sDate
                  nroCmp = nNumCompro
                  importeAfip = Format(cRv!TotalGeneral, "$0.00")
                  Moneda = "PES": ctz = 1
                  codAut = CDbl(sCae)
                                  
                  params = "{""ver"":" & ver & _
                           ", ""fecha"":" & Chr(34) & fecha & Chr(34) & _
                           ", ""cuit"":" & Cuit & _
                           ", ""ptoVta"":" & ptoVta & _
                           ", ""tipoCmp"":" & tipoComp & _
                           ", ""nroCmp"":" & nroCmp & _
                           ", ""importe"":" & importeAfip & _
                           ", ""moneda"":" & Chr(34) & Moneda & Chr(34) & _
                           ", ""ctz"":" & ctz & _
                           ", ""tipoDocRec"":" & tipoDocRec & _
                           ", ""nroDocRec"":" & nroDocRec & _
                           ", ""tipoCodAut"":" & Chr(34) & "E" & Chr(34) & _
                           ", ""codAut"":" & codAut & "}"
                 
                 valor64 = EncodeBase64(StrConv(params, vbFromUnicode))
                 afipUrl = "https://www.afip.gob.ar/fe/qr/?p=" & valor64
                                                                   
                 Picture1.Picture = cQrCode.GetPictureQrCode(afipUrl, 150, 150)
                 Printer.PaintPicture Picture1.Picture, X, Y
                                               
                 Y = Y + 2
           
                 If Y >= 28 Then
                    Y = 0
                    Printer.NewPage
                 End If

                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 Printer.Print "-"
                                                                                                                                        
                 Printer.EndDoc
        End Select
       
     Else
         Dim nCan As Currency, nUni As Currency, nTot As Currency, sDesProduc As String
         Dim i As Integer
                    
          sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & CmbCompro.ItemData(CmbCompro.ListIndex) & " AND Usuario=" & nUsuario)
          If Frame2.Caption = "J" Then
             sImpre = "Ninguna"
          End If
          If sImpre <> "Ninguna" And sImpre <> "0" And bAnulaPres <> True Then
                                            
              For Each P In Printers
                  If P.DeviceName = sImpre Then
                     Set Printer = P
                     Printer.FontName = "Courier New"
                     Printer.ScaleMode = 7
                     Printer.FontSize = 9
                     Exit For
                  End If
              Next
                                 
              Printer.FontSize = 14
                                 
              Y = 1
              X = 0
              
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.FontBold = True
              Printer.Print sNomFan
              Printer.FontBold = False
              
              Printer.FontSize = 7
              
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print Date & " - " & Time
                        
              
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print CmbCompro.text; " N. :" & Format(nSucursal, "0000") & "-" & Format(nNumero, "00000000")
              
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "Cliente: " & LblCliente.Caption
              
              Y = Y + 0.3
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "Domicilio: " & cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & TxtCliente.text & "'")
              
              Y = Y + 0.4
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print CmbCond.text
                          
              Rsd.MoveFirst
                
              Y = Y + 0.5
                                         
                
           '   Printer.FontName = "Courier New"
             Printer.FontBold = True
                
              Do While Not Rsd.EOF
                 Printer.CurrentY = Y
                 Printer.CurrentX = X
                 
                 
                 nTot = Rsd!PUnitario * Rsd!Cantidad

                 cantTotdet = Len(Format(nTot, "#0.00"))
                 cantCanti = Len(Rsd!Cantidad & " x " & Format(Rsd!PUnitario, "#0.00"))
                  
                  Y = Y + 0.3
                  If Y >= 28 Then
                     Y = 0
                     Printer.NewPage
                 End If
                  Printer.CurrentY = Y
                  Printer.CurrentX = X
                  Printer.Print Left(Rsd!Descripcion, 25)
                  
                  Y = Y + 0.2
                  If Y >= 28 Then
                     Y = 0
                     Printer.NewPage
                 End If
                  Printer.CurrentY = Y
                  Printer.CurrentX = X
                  Printer.Print Format(Rsd!Cantidad, "#0.00") & " x " & Format(Rsd!PUnitario, "#0.00") & Space(30 - cantCanti) & Space(12 - cantTotdet) & Format(nTot, "$0.00")
                 Rsd.MoveNext
              Loop
                              
              Printer.FontBold = False
                              '
              Y = Y + 0.5
             
              If Y >= 28 Then
                    Y = 0
                    Printer.NewPage
              End If
                 
             Printer.CurrentY = Y
             Printer.CurrentX = X
             Printer.FontBold = True
                
             Printer.Print "Importe Total :" & Format(LblTotal.Caption, "#0.00")
             Printer.FontBold = False
             
             Y = Y + 3
             If Y >= 28 Then
                    Y = 0
                    Printer.NewPage
              End If

            Printer.CurrentY = Y
            Printer.CurrentX = X
            Printer.Print " "

            Y = Y + 1
            If Y >= 28 Then
                    Y = 0
                    Printer.NewPage
              End If

            Printer.CurrentY = Y
            Printer.CurrentX = X
            Printer.Print " "

             Y = Y + 1
             If Y >= 28 Then
                    Y = 0
                    Printer.NewPage
              End If
            Printer.CurrentY = Y
            Printer.CurrentX = X
            Printer.Print " "
     
             Printer.CurrentY = Y
             Printer.CurrentX = X
             Printer.Print "----------------------------------------"
             
                                        
             Printer.EndDoc
          End If
       End If
  End If
     
grabarTicket:
     Dim RsCv As ADODB.Recordset
     Dim RsDa As ADODB.Recordset
     Dim rReci As ADODB.Recordset
     Dim rCant As ADODB.Recordset
     Dim rValo As ADODB.Recordset
     Dim rCaja As ADODB.Recordset
     Dim rApli As ADODB.Recordset
     Dim rRecibo As ADODB.Recordset
    ' Grabar Cabecera

     Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
     Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
     Set RsDa = cRsl.RsVacio("DetallesComprobantes", "Id", "N")
     Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
     Set rValo = cRsl.RsVacio("Valores", "Id", "N")
     Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
     Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
     Set rRecibo = cRsl.RsVacio("CabComprobantes", "Id", "N")
     
           
     cRv.MoveFirst
           
     RsCv.AddNew
     RsCv!Suc = nSucursal
     RsCv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
     RsCv!Movimiento = nMovimiento
     RsCv!Comprobante = cRv!Comprobante
     RsCv!fecha = Date
     RsCv!FechaIva = Date
     RsCv!Numero = IIf(nNumero = 0, nNumCompro, nNumero)
     RsCv!sucursal = cRv!sucursal
     RsCv!Cliente = cRv!Cliente
     RsCv!Lista = cRv!Lista
     RsCv!Vendedor = cRv!Vendedor
     RsCv!Caja = cRv!Caja
     RsCv!CondVenta = CmbCond.ItemData(CmbCond.ListIndex)
           
     LblTotal.Caption = 0
     LblTotalGrabado.Caption = 0
     LblTotalIva.Caption = 0
     LblTotalNoGrabado.Caption = 0
     LblDescuento.Caption = 0

     nTotal = 0
     nGrabado = 0
     nIvas = 0
     nNoGrabado = 0


     Rsd.MoveFirst

     Do While Not Rsd.EOF
        RsDa.AddNew
        RsDa!Movimiento = RsCv!Movimiento
        RsDa!Producto = Rsd!Producto
        RsDa!Descripcion = Rsd!Descripcion
        RsDa!Cantidad = Rsd!Cantidad
        RsDa!Cuenta = 19
        RsDa!Descuento = 0
        RsDa!Tasa = cRsl.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & Rsd!Producto & "'")
        RsDa!Deposito = 1
        RsDa!Medida = cRsl.TraerValorDeUnCampo("Productos", "UMVenta", "Producto='" & Rsd!Producto & "'")
        If Frame2.Caption = "N" Then
           RsDa!Impuesto = 0
           RsDa!PrecioTotal = Rsd!Total
           RsDa!PrecioUnitario = Rsd!Total / Rsd!Cantidad
           nTotal = nTotal + RsDa!PrecioTotal
           nGrabado = nGrabado + RsDa!PrecioTotal
        Else
           sInter = cRsl.TraerValorDeUnCampo("Productos", "TipoInterno", "Producto='" & Rsd!Producto & "'")
           nInter = 0
           pInterPorc = 0: vInterno = 0
           Select Case sInter
                  Case "No"
                        nInter = 0
                        sInter = "0"
                   Case "Porcentaje"
                         sInter = ""
'                          sInter = "0"
                         pInterPorc = (cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & Rsd!Producto & "'") / 100) + 1
                         vInterno = Int(((Rsd!PUnitario) / pInterPorc) * 100)
                         nInter = vInterno / 100
                   Case "Fijo"
                         nInter = cRsl.TraerValorDeUnCampo("Productos", "ValorInterno", "Producto='" & Rsd!Producto & "'")
          End Select
          nTasa = cRsl.TraerTasaProducto(Rsd!Producto)
          RsDa!Impuesto = ((Rsd!Total - (Rsd!Cantidad * nInter)) - (Rsd!Total - (Rsd!Cantidad * nInter)) / ((nTasa / 100) + 1)) '/ Rsd!Cantidad
  
          nNoGrabado = nNoGrabado + (Rsd!Cantidad * nInter)
          If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = False Then
             RsDa!PrecioTotal = Rsd!Total
             RsDa!PrecioUnitario = RsDa!PrecioTotal / Rsd!Cantidad
           '  RsDa!PrecioUnitario = Round((Rsd!PUnitario - nInter) / ((nIva / 100) + 1), 3)
             nGrabado = nGrabado + Rsd!Total - (Rsd!Cantidad * nInter)
          Else
             nGrabado = nGrabado + (Rsd!Total - (Rsd!Cantidad * nInter)) / ((nTasa / 100) + 1)
             RsDa!PrecioTotal = (Rsd!Total - (Rsd!Cantidad * nInter)) / ((nTasa / 100) + 1)
             RsDa!PrecioUnitario = RsDa!PrecioTotal / Rsd!Cantidad
             nIvas = nIvas + (RsDa!Impuesto * Rsd!Cantidad)
          End If
          nTotal = nTotal + Rsd!Total
        End If
        
        RsDa!MoviRemi = 0
        RsDa.Update

       If cRsl.TraerValorDeUnCampo("Productos", "ActuaStock", "Producto='" & Rsd!Producto & "'") = "Si" Then
          rCant.AddNew
          rCant!Movimiento = RsCv!Movimiento
          rCant!Estado = 1
          rCant!Producto = Rsd!Producto
          rCant!Deposito = 1
          Select Case cNum.QueMoviStock(CmbCompro.ItemData(CmbCompro.ListIndex))
                   Case "Suma"
                        rCant!CantidadDebe = Rsd!Cantidad
                        rCant!CantidadHaber = 0
                   Case "Resta"
                        rCant!CantidadDebe = 0
                        rCant!CantidadHaber = Rsd!Cantidad
          End Select
          rCant.Update
        End If
        Rsd.MoveNext
     Loop

     LblTotal.Caption = Round(nTotal, 2)
     LblTotalGrabado.Caption = nGrabado
     LblTotalIva.Caption = nIvas
     LblTotalNoGrabado.Caption = nNoGrabado
     
     RsCv!Neto = CCur(LblTotalGrabado.Caption)
     RsCv!Iva1 = CCur(LblTotalIva.Caption)
     RsCv!Iva2 = 0
     RsCv!Descuentos = 0
     RsCv!Financiacion = 0
     RsCv!NoGravados = CCur(LblTotalNoGrabado.Caption)
     RsCv!PIvaCompras = 0
     RsCv!IBrutosCompras = 0
      Select Case cRsl.TraerValorDeUnCampo("Comprobantes", "TipoMovimiento", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
             Case 1, 11
                  RsCv!Debe = CCur(LblTotal.Caption)
                  RsCv!Haber = 0
             Case 2, 12, 13
                  RsCv!Debe = 0
                  RsCv!Haber = CCur(LblTotal.Caption)
      End Select
      RsCv!Desde = 2
      RsCv!Hasta = 0
      RsCv!factura = 0
      RsCv!Motivo = ""
      RsCv!COTIZACION = 1
      RsCv!Anulado = 0
      RsCv.Update

        ' Grabar Cuenta Corriente
     If CmbCond.ItemData(CmbCond.ListIndex) = 0 And TxtCliente.text <> sConsFinal Then
         rReci.AddNew
         rReci!fecha = Date
         rReci!Cliente = cRv!Cliente
         rReci!Venta = RsCv!Movimiento
         rReci!Comprobante = cRv!Comprobante
         rReci!Numero = cRv!Numero
         rReci!sucursal = cRv!sucursal
         rReci!Cuota = 1
         Select Case cRsl.TraerValorDeUnCampo("Comprobantes", "TipoMovimiento", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex))
                Case 1, 11
                     rReci!Debe = Round(nTotal, 2)
                     rReci!Haber = 0
                Case 2, 12, 13
                     rReci!Debe = 0
                     rReci!Haber = Round(nTotal, 2)
         End Select
         rReci!Anulado = 0
         rReci.Update
        Dim cRcL As ClsClienteL
        Set cRcL = New ClsClienteL
        
        sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & CmbCompro.ItemData(CmbCompro.ListIndex) & " AND Usuario=" & nUsuario)
             
        If sImpre <> "Ninguna" And sImpre <> "0" And bAnulaPres <> True Then
                                                
             For Each P In Printers
                 If P.DeviceName = sImpre Then
                     Set Printer = P
                     Printer.FontName = "Courier New"
                     Printer.ScaleMode = 7
                     Printer.FontSize = 9
                     Exit For
                  End If
              Next
                                         
              Printer.FontSize = 14
                                         
              Y = 1
              X = 0
                      
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.FontBold = True
              Printer.Print sNomFan
              Printer.FontBold = False
              
              Printer.FontSize = 7
              
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print Date & " - " & Time
              
               Y = Y + 0.8
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.FontBold = True
              Printer.Print "CUENTA CORRIENTE"
              Printer.FontBold = False
        
              Y = Y + 0.4
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print Trim(LblCliente.Caption)
              
              Y = Y + 0.4
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "IMPORTE :" & CCur(LblTotal.Caption)
                           
              Y = Y + 0.4
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "Saldo Anterior :" & Format(cRcL.TraerSaldoCtaCte(TxtCliente.text), "#0.00")
               
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "----------------------------------------"
               
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "     Firma Cliente"
                                    
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "----------------------------------------"
              
              Y = Y + 0.5
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print "   Aclaracion"
              
              Y = Y + 0.8
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print " "
               
              Y = Y + 0.8
              Printer.CurrentY = Y
              Printer.CurrentX = X
              Printer.Print " "
                                          
             Printer.EndDoc
             
         End If
         
     Else
'        Dim nNum, nSuc As Long
      
      
       ' Grabar Valores
       ' Actualizo Caja
        If bFormaNueva = False Then
           rValo.AddNew
           rValo!Venta = RsCv!Movimiento
    '      rValo!Venta = nMoviAfip
           rValo!CodPago = IIf(cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & CmbForma.ItemData(CmbForma.ListIndex)) = 1, 1, CmbForma.ItemData(CmbForma.ListIndex))
           rValo!formapago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", CmbForma.ItemData(CmbForma.ListIndex))
           rValo!Nombre = ""
           If cNum.SumaCaja(CmbCompro.ItemData(CmbCompro.ListIndex)) Then
              rValo!importe = LblTotal.Caption
           Else
              rValo!formapago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", CmbForma.ItemData(CmbForma.ListIndex)) & "-Nota de Credito"
              rValo!importe = -LblTotal.Caption
         End If
         rValo.Update
      
      
      
         rValo.MoveFirst
         rCaja.AddNew
         rCaja!Suc = nSucursal
         rCaja!Movimiento = RsCv!Movimiento
         rCaja!Comprobante = cRv!Comprobante
         rCaja!Numero = IIf(cRv!Numero = 0, nNumCompro, cRv!Numero)
         rCaja!sucursal = cRv!sucursal
         rCaja!fecha = Date
         rCaja!Hora = Time
         rCaja!Caja = cRsl.RegPorDefecto("Caja")
         rCaja!formapago = rValo!formapago
         rCaja!Nombre = LblCliente.Caption
         rCaja!Banco = rValo!Banco
         rCaja!NumeroCheque = rValo!NumeroCheque
         rCaja!FechaAcreditacion = rValo!FechaAcreditacion
         rCaja!NumeroTarjeta = rValo!NumeroTarjeta
         rCaja!FechaVencimiento = rValo!FechaVencimiento
         rCaja!Autorizacion = rValo!Autorizacion
         rCaja!Cupon = rValo!Cupon
         If cNum.SumaCaja(CmbCompro.ItemData(CmbCompro.ListIndex)) Then
            rCaja!Debe = rValo!importe
            rCaja!Haber = 0
         Else
            rCaja!Debe = 0
            rCaja!Haber = rValo!importe
         End If
         rCaja!Abierta = "S"
         rCaja!Anulado = 0
         rCaja.Update
      Else
          bFormaNueva = False
      
          Dim rValoAux As ADODB.Recordset
          Dim rBanco As ADODB.Recordset
                            
          If nFormaPago < 2 And LblTotal.Caption <> 0 Then
             If Not VentaValores.RsV Is Nothing Then
               If VentaValores.RsV.RecordCount <> 0 Then
                  
                  If nFormaPago = 0 Then
                   ' Grabar Valores
                     rValo.AddNew
                       ' rValo!Id = 0
                     rValo!Venta = nMoviAfip
                     rValo!CodPago = rValoAux!CodPago 'cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "Condta=" & nFormaPago)
                     rValo!formapago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
                     rValo!Nombre = ""
                     rValo!importe = LblTotal.Caption
                     rValo.Update
                     
                        
                  ' Actualizo Caja
                     rValo.MoveFirst
                     rCaja.AddNew
                     rCaja!Suc = nSucursal
                     rCaja!Movimiento = nMoviAfip
                     rCaja!Comprobante = nRecibo
                     rCaja!Numero = cRv!Numero
                     rCaja!sucursal = cRv!sucursal
                     rCaja!fecha = Date
                     rCaja!Hora = Time
                     rCaja!Caja = cRsl.RegPorDefecto("Caja")
                     rCaja!formapago = rValo!formapago
                     rCaja!Nombre = LblCliente.Caption
                     rCaja!Banco = rValo!Banco
                     rCaja!NumeroCheque = rValo!NumeroCheque
                     rCaja!FechaAcreditacion = rValo!FechaAcreditacion
                     rCaja!NumeroTarjeta = rValo!NumeroTarjeta
                     rCaja!FechaVencimiento = rValo!FechaVencimiento
                     rCaja!Autorizacion = rValo!Autorizacion
                     rCaja!Cupon = rValo!Cupon
                     If cNum.SumaCaja(CmbCompro.ItemData(CmbCompro.ListIndex)) Then
                        rCaja!Debe = rValo!importe
                        rCaja!Haber = 0
                     Else
                        rCaja!Debe = 0
                        rCaja!Haber = rValo!importe
                     End If
                     rCaja!Abierta = "S"
                     rCaja!Anulado = 0
                     rCaja.Update
                Else
                
                   Set rValoAux = VentaValores.RsV
                   
                  
                   rValoAux.MoveFirst
                   Do While Not rValoAux.EOF
                      rValo.AddNew
                     ' rValo!Id = 0
                      rValo!Venta = nMoviAfip
                      rValo!CodPago = rValoAux!CodPago
                      rValo!formapago = rValoAux!formapago
                      rValo!Nombre = rValoAux!Nombre
                      rValo!Banco = rValoAux!Banco
                      rValo!IdCheque = rValoAux!IdCheque
                      rValo!NumeroCheque = rValoAux!NCheque
                      rValo!FechaAcreditacion = rValoAux!FAcredi
                      rValo!NumeroTarjeta = rValoAux!NTarjet
                      rValo!FechaVencimiento = rValoAux!FVencim
                      rValo!Autorizacion = rValoAux!NAutori
                      rValo!Cupon = rValoAux!NCupon
                      rValo!importe = Format(rValoAux!importe, "#0.00")
                      rValo!PorVenta = " "
                      rValo.Update
                      rValoAux.MoveNext
                  Loop
                
                  If rValo.EOF = False Or rValo.BOF = False Then
                     rValo.MoveFirst
                     Do While Not rValo.EOF
                        If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(rValo!formapago) & "'") = "Caja" Then
                           rCaja.AddNew
                           rCaja!Suc = nSucursal
                           rCaja!Movimiento = nMoviAfip
                           rCaja!Comprobante = cRv!Comprobante
                           rCaja!Numero = cRv!Numero
                           rCaja!sucursal = cRv!sucursal
                           rCaja!fecha = Date
                           rCaja!Hora = Time
                           rCaja!Caja = cRsl.RegPorDefecto("Caja")
                           rCaja!formapago = rValo!formapago
                           rCaja!Nombre = LblCliente.Caption
                           rCaja!Banco = rValo!Banco
                           rCaja!NumeroCheque = rValo!NumeroCheque
                           rCaja!FechaAcreditacion = rValo!FechaAcreditacion
                           rCaja!NumeroTarjeta = rValo!NumeroTarjeta
                           rCaja!FechaVencimiento = rValo!FechaVencimiento
                           rCaja!Autorizacion = rValo!Autorizacion
                           rCaja!Cupon = rValo!Cupon
                           If cNum.SumaCaja(CmbCompro.ItemData(CmbCompro.ListIndex)) Then
                              rCaja!Debe = rValo!importe
                              rCaja!Haber = 0
                           Else
                              rCaja!Debe = 0
                              rCaja!Haber = rValo!importe
                           End If
                           rCaja!Abierta = "S"
                           rCaja!Anulado = 0
                           rCaja.Update
                        Else
                            If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(rValo!formapago) & "'") = "Banco" Then
                               rBanco.AddNew
                               rBanco!Movimiento = nMoviAfip
                               rBanco!Comprobante = cRv!Comprobante
                               rBanco!Numero = cRv!Numero
                               rBanco!sucursal = cRv!sucursal
                               rBanco!fecha = Date
                               rBanco!Hora = Time
                               rBanco!Caja = cRsl.RegPorDefecto("Caja")
                               rBanco!formapago = rValo!formapago
                               rBanco!Nombre = rValo!Nombre
                               rBanco!Banco = rValo!Banco
                               rBanco!NumeroCheque = rValo!NumeroCheque
                               rBanco!FechaAcreditacion = rValo!FechaAcreditacion
                               rBanco!NumeroTarjeta = rValo!NumeroTarjeta
                               rBanco!FechaVencimiento = rValo!FechaVencimiento
                               rBanco!Autorizacion = rValo!Autorizacion
                               rBanco!Cupon = rValo!Cupon
                               If cNum.SumaCaja(rBanco!Comprobante) Then
                                  rBanco!Debe = rValo!importe
                                  rBanco!Haber = 0
                               Else
                                  rBanco!Debe = 0
                                  rBanco!Haber = rValo!importe
                               End If
                               rBanco!Abierta = "S"
                               rBanco!Anulado = 0
                               rBanco.Update
                            End If
                        End If
                        rValo.MoveNext
                    Loop
                  End If
                   ' cGrab.AgregarCobranzaVenta RsCv, rReci, rValo, rApli, rCaja, rBanco, rRet
                End If
              End If
            End If
         End If
     End If
     
    End If
                                                                                                                                                                                                        
    cPe.AgregarConRecibo RsCv, RsDa, rReci, rValo, rApli, rCant, rCaja, rRecibo
    If sCae <> "" Then
       GrabarCae
    End If
    
    Dim sNom As String
    
    sNom = CurDir() & "\venta.dat"

    Kill sNom
             
    bGraboTick = True
      
    Set cRv = Nothing
    Set cRle = Nothing
   
    nId = 1
    
    Erase nUnitario
    Frame2.Visible = False
    LimpiarDetalles
    HabilitarDetalles bCod, False
    CmdDetalle(0).Enabled = bBoto
    Grid1.Enabled = True
    dTotalAux = 0
    LblTotal.Caption = "0,00"
    LblTotalGrabado.Caption = "0,00"
    LblTotalNoGrabado.Caption = "0,00"
    LblTotalIva.Caption = "0,00"
    Rsd.MoveFirst
    Do While Not Rsd.EOF
       Rsd.Delete
       Rsd.MoveNext
    Loop
    TxtPago.text = "0.00"
    LblVuelto.Caption = "0.00"
        
    TxtCliente.text = sConsFinal
    LblCliente.Caption = sNomFinal
    CmbCond.text = sCondPago
    CmbForma.text = cRsl.TraerValorDeUnCampo("CondVenta", "Descripcion", "CondVta=" & nFormaPago)
    CmbCompro.Enabled = False
    If Frame2.Caption = "J" Then
       Frame2.Caption = "N"
    End If
    If Frame2.Caption = "P" Then
       CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinal"))
    Else
        CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinalN"))
    End If
    
    TxtProducto.SetFocus

   Set cRsl = Nothing
Exit Sub
Errores:

  If err.Number <> 0 Then
     If err.Number = 481 Or err.Description = "Desbordamiento" Then
        Resume Next
     End If
     MsgBox err.Description, vbCritical, "Atención"
     cRsl.GrabarError "CmdVuelto", err.Description, "Numero:" & cRsl.TraerValorDeUnCampo("Comprobantes", "Numero", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) & " Movimiento=" & nMovimiento
  Else
     MsgBox "Controlador Apagado, fuera de Línea, Cable Desconectado o Falta Papel..." & _
     Chr(13) & Chr(10) & "Reinicie Computadora e Impresora y Vuelva a Intentar", vbCritical, "Atención"
  End If

   Frame2.Visible = False
   HabilitarDetalles bCod, False
   CmdDetalle(0).Enabled = bBoto
   LblTotal.Caption = Format(dTotalAux, "#0.00")
   dTotalAux = 0
   Grid1.Enabled = True

   Exit Sub
End Sub

 Private Function TraerSpacios(ByVal nLen As Double) As Byte
    nLen = Format(nLen, "#0.00")
    TraerSpacios Len(nLen)
  End Function
  
Private Sub Command1_Click()
    bFormaNueva = True
    nLlama = 10
    If nFormaPago = 1 And LblTotal.Caption <> 0 Then
        VentaValores.Show 1
    End If
End Sub

Private Sub Form_Activate()
   Nuevo
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo Errores
   Dim cRsl As ClsLectura
   
   Set cRsl = New ClsLectura
   If KeyCode = vbKeyF3 Then
      If Frame2.Caption = "J" Then
         Frame2.Caption = "N"
      Else
          Frame2.Caption = "J"
      End If
      CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinalN"))
      CmbCompro.Enabled = False
   End If
   If KeyCode = vbKeyF4 Then
      If Frame2.Caption = "N" Or Frame2.Caption = "J" Then
         Frame2.Caption = "P"
         CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinal"))
         CmbCompro.Enabled = False
      Else
         Frame2.Caption = "N"
         CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinalN"))
         CmbCompro.Enabled = False
      End If
      TxtCliente.text = sConsFinal
      LblCliente.Caption = sNomFinal
      If TxtPago.Visible = True Then
         TxtCliente.SetFocus
      End If
   End If
   If Frame2.Visible = False Then
      If KeyCode = vbKeyF1 Then
          TxtProducto.text = ""
          nDat = 0
          nBuscar = 3
          nLlama = 5
          FrmBuscarEx.Show 1
          If nDat <> 0 Then
             TxtProducto.text = cRsl.TraerValorDeUnCampo("Productos", "Producto", "Id=" & nDat)
             TxtProducto_LostFocus
          End If
      End If
      If KeyCode = vbKeyF5 Then
         If Rsd.RecordCount <> 0 Then
            TxtProducto.text = ""
            Frame2.Visible = True
            Recalcular cRsl.TraerValorDeUnCampo("CondVenta", "RgTarjeta", "CondVta=" & CmbForma.ItemData(CmbForma.ListIndex))
            dTotalAux = LblTotal.Caption
            CargarFrame
            TxtCliente.Enabled = True
            BtnBuscar.Enabled = False
            CmbCond.Enabled = False
            CmdVuelto(0).Enabled = True
            TxtCliente.SetFocus
            TxtCliente.SelStart = 0
            TxtCliente.SelLength = Len(TxtCliente.text)
         End If
      End If
   ' Borrar el Ultimo
   If KeyCode = 46 Then
      If pClave = True And bDejo = True Then
         FrmClave.Show 1
         If bClave = True Then
           If Rsd.RecordCount <> 0 Then
              Rsd.MoveLast
              CalcularTotales Rsd, False
              Rsd.Delete
              GrabarTxtRsd Rsd
              If LblTotal.Caption = 0 Then
                 nId = 1
              End If
           Else
             LblTotal.Caption = "0.00"
           End If
         End If
         bDejo = False
      Else
         If Rsd.RecordCount <> 0 Then
              Rsd.MoveLast
              CalcularTotales Rsd, False
              Rsd.Delete
              GrabarTxtRsd Rsd
              If LblTotal.Caption = 0 Then
                 nId = 1
              End If
          Else
             LblTotal.Caption = "0.00"
          End If
      End If
   End If
   ' Borrar por Id
   If KeyCode = vbKeyF7 Then
      If pClave = True And bDejo = True Then
         FrmClave.Show 1
         If bClave = True Then
           If Rsd.RecordCount <> 0 Then
              TxtId.Visible = True
              CmdDetalle(0).Default = False
              TxtId.SetFocus
           End If
         End If
         bDejo = False
      Else
         If Rsd.RecordCount <> 0 Then
            TxtId.Visible = True
            CmdDetalle(0).Default = False
            TxtId.SetFocus
         End If
      End If
   End If
   ' Cambia Precio
   If KeyCode = vbKeyF8 Then
     
      If bCajaDesc = True Then
         LblTotalGrabado.Caption = 0
         If Frame3.Visible = False Then
            If Rsd.RecordCount <> 0 Then
               bF11 = True
               TxtCPrecio.Visible = True
               CmdDetalle(0).Default = False
               TxtCPrecio.SetFocus
            End If
          End If
     End If
    End If

   ' Codigo
   If KeyCode = vbKeyF9 Then
      If TxtProducto.Enabled = True Then
          TxtProducto.SetFocus
          TxtProducto.SelStart = 0
          TxtProducto.SelLength = Len(TxtProducto.text)
      End If
   End If
   If KeyCode = vbKeyF11 Then
      If TxtCantidad.Enabled = True Then
          bF11 = True
          TxtCantidad.SetFocus
          TxtCantidad.SelStart = 0
          TxtCantidad.SelLength = Len(TxtCantidad.text)
      End If
   End If
   If KeyCode = vbKeyF12 And bAdmin = True Then
      If TxtPrecio.Enabled = True And bAdmin = True Then
          TxtPrecio.SetFocus
          TxtPrecio.SelStart = 0
          TxtPrecio.SelLength = Len(TxtPrecio.text)
      End If
   End If
  Else
volverPorError:
    If KeyCode = vbKeyF9 Then
       Frame2.Visible = False
       HabilitarDetalles bCod, False
       CmdDetalle(0).Enabled = bBoto
       LblTotal.Caption = Format(dTotalAux, "#0.00")
       Grid1.Enabled = True
       Recalcular 0
       Erase nUnitario
   End If
      If KeyCode = vbKeyF5 Then
         TxtCliente.SetFocus
      End If
      If KeyCode = vbKeyF6 Then
         CmbCompro.Enabled = True
         CmbCompro.SetFocus
      End If
      If KeyCode = vbKeyF7 Then
         TxtPago.SetFocus
      End If
      If KeyCode = vbKeyF11 Then
         CmbCond.SetFocus
      End If
      If KeyCode = vbKeyF12 Then
         CmbForma.SetFocus
      End If
  End If
  Exit Sub
Errores:
   MsgBox err.Description, vbCritical, "Atención"
   cRsl.GrabarError "Key Down", err.Description, str(KeyCode)

End Sub

Private Function CalcularDescuento(pTotal As Double) As Double
  Dim nDesc As Double
  nDesc = 0
  If bDescuento = True Then
     If TxtCliente.text <> sConsFinal Then
        If CmbForma.text = sFormaPago Then
           If CmbCond.text = "Contado" Then
              pTotal = pTotal * dPorcDescuento / 100
              nDesc = pTotal
           End If
        End If
     End If
  End If
  
  CalcularDescuento = nDesc
  
End Function

Private Sub CargarFrame()
  
  Me.Refresh
  bCod = TxtProducto.Enabled
  bBoto = CmdDetalle(0).Enabled

  HabilitarDetalles False, False
  CmdDetalle(0).Enabled = False
  Grid1.Enabled = False
  TxtPago.SetFocus
  
  LblCompra.Caption = LblTotal.Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{TAB}"
      Exit Sub
   End If
End Sub

Private Sub Limpiar()

  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  

  TxtCliente.text = sConsFinal
  LblCliente.Caption = sNomFinal
  
  LblCompra.Caption = "0,00"
  TxtPago.text = "0,00"
  LblVuelto.Caption = "0,00"
  TxtId.text = ""
  TxtCPrecio.text = ""
  dTotalAux = 0
  LblTotalGrabado.Caption = "0,00"
  LblTotalNoGrabado.Caption = "0,00"
  LblTotalIva.Caption = "0,00"
  LblDescuento.Caption = "0.00"
  TxtProducto.text = ""
  TxtDetalle.Caption = ""
  TxtCantidad.text = 1
  TxtPrecio.text = "0,000"
  LblTotal.Caption = "0,00"
  
  Set cRsl = Nothing
End Sub

Private Sub Form_Load()
   On Error GoTo Errores
   Dim cRsl As ClsLectura

   Set cRsl = New ClsLectura

   CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
   Limpiar
   bError = False
   sw = True
   bF11 = True
   CargarCombos
   CrearRsDetalles
   HabilitarDetalles False, False
   BotonDetalles True, False
   nId = 1
   Frame2.Visible = False
   Frame2.Caption = IIf(cRsl.RegPorDefecto("TipoVenta") = 0, "N", "P")
   
   bDejo = True
       
   bCajaDesc = cRsl.RegPorDefecto("DescCaja")
   
   If Frame2.Caption = "N" Then
      CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinalN"))
   Else
      CmbCompro.text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", cRsl.RegPorDefecto("VentaConsFinal"))
   End If
   
   bPercibe = cRsl.TraerValorDeUnCampo("Sistema", "Percibe", "")
      
   Set cQrCode = New ClsQrCode

   Me.Top = 0
   Me.Left = 0
   CmbCompro.Enabled = False
   Erase nUnitario
  
  Cadena = ""
 
  LblTotal.Caption = "0.00"
  
  LlenarTxtRsd
Exit Sub
Errores:
  MsgBox "Debe Definir un Comprobante 'N' para Este Usuario...", vbCritical, "Atención"
End Sub

Private Sub CargarCombos()
 Dim cRsl As ClsLectura
 Dim nCons As Long, nCondPago As Byte

 Set cRsl = New ClsLectura
  
 CmbCond.AddItem "Contado"
 CmbCond.ItemData(CmbCond.NewIndex) = 1
 CmbCond.AddItem "Cuenta Corriente"
 CmbCond.ItemData(CmbCond.NewIndex) = 0
 
 cRsl.CargaCombo CmbCompro, "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
 cRsl.CargaCombo CmbForma, "CondVenta", "CondVta", "Descripcion", "Tipo=1 or Tipo=2"
 

 nCons = cRsl.TraerValorDeUnCampo("Registros", "ClienteConsFinal", "Suc=" & nSucursal & " and Usuario=" & nUsuario)
 sConsFinal = cRsl.TraerValorDeUnCampo("Clientes", "Cliente", "Id=" & nCons)
 sNomFinal = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & sConsFinal & "'")
 TxtCliente.text = sConsFinal
 LblCliente.Caption = sNomFinal
 
 nFormaPago = cRsl.TraerValorDeUnCampo("Registros", "Cobro", "Suc=" & nSucursal & " and Usuario=" & nUsuario)
 nListaVenta = cRsl.RegPorDefecto("ListaRes")
 
 nCondPago = cRsl.RegPorDefecto("CPago")
 
 CmbCond.text = "Cuenta Corriente"
 
 If nCondPago = 1 Then
    CmbCond.text = "Contado"
 End If
 
 sCondPago = CmbCond.text
 
 CmbForma.text = cRsl.TraerValorDeUnCampo("CondVenta", "Descripcion", "CondVta=" & nFormaPago)
  
 Set cRsl = Nothing
End Sub

Private Sub BotonDetalles(bAgr As Boolean, bCan As Boolean)
  CmdDetalle(0).Enabled = bAgr
End Sub

Private Sub CrearRsDetalles()
  Set Rsd = New ADODB.Recordset
  Rsd.Fields.Append "Id", adInteger
  Rsd.Fields.Append "Producto", adVarChar, 25, adFldIsNullable
  Rsd.Fields.Append "Descripcion", adVarChar, 50, adFldIsNullable
  Rsd.Fields.Append "Cantidad", adDouble
  Rsd.Fields.Append "PUnitario", adDouble
  Rsd.Fields.Append "Total", adDouble
  Rsd.Fields.Append "Internos", adSingle
  Rsd.Fields.Append "Impuestos", adSingle
  Rsd.Open
  CabGrid
End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Id"
      .Columns(0).Width = 500
      .Columns(1).Caption = "Producto"
      .Columns(1).Width = 1000
      .Columns(2).Caption = "Descripción"
      .Columns(2).Width = 3300
      .Columns(3).Caption = "Cant."
      .Columns(3).Width = 800
      .Columns(3).Alignment = dbgRight
      .Columns(3).NumberFormat = "#0.000"
      .Columns(4).Caption = "P.Uni."
      .Columns(4).Width = 800
      .Columns(4).NumberFormat = "0.00"
      .Columns(4).Alignment = dbgRight
      .Columns(5).Caption = "Total"
      .Columns(5).Width = 1000
      .Columns(5).NumberFormat = "0.00"
      .Columns(5).Alignment = dbgRight
      .Columns(6).Visible = False
      .Columns(7).Visible = False
 End With
End Sub

Private Sub HabilitarDetalles(bHabi As Boolean, bId As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtPrecio.Enabled = bHabi
  TxtId.Visible = bId
  TxtCPrecio.Visible = bId
End Sub

Private Sub TxtCantidad_Change()
  TodoBien
End Sub

Private Sub TxtCantidad_GotFocus()
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.text)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
   If SoloNumeroDecimalFinal(KeyAscii, TxtCantidad) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCantidad_LostFocus()
  If Len(TxtCantidad.text) < 8 Then
     If TxtCantidad.text = "" Then TxtCantidad.text = 1
     If TxtCantidad.text <= 0 Then
        MsgBox "La Cantidad no puede Ser Menor o Igual a Cero", vbCritical, "Atención"
        TxtCantidad.text = 1
        TxtCantidad.SetFocus
     End If
     TxtCantidad.text = Format(TxtCantidad.text, "#0.000")
  Else
      TxtCantidad.text = "1.000"
      Beep
      MsgBox "Invalidad Cantidad de Articulos", vbCritical, "Atención"
      TxtProducto.SetFocus
  End If
End Sub

Private Sub TxtCliente_LostFocus()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Dim dTotal As Double
  
  dTotal = dTotalAux
  
  If TxtCliente.text = "" Then
     TxtCliente.text = "."
  End If

  If TxtPago.Visible = True And TxtPago.Enabled = True And CmdVuelto(0).Enabled = True Then
     If cRsl.TraerCantidad("Clientes", "Cliente", "Cliente='" & TxtCliente.text & "'") = 0 Then
        BtnBuscar.Enabled = True
        CmbCond.Enabled = True
        BtnBuscar.SetFocus
     Else
        LblCliente.Caption = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & TxtCliente.text & "'")
        CmbCompro.text = cRsl.TraerComproCliente(TxtCliente.text, Frame2.Caption)
        If TxtCliente.text = "." Then
           CmbCond.text = "Contado"
        Else
           CmbCond.text = sCondPago
           CmbCond.Enabled = True
        End If
        
        dTotal = dTotalAux - CalcularDescuento(dTotal)
        LblTotal.Caption = Format(dTotal, "#0.00")
        LblCompra.Caption = LblTotal.Caption
        If TxtPago.Visible = True Then
           TxtPago.SetFocus
        End If
     End If
  End If
End Sub

Private Sub TxtDetalle_Change()
  TodoBien
End Sub

Private Sub TxtId_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtId_LostFocus()
  If TxtId.text <> "" Then
     Rsd.MoveFirst
     Rsd.Find "Id=" & Val(TxtId.text)
     If Not Rsd.EOF Then
        CalcularTotales Rsd, False
        Rsd.Delete
        GrabarTxtRsd Rsd
     End If
     TxtId.text = ""
     TxtId.Visible = False
     LimpiarDetalles
     CmdDetalle(0).Default = True
     If CmdDetalle(0).Caption = "Agregar" Then
        Nuevo
     End If
  End If
End Sub

Private Sub TxtPago_GotFocus()
   If CmbCompro.Enabled = False Then
      TxtPago.SelStart = 0
      TxtPago.SelLength = Len(TxtPago.text)
   Else
       If CmbCompro.Enabled = True And CmbCompro.Visible = True Then
          CmbCompro.SetFocus
       End If
   End If
End Sub

Private Sub TxtPago_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtPago) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPago_LostFocus()
  On Error GoTo Errores
  If TxtPago.text = "" Then TxtPago.text = 0
  TxtPago.text = Format(TxtPago.text, "#0.00")
  LblVuelto.Caption = Format(CDbl(TxtPago.text) - CDbl(LblCompra.Caption), "#0.00")
Exit Sub
Errores:
 TxtPago.text = "0.00"
 TxtPago.SetFocus
 TxtPago.SelStart = 0
 TxtPago.SelLength = Len(TxtPago.text)
End Sub

Private Sub TxtPrecio_Change()
  TodoBien
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtPrecio) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPrecio_LostFocus()
  On Error Resume Next
  If TxtProducto.text <> "" And TxtDetalle.Caption <> "" Then
     TxtPrecio.text = Format(TxtPrecio.text, "0.000")
     Nuevo
     Nuevo
  Else
     TxtPrecio.text = Format(0, "0.000")
     TxtProducto.SetFocus
  End If
End Sub

Private Sub TxtProducto_Change()
  TodoBien
End Sub

Private Sub TxtProducto_LostFocus()
  Dim cRsl As ClsLectura, nDigi As Byte, nLen As Byte, cPrecAux As Single, cPrec As Single, Rs As Recordset
  Dim cSPrec As String, sNu As Byte, cLen As Byte, nCant As Double, sProd As String
  
  Set cRsl = New ClsLectura
  Set Rs = New Recordset
  
    If TxtProducto.text <> "" Then
     If nDigito <> 0 Then
        nDigi = nDigito
     End If
     sProd = Left(TxtProducto.text, nDigi)  ' & "00000"
     Set Rs = cRsl.TraerProductoPrecio(sProd, nListaVenta)
     If Rs.RecordCount <> 0 Then
        nPrecioEnvase = Rs!ValorEnvase
        
        If Rs!CodCompu = "Si" Then
           TxtDetalle.Caption = Rs!Descripcion
           nLen = fBalanza ' ojo con esto 2 0 0 Mucoop es 2
           nDigi = nDigi
'           If Rs!Producto = sEnvase Then
'               TxtPrecio.text = Mid(TxtProducto.text, nDigi + 1, Len(TxtProducto) - nDigito - 1)
'           Else
                TxtPrecio.text = Mid(TxtProducto.text, nDigi + 1, nLen)
              '  TxtPrecio.text = Mid(TxtProducto.text, nDigi - 2, 6) ' Len(TxtProducto) - nDigi - 2)
       '    End If
           If Rs!Producto = sEnvase Then
              TxtPrecio.text = Mid(TxtProducto.text, 8, 5) * 100
              TxtCantidad.text = -(TxtCantidad.text)
           End If
           If TxtPrecio.text <> "" Then
              If nScaner = 0 Then
                 cPrec = Format(TxtPrecio.text / 100, "#0.0000")
              Else
                 cPrec = Format(TxtPrecio.text, "#0.0000")
              End If
              
              cPrecAux = Format(Rs!precio, "#0.00")
              If cPrecAux <> 0 Then
                 cSPrec = cPrec
                 sNu = InStr(1, cSPrec, ".")
                 cLen = Len(Mid(cSPrec, sNu + 1, 3))
                 If cLen > 1 Then
                    cPrec = Left(cSPrec, Len(cSPrec))
                 Else
                    cPrec = Left(cSPrec, Len(cSPrec))
                 End If
                 TxtProducto.text = sProd
                 
                 If Rs!fijo = "No" Then
                    nCant = Format(cPrec / cPrecAux, "#0.000")
                    TxtCantidad.text = nCant
                    TxtPrecio.text = Format(cPrecAux, "#0.000")
                 Else
                  If Rs!fijo = "Um" Then
                       Dim sCantiUm As Double
                       sCantiUm = TxtPrecio.text / 1000
                       TxtPrecio.text = Rs!precio
                       TxtCantidad.text = sCantiUm
                       TxtPrecio.text = Format(TxtPrecio.text, "#0.00")
                       TxtProducto.text = sProd
                       Nuevo
                       Nuevo
                    Else
                        If Rs!fijo = "Ux" Then
                            TxtCantidad.text = 1
                            TxtPrecio.text = Format(cPrecAux, "#0.00")
                        Else
                            TxtCantidad.text = 1
                            TxtPrecio.text = Format(cPrec, "#0.00")
                        End If
                    End If
                 End If
                 Nuevo
                 Nuevo
              Else
                 cSPrec = cPrec
                 sNu = InStr(1, cSPrec, ".")
                 cLen = Len(Mid(cSPrec, sNu + 1, 3))
                 If cLen > 1 Then
                   cPrec = Left(cSPrec, Len(cSPrec))
                 Else
                   cPrec = Left(cSPrec, Len(cSPrec))
                 End If
                 If Rs!fijo = "Si" Then
                    TxtCantidad.text = 1
                    TxtPrecio.text = Format(cPrec, "#0.00")
                    TxtProducto.text = sProd
                    Nuevo
                    Nuevo
                 Else
                    MsgBox "Producto sin Precio", vbCritical, "Atención"
                 End If
              End If
           Else
               LimpiarDetalles
               Beep
               ErrorCodigo.Show 1
               TxtProducto.SetFocus
           End If
        Else
           ' Trae Producto y Precio
            Set Rs = cRsl.TraerProductoPrecio(sProd, nListaVenta)
            TxtDetalle.Caption = Rs!Descripcion
            TxtPrecio.text = Format(Rs!precio, "0.000")
            nPrecioEnvase = Rs!ValorEnvase
           bPrec = Rs!Calcular
           mPrec = CSng(TxtPrecio.text)
           If TxtDetalle.Caption <> "0" Then
              If TxtPrecio.text <> 0 Then
                 If bPrec = "No" Then
                    Nuevo
                    Nuevo
                 Else
                    TxtPrecio.SelStart = 0
                    TxtPrecio.SelLength = Len(TxtPrecio.text)
                    TxtPrecio.SetFocus
                 End If
              Else
                 TxtPrecio.SelStart = 0
                 TxtPrecio.SelLength = Len(TxtPrecio.text)
                 TxtPrecio.SetFocus
              End If
           Else
              LimpiarDetalles
              Beep
              ErrorCodigo.Show 1
              TxtProducto.SetFocus
           End If
        End If
     Else
        Dim nCad As Byte, nTam As Byte
        nCad = InStrRev(TxtProducto.text, "*", Len(TxtProducto.text))
        If nCad <> 0 Then
           TxtCantidad.text = 1
           nTam = Len(TxtProducto.text) - nCad
           TxtCantidad.text = 1
           If nCad > 1 Then
              TxtCantidad.text = Left(TxtProducto.text, nCad - 1)
           End If
           TxtProducto.text = Mid(TxtProducto.text, nCad + 1, nTam)
        End If
        
        ' Trae Producto y Precio
        Set Rs = cRsl.TraerProductoPrecio(TxtProducto.text, nListaVenta)
        
        If Rs.RecordCount <> 0 Then
           TxtProducto.text = Rs!Producto
           TxtDetalle.Caption = Rs!Descripcion
           TxtPrecio.text = Format(Rs!precio, "0.000")
           nPrecioEnvase = Rs!ValorEnvase
           bPrec = Rs!Calcular
           mPrec = CSng(TxtPrecio.text)
           If TxtDetalle.Caption <> "0" Then
              If TxtPrecio.text <> 0 Then
                 If bPrec = "No" Then
                    Nuevo
                    Nuevo
                 Else
                    TxtPrecio.SelStart = 0
                    TxtPrecio.SelLength = Len(TxtPrecio.text)
                    TxtPrecio.SetFocus
                End If
           Else
              TxtPrecio.SelStart = 0
              TxtPrecio.SelLength = Len(TxtPrecio.text)
              TxtPrecio.SetFocus
           End If
        Else
          LimpiarDetalles
          Beep
         ErrorCodigo.Show 1
         TxtProducto.SetFocus
        End If
       Else
          LimpiarDetalles
          Beep
          ErrorCodigo.Show 1
          If TxtProducto.Enabled = True Then
             TxtProducto.SetFocus
          End If
       End If
     End If
  Else
     If bF11 = False Then
        If TxtProducto.Enabled = True Then
           TxtProducto.SetFocus
           bF11 = True
        End If
     Else
       bF11 = False
     End If
  End If
  Set cRsl = Nothing
End Sub

Private Sub TodoBien()
  CmdDetalle(0).Enabled = False
  If TxtPrecio.text <> "" And TxtPrecio.text <> "0" And Len(TxtDetalle.Caption) <> 0 Then
     CmdDetalle(0).Enabled = True
  End If
End Sub

Public Sub CalcularTotales(pRsd As Recordset, pSuma As Boolean)
  Dim Rs As Recordset, cRsl As ClsLectura
  Dim sInter As String, nInter As Double, nTotalGrabado As Double, nTotalNoGrabado As Double, nTotalIva As Double
  Dim sInter1 As String, nTasa As Single

  If pSuma = True Then
     LblTotal.Caption = LblTotal.Caption + Rsd!Total
     LblTotal.Caption = Format(LblTotal.Caption, "#0.00")
  Else
     LblTotal.Caption = LblTotal.Caption - Rsd!Total
     LblTotal.Caption = Format(LblTotal.Caption, "#0.00")
  End If
   If pRsd!Producto <> "ENV" Then
        
      Set cRsl = New ClsLectura
      Set Rs = cRsl.TraerProductoPrecio(pRsd!Producto, nListaVenta)
      sInter = Rs!TipoInterno
    
    
      nTasa = Rs!Iva
      Select Case sInter
             Case "No"
                  nInter = 0
                  sInter = "0"
             Case "Porcentaje"
                   sInter = "0"
                   nInter = 0 'pRsd!PUnitario / (pRsd!PUnitario + (pRsd!PUnitario * rs!VarlorInterno & "'") / 100)
             Case "Fijo"
                   nInter = Rs!Internos * pRsd!Cantidad
      End Select
      nTotalGrabado = nTotalGrabado + ((pRsd!Total - nInter) / ((nTasa / 100) + 1))
      nTotalNoGrabado = nTotalNoGrabado + nInter
      nTotalIva = nTotalIva + (nTotalGrabado * nTasa / 100)
       
      If pSuma = True Then
         LblTotalGrabado.Caption = LblTotalGrabado.Caption + nTotalGrabado
         LblTotalNoGrabado.Caption = LblTotalNoGrabado.Caption + nTotalNoGrabado
         LblTotalIva.Caption = LblTotalIva.Caption + nTotalIva
      Else
         LblTotalGrabado.Caption = LblTotalGrabado.Caption - nTotalGrabado
         LblTotalNoGrabado.Caption = LblTotalNoGrabado.Caption - nTotalNoGrabado
         LblTotalIva.Caption = LblTotalIva.Caption - nTotalIva
      End If
   Else
      nTasa = 21
      nInter = 0
      sInter = "0"
    
      nTotalGrabado = nTotalGrabado + ((pRsd!Total - nInter) / ((nTasa / 100) + 1))
      nTotalNoGrabado = nTotalNoGrabado + nInter
      nTotalIva = nTotalIva + (nTotalGrabado * nTasa / 100)
       
      If pSuma = True Then
         LblTotalGrabado.Caption = LblTotalGrabado.Caption + nTotalGrabado
         LblTotalNoGrabado.Caption = LblTotalNoGrabado.Caption + nTotalNoGrabado
         LblTotalIva.Caption = LblTotalIva.Caption + nTotalIva
      Else
         LblTotalGrabado.Caption = LblTotalGrabado.Caption - nTotalGrabado
         LblTotalNoGrabado.Caption = LblTotalNoGrabado.Caption - nTotalNoGrabado
         LblTotalIva.Caption = LblTotalIva.Caption - nTotalIva
      End If
   End If
  
  Set cRsl = Nothing
  Set Rs = Nothing
End Sub

Public Function Autocompletar_Combo(Combo As ComboBox)
  Dim i As Integer, posSelect As Integer

   Select Case (KeyRetroceso Or Len(Combo.text) = 0)
          Case True
                KeyRetroceso = False
                Exit Function
   End Select

   With Combo
        For i = 0 To .ListCount - 1
            If InStr(1, .List(i), .text, vbTextCompare) = 1 Then
               posSelect = .SelStart
               .text = .List(i)
               .SelStart = posSelect
               .SelLength = Len(.text) - posSelect
               Exit For
          End If
        Next i
   End With
End Function

Private Sub TxtCPrecio_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtCPrecio) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCPrecio_LostFocus()
  Dim nTotG As Double
  
  If TxtCPrecio.text <> "" Then
     Rsd.MoveLast
     LblTotal.Caption = Format(LblTotal.Caption - Rsd!Total, "#0.00")
     nTotG = LblTotalGrabado.Caption
     Rsd!PUnitario = Rsd!PUnitario - ((Rsd!PUnitario * TxtCPrecio.text) / 100)
   '  Rsd!PUnitario = Round(CDbl(TxtCPrecio.text), 2)
     Rsd!Total = Round(CDbl(Rsd!PUnitario) * Rsd!Cantidad, 2)
     Rsd.Update
     CalcularTotales Rsd, True
  
     TxtCPrecio.text = ""
     TxtCPrecio.Visible = False
  Else
     TxtCPrecio.Visible = False
  End If
End Sub

Function ObtenerDecimales(ByVal n As Double) As String
    Dim pos As Long
    n = Round(n, 3)
    pos = InStr(n, ".") + InStr(n, ",")
    If pos <> 0 Then ObtenerDecimales = Mid$(n, pos + 1)
End Function

Private Function VerificarCuenta(pCliente As String) As Boolean

 If UCase(CmbCond.text) <> "CONTADO" Then

   Dim nSaldo As Currency, nLimite As Currency, cRcL As ClsClienteL, cRsl As ClsLectura
   Dim dif As Integer, UltPago As Date, sSaldo As Currency
   Dim nApro As Boolean, nDiferencia As Currency
        
   nApro = True
  
   Set cRcL = New ClsClienteL
   Set cRsl = New ClsLectura
  
  If Not cRsl.TraerValorDeUnCampo("Clientes", "FechaBaja", "Cliente='" & pCliente & "'") = "" Then
     MsgBox "Cliente Con Cuenta Inactiva", vbCritical, "Atención"
     nApro = False
     VerificarCuenta = False
     Exit Function
  End If
  
   UltPago = cRcL.TraerUltimoPago(pCliente)
   dif = DateDiff("d", UltPago, Date)
   If dif > nDias Then
      nApro = False
   End If

   sSaldo = cRcL.ResumenSaldoAnteriorBus(pCliente)
   nLimite = cRsl.TraerValorDeUnCampo("Clientes", "Limite", "Cliente='" & pCliente & "'")
                    
   If sSaldo > 1 And nLimite <> 0 Then
      If sSaldo + CCur(LblTotal.Caption) > nLimite Then
         nDiferencia = sSaldo + CCur(LblTotal.Caption) - nLimite
         nApro = False
      End If
   Else
      If CCur(LblTotal.Caption) > nLimite And nLimite <> 0 Then
         nApro = False
         dif = 0
      Else
         nApro = True
      End If
   End If
  
   If nApro = False Then
     MsgBox "Limite de Fecha de Pago Exedido :" & dif & Chr(10) & Chr(13) & "Limite de Credito Exedido :" & Format(nDiferencia, "#0.00"), vbCritical, "Atención"
     FrmClave.Show 1
     If bClave = True Then
        VerificarCuenta = True
     Else
        VerificarCuenta = False
     End If
   Else
      VerificarCuenta = True
   End If
  Else
    VerificarCuenta = True
  End If
End Function

Private Sub Recalcular(pValor As Double)
  Dim i As Integer
  LblTotal.Caption = 0
  
  i = 0
  
  If Not Rsd Is Nothing Then
    If Rsd.RecordCount <> 0 Then
      If EstaArrayVacio(nUnitario) Then
          ReDim Preserve nUnitario(Rsd.RecordCount)
          Rsd.MoveFirst
          Do While Not Rsd.EOF
             nUnitario(i) = Rsd!PUnitario
             i = i + 1
             Rsd.MoveNext
          Loop
          
       End If
       Rsd.MoveFirst
       i = 0
       Do While Not Rsd.EOF
          Rsd!PUnitario = nUnitario(i) + (nUnitario(i) * pValor / 100)
          Rsd!Total = Rsd!Cantidad * Rsd!PUnitario
          LblDescuento.Caption = CCur(LblDescuento.Caption) + (nUnitario(i) * Abs(pValor) / 100)
          Rsd.Update
          LblTotal.Caption = LblTotal.Caption + Rsd!Total
          i = i + 1
          Rsd.MoveNext
       Loop
    End If
  End If
  
  LblTotal.Caption = Format(LblTotal.Caption, "#0.00")
  dTotalAux = LblTotal.Caption

End Sub

Function EstaArrayVacio(vArray As Variant) As Boolean
    On Error Resume Next
    EstaArrayVacio = UBound(vArray)
    EstaArrayVacio = err ' Error 9 (Subscript out of range)
End Function

Private Sub GrabarCae()
   Dim cRle As ClsEscritura
   Set cRle = New ClsEscritura
        
   Dim rCae As ADODB.Recordset
   Set rCae = New ADODB.Recordset
        
   rCae.Fields.Append "Movimiento", adInteger
   rCae.Fields.Append "Fecha", adVarChar, 10
   rCae.Fields.Append "CAE", adVarChar, 50
   rCae.Fields.Append "Documento", adVarChar, 20
   rCae.Fields.Append "Comprobante", adInteger
   rCae.Fields.Append "Numero", adInteger
   rCae.Open
                                            
   rCae.AddNew
   rCae!Movimiento = nMoviAfip
   rCae!fecha = sVencimiento
   rCae!CAE = sCae
   rCae!Documento = sNumDoc
   rCae!Comprobante = CmbCompro.ItemData(CmbCompro.ListIndex)
   rCae!Numero = nNumCompro
   rCae.Update
   
   cRle.Actualizar "AfipCaes_A", rCae
   
   
End Sub

Private Sub ErrorImpre()
  Frame2.Visible = False
  HabilitarDetalles bCod, False
  CmdDetalle(0).Enabled = bBoto
  LblTotal.Caption = Format(dTotalAux, "#0.00")
  Grid1.Enabled = True
  Recalcular 0
  Erase nUnitario
End Sub


Private Function Autorizar(pCompro As Byte, pConcepto As Byte, pTipoDoc As Byte, pNumDoc As String, pTotal As String, pNeto As String, _
   pIva As String, pPunto As Integer, pMovi As Long, pCosto As Byte, rSx As Recordset, Optional pInterno As String, Optional pAlInterno As String) As Boolean
   
    Dim cRsl As ClsLectura, cRle As ClsEscritura
    
    Set cRle = New ClsEscritura
    Set cRsl = New ClsLectura

    Dim WSAA As Object, WSFEv1 As Object
    Dim ttl As Long, tra As String, Path As String, certificado As String, ClavePrivada As String
    Dim cms As String, cacert As String, wsdl As String
    Dim cache As String, wrapper As String, proxy As String, ok As Boolean, ta As String
    Dim tipo_cbte As Byte, punto_vta As Integer, cbte_nro As String, v As Variant, fecha As String
    Dim concepto As Byte, tipo_doc As Byte, nro_doc As String, cbt_desde As Long, cbt_hasta As Long
    Dim imp_total As String, imp_tot_conc As String, imp_neto As String, imp_iva As String, imp_trib As String
    Dim imp_op_ex As String, fecha_cbte As String, fecha_venc_pago As String
    Dim fecha_serv_desde As String, fecha_serv_hasta As String, moneda_id As String, moneda_ctz As String
    Dim evento As Variant, token As String, Sign As String
    Dim ID As Byte, base_imp As String, importe As String, CAE As String, cae2 As String
    Dim fd As String, Excepcion As String
    
    On Error GoTo ManejoError
    
    
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    If WSAA.Version < "2.02c" Then
       MsgBox "Debe instalar una versión más actualizada de PyAfipWs WSAA!"
       End
    End If
            
   ta = cRsl.TraerValorDeUnCampo("Sistema", "Ta", "")
   If ta <> "" Then
      ok = WSAA.AnalizarXml(ta)
       If Not WSAA.Expirado() Then
           token = WSAA.ObtenerTagXml("token")
           Sign = WSAA.ObtenerTagXml("sign")
       End If
   End If
            
   If token = "" Or Sign = "" Then
      
          ' deshabilito errores no manejados (version 2.04 o superior)
       WSAA.LanzarExcepciones = False
              
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
      ttl = 43200 ' tiempo de vida = 12hs hasta expiración
      tra = WSAA.CreateTRA("wsfe", ttl)
      ControlarExcepcion WSAA
      Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    
      Path = CurDir() + "\"
 
      certificado = sCertificado '"dagsis.crt" ' certificado de prueba
      ClavePrivada = sClavePrivada '"dagsis.key" ' clave privada de prueba
        
    ' Generar el mensaje firmado (CMS)
      cms = WSAA.SignTRA(tra, Path + certificado, Path + ClavePrivada)
      ControlarExcepcion WSAA
      Debug.Print cms
    
    ' Conectarse con el webservice de autenticación:
      cache = ""
      proxy = "" '"usuario:clave@localhost:8000"
      wrapper = ""
      cacert = ""
          
      wsdl = sWebServiceLogin
   
      ok = WSAA.Conectar(cache, wsdl, proxy, wrapper, cacert) ' Homologación
      ControlarExcepcion WSAA
    
    ' Llamar al web service para autenticar:
      ta = WSAA.LoginCMS(cms)
      ControlarExcepcion WSAA

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
      Debug.Print ta
      Debug.Print "Token:", WSAA.token
      Debug.Print "Sign:", WSAA.Sign
    
      If ta <> "" Then
           cRle.GrabarValordeUnCampo "Sistema", "Ta", "'" & ta & "'", ""
      End If
      token = WSAA.token
      Sign = WSAA.Sign

    End If
    
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    If WSAA.Version < "1.12" Then
        MsgBox "Debe instalar una versión mas actualizada de PyAfipWs WSFEv1!"
        End
    End If
    'Debug.Print WSFEv1.InstallDir

    ' Setear tocken y sing de autorización (pasos previos)
    WSFEv1.token = token ' WSAA.Token
    WSFEv1.Sign = Sign ' WSAA.Sign

    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = sCuitAfip
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False

    ' Conectar al Servicio Web de Facturación
    proxy = "" ' "usuario:clave@localhost:8000"
    proxy = ""

    wsdl = Trim(sWebServiceCae)

   ' wsdl = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL"

    cache = "" 'Path
    wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    cacert = "" ' WSAA.InstallDir & "\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)"

    ok = WSFEv1.Conectar(cache, wsdl, proxy, wrapper, cacert) ' homologación
    Debug.Print WSFEv1.Version
    ControlarExcepcion WSFEv1

    ' mostrar bitácora de depuración:
    Debug.Print WSFEv1.DebugLog

    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
   ' WSFEv1.Dummy
   ' ControlarExcepcion WSFEv1
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

  
    ' Establezco los valores de la factura a autorizar:
    tipo_cbte = pCompro '6
    punto_vta = pPunto
    
    If bError = True Then
       Exit Function
    End If
    
    cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
    ControlarExcepcion WSFEv1
    
    For Each v In WSFEv1.Errores
        Debug.Print v
    Next
    Debug.Print WSFEv1.errmsg
    Debug.Print WSFEv1.errcode
    
    If cbte_nro = "" Then
        cbte_nro = 0                ' no hay comprobantes emitidos
    Else
        cbte_nro = CLng(cbte_nro)   ' convertir a entero largo
    End If
    Dim sDate As Date
    sDate = Date
    
    fecha = Format(sDate, "yyyymmdd")
    
    concepto = pConcepto
    tipo_doc = pTipoDoc: nro_doc = pNumDoc
    
    cbte_nro = cbte_nro + 1
    cbt_desde = cbte_nro: cbt_hasta = cbte_nro
    
    Dim imp_interno As Single
 
  
    imp_interno = Format(pInterno, "0.00")
   
    imp_total = pTotal: imp_tot_conc = "0.00": imp_neto = Format(pNeto, "0.00")  '16.22
    imp_iva = Format(pIva, "0.00"): imp_trib = Format(imp_interno, "0.00"): imp_op_ex = "0.00" ' Aca va la persepcion imp_trib
    
    imp_iva_fact = imp_iva
    imp_tributo_fact = imp_trib
    
    fecha_cbte = fecha: fecha_venc_pago = ""
    ' Fechas del período del servicio facturado (solo si concepto = 1?)
    If pConcepto = 1 Then
       fecha_serv_desde = "": fecha_serv_hasta = ""
    Else
       fecha_serv_desde = fecha: fecha_serv_hasta = fecha
       fecha_venc_pago = fecha
    End If
    moneda_id = "PES": moneda_ctz = "1.000"
  '  imp_tot_conc = 0 ' TxtNoGrav.text
  
    Dim cRcli As ClsClienteL
    Set cRcli = New ClsClienteL
  
    Dim pTipoReceptor As Integer
    pTipoReceptor = cRcli.BuscarTipoIvaReceptor(TxtCliente.text)

    ok = WSFEv1.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz)

'     ok = WSFEv1.AgregarOpcional(2101, "2850590940090418135201")  ' CBU
'     ok = WSFEv1.AgregarOpcional(2102, "pyafipws")                ' alias
'     if tipo_cbte in (203, 208, 213):
'    WSFEv1.AgregarOpcional(22, "S")
      
    ' Agrego los comprobantes asociados:
    If tipo_cbte = 3 Or tipo_cbte = 2 Or tipo_cbte = 7 Or tipo_cbte = 8 Or tipo_cbte = 12 Or tipo_cbte = 13 Or tipo_cbte = 212 Or tipo_cbte = 213 Then   ' solo nc/nd
        Dim nNumeroAsoc As Integer, rAsoc As Recordset, tipoAsoc, puntoAsoc As Integer, numeroAsoc As Long
        nLlama = 9
        ComproBuscar.Show 1
        If ComproBuscar.nId <> 0 Then
           Set rAsoc = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
           
           tipoAsoc = cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & rAsoc!Comprobante)
           puntoAsoc = rAsoc!sucursal
           numeroAsoc = rAsoc!Numero
          ok = WSFEv1.AgregarCmpAsoc(tipoAsoc, puntoAsoc, numeroAsoc, sCuitAfip, Format(rAsoc!fecha, "yyyymmdd"))
       End If
    End If
'
'    ' Agrego impuestos varios
'    id = 99
'    Desc = "Impuesto Municipal Matanza'"
'    base_imp = "100.00"
'    alic = "0.10"
'    importe = "0.10"
'    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
'
'    ' Agrego impuestos varios

   Dim desc As String, nPerc As Currency, alic
   Dim nTasa(3) As Byte
   Dim sBaseCam(3) As Currency, sImporte(3) As Currency

'
  
'
'    ' Agrego impuestos varios
'    id = 1
'    Desc = "Impuesto nacional"
'    base_imp = "50.00"
'    alic = "1.00"
'    importe = "0.50"
'    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
'
     Dim i As Byte
'
'     Dim nTasa(3) As Byte
'     Dim sBaseCam(3) As Currency, sImporte(3) As Currency
     
     
     rSx.MoveFirst
    
    
      'Agrego impuestos varios
'    If LblPercepcion.Caption <> 0 Then
'       Id = 2
'       nPerc = cRsl.TraerValorDeUnCampo("Clientes", "Percepcion", "Cliente='" & TxtCliente.text & "'")
'       ok = WSFEv1.AgregarTributo(Id, "Perc.Ing.Bruto Bs.As.", LblNeto.Caption, Str(nPerc), LblPercepcion.Caption)
'    End If

'   ok = WSFEv1.AgregarIva(5, imp_neto, imp_iva)
    
    If bMono = False Then
        Dim sAuxUnitario As Double

        Do While Not rSx.EOF
           sAuxUnitario = rSx!PUnitario
           If cRsl.TraerValorDeUnCampo("Productos", "CodCompu", "Producto='" & rSx!Producto & "'") = "Si" Then
             If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & CmbCompro.ItemData(CmbCompro.ListIndex)) = "1" Then
                rSx!PUnitario = rSx!pTotal - rSx!Impuestos
             Else
                rSx!PUnitario = rSx!pTotal
             End If
           End If
           Select Case rSx!Tasa

               Case 9
                  nTasa(0) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & rSx!Tasa)
                  sBaseCam(0) = sBaseCam(0) + IIf(tipo_cbte = 1 Or tipo_cbte = 3, (rSx!pTotal - rSx!Impuestos), rSx!pTotal) - rSx!NoGrabado
                  sImporte(0) = sImporte(0) + rSx!Impuestos
               Case 10
                  nTasa(1) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & rSx!Tasa)
                  sBaseCam(1) = sBaseCam(1) + IIf(tipo_cbte = 1 Or tipo_cbte = 3, (rSx!pTotal - rSx!Impuestos), rSx!pTotal)
                  sImporte(1) = sImporte(1) + rSx!Impuestos
               Case 15
                  nTasa(2) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & rSx!Tasa)
                  sBaseCam(2) = sBaseCam(2) + IIf(tipo_cbte = 1 Or tipo_cbte = 3, (rSx!pTotal - rSx!Impuestos), rSx!pTotal) - rSx!NoGrabado '  * rSx!Cantidad)    '- Rsd!Descuento
                  sImporte(2) = sImporte(2) + rSx!Impuestos
               Case 16
                  nTasa(3) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & rSx!Tasa)
                  sBaseCam(3) = sBaseCam(3) + IIf(tipo_cbte = 1 Or tipo_cbte = 3, (rSx!pTotal - rSx!Impuestos), rSx!pTotal)
                  sImporte(3) = sImporte(3) + rSx!Impuestos
           End Select
           rSx!PUnitario = sAuxUnitario
           rSx.MoveNext
        Loop
        For i = 0 To 3
            If sImporte(i) <> 0 Or nTasa(i) = 3 Then
               ID = nTasa(i)
               base_imp = Round(sBaseCam(i), 2) - IIf(pCosto = 0, sImporte(i), 0)
               base_imp = Round(base_imp, 2)
               importe = Format(sImporte(i), "#0.00")
               ok = WSFEv1.AgregarIva(ID, base_imp, importe)
              ' ok = WSFEv1.AgregarIva(Id, 80, 18.9)
            End If
        Next
     End If

    If imp_interno <> 0 Then
       ID = 1
       desc = "Impuestos internos"
       base_imp = imp_neto
       alic = "0.00"
       importe = imp_interno
       ok = WSFEv1.AgregarTributo(ID, desc, base_imp, alic, importe)
    End If
    
    ' Habilito reprocesamiento automático (predeterminado):
    WSFEv1.Reprocesar = True
    
    ok = WSFEv1.EstablecerCampoFactura("cancela_misma_moneda_ext", "N")
    ok = WSFEv1.EstablecerCampoFactura("condicion_iva_receptor_id", pTipoReceptor)

    ' Solicito CAE:
    CAE = WSFEv1.CAESolicitar()
    ControlarExcepcion WSFEv1

    MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & CAE & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly

    ' Muestro los errores
    If WSFEv1.errmsg <> "" Then
        MsgBox WSFEv1.errmsg, vbExclamation, "Error"
    End If


    If CAE = "" Then
        ' hubo error, no comparo
    Else
        
      sCae = CAE
      sVencimiento = WSFEv1.Vencimiento
      sNumDoc = pNumDoc
      nNumCompro = cbte_nro
      
      Autorizar = True
    End If
Exit Function
ManejoError:
    ' Si hubo error (tradicional, no controlado):
    
    ' Depuración (grabar a un archivo los detalles del error)
    fd = FreeFile
    Open "c:\error.txt" For Append As fd
    If Not WSAA Is Nothing Then
        If WSAA.Version >= "1.02a" Then
            Print #fd, WSAA.Excepcion
            Print #fd, WSAA.Traceback
            Print #fd, WSAA.XmlRequest
            Print #fd, WSAA.XmlResponse
            ' guardo mensaje de error para mostrarlo:
            Excepcion = WSAA.Excepcion
        End If
    End If
    If Not WSFEv1 Is Nothing Then
        If WSFEv1.Version >= "1.10a" Then
            Print #fd, WSFEv1.Excepcion
            Print #fd, WSFEv1.Traceback
            Print #fd, WSFEv1.XmlRequest
            Print #fd, WSFEv1.XmlResponse
            Print #fd, WSFEv1.DebugLog()
            ' guardo mensaje de error para mostrarlo:
            Excepcion = WSFEv1.Excepcion
        End If
    End If
    Close fd
    
    Debug.Print err.Description            ' descripción error afip
    Debug.Print err.Number - vbObjectError ' codigo error afip
    If Excepcion = "" Then                 ' si no tengo mensaje de excepcion
        Excepcion = err.Description        ' uso el error de VB
    End If
    
    ' Mostrar el mensaje de error
    Select Case MsgBox(Excepcion, vbCritical + vbRetryCancel, "Error:" & err.Number - vbObjectError & " en " & err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print err.Description
    End Select
End Function

Sub ControlarExcepcion(obj As Object)
    Dim fd As String

    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0
    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
        fd = FreeFile
        Open "c:\excepcion.txt" For Append As fd
        Print #fd, obj.Excepcion
        Print #fd, obj.Traceback
        Print #fd, obj.XmlRequest
        Print #fd, obj.XmlResponse
        Close fd
        MsgBox "Servidor de AFIP Caido. Intentelo Mas tarde....", vbCritical, "Atención"
        bError = True
        Frame2.Visible = False
        HabilitarDetalles bCod, False
        CmdDetalle(0).Enabled = bBoto
        LblTotal.Caption = Format(dTotalAux, "#0.00")
       Grid1.Enabled = True
       Recalcular 0
       Erase nUnitario
    End If
End Sub

Private Sub GrabarTxtRsd(pRsd As Recordset)
  Dim sNom As String, sVariable As String, nPrecio As Double

  If pRsd.RecordCount <> 0 Then
     pRsd.MoveFirst
     sNom = CurDir() & "\venta.dat"
     Open sNom For Output As #1
     Do While Not pRsd.EOF
          sVariable = ""
          Print #1, pRsd!ID & ";" & pRsd!Producto & ";" & pRsd!Descripcion & ";" & pRsd!Cantidad & ";" & pRsd!PUnitario & ";" & pRsd!Total
          pRsd.MoveNext
      Loop
      Close #1
  Else
      sNom = CurDir() & "\venta.dat"
      Kill sNom
  End If
End Sub

Private Sub LlenarTxtRsd()
     Dim sNom As String, sVariable As String, nPrecio As Double
     Dim Lineas() As String

     Dim arrTem() As String
     Dim i As Byte

     
     sNom = CurDir() & "\venta.dat"
     
     If Len(Dir(sNom)) > 0 Then
        Open sNom For Input As #1
     
        LblTotal.Caption = "0.00"
      
        Do While Not EOF(1)
           Line Input #1, sVariable
           Lineas = Split(sVariable, ";")
           If UBound(Lineas) <> -1 Then
              ReDim arrTem(UBound(Lineas))
              Rsd.AddNew
              If sVariable = vbNullString Then Return
              For i = 0 To UBound(Lineas)
                  Select Case i
                         Case 0
                              Rsd!ID = Lineas(i)
                         Case 1
                           Rsd!Producto = Lineas(i)
                           Case 2
                                Rsd!Descripcion = Lineas(i)
                           Case 3
                                Rsd!Cantidad = Lineas(i)
                           Case 4
                                Rsd!PUnitario = Lineas(i)
                           Case 5
                                Rsd!Total = Lineas(i)
                    End Select
                   
              Next
              Rsd.Update
            End If
            CalcularTotales Rsd, True
        Loop
        Close #1
     End If
                         
End Sub

