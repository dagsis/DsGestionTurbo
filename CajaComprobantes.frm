VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CajaComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Caja"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8415
   Begin VB.ComboBox CmbCuenta 
      Height          =   315
      Left            =   1245
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1455
      Width           =   3075
   End
   Begin VB.TextBox TxtSucursal 
      Height          =   315
      Left            =   5475
      TabIndex        =   43
      Text            =   "0000"
      Top             =   165
      Width           =   495
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   6270
      TabIndex        =   42
      Text            =   "00000000"
      Top             =   165
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   660
      Index           =   7
      Left            =   7290
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6045
      Width           =   915
   End
   Begin VB.ComboBox CmbCaja 
      Height          =   315
      Left            =   3135
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   585
      Width           =   1530
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   6
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3135
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   5
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2388
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   660
      Index           =   4
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4572
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Buscar"
      Height          =   660
      Index           =   3
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1642
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   660
      Index           =   2
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3835
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Anular"
      Height          =   660
      Index           =   1
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   896
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   660
      Index           =   0
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   135
      Width           =   915
   End
   Begin VB.TextBox TxtMotivo 
      Height          =   315
      Left            =   1245
      MaxLength       =   60
      TabIndex        =   5
      Text            =   "TxtMotivo"
      Top             =   1050
      Width           =   5865
   End
   Begin VB.ComboBox CmbCajero 
      Height          =   315
      Left            =   5415
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   585
      Width           =   1710
   End
   Begin MSComCtl2.DTPicker DTPfecha 
      Height          =   330
      Left            =   1245
      TabIndex        =   2
      Top             =   585
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      Format          =   709427201
      CurrentDate     =   36942
   End
   Begin VB.ComboBox Cmbcomprobante 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   2400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   105
      TabIndex        =   16
      Top             =   1800
      Width           =   7095
      Begin VB.TextBox TxtId 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   975
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "TxtId"
         Top             =   2370
         Width           =   1260
      End
      Begin VB.CommandButton CmbCheques 
         Caption         =   "..."
         Height          =   270
         Left            =   210
         TabIndex        =   41
         Top             =   2385
         Width           =   585
      End
      Begin VB.TextBox TxtLibrador 
         Height          =   315
         Left            =   3105
         MaxLength       =   30
         TabIndex        =   13
         Text            =   "TxtLibrador"
         Top             =   3180
         Width           =   2385
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3960
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3960
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   2145
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3960
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3960
         Width           =   870
      End
      Begin VB.ComboBox CmbBanco 
         Height          =   315
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker DTPfechaValor 
         Height          =   315
         Left            =   945
         TabIndex        =   12
         Top             =   3180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709427201
         CurrentDate     =   36942
      End
      Begin VB.TextBox TxtCheque 
         Height          =   315
         Left            =   3705
         MaxLength       =   30
         TabIndex        =   11
         Text            =   "TxtCheque"
         Top             =   2760
         Width           =   1755
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "TxtImporte"
         Top             =   1995
         Width           =   1095
      End
      Begin VB.ComboBox CmbValor 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1995
         Width           =   2040
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   1680
         Left            =   135
         TabIndex        =   18
         Top             =   240
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   2963
         _Version        =   393216
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line7 
         X1              =   4755
         X2              =   4755
         Y1              =   4650
         Y2              =   3870
      End
      Begin VB.Line Line6 
         X1              =   135
         X2              =   135
         Y1              =   3855
         Y2              =   4680
      End
      Begin VB.Line Line5 
         X1              =   150
         X2              =   4755
         Y1              =   4665
         Y2              =   4665
      End
      Begin VB.Line Line4 
         X1              =   135
         X2              =   5595
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line3 
         X1              =   135
         X2              =   135
         Y1              =   1935
         Y2              =   3825
      End
      Begin VB.Line Line2 
         X1              =   5595
         X2              =   5595
         Y1              =   1935
         Y2              =   3825
      End
      Begin VB.Label Label14 
         Caption         =   "Librador :"
         Height          =   210
         Left            =   2310
         TabIndex        =   37
         Top             =   3180
         Width           =   765
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5505
         TabIndex        =   36
         Top             =   4500
         Width           =   1440
      End
      Begin VB.Label Label12 
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4905
         TabIndex        =   35
         Top             =   4485
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   5595
         Y1              =   3825
         Y2              =   3825
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha :"
         Height          =   225
         Left            =   285
         TabIndex        =   33
         Top             =   3180
         Width           =   630
      End
      Begin VB.Label Label10 
         Caption         =   "Banco :"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label Label9 
         Caption         =   "Nº Cheque :"
         Height          =   225
         Left            =   2715
         TabIndex        =   31
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Importe :"
         Height          =   195
         Left            =   3090
         TabIndex        =   30
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label7 
         Caption         =   "Valores :"
         Height          =   210
         Left            =   270
         TabIndex        =   29
         Top             =   1995
         Width           =   840
      End
   End
   Begin VB.Label Label15 
      Caption         =   "Cuenta :"
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   1455
      Width           =   885
   End
   Begin VB.Label Label17 
      Caption         =   "--"
      Height          =   195
      Left            =   6060
      TabIndex        =   44
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Label3 
      Caption         =   "Caja :"
      Height          =   195
      Left            =   2610
      TabIndex        =   39
      Top             =   585
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Motivo :"
      Height          =   225
      Left            =   165
      TabIndex        =   23
      Top             =   1050
      Width           =   750
   End
   Begin VB.Label Label5 
      Caption         =   "Cuenta :"
      Height          =   180
      Left            =   4770
      TabIndex        =   22
      Top             =   585
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   165
      TabIndex        =   21
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   165
      Left            =   4800
      TabIndex        =   20
      Top             =   180
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   195
      Left            =   165
      TabIndex        =   19
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "CajaComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean                     ' Suich Cabecera Venta
Dim sw1 As Boolean                    ' Suich Detalles Venta
Dim cImpor As Currency                ' Importe Total de Venta
Dim Rs_Caj As ADODB.Recordset         ' Cabecera Caja
Dim RsCv As ADODB.Recordset           ' Cabecera Virtual Caja
Dim Rsd As ADODB.Recordset            ' Cabecera Vitual CajaDetalle
Dim Rs As ADODB.Recordset             ' Cabecera Caja Tabla Todos
Dim Rs_Sis As ADODB.Recordset         ' Tabla Sistema, Ultimo Movimiento
Dim cRsl As ClsLectura
Public nMovi As Long
Public nCas As Integer, nCh As Long, nPos As Variant

Private Sub CmbBanco_Click()
  Dim cXa As ClsComprobantesL, cRx As ClsLectura
  
  Set cXa = New ClsComprobantesL
  Set cRx = New ClsLectura
  
  If sw = False Then
     If cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 3 Then
        If CmbBanco.ListIndex <> -1 Then
           TxtLibrador.text = cRx.TraerValorDeUnCampo("CuentasBancarias", "Titular", "Id=" & CmbBanco.ItemData(CmbBanco.ListIndex))
           TxtCheque.text = cRx.TraerValorDeUnCampo("CuentasBancarias", "Numero", "Id=" & CmbBanco.ItemData(CmbBanco.ListIndex))
        End If
     End If
     Set cXa = Nothing
  End If
  TodoBien
End Sub

Private Sub CmbCheques_Click()
  Dim cRx As ClsLectura
  Set cRx = New ClsLectura
  
  BancoCheques.Show 1
  If nDat <> 0 Then
     TxtId.text = nDat
     If TxtId.text <> "" Then
        If cRx.TraerValorDeUnCampo("Cheques", "Id", "id=" & TxtId.text & " AND Anulado<>1") <> 0 Then
        TxtImporte.text = Format(cRx.TraerValorDeUnCampo("Cheques", "Importe", "Id=" & TxtId.text), nCantDecimales)
        TxtLibrador.text = cRx.TraerValorDeUnCampo("Cheques", "Titular", "Id=" & TxtId.text)
        CmbBanco.text = cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & cRsl.TraerValorDeUnCampo("Cheques", "Banco", "Id=" & TxtId.text))
        TxtCheque.text = cRx.TraerValorDeUnCampo("Cheques", "Numero", "Id=" & TxtId.text)
        DTPfechaValor.Value = cRx.TraerValorDeUnCampo("Cheques", "FechaAcre", "id=" & TxtId.text)
        If CmdDetalle(0).Enabled = True Then
           CmdDetalle(0).SetFocus
        End If
        nCh = TxtId
        Else
           MsgBox "El Cheque no Existe", vbCritical, "Atención"
           TxtId.SetFocus
        End If
     End If
  End If
  Set cRx = Nothing

End Sub

Private Sub CmbComprobante_Click()
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  Dim cRx As ClsLectura
  
  On Error GoTo errHandler

  If sw = False Then
     Set cRx = New ClsLectura
     Set cNum = New ClsComprobantesL
     nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
     nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     TxtSucursal.text = nSuc
     TxtNumero.text = nNum
     CmbCajero.ListIndex = -1
     If cNum.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex)) <> 14 Then
        CmbCajero.ListIndex = -1
        CmbCajero.Enabled = False
     Else
     If cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex)) <> 27 Then
        CmbCajero.ListIndex = 0
        CmbCajero.Enabled = True
        End If
     End If
     If cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 8 Then
        cRx.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", "Grupo=4"
        CmbCuenta.text = cRx.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & cRx.TraerValorDeUnCampo("OperacionesContables", "Cuenta", "Id=3"))
     Else
        cRsl.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", "Grupo=5"
        CmbCuenta.text = cRx.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & cRx.TraerValorDeUnCampo("OperacionesContables", "Cuenta", "Id=4"))
     End If
     Exit Sub
   End If
  Set cNum = Nothing
  Set cRx = Nothing
  VerEstado
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub HabilitarValores(bCompr As Boolean, bImpor As Boolean, bCheque As Boolean, bSal As Boolean)
  CmbValor.Enabled = bCompr
  TxtImporte.Enabled = bImpor
  TxtLibrador.Enabled = bCheque
  CmbBanco.Enabled = bCheque
  TxtCheque.Enabled = bCheque
  DTPfechaValor.Enabled = bCheque
  CmbCheques.Enabled = bSal
  TxtId.Enabled = bSal
End Sub

Private Sub CmbValor_Click()
'  Dim cTip As ClsComprobantesL, cRa As ClsLectura
'  Set cRa = New ClsLectura
'  If sw = False And sw1 = False Then
'     Set cTip = New ClsComprobantesL
'     LimpiarDetalles
'     If CmbValor.ListIndex <> -1 Then
'        Select Case cTip.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex))
'               Case 1
'                    HabilitarValores True, True, False, False
'                    TxtImporte.SetFocus
'               Case 3
'                    HabilitarValores True, True, True, False
'                    CmdDetalle(0).Enabled = False
'                    cRa.CargaCombo CmbBanco, "CuentasBancarias", "Id", "Descripcion", ""
'                    TxtImporte.SetFocus
'               Case 4
'                    CmdDetalle(0).Enabled = False
'                    If cTip.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 8 Then
'                       HabilitarValores True, True, True, False
'                       TxtImporte.SetFocus
'                    Else
'                       HabilitarValores True, True, False, True
'                   End If
'                   cRa.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
'               Case 5
'                    CmdDetalle(0).Enabled = False
'                    HabilitarValores True, True, True, False
'                    TxtImporte.SetFocus
'        End Select
'     End If
'  End If
'  TodoBien
'  Set cRa = Nothing
'  Set cTip = Nothing

  Dim cTip As ClsComprobantesL, cRa As ClsLectura
  Set cRa = New ClsLectura
  If sw = False And sw1 = False Then
     Set cTip = New ClsComprobantesL
     If CmbValor.ListIndex <> -1 Then
        Select Case cTip.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex))
               Case 1
                    HabilitarValores True, True, False, False
                    CmdDetalle(0).Enabled = True
                    TxtImporte.SetFocus
               Case 3
                    HabilitarValores True, True, True, False
                    cRa.CargaCombo CmbBanco, "CuentasBancarias", "Id", "Descripcion", ""
                    CmdDetalle(0).Enabled = True
                    TxtImporte.SetFocus
               Case 4
                    If cTip.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 8 Then
                       HabilitarValores True, True, True, False
                       TxtImporte.SetFocus
                    Else
                       HabilitarValores True, True, False, True
                       TxtId.SetFocus
                   End If
                   cRa.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
                   CmdDetalle(0).Enabled = True
               Case 5
                    HabilitarValores True, True, True, False
                    CmdDetalle(0).Enabled = True
                    TxtImporte.SetFocus
        End Select
        LimpiarDetalles
     End If
  End If
  Set cRa = Nothing
  Set cTip = Nothing

End Sub

Private Sub TxtCheque_Change()
  TodoBien
End Sub

Private Sub TxtId_Change()
  TodoBien
End Sub

Private Sub TxtId_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtId_LostFocus()
  Dim cRx As ClsLectura, sId As Long
  Set cRx = New ClsLectura
  If TxtId.text <> "" Then
     sId = cRx.TraerValorDeUnCampo("Cheques", "Id", "Numero='" & TxtId.text & "' AND Anulado=0 AND FechaSalida is Null")
     If sId <> 0 Then
        TxtId.text = sId
        TxtImporte.text = Format(cRx.TraerValorDeUnCampo("Cheques", "Importe", "Id=" & sId), nCantDecimales)
        TxtLibrador.text = cRx.TraerValorDeUnCampo("Cheques", "Titular", "Id=" & sId)
        CmbBanco.text = cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & cRsl.TraerValorDeUnCampo("Cheques", "Banco", "Id=" & sId))
        TxtCheque.text = cRx.TraerValorDeUnCampo("Cheques", "Numero", "Id=" & sId)
        DTPfechaValor.Value = cRx.TraerValorDeUnCampo("Cheques", "FechaAcre", "Id=" & sId)
        CmdDetalle(0).SetFocus
        nCh = sId
     Else
        MsgBox "El Cheque no Existe o Cheque ya procesado", vbCritical, "Atención"
        TxtId.SetFocus
    End If
  End If
  Set cRx = Nothing
End Sub

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Anular
         Case 3
              Buscar
         Case 4
              Imprimir
         Case 5
              Anterior
         Case 6
              Siguiente
         Case 7
              Salir
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  TxtNumero.ForeColor = &H80000012
  TxtSucursal.ForeColor = &H80000012
  lblTotal.Caption = nCantDecimales
  TxtImporte.text = nCantDecimales
  CmbComprobante.ListIndex = -1
  CmbBanco.ListIndex = -1
  CmbValor.ListIndex = -1
  CmbCajero.ListIndex = -1
  CmbCaja.ListIndex = -1
  DtpFecha.Value = Date
  DTPfechaValor.Value = Date
End Sub
Private Sub HabilitarTodo(bEstado As Boolean, bGrid As Boolean)
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
         Or TypeOf ctl Is ComboBox Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  Grid1.Enabled = bGrid
End Sub
Private Sub Form_Load()
 On Error GoTo errHandler

 Set Rs = New ADODB.Recordset
 Set Rs_Sis = New ADODB.Recordset

 sw = True
 sw1 = True
 nLlama = 3

 Botones True, False, False, False, False, False, False, True
 BotonDetalles False, False, False, False

 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
 CmdBotones(3).Picture = LoadResPicture("Buscar", 0)
 CmdBotones(4).Picture = LoadResPicture("Imprimir", 0)
 CmdBotones(7).Picture = LoadResPicture("Salir", 0)

 CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
 CmdDetalle(1).Picture = LoadResPicture("Borrar", 0)
 CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
 CmdDetalle(3).Picture = LoadResPicture("Cancelar", 0)

 CargarCombos
 CrearRsDetalles
 
 Limpiar
 HabilitarTodo False, False
 HabilitarDetalles False
'
 ArmaRsParaMovimiento
 Me.Top = 0
 Me.Left = 0

Exit Sub

errHandler:
   ManejaErrores
End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler
  Dim cRa As ClsLectura
  
  Set cRa = New ClsLectura
  cRa.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "(TipoOperacion=3 or TipoOperacion=14) and Suc=" & nSucursal
  cRa.CargaCombo CmbValor, "CondVenta", "CondVta", "Descripcion", "Tipo=1 or Tipo=4 or Tipo=5 or Tipo=3"
  cRa.CargaCombo CmbCajero, "CuentasBancarias", "Banco", "Descripcion", ""
  cRa.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  cRa.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
  cRa.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", ""
  
  Set cRa = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bBus As Boolean, _
  bVal As Boolean, bAnt As Boolean, bSig As Boolean, bSal As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  If TxtNumero.ForeColor = &H40C0& Then
     CmdBotones(1).Enabled = False
  End If
  CmdBotones(2).Enabled = bMod
  CmdBotones(3).Enabled = bBus
  CmdBotones(4).Enabled = bVal
  CmdBotones(5).Enabled = bAnt
  CmdBotones(6).Enabled = bSig
  CmdBotones(7).Enabled = bSal
End Sub
Private Sub BotonDetalles(bAgr As Boolean, bBorr As Boolean, bMod As Boolean, bCan As Boolean)
  CmdDetalle(0).Enabled = bAgr
  CmdDetalle(1).Enabled = bBorr
  CmdDetalle(2).Enabled = bMod
  CmdDetalle(3).Enabled = bCan
End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  CmbCheques.Enabled = bHabi
  TxtId.Enabled = bHabi
  CmbValor.Enabled = bHabi
  TxtCheque.Enabled = bHabi
  CmbBanco.Enabled = bHabi
  TxtImporte.Enabled = bHabi
  TxtLibrador.Enabled = bHabi
  DTPfechaValor.Enabled = bHabi
End Sub


Private Sub TxtImporte_GotFocus()
  TxtImporte.SelStart = 0
  TxtImporte.SelLength = Len(TxtImporte.text)
End Sub
Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtImporte) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtImporte_LostFocus()
  If TxtImporte.text = "" Then TxtImporte.text = "0.00"
  TxtImporte.text = Format(TxtImporte.text, nCantDecimales)
End Sub
Private Sub CrearRsDetalles()
 On Error GoTo errHandler

 Set Rsd = New ADODB.Recordset
 Set cRsl = New ClsLectura

 Set Rsd = cRsl.RsVacio("Valores", "Id", "N")
 CabGrid

Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CabGrid()
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Visible = False
      .Columns(2).Visible = False
      .Columns(3).Caption = "Valor"
      .Columns(3).Width = 5000
      .Columns(4).Visible = False
      .Columns(5).Visible = False
      .Columns(6).Visible = False
      .Columns(7).Visible = False
      .Columns(8).Visible = False
      .Columns(9).Visible = False
      .Columns(10).Visible = False
      .Columns(11).Visible = False
      .Columns(12).Visible = False
      .Columns(13).Caption = "Importe"
      .Columns(13).Width = 1250
      .Columns(13).NumberFormat = nCantDecimales
      .Columns(13).Alignment = dbgRight
      .Columns(14).Visible = False
 End With
End Sub

Private Sub ArmaRsParaMovimiento()
  Dim cComp As ClsComprobantesL
  On Error GoTo errHandler

  Set cComp = New ClsComprobantesL
  Set Rs_Caj = cComp.TraerTodos("CabComprobantes", "Id", 3)
  Botones True, False, False, False, False, False, False, True
  If Rs_Caj.BOF = False And Rs_Caj.EOF = False Then
     Rs_Caj.MoveLast
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Caj!ID, True)
     LinkearTexto
     LLenarDetalles
     Botones True, True, True, True, True, True, True, True
     PuedoAnular
  End If
  Set cRsl = Nothing
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LinkearTexto()
  Dim nSuc As String * 4, nNum As String * 8
  Dim cComp As ClsComprobantesL
  Dim cCompr As ClsLectura, cCaja As ClsLectura
  On Error GoTo errHandler

  Set cCompr = New ClsLectura
  Set cCaja = New ClsLectura
  Set cComp = New ClsComprobantesL

  CmbComprobante.text = cCompr.DatoCombo("Comprobantes", "Id", "Descripcion", Rs!Comprobante)
  If cComp.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 8 Then
     cCompr.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", "Grupo=4"
     CmbCuenta.text = cCompr.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & cCompr.TraerValorDeUnCampo("OperacionesContables", "Cuenta", "Id=3"))
  Else
     cCompr.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", "Grupo=5"
     CmbCuenta.text = cCompr.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & cCompr.TraerValorDeUnCampo("OperacionesContables", "Cuenta", "Id=4"))
  End If
  
  CmbCuenta.text = cCaja.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & Rs!Lista)
  DtpFecha.Value = Rs!fecha
  nNum = Format(Rs!Numero, "00000000")
  nSuc = Format(Rs!Sucursal, "0000")
  TxtNumero.ForeColor = &H80000012
  TxtSucursal.ForeColor = &H80000012
  If cComp.ComprobanteAnulado(Rs!Movimiento) = True Then
     TxtNumero.ForeColor = &H40C0&
     TxtSucursal.ForeColor = &H40C0&
  End If
  TxtNumero.text = nNum
  TxtSucursal.text = nSuc
  nMovi = Rs!Movimiento
  If Rs!Vendedor <> 0 Then
     CmbCajero.text = cCaja.DatoCombo("CuentasBancarias", "Banco", "Descripcion", Rs!Vendedor)
  End If
  CmbCaja.text = cCaja.DatoCombo("Cajas", "Caja", "Descripcion", Rs!Caja)
  TxtMotivo.text = Rs!Motivo
  lblTotal.Caption = Format(Rs!Neto, nCantDecimales)
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LLenarDetalles()
  Dim rAuxD As ADODB.Recordset
  On Error GoTo errHandler

  Set rAuxD = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set rAuxD = cRsl.TraerRS("MovimientoCajaTraer", Rs!Movimiento, True)
  If rAuxD.RecordCount <> 0 Then
     rAuxD.MoveFirst
     Do While Not rAuxD.EOF
        Rsd.AddNew
        Rsd!Venta = rAuxD!Movimiento
        Rsd!formapago = rAuxD!formapago
        Rsd!Nombre = rAuxD!Nombre
        Rsd!Banco = rAuxD!Banco
        Rsd!NumeroCheque = rAuxD!NumeroCheque
        Rsd!FechaAcreditacion = rAuxD!FechaAcreditacion
        Rsd!NumeroTarjeta = rAuxD!NumeroTarjeta
        Rsd!FechaVencimiento = rAuxD!FechaVencimiento
        Rsd!Autorizacion = rAuxD!Autorizacion
        Rsd!Cupon = rAuxD!Cupon
        Rsd!importe = rAuxD!Debe + rAuxD!Haber
        Rsd!PorVenta = " "
        Rsd.Update
        rAuxD.MoveNext
     Loop
   End If
  
  Set rAuxD = cRsl.TraerRS("MovimientoBancoTraer", Rs!Movimiento, True)
  If rAuxD.RecordCount <> 0 Then
     rAuxD.MoveFirst
     Do While Not rAuxD.EOF
        Rsd.AddNew
        Rsd!Venta = rAuxD!Movimiento
        Rsd!formapago = rAuxD!formapago & " (Banco)"
        Rsd!Nombre = rAuxD!Nombre
        Rsd!Banco = rAuxD!Banco
        Rsd!NumeroCheque = rAuxD!NumeroCheque
        Rsd!FechaAcreditacion = rAuxD!FechaAcreditacion
        Rsd!NumeroTarjeta = rAuxD!NumeroTarjeta
        Rsd!FechaVencimiento = rAuxD!FechaVencimiento
        Rsd!Autorizacion = rAuxD!Autorizacion
        Rsd!Cupon = rAuxD!Cupon
        Rsd!importe = rAuxD!Debe + rAuxD!Haber
        Rsd!PorVenta = " "
        Rsd.Update
        rAuxD.MoveNext
     Loop
  End If
  
  CabGrid
  Set cRsl = Nothing
  rAuxD.Close
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LinkearDetalles()
  On Error GoTo errHandler

  Dim cValor As ClsLectura, cBanco As ClsLectura

  Set cValor = New ClsLectura
  Set cBanco = New ClsLectura
  
     CmbValor.text = Rsd!formapago
     If Not IsNull(Rsd!Banco) Then
        CmbBanco.text = cBanco.DatoCombo("Bancos", "Banco", "Descripcion", Rsd!Banco)
        TxtCheque.text = "" & Rsd!NumeroCheque
        TxtLibrador.text = "" & Rsd!Nombre
        DTPfechaValor.Value = IIf(Rsd!FechaAcreditacion = Null, Date, Rsd!FechaAcreditacion)
     End If
     TxtImporte.text = Format(Rsd!importe, nCantDecimales)
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Imprimir()
   Dim rPrint As ADODB.Recordset
   Set rPrint = New ADODB.Recordset
   
   rPrint.Fields.Append "Fecha", adDate
   rPrint.Fields.Append "Comprobante", adVarChar, 30
   rPrint.Fields.Append "Numero", adVarChar, 13
   rPrint.Fields.Append "Motivo", adVarChar, 80
   rPrint.Fields.Append "FormaPago", adVarChar, 100
   rPrint.Fields.Append "Importe", adCurrency
   rPrint.Open
        
   Rsd.MoveFirst
   Do While Not Rsd.EOF
      rPrint.AddNew
      rPrint!fecha = DtpFecha.Value
      rPrint!Comprobante = CmbComprobante.text
      rPrint!Numero = TxtSucursal.text & "-" & TxtNumero.text
      rPrint!Motivo = TxtMotivo.text
      rPrint!formapago = Trim(Rsd!formapago) & " " & Trim(Rsd!NumeroCheque) & " " & Trim(Rsd!Banco) & " " & IIf(IsNull(Rsd!NumeroCheque), "", Rsd!FechaAcreditacion)
      rPrint!importe = Rsd!importe
      rPrint.Update
      Rsd.MoveNext
   Loop
   
  nImpr = 45
  FrmImpresor.Show

End Sub

Private Sub Nuevo()
  Dim cRx As ClsLectura
  
  On Error GoTo errHandler

  Set cRx = New ClsLectura
  If sw = True Then
     Limpiar
     HabilitarTodo True, False
     HabilitarDetalles False
     Botones False, False, False, False, False, False, False, True
     BotonDetalles False, False, False, False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(7).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(7).Caption = "Cancelar"
     sw = False
     CrearRsValor
     CrearRsDetalles
     CmbComprobante.SetFocus
     nMovi = 0
     nCas = 1
     CmbCajero.ListIndex = -1
     CmbCaja.text = cRx.DatoCombo("Cajas", "Caja", "Descripcion", cRsl.RegPorDefecto("Caja"))
     CmbComprobante.ListIndex = 1
  Else
     sw = True
     Botones False, False, False, False, False, False, False, True
     GrabarTodo
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(7).Picture = LoadResPicture("Salir", 0)
     CmdBotones(7).Caption = "Salir"
     HabilitarTodo False, False
     Botones True, True, True, True, True, True, True, True
     BotonDetalles False, False, False, False
     nMovi = RsCv!Movimiento
     cRx.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", ""
     nCas = 0
  End If
  Set cRx = Nothing
Exit Sub

errHandler:
     ManejaErrores
'     Recargar
End Sub
Private Sub Buscar()
  nLlama = 8
  ComproBuscar.Show 1
  If ComproBuscar.nId <> 0 Then
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(7).Picture = LoadResPicture("Salir", 0)
     CmdBotones(7).Caption = "Salir"
     HabilitarTodo False, False
     sw = True
     Limpiar
     LimpiarDetalles
     CrearRsDetalles
     Botones True, False, False, False, False, False, False, True
     BotonDetalles False, False, False, False
     If Rs_Caj.BOF = False Or Rs_Caj.EOF Then
        Rs_Caj.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True, True, True
     End If
  End If
End Sub

Private Sub Siguiente()
  On Error GoTo errHandler
  Rs_Caj.MoveNext
  If Rs_Caj.EOF Then
     Rs_Caj.MoveLast
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Caj!ID, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  PuedoAnular
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Anterior()
  On Error GoTo errHandler
  If Rs_Caj.RecordCount <> 0 Then
     Rs_Caj.MovePrevious
     If Rs_Caj.BOF Then
        Rs_Caj.MoveFirst
        Beep
     End If
     Set cRsl = New ClsLectura
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Caj!ID, True)
     Set cRsl = Nothing
     CrearRsDetalles
     Limpiar
     LinkearTexto
     LLenarDetalles
     Botones True, True, True, True, True, True, True, True
     HabilitarTodo False, True
     BotonDetalles False, False, False, False
     PuedoAnular
  End If
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Salir()
  If sw = False Then
     If CmdBotones(0).Enabled = True Then
        If MsgBox("El Comprobante no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
           Exit Sub
        End If
     Else
        If MsgBox("Cancela la Crearción del Comprobante ?", 20, "Atención") = vbNo Then
           Exit Sub
        End If
     End If
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(7).Picture = LoadResPicture("Salir", 0)
     CmdBotones(7).Caption = "Salir"
     HabilitarTodo False, False
     sw = True
     Limpiar
     LimpiarDetalles
     CrearRsDetalles
     Botones True, False, False, False, False, False, False, True
     BotonDetalles False, False, False, False
     If Rs_Caj.BOF = False Or Rs_Caj.EOF = False Then
      '  Rs_Caj.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Caj!ID, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True, True, True
     End If
  Else
     Unload Me
  End If
End Sub

Private Sub CrearRsValor()
 On Error GoTo errHandler

 Set RsCv = New ADODB.Recordset
 Set cRsl = New ClsLectura

 Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")

Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LimpiarDetalles()
 TxtId.text = ""
 TxtImporte.text = nCantDecimales
 TxtCheque.text = ""
 CmbBanco.ListIndex = -1
 TxtLibrador.text = ""
 DTPfechaValor.Value = Date
End Sub

Private Sub GrabarTodo()
  Dim cNum As ClsComprobantesL, cTipo As ClsComprobantesL
  Dim rCaja As ADODB.Recordset, cGrab As ClsComprobantesE
  Dim rBanco As ADODB.Recordset
  
  Dim i As Byte, bIngre As Boolean
  
  Set cNum = New ClsComprobantesL
  Set cTipo = New ClsComprobantesL
  Set rCaja = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set cGrab = New ClsComprobantesE
  
  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
  Set rBanco = cRsl.RsVacio("CtaBanco", "Id", "N")
  
  RsCv.AddNew
 ' RsCv!Id = 0
  RsCv!Suc = nSucursal
  RsCv!tipo = 3
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DtpFecha.Value
  RsCv!FechaIva = DtpFecha.Value
  RsCv!Numero = TxtNumero.text
  RsCv!Sucursal = TxtSucursal.text
  RsCv!Cliente = ""
  If CmbCajero.ListIndex <> -1 Then
     RsCv!Cliente = cRsl.TraerValorDeUnCampo("Proveedores", "Proveedor", "Id=" & CmbCajero.ItemData(CmbCajero.ListIndex))
  End If
  RsCv!Lista = CmbCuenta.ItemData(CmbCuenta.ListIndex)
  If CmbCajero.ListIndex <> -1 Then
     RsCv!Vendedor = CmbCajero.ItemData(CmbCajero.ListIndex)
  Else
     RsCv!Vendedor = 0
  End If
  RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
  RsCv!CondVenta = 0
  RsCv!Neto = CCur(lblTotal.Caption)
  RsCv!Iva1 = 0
  RsCv!Iva2 = 0
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
         Case 8, 26
              RsCv!Debe = CCur(lblTotal.Caption)
              RsCv!Haber = 0
         Case 9, 27
              RsCv!Debe = 0
              RsCv!Haber = CCur(lblTotal.Caption)
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = TxtMotivo.text
  RsCv!Anulado = 0
  RsCv.Update
  If CmbCajero.ListIndex <> -1 Then
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        If Rsd!formapago <> "CHEQUES PROPIOS" Then
           rCaja.AddNew
           rCaja!Suc = nSucursal
           rCaja!Movimiento = RsCv!Movimiento
           rCaja!Comprobante = RsCv!Comprobante
           rCaja!Numero = RsCv!Numero
           rCaja!Sucursal = RsCv!Sucursal
           rCaja!fecha = DtpFecha.Value
           rCaja!Hora = Time
           rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
           rCaja!formapago = Rsd!formapago
           rCaja!Nombre = Rsd!Nombre
           rCaja!Banco = Rsd!Banco
           rCaja!NumeroCheque = Rsd!NumeroCheque
           rCaja!FechaAcreditacion = IIf(IsNull(Rsd!FechaAcreditacion), Date, Rsd!FechaAcreditacion)
           rCaja!NumeroTarjeta = ""
           rCaja!FechaVencimiento = Null
           rCaja!Autorizacion = ""
           rCaja!Cupon = ""
           Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
                  Case 8, 26
                       rCaja!Debe = 0
                       rCaja!Haber = Rsd!importe
                  Case 9, 27
                       rCaja!Debe = Rsd!importe
                       rCaja!Haber = 0
           End Select
           rCaja!Abierta = "S"
           rCaja!Anulado = 0
           rCaja.Update
        End If
        Rsd.MoveNext
     Loop
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        rBanco.AddNew
        rBanco!Movimiento = RsCv!Movimiento
        rBanco!Comprobante = RsCv!Comprobante
        rBanco!Numero = RsCv!Numero
        rBanco!Sucursal = RsCv!Sucursal
        rBanco!fecha = DtpFecha.Value
        rBanco!Hora = Time
        rBanco!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
        rBanco!formapago = Rsd!formapago
        rBanco!Nombre = Rsd!Nombre
        rBanco!Banco = Rsd!Banco
        rBanco!NumeroCheque = Rsd!NumeroCheque
        rBanco!FechaAcreditacion = IIf(IsNull(Rsd!FechaAcreditacion), Date, Rsd!FechaAcreditacion)
        rBanco!NumeroTarjeta = ""
        rBanco!FechaVencimiento = Null
        rBanco!Autorizacion = ""
        rBanco!Cupon = ""
        Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
               Case 9, 27
                    rBanco!Debe = 0
                    rBanco!Haber = Rsd!importe
                Case 8, 26
                   rBanco!Debe = Rsd!importe
                   rBanco!Haber = 0
       End Select
       rBanco!Abierta = "S"
       rBanco!Anulado = 0
       rBanco.Update
       Rsd.MoveNext
    Loop
 Else
   Rsd.MoveFirst
   Do While Not Rsd.EOF
      If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(Rsd!formapago) & "'") = "Caja" Then
         rCaja.AddNew
         rCaja!Suc = nSucursal
         rCaja!Movimiento = RsCv!Movimiento
         rCaja!Comprobante = RsCv!Comprobante
         rCaja!Numero = RsCv!Numero
         rCaja!Sucursal = RsCv!Sucursal
         rCaja!fecha = DtpFecha.Value
         rCaja!Hora = Time
         rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
         rCaja!formapago = Rsd!formapago
         rCaja!Nombre = Rsd!Nombre
         rCaja!Banco = Rsd!Banco
         rCaja!NumeroCheque = Rsd!NumeroCheque
         rCaja!FechaAcreditacion = IIf(IsNull(Rsd!FechaAcreditacion), Date, Rsd!FechaAcreditacion)
         rCaja!NumeroTarjeta = ""
         rCaja!FechaVencimiento = Null
         rCaja!Autorizacion = ""
         rCaja!Cupon = ""
         Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
                Case 8, 26
                     rCaja!Debe = Rsd!importe
                     rCaja!Haber = 0
                Case 9, 27
                     rCaja!Debe = 0
                     rCaja!Haber = Rsd!importe
         End Select
         rCaja!Abierta = "S"
         rCaja!Anulado = 0
         rCaja.Update
      Else
         If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(Rsd!formapago) & "'") = "Banco" Then
            rBanco.AddNew
            rBanco!Movimiento = RsCv!Movimiento
            rBanco!Comprobante = RsCv!Comprobante
            rBanco!Numero = RsCv!Numero
            rBanco!Sucursal = RsCv!Sucursal
            rBanco!fecha = DtpFecha.Value
            rBanco!Hora = Time
            rBanco!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
            rBanco!formapago = Rsd!formapago
            rBanco!Nombre = Rsd!Nombre
            rBanco!Banco = Rsd!Banco
            rBanco!NumeroCheque = Rsd!NumeroCheque
            rBanco!FechaAcreditacion = Rsd!FechaAcreditacion
            rBanco!NumeroTarjeta = Rsd!NumeroTarjeta
            rBanco!FechaVencimiento = Rsd!FechaVencimiento
            rBanco!Autorizacion = Rsd!Autorizacion
            rBanco!Cupon = Rsd!Cupon
            If cTipo.SumaCaja(rBanco!Comprobante) Then
               rBanco!Debe = Rsd!importe
               rBanco!Haber = 0
            Else
               rBanco!Debe = 0
               rBanco!Haber = Rsd!importe
            End If
            rBanco!Abierta = "S"
            rBanco!Anulado = 0
            rBanco.Update
        End If
     End If
     Rsd.MoveNext
   Loop
 
 End If
 If CmdBotones(2).Caption = "Grabar" Then
    ' cgrab.Modificar rscv,rcaja
  Else
     Dim na As Integer
     na = cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     bIngre = IIf(na = 8, True, False)
     cGrab.GrabarCaja RsCv, rCaja, rBanco, Rsd, bIngre
  End If

  
 CrearRsDetalles
 Limpiar
 
 ArmaRsParaMovimiento
  
  ' Fin del proceso de facturacion fiscal
  
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing

End Sub

Private Sub VerEstado()
  CmdBotones(0).Enabled = False
  If Rsd.RecordCount <> 0 Then
     CmdBotones(0).Enabled = True
  End If
End Sub

Private Sub VerCabecera()
  CmdDetalle(0).Enabled = False
  If Len(TxtMotivo.text) <> 0 Then
     CmdDetalle(0).Enabled = True
  End If
End Sub

Private Sub TodoBien()
  Dim cTip As ClsComprobantesL
  CmdDetalle(0).Enabled = False
  If sw = False And sw1 = False Then
     Set cTip = New ClsComprobantesL
     If CmbValor.ListIndex <> -1 Then
        Select Case cTip.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex))
               Case 1
                     If Len(TxtImporte.text) <> 0 And Val(TxtImporte.text) <> 0 Then
                        CmdDetalle(0).Enabled = True
                     End If
               Case 3
                     If Len(TxtImporte.text) <> 0 And Len(CmbBanco.text) <> 0 And Len(TxtNumero.text) <> 0 And Len(TxtLibrador.text) <> 0 And Val(TxtImporte.text) <> 0 Then
                        CmdDetalle(0).Enabled = True
                     End If
               Case 4
                     If Len(TxtImporte.text) <> 0 And Len(CmbBanco.text) <> 0 And Len(TxtNumero.text) <> 0 And Len(TxtLibrador.text) <> 0 And Val(TxtImporte.text) <> 0 Then
                        CmdDetalle(0).Enabled = True
                     End If
               Case 5
        End Select
     End If
  End If
  Set cTip = Nothing
End Sub

Private Sub TxtLibrador_Change()
  TodoBien
End Sub

Private Sub TxtMotivo_Change()
  VerCabecera
End Sub

Private Sub CmdDetalle_Click(Index As Integer)
  Select Case Index
         Case 0
              NuevoRsD
         Case 1
              BorrarRsD
         Case 2
              ModificarRsD
         Case 3
              CancelarRsd
  End Select
End Sub

Private Sub NuevoRsD()
  Dim cTip  As ClsComprobantesL
  
  Set cTip = New ClsComprobantesL
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     CmdBotones(7).Enabled = False
     HabilitarTodo False, False
     HabilitarDetalles True
     LimpiarDetalles
     BotonDetalles True, False, False, True
     CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(0).Caption = "Grabar"
     If cTip.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 14 Then
        CmbValor.ListIndex = 1
     Else
        CmbValor.ListIndex = 0
     End If
     CmbValor.SetFocus
  Else
     CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
     CmdDetalle(0).Caption = "Agregar"
     Rsd.AddNew
     GrabarRsd
     If cTip.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 4 Then
        nCh = nCh + 1
     End If
     Rsd.MoveLast
     Grid1.Bookmark = Rsd.Bookmark
     BotonDetalles True, True, True, False
     sw1 = True
     HabilitarTodo True, True
     HabilitarDetalles False
     CmdBotones(0).Enabled = True
     CmdBotones(7).Enabled = True
  End If
  Set cTip = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub BorrarRsD()
  On Error GoTo errHandler
  If MsgBox("Borra el Item Seleccionado ?", 20, "Atención") = 6 Then
     LimpiarDetalles
     BotonDetalles True, False, False, False
     Rsd.Delete
     Rsd.MovePrevious
     If Rsd.BOF = False And Rsd.EOF = False Then
        If Rsd.BOF Then
           Rsd.MoveFirst
        End If
        LinkearDetalles
        Grid1.Bookmark = Rsd.Bookmark
        BotonDetalles True, True, True, False
        CmdBotones(0).Enabled = False
        HabilitarValores False, False, False, False
     Else
        Botones False, False, False, False, False, True, False, True
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub ModificarRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     HabilitarTodo True, False
     HabilitarDetalles True
     BotonDetalles False, False, True, True
     CmdDetalle(2).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(2).Caption = "Grabar"
  Else
     CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
     CmdDetalle(2).Caption = "Actualizar"
     sw1 = True
     GrabarRsd
     BotonDetalles True, True, True, False
     CalcularTotales
     Grid1.Bookmark = nPos
     HabilitarTodo True, True
     HabilitarDetalles False
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CancelarRsd()

  CmdBotones(5).Enabled = True
  If sw1 = False Then
     If CmdDetalle(0).Caption = "Grabar" Then
        If CmdDetalle(0).Enabled = True Then
           If MsgBox("El Detalle no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
        CmdDetalle(0).Caption = "Nuevo"
        HabilitarTodo True, True
        LimpiarDetalles
        Botones False, False, False, False, False, False, False, True
        If Rsd.EOF = False Or Rsd.BOF = False Then
           BotonDetalles True, True, True, False
           Botones True, False, False, False, False, False, False, True
           LinkearDetalles
        End If
        BotonDetalles True, False, False, False
        HabilitarDetalles False
        sw1 = True
     Else
        CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
        CmdDetalle(2).Caption = "Actualizar"
        BotonDetalles True, True, True, False
        HabilitarTodo True, True
        HabilitarDetalles False
        LinkearDetalles
        sw1 = True
     End If
  End If

End Sub
Private Sub GrabarRsd()
  Dim cXa As ClsComprobantesL, cRse As ClsEscritura
  
  On Error GoTo errHandler

  Set cXa = New ClsComprobantesL
  Set cRse = New ClsEscritura
  
 ' Rsd!Id = 0
  Rsd!formapago = CmbValor.text
  Rsd!importe = TxtImporte.text
  Rsd!IdCheque = 0
  If cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 4 Then
     Rsd!NumeroCheque = TxtCheque.text
     Rsd!Banco = CmbBanco.ItemData(CmbBanco.ListIndex)
     If CmdDetalle(0).Enabled = True Then
        Rsd!IdCheque = IIf(nCh = 0, -1, nCh)
     End If
     Rsd!Nombre = TxtLibrador.text
     Rsd!FechaAcreditacion = DTPfechaValor.Value
  Else
     If cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 3 Then
        Rsd!NumeroCheque = TxtCheque.text
        Rsd!Banco = CmbBanco.ItemData(CmbBanco.ListIndex)
        Rsd!Nombre = TxtLibrador.text
        Rsd!FechaAcreditacion = DTPfechaValor.Value
     Else
        Rsd!Nombre = TxtLibrador.text
     End If
  End If
  If cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 3 Then
     cRse.GrabarValordeUnCampo "CuentasBancarias", "Numero", "Numero+1", "Id=" & CmbBanco.ItemData(CmbBanco.ListIndex)
  End If

  Rsd.Update
  CabGrid
  VerEstado
  CalcularTotales
  
  Set cXa = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler

  If sw = False And sw1 = True Then
     If Grid1.Row <> -1 Then
        nPos = Grid1.Bookmark
        LimpiarDetalles
        LinkearDetalles
     End If
     BotonDetalles True, True, True, False
  End If
  HabilitarValores False, False, False, False
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CalcularTotales()
  Dim rSuma As ADODB.Recordset

  Set rSuma = New ADODB.Recordset
  Set rSuma = Rsd
  rSuma.MoveFirst
  lblTotal.Caption = "0.00"
  Do While Not rSuma.EOF
     lblTotal.Caption = lblTotal.Caption + rSuma!importe
     rSuma.MoveNext
  Loop
  lblTotal.Caption = Format(lblTotal.Caption, nCantDecimales)
End Sub
Private Sub Anular()
  Dim cBor As ClsComprobantesE, cLer As ClsComprobantesL
  
  Set cBor = New ClsComprobantesE
  Set cLer = New ClsComprobantesL
  
  Dim nOpe As Integer, nTipo As Integer
   
  nOpe = cLer.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  nTipo = cLer.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  
  
  If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
     cBor.AnularComprobantes Rs!Movimiento, nOpe, nTipo
     Rs_Caj.Delete
     CrearRsDetalles
     Limpiar
     Botones True, False, False, False, False, False, False, True
     Anterior
  End If
End Sub

Private Sub TxtSucursal_GotFocus()
  TxtSucursal.SelStart = 0
  TxtSucursal.SelLength = Len(TxtSucursal.text)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSucursal_LostFocus()
  If TxtSucursal.text = "" Then TxtSucursal.text = 1
  TxtSucursal.text = Format(TxtSucursal.text, "0000")
End Sub

Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtNumero_LostFocus()
  If TxtNumero.text = "" Then TxtNumero.text = 0
  TxtNumero.text = Format(TxtNumero.text, "00000000")
End Sub

Private Sub PuedoAnular()
  Dim cRx As ClsLectura
  
  Set cRx = New ClsLectura
  
  If cRx.TraerCantidad("Cheques", "Id", "Movimiento=" & Rs!Movimiento & " AND FechaSalida is not Null") >= 1 Then
     Botones True, False, True, True, True, True, True, True
  End If
End Sub
