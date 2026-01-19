VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasAdelantos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adelantos a Proveedores"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   8355
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
      TabIndex        =   28
      Top             =   1380
      Width           =   7095
      Begin VB.ComboBox CmbValor 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1995
         Width           =   3120
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4905
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "TxtImporte"
         Top             =   1995
         Width           =   1290
      End
      Begin VB.TextBox TxtCheque 
         Height          =   315
         Left            =   4425
         MaxLength       =   30
         TabIndex        =   9
         Text            =   "TxtCheque"
         Top             =   2760
         Width           =   1755
      End
      Begin VB.ComboBox CmbBanco 
         Height          =   315
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2760
         Width           =   2505
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3960
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   2145
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3960
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3960
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3960
         Width           =   870
      End
      Begin VB.TextBox TxtLibrador 
         Height          =   315
         Left            =   3810
         MaxLength       =   30
         TabIndex        =   11
         Text            =   "TxtLibrador"
         Top             =   3180
         Width           =   2385
      End
      Begin VB.CommandButton CmbCheques 
         Caption         =   "..."
         Height          =   270
         Left            =   210
         TabIndex        =   29
         Top             =   2385
         Width           =   585
      End
      Begin VB.TextBox TxtId 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   975
         MaxLength       =   15
         TabIndex        =   7
         Text            =   "TxtId"
         Top             =   2370
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPfechaValor 
         Height          =   315
         Left            =   945
         TabIndex        =   10
         Top             =   3180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   708247553
         CurrentDate     =   36942
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   1680
         Left            =   135
         TabIndex        =   33
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
      Begin VB.Label Label7 
         Caption         =   "Valores :"
         Height          =   210
         Left            =   270
         TabIndex        =   43
         Top             =   1995
         Width           =   840
      End
      Begin VB.Label Label8 
         Caption         =   "Importe :"
         Height          =   195
         Left            =   4155
         TabIndex        =   42
         Top             =   1980
         Width           =   705
      End
      Begin VB.Label Label9 
         Caption         =   "Nº Cheque :"
         Height          =   225
         Left            =   3435
         TabIndex        =   41
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Banco :"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha :"
         Height          =   225
         Left            =   285
         TabIndex        =   39
         Top             =   3180
         Width           =   630
      End
      Begin VB.Line Line1 
         X1              =   780
         X2              =   6240
         Y1              =   3825
         Y2              =   3825
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
         Left            =   4920
         TabIndex        =   38
         Top             =   4500
         Width           =   660
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
         Left            =   5670
         TabIndex        =   37
         Top             =   4500
         Width           =   1275
      End
      Begin VB.Label Label14 
         Caption         =   "Librador :"
         Height          =   210
         Left            =   3090
         TabIndex        =   36
         Top             =   3180
         Width           =   765
      End
      Begin VB.Line Line2 
         X1              =   6225
         X2              =   6225
         Y1              =   1935
         Y2              =   3825
      End
      Begin VB.Line Line3 
         X1              =   135
         X2              =   135
         Y1              =   1935
         Y2              =   3825
      End
      Begin VB.Line Line4 
         X1              =   765
         X2              =   6225
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line5 
         X1              =   150
         X2              =   4755
         Y1              =   4665
         Y2              =   4665
      End
      Begin VB.Line Line6 
         X1              =   135
         X2              =   135
         Y1              =   3855
         Y2              =   4680
      End
      Begin VB.Line Line7 
         X1              =   4755
         X2              =   4755
         Y1              =   4650
         Y2              =   3870
      End
      Begin VB.Label LblId 
         Caption         =   "LblId"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1185
         TabIndex        =   35
         Top             =   3585
         Width           =   780
      End
      Begin VB.Label Label13 
         Caption         =   "Id Cheque :"
         Height          =   210
         Left            =   315
         TabIndex        =   34
         Top             =   3585
         Width           =   900
      End
   End
   Begin VB.TextBox TxtCliente 
      Height          =   315
      Left            =   1185
      MaxLength       =   15
      TabIndex        =   3
      Text            =   "TxtCliente"
      Top             =   1065
      Width           =   795
   End
   Begin VB.CommandButton CmdProv 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   6375
      TabIndex        =   4
      Top             =   1065
      Width           =   735
   End
   Begin VB.ComboBox Cmbcomprobante 
      Height          =   315
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   3600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   660
      Index           =   0
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Aplicacion"
      Height          =   660
      Index           =   2
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3810
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Buscar"
      Height          =   660
      Index           =   3
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   660
      Index           =   4
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4545
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
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2325
      Width           =   915
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
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3075
      Width           =   915
   End
   Begin VB.ComboBox CmbCaja 
      Height          =   315
      Left            =   2895
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   1530
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   660
      Index           =   7
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   915
   End
   Begin MSComCtl2.DTPicker DTPfecha 
      Height          =   315
      Left            =   1185
      TabIndex        =   1
      Top             =   630
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   708378625
      CurrentDate     =   36942
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Anular"
      Height          =   660
      Index           =   1
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   915
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
      Height          =   315
      Left            =   2070
      TabIndex        =   27
      Top             =   1065
      Width           =   4245
   End
   Begin VB.Label Label6 
      Caption         =   "Proveedor :"
      Height          =   255
      Left            =   90
      TabIndex        =   26
      Top             =   1065
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   195
      Left            =   90
      TabIndex        =   25
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   165
      Left            =   4830
      TabIndex        =   24
      Top             =   150
      Width           =   690
   End
   Begin VB.Label LblNumero 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblNumero"
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
      Left            =   5685
      TabIndex        =   23
      Top             =   150
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   90
      TabIndex        =   22
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Caja :"
      Height          =   195
      Left            =   2460
      TabIndex        =   21
      Top             =   630
      Width           =   420
   End
End
Attribute VB_Name = "ComprasAdelantos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean                     ' Suich Cabecera Venta
Dim sw1 As Boolean                    ' Suich Detalles Venta
Dim cImpor As Currency                ' Importe Total de Venta
Dim Rs_Caj As ADODB.Recordset         ' Cabecera Caja
Dim RsCv As ADODB.Recordset           ' Cabecera Virtual Adelanto
Dim Rsd As ADODB.Recordset            ' Cabecera Vitual AdelntoDetalle
Dim Rs As ADODB.Recordset             ' Cabecera Adelanto Tabla Todos
Dim Rs_Sis As ADODB.Recordset         ' Tabla Sistema, Ultimo Movimiento
Dim cRsl As ClsLectura
Private nMovi As Long
Private nCas As Integer


Dim rReci As ADODB.Recordset, rCaja As ADODB.Recordset
Dim rValo As ADODB.Recordset, rApli As ADODB.Recordset
Dim rAde As ADODB.Recordset, nCh As Long
Dim rBanco As ADODB.Recordset

Private Sub CmbBanco_Click()
  Dim cXa As ClsComprobantesL, cRx As ClsLectura
  
  Set cXa = New ClsComprobantesL
  Set cRx = New ClsLectura
  
  If sw = False Then
     If cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 3 Or cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 12 Then
        If CmbBanco.ListIndex <> -1 Then
           TxtLibrador.text = cRx.TraerValorDeUnCampo("CuentasBancarias", "Titular", "Id=" & CmbBanco.ItemData(CmbBanco.ListIndex))
           TxtCheque.text = cRx.TraerValorDeUnCampo("CuentasBancarias", "Numero", "Id=" & CmbBanco.ItemData(CmbBanco.ListIndex))
           CmdDetalle(0).Enabled = True
        End If
     End If
     Set cXa = Nothing
  End If
End Sub

Private Sub CmbCheques_Click()
  Dim cRx As ClsLectura
  Set cRx = New ClsLectura
  
  BancoCheques.Show 1
  If nDat <> 0 Then
     TxtId.text = nDat
     If TxtId.text <> "" Then
        If cRx.TraerValorDeUnCampo("Cheques", "Id", "id=" & TxtId.text & " AND Anulado<>1") <> 0 Then
        TxtImporte.text = Format(cRx.TraerValorDeUnCampo("Cheques", "Importe", "Id=" & TxtId.text), "#0.00")
        TxtLibrador.text = cRx.TraerValorDeUnCampo("Cheques", "Titular", "Id=" & TxtId.text)
        CmbBanco.text = cRx.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & cRx.TraerValorDeUnCampo("Cheques", "Banco", "Id=" & TxtId.text))
        TxtCheque.text = cRx.TraerValorDeUnCampo("Cheques", "Numero", "Id=" & TxtId.text)
        DTPfechaValor.Value = cRx.TraerValorDeUnCampo("Cheques", "FechaAcre", "Id=" & TxtId.text)
        CmdDetalle(0).Enabled = True
        CmdDetalle(0).SetFocus
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
  On Error GoTo errHandler

  If sw = False Then
     Set cNum = New ClsComprobantesL
     nNum = Format(cNum.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
     nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     LblNumero.Caption = nSuc & "-" & nNum
  End If
  Set cNum = Nothing
  VerEstado
Exit Sub

errHandler:
   ManejaErrores

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
        sw = True
        TxtId.text = sId
        TxtImporte.text = Format(cRx.TraerValorDeUnCampo("Cheques", "Importe", "Id=" & TxtId.text), "#0.00")
        TxtLibrador.text = cRx.TraerValorDeUnCampo("Cheques", "Titular", "Id=" & TxtId.text)
        CmbBanco.text = cRx.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & cRx.TraerValorDeUnCampo("Cheques", "Banco", "Id=" & TxtId.text))
        TxtCheque.text = cRx.TraerValorDeUnCampo("Cheques", "Numero", "Id=" & TxtId.text)
        DTPfechaValor.Value = cRx.TraerValorDeUnCampo("Cheques", "FechaAcre", "Id=" & TxtId.text)
        sw = False
        CmdDetalle(0).Enabled = True
        CmdDetalle(0).SetFocus
        nCh = TxtId
     Else
        MsgBox "El Cheque no Existe", vbCritical, "Atención"
        TxtId.SetFocus
    End If
  End If
  Set cRx = Nothing

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
  Dim cTip As ClsComprobantesL, cRa As ClsLectura
  Set cRa = New ClsLectura
  If sw = False And sw1 = False Then
     Set cTip = New ClsComprobantesL
     If CmbValor.ListIndex <> -1 Then
        Select Case cTip.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex))
               Case 1
                    HabilitarValores True, True, False, False
                    CmdDetalle(0).Enabled = True
                    LblId.Caption = "000000"
                    TxtImporte.SetFocus
               Case 3, 12
                    HabilitarValores True, True, True, False
                    cRa.CargaCombo CmbBanco, "CuentasBancarias", "Id", "Descripcion", ""
                    LblId.Caption = "000000"
                    CmdDetalle(0).Enabled = False
                    If CmbBanco.ListIndex <> -1 Then
                       CmdDetalle(0).Enabled = True
                    End If
                    TxtImporte.SetFocus
               Case 4
                    If cTip.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 8 Then
                      HabilitarValores True, True, True, False
                      LblId.Caption = Format(nCh, "000000")
                      TxtImporte.SetFocus
                    Else
                       LblId.Caption = Format(0, "000000")
                       HabilitarValores True, True, False, True
                       TxtId.SetFocus
                   End If
                   cRa.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
                   CmdDetalle(0).Enabled = True
                   CmdDetalle(0).Enabled = False
                   If CmbBanco.ListIndex <> -1 Then
                      CmdDetalle(0).Enabled = True
                    End If
               Case 5
                    HabilitarValores True, True, True, False
                    CmdDetalle(0).Enabled = True
                    LblId.Caption = "000000"
                    TxtImporte.SetFocus
        End Select
        LimpiarDetalles
     End If
  End If
  Set cRa = Nothing
  Set cTip = Nothing

End Sub

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Anular
         Case 2
              Aplicacion
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
Private Sub Aplicacion()
 Dim cTi As ClsComprobantesL, nCom As Integer
 
 Set cTi = New ClsComprobantesL
 
 nCom = CmbComprobante.ItemData(CmbComprobante.ListIndex)
 
' If cTi.TraerTipoMovimiento(nCom) = 1 Or cTi.TraerTipoMovimiento(nCom) = 24 Or cTi.TraerTipoMovimiento(nCom) = 25 Then
     ComprasAdeAplica.Show
' End If

End Sub

Private Sub Anular()
  Dim cBor As ClsComprobantesE, nMovAs As Long, cRx As ClsLectura
  
  Set cBor = New ClsComprobantesE
  Set cRx = New ClsLectura
  
  If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
     nMovAs = cRx.TraerValorDeUnCampo("CabComprobantes", "Movimiento", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Sucursal=" & Rs!Sucursal & " AND Numero=" & Rs!Numero)
     cBor.AnularAdeProveedores CmbComprobante.ItemData(CmbComprobante.ListIndex), Rs!Sucursal, Rs!Numero, Rs!Movimiento, True
     Rs_Caj.Delete
     Limpiar
     If Rs_Caj.RecordCount <> 0 Then
        Anterior
        Botones True, True, True, True, True, True, True, True, True
     Else
        Botones True, False, False, False, False, False, False, True, False
        BotonDetalles False, False, False, False
        CrearRsDetalles
     End If
  End If
  Set cRx = Nothing
End Sub
Private Sub CmdProv_Click()
  Dim RsC As ADODB.Recordset
  On Error GoTo errHandler
  
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 2
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ProveedorTraerUno", nDat, True)
     TxtCliente.text = RsC!Proveedor
     LblCliente.Caption = RsC!RazonSocial
     nLlama = 1
     RsC.Close
     CmdDetalle(0).SetFocus
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LblCliente_Change()
  VerCabecera
End Sub
Private Sub TxtCliente_GotFocus()
  TxtCliente.SelStart = 0
  TxtCliente.SelLength = Len(TxtCliente.text)
End Sub
Private Sub TxtCliente_LostFocus()
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set cRcL = New ClsClienteL
  LblCliente.Caption = ""
  If TxtCliente.text <> "" Then
     If cRcL.BuscarNombreProveedor(TxtCliente.text) <> "" Then
        LblCliente.Caption = cRcL.BuscarNombreProveedor(TxtCliente.text)
        CmdDetalle(0).SetFocus
        nLlama = 1
     Else
        MsgBox "El Proveedor no es el Correto", vbCritical, "Atención"
        CmdProv.SetFocus
     End If
  End If
  Set cRcL = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
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
  LblNumero.Caption = "0000-00000000"
  lblTotal.Caption = "0.00"
  TxtImporte.text = "0.00"
  LblCliente.Caption = ""
  CmbComprobante.ListIndex = -1
  CmbBanco.ListIndex = -1
  CmbValor.ListIndex = -1
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

 Botones True, False, False, False, False, False, False, True, False
 BotonDetalles False, False, False, False

 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(2).Picture = LoadResPicture("Valores", 0)
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
 LblId.Caption = "00000"
 nCh = Format(cRsl.TraerValorDeUnCampo("Sistema", "IdCheque", ""), "00000")

 ArmaRsParaMovimiento
 Me.Top = 0
 Me.Left = 0

Exit Sub

errHandler:
   ManejaErrores
End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=9 and Suc=" & nSucursal

  Set cRsl = New ClsLectura
'  cRsl.CargaCombo CmbValor, "CondVenta", "CondVta", "Descripcion", "Tipo<=4"
  cRsl.CargaCombo CmbValor, "CondVenta", "CondVta", "Descripcion", "tipo=1 or tipo=3 or tipo=4 or tipo=12"

  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
  Set cRsl = New ClsLectura

Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bBus As Boolean, _
  bVal As Boolean, bAnt As Boolean, bSig As Boolean, bSal As Boolean, bProv As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  CmdBotones(2).Enabled = bMod
  CmdBotones(3).Enabled = bBus
  CmdBotones(4).Enabled = bVal
  CmdBotones(5).Enabled = bAnt
  CmdBotones(6).Enabled = bSig
  CmdBotones(7).Enabled = bSal
  CmdProv.Enabled = bProv
End Sub
Private Sub BotonDetalles(bAgr As Boolean, bBorr As Boolean, bMod As Boolean, bCan As Boolean)
  CmdDetalle(0).Enabled = bAgr
  CmdDetalle(1).Enabled = bBorr
  CmdDetalle(2).Enabled = bMod
  CmdDetalle(3).Enabled = bCan
End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtId.Enabled = bHabi
  CmbValor.Enabled = bHabi
  TxtCheque.Enabled = bHabi
  CmbBanco.Enabled = bHabi
  TxtImporte.Enabled = bHabi
  TxtLibrador.Enabled = bHabi
  DTPfechaValor.Enabled = bHabi
  CmbCheques.Enabled = bHabi
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
  TxtImporte.text = Format(TxtImporte.text, "#0.00")
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
      .Columns(13).Alignment = dbgRight
      .Columns(13).Caption = "Importe"
      .Columns(13).Width = 1250
      .Columns(13).NumberFormat = "#0.00"
      .Columns(14).Visible = False
 End With
End Sub

Private Sub ArmaRsParaMovimiento()
  Dim cComp As ClsComprobantesL
  On Error GoTo errHandler

  Set cComp = New ClsComprobantesL
  Set Rs_Caj = cComp.TraerTodosCon("CabComprobantes", "Id", "Tipo=9 and Motivo='ADE' and Suc=" & nSucursal)
  Botones True, False, False, False, False, False, False, True, False
  If Rs_Caj.BOF = False And Rs_Caj.EOF = False Then
     Rs_Caj.MoveLast
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Caj!ID, True)
     LinkearTexto
     LLenarDetalles
     Botones True, True, True, True, True, True, True, True, False
  End If
  Set cRsl = Nothing
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LinkearTexto()
  Dim nSuc As String * 4, nNum As String * 8, cRcL As ClsClienteL

  Dim cCompr As ClsLectura, cCaja As ClsLectura, cComp As ClsComprobantesL
  On Error GoTo errHandler

  Set cCompr = New ClsLectura
  Set cCaja = New ClsLectura
  Set cRcL = New ClsClienteL
  Set cComp = New ClsComprobantesL

  CmbComprobante.text = cCompr.DatoCombo("Comprobantes", "Id", "Descripcion", Rs!Comprobante)
  DtpFecha.Value = Rs!fecha
  nNum = Format(Rs!Numero, "00000000")
  nSuc = Format(Rs!Sucursal, "0000")
  LblNumero.ForeColor = &H80000012
  If cComp.ComprobanteAnulado(Rs!Movimiento) = True Then
     LblNumero.ForeColor = &H40C0&
  End If
  LblNumero.Caption = nSuc & "-" & nNum
  nMovi = Rs!Movimiento
  nMoviAde = Rs!Movimiento
  CmbCaja.text = cCaja.DatoCombo("Cajas", "Caja", "Descripcion", Rs!Caja)
  TxtCliente.text = Rs!Cliente
  LblCliente.Caption = cRcL.BuscarNombreProveedor(Rs!Cliente)
  lblTotal.Caption = Format(Rs!Neto, "#0.00")
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LLenarDetalles()
  Dim rAuxD As ADODB.Recordset, nMo As Long
  On Error GoTo errHandler

  Set rAuxD = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set rAuxD = cRsl.TraerRS("MovimientoCajaTraer", Rs!Movimiento, True)
  
  If rAuxD.RecordCount = 0 Then
     nMo = cRsl.TraerValorDeUnCampo("Caja", "Movimiento", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Numero=" & Rs!Numero & " AND Sucursal=" & Rs!Sucursal)
     If nMo <> 0 Then
        Set rAuxD = cRsl.TraerRS("MovimientoCajaTraer", nMo, True)
     End If
  End If
  
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
  
  Set rAuxD = cRsl.TraerRS("MovimientoBancoTraer", Rs!Movimiento, True)
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

  Set cRsl = Nothing
  rAuxD.Close
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LinkearDetalles()
  On Error GoTo errHandler

  Dim cValor As ClsLectura, cBanco As ClsLectura, cTip As ClsComprobantesL

  Set cValor = New ClsLectura
  Set cBanco = New ClsLectura
  Set cTip = New ClsComprobantesL
  
  If Not IsNull(Rsd!formapago) Then
     CmbValor.text = Rsd!formapago
     If Not IsNull(Rsd!Banco) Then
     If cTip.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 4 Then
        CmbBanco.text = cBanco.DatoCombo("Bancos", "Banco", "Descripcion", Rsd!Banco)
     Else
        CmbBanco.text = cBanco.DatoCombo("CuentasBancarias", "Id", "Descripcion", Rsd!Banco)
     End If
        TxtCheque.text = Rsd!NumeroCheque
     End If
     TxtImporte.text = Format(Rsd!importe, "#0.00")
     TxtLibrador.text = Rsd!Nombre
'  DTPfechaValor.Value = IIf(RsD!FechaAcreditacion = Null, Date, RsD!FechaAcreditacion)
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Nuevo()
  On Error GoTo errHandler

  If sw = True Then
     Limpiar
     HabilitarTodo True, False
     HabilitarDetalles False
     Botones False, False, False, False, False, False, False, True, True
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
     CmbComprobante.ListIndex = 0
     CmbCaja.ListIndex = 0
  Else
     sw = True
     Botones False, False, False, False, False, False, False, True, False
     GrabarTodo
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(7).Picture = LoadResPicture("Salir", 0)
     CmdBotones(7).Caption = "Salir"
     HabilitarTodo False, False
     Botones True, True, True, True, True, True, True, True, False
     BotonDetalles False, False, False, False
     nMovi = RsCv!Movimiento
     nCas = 0
  End If
Exit Sub

errHandler:
     ManejaErrores
'     Recargar
End Sub
Private Sub Buscar()
  nLlama = 7
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
     Botones True, False, False, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     If Rs_Caj.BOF = False Or Rs_Caj.EOF Then
        Rs_Caj.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True, True, True, False
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
  Botones True, True, True, True, True, True, True, True, False
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Anterior()
  On Error GoTo errHandler
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
  Botones True, True, True, True, True, True, True, True, False
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
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
     Botones True, False, False, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     If Rs_Caj.BOF = False Or Rs_Caj.EOF = False Then
        Rs_Caj.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Caj!ID, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True, True, True, False
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
' CmbValor.ListIndex = -1
 TxtId.text = ""
 TxtImporte.text = "0.00"
 TxtCheque.text = ""
 CmbBanco.ListIndex = -1
 TxtLibrador.text = ""
 DTPfechaValor.Value = Date
End Sub

Private Sub GrabarTodo()
  On Error GoTo Errores
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cStock As ClsComprobantesL
  
  Dim nNum As String * 8, nSuc As String * 4
  Dim i As Byte
  
  Dim Rx As ADODB.Recordset
  Dim Rx1 As ADODB.Recordset

  Set cRsl = New ClsLectura
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  
 ' cSaldo = 0

  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
  Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
  
  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
  Set rValo = cRsl.RsVacio("Valores", "Id", "N")
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
  Set rAde = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rBanco = cRsl.RsVacio("CtaBanco", "Id", "N")
      
  RsCv.AddNew
 ' RsCv!Id = 0
  RsCv!Suc = nSucursal
  RsCv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DtpFecha.Value
  RsCv!FechaIva = DtpFecha.Value
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!Sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = TxtCliente.text
  RsCv!Lista = 0
  RsCv!Vendedor = 0
  RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
  RsCv!CondVenta = 0
  RsCv!Neto = lblTotal.Caption
  RsCv!Iva1 = 0
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  RsCv!NoGravados = 0
  Select Case cStock.TraerTipoMovimiento(RsCv!Comprobante)
         Case 15
              RsCv!Debe = lblTotal.Caption
              RsCv!Haber = 0
         Case 10
              RsCv!Debe = 0
              RsCv!Haber = lblTotal.Caption
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = "ADE"
  RsCv!COTIZACION = 1
  RsCv!Anulado = 0
  RsCv.Update

     ' Grabo Recibo de Pago
  rReci.AddNew
 ' rReci!Id = 0
  rReci!fecha = DtpFecha.Value
  rReci!Cliente = TxtCliente.text
  rReci!Venta = RsCv!Movimiento
  rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  rReci!Numero = cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  rReci!Sucursal = cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  Select Case cStock.TraerTipoMovimiento(rReci!Comprobante)
         Case 10
              rReci!Cuota = 0
              rReci!Debe = 0
              rReci!Haber = lblTotal.Caption
         Case 15
              rReci!Cuota = 1000
              rReci!Debe = lblTotal.Caption
              rReci!Haber = 0
  End Select
  rReci!Anulado = 0
  rReci.Update
  
  Rsd.MoveFirst
  Do While Not Rsd.EOF
     If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Rsd!formapago & "'") = "Caja" Then
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
        rCaja!Nombre = LblCliente.Caption
        rCaja!Banco = Rsd!Banco
        rCaja!NumeroCheque = Rsd!NumeroCheque
        rCaja!FechaAcreditacion = IIf(IsNull(Rsd!FechaAcreditacion), Date, Rsd!FechaAcreditacion)
        rCaja!NumeroTarjeta = ""
        rCaja!FechaVencimiento = Null
        rCaja!Autorizacion = ""
        rCaja!Cupon = ""
        If cStock.SumaCaja(rCaja!Comprobante) Then
           rCaja!Debe = Rsd!importe
           rCaja!Haber = 0
        Else
          rCaja!Debe = 0
          rCaja!Haber = Rsd!importe
        End If
        rCaja!Abierta = "S"
        rCaja!Anulado = 0
        rCaja.Update
     Else
     If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Rsd!formapago & "'") <> "No" Then
        rBanco.AddNew
        rBanco!Movimiento = RsCv!Movimiento
        rBanco!Comprobante = RsCv!Comprobante
        rBanco!Numero = RsCv!Numero
        rBanco!Sucursal = RsCv!Sucursal
        rBanco!fecha = DTPfechaValor
        rBanco!Hora = Time
        rBanco!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
        rBanco!formapago = Rsd!formapago
        rBanco!Nombre = LblCliente.Caption
        rBanco!Banco = Rsd!Banco
        rBanco!NumeroCheque = Rsd!NumeroCheque
        rBanco!FechaAcreditacion = IIf(IsNull(Rsd!FechaAcreditacion), Date, Rsd!FechaAcreditacion)
        rBanco!NumeroTarjeta = ""
        rBanco!FechaVencimiento = Null
        rBanco!Autorizacion = ""
        rBanco!Cupon = ""
        If cStock.SumaCaja(rBanco!Comprobante) Then
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
  
          
  cGrab.AgregarAdelantos RsCv, rReci, Rsd, rCaja, rBanco, False
    
 CrearRsDetalles
 Limpiar
 
 ArmaRsParaMovimiento
    
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cStock = Nothing
  
Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "GrabarTodo"


End Sub

Private Sub VerEstado()
  CmdBotones(0).Enabled = False
  If Rsd.RecordCount <> 0 Then
     CmdBotones(0).Enabled = True
  End If
End Sub

Private Sub VerCabecera()
  CmdDetalle(0).Enabled = False
  If Len(LblCliente.Caption) <> 0 Then
     CmdDetalle(0).Enabled = True
  End If
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
     CmbValor.ListIndex = 0
     CmbValor.SetFocus
  Else
     CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
     CmdDetalle(0).Caption = "Agregar"
     Rsd.AddNew
     GrabarRsd
     Rsd.MoveLast
     Grid1.Bookmark = Rsd.Bookmark
     BotonDetalles False, True, True, False
     If Rsd.RecordCount <= 30 Then
        BotonDetalles True, True, True, False
     End If
     sw1 = True
     HabilitarTodo True, True
     HabilitarDetalles False
     CmdBotones(0).Enabled = True
     CmdBotones(7).Enabled = True
     CmdDetalle(0).SetFocus
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub BorrarRsD()
  On Error GoTo errHandler
  If MsgBox("Borra el Producto Seleccionado ?", 20, "Atención") = 6 Then
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
     Else
        Botones False, False, False, False, False, True, False, True, False
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
     GrabarRsd
     BotonDetalles True, False, False, False
     CalcularTotales
     sw1 = True
     HabilitarTodo True, True
     HabilitarDetalles False
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CancelarRsd()
  On Error GoTo errHandler

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
        BotonDetalles True, False, False, False
        HabilitarTodo True, True
        HabilitarDetalles False
        LimpiarDetalles
        If Rsd.EOF = False Or Rsd.BOF = False Then
           BotonDetalles True, True, True, False
           LinkearDetalles
        End If
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
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub GrabarRsd()
  Dim cXa As ClsComprobantesL, cRse As ClsEscritura
  

  Set cXa = New ClsComprobantesL
  Set cRse = New ClsEscritura
'  Rsd!Id = 0

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
     If cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 3 Or cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 12 Then
        Rsd!NumeroCheque = TxtCheque.text
        Rsd!Banco = CmbBanco.ItemData(CmbBanco.ListIndex)
        Rsd!Nombre = TxtLibrador.text
        Rsd!FechaAcreditacion = DTPfechaValor.Value
     Else
        Rsd!Nombre = TxtLibrador.text
     End If
  End If
  If cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 3 Or cXa.TraerTipoCondicion(CmbValor.ItemData(CmbValor.ListIndex)) = 12 Then
     cRse.GrabarValordeUnCampo "CuentasBancarias", "Numero", "Numero+1", "Id=" & CmbBanco.ItemData(CmbBanco.ListIndex)
  End If
  
  Rsd.Update
  VerEstado
  CalcularTotales

End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler
  
  If sw = False And sw1 = True Then
     BotonDetalles True, False, False, False
     If Grid1.Row <> -1 Then
        LimpiarDetalles
        LinkearDetalles
       ' LinkearTexto
        BotonDetalles True, True, True, False
     End If
  Else
    If CmdBotones(0).Caption = "Nuevo" Then
     If Grid1.Row <> -1 Then
        LimpiarDetalles
        LinkearDetalles
     '   LinkearTexto
     End If
     End If
  End If
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
  lblTotal.Caption = Format(lblTotal.Caption, "#0.00")
End Sub

Private Sub Imprimir()
  Dim cCli As ClsClienteL, pY As Single, cRsl As ClsLectura
  Dim cNum As ClsComprobantesL
  
  
  Dim Rp As ADODB.Recordset
  
  Set cCli = New ClsClienteL
  Set cRsl = New ClsLectura
  Set cNum = New ClsComprobantesL
  
  Printer.font.Name = "Ms Sans Seriff"
  Printer.font.Size = 14
  Printer.font.bold = True
  Printer.ScaleMode = 7
  Printer.CurrentY = 1.5
  Printer.CurrentX = 1
  Printer.Print cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
  Printer.font.Size = 14
  Printer.font.bold = False
  Printer.CurrentY = 2.5
  Printer.CurrentX = 1
  Printer.Print "ORDEN DE PAGO"
  Printer.font.Size = 10
  Printer.CurrentY = 3.5
  Printer.CurrentX = 13
  Printer.Print "Nro.de Documento :" & LblNumero.Caption
  Printer.CurrentY = 4
  Printer.CurrentX = 13
  Printer.Print "Fecha de Pago :" & DtpFecha.Value
  Printer.CurrentY = 4.5
  Printer.CurrentX = 1
  Printer.Print "Proveedor :" & Trim(TxtCliente.text) & " - " & LblCliente.Caption
  Printer.CurrentY = 5
  Printer.CurrentX = 1
  Printer.Print "Domicilio :" & cRsl.TraerValorDeUnCampo("Proveedores", "Domicilio", "Proveedor='" & TxtCliente.text & "'")
  Printer.CurrentY = 5.5
  Printer.CurrentX = 1
  Printer.Print "Localidad :"
  Printer.CurrentY = 6
  Printer.CurrentX = 1
  Printer.Print "CUIT :" & cRsl.TraerValorDeUnCampo("Proveedores", "Cuit", "Proveedor='" & TxtCliente.text & "'")
  Printer.CurrentY = 6.5
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  Printer.CurrentY = 7
  Printer.CurrentX = 1
  Printer.Print "       FECHA                                      TIPO                        NUMERO                                                    IMPORTE"
  Printer.CurrentY = 7.5
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  pY = 8
  Printer.CurrentX = 1
'  Printer.Print Format(Rs!Asiento, "00000000")
  Printer.CurrentY = pY
  Printer.CurrentX = 1.7
  Printer.Print DtpFecha.Value
  Printer.CurrentY = pY
  Printer.CurrentX = 5.5
  Printer.Print CmbComprobante.text
  Printer.CurrentY = pY
  Printer.CurrentX = 9
  Printer.Print LblNumero.Caption
  Printer.CurrentY = pY
  Printer.CurrentX = 18 - Printer.TextWidth(Format(lblTotal.Caption, "#0.00"))
  Printer.Print Format(lblTotal.Caption, "#0.00")
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 8
  Printer.Print "T O T A L   A  P A G A R :"
  Printer.CurrentY = pY
  Printer.CurrentX = 18 - Printer.TextWidth(Format(lblTotal.Caption, "#0.00"))
  Printer.Print Format(lblTotal.Caption, "#0.00")
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "CANCELACION :"
  Rsd.MoveFirst
  pY = pY + 1
  
  Do While Not Rsd.EOF
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     Printer.Print Rsd!formapago
     Printer.CurrentY = pY
     Printer.CurrentX = 7
     If Rsd!NumeroCheque <> "" Then
        Printer.Print cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & Rsd!Banco) & " " & Rsd!NumeroCheque & " " & Rsd!FechaAcreditacion
     Else
        If Rsd!Cupon <> "" Then
           Printer.Print cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & Rsd!Banco) & " " & Rsd!NumeroTarjeta & " " & Rsd!Cupon
        End If
     End If
     Printer.CurrentY = pY
     Printer.CurrentX = 18 - Printer.TextWidth(Format(Rsd!importe, "#0.00"))
     Printer.Print Format(Rsd!importe, "#0.00")
     Rsd.MoveNext
     pY = pY + 0.5
  Loop
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 8
  Printer.Print "T O T A L   D E L   P A G O :"
  Printer.CurrentY = pY
  Printer.CurrentX = 18 - Printer.TextWidth(Format(lblTotal.Caption, "#0.00"))
  Printer.Print Format(lblTotal.Caption, "#0.00")
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "CONFECCIONO :"
  Printer.CurrentY = pY
  Printer.CurrentX = 10
  Printer.Print "AUTORIZO :_____________________________________________"
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "Recibimos de " & cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " la cantidad de $ " & Format(lblTotal.Caption, "#0.00")
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "En todo de acuerdo con la presente liquidación."
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "ACLARACION_______________________________  DOC:_____________________ FIRMA_________________________"
  Printer.EndDoc

End Sub

