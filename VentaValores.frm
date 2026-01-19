VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
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
      Height          =   4185
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   8010
      Begin VB.CommandButton CmdValores 
         Caption         =   "Cuotas"
         Height          =   720
         Index           =   2
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdValores 
         Caption         =   "Salir"
         Height          =   720
         Index           =   3
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdValores 
         Caption         =   "Borrar"
         Height          =   720
         Index           =   1
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdValores 
         Caption         =   "Nuevo"
         Height          =   720
         Index           =   0
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3315
         Width           =   945
      End
      Begin VB.TextBox TxtAutorizacion 
         Height          =   315
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "TxtAutorizacion"
         Top             =   2745
         Width           =   1905
      End
      Begin VB.TextBox TxtCupon 
         Height          =   315
         Left            =   4125
         MaxLength       =   30
         TabIndex        =   9
         Text            =   "TxtCupon"
         Top             =   2745
         Width           =   1860
      End
      Begin VB.TextBox TxtTarjeta 
         Height          =   315
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "TxtTarjeta"
         Top             =   2745
         Width           =   2550
      End
      Begin VB.TextBox TxtCheque 
         Height          =   315
         Left            =   4935
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "TxtCheque"
         Top             =   2145
         Width           =   1665
      End
      Begin VB.ComboBox CmbBanco 
         Height          =   315
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2145
         Width           =   2175
      End
      Begin VB.TextBox TxtNombre 
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "TxtNombre"
         Top             =   2145
         Width           =   2550
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "TxtImporte"
         Top             =   1200
         Width           =   1065
      End
      Begin VB.ComboBox CmbPago 
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Width           =   2730
      End
      Begin MSComCtl2.DTPicker DTPVencimiento 
         Height          =   315
         Left            =   2685
         TabIndex        =   8
         Top             =   2745
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709951489
         CurrentDate     =   36795
      End
      Begin MSComCtl2.DTPicker DTPAcreditacion 
         Height          =   315
         Left            =   6615
         TabIndex        =   6
         Top             =   2145
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709951489
         CurrentDate     =   36795
      End
      Begin MSDataGridLib.DataGrid Grid2 
         Height          =   1635
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   2884
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7905
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Label LblTPago 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTPago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         TabIndex        =   25
         Top             =   3750
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label29 
         Caption         =   "Nº Cupón :"
         Height          =   195
         Left            =   4140
         TabIndex        =   24
         Top             =   2535
         Width           =   870
      End
      Begin VB.Label Label28 
         Caption         =   "Nº Autorización :"
         Height          =   180
         Left            =   6015
         TabIndex        =   23
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label Label27 
         Caption         =   "F.Vencimiento :"
         Height          =   180
         Left            =   2685
         TabIndex        =   22
         Top             =   2535
         Width           =   1110
      End
      Begin VB.Label Label26 
         Caption         =   "Nº Tarjeta :"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "F.Acreditación :"
         Height          =   240
         Left            =   6630
         TabIndex        =   20
         Top             =   1905
         Width           =   1140
      End
      Begin VB.Label Label23 
         Caption         =   "Nº de Cheque :"
         Height          =   195
         Left            =   4935
         TabIndex        =   19
         Top             =   1905
         Width           =   1200
      End
      Begin VB.Label Label22 
         Caption         =   "Banco :"
         Height          =   210
         Left            =   2685
         TabIndex        =   18
         Top             =   1905
         Width           =   750
      End
      Begin VB.Label LblNombre 
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   1905
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Importe :"
         Height          =   210
         Left            =   150
         TabIndex        =   16
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label18 
         Caption         =   "Forma de Pago :"
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   240
         Width           =   1320
      End
   End
End
Attribute VB_Name = "VentaValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean, cImpor As Currency, cImporte As Currency, sw1 As Boolean
Dim cRsl As ClsLectura, nCh As Integer
Public RsV As ADODB.Recordset

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Public Sub RemoveCancelMenuItem(frm As Form)
  Dim hSysMenu As Long
  hSysMenu = GetSystemMenu(frm.hwnd, 0)
  Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
  Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

Private Sub CrearRsValores()
 On Error GoTo errHandler

 Set RsV = New ADODB.Recordset
 Set cRsl = New ClsLectura
  
 RsV.Fields.Append "Comprobante", adInteger
 RsV.Fields.Append "FormaPago", adVarChar, 30
 RsV.Fields.Append "Importe", adCurrency
 RsV.Fields.Append "Nombre", adVarChar, 60, adFldIsNullable
 RsV.Fields.Append "Banco", adInteger, , adFldIsNullable
 RsV.Fields.Append "NCheque", adVarChar, 30, adFldIsNullable
 RsV.Fields.Append "FAcredi", adDate, , adFldIsNullable
 RsV.Fields.Append "NTarjet", adVarChar, 30, adFldIsNullable
 RsV.Fields.Append "FVencim", adDate, , adFldIsNullable
 RsV.Fields.Append "NCupon", adVarChar, 30, adFldIsNullable
 RsV.Fields.Append "NAutori", adVarChar, 30, adFldIsNullable
 RsV.Fields.Append "CodPago", adInteger, , adFldIsNullable
 RsV.Fields.Append "IdCheque", adInteger, , adFldIsNullable
 RsV.CursorType = adOpenKeyset
 RsV.LockType = adLockOptimistic
 RsV.Open
 
 CabGridValores
 
Exit Sub

errHandler:
   ManejaErrores
 
End Sub

Private Sub CabGridValores()
  Set Grid2.DataSource = RsV
  Grid2.HeadFont.Size = 10
  Grid2.HeadFont.bold = True
  With Grid2
      .Columns(0).Visible = False
      .Columns(1).Width = 3100
      .Columns(1).Caption = "Forma de Pago"
      .Columns(2).Width = 1200
      .Columns(2).Caption = "Importe"
      .Columns(2).NumberFormat = "#0.00"
      .Columns(2).Alignment = dbgRight
      .Columns(3).Visible = False
      .Columns(4).Visible = False
      .Columns(5).Visible = False
      .Columns(6).Visible = False
      .Columns(7).Visible = False
      .Columns(8).Visible = False
      .Columns(9).Visible = False
      .Columns(10).Visible = False
      .Columns(11).Visible = False
      .Columns(12).Visible = False
 End With
 
End Sub

Private Sub LimpiarValores()
  TxtNombre.text = ""
  CmbBanco.ListIndex = -1
  TxtCheque.text = ""
  DTPAcreditacion.Value = Date
  TxtTarjeta.text = ""
  DTPVencimiento.Value = Date
  TxtCupon.text = ""
  TxtAutorizacion.text = ""
  TxtImporte.text = "0.00"
   If nLlama = 10 Then
      TxtNombre.text = VentaRapido.LblCliente.Caption
      TxtTarjeta.text = "S/D"
      TxtCupon.text = "S/D"
      TxtAutorizacion.text = "S/D"
   End If
End Sub


Private Sub LinkearValores()
  Set cRsl = New ClsLectura
  
  TxtImporte.text = Format(RsV!importe, "#0.00")
  TxtNombre.text = RsV!Nombre
  TxtCheque.text = RsV!NCheque
  DTPAcreditacion.Value = IIf(IsNull(RsV!FAcredi), Date, RsV!FAcredi)
  TxtTarjeta.text = RsV!NTarjet
  DTPVencimiento.Value = IIf(IsNull(RsV!FVencim), Date, RsV!FVencim)
  TxtCupon.text = RsV!NCupon
  TxtAutorizacion.text = RsV!NAutori
  CmbPago.text = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", RsV!CodPago)
End Sub

Private Sub CmbBanco_Click()
   If TxtNombre.text = "" Then
      If nLlama = 4 Then
         TxtNombre.text = VentaCobranzas.LblCliente.Caption
      Else
         If nLlama = 10 Then
            TxtNombre.text = VentaRapido.LblCliente.Caption
         Else
           TxtNombre.text = VentaComprobantes.LblCliente.Caption
         End If
      End If
   End If
   TodoBien
End Sub

Private Sub CmbPago_Click()
  Dim cTip As ClsComprobantesL, nCuotas As Integer, cInter As Currency, rAux As ADODB.Recordset
  Dim cRsl As ClsLectura
  Set cRsl = New ClsLectura
  Dim sFinan As Single, sIvaTarjeta As Single, sRgTarjeta As Single, sIbTarjeta As Single
  Dim cFina As Currency
  If sw = False And sw1 = False Then
     Set cTip = New ClsComprobantesL
     If CmbPago.ListIndex <> -1 Then
     Label23.Caption = "Nº de Cheque :"
     Label24.Caption = "F.Acreditación"
     Label26.Caption = "Nº Tarjeta :"

        Select Case cTip.TraerTipoCondicion(CmbPago.ItemData(CmbPago.ListIndex))
               Case 1
                    HabilitarValores True, True, True, False, False, False, False, False, False, False, False, False
                    CmdValores(0).Enabled = True
               Case 2
                   
                    HabilitarValores True, True, True, True, False, False, False, True, True, True, True, False
                    CmdValores(0).Enabled = True
               Case 3
                    HabilitarValores True, True, True, True, True, True, True, False, False, False, False, False
                    CmdValores(0).Enabled = True
               Case 4
                    HabilitarValores True, True, True, True, True, True, True, True, False, False, False, False
                    cRsl.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
                    CmdValores(0).Enabled = True
                     Label26.Caption = "Cuit Titular :"
               Case 5
                    HabilitarValores True, True, True, True, False, False, False, False, False, False, False, False
                    CmdValores(0).Enabled = True
               Case 6
                    HabilitarValores True, True, True, False, False, False, False, False, False, False, False, False
                    CmdValores(0).Enabled = True
               Case 7
                    sw1 = True
                    HabilitarValores False, False, False, False, False, False, False, False, False, False, False, True
                    BotonValor False, False, True, True
               Case 8
                    HabilitarValores True, True, True, True, True, True, True, True, False, False, False, False
                    cRsl.CargaCombo CmbBanco, "CuentasBancarias", "Id", "Descripcion", ""
                    CmdValores(0).Enabled = True
                    Label23.Caption = "Nº de Inter.:"
               Case 9, 10, 11, 12, 13
                    HabilitarValores True, True, True, True, False, True, True, True, False, False, False, False
                    BotonValor True, False, False, False
                    Label23.Caption = "Nº de Comp.:"
                    Label26.Caption = "Cuit : "
               Case 14
                    HabilitarValores True, True, True, True, False, False, True, False, False, False, False, False
                    BotonValor True, False, False, False
        End Select
        LimpiarValores
        If cTip.TraerTipoCondicion(CmbPago.ItemData(CmbPago.ListIndex)) < 6 Then
           TxtImporte.text = Format(LblTPago.Caption, "#0.00")
           cImporte = CCur(TxtImporte.text)
        Else
           nCuotas = 1
           Set rAux = New ADODB.Recordset
           Set rAux = cRsl.TraerRS("CondVentaTraerUno", CmbPago.ItemData(CmbPago.ListIndex), True)
           sIvaTarjeta = LblTPago * rAux!IvaTarjeta / 100
           sRgTarjeta = LblTPago * rAux!RgTarjeta / 100
           sIbTarjeta = LblTPago * rAux!IBTarjeta / 100
           sFinan = (LblTPago + cInter) * (rAux!Cuotas * rAux!Financiacion) / 100
           nCuotas = cTip.CantCuotas(CmbPago.ItemData(CmbPago.ListIndex))
           If nCuotas >= 1 Then
              TxtImporte.text = (CCur(LblTPago.Caption) + sFinan) / nCuotas
              TxtImporte.text = Format(TxtImporte.text, "0.00")
           End If
        End If
        If TxtImporte.Enabled = True Then
    '       TxtImporte.SetFocus
        End If
        Set cTip = Nothing
        Set cRsl = Nothing
     End If
     TodoBien
  End If
End Sub

Private Sub CmdValores_Click(Index As Integer)
  Select Case Index
         Case 0
              NuevoRsV
         Case 1
              BorrarRsV
         Case 2
              CuotasRsv
         Case 3
              SalirRsv
  End Select
End Sub
Private Sub DTPAcreditacion_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim WshShell As Object
   If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub DTPVencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim WshShell As Object
   If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  If nLlama = 10 Then
     cRsl.CargaCombo CmbPago, "CondVenta", "CondVta", "Descripcion", "tipo=2 or Descripcion='EFECTIVO'"
  Else
     cRsl.CargaCombo CmbPago, "CondVenta", "CondVta", "Descripcion", "tipo<6 and Tipo<>3 or Tipo>=8"
  End If
  CmbPago.text = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub BotonValor(bAgr As Boolean, bBorr As Boolean, bCan As Boolean, bCuot As Boolean)
  CmdValores(0).Enabled = bAgr
  CmdValores(1).Enabled = bBorr
  CmdValores(2).Enabled = bCuot
  CmdValores(3).Enabled = bCan
End Sub

Private Sub HabilitarValores(bCompr As Boolean, bPago As Boolean, bImpor As Boolean, _
  bNomb As Boolean, bBanco As Boolean, bNcheque As Boolean, bFAcre As Boolean, bNtarje As Boolean, _
  bFVenc As Boolean, bNcupon As Boolean, bNAuto As Boolean, bGrid As Boolean)
  CmbPago.Enabled = bCompr
  TxtImporte.Enabled = bImpor
  TxtNombre.Enabled = bNomb
  CmbBanco.Enabled = bBanco
  TxtCheque.Enabled = bNcheque
  DTPAcreditacion.Enabled = bFAcre
  TxtTarjeta.Enabled = bNtarje
  DTPVencimiento.Enabled = bFVenc
  TxtCupon.Enabled = bNcupon
  TxtAutorizacion.Enabled = bNAuto
  Grid2.Enabled = bGrid
End Sub
Private Sub Form_Load()

  Dim cValor As ClsComprobantesL, rAux As ADODB.Recordset, nCasVal
  Dim cRsl As ClsLectura
  
  Set cValor = New ClsComprobantesL
  Set rAux = New ADODB.Recordset
  Set cRsl = New ClsLectura
  
  On Error GoTo errHandler
  
  CmdValores(0).Picture = LoadResPicture("Nuevo", 0)
  CmdValores(1).Picture = LoadResPicture("Borrar", 0)
  CmdValores(2).Picture = LoadResPicture("Valores", 0)
  CmdValores(3).Picture = LoadResPicture("Salir", 0)
  sw = True
  
  cImpor = 0

  LimpiarValores
  CargarCombos
  Set VentaDefinirCuotas.rAReci = Nothing
  If nLlama = 1 Or nLlama = 10 Then
     nCasVal = 1 'VentaComprobantes.nCas
  Else
     nCasVal = VentaCobranzas.nCas
  End If
  Select Case nCasVal
         Case 0                        '        Ya Grabado en la Tabla Valores
              Set rAux = cValor.DameValores(VentaComprobantes.nMovi)
              Set Grid2.DataSource = rAux
              Grid2.HeadFont.Size = 10
              Grid2.HeadFont.bold = True
              With Grid2
                   .Columns(0).Width = 3100
                   .Columns(0).Caption = "Forma de Pago"
                   .Columns(1).Width = 1200
                   .Columns(1).Caption = "Importe"
                   .Columns(1).NumberFormat = "#0.00"
                   .Columns(1).Alignment = dbgRight
              End With
              BotonValor False, False, True, False
              HabilitarValores True, False, False, False, False, False, False, False, False, False, False, False
              Set RsV = New ADODB.Recordset
         Case 1         '        Nuevo Recordset Virtual
              BotonValor True, False, True, False
              HabilitarValores True, False, True, False, False, False, False, False, False, False, False, False
              Set RsV = New ADODB.Recordset
              CrearRsValores
              If nLlama = 1 Then
                 LblTPago.Caption = VentaComprobantes.lblTotal.Caption
              Else
                 If nLlama = 10 Then
                    LblTPago.Caption = VentaRapido.lblTotal.Caption
                 Else
                    LblTPago.Caption = VentaCobranzas.TxtEfectivo.text + CDbl(VentaCobranzas.cSaldo)
                 End If
              End If
              LblTPago.Caption = Format(LblTPago.Caption, "#0.00")
              NuevoRsV
         Case 2                        '        Recorset aun no grabado en la tabla Valores
              If RsV.BOF = False Or RsV.EOF = False Then
                 CabGridValores
                 RsV.MoveFirst
                 LinkearValores
                 BotonValor False, True, True, False
                 Grid2.Enabled = True
                 LblTPago.Caption = "0.00"
                 TxtImporte.text = "0.00"
                 BotonValor False, True, True, False
                 HabilitarValores True, False, False, False, False, False, False, False, False, False, False, False
              End If
        Case 3
              BotonValor True, False, True, False
              HabilitarValores True, False, False, False, False, False, False, False, False, False, False, False
              Set RsV = New ADODB.Recordset
              CrearRsValores
              If nLlama = 1 Then
                 LblTPago.Caption = VentaComprobantes.lblTotal.Caption
              Else
                 LblTPago.Caption = VentaCobranzas.TxtEfectivo.text + CDbl(VentaCobranzas.cSaldo)
              End If
              LblTPago.Caption = Format(LblTPago.Caption, "#0.00")
  End Select
  RemoveCancelMenuItem Me
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub NuevoRsV()
  Dim cTip As ClsComprobantesL, nCuota As Integer
  Set cTip = New ClsComprobantesL
  If sw = True Then
     CmdValores(0).Picture = LoadResPicture("Grabar", 0)
     CmdValores(3).Picture = LoadResPicture("Cancelar", 0)
     CmdValores(0).Caption = "Grabar"
     CmdValores(3).Caption = "Cancelar"
     BotonValor False, False, True, False
     CmbPago.Enabled = True
     sw = False
     CmbPago_Click
  Else
     CmdValores(0).Picture = LoadResPicture("Nuevo", 0)
     CmdValores(3).Picture = LoadResPicture("Salir", 0)
     CmdValores(0).Caption = "Nuevo"
     CmdValores(3).Caption = "Salir"
     BotonValor True, True, False, False
     CmbPago.Enabled = False
     RsV.AddNew
     Grabar
     sw = True
     HabilitarValores False, False, False, False, False, False, False, False, False, False, False, False
     cImpor = TxtImporte.text
     TxtImporte.text = Format(cImpor, "#0.00")
     nCuota = cTip.CantCuotas(CmbPago.ItemData(CmbPago.ListIndex))
     LblTPago.Caption = LblTPago.Caption - (RsV!importe * nCuota)
     LblTPago.Caption = Format(LblTPago.Caption, "#0.00")
     If LblTPago.Caption <= 0 Then
        CmdValores(0).Enabled = False
        CmdValores(3).Enabled = True
     End If
     Grid2.Enabled = True
  End If
End Sub

Private Sub Grid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler
  If sw = True Then
     If Grid2.Row <> -1 Then
        RsV.Bookmark = Grid2.Bookmark
        LinkearValores
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub TxtAutorizacion_Change()
  TodoBien
End Sub

Private Sub TxtCheque_Change()
  TodoBien
End Sub

Private Sub TxtCupon_Change()
  TodoBien
End Sub

Private Sub TxtImporte_Change()
  TodoBien
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
  Dim cTip As ClsComprobantesL
  Set cTip = New ClsComprobantesL
  If TxtImporte.text = "" Then TxtImporte.text = 0
  If cTip.TraerTipoCondicion(CmbPago.ItemData(CmbPago.ListIndex)) <> 6 Then
     If CCur(TxtImporte.text) > cImporte Then TxtImporte.text = cImporte
  End If
  TxtImporte.text = Format(TxtImporte.text, "#0.00")
  If CmdValores(0).Enabled = True Then
     CmdValores(0).SetFocus
  End If
End Sub
Private Sub Grabar()
  Dim cTip As ClsComprobantesL, nCuotas As Integer, cInter As Currency, rAux As ADODB.Recordset
  Dim sFinan As Single, sIvaTarjeta As Single, sRgTarjeta As Single, sIbTarjeta As Single
  
  
  Set cTip = New ClsComprobantesL
  Set rAux = New ADODB.Recordset
  RsV!formapago = CmbPago.text
  nCuotas = 1
  
  Set rAux = cRsl.TraerRS("CondVentaTraerUno", CmbPago.ItemData(CmbPago.ListIndex), True)
     
  sIvaTarjeta = LblTPago * rAux!IvaTarjeta / 100
  sRgTarjeta = LblTPago * rAux!RgTarjeta / 100
  sIbTarjeta = LblTPago * rAux!IBTarjeta / 100
  cInter = Format(sRgTarjeta + sIbTarjeta + sIvaTarjeta, "#0.00")
  Dim sCam As Currency
  sCam = (TxtImporte.text * rAux!Cuotas) - LblTPago.Caption
  sFinan = (LblTPago + cInter) * (rAux!Cuotas * rAux!Financiacion) / 100
  RsV!importe = CCur(TxtImporte.text) + sFinan
  RsV!Nombre = Left(TxtNombre.text, 30)
  RsV!Banco = Null
  RsV!IdCheque = 0
  If CmbBanco.ListIndex <> -1 Then
     RsV!Banco = CmbBanco.ItemData(CmbBanco.ListIndex)
     RsV!IdCheque = 1
  End If
  RsV!NCheque = TxtCheque.text
  RsV!FAcredi = IIf(DTPAcreditacion.Enabled = True, DTPAcreditacion.Value, Null)
  RsV!NTarjet = TxtTarjeta.text
  RsV!FVencim = IIf(DTPVencimiento.Enabled = True, DTPVencimiento.Value, Null)
  RsV!NCupon = TxtCupon.text
  RsV!NAutori = TxtAutorizacion.text
  RsV!CodPago = CmbPago.ItemData(CmbPago.ListIndex)
  RsV.Update
End Sub
Private Sub SalirRsv()
  If sw = False Then
     If CmdValores(0).Caption = "Grabar" Or CmdValores(2).Caption = "Cuotas" Then
        If CmdValores(0).Enabled = True Then
           If MsgBox("El Valor no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           Else
              CmdValores(0).Picture = LoadResPicture("Nuevo", 0)
              CmdValores(3).Picture = LoadResPicture("Salir", 0)
              CmdValores(0).Caption = "Nuevo"
              CmdValores(3).Caption = "Salir"
              BotonValor True, False, True, False
              If RsV.RecordCount <> 0 Then
                 BotonValor True, True, True, False
              Else
                 bFormaNueva = False
                 Unload Me
                ' nLlama = 1
              End If
              sw = True
              Exit Sub
           End If
        End If
        CmdValores(0).Picture = LoadResPicture("Nuevo", 0)
        CmdValores(2).Picture = LoadResPicture("Valores", 0)
        CmdValores(3).Picture = LoadResPicture("Salir", 0)
        CmdValores(0).Caption = "Nuevo"
        CmdValores(2).Caption = "Cuotas"
        CmdValores(3).Caption = "Salir"
        BotonValor True, False, True, False
        If RsV.RecordCount <> 0 Then
           BotonValor True, True, True, False
        Else
          ' nLlama = 1
          sw = False
          sw1 = False
        End If
'        sw = True
'        sw1 = True
     Else
        CmdValores(2).Picture = LoadResPicture("Modificar", 0)
        CmdValores(3).Picture = LoadResPicture("Salir", 0)
        CmdValores(2).Caption = "Modificar"
        CmdValores(3).Caption = "Salir"
     End If
  '   sw = False
'     HabilitarValores True, False, True, False, False, False, False, False, False, False, False, False
'     Form_Load
  Else
  
        If CmdValores(0).Enabled = False Or RsV.RecordCount = 0 Then
           If LblTPago.Caption = 0 Or RsV.RecordCount = 0 Then
              Unload Me
           End If
        Else
           MsgBox "Debe Ingresar el Pago para Poder Salir", vbCritical, "Atención"
        End If
   
  End If
End Sub

Private Sub BorrarRsV()
  Dim cTip As ClsComprobantesL
  Set cTip = New ClsComprobantesL
  
  If MsgBox("Borra el Registro Seleccionado ?", 20, "Atención") = 6 Then
     LblTPago.Caption = LblTPago.Caption + (RsV!importe * cTip.CantCuotas(CmbPago.ItemData(CmbPago.ListIndex)))
     LblTPago.Caption = Format(LblTPago.Caption, "#0.00")
     RsV.Delete
     RsV.MovePrevious
     Set VentaDefinirCuotas.rAReci = Nothing
     If RsV.BOF Then
        RsV.MoveFirst
     End If
     CmbPago.ListIndex = -1
     LimpiarValores
     BotonValor True, False, True, False
     Grid2.Enabled = False
     If RsV.RecordCount <> 0 Then
        LinkearValores
        Grid2.Enabled = True
        BotonValor True, True, False, False
     Else
        If nLlama = 1 Then
           VentaComprobantes.nCas = 3
           VentaComprobantes.LblFinanciacion = Format(0, "#0.00")
           If RsV.RecordCount <> 0 Then
              VentaComprobantes.CalcularTotales
           End If
        End If
     End If
     CabGridValores
  End If
End Sub

Private Sub CuotasRsv()
     VentaDefinirCuotas.Show 1
     BotonValor True, False, True, True
     If VentaDefinirCuotas.rAReci.RecordCount = 0 Then
        BotonValor False, False, True, True
     End If
End Sub

Private Sub TodoBien()
  Dim cTip As ClsComprobantesL
  If sw = False And sw1 = False Then
     CmdValores(0).Enabled = False
     Set cTip = New ClsComprobantesL
     If CmbPago.ListIndex <> -1 Then
        Select Case cTip.TraerTipoCondicion(CmbPago.ItemData(CmbPago.ListIndex))
               Case 1
                     If Len(TxtImporte.text) <> 0 And Val(TxtImporte.text) <> 0 Then
                        CmdValores(0).Enabled = True
                     End If
               Case 2
                     If Len(TxtImporte.text) <> 0 And Len(TxtTarjeta.text) <> 0 And Len(TxtCupon.text) <> 0 And Len(TxtAutorizacion.text) <> 0 And Val(TxtImporte.text) <> 0 Then
                        CmdValores(0).Enabled = True
                     End If
               Case 3
                     If Len(TxtImporte.text) <> 0 And Len(CmbBanco.text) <> 0 And Len(TxtCheque.text) <> 0 And Len(TxtNombre.text) <> 0 And Val(TxtImporte.text) <> 0 Then
                        CmdValores(0).Enabled = True
                     End If
               Case 4, 8
                     If Len(TxtImporte.text) <> 0 And Len(CmbBanco.text) <> 0 And Len(TxtCheque.text) <> 0 And Len(TxtNombre.text) <> 0 And Val(TxtImporte.text) <> 0 Then
                        CmdValores(0).Enabled = True
                     End If
               Case 5
               Case 9 To 13
                     If Len(TxtImporte.text) <> 0 And Len(TxtCheque.text) <> 0 And Val(TxtImporte.text) <> 0 And Len(TxtTarjeta.text) <> 0 Then
                        CmdValores(0).Enabled = True
                     End If
               Case 14
                     If Len(TxtImporte.text) <> 0 And Len(TxtNombre.text) <> 0 Then
                        CmdValores(0).Enabled = True
                     End If
        End Select
        CmdValores(3).Enabled = True
     End If
  End If
  Set cTip = Nothing
  
End Sub

Private Sub TxtNombre_Change()
    TodoBien
End Sub

Private Sub TxtTarjeta_Change()
  TodoBien
End Sub
