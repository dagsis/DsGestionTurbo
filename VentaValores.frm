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
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   720
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "Borrar"
         Height          =   720
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3315
         Width           =   945
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   3720
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
      Begin VB.ComboBox CmbFPago 
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Width           =   2730
      End
      Begin MSComCtl2.DTPicker DtpFVencimiento 
         Height          =   315
         Left            =   2685
         TabIndex        =   8
         Top             =   2745
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   158400513
         CurrentDate     =   36795
      End
      Begin MSComCtl2.DTPicker DtpFAcreditacion 
         Height          =   315
         Left            =   6615
         TabIndex        =   6
         Top             =   2160
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   158400513
         CurrentDate     =   36795
      End
      Begin MSDataGridLib.DataGrid GridValores 
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
         Left            =   120
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

Private mRs As ADODB.Recordset
Private mTotal As Double
Public Cancelado As Boolean

Private mLoading As Boolean
Private cRsl As ClsLectura
Private p    As ClsPrograma
Private mTipoByCondVta As Object 'Scripting.Dictionary

Private Sub ActualizarBotonesSegunResta()
    Dim resta As Double
    resta = Round(mTotal - SumaImportes(), 2)

    'Nuevo solo si falta pagar
    CmdNuevo.Enabled = (resta > 0.01)

    'Borrar si hay al menos 1 registro (aunque el puntero esté BOF/EOF)
    If (Not mRs Is Nothing) Then
        CmdBorrar.Enabled = (mRs.State = adStateOpen And mRs.RecordCount > 0)
    Else
        CmdBorrar.Enabled = False
    End If
End Sub


Private Sub FocoSegunEstado()
    On Error Resume Next

    If CmdNuevo.Enabled Then
        If TxtImporte.Enabled And TxtImporte.Visible Then
            TxtImporte.SetFocus
            SeleccionarTodo TxtImporte
        End If
    Else
        'si ya completó el total, que vayan a salir directo
        CmdSalir.SetFocus
    End If

    err.Clear
End Sub

Private Sub CargarCacheCondVenta()
    On Error GoTo errHandler

    Set mTipoByCondVta = CreateObject("Scripting.Dictionary")
    mTipoByCondVta.RemoveAll

    'Traemos TODAS las condiciones que entran en el combo
    Dim rs As ADODB.Recordset
    Set rs = cRsl.TraerRsCondi("CondVenta", "CondVta", "Tipo<5")

    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
                mTipoByCondVta(CLng(rs!CondVta)) = CInt(rs!tipo)
                rs.MoveNext
            Loop
        End If
    End If

    Exit Sub
errHandler:
    MsgBox err.Description, vbCritical, "CargarCacheCondVenta"
End Sub


Private Sub CmbFPago_Click()
    If mLoading Then Exit Sub

    Dim t As Integer
    t = TipoFPagoSeleccionado()

    LimpiarCamposNoUsados t
    AplicarHabilitacionSegunTipo

    SetImporteResta
    RefrescarSumatoria
    FocoImporteSeguro
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim WshShell As Object
  If KeyAscii = vbKeyReturn Then
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

'========================
'  API
'========================
Public Sub Inicializar(ByRef rs As ADODB.Recordset, ByVal totalPagar As Double)
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set p = New ClsPrograma

  Cancelado = True
  Set mRs = rs
  mTotal = Round(totalPagar, 2)
  
  'Enlazar grid
  Set GridValores.DataSource = mRs
  CabGridValores
  
  'Total a pagar en label inferior
  RefrescarSumatoria
  
  'Cargar combos (ya lo tenés hecho con defaults de config)
  CargarCombos
  
  '<<< NUEVO: cache CondVta->Tipo para habilitar campos sin consultar cada click
  CargarCacheCondVenta
  
  'Defaults
  If DtpFAcreditacion.Value = 0 Then DtpFAcreditacion.Value = Date
  If DtpFVencimiento.Value = 0 Then DtpFVencimiento.Value = Date
  
  TxtNombre.Text = ""
  TxtTarjeta.Text = ""
  TxtCheque.Text = ""
  TxtCupon.Text = ""
  TxtAutorizacion.Text = ""
  
  '<<< NUEVO: por defecto el importe = RESTA (si recién abre, es el total)
  RefrescarSumatoria
  SetImporteResta
  AplicarHabilitacionSegunTipo
  ActualizarBotonesSegunResta
  FocoImporteSeguro

  
Exit Sub

errHandler:
    MsgBox err.Description, vbCritical, "Inicializar Valores"

End Sub

Private Sub CabGridValores()
  GridValores.HeadFont.Size = 10
  GridValores.HeadFont.Bold = True
  With GridValores
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

'========================
'  Combos
'========================
Private Sub CargarCombos()
    On Error GoTo errHandler

    '--- Formas de pago (simple, reemplazable por carga desde tabla)
    cRsl.CargaCombo CmbFPago, "CondVenta", "CondVta", "Descripcion", "Tipo<7"
    p.SetComboByItemData CmbFPago, nFormaPago

    '--- Bancos (si lo querés de tabla, lo cambiamos)
     cRsl.CargaCombo CmbBanco, "Bancos", "Banco", "Descripcion", ""
     p.SetComboByItemData CmbBanco, nBanco

    Exit Sub
errHandler:
    MsgBox err.Description, vbCritical, "CargarCombos"
End Sub

'========================
'  Helpers
'========================
Private Function SumaImportes() As Double
    On Error GoTo errHandler

    Dim s As Double
    s = 0

    If mRs Is Nothing Then GoTo fin
    If mRs.State <> adStateOpen Then GoTo fin
    If mRs.RecordCount <= 0 Then GoTo fin
    If (mRs.BOF And mRs.EOF) Then GoTo fin

    Dim rsC As ADODB.Recordset
    Set rsC = mRs.Clone

    rsC.MoveFirst
    Do While Not rsC.EOF
        s = s + CDbl(Val(rsC!importe & ""))
        rsC.MoveNext
    Loop

    rsC.Close
    Set rsC = Nothing

fin:
    SumaImportes = Round(s, 2)
    Exit Function

errHandler:
    On Error Resume Next
    If Not rsC Is Nothing Then
        rsC.Close
        Set rsC = Nothing
    End If
    MsgBox err.Description, vbCritical, "SumaImportes"
End Function


Private Sub RefrescarSumatoria()
    On Error GoTo errHandler

    Dim pagado As Double, resta As Double
    pagado = SumaImportes()
    resta = Round(mTotal - pagado, 2)

    'Usamos tu label existente para mostrar todo:
    LblTPago.Caption = "Total: " & Format(mTotal, "#0.00") & _
                       "  |  Pagado: " & Format(pagado, "#0.00") & _
                       "  |  Resta: " & Format(resta, "#0.00")

    Exit Sub
errHandler:
    MsgBox err.Description, vbCritical, "RefrescarSumatoria"
End Sub

Private Sub SeleccionarTodo(ByVal tb As TextBox)
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
End Sub

Private Sub SetImporteResta()
    Dim pagado As Double, resta As Double
    pagado = SumaImportes()
    resta = Round(mTotal - pagado, 2)
    If resta < 0 Then resta = 0

    TxtImporte.Text = Format(resta, "#0.00")
End Sub

Private Sub FocoImporteSeguro()
    On Error Resume Next
    If TxtImporte.Enabled And TxtImporte.Visible Then
        TxtImporte.SetFocus
        SeleccionarTodo TxtImporte
    End If
    err.Clear
End Sub


Private Function TotalOK() As Boolean
    TotalOK = (Abs(SumaImportes() - mTotal) <= 0.01)
End Function

Private Sub HabilitarCamposPorForma()
    'Opcional: habilitar/inhabilitar campos según forma de pago elegida
    'Para no complicarte ahora, lo dejo simple: todo habilitado.
End Sub

Private Sub Form_Load()

  CmdNuevo.Picture = LoadResPicture("Nuevo", 0)
  CmdBorrar.Picture = LoadResPicture("Borrar", 0)
  
  CmdSalir.Picture = LoadResPicture("Salir", 0)
  
End Sub

'========================
'  Botones (con TUS nombres)
'========================
Private Sub CmdNuevo_Click()
    On Error GoTo errHandler
    If Round(mTotal - SumaImportes(), 2) <= 0.01 Then Exit Sub
    If mLoading Then Exit Sub

    If CmbFPago.ListIndex = -1 Then
        MsgBox "Elegí la forma de pago.", vbExclamation, "Valores"
        Exit Sub
    End If

    Dim impu As Double
    impu = Round(CDbl(Val(TxtImporte.Text)), 2)
   If impu <= 0 Then
        'si está en cero, ponemos la resta automáticamente
        Dim resta As Double
        resta = Round(mTotal - SumaImportes(), 2)
        If resta < 0 Then resta = 0
        TxtImporte.Text = Format(resta, "#0.00")
        impu = Round(CDbl(Val(TxtImporte.Text)), 2)
    End If

    'No permitir pasarse del total (recomendado)
    Dim pagado As Double
    pagado = SumaImportes()
    If Round(pagado + impu, 2) - mTotal > 0.01 Then
        MsgBox "Con este importe te pasás del total. Ajustá el importe.", vbExclamation, "Valores"
        TxtImporte.SetFocus
        Exit Sub
    End If

    mLoading = True

    mRs.AddNew
    mRs!FormaPagoId = CmbFPago.ItemData(CmbFPago.ListIndex)
    mRs!FormaPagoDesc = CmbFPago.Text
    mRs!importe = impu

    mRs!Nombre = TxtNombre.Text
    If CmbBanco.ListIndex >= 0 Then
    mRs!BancoId = CmbBanco.ItemData(CmbBanco.ListIndex)
    Else
        mRs!BancoId = 0
    End If
    mRs!BancoDesc = CmbBanco.Text

    mRs!BancoDesc = CmbBanco.Text
    mRs!NroCheque = TxtCheque.Text
    mRs!FechaAcreditacion = DtpFAcreditacion.Value

    mRs!NroTarjeta = TxtTarjeta.Text
    mRs!Vencimiento = DtpFVencimiento.Value
    mRs!NroCupon = TxtCupon.Text
    mRs!NroAutorizacion = TxtAutorizacion.Text

    mRs!Observaciones = "" 'si después agregás un txt observaciones
   
    mRs.Update
    mLoading = False
    
    'limpio solo campos específicos (NO TxtNombre)
    TxtCheque.Text = ""
    TxtTarjeta.Text = ""
    TxtCupon.Text = ""
    TxtAutorizacion.Text = ""
    
    RefrescarSumatoria
    SetImporteResta
    AplicarHabilitacionSegunTipo
    ActualizarBotonesSegunResta
    FocoSegunEstado
Exit Sub

errHandler:
    mLoading = False
    MsgBox err.Description, vbCritical, "Nuevo"
End Sub

Private Sub CmdBorrar_Click()
    On Error GoTo errHandler
    If mLoading Then Exit Sub

    If mRs Is Nothing Then Exit Sub
    If mRs.RecordCount <= 0 Then Exit Sub
    If (mRs.BOF Or mRs.EOF) Then mRs.MoveFirst
    If (mRs.BOF Or mRs.EOF) Then Exit Sub

    If MsgBox("¿Eliminar el valor seleccionado?", vbYesNo + vbQuestion, "Valores") = vbNo Then Exit Sub

    Dim bkActual As Variant
    Dim bkIr As Variant
    Dim tieneBkIr As Boolean
    Dim rsC As ADODB.Recordset

    bkActual = mRs.Bookmark
    tieneBkIr = False

    '--- calculo a dónde ir después (sin tocar el real)
    Set rsC = mRs.Clone
    rsC.Bookmark = bkActual

    rsC.MoveNext
    If Not rsC.EOF Then
        bkIr = rsC.Bookmark
        tieneBkIr = True
    Else
        rsC.MovePrevious
        If Not rsC.BOF Then
            bkIr = rsC.Bookmark
            tieneBkIr = True
        End If
    End If

    rsC.Close
    Set rsC = Nothing

    '--- borrar el actual
    mLoading = True
    mRs.Bookmark = bkActual
    mRs.Delete            '<<< NO Update
    mLoading = False

    '--- reposicionar si quedan registros
    If mRs.RecordCount > 0 Then
        On Error Resume Next
        If tieneBkIr Then
            mRs.Bookmark = bkIr
        Else
            mRs.MoveFirst
        End If
        If err.Number <> 0 Then
            err.Clear
            mRs.MoveFirst
        End If
        On Error GoTo errHandler
    End If

    RefrescarSumatoria
    SetImporteResta
    ActualizarBotonesSegunResta
    FocoSegunEstado
    Exit Sub

errHandler:
    mLoading = False
    MsgBox err.Description, vbCritical, "Borrar"
End Sub



Private Sub CmdSalir_Click()
    On Error GoTo errHandler

    If Not TotalOK() Then
        MsgBox "Los valores cargados no suman el total a pagar." & vbCrLf & _
               LblTPago.Caption, vbExclamation, "Valores"
        Exit Sub
    End If

    Cancelado = False
    Unload Me
    Exit Sub

errHandler:
    MsgBox err.Description, vbCritical, "Salir"
End Sub



Private Sub TxtImporte_GotFocus()
   SeleccionarTodo TxtImporte
End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
    SoloNumDecimal KeyAscii, TxtImporte
End Sub
Private Sub SoloNumDecimal(ByRef KeyAscii As Integer, ByVal tb As TextBox)
    If SoloNumeroDecimalFinal(KeyAscii, tb) = False Then KeyAscii = 0
End Sub



'========================
'  Eventos de controles
'========================
Private Sub TxtImporte_LostFocus()
    If Trim$(TxtImporte.Text) = "" Then TxtImporte.Text = "0.00"
    TxtImporte.Text = Format(CDbl(Val(TxtImporte.Text)), "#0.00")
End Sub


Private Sub GridValores_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If mLoading Then Exit Sub
    'Si querés, acá podés “cargar” los txt con el registro actual.
End Sub

Private Function TipoFPagoSeleccionado() As Integer
    On Error GoTo fallback

    TipoFPagoSeleccionado = 0
    If CmbFPago.ListIndex < 0 Then Exit Function

    Dim id As Long
    id = CLng(CmbFPago.ItemData(CmbFPago.ListIndex))

    If Not mTipoByCondVta Is Nothing Then
        If mTipoByCondVta.Exists(id) Then
            TipoFPagoSeleccionado = CInt(mTipoByCondVta(id))
            Exit Function
        End If
    End If

fallback:
    'Plan B (por si no está cacheado): 1 consulta puntual (poco frecuente)
    TipoFPagoSeleccionado = CInt(Val(cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & id)))
End Function


Private Sub AplicarHabilitacionSegunTipo()
    Dim t As Integer
    t = TipoFPagoSeleccionado()

    'apagar todo
    TxtNombre.Enabled = False
    CmbBanco.Enabled = False
    TxtCheque.Enabled = False
    DtpFAcreditacion.Enabled = False

    TxtTarjeta.Enabled = False
    DtpFVencimiento.Enabled = False
    TxtCupon.Enabled = False
    TxtAutorizacion.Enabled = False

    'TxtImporte siempre
    TxtImporte.Enabled = True

    Select Case t
        Case 1
            'solo importe

        Case 3, 4
            TxtNombre.Enabled = True
            CmbBanco.Enabled = True
            TxtCheque.Enabled = True
            DtpFAcreditacion.Enabled = True

        Case 2, 5
            TxtTarjeta.Enabled = True
            DtpFVencimiento.Enabled = True
            TxtCupon.Enabled = True
            TxtAutorizacion.Enabled = True
    End Select
End Sub

Private Sub LimpiarCamposNoUsados(ByVal t As Integer)
    Select Case t
        Case 1
            TxtCheque.Text = ""
            TxtTarjeta.Text = ""
            TxtCupon.Text = ""
            TxtAutorizacion.Text = ""

        Case 3, 4
            TxtTarjeta.Text = ""
            TxtCupon.Text = ""
            TxtAutorizacion.Text = ""

        Case 2, 5
            TxtCheque.Text = ""
    End Select
End Sub

