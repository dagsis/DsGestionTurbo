VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscarComprobante 
   Caption         =   "Listar Comprobantes"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdNext 
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton CmdPrev 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox TxtNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtSucursal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4500
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox CmbComprobante 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin MSDataGridLib.DataGrid GridBus 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10186
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
   Begin VB.Label Label3 
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "P.Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FrmBuscarComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TipoOperacionFiltro As Long  '1=Ventas, 5=NC ventas, etc (puede ser uno o varios)

'Si querés varios tipos, mejor usar string:
Public TipoOperacionWhere As String  'ej: "1,5" o "2,8"


Private cRsl As ClsLectura
Private p As ClsPrograma
Private mRs As ADODB.Recordset

Public MovimientoElegido As Long

Public OwnerForm As Object   'o As FrmVentas si querés fuerte tipado
Private mPage As Long
Private Const PAGE_SIZE As Long = 23

Private Sub SoloNumEntero(ByRef KeyAscii As Integer)
    If SoloNumero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub SeleccionarTodo(ByVal tb As TextBox)
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
End Sub

Private Sub CmdNext_Click()
   If Not mRs Is Nothing Then
        If mRs.RecordCount >= 0 Then
            If mRs.RecordCount < GetPageSize() Then
                Exit Sub
            End If
        End If
    End If

    mPage = mPage + 1
    Buscar
End Sub

Private Sub CmdPrev_Click()
   If mPage > 1 Then mPage = mPage - 1
    Buscar
End Sub

Private Function GetPageSize() As Long
    Dim ps As Long
    ps = CLng(Val(PAGE_SIZE))
    If ps <= 0 Then ps = 50
    GetPageSize = ps
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim WshShell As Object
  If KeyAscii = vbKeyReturn Then
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
    Set cRsl = New ClsLectura
    Set p = New ClsPrograma

    'IMPORTANTE: que arranque vacío, no "0"
    txtSucursal.Text = ""
    TxtNumero.Text = ""
    CargarCombos
    
    mPage = 1
    Buscar
End Sub

Private Sub TxtNumero_GotFocus()
    SeleccionarTodo TxtNumero
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
    SoloNumEntero KeyAscii
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
    SoloNumEntero KeyAscii
End Sub

Private Sub TxtSucursal_GotFocus()
    SeleccionarTodo txtSucursal
End Sub


Private Sub GridBus_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    PreviewSeleccion
End Sub

Private Sub GridBus_Click()
    PreviewSeleccion
End Sub

Private Sub GridBus_DblClick()
    AceptarSeleccion
End Sub

Private Sub PreviewSeleccion()
    On Error GoTo errHandler

    If mRs Is Nothing Then Exit Sub
    If mRs.BOF Or mRs.EOF Then Exit Sub
    If OwnerForm Is Nothing Then Exit Sub

    OwnerForm.CargarComprobantePorMovimiento CLng(mRs!Movimiento)
    Exit Sub

errHandler:
    MsgBox "PreviewSeleccion: " & err.Description, vbCritical
End Sub


Private Sub AceptarSeleccion()
    On Error Resume Next
    If mRs Is Nothing Then Exit Sub
    If mRs.BOF Or mRs.EOF Then Exit Sub

    MovimientoElegido = CLng(mRs!Movimiento)

    If Not OwnerForm Is Nothing Then
        OwnerForm.CargarComprobantePorMovimiento MovimientoElegido
    End If

    Unload Me
End Sub


Private Sub CargarCombos()
    CmbComprobante.Clear
   
    'Luego los comprobantes
   If ListaTieneValor(TipoOperacionWhere, 1) Then
     cRsl.CargaCombo CmbComprobante, "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Usuario=" & nUsuario & " AND Ver=1"
   Else
     cRsl.CargaComboIn CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion", TipoOperacionWhere
   End If
  

    CmbComprobante.AddItem "TODOS"
    CmbComprobante.ItemData(CmbComprobante.NewIndex) = 0

    CmbComprobante.Text = "TODOS"
End Sub

Private Function ListaTieneValor(ByVal lista As String, ByVal valor As Long) As Boolean
    Dim a() As String, i As Long
    lista = Replace$(lista, " ", "")
    a = Split(lista, ",")
    For i = LBound(a) To UBound(a)
        If CLng(Val(a(i))) = valor Then
            ListaTieneValor = True
            Exit Function
        End If
    Next
End Function

Private Sub CmdBuscar_Click()
    Buscar
End Sub


Private Sub Buscar()
    On Error GoTo errHandler

    Dim sSql As String
    Dim whereSql As String
    Dim idComp As Long
    Dim nSuc As Long
    Dim nNum As Long
    Dim page As Long
    Dim pageSize As Long

    '---------------------------
    ' Paginación
    '---------------------------
    page = mPage
    If page <= 0 Then page = 1

    pageSize = PAGE_SIZE
    If Len(Trim$(pageSize)) > 0 Then
        pageSize = CLng(Val(pageSize))
    End If
    If pageSize <= 0 Then pageSize = 50

    '---------------------------
    ' WHERE (una sola vez)
    '---------------------------
    whereSql = " WHERE 1=1 " & vbCrLf

    'TipoOperacionWhere: "1,5" o "2,8"
    If Len(Trim$(TipoOperacionWhere)) > 0 Then
        whereSql = whereSql & " AND comp.TipoOperacion IN (" & TipoOperacionWhere & ") " & vbCrLf
    End If

    'Comprobante (si no es TODOS)
    If CmbComprobante.ListIndex >= 0 Then
        idComp = CLng(CmbComprobante.ItemData(CmbComprobante.ListIndex))
        If idComp <> 0 Then
            whereSql = whereSql & " AND c.Comprobante=" & idComp & vbCrLf
        End If
    End If

    'Sucursal (ignorar 0 / vacío)
    nSuc = CLng(Val(Trim$(txtSucursal.Text)))
    If nSuc > 0 Then
        whereSql = whereSql & " AND c.Sucursal=" & nSuc & vbCrLf
    End If

    'Número (ignorar 0 / vacío)
    nNum = CLng(Val(Trim$(TxtNumero.Text)))
    If nNum > 0 Then
        whereSql = whereSql & " AND c.Numero=" & nNum & vbCrLf
    End If

    '---------------------------
    ' SQL SQL2008 paginado (ROW_NUMBER)
    '---------------------------
    sSql = ""
    sSql = sSql & "DECLARE @Page INT = " & page & ";" & vbCrLf
    sSql = sSql & "DECLARE @PageSize INT = " & pageSize & ";" & vbCrLf

    sSql = sSql & "WITH Q AS (" & vbCrLf
    sSql = sSql & " SELECT " & vbCrLf
    sSql = sSql & "  c.Movimiento, " & vbCrLf
    sSql = sSql & "  c.Fecha, " & vbCrLf
    sSql = sSql & "  (comp.Descripcion + ' ' + " & vbCrLf
    sSql = sSql & "   RIGHT('0000'+CAST(c.Sucursal AS varchar(4)),4) + '-' + " & vbCrLf
    sSql = sSql & "   RIGHT('00000000'+CAST(c.Numero AS varchar(8)),8)) AS CompCol, " & vbCrLf
    sSql = sSql & "  c.Cliente, cli.RazonSocial, " & vbCrLf
    sSql = sSql & "  CASE " & vbCrLf
    sSql = sSql & "   WHEN ISNULL(c.Debe,0) <> 0 THEN c.Debe " & vbCrLf
    sSql = sSql & "   WHEN ISNULL(c.Haber,0) <> 0 THEN c.Haber " & vbCrLf
    sSql = sSql & "   ELSE ISNULL(c.Neto,0) + ISNULL(c.Iva1,0) + ISNULL(c.NoGravados,0) + ISNULL(c.PIvaCompras,0) " & vbCrLf
    sSql = sSql & "  END AS Total, " & vbCrLf
    sSql = sSql & "  comp.Iva AS EsFiscal, c.Anulado, " & vbCrLf
    sSql = sSql & "  CASE " & vbCrLf
    sSql = sSql & "   WHEN ISNULL(c.Anulado,0)<>0 THEN 'ANU' " & vbCrLf
    sSql = sSql & "   WHEN ISNULL(comp.Iva,0)=1 AND EXISTS(SELECT 1 FROM AfipCaes a WHERE a.Movimiento=c.Movimiento) THEN 'A' " & vbCrLf
    sSql = sSql & "   WHEN ISNULL(comp.Iva,0)=1 THEN 'N' " & vbCrLf
    sSql = sSql & "   ELSE 'C' " & vbCrLf
    sSql = sSql & "  END AS Estado, " & vbCrLf
    sSql = sSql & "  ROW_NUMBER() OVER (ORDER BY c.Fecha DESC, c.Movimiento DESC) AS RN " & vbCrLf
    sSql = sSql & " FROM CabComprobantes c " & vbCrLf
    sSql = sSql & " LEFT JOIN Clientes cli ON cli.Cliente=c.Cliente " & vbCrLf
    sSql = sSql & " LEFT JOIN Comprobantes comp ON comp.Id=c.Comprobante " & vbCrLf
    sSql = sSql & whereSql
    sSql = sSql & ") " & vbCrLf

    sSql = sSql & "SELECT Movimiento, Fecha, CompCol, Cliente, RazonSocial, Total, EsFiscal, Anulado, Estado " & vbCrLf
    sSql = sSql & "FROM Q " & vbCrLf
    sSql = sSql & "WHERE RN BETWEEN ((@Page-1)*@PageSize + 1) AND (@Page*@PageSize) " & vbCrLf
    sSql = sSql & "ORDER BY RN " & vbCrLf

    '---------------------------
    ' Ejecutar
    '---------------------------
    Set mRs = cRsl.TraerRsSQL(sSql)

    '--- Si no hay registros y no es la primer página, retroceder automáticamente ---
    If mRs Is Nothing Then Exit Sub
    
    If mRs.RecordCount = 0 Then
        If mPage > 1 Then
            mPage = mPage - 1
            Buscar
            Exit Sub
        End If
    End If

    '--- Habilitar / deshabilitar navegación ---
    CmdPrev.Enabled = (mPage > 1)
    CmdNext.Enabled = (mRs.RecordCount >= GetPageSize())


    Set GridBus.DataSource = mRs
    GridBus.Refresh
    FormatearGrid
    Exit Sub

errHandler:
    MsgBox err.Description, vbCritical, "Buscar"
End Sub

Private Sub FormatearGrid()
    On Error Resume Next

    GridBus.HeadFont.Bold = True

    '0 Movimiento
    GridBus.Columns(0).Visible = False

    '1 Fecha
    GridBus.Columns(1).Caption = "Fecha"
    GridBus.Columns(1).Width = 1100

    '2 CompCol  <<< ACA
    GridBus.Columns(2).Caption = "Comprobante"
    GridBus.Columns(2).Width = 3000

    '3 Cliente
    GridBus.Columns(3).Caption = "Cliente"
    GridBus.Columns(3).Width = 1000

    '4 RazonSocial
    GridBus.Columns(4).Caption = "Nombre"
    GridBus.Columns(4).Width = 3000

    '5 Total
    GridBus.Columns(5).Caption = "Total"
    GridBus.Columns(5).Width = 1200
    GridBus.Columns(5).Alignment = dbgRight
    GridBus.Columns(5).NumberFormat = "#0.00"

    '6 EsFiscal
    GridBus.Columns(6).Visible = False

    '7 Anulado
    GridBus.Columns(7).Visible = False

    '8 Estado
    GridBus.Columns(8).Caption = "Estado"
    GridBus.Columns(8).Width = 700

    GridBus.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not mRs Is Nothing Then
        If mRs.State = adStateOpen Then mRs.Close
    End If
    Set mRs = Nothing
End Sub


