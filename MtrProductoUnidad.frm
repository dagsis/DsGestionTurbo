VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MtrProductoUnidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidad de Medida de Los Productos"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   9750
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   3
      Left            =   8475
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3450
      Width           =   1110
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2550
      Width           =   1110
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1650
      Width           =   1110
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Agregar"
      Height          =   810
      Index           =   0
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   750
      Width           =   1110
   End
   Begin VB.TextBox TxtCantidadB 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5700
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "TxtCant"
      Top             =   4635
      Width           =   825
   End
   Begin VB.ComboBox CmbUnidadB 
      Height          =   315
      Left            =   7245
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4635
      Width           =   2130
   End
   Begin VB.TextBox TxtCantidadA 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   990
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "TxtCant"
      Top             =   4635
      Width           =   825
   End
   Begin VB.ComboBox CmbUnidadA 
      Height          =   315
      Left            =   2535
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4635
      Width           =   2130
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   9
      Top             =   225
      Width           =   720
   End
   Begin VB.TextBox TxtProducto 
      Height          =   315
      Left            =   990
      MaxLength       =   25
      TabIndex        =   7
      Text            =   "TxtProducto"
      Top             =   225
      Width           =   930
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   18
      Top             =   765
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   6165
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
   Begin VB.Line Line4 
      X1              =   8520
      X2              =   9480
      Y1              =   4425
      Y2              =   4425
   End
   Begin VB.Line Line3 
      X1              =   4905
      X2              =   6000
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   3930
      X2              =   4740
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      X1              =   1335
      X2              =   195
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label8 
      Caption         =   "Unidad de Medida de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6075
      TabIndex        =   15
      Top             =   4335
      Width           =   3075
   End
   Begin VB.Label Label7 
      Caption         =   "Unidad de Medida Principal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1485
      TabIndex        =   14
      Top             =   4335
      Width           =   2355
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad :"
      Height          =   240
      Left            =   4890
      TabIndex        =   13
      Top             =   4635
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "Unidad :"
      Height          =   240
      Left            =   6600
      TabIndex        =   12
      Top             =   4635
      Width           =   705
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad :"
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   4635
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Unidad :"
      Height          =   240
      Left            =   1890
      TabIndex        =   10
      Top             =   4635
      Width           =   705
   End
   Begin VB.Label LblProducto 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblProducto"
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
      Left            =   2025
      TabIndex        =   8
      Top             =   240
      Width           =   5595
   End
   Begin VB.Label Label1 
      Caption         =   "Producto :"
      Height          =   225
      Left            =   195
      TabIndex        =   6
      Top             =   255
      Width           =   915
   End
End
Attribute VB_Name = "MtrProductoUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rsd As ADODB.Recordset, sw As Boolean, nId As Integer

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Anular
         Case 2
              Modificar
         Case 3
              Salir
  End Select
End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bSal As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  CmdBotones(2).Enabled = bMod
  CmdBotones(3).Enabled = bSal
End Sub

Private Sub Nuevo()
  On Error GoTo errHandler

  If sw = True Then
     Habilitar True, False
     Botones True, False, False, True
     HabilitarCab False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(3).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(3).Caption = "Cancelar"
     sw = False
     CmbUnidadA.ListIndex = 0
     CmbUnidadB.ListIndex = 0
     TxtCantidadA.SetFocus
  Else
     GrabarTodo
     sw = True
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(3).Picture = LoadResPicture("Salir", 0)
     CmdBotones(3).Caption = "Salir"
     Habilitar False, True
     Botones True, True, True, True
     HabilitarCab True
  End If
Exit Sub

errHandler:
     ManejaErrores

End Sub

Private Sub GrabarTodo()
  Dim cRsl As ClsLectura, sSql As String
  
  Set cRsl = New ClsLectura
  If CmdBotones(0).Caption = "Grabar" Then
     Rsd.AddNew
     Rsd!CantidadA = TxtCantidadA.text
     Rsd!UnidadA = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & CmbUnidadA.ItemData(CmbUnidadA.ListIndex))
     Rsd!CantidadB = TxtCantidadB.text
     Rsd!Unidadb = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & CmbUnidadB.ItemData(CmbUnidadB.ListIndex))
     Rsd.Update
     
     sSql = "INSERT INTO ProductoUnidades ( Producto, CantidadA, UnidadA, CantidadB,UnidadB) "
     sSql = sSql & "VALUES ('" & TxtProducto.text & "'," & Rsd!CantidadA & "," & CmbUnidadA.ItemData(CmbUnidadA.ListIndex) & ","
     sSql = sSql & Rsd!CantidadB & "," & CmbUnidadB.ItemData(CmbUnidadB.ListIndex) & ")"
  Else
     Rsd!CantidadA = TxtCantidadA.text
     Rsd!UnidadA = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & CmbUnidadA.ItemData(CmbUnidadA.ListIndex))
     Rsd!CantidadB = TxtCantidadB.text
     Rsd!Unidadb = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & CmbUnidadB.ItemData(CmbUnidadB.ListIndex))
     Rsd.Update
          
    sSql = "UPDATE ProductoUnidades SET CantidadA=" & Rsd!CantidadA & ",UnidadA=" & CmbUnidadA.ItemData(CmbUnidadA.ListIndex) & ","
    sSql = sSql & "CantidadB=" & Rsd!CantidadB & ",UnidadB=" & CmbUnidadB.ItemData(CmbUnidadB.ListIndex) & " "
    sSql = sSql & "WHERE Id=" & Rsd!ID
    
  End If
  db.Open sDb
  db.Execute sSql
  db.Close
  CrearRs
  LlenarRs
  Set cRsl = Nothing
End Sub

Private Sub HabilitarCab(bHabi As Boolean)
  CmdBuscar.Enabled = bHabi
  TxtProducto.Enabled = bHabi
End Sub
Private Sub Anular()
  If MsgBox("Anula el Registro Seleccionado ?", 20, "Atención") = 6 Then
     Rsd.Bookmark = Grid1.Bookmark
     db.Open sDb
     db.Execute "DELETE FROM ProductoUnidades WHERE Id=" & Rsd!ID
     db.Close
     Rsd.Delete
  End If
End Sub

Private Sub Modificar()
  On Error GoTo errHandler

  If sw = True Then
     Habilitar True, False
     Botones False, False, True, True
     HabilitarCab False
     CmdBotones(2).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(2).Caption = "Grabar"
     CmdBotones(3).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(3).Caption = "Cancelar"
     sw = False
  Else
     GrabarTodo
     sw = True
     CmdBotones(2).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(2).Caption = "Modificar"
     CmdBotones(3).Picture = LoadResPicture("Salir", 0)
     CmdBotones(3).Caption = "Salir"
     Habilitar False, True
     Botones True, True, True, True
     HabilitarCab True
  End If
Exit Sub

errHandler:
     ManejaErrores

End Sub

Private Sub Salir()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        If CmdBotones(0).Enabled = True Then
           If MsgBox("El Registro no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
        CmdBotones(0).Caption = "Nuevo"
        CmdBotones(3).Picture = LoadResPicture("Salir", 0)
        CmdBotones(3).Caption = "Salir"
     Else
        If CmdBotones(2).Enabled = True Then
           If MsgBox("El Registro no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        Else
           If MsgBox("Cancela la Modificación del Comprobante ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
        CmdBotones(2).Caption = "Modificar"
        CmdBotones(3).Picture = LoadResPicture("Salir", 0)
        CmdBotones(3).Caption = "Salir"
     End If
     Habilitar False, True
     sw = True
     Botones True, True, True, True
     HabilitarCab False
  Else
     Unload Me
  End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
'     KeyAscii = Upper(KeyAscii)
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

Private Sub CmdBuscar_Click()
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL, cRsl As ClsLectura
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  sw = False
  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.text = RsP!Producto
     LblProducto.Caption = RsP!Descripcion
     Botones True, True, True, True
     Habilitar False, True
     CrearRs
     LlenarRs
     RsP.Close
     sw = True
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub LlenarRs()
  Dim cRsl As ClsLectura, Rx As ADODB.Recordset
  
  Set cRsl = New ClsLectura
  
  Set Rx = cRsl.TraerRsCondi("ProductoUnidades", "Id", "Producto='" & TxtProducto.text & "'")
  Do While Not Rx.EOF
     Rsd.AddNew
     Rsd!ID = Rx!ID
     Rsd!CantidadA = Rx!CantidadA
     Rsd!UnidadA = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rx!UnidadA)
     Rsd!CantidadB = Rx!CantidadB
     Rsd!Unidadb = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rx!Unidadb)
     Rsd.Update
     Rx.MoveNext
  Loop
  If Rx.RecordCount <> 0 Then
     Rsd.MoveFirst
     LinkearTexto
  End If
End Sub

Private Sub Form_Load()
  Limpiar
  CargarCombos
  Habilitar False, False
  CrearRs
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
  CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
  CmdBotones(3).Picture = LoadResPicture("Salir", 0)
  Botones False, False, False, False
  sw = False
  
End Sub

Private Sub CrearRs()
 On Error GoTo errHandler

 Set Rsd = New ADODB.Recordset
 
 Rsd.Fields.Append "Id", adInteger
 Rsd.Fields.Append "CantidadA", adSingle, , adFldIsNullable
 Rsd.Fields.Append "UnidadA", adVarChar, 30, adFldIsNullable
 Rsd.Fields.Append "CantidadB", adSingle, , adFldIsNullable
 Rsd.Fields.Append "UnidadB", adVarChar, 30, adFldIsNullable
 Rsd.Open
 
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
      .Columns(1).Caption = "Cant."
      .Columns(1).Width = 800
      .Columns(1).Alignment = dbgRight
      .Columns(1).NumberFormat = nCantDecimales
      .Columns(2).Caption = "Unidad Principal"
      .Columns(2).Width = 3000
      .Columns(3).Caption = "Cant."
      .Columns(3).Width = 800
      .Columns(3).Alignment = dbgRight
      .Columns(3).NumberFormat = nCantDecimales
      .Columns(4).Caption = "Unidad de Stock"
      .Columns(4).Width = 3000
 End With
End Sub

Private Sub Habilitar(bEstado As Boolean, bGrid As Boolean)
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

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbUnidadA, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbUnidadB, "UnidadMedida", "Unidad", "Descripcion", ""

  Set cRsl = Nothing
End Sub
Private Sub Limpiar()
  TxtProducto.text = ""
  LblProducto.Caption = ""
  TxtCantidadA.text = "1.00"
  TxtCantidadB.text = "1.00"
  CmbUnidadA.ListIndex = -1
  CmbUnidadB.ListIndex = -1
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler

  If sw = True Then
     If Grid1.Row <> -1 Then
        Rsd.Bookmark = Grid1.Bookmark
        LinkearTexto
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub LinkearTexto()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  TxtCantidadA.text = Format(Rsd!CantidadA, nCantDecimales)
  CmbUnidadA.text = Rsd!UnidadA
  TxtCantidadB.text = Format(Rsd!CantidadB, nCantDecimales)
  CmbUnidadB.text = Rsd!Unidadb
  Set cRsl = Nothing
End Sub

Private Sub TxtCantidadA_GotFocus()
  TxtCantidadA.SelStart = 0
  TxtCantidadA.SelLength = Len(TxtCantidadA.text)
End Sub

Private Sub TxtCantidadA_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtCantidadA) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCantidadB_LostFocus()
  If TxtCantidadB.text = "" Then TxtCantidadB.text = 1
  TxtCantidadB.text = Format(TxtCantidadB.text, nCantDecimales)
End Sub

Private Sub TxtCantidadB_GotFocus()
  TxtCantidadB.SelStart = 0
  TxtCantidadB.SelLength = Len(TxtCantidadB.text)
End Sub

Private Sub TxtCantidadB_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtCantidadB) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCantidadA_LostFocus()
  If TxtCantidadA.text = "" Then TxtCantidadA.text = 1
  TxtCantidadA.text = Format(TxtCantidadA.text, nCantDecimales)
End Sub

Private Sub TxtProducto_GotFocus()
  TxtProducto.SelStart = 0
  TxtProducto.SelLength = Len(TxtProducto.text)
End Sub

Private Sub TxtProducto_LostFocus()
  Dim cDes As ClsProductoL
  
  Set cDes = New ClsProductoL
  
  LblProducto.Caption = cDes.BuscarNombreProducto(TxtProducto.text)
  If LblProducto.Caption = "" Then
     TxtProducto.text = ""
  End If
  Set cDes = Nothing
End Sub
