VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MtrEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas de Productos"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   7020
   Begin VB.CommandButton CmdTranfiere 
      Caption         =   "Actualizar Precios a Caja/s"
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   4005
      Width           =   2565
   End
   Begin VB.CommandButton CmbTerminar 
      Caption         =   "Terminar"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton CmbAplicar 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6780
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   5
         Text            =   "TxtProducto"
         Top             =   270
         Width           =   1290
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Left            =   5085
         TabIndex        =   4
         Top             =   240
         Width           =   765
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
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
               LCID            =   3082
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
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
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
         Height          =   285
         Left            =   1455
         TabIndex        =   6
         Top             =   270
         Width           =   3600
      End
   End
End
Attribute VB_Name = "MtrEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRxl As ClsLectura
Dim Cadena As String, ic As Byte
Public rImp As ADODB.Recordset
Dim Rx As ADODB.Recordset

Private Sub CmbAplicar_Click()
  nImpr = 47
  If pEti = 2 Then
     nImpr = 48
  End If
  FrmImpresor.Show
End Sub

Private Sub CmdTranfiere_Click()
  Dim cRsl As ClsProductoL
      
  Set cRsl = New ClsProductoL
  
  If MsgBox("Transfiere Los Precios Del Server a las Caja/s ?", 20, "Atención") = 6 Then
     If rImp.RecordCount <> 0 Then
         rImp.MoveFirst
         Do While Not rImp.EOF
            cRsl.ActualizarPreciosCajaMovi rImp!codigo, rImp!precio
            rImp.MoveNext
         Loop
         MsgBox "Operación Realizada Con Exito", vbInformation, "Atención"
     End If
 End If
 
End Sub

Private Sub Form_Activate()
  TxtProducto.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub CmbTerminar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CrearRx
  Cadena = ""
  ic = 1
  Limpiar
End Sub

Private Sub CrearRx()
  Set rImp = New ADODB.Recordset
  
  rImp.Fields.Append "Codigo", adVarChar, 15, adFldIsNullable
  rImp.Fields.Append "Barra", adDouble
  rImp.Fields.Append "Descripcion", adVarChar, 50, adFldIsNullable
  rImp.Fields.Append "Descripcion1", adVarChar, 50, adFldIsNullable
  rImp.Fields.Append "Precio", adCurrency
  rImp.Open
  
  CabGrid
End Sub
Private Sub CabGrid()
  Set Grid1.DataSource = rImp
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Código"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Barra"
      .Columns(1).Width = 0
      .Columns(2).Caption = "Descripción"
      .Columns(2).Width = 4000
      .Columns(3).Visible = False
      .Columns(4).Caption = "Precio"
      .Columns(4).Width = 950
      .Columns(4).NumberFormat = "#0.00"
      .Columns(4).Alignment = dbgRight
 End With
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
  Else
   '  If nPuerto = 0 Then
        Grabar
        TxtProducto.text = ""
        LblProducto.Caption = ""
        TxtProducto.SetFocus
    ' End If
  End If

End Sub


Private Sub CmdBuscar_Click()
  Dim RxP As ADODB.Recordset, cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set RxP = New ADODB.Recordset
  
  Set cRxl = New ClsLectura
  Set cRpl = New ClsProductoL
  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  Limpiar
  If nDat <> 0 Then
     Set RxP = cRxl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto = RxP!Producto
     LblProducto.Caption = RxP!Descripcion
     TxtProducto.SetFocus
     RxP.Close
     Grabar
     TxtProducto.text = ""
  End If
  Set cRxl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub


Private Sub Grabar()
  Dim cRxl As ClsLectura
  
  Set cRxl = New ClsLectura
  Dim nListaVenta As Integer
  
  nListaVenta = cRxl.RegPorDefecto("ListaRes")
  
  rImp.AddNew
  rImp!codigo = TxtProducto.text
  rImp!Barra = cRxl.TraerValorDeUnCampo("Productos", "Peso", "Producto='" & TxtProducto.text & "'")
  LblProducto.Caption = cRxl.TraerValorDeUnCampo("Productos", "Descripcion", "Producto='" & TxtProducto.text & "'")
  rImp!Descripcion = cRxl.TraerValorDeUnCampo("Productos", "Descripcion", "Producto='" & TxtProducto.text & "'")
  rImp!Descripcion1 = ""
  rImp!precio = cRxl.TraerValorDeUnCampo("Precios", "Precio", "Lista=" & nListaVenta & " and Producto='" & TxtProducto.text & "'")
  rImp.Update
  
  Set cRxl = Nothing
End Sub
Private Sub Limpiar()
  TxtProducto.text = ""
  LblProducto.Caption = ""
End Sub

