VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form StockConciliar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliar Transferencias"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   6960
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
      Height          =   5370
      Left            =   90
      TabIndex        =   7
      Top             =   15
      Width           =   6780
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "Aplicar"
         Height          =   255
         Left            =   2610
         TabIndex        =   4
         Top             =   1230
         Width           =   1200
      End
      Begin VB.ComboBox CmbDesde 
         Height          =   315
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1215
         Width           =   1050
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Left            =   5850
         TabIndex        =   2
         Top             =   870
         Width           =   765
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   1
         Text            =   "TxtProducto"
         Top             =   885
         Width           =   1290
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   3645
         Left            =   120
         TabIndex        =   8
         Top             =   1605
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6429
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
      Begin VB.Label Label2 
         Caption         =   "Deposito :"
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
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   10
         Top             =   1215
         Width           =   1020
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
         Left            =   1440
         TabIndex        =   9
         Top             =   870
         Width           =   4305
      End
   End
   Begin VB.CommandButton CmbAplicar 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   4005
      TabIndex        =   6
      Top             =   5475
      Width           =   1335
   End
   Begin VB.CommandButton CmbTerminar 
      Caption         =   "Terminar"
      Height          =   375
      Left            =   5445
      TabIndex        =   5
      Top             =   5475
      Width           =   1335
   End
End
Attribute VB_Name = "StockConciliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRxl As ClsLectura
Dim Cadena As String, ic As Byte
Public rImp As ADODB.Recordset
Dim Rx As ADODB.Recordset, nCantOld As Double

Private Sub CmbAplicar_Click()
  On Error GoTo errHandler
    
  If rImp.RecordCount <> 0 Then
     db.Open sDb
     db.BeginTrans

     rImp.MoveFirst
     Do While Not rImp.EOF
        db.Execute "UPDATE Cantidades SET Estado=1 WHERE Producto='" & rImp!codigo & "' AND Estado=2 and  Deposito=" & CmbDesde.ItemData(CmbDesde.ListIndex)
        rImp.MoveNext
     Loop
     db.CommitTrans
     db.Close

     CrearRx
  End If
Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Private Sub CmdAplicar_Click()
  Dim cRxl As ClsLectura
'
  Set cRxl = New ClsLectura
'
  rImp.AddNew
  rImp!codigo = TxtProducto.text
  rImp!Barra = cRxl.TraerValorDeUnCampo("Productos", "Peso", "Producto='" & TxtProducto.text & "'")
  LblProducto.Caption = cRxl.TraerValorDeUnCampo("Productos", "Descripcion", "Producto='" & TxtProducto.text & "'")
  rImp!Descripcion = cRxl.TraerValorDeUnCampo("Productos", "Descripcion", "Producto='" & TxtProducto.text & "'")
'
  rImp!Cantidad = TxtCantidad.text
  rImp.Update
'
  nCantOld = 0
  Limpiar
  TxtProducto.SetFocus
  Set cRxl = Nothing

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
  CargarCombos
End Sub

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  CmbDesde.ListIndex = 0 ' cRsl.RegPorDefecto("Deposito")
  Set cRsl = Nothing
End Sub
Private Sub CrearRx()
  Set rImp = New ADODB.Recordset
  
  rImp.Fields.Append "Codigo", adVarChar, 15, adFldIsNullable
  rImp.Fields.Append "Barra", adDouble
  rImp.Fields.Append "Descripcion", adVarChar, 50, adFldIsNullable
  rImp.Fields.Append "Cantidad", adCurrency
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
      .Columns(3).Caption = "Cant."
      .Columns(3).Width = 950
      .Columns(3).NumberFormat = "#0.00"
      .Columns(3).Alignment = dbgRight
 End With
End Sub


Private Sub TxtProducto_GotFocus()
  TxtProducto.SelStart = 0
  TxtProducto.SelLength = Len(TxtProducto.text)
End Sub

Private Sub TxtProducto_LostFocus()
  Dim cDes As ClsProductoL, nCanti As Double, cRsl As ClsLectura
  
  Set cDes = New ClsProductoL
  Set cRsl = New ClsLectura
  LblProducto.Caption = cDes.BuscarNombreProducto(TxtProducto.text)
  If LblProducto.Caption = "" Then
     TxtProducto.text = ""
  Else
    ' If nPuerto = 0 Then
        nCanti = cRsl.SumarCantidad("Cantidades", "CantidadDebe", "producto='" & TxtProducto.text & "' and estado=2 and deposito=" & CmbDesde.ItemData(CmbDesde.ListIndex))
        nCantOld = nCanti
        If nCanti <> 0 Then
              TxtCantidad.text = Format(nCanti, "#0.00")
              TxtCantidad.SetFocus
        Else
           MsgBox "Producto sin Stock en transito", vbInformation, "Atención"
           Limpiar
           TxtProducto.SetFocus
       End If
    ' End If
  End If

End Sub

Private Sub CmdBuscar_Click()
  Dim RxP As ADODB.Recordset, cRpl As ClsProductoL, nCanti As Double
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
     nCanti = cRxl.SumarCantidad("Cantidades", "CantidadDebe", "producto='" & TxtProducto.text & "' and estado=2 and deposito=" & CmbDesde.ItemData(CmbDesde.ListIndex))
     nCantOld = nCanti
     If nCanti <> 0 Then
        TxtCantidad.text = Format(nCanti, "#0.00")
        TxtCantidad.SetFocus
     Else
        MsgBox "Producto sin Stock en transito", vbInformation, "Atención"
        Limpiar
        TxtProducto.SetFocus
     End If
     RxP.Close
  End If
  Set cRxl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Limpiar()
  TxtProducto.text = ""
  LblProducto.Caption = ""
  TxtCantidad.text = "0.00"
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
  Dim nCanti As Double
  If TxtCantidad.text = "" Then TxtCantidad.text = 1
     If nCantOld < TxtCantidad.text Then
        MsgBox "No Puede Recibir mas de los Productos Transferidos", vbCritical, "Atención"
        TxtCantidad.text = Format(nCantOld, "#0.00")
        TxtCantidad.SetFocus
   End If
  TxtCantidad.text = Format(TxtCantidad.text, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If rImp.RecordCount <> 0 Then
    If MsgBox("Tiene Registros Pendientes para Conciliar. Sale sin Procesar ?", 20, "Atención") = 6 Then
       Exit Sub
    End If
    Cancel = 1
  End If
End Sub


