VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentaGrillaProduc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas de Producto/Stock/Precios"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   11565
   Begin VB.ComboBox CmbSucursal 
      Height          =   315
      Left            =   7695
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   75
      Width           =   3390
   End
   Begin VB.ComboBox CmbEmpresa 
      Height          =   315
      Left            =   1095
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   3390
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Desc. Adicional"
      Height          =   225
      Index           =   2
      Left            =   9375
      TabIndex        =   5
      Top             =   525
      Width           =   1830
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Descripción"
      Height          =   225
      Index           =   1
      Left            =   7725
      TabIndex        =   4
      Top             =   525
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Código"
      Height          =   225
      Index           =   0
      Left            =   6510
      TabIndex        =   3
      Top             =   525
      Width           =   1335
   End
   Begin VB.TextBox TxtBuscar 
      Height          =   285
      Left            =   1095
      TabIndex        =   0
      Text            =   "TxtBuscar"
      Top             =   510
      Width           =   5130
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   4740
      Left            =   75
      TabIndex        =   1
      Top             =   930
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   8361
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
      Caption         =   "Sucursal :"
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
      Left            =   6720
      TabIndex        =   9
      Top             =   90
      Width           =   885
   End
   Begin VB.Label Label2 
      Caption         =   "Empresa :"
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
      Left            =   120
      TabIndex        =   7
      Top             =   135
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Criterio :"
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
      Left            =   135
      TabIndex        =   2
      Top             =   540
      Width           =   825
   End
End
Attribute VB_Name = "VentaGrillaProduc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Rsd As Recordset, sw As Boolean, sAux As String

Private Sub CmbEmpresa_Click()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  If sw = True Then
     TxtBuscar_LostFocus
  End If
  sw = True
  
  Set cRsl = Nothing
End Sub

Private Sub CmbSucursal_Click()
  If sw = True Then
     TxtBuscar_LostFocus
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
  KeyAscii = Upper(KeyAscii)
End Sub

Private Sub LlenarGrilla(sBus As String, X As Byte)
  Dim Rx As Recordset, cRx As ClsProductoL, cRsl As ClsLectura
  Dim pSuc As Byte
  
  pSuc = CmbSucursal.ItemData(CmbSucursal.ListIndex)
  If CmbSucursal.ItemData(CmbSucursal.ListIndex) = nSucursal Then
     pSuc = 0
  End If
  
  Set cRx = New ClsProductoL
  Set cRsl = New ClsLectura
  Set Rx = cRx.TraerListaProduc(sBus, X, pSuc)
    
 ' CabGrid
  If Rx.RecordCount <> 0 Then
     Do While Not Rx.EOF
        Rsd.AddNew
        Rsd!Producto = Rx!Producto
        Rsd!Descripcion = Rx!Descripcion
        Rsd!Adicional = Rx!Descripcion1
        Rsd!Cantidad = Rx!Cantidad
        Rsd!precio = cRsl.TraerValorDeUnCampo("Precios", "Precio", "Producto='" & Rx!Producto & "' AND Lista=1")
        Rsd.Update
        Rx.MoveNext
     Loop
     Rsd.MoveFirst
  End If
End Sub

Private Sub Form_Activate()
  TxtBuscar.SetFocus
  Option1(1).Value = True
End Sub

Private Sub Form_Load()
  TxtBuscar.text = ""
  sAux = sDb
  CabGrid
  CargarCombos
End Sub

Private Sub CargarCombos()
  Dim cRsl As ClsLectura, sAux As String
  Set cRsl = New ClsLectura

  sAux = sDb
  sDb = sUdl
  cRsl.CargaCombo CmbEmpresa, "Empresa", "Id", "Razon", ""
  CmbEmpresa.ListIndex = nMenu - 1
  sDb = sAux
  cRsl.CargaCombo CmbSucursal, "Depositos", "Deposito", "Descripcion", ""

  CmbSucursal.text = cRsl.TraerValorDeUnCampo("Depositos", "Descripcion", "Suc=" & nSucursal)
  
  Set cRsl = Nothing
End Sub
Private Sub CabGrid()
  Set Rsd = New Recordset

  Rsd.Fields.Append "Producto", adVarChar, 25
  Rsd.Fields.Append "Descripcion", adVarChar, 50
  Rsd.Fields.Append "Adicional", adVarChar, 50, adFldIsNullable
  Rsd.Fields.Append "Cantidad", adSingle
  Rsd.Fields.Append "Precio", adCurrency
  Rsd.Open
  
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Producto"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Descripción"
      .Columns(1).Width = 4000
      .Columns(2).Caption = "Desc. Adicional"
      .Columns(2).Width = 3900
      .Columns(3).Caption = "Cant."
      .Columns(3).Width = 800
      .Columns(3).Alignment = dbgRight
      .Columns(3).NumberFormat = "#0.000"
      .Columns(4).Caption = "P.Uni."
      .Columns(4).Width = 1100
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(4).Alignment = dbgRight
 End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  db.Close
'  db.Open sDb
End Sub

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
  Rsd.Sort = Rsd.Fields(ColIndex).Name & " Asc"
End Sub

Private Sub Option1_Click(Index As Integer)
  TxtBuscar.SetFocus
  TxtBuscar.text = ""
  CabGrid
End Sub

Private Sub TxtBuscar_LostFocus()
  Dim X As Byte, cRsl As ClsLectura, sAhora As String
  
  Set cRsl = New ClsLectura
  If Option1(0).Value = True Then X = 1
  If Option1(1).Value = True Then X = 2
  If Option1(2).Value = True Then X = 3
  
  If TxtBuscar.text <> "" Then
   '  db.Close
     
     sDb = sUdl

     sAhora = "FILE NAME=" & App.Path & "\" & cRsl.TraerValorDeUnCampo("Empresa", "Udl", "Id=" & CmbEmpresa.ListIndex)
     sDb = sAhora
     db.Open sAhora
     CabGrid
     LlenarGrilla TxtBuscar.text, X
  End If
'  db.Open sDb
  sDb = sAux
  
End Sub
