VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentaCobranzasDetalles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Facturas"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   7575
      TabIndex        =   0
      Top             =   3195
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   2925
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   5159
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
End
Attribute VB_Name = "VentaCobranzasDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rDet As ADODB.Recordset, rs As ADODB.Recordset, rAux As ADODB.Recordset

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CrearRs
  LlenarGrilla
End Sub
Private Sub CrearRs()
  Set rDet = New ADODB.Recordset
  rDet.Fields.Append "Fecha", adDate
  rDet.Fields.Append "Comprobante", adVarChar, 20
  rDet.Fields.Append "Numero", adDouble
  rDet.Fields.Append "Producto", adVarChar, 25, adFldIsNullable
  rDet.Fields.Append "Cantidad", adDouble
  rDet.Fields.Append "Detalle", adVarChar, 100
  rDet.Open
  CabGrid
End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = rDet
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Width = 950
      .Columns(0).Caption = "Fecha"
      .Columns(1).Width = 1600
      .Columns(1).Caption = "Comprobante"
      .Columns(2).Width = 950
      .Columns(2).Caption = "Número"
      .Columns(2).NumberFormat = "000000"
      .Columns(3).Width = 1000
      .Columns(3).Caption = "Producto"
      .Columns(4).Width = 700
      .Columns(4).Caption = "Cant."
      .Columns(4).Alignment = dbgRight
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(5).Width = 3000
      .Columns(5).Caption = "Descripción"
  End With

End Sub

Private Sub LlenarGrilla()
  Dim cComp As ClsComprobantesL, nMovi As Long
  
  Set rAux = New ADODB.Recordset
  Set cComp = New ClsComprobantesL
  Set rs = New ADODB.Recordset
  
  Set rAux = VentaCobranzas.RsC
  rAux.MoveFirst
  Do While Not rAux.EOF
     Do While nMovi <> rAux!Venta
        Set rs = cComp.TraerDetalles(rAux!Venta)
        nMovi = rAux!Venta
        Do While Not rs.EOF
           rDet.AddNew
           rDet!fecha = rs!fecha
           rDet!Comprobante = rs!CompCompro
           rDet!Numero = rs!Numero
           rDet!Producto = rs!Producto
           rDet!Cantidad = rs!Cantidad
           rDet!Detalle = rs!DesCompro
           rDet.Update
           rs.MoveNext
        Loop
     Loop
     rAux.MoveNext
  Loop
  CabGrid
End Sub

