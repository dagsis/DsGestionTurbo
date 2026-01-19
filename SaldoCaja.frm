VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SaldoCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Venta"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6375
   Begin VB.ComboBox CmbFamilia 
      Height          =   315
      Left            =   945
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   585
      Width           =   3315
   End
   Begin VB.CommandButton CmdPlay 
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
      Height          =   285
      Left            =   4365
      TabIndex        =   2
      Top             =   195
      Width           =   510
   End
   Begin MSComCtl2.DTPicker DtpHasta 
      Height          =   300
      Left            =   3015
      TabIndex        =   1
      Top             =   180
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   209321985
      CurrentDate     =   39108
   End
   Begin MSComCtl2.DTPicker DtpDesde 
      Height          =   300
      Left            =   915
      TabIndex        =   0
      Top             =   180
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   209321985
      CurrentDate     =   39108
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   2460
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   4339
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Familia :"
      Height          =   210
      Left            =   225
      TabIndex        =   9
      Top             =   615
      Width           =   690
   End
   Begin VB.Label Label4 
      Caption         =   "Desde :"
      Height          =   240
      Left            =   2340
      TabIndex        =   8
      Top             =   195
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Desde :"
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   195
      Width           =   645
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Caja :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4245
      TabIndex        =   6
      Top             =   3630
      Width           =   1950
   End
   Begin VB.Label Label2 
      Caption         =   "Total Caja :"
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
      Left            =   3105
      TabIndex        =   5
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ESC = Volver"
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
      Left            =   150
      TabIndex        =   4
      Top             =   3855
      Width           =   1335
   End
End
Attribute VB_Name = "SaldoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As Recordset

Private Sub CmdPlay_Click()
  Dim Rx As Recordset, cRl As ClsComprobantesL
  
  Set Rx = New Recordset
  Set cRl = New ClsComprobantesL
  
  lblTotal.Caption = "0.00"
  If CmbFamilia.text = "Todos" Then
      Set Rx = cRl.TraerVentas(DtpDesde.Value, DtpHasta.Value)
      If Rs.RecordCount <> 0 Then
         Rs.MoveFirst
         Do While Not Rs.EOF
            Rs.Delete
            Rs.MoveNext
         Loop
      End If
      Do While Not Rx.EOF
         Rs.AddNew
         Rs!CCompro = Rx!Cantidad
         Rs!Comprobante = Rx!Descripcion
         Rs!importe = Rx!SumaDePTotal
         Rs.Update
         lblTotal.Caption = lblTotal.Caption + Rx!SumaDePTotal
         Rx.MoveNext
      Loop
  Else
      Set Rx = cRl.TraerIvaPorFamilia(CmbFamilia.ItemData(CmbFamilia.ListIndex), True, DtpDesde.Value, DtpHasta.Value)
      If Rs.RecordCount <> 0 Then
         Rs.MoveFirst
         Do While Not Rs.EOF
            Rs.Delete
            Rs.MoveNext
         Loop
      End If
          
      Do While Not Rx.EOF
         Rs.AddNew
         Rs!CCompro = 0
         Rs!Comprobante = Rx!Descripcion & " Iva 21% "
         Rs!importe = Rx!SumaDePrecioTotal
         Rs.Update
         lblTotal.Caption = lblTotal.Caption + Rx!SumaDePrecioTotal
         Rx.MoveNext
      Loop
      Set Rx = cRl.TraerIvaPorFamilia(CmbFamilia.ItemData(CmbFamilia.ListIndex), False, DtpDesde.Value, DtpHasta.Value)
      Do While Not Rx.EOF
        Rs.AddNew
        Rs!CCompro = 0
        Rs!Comprobante = Rx!Descripcion & " Iva 0% "
        Rs!importe = Rx!SumaDePrecioTotal
        Rs.Update
        lblTotal.Caption = lblTotal.Caption + Rx!SumaDePrecioTotal
        Rx.MoveNext
      Loop
  End If
  lblTotal.Caption = Format(lblTotal.Caption, "#0.00")
End Sub

Private Sub DtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub DtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Unload Me
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
  lblTotal.Caption = "0.00"
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  CargarCombos
  CrearRs
End Sub

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbFamilia, "Familias", "Familia", "Descripcion", ""
  CmbFamilia.AddItem "Todos", 0
  
  CmbFamilia.text = "Todos"
End Sub

Private Sub CrearRs()
  Set Rs = New Recordset
  
  Rs.Fields.Append "CCompro", adInteger
  Rs.Fields.Append "Comprobante", adVarChar, 30
  Rs.Fields.Append "Importe", adCurrency
  Rs.Open
  
  CabGrid
End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = Rs
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "C.Comp."
      .Columns(0).Width = 1200
      .Columns(0).Alignment = dbgRight
      .Columns(1).Caption = "Descripción"
      .Columns(1).Width = 2800
      .Columns(2).Caption = "Total"
      .Columns(2).Width = 1500
      .Columns(2).NumberFormat = "0.00"
      .Columns(2).Alignment = dbgRight
 End With
End Sub
  
