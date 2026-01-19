VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaDefinirCuotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definir Cuotas"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   735
      Index           =   2
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3615
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   735
      Index           =   1
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3615
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Agregar"
      Height          =   735
      Index           =   0
      Left            =   2745
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3615
      Width           =   855
   End
   Begin VB.TextBox TxtImporte 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1245
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "TxtImporte"
      Top             =   4005
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   315
      Left            =   1245
      TabIndex        =   3
      Top             =   3615
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   37053
   End
   Begin VB.Frame Frame1 
      Caption         =   "Definir Cobros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5295
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2145
         Left            =   105
         TabIndex        =   12
         Top             =   270
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   3784
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
               LCID            =   11274
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
               LCID            =   11274
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
   End
   Begin VB.Label LblResto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblResto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Resto :"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Importe Total :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Importe :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Vencimiento :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "VentaDefinirCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rAReci As ADODB.Recordset, sw As Boolean
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Public Sub RemoveCancelMenuItem(frm As Form)
  Dim hSysMenu As Long
  hSysMenu = GetSystemMenu(frm.hwnd, 0)
  Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
  Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub


Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Agregar
         Case 1
              Borrar
         Case 2
              salir
  End Select
End Sub

Private Sub Form_Load()
  
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
  CmdBotones(2).Picture = LoadResPicture("Salir", 0)
 Limpiar
 If rAReci Is Nothing Then
    CrearRs
 Else
    CabGrid
    Calcular
 End If
 Botones True, False, True
 RemoveCancelMenuItem Me
 TxtImporte.text = Format(LblResto.Caption, nCantDecimales)
End Sub
Private Sub Limpiar()
 DTPFecha.Value = Date
 TxtImporte.text = "0.00"
 LblTotal.Caption = Format(VentaComprobantes.LblTotal.Caption, "0.00")
 LblResto.Caption = Format(LblTotal.Caption, nCantDecimales)
End Sub
Private Sub CrearRs()
  Set rAReci = New ADODB.Recordset
  rAReci.Fields.Append "Fecha", adDate
  rAReci.Fields.Append "Importe", adCurrency
  rAReci.Open
  CabGrid
End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = rAReci
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Width = 3100
      .Columns(0).Caption = "Vencimiento"
      .Columns(1).Width = 1200
      .Columns(1).Caption = "Importe"
      .Columns(1).NumberFormat = nCantDecimales
      .Columns(1).Alignment = dbgRight
 End With

End Sub

Private Sub Agregar()
 sw = True
 rAReci.AddNew
 rAReci!fecha = DTPFecha.Value
 rAReci!importe = TxtImporte.text
 rAReci.Update
 Calcular
 sw = False
End Sub
Private Sub Borrar()
  If rAReci.RecordCount <> 0 Then
     rAReci.Delete
     Calcular
  End If
End Sub
Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler
  If sw = False Then
     If Grid1.Row <> -1 Then
        rAReci.Bookmark = Grid1.Bookmark
        Botones True, True, False
        If LblResto.Caption = 0 Then
           Botones False, True, True
        End If
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub Botones(bAgr As Boolean, bBorr As Boolean, bSal As Boolean)
  CmdBotones(0).Enabled = bAgr
  CmdBotones(1).Enabled = bBorr
  CmdBotones(2).Enabled = bSal
End Sub
Private Sub salir()
  Unload Me
End Sub

Private Sub Calcular()
 Dim Rs As ADODB.Recordset
 
 Botones True, False, False
 Set Rs = New ADODB.Recordset
 Set Rs = rAReci
 LblResto.Caption = Format(LblTotal.Caption, nCantDecimales)
 If Rs.RecordCount <> 0 Then
    Rs.MoveFirst
    Do While Not Rs.EOF
       LblResto.Caption = LblResto.Caption - Rs!importe
       LblResto.Caption = Format(LblResto.Caption, nCantDecimales)
       Rs.MoveNext
    Loop
 End If
 TxtImporte.text = Format(LblResto.Caption, nCantDecimales)
 Botones True, False, False
 If LblResto.Caption = 0 Then
    Botones False, False, True
 End If
 If rAReci.RecordCount = 0 Then
    Botones True, False, True
 End If
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
  If TxtImporte.text = "" Then TxtImporte.text = 0
  TxtImporte.text = Format(TxtImporte.text, nCantDecimales)
End Sub
