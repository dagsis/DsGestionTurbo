VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Delivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   10695
   Begin VB.CheckBox ChkVer 
      Caption         =   "Ver Entregados por Motos"
      Height          =   255
      Left            =   2940
      TabIndex        =   9
      Top             =   150
      Width           =   2505
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Cierres"
      Height          =   435
      Index           =   3
      Left            =   2820
      TabIndex        =   6
      Top             =   6855
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ComboBox CmbMoto 
      Height          =   315
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   2085
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   435
      Index           =   1
      Left            =   7545
      TabIndex        =   3
      Top             =   6855
      Width           =   1410
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   435
      Index           =   0
      Left            =   9075
      TabIndex        =   2
      Top             =   6855
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Delivery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   135
      TabIndex        =   0
      Top             =   570
      Width           =   10425
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   5775
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   10140
         _ExtentX        =   17886
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
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblTotal"
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
      Left            =   1425
      TabIndex        =   8
      Top             =   6870
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Total Moto :"
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
      Left            =   285
      TabIndex        =   7
      Top             =   6870
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Moto :"
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
      Left            =   165
      TabIndex        =   5
      Top             =   150
      Width           =   660
   End
End
Attribute VB_Name = "Delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rPed As Recordset
Dim nTotal As Double

Private Sub ChkVer_Click()
  Dim cRsl As ClsPrograma
  
  Set cRsl = New ClsPrograma
  
  nTotal = 0
  Limpiar
  Set rPed = cRsl.TraerDelivery(CmbMoto.ItemData(CmbMoto.ListIndex), IIf(ChkVer.Value = 1, True, False))
  If ChkVer.Value = 1 Then
     LblTotal.Caption = Format(cRsl.DameTotalMoto(CmbMoto.ItemData(CmbMoto.ListIndex)), "#0.00")
  End If

  CabGrid

End Sub

Private Sub CmbMoto_Click()
  Dim cRsl As ClsPrograma
  
  Set cRsl = New ClsPrograma
  
  nTotal = 0
  Limpiar
  Set rPed = cRsl.TraerDelivery(CmbMoto.ItemData(CmbMoto.ListIndex), IIf(ChkVer.Value = 1, True, False))
  CabGrid
  If ChkVer.Value = 1 Then
     LblTotal.Caption = Format(cRsl.DameTotalMoto(CmbMoto.ItemData(CmbMoto.ListIndex)), "#0.00")
  End If

End Sub

Private Sub CmdBotones_Click(Index As Integer)
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Select Case Index
    Case 1
         CmbMoto.text = cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & cRsl.RegPorDefecto("Vendedor"))
         ChkVer.Value = 0
         Limpiar
         NuevoDelivery.Show 1
    Case 0
         Unload Me
  End Select
End Sub


Private Sub Form_Load()
  Limpiar
  CargarCombos
  LlenarGrilla
End Sub

Private Sub Limpiar()
  LblTotal.Caption = "0.00"
End Sub
Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbMoto, "Vendedores", "Vendedor", "Descripcion", ""
  CmbMoto.text = cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & cRsl.RegPorDefecto("Vendedor"))
  
  Set cRsl = New ClsLectura
  
End Sub

Public Sub LlenarGrilla()
  Dim cRsl As ClsPrograma
  
  Set cRsl = New ClsPrograma
  
  nTotal = 0
  Limpiar
  Set rPed = cRsl.TraerDelivery(CmbMoto.ItemData(CmbMoto.ListIndex), IIf(ChkVer.Value = 1, True, False))
  CabGrid

End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  If KeyCode = 13 Then
     CmbMoto.text = cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & cRsl.RegPorDefecto("Vendedor"))
     ChkVer.Value = 0
     If Grid1.Row <> -1 Then
        rPed.Bookmark = Grid1.Bookmark
        nPedido = rPed![a_Mesa]
        bDelivery = True
        nLista = cRsl.RegPorDefecto("Precio")
        nMesa = nPedido
      
       frmMesa.Show 1
    End If
  End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  If Button = 2 Then
'     nTotal = 0
'     Limpiar
'     CmbMoto.text = cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & cRsl.RegPorDefecto("Vendedor"))
'     ChkVer.Value = 0
  
     If Grid1.ApproxCount > 0 Then
         MenuDelivery.MenuMovi.Enabled = True
         MenuDelivery.MenuSalidaMoto.Enabled = True
         MenuDelivery.MenuAnular.Enabled = False
         MenuDelivery.MenuPedido.Enabled = True
     Else
         MenuDelivery.MenuMovi.Enabled = True
         MenuDelivery.MenuSalidaMoto.Enabled = False
         MenuDelivery.MenuAnular.Enabled = False
         MenuDelivery.MenuPedido.Enabled = False
     End If
     Delivery.PopupMenu MenuDelivery.MenuDelivery
  End If
End Sub

Public Sub CabGrid()
  Set Grid1.DataSource = rPed
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Caption = "Ped"
      .Columns(1).Width = 650
      .Columns(2).Caption = "Hora"
      .Columns(2).Width = 750
      .Columns(3).Caption = "Entrega"
      .Columns(3).Width = 1200
      .Columns(4).Caption = "Observación"
      .Columns(4).Width = 2000
      .Columns(5).Caption = "Cliente"
      .Columns(5).Width = 2400
      .Columns(6).Caption = "Dirección"
      .Columns(6).Width = 2400
      .Columns(7).Visible = False
      .Columns(8).Visible = False
'      .Columns(9).Visible = False
'      .Columns(10).Visible = False
 End With
End Sub
