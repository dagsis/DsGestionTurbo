VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentaRepartoEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos Entregados"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   6945
   Begin VB.CommandButton CmbTerminar 
      Caption         =   "Terminar"
      Height          =   375
      Left            =   5355
      TabIndex        =   7
      Top             =   5460
      Width           =   1335
   End
   Begin VB.CommandButton CmbAplicar 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   3915
      TabIndex        =   0
      Top             =   5460
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Remitos"
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6780
      Begin VB.CommandButton CmbOk 
         Caption         =   "Aplicar"
         Height          =   285
         Left            =   5970
         TabIndex        =   4
         Top             =   990
         Width           =   645
      End
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Left            =   2790
         TabIndex        =   3
         Top             =   390
         Width           =   600
      End
      Begin VB.TextBox TxtSucursal 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "0000"
         Top             =   390
         Width           =   495
      End
      Begin VB.TextBox TxtNumero 
         Height          =   315
         Left            =   1875
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "00000000"
         Top             =   390
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   3795
         Left            =   120
         TabIndex        =   6
         Top             =   1425
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6694
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
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCliente"
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
         Left            =   150
         TabIndex        =   10
         Top             =   975
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   "Número :"
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
         Left            =   195
         TabIndex        =   9
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label17 
         Caption         =   "--"
         Height          =   195
         Left            =   1650
         TabIndex        =   8
         Top             =   375
         Width           =   150
      End
   End
End
Attribute VB_Name = "VentaRepartoEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rImp As Recordset, Rs As Recordset


Private Sub CmbAplicar_Click()
 On Error GoTo errHandler
  
  
  If rImp.RecordCount <> 0 Then
     db.Open sDb
     db.BeginTrans

     rImp.MoveFirst
     Do While Not rImp.EOF
        db.Execute "UPDATE RegTransporte SET Estado=3 WHERE Movimiento=" & rImp!Movimiento
        rImp.MoveNext
     Loop
     db.CommitTrans
     CrearRx
     db.Close
  End If
  
Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Private Sub CmbOk_Click()

  rImp.AddNew
  rImp!REMITO = Rs!Numero
  rImp!fecha = Rs!fecha
  rImp!Cliente = Rs!Destinatario
  rImp!Movimiento = Rs!Movimiento
  rImp.Update
  
  Limpiar
  TxtNumero.SetFocus
  
End Sub

Private Sub CmbTerminar_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
  Dim cRx As ClsComprobantesL
  
  Set cRx = New ClsComprobantesL
  
  If TxtSucursal.text <> 0 And TxtNumero.text <> 0 Then
     Set Rs = cRx.TraerRemitoEntrega(TxtSucursal.text, TxtNumero.text)
     If Rs.RecordCount <> 0 Then
        LblCliente.Caption = Rs!fecha & " " & Rs!Destinatario & " " & Rs!Motivo
        CmbOk.SetFocus
     Else
        MsgBox "Remito No Encontrado ...", vbCritical, "Atención"
        TxtNumero.text = "00000000"
        TxtNumero.SelStart = 0
        TxtNumero.SelLength = Len(TxtNumero.text)
     End If
  End If
End Sub

Private Sub Form_Load()
  Limpiar
  CrearRx
  CabGrid
End Sub

Private Sub Limpiar()
  LblCliente.Caption = ""
  TxtNumero.text = "00000000"
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub
Private Sub CrearRx()
  Set rImp = New ADODB.Recordset
  
  rImp.Fields.Append "Remito", adVarChar, 15, adFldIsNullable
  rImp.Fields.Append "Fecha", adVarChar, 10, adFldIsNullable
  rImp.Fields.Append "Cliente", adVarChar, 50, adFldIsNullable
  rImp.Fields.Append "Movimiento", adInteger
    
  rImp.Open
  
  CabGrid
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
     Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = rImp
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Remito"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Fecha"
      .Columns(1).Width = 950
      .Columns(2).Caption = "Cliente"
      .Columns(2).Width = 3950
      .Columns(3).Visible = False
 End With
End Sub

Private Sub TxtSucursal_GotFocus()
  TxtSucursal.SelStart = 0
  TxtSucursal.SelLength = Len(TxtSucursal.text)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSucursal_LostFocus()
  If TxtSucursal.text = "" Then TxtSucursal.text = 1
  TxtSucursal.text = Format(TxtSucursal.text, "0000")
End Sub

Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtNumero_LostFocus()
  If TxtNumero.text = "" Then TxtNumero.text = 0
  TxtNumero.text = Format(TxtNumero.text, "00000000")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If rImp.RecordCount <> 0 Then
    If MsgBox("Tiene Registros Pendientes para Cambiar de Estado. Sale sin Procesar ?", 20, "Atención") = 6 Then
       Exit Sub
    End If
    Cancel = 1
  End If
End Sub
