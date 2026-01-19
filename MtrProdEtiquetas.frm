VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MtrProdEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de Etiquetas"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   4980
   Begin VB.CommandButton CmdTranfiere 
      Caption         =   "Actualizar Precios a Caja/s"
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   2565
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   450
      Left            =   3570
      TabIndex        =   1
      Top             =   2850
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   4740
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   390
         Left            =   135
         TabIndex        =   6
         Top             =   2220
         Width           =   1440
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   330
         Left            =   1980
         TabIndex        =   4
         Top             =   675
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   39035
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   330
         Left            =   1980
         TabIndex        =   5
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   39035
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
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
         Left            =   1245
         TabIndex        =   3
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
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
         Left            =   1245
         TabIndex        =   2
         Top             =   690
         Width           =   735
      End
   End
End
Attribute VB_Name = "MtrProdEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdLimpiar_Click()
  Dim sSql As String, cRsl As ClsLectura
    
  Dim sConex As Connection
  
  Set sConex = New Connection
  
  Set cRsl = New ClsLectura
  If MsgBox("Limpia Los Movimientos de La Fecha Indicada ?", 20, "Atención") = 6 Then
     sConex.Open sDb
     sSql = "UPDATE Productos SET FechaBaja=null "
     sSql = sSql & "WHERE Productos.FechaBaja>='" & FechaAmericana(DtpDesde.Value) & "' AND Productos.FechaBaja<='" & FechaAmericana(DtpHasta.Value) & "'"
     sConex.Execute sSql
     sConex.Close
     MsgBox "Operación Realizada Con Exito", vbInformation, "Atención"
  End If
  
End Sub

Private Sub CmdTranfiere_Click()
  Dim cRsl As ClsProductoL
      
  Set cRsl = New ClsProductoL
  
  If MsgBox("Transfiere Los Precios Del Server a las Caja/s ?", 20, "Atención") = 6 Then
     cRsl.ActualizarPreciosCajas DtpDesde.Value, DtpHasta.Value
     MsgBox "Operación Realizada Con Exito", vbInformation, "Atención"
 End If

End Sub

Private Sub Command1_Click()
  nImpr = 46
  FrmImpresor.Show
End Sub

Private Sub Form_Load()
   DtpDesde.Value = Date
   DtpHasta.Value = Date
End Sub
