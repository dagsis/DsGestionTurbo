VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasInfRetIb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Retención de IB"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   3840
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
      Height          =   2250
      Left            =   180
      TabIndex        =   1
      Top             =   30
      Width           =   3510
      Begin VB.CheckBox ChkExportar 
         Caption         =   "Generar Txt para Arba"
         Height          =   225
         Left            =   225
         TabIndex        =   6
         Top             =   1905
         Width           =   2160
      End
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   450
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52690945
         CurrentDate     =   36983
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   345
         Left            =   1140
         TabIndex        =   3
         Top             =   1185
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52690945
         CurrentDate     =   36983
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   435
         TabIndex        =   5
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   420
         TabIndex        =   4
         Top             =   1185
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Aplicar"
      Height          =   420
      Left            =   2730
      TabIndex        =   0
      Top             =   2310
      Width           =   945
   End
End
Attribute VB_Name = "ComprasInfRetIb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdImprimir_Click()
 If ChkExportar.Value = 0 Then
     nImpr = 63
     FrmImpresor.Show
  Else
     ExportarTxt
  End If
End Sub

Private Sub Form_Load()
  DTPDesde.Value = Date
  DTPHasta.Value = Date
End Sub

Private Sub ExportarTxt()
  Dim sCuit As String, sFecha As String, sTipo As String, sLetra As String, sSuc As String, sNum As String, sMonto As String, sPercep As String, sOpe As String
  Dim Rx As Recordset, cRre As ClsComprobantesL, sTotal As String
  
  Set cRre = New ClsComprobantesL
  
  Set Rx = cRre.TraerInfRetIBex(DTPDesde.Value, DTPHasta.Value)


  If Rx.RecordCount <> 0 Then
        Dim sNom As String, sVariable As String
        Dim sNomDet As String, sVariableDet As String
        
        sNom = CurDir() & "\ArbaRetencion\CprCab1" & Format(Day(Date), "00") & "-" & Format(Month(Date), "00") & "-" & Year(DTPDesde.Value) & "-1421.txt"
          
        Open sNom For Output As #1
        Do While Not Rx.EOF
           sCuit = Left(Rx!Cuit, 2) & "-" & Mid(Rx!Cuit, 3, 8) & "-" & Right(Rx!Cuit, 1)
           sFecha = Rx!FechaRec
           sSuc = Format(Rx!Sucursal, "0000")
           sNum = Format(Rx!Numero, "00000000")
           sPercep = Format(Rx!importe, "00000000.00")
           sOpe = "A"
           sTotal = sCuit & sFecha & sSuc & sNum & sPercep & sOpe
           Print #1, sTotal
           Rx.MoveNext
        Loop
        Close #1
        MsgBox "Archivos de Exportación Para Arba Creados Con Exito", vbExclamation, "Atención"
  End If
End Sub


