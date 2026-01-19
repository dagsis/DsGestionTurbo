VERSION 5.00
Begin VB.Form BancoCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques en Cartera"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   240
      TabIndex        =   1
      Top             =   390
      Width           =   8610
   End
   Begin VB.CommandButton CmdSeleccionar 
      Caption         =   "Salir"
      Height          =   390
      Left            =   7560
      TabIndex        =   0
      Top             =   3555
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "   Id               Banco               Número                Titular        F.Acreditación          Importe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   75
      Width           =   8565
   End
End
Attribute VB_Name = "BancoCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset

Private Sub CmdSeleccionar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Dim cRx As ClsLectura, sBan As String, sTitu As String, n As Integer, sTotal As String
   
   Set cRx = New ClsLectura
   
   Set Rs = cRx.TraerRsCondi("Cheques", "FechaAcre", "Anulado=0 AND FechaSalida is Null and Suc=" & nSucursal)
   Do While Not Rs.EOF
      sBan = Left(Trim(cRx.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & Rs!Banco)), 20)
      If sBan <> "0" Then
         n = InStr(1, Rs!Titular, "-") - 2
         If n <= 0 Then n = 25
            sTitu = Left(Rs!Titular, n)
            sTitu = Left(sTitu, 15)
            sTotal = Format(Rs!ID, "000000") & " " & sBan & Space(20 - Len(sBan)) & " " & Rs!Numero & Space(12 - Len(Rs!Numero)) & sTitu & Space(15 - Len(sTitu)) & Rs!FechaAcre & " " & Space(14 - Len(Format(Rs!importe, nCantDecimales))) & Format(Rs!importe, nCantDecimales)
            List1.AddItem sTotal
            List1.ItemData(List1.NewIndex) = Rs!ID
      End If
      Rs.MoveNext
   Loop
End Sub

Private Sub List1_DblClick()
   If List1.ListIndex <> -1 Then
      nDat = List1.ItemData(List1.ListIndex)
   End If
   Unload Me
End Sub
