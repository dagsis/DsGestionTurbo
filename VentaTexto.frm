VERSION 5.00
Begin VB.Form VentaTexto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Texto del Comprobante"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   0
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2415
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Height          =   690
      Index           =   1
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2415
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Texto"
      Height          =   2220
      Left            =   150
      TabIndex        =   4
      Top             =   135
      Width           =   6060
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   195
         MaxLength       =   60
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1380
         Width           =   5565
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   195
         MaxLength       =   60
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1055
         Width           =   5565
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   195
         MaxLength       =   60
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   730
         Width           =   5565
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   195
         MaxLength       =   60
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   405
         Width           =   5565
      End
   End
End
Attribute VB_Name = "VentaTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset, pMovi As Long


Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
            Unload Me
         Case 1
            Grabar
            CmdBotones(1).Enabled = False
  End Select
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
 CmdBotones(0).Picture = LoadResPicture("Salir", 0)
 CmdBotones(1).Picture = LoadResPicture("Grabar", 0)
 pMovi = VentaComprobantes.nMovi
 LlenarRs
 CmdBotones(1).Enabled = False
End Sub

Private Sub LlenarRs()
  Dim cRsl As ClsLectura, sCli As String, nTra As Integer
  
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRsCondi("CabTexto", "Id", "Movimiento=" & pMovi)
    
  Text1.text = ""
  Text2.text = ""
  Text3.text = ""
  Text4.text = ""
  
  If Rs.RecordCount = 0 Then
     Rs.AddNew
     Rs!Movimiento = pMovi
     Rs!Texto1 = Text1.text
     Rs!Texto2 = Text2.text
     Rs!Texto3 = Text3.text
     Rs!Texto4 = Text4.text
     Rs.Update
  End If

  Text1.text = "" & Rs!Texto1
  Text2.text = "" & Rs!Texto2
  Text3.text = "" & Rs!Texto3
  Text4.text = "" & Rs!Texto4
  
  Set cRsl = Nothing

End Sub

Private Sub Grabar()
  Dim cRle As ClsComprobantesE
  
  Set cRle = New ClsComprobantesE
  Rs!Movimiento = pMovi
  Rs!Texto1 = Text1.text
  Rs!Texto2 = Text2.text
  Rs!Texto3 = Text3.text
  Rs!Texto4 = Text4.text
  Rs!TextoFiscal = ""
  Rs.Update
  
  cRle.GrabarTexto Rs, pMovi
  LlenarRs
  Set cRle = Nothing
End Sub

Private Sub TodoBien()
  CmdBotones(1).Enabled = False
  If Len(Text1.text) <> 0 Then
     CmdBotones(1).Enabled = True
  End If
End Sub

Private Sub Text1_Change()
  TodoBien
End Sub

Private Sub Text2_Change()
  TodoBien
End Sub

Private Sub Text3_Change()
  TodoBien
End Sub

Private Sub Text4_Change()
  TodoBien
End Sub
