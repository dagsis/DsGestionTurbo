VERSION 5.00
Begin VB.Form MenuMesa 
   Caption         =   "Mesa"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnumesa 
      Caption         =   "Mesa"
      Begin VB.Menu mnuabrir 
         Caption         =   "Abrir Mesa"
      End
      Begin VB.Menu MnuCopiar 
         Caption         =   "Copiar Mesa"
      End
      Begin VB.Menu MnuPendiente 
         Caption         =   "Cerrar Mesa"
      End
      Begin VB.Menu CierreRapido 
         Caption         =   "Cierre Rápido"
      End
      Begin VB.Menu MenuAnular 
         Caption         =   "Anular"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mncobranza 
         Caption         =   "Cobranza"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "MenuMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub MenuAbrir_Click()
  Dim ok As Boolean, Mesas As ClsPrograma
    
   Set Mesas = New ClsPrograma
   
   ok = Mesas.DameEstadoMesa("VACIA", nMesa)
   If ok Then
      AbrirMesa.Show 1
   End If
   
End Sub

Private Sub MenuAnular_Click()
  Dim cRle As ClsPrograma
  
  Set cRle = New ClsPrograma
  If MsgBox("Anula la Mesa elegida  ?", 20, "Atención") = 6 Then
     cRle.AnularMesa
     FrmAdicion.CmdMesa(nMesa).BackColor = &HC000&
  End If
End Sub

Private Sub MnuCopiar_Click()
    CopiarMesa.Show 1
End Sub
