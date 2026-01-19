VERSION 5.00
Begin VB.Form HerraDatosEmpre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de la Empresa"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   7410
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   465
      Left            =   5715
      TabIndex        =   0
      Top             =   3585
      Width           =   1545
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
      Height          =   3480
      Left            =   60
      TabIndex        =   11
      Top             =   15
      Width           =   7230
      Begin VB.TextBox TxtPunto 
         Height          =   285
         Left            =   3345
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "1"
         Top             =   2940
         Width           =   450
      End
      Begin VB.CommandButton CmbCuit 
         Caption         =   "A"
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
         Left            =   3270
         TabIndex        =   22
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox TxtTelefono 
         Height          =   285
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "TxtTelefono"
         Top             =   2215
         Width           =   2895
      End
      Begin VB.TextBox TxtCopias 
         Height          =   300
         Left            =   1890
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "Text9"
         Top             =   2940
         Width           =   405
      End
      Begin VB.TextBox TxtInicio 
         Height          =   300
         Left            =   1875
         MaxLength       =   7
         TabIndex        =   8
         Text            =   "TxtInicio"
         Top             =   2583
         Width           =   885
      End
      Begin VB.TextBox TxtIB 
         Height          =   285
         Left            =   1875
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "TxtIB"
         Top             =   2215
         Width           =   1470
      End
      Begin VB.TextBox TxtCuit 
         Height          =   285
         Left            =   1875
         MaxLength       =   14
         TabIndex        =   5
         Text            =   "TxtCuit"
         Top             =   1845
         Width           =   1335
      End
      Begin VB.ComboBox CmbIva 
         Height          =   315
         Left            =   5070
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1845
         Width           =   2055
      End
      Begin VB.TextBox TxtLocalidad 
         Height          =   315
         Left            =   1875
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "TxtLocalidad"
         Top             =   1479
         Width           =   3360
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   300
         Left            =   1875
         TabIndex        =   3
         Text            =   "TxtDireccion"
         Top             =   1096
         Width           =   5250
      End
      Begin VB.TextBox TxtRazon 
         Height          =   300
         Left            =   1875
         TabIndex        =   2
         Text            =   "TxtRazon"
         Top             =   713
         Width           =   5250
      End
      Begin VB.TextBox TxtFantasia 
         Height          =   300
         Left            =   1875
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "TxtFantasia"
         Top             =   330
         Width           =   5250
      End
      Begin VB.Label Label4 
         Caption         =   "P.de Venta :"
         Height          =   240
         Left            =   2400
         TabIndex        =   24
         Top             =   2940
         Width           =   1140
      End
      Begin VB.Label Label17 
         Caption         =   "Teléfono :"
         Height          =   240
         Left            =   3420
         TabIndex        =   21
         Top             =   2215
         Width           =   810
      End
      Begin VB.Label Label12 
         Caption         =   "C.Postal/Localidad :"
         Height          =   225
         Left            =   195
         TabIndex        =   20
         Top             =   1485
         Width           =   1530
      End
      Begin VB.Label Label9 
         Caption         =   "Cant.Copias :"
         Height          =   240
         Left            =   195
         TabIndex        =   19
         Top             =   2940
         Width           =   1140
      End
      Begin VB.Label Label8 
         Caption         =   "Inicio de Actividades :"
         Height          =   255
         Left            =   195
         TabIndex        =   18
         Top             =   2583
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "Condición de Iva :"
         Height          =   270
         Left            =   3690
         TabIndex        =   17
         Top             =   1845
         Width           =   1350
      End
      Begin VB.Label Label6 
         Caption         =   "IIBB :"
         Height          =   240
         Left            =   195
         TabIndex        =   16
         Top             =   2215
         Width           =   600
      End
      Begin VB.Label Label5 
         Caption         =   "Cuit :"
         Height          =   210
         Left            =   195
         TabIndex        =   15
         Top             =   1860
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección :"
         Height          =   285
         Left            =   195
         TabIndex        =   14
         Top             =   1096
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Razón Social :"
         Height          =   285
         Left            =   195
         TabIndex        =   13
         Top             =   713
         Width           =   1350
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de Fantasia :"
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   330
         Width           =   1725
      End
   End
End
Attribute VB_Name = "HerraDatosEmpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbCuit_Click()
 On Error GoTo errHandler
  Dim Padron As Object, ok As Variant, Impuesto As Variant, Actividad As Variant
    
  If TxtCuit.text <> "" Then

  Dim obj As DsPadronA4
         
     Set obj = New DsPadronA4
     Dim Resultado As Boolean
          
     Resultado = obj.SolicitarCuit(TxtCuit.text)
     
     If Resultado = True Then
        TxtRazon.text = obj.Nombre
        TxtDireccion.text = obj.Direccion
        TxtCuit.text = obj.Cuit
        TxtLocalidad.text = obj.CodPostal & " - " & obj.Localidad
        CmbIva.text = obj.CategoriaIva
        TxtInicio.text = obj.Inicio
     Else
        MsgBox "El Cuit solicitado No Corresponde a una Persona Física" & Chr(10) & Chr(13), vbInformation, "Atención"
        CmbCuit.Enabled = True
        Exit Sub
     End If
     If obj.Resultado = "OK" Then
        MsgBox obj.Nombre & " " & obj.Estado & vbCrLf & obj.Direccion & vbCrLf _
                                   & obj.Localidad & vbCrLf & obj.Provincia & _
                                   vbCrLf & obj.Actividad, vbInformation, _
                                   "Resultado CUIT " & obj.Cuit & " (online AFIP)"
      Else
         ' respuesta del servidor (para depuración)
          Debug.Print Padron.response
          MsgBox "Error AFIP: " & obj.Resultado
          TxtCuit.text = ""
          TxtCuit.SetFocus
     End If
    Else
       MsgBox "Debe Ingresar un Cuit Válido", vbCritical, "Atención"
   End If
   CmbCuit.Enabled = True
 
Exit Sub
    
errHandler:
   ManejaErrores
End Sub

Private Sub CmdAplicar_Click()
   Dim sSql As String
  On Error GoTo errHandler
  
  db.Open sDb
  db.BeginTrans
  
    sSql = "UPDATE AfipDatos SET Fantasia='" & TxtFantasia.text & "',Empresa='" & Replace(TxtRazon.text, "'", " ") & "',Direccion='" & TxtDireccion.text & "',"
    sSql = sSql & "Localidad='" & TxtLocalidad.text & "',Cuit='" & TxtCuit.text & "',IIBB='" & TxtIB.text & "',Iva='" & CmbIva.text & "',"
    sSql = sSql & "Inicio='" & TxtInicio.text & "',Copias=" & TxtCopias.text & ",Id_Datos=" & TxtPunto.text
    db.Execute sSql
            
    sSql = "UPDATE Sistema SET Fantasia='" & TxtFantasia.text & "',Direccion='" & TxtDireccion.text & "',Telefono='" & TxtTelefono.text & "',Ciudad='" & TxtLocalidad.text & "',"
    sSql = sSql & "Cuit='" & Replace(TxtCuit.text, "-", "") & "'"
    db.Execute sSql
            
    sSql = "UPDATE Empresa SET Razon='" & Replace(TxtRazon.text, "'", " ") & "',Direccion='" & TxtDireccion.text & "',Localidad='" & TxtLocalidad.text & "',Cuit='" & TxtCuit.text & "'"
    db.Execute sSql
    
    sSql = "UPDATE Comprobantes Set Sucursal=" & TxtPunto.text & " WHERE     (Afip <> 0) AND (TipoOperacion = 5) AND (Afip < 20) OR (Afip <> 0) AND (TipoOperacion = 1) AND (Afip < 20)"
    db.Execute sSql
  
  db.CommitTrans
  db.Close
  
  MsgBox "Deberá Reiniciar el Sistema para Aplicar los Cambios", vbInformation, "Atención"
  
  End
  
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
End Sub

Private Sub Form_Load()
  CargarCombos
  LinkearTexto
End Sub

Private Sub CargarCombos()
   Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbIva, "TipoIva", "TipoIva", "Descripcion", ""

End Sub

Private Sub LinkearTexto()
  Dim cRsl As ClsLectura, Rx As Recordset
  
  Set cRsl = New ClsLectura
  
  Set Rx = cRsl.TraerTodos("AfipDatos", "*")

  TxtFantasia.text = "" & Rx!Fantasia
  TxtRazon.text = Rx!Empresa
  TxtDireccion.text = Rx!Direccion
  TxtTelefono.text = "" & sTelFan
  TxtLocalidad.text = Rx!Localidad
  TxtCuit.text = Rx!Cuit
  TxtIB.text = Rx!IIBB
  CmbIva.text = Rx!Iva
  TxtInicio.text = Rx!Inicio
  TxtCopias.text = Rx!copias
  TxtPunto.text = Rx!Id_datos
End Sub

