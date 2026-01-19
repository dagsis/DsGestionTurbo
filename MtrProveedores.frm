VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MtrProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Proveedores"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   7665
   Begin VB.Frame FraCliente 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   315
      TabIndex        =   24
      Top             =   540
      Width           =   5895
      Begin VB.Frame FraDireccion 
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   5895
         Begin VB.Frame FraImpuesto 
            Caption         =   "Impuestos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3165
            Left            =   -15
            TabIndex        =   42
            Top             =   0
            Width           =   5895
            Begin VB.CommandButton CmdRetencion 
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
               Height          =   345
               Left            =   2490
               TabIndex        =   62
               ToolTipText     =   "Traer datos Afip"
               Top             =   2025
               Width           =   255
            End
            Begin VB.CommandButton CmdAdquirir 
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
               Height          =   375
               Left            =   2760
               TabIndex        =   61
               ToolTipText     =   "Traer datos Afip"
               Top             =   960
               Width           =   255
            End
            Begin VB.CommandButton CmdContancia 
               Caption         =   "C"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               TabIndex        =   60
               ToolTipText     =   "Constancia de Cuit"
               Top             =   480
               Width           =   255
            End
            Begin VB.ComboBox CmbGanancia 
               Height          =   315
               Left            =   1470
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   2415
               Width           =   2460
            End
            Begin VB.TextBox TxtRetencion 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1470
               MaxLength       =   7
               TabIndex        =   48
               Text            =   "TxtRete"
               Top             =   2040
               Width           =   930
            End
            Begin VB.Frame Frame1 
               Caption         =   "Precios de Venta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Left            =   3270
               TabIndex        =   56
               Top             =   375
               Width           =   2355
               Begin VB.TextBox TxtPorc 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1200
                  MaxLength       =   7
                  TabIndex        =   51
                  Top             =   495
                  Width           =   930
               End
               Begin VB.Label Label15 
                  Caption         =   "Porcentaje :"
                  Height          =   210
                  Left            =   150
                  TabIndex        =   57
                  Top             =   495
                  Width           =   945
               End
            End
            Begin VB.CheckBox ChkCtaCte 
               Caption         =   "Cta Cte Bancaria"
               Height          =   210
               Left            =   180
               TabIndex        =   50
               Top             =   2850
               Width           =   1605
            End
            Begin VB.TextBox TxtGanancias 
               Height          =   285
               Left            =   1470
               MaxLength       =   15
               TabIndex        =   47
               Text            =   "TxtGanancias"
               Top             =   1680
               Width           =   1200
            End
            Begin VB.TextBox TxtIBrutos 
               Height          =   285
               Left            =   1470
               MaxLength       =   15
               TabIndex        =   46
               Text            =   "TxtIBrutos"
               Top             =   1305
               Width           =   1200
            End
            Begin VB.ComboBox CmbIva 
               Height          =   315
               Left            =   210
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   465
               Width           =   2460
            End
            Begin VB.TextBox TxtCuit 
               Height          =   285
               Left            =   1470
               MaxLength       =   13
               TabIndex        =   44
               Text            =   "TxtCuit"
               Top             =   915
               Width           =   1200
            End
            Begin VB.TextBox TxtId 
               Height          =   315
               Left            =   1350
               TabIndex        =   43
               Text            =   "TxtId"
               Top             =   -195
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label21 
               Caption         =   "R.Ganancias :"
               Height          =   210
               Left            =   210
               TabIndex        =   59
               Top             =   2415
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "Retención IB :"
               Height          =   210
               Left            =   195
               TabIndex        =   58
               Top             =   2040
               Width           =   1125
            End
            Begin VB.Label Label19 
               Caption         =   "Nº Ganancias :"
               Height          =   270
               Left            =   225
               TabIndex        =   55
               Top             =   1680
               Width           =   1080
            End
            Begin VB.Label Label18 
               Caption         =   "Nº I.B. / C.M. :"
               Height          =   255
               Left            =   210
               TabIndex        =   54
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Label Label16 
               Caption         =   "C.de Iva :"
               Height          =   270
               Left            =   210
               TabIndex        =   53
               Top             =   240
               Width           =   825
            End
            Begin VB.Label Label17 
               Caption         =   "C.U.I.T :"
               Height          =   195
               Left            =   210
               TabIndex        =   52
               Top             =   915
               Width           =   735
            End
         End
         Begin VB.TextBox TxtLocalidad 
            Height          =   285
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "TxtLocalidad"
            Top             =   825
            Width           =   2730
         End
         Begin VB.TextBox TxtDomicilio 
            Height          =   300
            Left            =   1245
            MaxLength       =   45
            TabIndex        =   6
            Text            =   "TxtDomicilio"
            Top             =   420
            Width           =   4215
         End
         Begin VB.TextBox TxtPostal 
            Height          =   300
            Left            =   1245
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "TxtPostal"
            Top             =   831
            Width           =   570
         End
         Begin VB.ComboBox CmbProvincia 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1242
            Width           =   1815
         End
         Begin VB.ComboBox CmbPais 
            Height          =   315
            Left            =   3690
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1242
            Width           =   1800
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   300
            Left            =   1245
            MaxLength       =   17
            TabIndex        =   11
            Text            =   "TxtTelefono"
            Top             =   1668
            Width           =   1635
         End
         Begin VB.TextBox TxtFax 
            Height          =   300
            Left            =   3825
            MaxLength       =   17
            TabIndex        =   12
            Text            =   "TxtFax"
            Top             =   1668
            Width           =   1635
         End
         Begin VB.TextBox TxtEmail 
            Height          =   300
            Left            =   1245
            MaxLength       =   45
            TabIndex        =   13
            Text            =   "TxtEmail"
            Top             =   2079
            Width           =   4230
         End
         Begin VB.Label Label7 
            Caption         =   "Domicilio :"
            Height          =   300
            Left            =   375
            TabIndex        =   33
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Localidad :"
            Height          =   240
            Left            =   1905
            TabIndex        =   32
            Top             =   825
            Width           =   900
         End
         Begin VB.Label Label9 
            Caption         =   "C.Postal :"
            Height          =   225
            Left            =   375
            TabIndex        =   31
            Top             =   825
            Width           =   750
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia :"
            Height          =   240
            Left            =   390
            TabIndex        =   30
            Top             =   1245
            Width           =   930
         End
         Begin VB.Label Label11 
            Caption         =   "Pais :"
            Height          =   240
            Left            =   3105
            TabIndex        =   29
            Top             =   1260
            Width           =   525
         End
         Begin VB.Label Label12 
            Caption         =   "Teléfono :"
            Height          =   195
            Left            =   375
            TabIndex        =   28
            Top             =   1665
            Width           =   915
         End
         Begin VB.Label Label13 
            Caption         =   "Fax :"
            Height          =   300
            Left            =   2955
            TabIndex        =   27
            Top             =   1665
            Width           =   510
         End
         Begin VB.Label Label14 
            Caption         =   "E-Mail :"
            Height          =   210
            Left            =   375
            TabIndex        =   26
            Top             =   2100
            Width           =   720
         End
      End
      Begin MSComCtl2.DTPicker DTPBaja 
         Height          =   330
         Left            =   4155
         TabIndex        =   5
         Top             =   2400
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   119209985
         CurrentDate     =   36768
      End
      Begin MSComCtl2.DTPicker DTPAlta 
         Height          =   330
         Left            =   1695
         TabIndex        =   4
         Top             =   2400
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   119209985
         CurrentDate     =   36768
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   540
         Left            =   1710
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "MtrProveedores.frx":0000
         Top             =   1725
         Width           =   3720
      End
      Begin VB.TextBox TxtContacto 
         Height          =   285
         Left            =   1710
         MaxLength       =   40
         TabIndex        =   2
         Text            =   "TxtContacto"
         Top             =   1335
         Width           =   3720
      End
      Begin VB.TextBox TxtRazon 
         Height          =   285
         Left            =   1710
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "TxtRazon"
         Top             =   975
         Width           =   3720
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   285
         Left            =   1710
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "TxtCodigo"
         Top             =   570
         Width           =   1500
      End
      Begin VB.CheckBox ChkActivo 
         Caption         =   "Activo"
         Height          =   270
         Left            =   4635
         TabIndex        =   34
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   405
         TabIndex        =   40
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Razón Social :"
         Height          =   255
         Left            =   420
         TabIndex        =   39
         Top             =   935
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones :"
         Height          =   285
         Left            =   420
         TabIndex        =   38
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Alta :"
         Height          =   240
         Left            =   420
         TabIndex        =   37
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Baja :"
         Height          =   240
         Left            =   3120
         TabIndex        =   36
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Contacto :"
         Height          =   330
         Left            =   420
         TabIndex        =   35
         Top             =   1275
         Width           =   1125
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   780
      Index           =   0
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   780
      Index           =   1
      Left            =   1147
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   780
      Index           =   2
      Left            =   2129
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Buscar"
      Height          =   780
      Index           =   3
      Left            =   3111
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   780
      Index           =   4
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4050
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   5
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   420
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   6
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1275
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
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
      Height          =   780
      Index           =   7
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2115
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   8
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2970
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   780
      Index           =   9
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4050
      Width           =   915
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3765
      Left            =   150
      TabIndex        =   41
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6641
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificación"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dirección"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Impuesto"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MtrProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCod As Integer, sw As Boolean, nId As Integer
Dim sCod1 As String
Dim Rs As ADODB.Recordset, Rs_Pro As ADODB.Recordset
Dim cRsl As ClsLectura
Dim cRse As ClsEscritura
Dim P As ClsPrograma

Private Sub CmbGanancia_Click()
  TodoBien
End Sub

Private Sub CmbIva_Click()
  TodoBien
End Sub


Private Sub ChkActivo_Click()
  If sw = False Then
     DTPBaja.Enabled = False
     If ChkActivo.Value = 0 Then
        DTPBaja.Enabled = True
     End If
  End If
End Sub


Private Sub CmdAdquirir_Click()
    On Error GoTo errHandler
  Dim Padron As Object, ok As Variant, Impuesto As Variant, Actividad As Variant
  
  CmdAdquirir.Enabled = False
  
  If TxtCuit.Text <> "" Then

     Dim obj As DsPadronA4
         
     Set obj = New DsPadronA4
     Dim Resultado As Boolean
     
     Set cRsl = New ClsLectura
         
     Resultado = obj.SolicitarCuit(TxtCuit.Text)
     
     If Resultado = True Then
        If cRsl.TraerCantidad("Proveedores", "Cuit", "Cuit='" & TxtCuit.Text & "'") = 0 Or CmdBotones(2).Caption = "Grabar" Then
            TxtRazon.Text = obj.Nombre
            TxtDomicilio.Text = obj.Direccion
            TxtCuit.Text = obj.Cuit
            TxtPostal.Text = obj.CodPostal
            TxtLocalidad.Text = obj.Localidad
            CmbProvincia.Text = obj.Provincia
            CmbIva.Text = obj.CategoriaIva
        Else
           MsgBox "El Cuit Ingresado ya existe en la Base de Datos", vbCritical, "Atención"
           TxtCuit.Text = ""
           CmdAdquirir.Enabled = True
           Exit Sub
       End If
     Else
        MsgBox "El Cuit solicitado No Corresponde a una Persona Física" & Chr(10) & Chr(13), vbInformation, "Atención"
        CmdAdquirir.Enabled = True
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
          TxtCuit.Text = ""
          TxtCuit.SetFocus
     End If
    Else
       MsgBox "Debe Ingresar un Cuit Válido", vbCritical, "Atención"
   End If
   CmdAdquirir.Enabled = True
Exit Sub
    
errHandler:
   ManejaErrores
End Sub

Private Sub CmdContancia_Click()
' On Error GoTo errHandler
'   Dim Padron As Object, ok As Variant, Impuesto As Variant, actividad As Variant
'
'   If TxtCuit.text <> "" Then
'    ' Crear la interfaz COM
'        Set Padron = CreateObject("PadronAFIP")
'
'        ok = Padron.Conectar()
'
'        ok = Padron.DescargarConstancia(TxtCuit.text, "constancia.pdf")
'        If Padron.Excepcion = "" Then
'           Padron.MostrarPDF "constancia.pdf", False
'        Else
'          Debug.Print Padron.response
'          MsgBox "Error AFIP: " & Padron.Excepcion, vbCritical, "Resultado CUIT " & TxtCuit.text & " (online)"
'        End If
'    Else
'       MsgBox "Debe Ingresar un Cuit Válido", vbCritical, "Atención"
'       TxtCuit.text = ""
'       TxtCuit.SetFocus
'    End If
'Exit Sub
'
'errHandler:
'   ManejaErrores
End Sub

Private Sub CmdRetencion_Click()
 Dim obj As DsPadronA4
   
   Set obj = New DsPadronA4
   
    If TxtCuit.Text <> "" Then
       Dim Resultado As Boolean
       CmdRetencion.Enabled = False
       Resultado = obj.SolicitarIB(TxtCuit.Text)
     
       If Resultado = True Then
           If obj.Resultado = "Ok" Then
              TxtRetencion.Text = Replace(obj.Retencion, ",", ".")
           Else
              MsgBox obj.Mensaje, vbCritical, "Atención"
           End If
       End If
       CmdRetencion.Enabled = True
    Else
       MsgBox "El Campo CUIT esta vacio...", vbCritical, "Atención"
    End If
End Sub

Private Sub Form_Load()
  On Error GoTo errHandler

  Set Rs = New ADODB.Recordset
  Set Rs_Pro = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set P = New ClsPrograma
  Set cRse = New ClsEscritura

  VerFrame True, False, False
  sw = True
  Limpiar
  Habilitar False
  CargarCombos
  Botones pEdit3, False, False, False, False, False, False, False, False, True
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
  CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
  CmdBotones(3).Picture = LoadResPicture("Buscar", 0)
  CmdBotones(4).Picture = LoadResPicture("Imprimir", 0)
  CmdBotones(9).Picture = LoadResPicture("Salir", 0)
  
  ArmaRsParaMovimiento
  
  
  Me.Top = 0
  Me.Left = 0
  
Exit Sub

errHandler:
   ManejaErrores
 
End Sub

Private Sub Form_Resize()
  If CmdBotones(0).Enabled = True Then
     CmdBotones(0).SetFocus
  End If
End Sub

Private Sub TabStrip1_Click()
 Select Case TabStrip1.SelectedItem.Index
        Case 1
             VerFrame True, False, False
        Case 2
             VerFrame True, True, False
        Case 3
             VerFrame True, True, True
 End Select
End Sub

Private Sub VerFrame(bIden As Boolean, bDomi As Boolean, bImp As Boolean)
  FraCliente.Visible = bIden
  FraDireccion.Visible = bDomi
  FraImpuesto.Visible = bImp
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
   KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      End If
  Next ctl
  For Each ctl In Controls
      If TypeOf ctl Is ComboBox Then
         ctl.ListIndex = -1
      End If
  Next ctl
  ChkCtaCte.Value = 0
  TxtPorc.Text = "0.00"
  DTPAlta.Value = Date
  DTPBaja.Value = Date
  TxtRetencion.Text = "0.00"
End Sub

Private Sub Habilitar(bEstado As Boolean)
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
         Or TypeOf ctl Is ComboBox Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  CmdAdquirir.Enabled = bEstado
  CmdRetencion.Enabled = bEstado
End Sub

Private Sub Botones(bNue As Boolean, bBor As Boolean, bMod As Boolean, bBus As Boolean, _
 bImp As Boolean, bPri As Boolean, bSig As Boolean, bAnt As Boolean, bUlt As Boolean, _
 bSal As Boolean)
   CmdBotones(0).Enabled = bNue
   CmdBotones(1).Enabled = bBor
   CmdBotones(2).Enabled = bMod
   CmdBotones(3).Enabled = bBus
   CmdBotones(4).Enabled = bImp
   CmdBotones(5).Enabled = bPri
   CmdBotones(6).Enabled = bSig
   CmdBotones(7).Enabled = bAnt
   CmdBotones(8).Enabled = bUlt
   CmdBotones(9).Enabled = bSal
 End Sub

Private Sub CargarCombos()
  cRsl.CargaCombo CmbPais, "Paises", "Pais", "Descripcion", ""
  cRsl.CargaCombo CmbProvincia, "Provincias", "Provincia", "Descripcion", ""
  cRsl.CargaCombo CmbIva, "TipoIva", "TipoIva", "Descripcion", ""
  cRsl.CargaCombo CmbGanancia, "TblGanancia", "Id", "Descripcion", ""
  CmbGanancia.Text = "Exento"
  
End Sub
Private Sub LinkearTexto()
  ChkActivo.Value = 1
  If Not IsNull(Rs!FechaBaja) Then
     ChkActivo.Value = 0
     DTPBaja.Value = Rs!FechaBaja
  End If
  TxtId.Text = Rs!ID
  TxtCodigo.Text = Rs!Proveedor
  TxtRazon.Text = Rs!RazonSocial
  TxtContacto.Text = "" & Rs!Contacto
  TxtObservaciones.Text = "" & Rs!Observaciones
  TxtDomicilio.Text = "" & Rs!Domicilio
  TxtPostal.Text = "" & Rs!CodigoPostal
  TxtTelefono.Text = "" & Rs!Telefono
  TxtFax.Text = "" & Rs!Fax
  TxtEmail.Text = "" & Rs!email
  TxtCuit.Text = "" & Rs!Cuit
  DTPAlta.Value = Rs!FechaAlta
  
  TxtLocalidad.Text = "" & Rs!Localidad
  If Not IsNull(Rs!Pais) Then P.SetComboByItemData CmbPais, CLng(Rs!Pais)
  If Not IsNull(Rs!Provincia) Then P.SetComboByItemData CmbProvincia, CLng(Rs!Provincia)
  If Not IsNull(Rs!TipoIva) Then P.SetComboByItemData CmbIva, CLng(Rs!TipoIva)
  ChkCtaCte.Value = IIf(Rs!ctacte = True, 1, 0)
  TxtPorc.Text = Format(Rs!Porcentaje, "#0.00")
  TxtRetencion.Text = Format(Rs!Retencion, "#0.00")
  If Not IsNull(Rs!CatGanancia) Then P.SetComboByItemData CmbGanancia, CLng(Rs!CatGanancia)
End Sub
Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Borrar
         Case 2
              Modificar
         Case 3
              Buscar
         Case 5
              Primero
         Case 6
              Anterior
         Case 7
              Siguiente
         Case 8
              ultimo
         Case 9
              Salir
  End Select
End Sub

Private Sub Nuevo()
  On Error GoTo Errores
  
  If sw = True Then
     Habilitar True
     Botones False, False, False, False, False, False, False, False, False, True
     Limpiar
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(9).Caption = "Cancelar"
     CmdBotones(9).Picture = LoadResPicture("Cancelar", 0)
     If bAutoNumera = True Then
        TxtCodigo.Text = "P" & cRsl.ActualizarCampoTraer("Proveedor")
        TxtRazon.SetFocus
     Else
        TxtCodigo.SetFocus
     End If
     sw = False
     ChkActivo.Value = 1
     Dim RsReg As Recordset
     Set RsReg = cRsl.TraerRsCondi("Registros", "Usuario", "Usuario=" & nUsuario & " And Suc = " & nSucursal)
         
     P.SetComboByItemData CmbProvincia, RsReg!Provincia ' CmbProvincia.text = cRsl.DatoCombo("Provincias", "Provincia", "Descripcion", cRsl.RegPorDefecto("Provincia"))
     P.SetComboByItemData CmbPais, 1 ' CmbPais.text = cRsl.DatoCombo("Paises", "Pais", "Descripcion", 1)
     P.SetComboByItemData CmbIva, RsReg!Impuesto 'CmbIva.text = cRsl.DatoCombo("TipoIva", "TipoIva", "Descripcion", cRsl.RegPorDefecto("Impuesto"))
     CmbGanancia.Text = "Exento"
     DTPBaja.Enabled = False
  Else
     Botones True, True, True, True, True, True, True, True, True, True
     Habilitar False
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(9).Caption = "Salir"
     CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     sw = True
     Grabar
     ArmaRsParaMovimiento
  End If
  Exit Sub
Errores:
   ManejaErrores
End Sub

Private Sub Salir()
 On Error GoTo errHandler

  Botones True, False, False, False, False, False, False, False, False, True
  If sw = False Then
     sw = True
     Limpiar
     If CmdBotones(0).Caption = "Grabar" Then
        Habilitar False
        CmdBotones(0).Caption = "Nuevo"
        CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
        CmdBotones(9).Caption = "Salir"
        CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     Else
        Habilitar False
        CmdBotones(2).Caption = "Modificar"
        CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
        CmdBotones(9).Caption = "Salir"
        CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     End If
      If Rs_Pro.BOF = False And Rs_Pro.EOF = False Then
        Botones True, True, True, True, True, True, True, True, True, True
        LinkearTexto
     End If
  Else
     Unload Me
  End If
  Exit Sub
errHandler:
   ManejaErrores
End Sub
Private Sub Modificar()
  On Error GoTo Errores
  
  Static rList As Integer
  
  If sw = True Then
     sw = False
     sCod1 = Rs!Proveedor
     Botones False, False, True, False, False, False, False, False, False, True
     Habilitar True
     CmdBotones(2).Caption = "Grabar"
     CmdBotones(2).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(9).Caption = "Cancelar"
     CmdBotones(9).Picture = LoadResPicture("Cancelar", 0)
     DTPBaja.Enabled = False
  Else
     Botones True, True, True, True, True, True, True, True, True, True
     Habilitar False
     Grabar
     CmdBotones(2).Caption = "Modificar"
     CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
     CmdBotones(9).Caption = "Salir"
     CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     sw = True
  End If
Exit Sub
Errores:
   ManejaErrores

End Sub
Private Sub Borrar()
  On Error GoTo errHandler
    
  If cRsl.TraerCantidad("CabComprobantes", "Cliente", "Cliente='" & TxtCodigo.Text & "'") = 0 And cRsl.TraerCantidad("Registros", "Proveedor", "Proveedor=" & Val(TxtId.Text)) = 0 Then
    If MsgBox("Borra el Proveedor Seleccionado ?", 20, "Atención") = 6 Then
       Rs_Pro.Delete
       cRse.Actualizar "ProveedorBorrar", , TxtId.Text
       Limpiar
       Set cRse = Nothing
       If Rs_Pro.RecordCount <> 0 Then
          Anterior
       Else
          Botones True, False, False, False, False, False, False, False, False, True
       End If
    End If
  Else
     MsgBox "Proveedor Asociado a una Factura de Compra o En Proveedor por Defecto", vbCritical, "Atención"
  End If
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Grabar()
  On Error GoTo errHandler
  
  Dim Rs1 As ADODB.Recordset
  Dim i As Integer, cRsp As ClsClienteE
      
  Set cRsp = New ClsClienteE
    
  Set Rs1 = cRsl.RsVacio("Proveedores", "Proveedor", "S")
  
  Rs1.AddNew
'  Rs1!Id = 0
  Rs1!Proveedor = Replace(TxtCodigo.Text, "'", "")
  Rs1!RazonSocial = Replace(TxtRazon.Text, "'", "")
  Rs1!Contacto = Replace(TxtContacto.Text, "'", "")
  Rs1!Observaciones = Replace(TxtObservaciones.Text, "'", "")
  Rs1!FechaAlta = DTPAlta.Value
  Rs1!FechaBaja = IIf(ChkActivo.Value = 0, DTPBaja.Value, Null)
  Rs1!Domicilio = TxtDomicilio.Text
  Rs1!CodigoPostal = IIf(TxtPostal.Text = "", 0, TxtPostal.Text)
  
  Rs1!Localidad = TxtLocalidad.Text
  
  If CmbProvincia.ListIndex <> -1 Then
     Rs1!Provincia = CmbProvincia.ItemData(CmbProvincia.ListIndex)
  End If
  
  If CmbPais.ListIndex <> -1 Then
     Rs1!Pais = CmbPais.ItemData(CmbPais.ListIndex)
  End If

  Rs1!Telefono = TxtTelefono.Text
  Rs1!Fax = TxtFax.Text
  
  If CmbIva.ListIndex <> -1 Then
     Rs1!TipoIva = CmbIva.ItemData(CmbIva.ListIndex)
  End If
  
  Rs1!Cuit = TxtCuit.Text
  
  
  Rs1!email = TxtEmail.Text
  Rs1!IngBrutos = TxtIBrutos.Text
  Rs1!Ganancias = TxtGanancias.Text
  Rs1!ctacte = IIf(ChkCtaCte.Value = 1, True, False)
  Rs1!Porcentaje = TxtPorc.Text
  Rs1!Retencion = TxtRetencion.Text
  
  If CmbGanancia.ListIndex <> -1 Then
     Rs1!CatGanancia = CmbGanancia.ItemData(CmbGanancia.ListIndex)
  End If
    
  Rs1.Update
        
  If CmdBotones(2).Caption = "Grabar" Then
     cRsp.Modificar Rs1, sCod1, 1
  Else
     cRsp.Agregar Rs1, 1
  End If
  
  Rs1.Close
  
  Set cRse = Nothing
  bProveedores = True
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub ArmaRsParaMovimiento()
  Set Rs_Pro = cRsl.TraerTodos("Proveedores", "Id")
  Botones pEdit3, False, False, False, False, False, False, False, False, True
  If Rs_Pro.BOF = False And Rs_Pro.EOF = False Then
     Rs_Pro.MoveLast
     Set Rs = cRsl.TraerRS("ProveedorTraerUno", Rs_Pro!ID, True)
     LinkearTexto
     Botones pEdit3, pBorr3, pEdit3, True, True, True, True, True, True, True
  End If
  
End Sub

Private Sub Siguiente()
  On Error GoTo errHandler
  
  Rs_Pro.MoveNext
  If Rs_Pro.EOF Then
     Rs_Pro.MoveLast
     Beep
  End If
  Set Rs = cRsl.TraerRS("ProveedorTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Anterior()
  On Error GoTo errHandler
  Rs_Pro.MovePrevious
  If Rs_Pro.BOF Then
     Rs_Pro.MoveFirst
     Beep
  End If
  Set Rs = cRsl.TraerRS("ProveedorTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Primero()
  On Error GoTo errHandler
  Rs_Pro.MoveFirst
  
  Set Rs = cRsl.TraerRS("ProveedorTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub ultimo()
  On Error GoTo errHandler
  Rs_Pro.MoveLast
  
  Set Rs = cRsl.TraerRS("ProveedorTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub TxtCodigo_Change()
  TodoBien
End Sub

Private Sub TxtCodigo_LostFocus()
  Dim cRsl As ClsLectura, Rs1 As ADODB.Recordset
  Set cRsl = New ClsLectura
  Set Rs1 = New ADODB.Recordset

  If CmdBotones(2).Caption <> "Grabar" Then
     If TxtCodigo.Text <> "" Then
        Set Rs1 = cRsl.TraerRS("ProveedorTraerUnoS", TxtCodigo.Text, True)
        If Rs1.EOF = False Or Rs1.BOF = False Then
           MsgBox "Codigo de Proveedor Repetido", vbCritical, "Atención"
           TxtCodigo.SetFocus
           TxtCodigo.Text = ""
        End If
     End If
  End If
  
End Sub

Private Sub TxtPorc_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtPorc) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPorc_GotFocus()
  TxtPorc.SelStart = 0
  TxtPorc.SelLength = Len(TxtPorc.Text)
End Sub

Private Sub TxtMaxima_LostFocus()
  If TxtPorc.Text = "" Then TxtPorc.Text = 0
  TxtPorc.Text = Format(TxtPorc.Text, "0.00")
End Sub

Private Sub TxtPostal_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtRazon_Change()
  TodoBien
End Sub

Private Sub Buscar()
  nDat = 0
  nBuscar = 2
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs_Pro = cRsl.TraerTodos("Proveedores", "Id")
     Rs_Pro.MoveFirst
     Rs_Pro.Find "Id=" & nDat
     Set Rs = cRsl.TraerRS("ProveedorTraerUno", nDat, True)
     LinkearTexto
  End If
  
End Sub

Private Sub TodoBien()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtCodigo.Text) <> 0 And Len(TxtRazon.Text) <> 0 And Len(CmbIva.Text) <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtCodigo.Text) <> 0 And Len(TxtRazon.Text) <> 0 And Len(CmbIva.Text) <> 0 Then
           CmdBotones(2).Enabled = True
        End If
     End If
  End If
End Sub

Private Sub TxtRetencion_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtRetencion) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtRetencion_GotFocus()
  TxtRetencion.SelStart = 0
  TxtRetencion.SelLength = Len(TxtRetencion.Text)
End Sub

Private Sub TxtRetencion_LostFocus()
  If TxtRetencion.Text = "" Then TxtRetencion.Text = 0
  TxtRetencion.Text = Format(TxtRetencion.Text, "0.00")
End Sub

