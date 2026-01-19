VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MtrClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes"
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
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   780
      Index           =   9
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4050
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
      TabIndex        =   23
      Top             =   2970
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
      TabIndex        =   22
      Top             =   2115
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
      TabIndex        =   21
      Top             =   1275
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
      TabIndex        =   20
      Top             =   420
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   780
      Index           =   4
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4050
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Buscar"
      Height          =   780
      Index           =   3
      Left            =   3111
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   780
      Index           =   2
      Left            =   2115
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   780
      Index           =   1
      Left            =   1147
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   780
      Index           =   0
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4050
      Width           =   915
   End
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
      TabIndex        =   10
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
         TabIndex        =   27
         Top             =   -15
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
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   5895
            Begin VB.Frame FraDescuento 
               Caption         =   "Otros"
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
               Left            =   15
               TabIndex        =   55
               Top             =   15
               Width           =   5895
               Begin VB.CommandButton CmdPercepcion 
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
                  Height          =   270
                  Left            =   5070
                  TabIndex        =   73
                  Top             =   540
                  Width           =   405
               End
               Begin VB.TextBox TxtId 
                  Height          =   315
                  Left            =   1485
                  TabIndex        =   64
                  Text            =   "Text1"
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   1110
               End
               Begin VB.ComboBox CmbVendedor 
                  Height          =   315
                  Left            =   210
                  Style           =   2  'Dropdown List
                  TabIndex        =   63
                  Top             =   1875
                  Width           =   2580
               End
               Begin VB.ComboBox CmbListaPrecio 
                  Height          =   315
                  Left            =   195
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   525
                  Width           =   2580
               End
               Begin VB.ComboBox CmbZona 
                  Height          =   315
                  Left            =   210
                  Style           =   2  'Dropdown List
                  TabIndex        =   61
                  Top             =   1170
                  Width           =   2580
               End
               Begin VB.ComboBox CmbDescuento 
                  Height          =   315
                  Left            =   210
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   2520
                  Width           =   2580
               End
               Begin VB.TextBox TxtLimite 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4485
                  TabIndex        =   59
                  Text            =   "Text1"
                  Top             =   2535
                  Width           =   990
               End
               Begin VB.ComboBox CmbRecargo 
                  Height          =   315
                  Left            =   2985
                  Style           =   2  'Dropdown List
                  TabIndex        =   58
                  Top             =   1875
                  Width           =   2580
               End
               Begin VB.TextBox TxtPercepcion 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3990
                  TabIndex        =   57
                  Text            =   "Text1"
                  Top             =   525
                  Width           =   990
               End
               Begin VB.TextBox TxtRetencion 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4035
                  TabIndex        =   56
                  Text            =   "Text1"
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   990
               End
               Begin VB.Label Label19 
                  Caption         =   "Vendedor :"
                  Height          =   240
                  Left            =   210
                  TabIndex        =   72
                  Top             =   1620
                  Width           =   975
               End
               Begin VB.Label Label23 
                  Caption         =   "Descuentos"
                  Height          =   255
                  Left            =   210
                  TabIndex        =   71
                  Top             =   2295
                  Width           =   1185
               End
               Begin VB.Label Label21 
                  Caption         =   "Lista de Precio :"
                  Height          =   270
                  Left            =   210
                  TabIndex        =   70
                  Top             =   285
                  Width           =   1245
               End
               Begin VB.Label Label20 
                  Caption         =   "Zona :"
                  Height          =   285
                  Left            =   210
                  TabIndex        =   69
                  Top             =   930
                  Width           =   630
               End
               Begin VB.Label Label16 
                  Caption         =   "Limite de Compra :"
                  Height          =   210
                  Left            =   3045
                  TabIndex        =   68
                  Top             =   2535
                  Width           =   1380
               End
               Begin VB.Label Label22 
                  Caption         =   "Recargo :"
                  Height          =   255
                  Left            =   2985
                  TabIndex        =   67
                  Top             =   1620
                  Width           =   1185
               End
               Begin VB.Label Label25 
                  Caption         =   "Percepción :"
                  Height          =   255
                  Left            =   3015
                  TabIndex        =   66
                  Top             =   525
                  Width           =   945
               End
               Begin VB.Label Label26 
                  Caption         =   "Retención :"
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   65
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   945
               End
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
               TabIndex        =   54
               ToolTipText     =   "Imprimir Contancia de Cuit"
               Top             =   570
               Width           =   375
            End
            Begin VB.CommandButton CmdCuit 
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
               TabIndex        =   53
               ToolTipText     =   "Traer Datos Afip"
               Top             =   1080
               Width           =   375
            End
            Begin VB.ComboBox CmbIva 
               Height          =   315
               Left            =   210
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   570
               Width           =   2460
            End
            Begin VB.TextBox TxtCuit 
               Height          =   285
               Left            =   1470
               MaxLength       =   11
               TabIndex        =   49
               Text            =   "TxtCuit"
               Top             =   1050
               Width           =   1200
            End
            Begin VB.TextBox TxtBrutos 
               Height          =   285
               Left            =   1020
               MaxLength       =   20
               TabIndex        =   51
               Text            =   "TxtBrutos"
               Top             =   1560
               Width           =   1650
            End
            Begin VB.Label Label24 
               Caption         =   "C.de Iva :"
               Height          =   270
               Left            =   210
               TabIndex        =   52
               Top             =   315
               Width           =   825
            End
            Begin VB.Label Label17 
               Caption         =   "C.U.I.T :"
               Height          =   195
               Left            =   180
               TabIndex        =   50
               Top             =   1050
               Width           =   735
            End
            Begin VB.Label Label18 
               Caption         =   "Ing.Brutos :"
               Height          =   285
               Left            =   165
               TabIndex        =   48
               Top             =   1530
               Width           =   900
            End
         End
         Begin VB.TextBox CmbLocalidad 
            Height          =   300
            Left            =   2895
            MaxLength       =   30
            TabIndex        =   31
            Text            =   "CmbLocalidad"
            Top             =   780
            Width           =   2670
         End
         Begin VB.ComboBox CmbTransporte 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   2430
            Width           =   4230
         End
         Begin VB.TextBox TxtDomicilio 
            Height          =   300
            Left            =   1335
            MaxLength       =   45
            TabIndex        =   28
            Text            =   "TxtDomicilio"
            Top             =   390
            Width           =   4215
         End
         Begin VB.TextBox TxtPostal 
            Height          =   300
            Left            =   1305
            MaxLength       =   4
            TabIndex        =   30
            Text            =   "TxtPostal"
            Top             =   795
            Width           =   570
         End
         Begin VB.ComboBox CmbProvincia 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1215
            Width           =   1815
         End
         Begin VB.ComboBox CmbPais 
            Height          =   315
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1215
            Width           =   1800
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   300
            Left            =   1305
            MaxLength       =   17
            TabIndex        =   37
            Text            =   "TxtTelefono"
            Top             =   1635
            Width           =   1635
         End
         Begin VB.TextBox TxtFax 
            Height          =   300
            Left            =   3945
            MaxLength       =   17
            TabIndex        =   39
            Text            =   "TxtFax"
            Top             =   1635
            Width           =   1635
         End
         Begin VB.TextBox TxtEmail 
            Height          =   300
            Left            =   1305
            MaxLength       =   200
            TabIndex        =   41
            Text            =   "TxtEmail"
            Top             =   2055
            Width           =   4230
         End
         Begin VB.Label Label7 
            Caption         =   "Domicilio :"
            Height          =   300
            Index           =   1
            Left            =   210
            TabIndex        =   45
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Localidad :"
            Height          =   240
            Left            =   2025
            TabIndex        =   44
            Top             =   795
            Width           =   900
         End
         Begin VB.Label Label9 
            Caption         =   "C.Postal :"
            Height          =   225
            Left            =   180
            TabIndex        =   43
            Top             =   795
            Width           =   750
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia :"
            Height          =   240
            Left            =   180
            TabIndex        =   40
            Top             =   1215
            Width           =   930
         End
         Begin VB.Label Label11 
            Caption         =   "Pais :"
            Height          =   240
            Left            =   3240
            TabIndex        =   38
            Top             =   1290
            Width           =   525
         End
         Begin VB.Label Label12 
            Caption         =   "Teléfono :"
            Height          =   195
            Left            =   180
            TabIndex        =   36
            Top             =   1635
            Width           =   915
         End
         Begin VB.Label Label13 
            Caption         =   "Fax :"
            Height          =   300
            Left            =   3075
            TabIndex        =   34
            Top             =   1635
            Width           =   510
         End
         Begin VB.Label Label14 
            Caption         =   "E-Mail :"
            Height          =   210
            Left            =   180
            TabIndex        =   32
            Top             =   2070
            Width           =   720
         End
         Begin VB.Label Label15 
            Caption         =   "Transporte :"
            Height          =   225
            Left            =   180
            TabIndex        =   29
            Top             =   2460
            Width           =   1005
         End
      End
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   855
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker DTPBaja 
         Height          =   330
         Left            =   4110
         TabIndex        =   8
         Top             =   2745
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   152698881
         CurrentDate     =   36768
      End
      Begin MSComCtl2.DTPicker DTPAlta 
         Height          =   330
         Left            =   1650
         TabIndex        =   7
         Top             =   2745
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   152698881
         CurrentDate     =   36768
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   540
         Left            =   1665
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "MtrClientes.frx":0000
         Top             =   2070
         Width           =   3720
      End
      Begin VB.TextBox TxtContacto 
         Height          =   285
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   5
         Text            =   "TxtContacto"
         Top             =   1680
         Width           =   3720
      End
      Begin VB.TextBox TxtRazon 
         Height          =   285
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   4
         Text            =   "TxtRazon"
         Top             =   1305
         Width           =   3720
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   285
         Left            =   1695
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "TxtCodigo"
         Top             =   465
         Width           =   1500
      End
      Begin VB.CheckBox ChkActivo 
         Caption         =   "Activo"
         Height          =   270
         Left            =   4635
         TabIndex        =   1
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo :"
         Height          =   225
         Index           =   0
         Left            =   405
         TabIndex        =   26
         Top             =   870
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Contacto :"
         Height          =   330
         Left            =   375
         TabIndex        =   25
         Top             =   1620
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Baja :"
         Height          =   240
         Left            =   3075
         TabIndex        =   15
         Top             =   2745
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Alta :"
         Height          =   240
         Left            =   375
         TabIndex        =   14
         Top             =   2745
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones :"
         Height          =   285
         Left            =   375
         TabIndex        =   13
         Top             =   2025
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "Razón Social :"
         Height          =   255
         Left            =   375
         TabIndex        =   12
         Top             =   1275
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   405
         TabIndex        =   11
         Top             =   450
         Width           =   885
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3765
      Left            =   165
      TabIndex        =   9
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6641
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificación"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dirección"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Impuestos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros"
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
Attribute VB_Name = "MtrClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCod As Integer, sw As Boolean, nId As Integer
Dim sCod1 As String
Dim rs As ADODB.Recordset, Rs_Cli As ADODB.Recordset
Dim cRsl As ClsLectura
Dim cRsC As ClsClienteL
Dim cRse As ClsClienteE
Dim p As ClsPrograma

Dim bMinuscula As Boolean

Private Sub CmbIva_Click()
  TodoBien
End Sub
Private Sub CmbListaPrecio_Click()
  TodoBien
End Sub

Private Sub CmbRecargo_Click()
  TodoBien
End Sub

Private Sub CmbVendedor_Click()
  TodoBien
End Sub

Private Sub CmbZona_Click()
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

Private Sub CmdContancia_Click()
'  On Error GoTo errHandler
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

Private Sub CmdCuit_Click()
  On Error GoTo errHandler
  Dim Padron As Object, ok As Variant, Impuesto As Variant, Actividad As Variant, cRsl As ClsLectura
     
  CmdCuit.Enabled = False
  
  If TxtCuit.text <> "" Then

     Dim obj As DsPadronA4
         
     Set obj = New DsPadronA4
     Dim Resultado As Boolean
          
     Resultado = obj.SolicitarCuit(TxtCuit.text)
     
     If Resultado = True Then
      '  If CmdBotones(2).Caption = "Grabar" Then
          TxtRazon.text = obj.Nombre
          TxtDomicilio.text = obj.Direccion
          TxtCuit.text = obj.Cuit
          TxtPostal.text = obj.CodPostal
          CmbLocalidad.text = obj.Localidad
          CmbProvincia.text = obj.Provincia
          CmbIva.text = obj.CategoriaIva
'       Else
'         MsgBox "El Cuit Ingresado ya existe en la Base de Datos", vbCritical, "Atención"
'         TxtCuit.text = ""
'         CmdCuit.Enabled = True
'         Exit Sub
      ' End If
     Else
        MsgBox "El Cuit solicitado No Corresponde a una Persona Física" & Chr(10) & Chr(13), vbInformation, "Atención"
        CmdCuit.Enabled = True
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
   CmdCuit.Enabled = True
Exit Sub
    
errHandler:
   ManejaErrores
End Sub

Private Sub CmdPercepcion_Click()
   Dim obj As DsPadronA4
   
   Set obj = New DsPadronA4
   
    If TxtCuit.text <> "" Then
       Dim Resultado As Boolean
          
       CmdPercepcion.Enabled = False
       Resultado = obj.SolicitarIB(TxtCuit.text)
     
       If Resultado = True Then
           If obj.Resultado = "Ok" Then
              TxtPercepcion.text = Replace(obj.Percepcion, ",", ".")
           Else
              MsgBox obj.Mensaje, vbCritical, "Atención"
           End If
       End If
       CmdPercepcion.Enabled = True
    Else
       MsgBox "El Campo CUIT esta vacio...", vbCritical, "Atención"
    End If
   
End Sub

Private Sub Form_Load()
  On Error GoTo errHandler

  Set rs = New ADODB.Recordset
  Set Rs_Cli = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set p = New ClsPrograma

  VerFrame True, False, False, False
  sw = True
  Limpiar
  Habilitar False
  CargarCombos
  Botones pEdit1, pBorr1, pEdit1, False, False, False, False, False, False, True
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
  CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
  CmdBotones(3).Picture = LoadResPicture("Buscar", 0)
  CmdBotones(4).Picture = LoadResPicture("Imprimir", 0)
  CmdBotones(9).Picture = LoadResPicture("Salir", 0)
  
  ArmaRsParaMovimiento
  
 ' Set cRsl = Nothing
  Me.Top = 0
  Me.Left = 0
  
Exit Sub

errHandler:
   ManejaErrores
 
End Sub


Private Sub LblCuenta_Change()
  TodoBien
End Sub

Private Sub LstImpuestos_Click()
  TodoBien
End Sub

Private Sub TabStrip1_Click()
 Select Case TabStrip1.SelectedItem.Index
        Case 1
             VerFrame True, False, False, False
        Case 2
             VerFrame True, True, False, False
             If TxtDomicilio.Enabled = True Then
                TxtDomicilio.SetFocus
             End If
        Case 3
             VerFrame True, True, True, False
             If CmbIva.Enabled = True Then
                CmbIva.SetFocus
             End If
        Case 4
             VerFrame True, True, True, True
             If CmbZona.Enabled = True Then
                CmbZona.SetFocus
             End If
 End Select
End Sub

Private Sub VerFrame(bIden As Boolean, bDomi As Boolean, bImp As Boolean, bDes As Boolean)
  FraCliente.Visible = bIden
  FraDireccion.Visible = bDomi
  FraImpuesto.Visible = bImp
  FraDescuento.Visible = bDes
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
      KeyAscii = 0
   End If
   If bMinuscula = False Then
      KeyAscii = Upper(KeyAscii)
   End If
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  For Each ctl In Controls
      If TypeOf ctl Is ComboBox Then
         ctl.ListIndex = -1
      End If
  Next ctl
  TxtRetencion.text = "0.00"
  TxtPercepcion.text = "0.00"
  TxtLimite.text = "0.00"
  DTPAlta.Value = Date
  DTPBaja.Value = Date
  bMinuscula = False
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
  CmdCuit.Enabled = bEstado
  CmdContancia.Enabled = bEstado
  CmdPercepcion.Enabled = bEstado
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
  cRsl.CargaCombo CmbZona, "Zonas", "Zona", "Descripcion", ""
  cRsl.CargaCombo CmbIva, "TipoIva", "TipoIva", "Descripcion", ""
  cRsl.CargaCombo CmbListaPrecio, "ListaDePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmBVendedor, "Vendedores", "Vendedor", "Descripcion", ""
  cRsl.CargaCombo CmbTipo, "TipoRazon", "Id", "Descripcion", ""
  cRsl.CargaCombo CmbTransporte, "Transporte", "Id", "Descripcion", ""
  cRsl.CargaCombo CmbDescuento, "Descuentos", "Descuento", "Descripcion", ""
  cRsl.CargaCombo CmbRecargo, "Recargos", "Recargo", "Descripcion", ""
  
  p.BuildComboIndex "Tipo", CmbTipo
  p.BuildComboIndex "Pais", CmbPais
  p.BuildComboIndex "Provincia", CmbProvincia
  p.BuildComboIndex "Zona", CmbZona
  p.BuildComboIndex "Iva", CmbIva
  p.BuildComboIndex "ListaPrecio", CmbListaPrecio
  p.BuildComboIndex "Vendedor", CmBVendedor
  p.BuildComboIndex "Descuento", CmbDescuento
  p.BuildComboIndex "Recargo", CmbRecargo
  p.BuildComboIndex "Transporte", CmbTransporte
 
End Sub
Private Sub LinkearTexto()
  Dim Rsi As ADODB.Recordset
     
  Limpiar
  If rs.RecordCount <> 0 Then
    ChkActivo.Value = 1
    If Not IsNull(rs!FechaBaja) Then
       ChkActivo.Value = 0
       DTPBaja.Value = rs!FechaBaja
     End If
     TxtId = rs!id
     'p.SetComboByItemData CmbTipo, rs!tipo ' cLoc.TraerValorDeUnCampo("TipoRazon", "Descripcion", "Id=" & rs!tipo)
     p.SetComboByIdFast "Tipo", CmbTipo, rs!tipo
     TxtCodigo.text = rs!Cliente
     TxtRazon.text = "" & rs!RazonSocial
     TxtContacto.text = "" & rs!Contacto
     TxtObservaciones.text = "" & rs!Observaciones
     TxtDomicilio.text = "" & rs!Domicilio
     TxtPostal.text = "" & rs!CodigoPostal
     TxtTelefono.text = "" & rs!Telefono
     TxtFax.text = "" & rs!Fax
     TxtEmail.text = "" & rs!email
     If Not IsNull(rs!Transporte) Then
        ' p.SetComboByItemData CmbTransporte, CLng(rs!Transporte) 'cVen.DatoCombo("Transporte", "Id", "Descripcion", rs!Transporte)
         p.SetComboByIdFast "Transporte", CmbTransporte, rs!Transporte
     End If
     TxtCuit.text = "" & rs!Cuit
     TxtBrutos.text = "" & rs!IngBruto
     DTPAlta.Value = rs!FechaAlta
     TxtLimite.text = Format(rs!Limite, nCantDecimales)
     TxtRetencion.text = Format(rs!Retencion, nCantDecimales)
     TxtPercepcion.text = Format(rs!Percepcion, nCantDecimales)
  
     CmbLocalidad.text = "" & rs!Ciudad
     
     If Not IsNull(rs!Pais) Then
       ' p.SetComboByItemData CmbPais, CLng(rs!Pais) ' CmbPais.text = cPai.DatoCombo("Paises", "Pais", "Descripcion", rs!Pais)
       p.SetComboByIdFast "Pais", CmbPais, rs!Pais
     End If
     If Not IsNull(rs!Provincia) Then
        p.SetComboByIdFast "Provincia", CmbProvincia, rs!Provincia ' cPro.DatoCombo("Provincias", "Provincia", "Descripcion", rs!Provincia)
     End If
     If Not IsNull(rs!TipoIva) Then
        'p.SetComboByItemData CmbIva, CLng(rs!TipoIva) ' CmbIva.text = cTii.DatoCombo("TipoIva", "TipoIva", "Descripcion", rs!TipoIva)
        p.SetComboByIdFast "Iva", CmbIva, rs!TipoIva
     End If
     If Not IsNull(rs!Zona) Then
       ' p.SetComboByItemData CmbZona, CLng(rs!Zona) 'CmbZona.text = cZon.DatoCombo("Zonas", "Zona", "Descripcion", rs!Zona)
       p.SetComboByIdFast "Zona", CmbZona, rs!Zona
     End If
     If Not IsNull(rs!precio) Then
        'p.SetComboByItemData CmbListaPrecio, CLng(rs!precio) ' CmbListaPrecio.text = cLip.DatoCombo("ListaDePrecio", "ListaPrecio", "Descripcion", rs!precio)
        p.SetComboByIdFast "ListaPrecio", CmbListaPrecio, rs!precio
     End If
     If Not IsNull(rs!Vendedor) Then
        'p.SetComboByItemData CmbVendedor, CLng(rs!Vendedor) 'CmBVendedor.text = cVen.DatoCombo("Vendedores", "Vendedor", "Descripcion", rs!Vendedor)
        p.SetComboByIdFast "Vendedor", CmBVendedor, rs!Vendedor
     End If
     If Not IsNull(rs!Descuento) Then
        'p.SetComboByItemData CmbDescuento, CLng(rs!Descuento) ' CmbDescuento.text = cVen.DatoCombo("Descuentos", "Descuento", "Descripcion", rs!Descuento)
        p.SetComboByIdFast "Descuento", CmbDescuento, rs!Descuento
     End If
     If Not IsNull(rs!recargo) Then
       ' p.SetComboByItemData CmbRecargo, CLng(rs!recargo) 'CmbRecargo.text = cVen.DatoCombo("Recargos", "Recargo", "Descripcion", rs!Recargo)
       p.SetComboByIdFast "Recargo", CmbRecargo, rs!recargo
     End If
     
  End If
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
  Dim RsReg As Recordset
    
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
        TxtCodigo.text = cRsl.ActualizarCampoTraer("Cliente")
        TxtRazon.SetFocus
     Else
        TxtCodigo.SetFocus
     End If
     sw = False
     ChkActivo.Value = 1
     TxtLimite.text = "0.00"
     
     Set RsReg = cRsl.TraerRsCondi("Registros", "Usuario", "Usuario=" & nUsuario & " And Suc = " & nSucursal)
     
     CmbTipo.text = "Cliente"
     p.SetComboByItemData CmbProvincia, RsReg!Provincia ' CmbProvincia.text = cRsl.DatoCombo("Provincias", "Provincia", "Descripcion", cRsl.RegPorDefecto("Provincia"))
     p.SetComboByItemData CmbPais, 1 ' CmbPais.text = cRsl.DatoCombo("Paises", "Pais", "Descripcion", 1)
     p.SetComboByItemData CmBVendedor, RsReg!Vendedor 'CmbVendedor.text = cRsl.DatoCombo("Vendedores", "Vendedor", "Descripcion", cRsl.RegPorDefecto("Vendedor"))
     p.SetComboByItemData CmbListaPrecio, RsReg!ListaPrecio 'CmbListaPrecio.text = cRsl.DatoCombo("ListaDePrecio", "ListaPrecio", "Descripcion", cRsl.RegPorDefecto("ListaPrecio"))
     p.SetComboByItemData CmbZona, RsReg!Zona 'CmbZona.text = cRsl.DatoCombo("Zonas", "Zona", "Descripcion", cRsl.RegPorDefecto("Zona"))
     p.SetComboByItemData CmbIva, RsReg!Impuesto 'CmbIva.text = cRsl.DatoCombo("TipoIva", "TipoIva", "Descripcion", cRsl.RegPorDefecto("Impuesto"))
     p.SetComboByItemData CmbDescuento, RsReg!Descuento 'CmbDescuento.text = cRsl.DatoCombo("Descuentos", "Descuento", "Descripcion", cRsl.RegPorDefecto("Descuento"))
     p.SetComboByItemData CmbTransporte, RsReg!Transporte 'CmbTransporte.text = cRsl.DatoCombo("Transporte", "Id", "Descripcion", cRsl.RegPorDefecto("Transporte"))
     CmbRecargo.text = "Sin Recargo"
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
      If Rs_Cli.BOF = False And Rs_Cli.EOF = False Then
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
     sCod1 = rs!Cliente
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
  
  Set cRse = New ClsClienteE
  
  
  If cRsl.TraerCantidad("CabComprobantes", "Cliente", "Cliente='" & TxtCodigo.text & "'") = 0 And cRsl.TraerCantidad("Registros", "ClienteConsFinal", "ClienteConsFinal=" & Val(TxtId.text)) = 0 Then
    If MsgBox("Borra el Cliente Seleccionado ?", 20, "Atención") = 6 Then
       cRse.Borrar "ClienteBorrar", , TxtId.text
       Limpiar
       Set cRse = Nothing
       Rs_Cli.Delete
       If Rs_Cli.RecordCount <> 0 Then
          Anterior
       Else
          Botones True, False, False, False, False, False, False, False, False, True
       End If
    End If
  Else
     MsgBox "Cliente Asociado a una Factura de Venta o Cliente Cons.Final del Sistema", vbCritical, "Atención"
  End If
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Grabar()
  On Error GoTo errHandler
  
  Dim Rs1 As ADODB.Recordset
  Dim i As Integer
  
  Set Rs1 = New ADODB.Recordset
    
  Set cRse = New ClsClienteE
    
  Set Rs1 = cRsl.RsVacio("Clientes", "Cliente", "S")
  
  Rs1.AddNew
 ' Rs1!Id = 0
  Rs1!tipo = CmbTipo.ItemData(CmbTipo.ListIndex)
  Rs1!Cliente = TxtCodigo.text
  Rs1!RazonSocial = Replace(TxtRazon.text, "'", "")
  Rs1!Contacto = Replace(TxtContacto.text, "'", "")
  Rs1!Observaciones = Replace(TxtObservaciones.text, "'", "")
  Rs1!FechaAlta = Date
  Rs1!FechaBaja = IIf(ChkActivo.Value = 0, DTPBaja.Value, Null)
  Rs1!Domicilio = Replace(TxtDomicilio.text, "'", "")
  Rs1!CodigoPostal = IIf(TxtPostal.text = "", 0, TxtPostal.text)
  Rs1!Localidad = 1
  Rs1!Ciudad = Replace(CmbLocalidad.text, "'", "")
  
  If CmbProvincia.ListIndex <> -1 Then
     Rs1!Provincia = CmbProvincia.ItemData(CmbProvincia.ListIndex)
  End If
  
  If CmbPais.ListIndex <> -1 Then
     Rs1!Pais = CmbPais.ItemData(CmbPais.ListIndex)
  End If

  Rs1!Telefono = TxtTelefono.text
  Rs1!Fax = TxtFax.text
  If CmbTransporte.ListIndex <> -1 Then
     Rs1!Transporte = CmbTransporte.ItemData(CmbTransporte.ListIndex)
  End If
  
  If CmbIva.ListIndex <> -1 Then
     Rs1!TipoIva = CmbIva.ItemData(CmbIva.ListIndex)
  End If
  
  Rs1!Cuit = TxtCuit.text
  Rs1!IngBruto = TxtBrutos.text
  
  If CmbZona.ListIndex <> -1 Then
     Rs1!Zona = CmbZona.ItemData(CmbZona.ListIndex)
  End If
  
  If CmbListaPrecio.ListIndex <> -1 Then
     Rs1!precio = CmbListaPrecio.ItemData(CmbListaPrecio.ListIndex)
  End If
  
  If CmBVendedor.ListIndex <> -1 Then
     Rs1!Vendedor = CmBVendedor.ItemData(CmBVendedor.ListIndex)
  End If
  
  Rs1!email = TxtEmail.text
  
  If CmbDescuento.ListIndex <> -1 Then
     Rs1!Descuento = CmbDescuento.ItemData(CmbDescuento.ListIndex)
  End If
  
  If CmbRecargo.ListIndex <> -1 Then
     Rs1!recargo = CmbRecargo.ItemData(CmbRecargo.ListIndex)
  End If

  If pClave = False Then
     Rs1!Limite = TxtLimite.text
  Else
     If CmdBotones(2).Caption <> "Grabar" Then
        Rs1!Limite = 0
     Else
        FrmClave.Show 1
        If bClave = True Then
           Rs1!Limite = TxtLimite.text
        Else
          Rs1!Limite = 0
        End If
     End If
  End If
  Rs1!Retencion = TxtRetencion.text
  Rs1!Percepcion = TxtPercepcion.text
  Rs1.Update
  ' Carga Detalle de Impuestos
      
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Modificar Rs1, "", 0
  Else
     cRse.Agregar Rs1, 0
  End If
  
  Rs1.Close
  Set cRse = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub ArmaRsParaMovimiento()
  Set Rs_Cli = cRsl.TraerTodos("Clientes", "Id")
  Botones pEdit, False, False, False, False, False, False, False, False, True
  If Rs_Cli.BOF = False And Rs_Cli.EOF = False Then
     Rs_Cli.MoveLast
     Set rs = cRsl.TraerRS("ClienteTraerUno", Rs_Cli!id, True)
     LinkearTexto
     Botones pEdit1, pBorr1, pEdit1, True, True, True, True, True, True, True
  End If
End Sub

Private Sub Siguiente()
  On Error GoTo errHandler
  Rs_Cli.MoveNext
  If Rs_Cli.EOF Then
     Rs_Cli.MoveLast
     Beep
  End If
  
  Set rs = cRsl.TraerRS("ClienteTraerUno", Rs_Cli!id, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Anterior()
  On Error GoTo errHandler
  Rs_Cli.MovePrevious
  If Rs_Cli.BOF Then
     Rs_Cli.MoveFirst
     Beep
  End If
  
  Set rs = cRsl.TraerRS("ClienteTraerUno", Rs_Cli!id, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Primero()
  On Error GoTo errHandler
  Rs_Cli.MoveFirst
 
  Set rs = cRsl.TraerRS("ClienteTraerUno", Rs_Cli!id, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub ultimo()
  On Error GoTo errHandler
  Rs_Cli.MoveLast
  
  Set rs = cRsl.TraerRS("ClienteTraerUno", Rs_Cli!id, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub TxtCodigo_Change()
  TodoBien
End Sub

Private Sub TxtCodigo_LostFocus()
  Dim Rs1 As ADODB.Recordset

  Set Rs1 = New ADODB.Recordset

  If CmdBotones(2).Caption <> "Grabar" Then
     If TxtCodigo.text <> "" Then
        Set Rs1 = cRsl.TraerRS("ClienteTraerUnoS", TxtCodigo.text, True)
        If Rs1.EOF = False Or Rs1.BOF = False Then
           MsgBox "Codigo de Cliente Repetido", vbCritical, "Atención"
           TxtCodigo.SetFocus
           TxtCodigo.text = ""
        End If
     End If
  End If
 
End Sub

Private Sub TxtCuit_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCuit_LostFocus()
  If TxtCuit.text <> "" Then
    If CmbIva.ItemData(CmbIva.ListIndex) <> 1 Then
       If ValidaCuit(TxtCuit) = False Then
          TxtCuit.text = ""
          TxtCuit.SetFocus
        Else
           If cRsl.TraerCantidad("Clientes", "Cuit", "Cuit='" & TxtCuit.text & "'") = 1 Then
              Dim sCli As String
              sCli = cRsl.TraerValorDeUnCampo("Clientes", "Cliente", "Cuit='" & TxtCuit.text & "'")
              MsgBox "Cuit de Cliente Repetido... " & sCli, vbCritical, "Atención"
              TxtCuit.text = ""
              TxtCuit.SetFocus
          End If
       End If
    End If
  End If
End Sub

Private Sub TxtEmail_GotFocus()
   bMinuscula = True
End Sub

Private Sub TxtEmail_LostFocus()
   bMinuscula = False
End Sub

Private Sub TxtLimite_Change()
  TodoBien
End Sub

Private Sub TxtPercepcion_Change()
  TodoBien
End Sub

Private Sub TxtRazon_Change()
  TodoBien
End Sub
Private Sub Buscar()
 
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs_Cli = cRsl.TraerTodos("Clientes", "Id")
     Rs_Cli.MoveFirst
     Rs_Cli.Find "Id=" & nDat
     Set rs = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     LinkearTexto
  End If
 
End Sub

Private Sub TodoBien()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtCodigo.text) <> 0 And Len(TxtRazon.text) <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtCodigo.text) <> 0 And Len(TxtRazon.text) <> 0 Then
           CmdBotones(2).Enabled = True
        End If
     End If
  End If
End Sub

Private Sub TxtLimite_GotFocus()
  TxtLimite.SelStart = 0
  TxtLimite.SelLength = Len(TxtLimite.text)
End Sub
Private Sub TxtLimite_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtLimite) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtLimite_LostFocus()
  TxtLimite.text = Format(TxtLimite.text, nCantDecimales)
End Sub

Private Sub TxtRetencion_Change()
  TodoBien
End Sub

Private Sub TxtRetencion_GotFocus()
  TxtRetencion.SelStart = 0
  TxtRetencion.SelLength = Len(TxtRetencion.text)
End Sub
Private Sub TxtRetencion_KeyPress(KeyAscii As Integer)
   If SoloNumeroDecimalFinal(KeyAscii, TxtRetencion) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtRetencion_LostFocus()
  TxtRetencion.text = Format(TxtRetencion.text, nCantDecimales)
End Sub

Private Sub TxtPercepcion_GotFocus()
  TxtPercepcion.SelStart = 0
  TxtPercepcion.SelLength = Len(TxtPercepcion.text)
End Sub
Private Sub TxtPercepcion_KeyPress(KeyAscii As Integer)
   If SoloNumeroDecimalFinal(KeyAscii, TxtPercepcion) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtPercepcion_LostFocus()
  TxtPercepcion.text = Format(TxtPercepcion.text, nCantDecimales)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cRsl = Nothing  ' esto dispara Class_Terminate y cierra conexión
End Sub
