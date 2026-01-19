VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form HerraPermisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos de Usuarios"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10590
   Begin VB.Frame Frame11 
      Caption         =   "Asignar Impresoras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   5100
      TabIndex        =   23
      Top             =   1020
      Width           =   5055
      Begin VB.ListBox LstCompro 
         Height          =   2985
         Left            =   150
         TabIndex        =   25
         Top             =   255
         Width           =   4785
      End
      Begin VB.ComboBox CmbImpre 
         Height          =   315
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3525
         Width           =   4785
      End
      Begin VB.Label Label2 
         Caption         =   "Impresora Local :"
         Height          =   210
         Left            =   180
         TabIndex        =   26
         Top             =   3255
         Width           =   1260
      End
   End
   Begin VB.ComboBox CmbUsuarios 
      Height          =   315
      Left            =   945
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   90
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   150
      TabIndex        =   3
      Top             =   1035
      Width           =   4800
      Begin VB.Frame Frame3 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3945
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   4800
         Begin VB.Frame Frame4 
            Caption         =   "Proveedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3945
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   4800
            Begin VB.Frame Frame5 
               Caption         =   "Stock"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3945
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   4800
               Begin VB.Frame Frame6 
                  Caption         =   "Compras"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3945
                  Left            =   0
                  TabIndex        =   13
                  Top             =   0
                  Width           =   4800
                  Begin VB.Frame Frame7 
                     Caption         =   "Ventas"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   3945
                     Left            =   0
                     TabIndex        =   15
                     Top             =   0
                     Width           =   4800
                     Begin VB.Frame Frame8 
                        Caption         =   "Caja"
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   3945
                        Left            =   0
                        TabIndex        =   17
                        Top             =   0
                        Width           =   4800
                        Begin VB.Frame Frame9 
                           Caption         =   "Herramientas"
                           BeginProperty Font 
                              Name            =   "MS Sans Serif"
                              Size            =   8.25
                              Charset         =   0
                              Weight          =   700
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   3945
                           Left            =   0
                           TabIndex        =   19
                           Top             =   0
                           Width           =   4800
                           Begin VB.Frame Frame10 
                              Caption         =   "Otros Permisos"
                              BeginProperty Font 
                                 Name            =   "MS Sans Serif"
                                 Size            =   8.25
                                 Charset         =   0
                                 Weight          =   700
                                 Underline       =   0   'False
                                 Italic          =   0   'False
                                 Strikethrough   =   0   'False
                              EndProperty
                              Height          =   3945
                              Left            =   0
                              TabIndex        =   21
                              Top             =   15
                              Width           =   4800
                              Begin VB.ListBox List1 
                                 Height          =   3435
                                 Index           =   9
                                 Left            =   120
                                 Style           =   1  'Checkbox
                                 TabIndex        =   22
                                 Top             =   255
                                 Width           =   4530
                              End
                           End
                           Begin VB.ListBox List1 
                              Height          =   3435
                              Index           =   8
                              Left            =   180
                              Style           =   1  'Checkbox
                              TabIndex        =   20
                              Top             =   210
                              Width           =   4530
                           End
                        End
                        Begin VB.ListBox List1 
                           Height          =   3435
                           Index           =   7
                           Left            =   180
                           Style           =   1  'Checkbox
                           TabIndex        =   18
                           Top             =   240
                           Width           =   4530
                        End
                     End
                     Begin VB.ListBox List1 
                        Height          =   3435
                        Index           =   6
                        Left            =   120
                        Style           =   1  'Checkbox
                        TabIndex        =   16
                        Top             =   240
                        Width           =   4530
                     End
                  End
                  Begin VB.ListBox List1 
                     Height          =   3435
                     Index           =   5
                     Left            =   120
                     Style           =   1  'Checkbox
                     TabIndex        =   14
                     Top             =   240
                     Width           =   4530
                  End
               End
               Begin VB.ListBox List1 
                  Height          =   3435
                  Index           =   4
                  Left            =   120
                  Style           =   1  'Checkbox
                  TabIndex        =   12
                  Top             =   240
                  Width           =   4530
               End
            End
            Begin VB.ListBox List1 
               Height          =   3435
               Index           =   3
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   10
               Top             =   240
               Width           =   4530
            End
         End
         Begin VB.ListBox List1 
            Height          =   3435
            Index           =   2
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   240
            Width           =   4530
         End
      End
      Begin VB.ListBox List1 
         Height          =   3435
         Index           =   1
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   4530
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprobantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   135
      TabIndex        =   1
      Top             =   1050
      Width           =   4800
      Begin VB.ListBox List1 
         Height          =   3435
         Index           =   0
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   240
         Width           =   4530
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4560
      Left            =   -30
      TabIndex        =   0
      Top             =   705
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   8043
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comprobantes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Productos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Proveedores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Stock"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Compras"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ventas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Caja"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Herramientas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros Permisos"
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
   Begin VB.Label Label1 
      Caption         =   "Usuario :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   135
      Width           =   885
   End
End
Attribute VB_Name = "HerraPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean, sw1 As Boolean, nUserPermiso As Integer


Private Sub CmbImpre_Click()
  Dim cRsl As ClsLectura, nCant As Integer
  
  Set cRsl = New ClsLectura
  
  If LstCompro.ListIndex <> -1 Then
     nCant = cRsl.TraerCantidad("Impresoras", "Id", "Computer='" & sComputadora & "' AND Comprobante=" & LstCompro.ItemData(LstCompro.ListIndex) & " And Usuario = " & CmbUsuarios.ItemData(CmbUsuarios.ListIndex))
     If sw1 = True Then
        db.Open sDb
        If nCant = 0 Then
           db.Execute "INSERT INTO Impresoras (Usuario,Computer,Comprobante,Impresora) VALUES (" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex) & ",'" & sComputadora & "'," & LstCompro.ItemData(LstCompro.ListIndex) & ",'" & CmbImpre.text & "')"
        Else
          db.Execute "UPDATE Impresoras SET Impresora='" & CmbImpre.text & "' WHERE Computer='" & sComputadora & "' AND Comprobante=" & LstCompro.ItemData(LstCompro.ListIndex) & " And Usuario = " & CmbUsuarios.ItemData(CmbUsuarios.ListIndex)
        End If
        db.Close
        sw1 = False
     End If
  Else
    CmbImpre.ListIndex = -1
  End If
  Set cRsl = Nothing
End Sub

Private Sub CmbUsuarios_Click()
  If sw = True Then
     sw = False
     LinkearTextos
  End If
End Sub

Private Sub LinkearTextos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  If CmbUsuarios.ItemData(CmbUsuarios.ListIndex) <> 1 And nUsuario <> 1 Then
    cRsl.CargaComboLst List1(0), "UsuariosComprobantesVentas", "Id", "Descripcion", "Valor", "Valor=1 and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Descripcion,Comprobante"
    cRsl.CargaComboLst List1(1), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Clientes' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(2), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Productos' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(3), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Proveedor' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(4), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Stock' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(5), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Compras' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(6), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Ventas' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(7), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Caja' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(8), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Herrami' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(9), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Valor=1 and Tipo='Tablas' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
  Else
    cRsl.CargaComboLst List1(0), "UsuariosComprobantesVentas", "Id", "Descripcion", "Valor", "Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Descripcion,Comprobante"
    cRsl.CargaComboLst List1(1), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Clientes' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(2), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Productos' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(3), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Proveedor' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(4), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Stock' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(5), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Compras' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(6), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Ventas' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(7), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Caja' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(8), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Herrami' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
    cRsl.CargaComboLst List1(9), "UsuariosPermisos", "Id", "Descripcion", "Valor", "Tipo='Tablas' and Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex), "Orden"
  End If
  cRsl.CargaCombo LstCompro, "Comprobantes", "Id", "Descripcion", "Suc=" & nSucursal & " and (TipoOperacion=1 or TipoOperacion=5 or TipoOperacion=6 or TipoOperacion=2 or TipoOperacion=4)"

 ' cRsl.CargaCombo LstCompro, "Comprobantes", "Id", "Descripcion", " (TipoOperacion=1 or TipoOperacion=5 or TipoOperacion=6 or TipoOperacion=2)"

  Dim P As Printer
    
  For Each P In Printers
      CmbImpre.AddItem P.DeviceName
  Next
  
  CmbImpre.AddItem "Ninguna"
  
  sw = True
  
  Set cRsl = Nothing
End Sub


Private Sub Form_Load()
  sw = False
  sw1 = False
  VerFrame True, False, False, False, False, False, False, False, False, False
  CargarCombos
End Sub

Private Sub CargarCombos()
  Dim cRsl As New ClsLectura, sCond As String
  Set cRsl = New ClsLectura
  
  sCond = "Id_Usuario<>1"
  If nUsuario = 1 Then
     sCond = "Id_Usuario<>0"
  End If
  
  cRsl.CargaCombo CmbUsuarios, "Usuarios", "Id_Usuario", "Nombre", sCond, "Nombre"
  CmbUsuarios.ListIndex = 0
  LinkearTextos

  Set cRsl = Nothing
  
End Sub

Private Sub VerFrame(bFrame1 As Boolean, bFrame2 As Boolean, bFrame3 As Boolean, bFrame4 As Boolean, bFrame5 As Boolean, bFrame6 As Boolean, bFrame7 As Boolean, bFrame8 As Boolean, bFrame9 As Boolean, bFrame10 As Boolean)
  Frame1.Visible = bFrame1
  Frame2.Visible = bFrame2
  Frame3.Visible = bFrame3
  Frame4.Visible = bFrame4
  Frame5.Visible = bFrame5
  Frame6.Visible = bFrame6
  Frame7.Visible = bFrame7
  Frame8.Visible = bFrame8
  Frame9.Visible = bFrame9
  Frame10.Visible = bFrame10
  Frame11.Visible = bFrame10
End Sub

Private Sub List1_Click(Index As Integer)
  Dim cRsl As ClsLectura, nComp As Integer
  
  Set cRsl = New ClsLectura
  
  If sw = True Then
     If Index = 0 Then
        If CmbUsuarios.ItemData(CmbUsuarios.ListIndex) <> 1 Then
           db.Open sDb
           db.Execute "UPDATE UsuariosComprobantesVentas SET Valor=" & IIf(List1(Index).Selected(List1(Index).ListIndex), 1, 0) & " WHERE Id=" & List1(Index).ItemData(List1(Index).ListIndex) & " And Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex)
           db.Close
        Else
           db.Open sDb
           db.Execute "UPDATE UsuariosComprobantesVentas SET Ver=" & IIf(List1(Index).Selected(List1(Index).ListIndex), 1, 0) & ",Valor=" & IIf(List1(Index).Selected(List1(Index).ListIndex), 1, 0) & " WHERE Id=" & List1(Index).ItemData(List1(Index).ListIndex) & " And Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex)
           db.Close
        End If
     Else
        If CmbUsuarios.ItemData(CmbUsuarios.ListIndex) <> 1 Then
           db.Open sDb
           db.Execute "UPDATE UsuariosPermisos SET Valor=" & IIf(List1(Index).Selected(List1(Index).ListIndex), 1, 0) & " WHERE Id=" & List1(Index).ItemData(List1(Index).ListIndex) & " And Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex)
           db.Close
        Else
           db.Open sDb
           db.Execute "UPDATE UsuariosPermisos SET Ver=" & IIf(List1(Index).Selected(List1(Index).ListIndex), 1, 0) & ",Valor=" & IIf(List1(Index).Selected(List1(Index).ListIndex), 1, 0) & " WHERE Id=" & List1(Index).ItemData(List1(Index).ListIndex) & " And Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex)
           db.Close
        End If
     End If
     nComp = cRsl.TraerValorDeUnCampo("UsuariosComprobantesVentas", "Comprobante", "Id=" & List1(Index).ItemData(List1(Index).ListIndex) & " And Usuario=" & CmbUsuarios.ItemData(CmbUsuarios.ListIndex))
  End If
End Sub

Private Sub LstCompro_Click()
   Dim cRsl As ClsLectura
   
   Set cRsl = New ClsLectura
   
   On Error GoTo errHandler
   
   CmbImpre.text = cRsl.TraerValorDeUnCampo("Impresoras", " Impresora", "Computer='" & sComputadora & "' and Comprobante=" & LstCompro.ItemData(LstCompro.ListIndex))
   sw1 = True
    
errHandler:
  If err.Number = 383 Then
  CmbImpre.ListIndex = -1
    Resume Next
  End If
End Sub

Private Sub TabStrip1_Click()
 Select Case TabStrip1.SelectedItem.Index
        Case 1
             VerFrame True, False, False, False, False, False, False, False, False, False
        Case 2
             VerFrame True, True, False, False, False, False, False, False, False, False
        Case 3
             VerFrame True, True, True, False, False, False, False, False, False, False
        Case 4
             VerFrame True, True, True, True, False, False, False, False, False, False
        Case 5
             VerFrame True, True, True, True, True, False, False, False, False, False
        Case 6
             VerFrame True, True, True, True, True, True, False, False, False, False
        Case 7
             VerFrame True, True, True, True, True, True, True, False, False, False
        Case 8
             VerFrame True, True, True, True, True, True, True, True, False, False
        Case 9
             VerFrame True, True, True, True, True, True, True, True, True, False
        Case 10
             VerFrame True, True, True, True, True, True, True, True, True, True
 End Select
End Sub

