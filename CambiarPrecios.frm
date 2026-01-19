VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CambiarPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11955
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   9630
      TabIndex        =   14
      Top             =   540
      Width           =   1440
   End
   Begin VB.ComboBox CmbProveedor 
      Height          =   315
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   90
      Width           =   4200
   End
   Begin VB.ComboBox CmbLista 
      Height          =   315
      Left            =   765
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   105
      Width           =   4530
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   ">>|"
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
      Index           =   3
      Left            =   2970
      TabIndex        =   9
      Top             =   5670
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
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
      Height          =   345
      Index           =   2
      Left            =   1995
      TabIndex        =   8
      Top             =   5670
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
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
      Height          =   345
      Index           =   1
      Left            =   1065
      TabIndex        =   7
      Top             =   5670
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   "|<<"
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
      Index           =   0
      Left            =   165
      TabIndex        =   6
      Top             =   5670
      Width           =   840
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Descripción"
      Height          =   225
      Index           =   1
      Left            =   1485
      TabIndex        =   3
      Top             =   6165
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Código"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   6165
      Width           =   1335
   End
   Begin VB.TextBox TxtBuscar 
      Height          =   285
      Left            =   765
      TabIndex        =   1
      Text            =   "TxtBuscar"
      Top             =   495
      Width           =   8100
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Descripción"
      Height          =   225
      Index           =   2
      Left            =   3240
      TabIndex        =   0
      Top             =   6165
      Width           =   2370
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   4740
      Left            =   135
      TabIndex        =   4
      Top             =   855
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   8361
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   2058
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
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Proveedor :"
      Height          =   180
      Left            =   5445
      TabIndex        =   13
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "Lista :"
      Height          =   180
      Left            =   195
      TabIndex        =   10
      Top             =   135
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Criterio :"
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   540
      Width           =   825
   End
End
Attribute VB_Name = "CambiarPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cn As Connection
Private Rs2 As Recordset
Const PAGE_SIZE = 50
Dim Registro As Integer

Private WithEvents Rsd As ADODB.Recordset
Attribute Rsd.VB_VarHelpID = -1

Private Sub CmbLista_Click()
  Conectar
End Sub

Private Sub CmbProveedor_Click()
  Registro = 0
  Conectar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
  KeyAscii = Upper(KeyAscii)
End Sub

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
  Rsd.Sort = Rsd.Fields(ColIndex).Name & " Asc"
  If ColIndex = 0 Then
     Option1(0).Value = True
  End If
  If ColIndex = 1 Then
     Option1(1).Value = True
  End If
  If ColIndex = 2 Then
     Option1(2).Value = True
  End If
  CmdMover_Click 0
End Sub

Private Sub Option1_Click(Index As Integer)
  TxtBuscar.SetFocus
  TxtBuscar.text = ""
End Sub

Private Sub TxtBuscar_LostFocus()
   If TxtBuscar.text <> "" Then
      If Rsd.RecordCount <> 0 Then
          Dim Anterior As Long
        
          Anterior = Rsd.AbsolutePosition
          Dim nPagina As Long
          Rsd.MoveFirst

          If Option1(0).Value = True Then Rsd.Find "Producto='" & TxtBuscar.text & "'", , adSearchForward
          If Option1(1).Value = True Then Rsd.Find "Descripcion  LIKE '*" & TxtBuscar.text & "*'", , adSearchForward
          If Option1(2).Value = True Then Rsd.Find "Descripcion1 LIKE '*" & TxtBuscar.text & "*'", , adSearchForward
          nPagina = Rsd.AbsolutePage
          If nPagina >= 0 Then
             If Rsd.BOF Then
                Rsd.AbsolutePosition = Anterior
             Else
                Rsd.AbsolutePage = nPagina
    
            ' Asigna el recordset paginado al control Datagrid
               Set Grid1.DataSource = Paginar_Recordset(Rsd)
               CabGrid
               If Option1(0).Value = True Then Rs2.Find "Producto='" & TxtBuscar.text & "'", , adSearchForward
               If Option1(1).Value = True Then Rs2.Find "Descripcion  LIKE '*" & TxtBuscar.text & "*'", , adSearchForward
               If Option1(2).Value = True Then Rs2.Find "Descripcion1 LIKE '*" & TxtBuscar.text & "*'", , adSearchForward
    
               Anterior = Rs2.Bookmark
               Rs2.AbsolutePosition = Anterior
               Grid1.Bookmark = Rs2.AbsolutePosition
            End If
        End If
      Else
        MsgBox "Registro No encontrado", vbInformation, "Atención"
      End If
   End If
End Sub

Private Sub CmdMover_Click(Index As Integer)
   Dim Pagina_Actual As Long
       
    With Rsd
        Select Case Index
               
            ' Primera página del recordset
            Case 0
                Pagina_Actual = 1
                Registro = PAGE_SIZE
            ' Anterior página del recordset
            Case 1
                Pagina_Actual = .AbsolutePage - 1
                Registro = Registro - PAGE_SIZE
                If Registro < PAGE_SIZE Then
                    Registro = PAGE_SIZE
                End If
            ' Siguiente página
            Case 2
                Pagina_Actual = .AbsolutePage + 1
                Registro = Registro + PAGE_SIZE
                If Registro > Rsd.RecordCount Then
                    Registro = Rsd.RecordCount
                End If
            ' Ultima
            Case 3
                Pagina_Actual = .PageCount
                Registro = Rsd.RecordCount
        End Select
           
        ' Si estamos en la primera o última pagina sale
        If Pagina_Actual < 1 Or Pagina_Actual > .PageCount Then
            Exit Sub
        End If
           
        ' Establece la Página actual al rec
        .AbsolutePage = Pagina_Actual
           
        ' Asigna el recordset paginado al control Datagrid
        Set Grid1.DataSource = Paginar_Recordset(Rsd)
        CabGrid
        ' visualiza en el caption la página actual y los registros
        Me.Caption = Me.Caption & "  -  Registro actual: " & Registro & _
                    " de " & Rsd.RecordCount & " registros ."
           
    End With

End Sub

Private Sub Command1_Click()
'   Dim Anterior As Long
'
'   Dim nPagina As Long
'   Rsd.MoveFirst
'   Rsd.Find "Producto='" & TxtBuscar.text & "'", , adSearchForward
'   nPagina = Rsd.AbsolutePosition
'   If nPagina >= 0 Then
'      If Rsd.BOF Then
'         Rsd.AbsolutePosition = Anterior
'      Else
'         Rsd.AbsolutePage = nPagina
'
'        ' Asigna el recordset paginado al control Datagrid
'        Set Grid1.DataSource = Paginar_Recordset(Rsd)
'        CabGrid
'        Rs2.Find "Producto='" & TxtBuscar.text & "'", , adSearchForward
'        If Rs2.AbsolutePosition <> adPosEOF Then
'            Anterior = Rs2.Bookmark
'            Rs2.AbsolutePosition = Anterior
'            Grid1.Bookmark = Rs2.AbsolutePosition
'        End If
'      End If
'   Else
'      MsgBox "Registro No encontrado", vbInformation, "Atención"
'   End If
  TxtBuscar_LostFocus
End Sub

Private Sub Form_Activate()
  TxtBuscar.SetFocus
  Option1(0).Value = True
End Sub

Private Sub Form_Load()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  TxtBuscar.text = ""
  Option1(2).Caption = "Descripción Adicional"
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbProveedor, "Proveedores", "Id", "RazonSocial", ""
  
  CmbLista.ListIndex = 0
End Sub

Private Sub Conectar()
  Dim sSql As String
  
  Dim Cadena As String
  Set cn = New ADODB.Connection
  
  cn.ConnectionString = sDb
  cn.Open
  
  Set Rsd = New Recordset
  Set Rs2 = New Recordset
  

  sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.Descripcion1,Productos.PrecioCompra, Precios.Porcentaje, Precios.Precio, Productos.FechaCompra "
  sSql = sSql & "FROM Productos LEFT JOIN Precios ON Productos.Producto = Precios.Producto "
  sSql = sSql & "WHERE Precios.Lista=" & CmbLista.ItemData(CmbLista.ListIndex)
  
  If CmbProveedor.ListIndex <> -1 Then
      sSql = sSql & " AND Productos.Proveedor=" & CmbProveedor.ItemData(CmbProveedor.ListIndex) & " "
  End If
  sSql = sSql & " ORDER BY Productos.Producto"
  With Rsd
       .CursorLocation = adUseClient
       .PageSize = PAGE_SIZE
       .Properties("Initial Fetch Size") = PAGE_SIZE
       .Open sSql, cn, adOpenDynamic, adLockOptimistic
       .ActiveConnection = Nothing
       Set Grid1.DataSource = Paginar_Recordset(Rsd)
       CabGrid
  End With
  Registro = PAGE_SIZE
  cn.Close
End Sub

Private Sub CabGrid()
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  Grid1.AllowUpdate = True
  With Grid1
       .Columns(0).Caption = "Produc."
       .Columns(0).Width = 1500
       .Columns(1).Caption = "Descripción"
       .Columns(1).Width = 3650
       .Columns(2).Caption = "Desc. Adicional"
       .Columns(2).Width = 1900
       .Columns(3).Caption = "P.Compr"
       .Columns(3).Width = 1000
       .Columns(3).NumberFormat = nCantDecimales
       .Columns(3).Alignment = dbgRight
       .Columns(4).Caption = "% Venta"
       .Columns(4).Width = 1000
       .Columns(4).NumberFormat = nCantDecimales
       .Columns(4).Alignment = dbgRight
       .Columns(5).Caption = "P.Venta"
       .Columns(5).Width = 1100
       .Columns(5).NumberFormat = nCantDecimales
       .Columns(5).Alignment = dbgRight
       .Columns(6).Caption = "F.U.Modif."
       .Columns(6).Width = 1000
       
  End With
End Sub
Private Function Paginar_Recordset(recset As ADODB.Recordset) As ADODB.Recordset
  
    Dim Fila As Long
    Dim Campo As Field
    Dim P As Long
  
    P = IIf(recset.AbsolutePage > 0, recset.AbsolutePage, 1)
       
    'Recordset desconectado
    With Rs2
           
        ' Si el recordset está abierto entonces lo cierra
        If .State = adStateOpen Then
            .Close
        End If
           
        ' Define y Crea los campos al recordset con Append
        For Each Campo In recset.Fields
            .Fields.Append Campo.Name, IIf(Campo.Type = 131, 6, Campo.Type), _
                           Campo.DefinedSize, Campo.Attributes
        Next Campo
           
        ' Lo Abre para agregarle los registros
           
        Rs2.Open
           
        For Fila = 1 To PAGE_SIZE
               
            ' si está en el final sale
            If recset.EOF Then
                Exit For
            End If
               
            ' va añadiendo los registros
            .AddNew
               
            ' Recorre los campos
             For Each Campo In recset.Fields
                 ' Asigna el nombre del campo
                 Rs2(Campo.Name) = Campo.Value
             Next Campo
                
             'Actualiza el recordset desconectado y el otro recordset
             .Update
             ' mueve al siguiente registro
             recset.MoveNext
  
        Next Fila
        ' Mueve al primer registro
        If Not .BOF Then
           .MoveFirst
            recset.AbsolutePage = P
        End If
        ' Establece la página actual
  
    End With
    ' devuelve a la función el conjunto de datos, es decir el recordset paginado
    Set Paginar_Recordset = Rs2
       
    ' Muestra la página actual en el caption del formulario
    Me.Caption = " Página: " & CStr(Rsd.AbsolutePage) & _
                 " de " & CStr(Rsd.PageCount)
  
End Function

Private Sub Grid1_AfterColEdit(ByVal ColIndex As Integer)
  Dim cRse As ClsEscritura, nPrecio As Currency
  
   Set cRse = New ClsEscritura
  
   cRse.GrabarValordeUnCampo "Precios", "Porcentaje", Grid1.Columns(4).Value, "Producto='" & Grid1.Columns(0).Value & "'"
   nPrecio = (Grid1.Columns(3).Value * Grid1.Columns(4).Value / 100) + Grid1.Columns(3).Value
   cRse.GrabarValordeUnCampo "Productos", "Actualiza", 0, "Producto='" & Grid1.Columns(0).Value & "'"
   cRse.GrabarValordeUnCampo "Productos", "PrecioCompra", Grid1.Columns(3).Value, "Producto='" & Grid1.Columns(0).Value & "'"
   cRse.GrabarValordeUnCampo "Productos", "FechaAlta", "'" & FechaAmericana(Date) & "'", "Producto='" & Grid1.Columns(0).Value & "'"
   cRse.GrabarValordeUnCampo "Precios", "Precio", str(nPrecio), "Producto='" & Grid1.Columns(0).Value & "' and Lista=" & CmbLista.ItemData(CmbLista.ListIndex)
   cRse.GrabarValordeUnCampo "Productos", "FechaCompra", "'" & FechaAmericana(Date) & "'", "Producto='" & Grid1.Columns(0).Value & "'"
   Grid1.Columns(5).Value = IIf(nPrecio = 0, Grid1.Columns(5).Value, nPrecio)
   Grid1.Columns(6).Value = Date
   Rsd.MoveFirst
   Rsd.Find "Producto='" & Grid1.Columns(0).Value & "'"
   Rsd!PrecioCompra = Grid1.Columns(3).Value
   Rsd!FechaCompra = Date
   Rsd!Porcentaje = Grid1.Columns(4).Value
   Rsd!precio = nPrecio
   Rsd.Update
  End Sub

