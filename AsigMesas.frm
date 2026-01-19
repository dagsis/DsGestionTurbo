VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AsigMesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignaciones de Mesa"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   10365
   Begin VB.ComboBox CmbMozo 
      Height          =   315
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   106
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Terminar"
      Height          =   465
      Left            =   9195
      TabIndex        =   0
      Top             =   4665
      Width           =   930
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desasignar Todo"
      Height          =   465
      Left            =   195
      TabIndex        =   105
      Top             =   4665
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   99
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   98
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   97
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   96
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   95
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   94
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   93
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   92
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   91
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   90
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   4020
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   89
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   88
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   87
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   86
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   85
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   84
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   83
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   82
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   81
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   80
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   3615
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   79
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   78
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   77
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   76
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   75
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   74
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   73
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   72
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   71
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   70
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3210
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   69
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   68
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   67
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   66
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   65
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   64
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   63
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   62
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   61
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   60
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2805
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   59
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   58
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   57
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   56
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   55
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   54
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   53
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   52
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   51
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   50
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2400
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   49
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   48
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   47
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   46
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   45
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   44
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   43
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   42
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   41
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   40
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2010
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   39
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   38
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   37
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   36
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   35
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   34
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   33
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   32
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   31
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   30
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1605
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   8
      Left            =   8850
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   375
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   7
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   375
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   6
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   375
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   5
      Left            =   7035
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   375
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   4
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   375
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   3
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   375
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   2
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   375
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   360
      Index           =   1
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   375
      Width           =   570
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   3450
      Left            =   195
      TabIndex        =   3
      Top             =   1005
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   6085
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
            LCID            =   11274
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
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mesas"
      Height          =   4395
      Left            =   3855
      TabIndex        =   1
      Top             =   105
      Width           =   6270
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   29
         Left            =   5595
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   28
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   27
         Left            =   4410
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   26
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   25
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   24
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   23
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   22
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   21
         Left            =   705
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   20
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   19
         Left            =   5595
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   18
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   17
         Left            =   4395
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   16
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   15
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   14
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   13
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   12
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   11
         Left            =   705
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   10
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   675
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   9
         Left            =   5595
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   270
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   360
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   270
         Width           =   570
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Mesas Asignadas :"
      Height          =   210
      Left            =   195
      TabIndex        =   4
      Top             =   750
      Width           =   1650
   End
   Begin VB.Label Label1 
      Caption         =   "Mozo :"
      Height          =   225
      Left            =   195
      TabIndex        =   2
      Top             =   135
      Width           =   1410
   End
End
Attribute VB_Name = "AsigMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMozo As Integer, rMesa As New ADODB.Recordset, sw1 As Boolean
Private Sub CmbMozo_Click()
  Dim cMesa As ClsPrograma, i As Integer

  Set cMesa = New ClsPrograma
  
  Set rMesa = cMesa.DameMozoMesa(CmbMozo.ItemData(CmbMozo.ListIndex))
  Set Grid1.DataSource = rMesa
  CabGrid
  
  For i = 0 To 99
      Command1(i).Caption = i + 1
      Refrescar i + 1
  Next
End Sub
Private Sub Command1_Click(Index As Integer)
  Dim sSql As String, cRsl As ClsLectura, cMesa As ClsPrograma
  
  Set cRsl = New ClsLectura
  Set cMesa = New ClsPrograma
  
  If Not Command1(Index).BackColor = &HFF8080 Then
     If Not Command1(Index).BackColor = &HFF& Then
        sw1 = True
        rMesa.AddNew
        rMesa![Mozo] = CmbMozo.ItemData(CmbMozo.ListIndex)
        rMesa![Mesa] = Index + 1
        rMesa.Update
        Set Grid1.DataSource = rMesa
        CabGrid
        Command1(Index).BackColor = &HFF&
        db.Open sDb
        db.BeginTrans
        
        sSql = "INSERT INTO MozoMesa (Mozo,Mesa) VALUES(" & CmbMozo.ItemData(CmbMozo.ListIndex) & "," & Index + 1 & ")"
        
        db.Execute sSql
        db.CommitTrans
        db.Close

     Else
         db.Open sDb
         sSql = "Delete  from Mozomesa WHERE Mesa=" & Index + 1
        
         db.Execute sSql
         db.Close
         Command1(Index).BackColor = &HC000&
         Set rMesa = cMesa.DameMozoMesa(CmbMozo.ItemData(CmbMozo.ListIndex))
         Set Grid1.DataSource = rMesa
         CabGrid
     End If
  Else
      Dim Rx As New Recordset
      
      Set Rx = cRsl.TraerRsCondi("MozoMesa", "Mozo", "Mesa=" & Command1(Index).Caption)
      If Rx.RecordCount <> 0 Then
         MsgBox "Mesa Asignada al Mozo :" & cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & Rx!Mozo), vbCritical, "Atención"
      End If
  End If
  Exit Sub
    
errHandler:
  db.RollbackTrans
  db.Close
  MsgBox Err.Description, vbCritical, "Atención"
  
End Sub
Private Sub Command2_Click()
  Dim i As Byte, sSql As String
  
  If MsgBox("Borra Todas las Mesas Asignadas a este Mozo!!!!  ?", 20, "Atención") = 6 Then
        db.Open sDb
        
        sSql = "Delete  from Mozomesa WHERE mozo=" & CmbMozo.ItemData(CmbMozo.ListIndex)
        
        db.Execute sSql
        db.Close
        
        Dim cMesa As ClsPrograma

        Set cMesa = New ClsPrograma
  
        Set rMesa = cMesa.DameMozoMesa(CmbMozo.ItemData(CmbMozo.ListIndex))
        Set Grid1.DataSource = rMesa
        CabGrid
        
        For i = 0 To 99
         Command1(i).Caption = i + 1
         Refrescar i + 1
     Next
  End If
End Sub
Private Sub Command3_Click()
'  Dim nMesa As Integer
'  nMesa = rMesa![Mesa]
'  rMesa.Delete
'  Command3.Enabled = False
'  Command1(nMesa - 1).BackColor = &HC000&
'  rMesa.Requery
'  Set Grid1.DataSource = rMesa
'  CabGrid
'  rMesa.MovePrevious
'  If rMesa.BOF Then
'     rMesa.MoveFirst
'  End If
'  sw1 = False
'  Command3.Enabled = False
End Sub
Private Sub Command4_Click()
  Unload Me
End Sub
Private Sub Form_Load()
  Dim cMesa As ClsPrograma, cRsl As ClsLectura
  
  Set cMesa = New ClsPrograma
  Set cRsl = New ClsLectura
  
  Dim i As Byte
  
  CargarCombos
  CmbMozo.ListIndex = cRsl.RegPorDefecto("Vendedor")

  For i = 0 To 99
      Command1(i).Caption = i + 1
      Refrescar i + 1
  Next
  
  sw1 = False

  
End Sub

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbMozo, "Vendedores", "Vendedor", "Descripcion", ""
    
  Set cRsl = Nothing
End Sub
Public Sub Refrescar(nMesa As Byte)
  Dim Rx As Recordset, cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Set Rx = cRsl.TraerRsCondi("MozoMesa", "Mozo", "Mesa=" & nMesa)
  
  If Rx.RecordCount <> 0 Then
     If Rx![Mozo] <> CmbMozo.ItemData(CmbMozo.ListIndex) Then
        Command1(nMesa - 1).BackColor = &HFF8080
     Else
        Command1(nMesa - 1).BackColor = &HFF&
     End If
  Else
     Command1(nMesa - 1).BackColor = &HC000&
  End If
  Rx.Close
End Sub
Private Sub CabGrid()
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Caption = "Mesa Asignada"
      .Columns(0).Width = 3000
      .Columns(0).Alignment = dbgCenter
      .Columns(1).Visible = False
 End With
End Sub
Private Sub Grid1_Click()
  sw1 = False
End Sub
Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If Grid1.Row <> -1 Then
     If sw1 = False Then
        rMesa.Bookmark = Grid1.Bookmark
     End If
  End If
End Sub

