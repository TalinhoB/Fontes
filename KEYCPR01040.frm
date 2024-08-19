VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form KEYCPR01040 
   BackColor       =   &H00FDDEC6&
   BorderStyle     =   0  'None
   Caption         =   "Cotação"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14010
   FillStyle       =   0  'Solid
   Icon            =   "KEYCPR01040.frx":0000
   LinkTopic       =   "KEYCPR01040"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_forNomeCp 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      MaxLength       =   40
      TabIndex        =   92
      Top             =   9360
      Width           =   150
   End
   Begin VB.PictureBox pic_observacao 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   1635
      Left            =   3300
      ScaleHeight     =   1635
      ScaleWidth      =   7875
      TabIndex        =   87
      Top             =   10500
      Visible         =   0   'False
      Width           =   7875
      Begin VB.TextBox txt_Observ 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   88
         Top             =   480
         Width           =   7695
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   7770
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Shape shp_Borda2 
         BorderColor     =   &H80000012&
         Height          =   1635
         Left            =   0
         Top             =   0
         Width           =   7875
      End
      Begin VB.Label lbl_TituloEscObs 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "<ESC> Sair"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   60
         TabIndex        =   90
         Top             =   1320
         Width           =   7665
      End
      Begin VB.Label lbl_obs_titulo01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Observação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   89
         Top             =   60
         Width           =   7665
      End
   End
   Begin VB.TextBox txt_forCodigo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5700
      MaxLength       =   6
      TabIndex        =   86
      Top             =   9360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmd_Limpar 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Limpar"
      Height          =   675
      Left            =   900
      Picture         =   "KEYCPR01040.frx":1EAE2
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7860
      Width           =   855
   End
   Begin VB.Frame frm_Selecao 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Seleção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   60
      TabIndex        =   64
      Top             =   420
      Width           =   13875
      Begin VB.TextBox txt_rqsNumero 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txt_rqsDataRf 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   68
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_rqsStatus 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11460
         TabIndex        =   67
         Top             =   240
         Width           =   1875
      End
      Begin VB.TextBox txt_rqsSolici 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   66
         Top             =   660
         Width           =   4635
      End
      Begin VB.TextBox txt_rqsGerent 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9960
         MaxLength       =   40
         TabIndex        =   65
         Top             =   660
         Width           =   3375
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Gerente:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   8760
         TabIndex        =   72
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Dt.Requisição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   8760
         TabIndex        =   71
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Solicitante:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   70
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Nº Requisição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   69
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame pnl_Requisicoes 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Requisição"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   60
      TabIndex        =   62
      Top             =   1680
      Width           =   13875
      Begin VSFlex8Ctl.VSFlexGrid grd_Dados 
         Height          =   3195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   13635
         _cx             =   268787187
         _cy             =   268768772
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8454143
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16637638
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483630
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame pnl_Cotacoes 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Cotações"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2475
      Left            =   60
      TabIndex        =   61
      Top             =   5280
      Width           =   13875
      Begin VB.ComboBox cbo_cot_cdpCodigo 
         Height          =   315
         Index           =   0
         Left            =   9180
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cbo_cot_cdpCodigo 
         Height          =   315
         Index           =   1
         Left            =   9180
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cbo_cot_cdpCodigo 
         Height          =   315
         Index           =   2
         Left            =   9180
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cbo_cot_cdpCodigo 
         Height          =   315
         Index           =   3
         Left            =   9180
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cbo_cot_cdpCodigo 
         Height          =   315
         Index           =   4
         Left            =   9180
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txt_cot_crqCondPg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   9180
         MaxLength       =   30
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txt_cot_crqCondPg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   9180
         MaxLength       =   30
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txt_cot_crqCondPg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   9180
         MaxLength       =   30
         TabIndex        =   30
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txt_cot_crqCondPg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   9180
         MaxLength       =   30
         TabIndex        =   40
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt_cot_crqCondPg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9180
         MaxLength       =   30
         TabIndex        =   50
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txt_cot_forNomeCp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   40
         TabIndex        =   3
         Top             =   600
         Width           =   3915
      End
      Begin VB.TextBox txt_cot_forCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   6
         TabIndex        =   2
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txt_cot_crqPrUnit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txt_cot_crqPrazoE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   11340
         MaxLength       =   3
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txt_cot_forNomeCp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   40
         TabIndex        =   13
         Top             =   960
         Width           =   3915
      End
      Begin VB.TextBox txt_cot_forCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         MaxLength       =   6
         TabIndex        =   12
         Top             =   960
         Width           =   675
      End
      Begin VB.TextBox txt_cot_crqPrUnit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4920
         MaxLength       =   16
         TabIndex        =   14
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox txt_cot_crqPrazoE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   11340
         MaxLength       =   3
         TabIndex        =   21
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txt_cot_forNomeCp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   40
         TabIndex        =   23
         Top             =   1320
         Width           =   3915
      End
      Begin VB.TextBox txt_cot_forCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   240
         MaxLength       =   6
         TabIndex        =   22
         Top             =   1320
         Width           =   675
      End
      Begin VB.TextBox txt_cot_crqPrUnit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4920
         MaxLength       =   16
         TabIndex        =   24
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txt_cot_crqPrazoE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   11340
         MaxLength       =   3
         TabIndex        =   31
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txt_cot_forNomeCp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   40
         TabIndex        =   33
         Top             =   1680
         Width           =   3915
      End
      Begin VB.TextBox txt_cot_forCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   240
         MaxLength       =   6
         TabIndex        =   32
         Top             =   1680
         Width           =   675
      End
      Begin VB.TextBox txt_cot_crqPrUnit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4920
         MaxLength       =   16
         TabIndex        =   34
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox txt_cot_crqPrazoE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   11340
         MaxLength       =   3
         TabIndex        =   41
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txt_cot_forNomeCp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   960
         MaxLength       =   40
         TabIndex        =   43
         Top             =   2040
         Width           =   3915
      End
      Begin VB.TextBox txt_cot_forCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   240
         MaxLength       =   6
         TabIndex        =   42
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox txt_cot_crqPrUnit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   4920
         MaxLength       =   16
         TabIndex        =   44
         Top             =   2040
         Width           =   1515
      End
      Begin VB.TextBox txt_cot_crqPrazoE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   11340
         MaxLength       =   3
         TabIndex        =   51
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txt_cot_crqAlqIPI 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   5
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqIPI 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6480
         MaxLength       =   16
         TabIndex        =   15
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqIPI 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6480
         MaxLength       =   16
         TabIndex        =   25
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqIPI 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6480
         MaxLength       =   16
         TabIndex        =   35
         Top             =   1680
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqIPI 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6480
         MaxLength       =   16
         TabIndex        =   45
         Top             =   2040
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqICMS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7860
         MaxLength       =   8
         TabIndex        =   7
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqICMS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7860
         MaxLength       =   6
         TabIndex        =   17
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqICMS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   7860
         MaxLength       =   6
         TabIndex        =   27
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqICMS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7860
         MaxLength       =   6
         TabIndex        =   37
         Top             =   1680
         Width           =   915
      End
      Begin VB.TextBox txt_cot_crqAlqICMS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   7860
         MaxLength       =   6
         TabIndex        =   47
         Top             =   2040
         Width           =   915
      End
      Begin VB.CheckBox chk_creditaIPI 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   0
         Left            =   7440
         TabIndex        =   6
         ToolTipText     =   "Credita IPI? Se sim marque a caixa"
         Top             =   600
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaIPI 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   1
         Left            =   7440
         TabIndex        =   16
         Top             =   960
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaIPI 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   2
         Left            =   7440
         TabIndex        =   26
         Top             =   1320
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaIPI 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   3
         Left            =   7440
         TabIndex        =   36
         Top             =   1680
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaIPI 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   4
         Left            =   7440
         TabIndex        =   46
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaICMS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   0
         Left            =   8820
         TabIndex        =   8
         ToolTipText     =   "Credita ICMS? Se sim marque a caixa"
         Top             =   600
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaICMS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   1
         Left            =   8820
         TabIndex        =   18
         ToolTipText     =   "Credita ICMS? Se sim marque a caixa"
         Top             =   960
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaICMS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   2
         Left            =   8820
         TabIndex        =   28
         ToolTipText     =   "Credita ICMS? Se sim marque a caixa"
         Top             =   1320
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaICMS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   3
         Left            =   8820
         TabIndex        =   38
         ToolTipText     =   "Credita ICMS? Se sim marque a caixa"
         Top             =   1680
         Width           =   195
      End
      Begin VB.CheckBox chk_creditaICMS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         Height          =   285
         Index           =   4
         Left            =   8820
         TabIndex        =   48
         ToolTipText     =   "Credita ICMS? Se sim marque a caixa"
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Dias entrega"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   12
         Left            =   11340
         TabIndex        =   81
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Cond.Pgto."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   11
         Left            =   9180
         TabIndex        =   80
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Crd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   8820
         TabIndex        =   79
         Top             =   300
         Width           =   285
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "ICMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   7860
         TabIndex        =   78
         Top             =   300
         Width           =   435
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Crd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   7440
         TabIndex        =   77
         Top             =   300
         Width           =   285
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   6480
         TabIndex        =   76
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Pr.Unitário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   5460
         TabIndex        =   75
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Fornecedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   74
         Top             =   300
         Width           =   960
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   73
         Top             =   300
         Width           =   570
      End
   End
   Begin VB.PictureBox pic_titulo 
      Appearance      =   0  'Flat
      BackColor       =   &H009C832C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5205
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   0
      Width           =   5200
      Begin VB.Image img_fechar 
         Height          =   270
         Left            =   4740
         Picture         =   "KEYCPR01040.frx":1EC2C
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image img_minimizar 
         Height          =   270
         Left            =   4440
         Picture         =   "KEYCPR01040.frx":1F0AE
         ToolTipText     =   "Minimizar"
         Top             =   60
         Width           =   270
      End
      Begin VB.Label lbl_Tituloform 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cotação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   59
         Top             =   60
         Width           =   750
      End
      Begin VB.Label lbl_fundo 
         Appearance      =   0  'Flat
         BackColor       =   &H009C832C&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.PictureBox pic_rodape 
      BackColor       =   &H009C832C&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5205
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   8580
      Width           =   5200
      Begin VB.Label lbl_rodape 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   56
         Top             =   60
         Width           =   1680
      End
      Begin VB.Label lbl_fundo 
         BackColor       =   &H009C832C&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.CommandButton cmd_sair 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Sair"
      Height          =   675
      Left            =   13080
      Picture         =   "KEYCPR01040.frx":1F530
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7860
      Width           =   855
   End
   Begin VB.CommandButton cmd_Confirmar 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Confirmar"
      Height          =   675
      Left            =   60
      Picture         =   "KEYCPR01040.frx":1F67A
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7860
      Width           =   855
   End
   Begin VB.Label lbl_Titulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDEC6&
      Caption         =   "<F2> Fornecedores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   17
      Left            =   7080
      TabIndex        =   91
      Top             =   7920
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lbl_Titulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDEC6&
      Caption         =   "<F3> Observações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   16
      Left            =   4200
      TabIndex        =   85
      Top             =   8220
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lbl_Titulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDEC6&
      Caption         =   "<F7> Normas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   15
      Left            =   10140
      TabIndex        =   84
      Top             =   8220
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lbl_Titulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDEC6&
      Caption         =   "<F6> Média de Compras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   14
      Left            =   7980
      TabIndex        =   83
      Top             =   8220
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.Label lbl_Titulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDEC6&
      Caption         =   "<F5> Últimas Compras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   13
      Left            =   5880
      TabIndex        =   82
      Top             =   8220
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDEC6&
      Caption         =   "<F5> Últimas Compras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Shape shp_borda 
      BorderColor     =   &H80000012&
      Height          =   315
      Left            =   5340
      Top             =   8640
      Width           =   255
   End
End
Attribute VB_Name = "KEYCPR01040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WMTyp_CtReqCompras()    As typ_AlmoxCTReqCompras

Dim WMBol_Inicio            As Boolean
Dim WMBol_Controle          As Boolean
Dim WMBol_Pesquisa          As Boolean

Dim WMInt_Index             As Integer

'Grid Requisição
Private Enum eCol
    Selecao = 0
    Grupo = 1
    DescGrupo = 2
    Produto = 3
    DescProd = 4
    Situacao = 5
    DtNecess = 6
    Quantidade = 7
    Norma = 8
    EmpMrc = 9
    DescMarca = 10
    Observacao = 11
    LastCol = 11
End Enum

Private Enum eTecla
    Observacao = 16
    UltCompras = 13
    MedCompras = 14
    Normas = 15
    Fornecedores = 17
End Enum

Private Sub cbo_cot_cdpCodigo_GotFocus(Index As Integer)
    cbo_cot_cdpCodigo(Index).BackColor = vbYellow
End Sub

Private Sub cbo_cot_cdpCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    End If
End Sub

Private Sub cbo_cot_cdpCodigo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub cbo_cot_cdpCodigo_LostFocus(Index As Integer)
    cbo_cot_cdpCodigo(Index).BackColor = vbWhite
End Sub

Private Sub chk_creditaICMS_GotFocus(Index As Integer)
    chk_creditaICMS(Index).BackColor = vbYellow
End Sub

Private Sub chk_creditaICMS_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    End If
End Sub

Private Sub chk_creditaICMS_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub chk_creditaICMS_LostFocus(Index As Integer)
    chk_creditaICMS(Index).BackColor = vbWhite
End Sub

Private Sub chk_creditaIPI_GotFocus(Index As Integer)
    chk_creditaIPI(Index).BackColor = vbYellow
End Sub

Private Sub chk_creditaIPI_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    End If
End Sub

Private Sub chk_creditaIPI_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub chk_creditaIPI_LostFocus(Index As Integer)
    chk_creditaIPI(Index).BackColor = vbWhite
End Sub

Private Sub cmd_confirmar_Click()
    Dim WLLng_Long  As Long
    
    On Local Error GoTo Erro
    
    If txt_rqsNumero.Text = Empty Then Exit Sub
    If Not MFcn_VerificaSelecao(grd_Dados) Then Exit Sub
    If Not MFcn_ValidaGravacao Then Exit Sub
    
    WGCnx_DBPrim.BeginTrans
    
    If Not MFcn_GravarDados Then
        WGCnx_DBPrim.RollbackTrans
        Exit Sub
    End If
    
    GoTo Fim
Erro:
    WGCnx_DBPrim.RollbackTrans
    GFKEY_MsgBox "Erro ao gravar cotação!" & vbCrLf & vbCrLf & GFKEY_TrataErro(Err.Number, Err.Description)
    Exit Sub
Fim:
    WGCnx_DBPrim.CommitTrans
    GFKEY_MsgBox "Gravação Ok!"
    
    Call MPrc_LimpaCamposCotacao("T")
    Call MPrc_Habilita(False)
    Call MFcn_CarregaDadosReq
End Sub
Private Sub cmd_confirmar_GotFocus()
    cmd_Confirmar.BackColor = vbYellow
End Sub
Private Sub cmd_confirmar_LostFocus()
    cmd_Confirmar.BackColor = keyCBackColor
End Sub

Private Sub cmd_Limpar_Click()
    Call MPrc_LimpaSelecao
    Call MPrc_LimpaCamposCotacao("T")
    Call MPrc_Habilita(False)
    pnl_Cotacoes.Enabled = False
    txt_rqsNumero.SetFocus
End Sub

Private Sub cmd_Limpar_GotFocus()
    cmd_Limpar.BackColor = vbYellow
End Sub

Private Sub cmd_Limpar_LostFocus()
    cmd_Limpar.BackColor = keyCBackColor
End Sub

Private Sub cmd_sair_Click()
    Unload Me
End Sub

Private Sub cmd_sair_GotFocus()
    cmd_sair.BackColor = vbYellow
End Sub

Private Sub cmd_sair_LostFocus()
    cmd_sair.BackColor = keyCBackColor
End Sub

Private Sub Form_Activate()
    
    On Local Error GoTo Erro
    '
    If WMBol_Inicio = True Then
        WMBol_Inicio = False
        Call GPKEY_CantoRedondo(Me, 25)
        
        If WGBol_CPRCadCondPg Then
            Call MPrc_BuscaCondPgtos
        End If
    End If
    
    GoTo Fim
Erro:
    If Err.Number <> 0 Then
        GFKEY_MsgBox "Erro ao iniciar programa" & vbCrLf & vbCrLf & GFKEY_TrataErro(Err.Number, Err.Description)
    End If
    Unload Me
Fim:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call MPrc_VerTeclaAtalho(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub Form_Load()
    WMBol_Inicio = True
End Sub

Private Sub grd_Dados_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grd_Dados
        If .TextMatrix(Row, eCol.Selecao) = True Then
            Call MPrc_Habilita(True)
            .Cell(flexcpBackColor, Row, eCol.Grupo, Row, eCol.LastCol) = keyCAmarelo
            
            If Not MFcn_BuscaCotacao(Row) Then
                Call MPrc_LimpaCamposCotacao("T")
            End If
        Else
            .Cell(flexcpBackColor, Row, eCol.Grupo, Row, eCol.LastCol) = keyCBranco
            Call MPrc_Habilita(False)
            Call MPrc_LimpaCamposCotacao("T")
        End If
    End With
End Sub

Private Sub grd_Dados_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With grd_Dados
        If .rows > 1 And NewRow > 0 Then
            If OldRow <> NewRow And OldRow > 0 Then
                .TextMatrix(OldRow, eCol.Selecao) = False
                .Cell(flexcpBackColor, OldRow, eCol.Grupo, OldRow, eCol.LastCol) = keyCBranco
            End If
        End If
    End With
End Sub

Private Sub grd_Dados_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> eCol.Selecao Then
        Cancel = True
    End If
End Sub

Private Sub grd_Dados_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub lbl_Titulo_Click(Index As Integer)
    Select Case Index
        Case 0
            txt_rqsNumero.Text = ""
            Set WGObj_FormPsqRqsPend = Me
            
            With PSQRQSPENDENTES
                .Params.Status = "'N'"
                .Show vbModal
            End With

            Set WGObj_FormPsqRqsPend = Nothing
            
            If txt_rqsNumero.Text <> Empty Then
                Call MFcn_CarregaDadosReq
            End If
            
        Case 13
            Call MPrc_VerTeclaAtalho(vbKeyF5)
            
        Case 14
            Call MPrc_VerTeclaAtalho(vbKeyF6)
            
        Case 15
            Call MPrc_VerTeclaAtalho(vbKeyF7)
            
        Case 16
            Call MPrc_VerTeclaAtalho(vbKeyF3)
        
        Case 17
            Call MPrc_VerTeclaAtalho(vbKeyF2, WMInt_Index)
            
    End Select
End Sub

Private Sub pic_ucp_KeyUp(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub img_fechar_Click()
    Unload Me
End Sub

Private Sub img_minimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub lbl_fundo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call GPKEY_MoveTela(Me)
End Sub

Private Sub lbl_rodape_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call GPKEY_MoveTela(Me)
End Sub

Private Sub lbl_Tituloform_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call GPKEY_MoveTela(Me)
End Sub

Private Sub pic_rodape_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call GPKEY_MoveTela(Me)
End Sub

Private Sub pic_titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call GPKEY_MoveTela(Me)
End Sub

Private Sub MPrc_MontaGrid()
    '
    Dim WLStr_TituloCols    As String
    
    On Local Error GoTo Erro
    
    WLStr_TituloCols = " |Grp|Descrição|Código|Descrição|Situação|Dt.Necessidade|" & _
        "Quantidade|Norma Técnica|Emp/Mrc|Descrição Marca|Observação"
        
    With grd_Dados
        .Clear
        .Refresh
        .rows = 1
        .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AllowUserResizing = flexResizeColumns
        .FormatString = WLStr_TituloCols
        .ScrollTrack = True
        .MergeCells = flexMergeFree
        .AllowBigSelection = False
        .FocusRect = flexFocusLight
        .Cols = eCol.LastCol + 1
        
        .ColWidth(eCol.Selecao) = 255
        .ColDataType(eCol.Selecao) = flexDTBoolean
        .FixedAlignment(eCol.Selecao) = flexAlignCenterCenter
        .ColAlignment(eCol.Selecao) = flexAlignCenterCenter
        
        .ColWidth(eCol.Grupo) = 500
        .ColDataType(eCol.Grupo) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.Grupo, .rows - 1, eCol.Grupo) = flexAlignLeftTop
        
        .ColWidth(eCol.DescGrupo) = IIf(WGBol_CPRDtNecessid = True, 2300, 3600)
        .ColDataType(eCol.DescGrupo) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.DescGrupo, .rows - 1, eCol.DescGrupo) = flexAlignLeftTop
        
        .ColWidth(eCol.Produto) = 1200
        .ColDataType(eCol.Produto) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.Produto, .rows - 1, eCol.Produto) = flexAlignLeftTop
        
        .ColWidth(eCol.DescProd) = IIf(WGBol_CPRDtNecessid = True, 3600, 4035)
        .ColDataType(eCol.DescProd) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.DescProd, .rows - 1, eCol.DescProd) = flexAlignLeftTop
        
        .ColWidth(eCol.Situacao) = 1200
        .ColDataType(eCol.Situacao) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.Situacao, .rows - 1, eCol.Situacao) = flexAlignLeftTop
        
        .ColWidth(eCol.DtNecess) = 1470
        .ColDataType(eCol.DtNecess) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.DtNecess, .rows - 1, eCol.DtNecess) = flexAlignCenterTop
        .ColHidden(eCol.DtNecess) = IIf(WGBol_CPRDtNecessid = True, False, True)
        
        .ColWidth(eCol.Quantidade) = IIf(WGBol_CPRDtNecessid = True, 1200, 1395)
        .ColDataType(eCol.Quantidade) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.Quantidade, .rows - 1, eCol.Quantidade) = flexAlignLeftTop
        
        .ColWidth(eCol.Norma) = 800
        .ColDataType(eCol.Norma) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.Norma, .rows - 1, eCol.Norma) = flexAlignLeftTop
        .ColHidden(eCol.Norma) = True
        
        .ColWidth(eCol.EmpMrc) = 1000
        .ColDataType(eCol.EmpMrc) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.EmpMrc, .rows - 1, eCol.EmpMrc) = flexAlignLeftTop
        
        .ColWidth(eCol.DescMarca) = 1500
        .ColDataType(eCol.DescMarca) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.DescMarca, .rows - 1, eCol.DescMarca) = flexAlignLeftTop
        .ColHidden(eCol.DescMarca) = True
        
        .ColWidth(eCol.Observacao) = 500
        .ColDataType(eCol.Observacao) = flexDTString
        .Cell(flexcpAlignment, 0, eCol.Observacao, .rows - 1, eCol.Observacao) = flexAlignLeftTop
        .ColHidden(eCol.Observacao) = True
        
        .Cell(flexcpFontBold, .rows - 1, eCol.Selecao, .rows - 1, eCol.LastCol) = True
    End With
    '
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
Fim:
End Sub


Private Function MFcn_MontaQuery() As String
    Dim WLStr_Sql       As String
    Dim WLStr_Colunas   As String
   
    On Local Error GoTo Erro
    
    MFcn_MontaQuery = ""
    
    WLStr_Colunas = _
        "b.emp_Codigo,b.mrc_Codigo,a.rqs_Numero,a.rqs_DataRf, " & _
        "a.rqs_Obser1,a.rqs_Obser2,a.rqs_Obser3,a.rqs_Solici, " & _
        "a.rqs_Gerent,a.rqs_Aprova,b.gpx_Codigo,b.prx_Codigo, " & _
        "b.nor_Codigo,b.irq_DtNece,b.irq_Quanti,b.irq_Status, " & _
        "d.prx_Descri,f.mrc_Descri,e.gpx_Descri "
    
    WLStr_Sql = "select  " & _
            WLStr_Colunas & _
        "from " & _
            "ALMOX_REQ_COMPRAS a  " & _
            "left join ALMOX_ITREQ_COMPRAS b ON b.rqs_Numero = a.rqs_Numero " & _
            "left join ALMOX_CTREQ_COMPRAS c ON c.emp_Codigo = b.emp_Codigo and  " & _
                "c.mrc_Codigo = b.mrc_Codigo and c.rqs_Numero = b.rqs_Numero and  " & _
                "c.gpx_Codigo = b.gpx_Codigo and c.prx_Codigo = b.prx_Codigo and " & _
                "c.crq_DtNece = b.irq_DtNece " & _
            "left join PROD_ALMOXARIFADO d ON d.emp_Codigo = b.emp_Codigo and " & _
                "d.mrc_Codigo = b.mrc_Codigo and d.gpx_Codigo = b.gpx_Codigo and  " & _
                "d.prx_Codigo = b.prx_Codigo " & _
            "left join GRUPOS_ALMOXARIFADO e ON e.gpx_Codigo = d.gpx_Codigo " & _
            "left join MARCAS f ON f.emp_Codigo = b.emp_Codigo and f.mrc_Codigo = b.mrc_Codigo " & _
        "where " & _
            "a.rqs_Numero = '" & Trim(txt_rqsNumero.Text) & "' and  " & _
            "trim(a.rqs_Aprova) = '' " & _
        "group by rqs_Numero,gpx_Codigo,prx_Codigo,irq_DtNece " & _
        "order by gpx_Codigo , prx_Codigo , irq_DtNece "
    
    GoTo Fim
Erro:
    Exit Function
Fim:
    MFcn_MontaQuery = WLStr_Sql
End Function
Private Function MFcn_CarregaDadosReq() As Boolean
    Dim WLRst_Tabela        As ADODB.Recordset
    Dim WLObj_DtNeces       As DataHora
    Dim WLObj_DataReq       As DataHora
    Dim WLLng_TotRow        As Long
    Dim WLLng_Cont          As Long
    
    Dim WLStr_Sql           As String
    Dim WLStr_Observ        As String
    
    Dim WLBol_Cotada        As Boolean
    Dim WLBol_NaoCotada     As Boolean
    Dim WLBol_CotadaParc    As Boolean
    
    On Local Error GoTo Erro
    
    MFcn_CarregaDadosReq = False
    
    Call MPrc_MontaGrid
    
    WLStr_Sql = MFcn_MontaQuery
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
    
    With grd_Dados
        If WLLng_TotRow > 0 Then
            WLRst_Tabela.MoveFirst
            Do While Not WLRst_Tabela.EOF
                .rows = .rows + 1
                
                Set WLObj_DtNeces = New DataHora
                WLObj_DtNeces.Init WLRst_Tabela!irq_DtNece
                
                .TextMatrix(.rows - 1, eCol.Selecao) = False
                .TextMatrix(.rows - 1, eCol.Grupo) = Format(Trim(WLRst_Tabela!gpx_Codigo), "00")
                .TextMatrix(.rows - 1, eCol.DescGrupo) = Trim$(WLRst_Tabela!gpx_Descri)
                .TextMatrix(.rows - 1, eCol.Produto) = Trim$(WLRst_Tabela!prx_Codigo)
                .TextMatrix(.rows - 1, eCol.DescProd) = Trim$(WLRst_Tabela!prx_Descri)
                .TextMatrix(.rows - 1, eCol.DtNecess) = WLObj_DtNeces.Data
                .TextMatrix(.rows - 1, eCol.Quantidade) = NumeroHelper.Formatar(WLRst_Tabela!irq_Quanti, 3)
                .TextMatrix(.rows - 1, eCol.Norma) = Format(WLRst_Tabela!nor_Codigo, "00000")
                .TextMatrix(.rows - 1, eCol.EmpMrc) = WLRst_Tabela!emp_Codigo & WLRst_Tabela!mrc_Codigo
                .TextMatrix(.rows - 1, eCol.DescMarca) = WLRst_Tabela!mrc_Descri
                .TextMatrix(.rows - 1, eCol.Observacao) = WLRst_Tabela!rqs_Obser1 & Space(1) & _
                    WLRst_Tabela!rqs_Obser2 & Space(1) & WLRst_Tabela!rqs_Obser3
                
                Select Case WLRst_Tabela!irq_Status
                    Case "F"
                        .TextMatrix(.rows - 1, eCol.Situacao) = "Comprada"
                        
                    Case "C"
                        .TextMatrix(.rows - 1, eCol.Situacao) = "Cotada"
                        WLBol_Cotada = True
                        
                    Case "A"
                        .TextMatrix(.rows - 1, eCol.Situacao) = "Aprovada"
                        WLBol_Cotada = True
                        
                    Case Else
                        .TextMatrix(.rows - 1, eCol.Situacao) = "Não Cotada"
                        WLBol_NaoCotada = True
                        
                End Select
                
                Set WLObj_DataReq = New DataHora
                WLObj_DataReq.Init WLRst_Tabela!rqs_DataRf
                
                txt_rqsSolici.Text = Trim$(WLRst_Tabela!rqs_Solici)
                txt_rqsDataRf.Text = WLObj_DataReq.Data
                txt_rqsGerent.Text = Trim$(WLRst_Tabela!rqs_Gerent)
                
                WLRst_Tabela.MoveNext
            Loop
        Else
            GFKEY_MsgBox "Nenhum dado encontrado nas condições informadas! "
            Exit Function
        End If
        
        If WLBol_Cotada = True And WLBol_NaoCotada = True Then
            WLBol_CotadaParc = True
        End If
        
        If WLBol_CotadaParc Then
            txt_rqsStatus.Text = "Cotada Parcial"
            txt_rqsStatus.BackColor = vbYellow
            
        ElseIf WLBol_NaoCotada Then
            txt_rqsStatus.Text = "Não Cotada"
            txt_rqsStatus.BackColor = keycVermelhoBC
            txt_rqsStatus.ForeColor = vbWhite
            
        ElseIf WLBol_Cotada Then
            txt_rqsStatus.Text = "Cotada"
            txt_rqsStatus.BackColor = vbGreen
        End If
        
    End With
    
    pnl_Requisicoes.Enabled = True
    grd_Dados.Enabled = True
    grd_Dados.SetFocus
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    MFcn_CarregaDadosReq = True
End Function

Private Sub MPrc_VerTeclaAtalho(KeyAscii As Integer, Optional Index As Integer)
    Select Case KeyAscii
        Case vbKeyF2
            If lbl_Titulo(eTecla.Fornecedores).Visible = True Then
                Call MPrc_ChamaPesqFornec(Index)
            End If
            
        Case vbKeyF5
            If lbl_Titulo(eTecla.UltCompras).Visible = True Then
                Call MPrc_ChamaPesqUCP
            End If
        
        Case vbKeyEscape
            pic_observacao.Visible = False
        
        Case vbKeyF3
            WMBol_Pesquisa = True
            
            If lbl_Titulo(eTecla.Observacao).Visible = True Then
                txt_Observ.Text = grd_Dados.TextMatrix(grd_Dados.Row, eCol.Observacao)
                
                With pic_observacao
                    .Top = ((Me.Height - .Height) / 2)
                    .Left = ((Me.Width - .Width) / 2)
                    .Visible = True
                End With
            End If
            
            WMBol_Pesquisa = False
            
        Case vbKeyF6
            If lbl_Titulo(eTecla.MedCompras).Visible = True Then
                Call MPrc_ChamaPesqMedCompr
            End If
            
        Case vbKeyF7
            If lbl_Titulo(eTecla.Normas).Visible = True Then
                Call MPrc_ChamaPesqNorma
            End If
            
    End Select
End Sub
Private Sub MPrc_ChamaPesqUCP()

    With PSQULTCOMPR
        WMBol_Pesquisa = True
        .Params.EmpCod = Left$(grd_Dados.TextMatrix(grd_Dados.Row, eCol.EmpMrc), 2)
        .Params.MrcCod = Right$(grd_Dados.TextMatrix(grd_Dados.Row, eCol.EmpMrc), 2)
        .Params.MrcDescri = grd_Dados.TextMatrix(grd_Dados.Row, eCol.DescMarca)
        .Params.PrxCod = grd_Dados.TextMatrix(grd_Dados.Row, eCol.Produto)
        .Params.GpxCod = grd_Dados.TextMatrix(grd_Dados.Row, eCol.Grupo)
        .Params.GpxDescri = grd_Dados.TextMatrix(grd_Dados.Row, eCol.DescGrupo)
        .Params.PrxDescri = grd_Dados.TextMatrix(grd_Dados.Row, eCol.DescProd)
        .Params.Quantidade = grd_Dados.TextMatrix(grd_Dados.Row, eCol.Quantidade)
        
        .Show vbModal
        
        WMBol_Pesquisa = False
    End With
    
End Sub

Private Sub txt_cot_crqAlqICMS_GotFocus(Index As Integer)
    With txt_cot_crqAlqICMS(Index)
        .BackColor = vbYellow
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txt_cot_crqAlqICMS_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    Else
        If KeyAscii <> vbKeyBack And InStr("0123456789.,", UCase(Chr(KeyAscii))) <= 0 Then KeyAscii = 0
        
        If Chr(KeyAscii) = "," Then
            KeyAscii = IIf(InStr(txt_cot_crqAlqICMS(Index).Text, ".") > 0, 0, Asc("."))
        End If
        
        If InStr("0123456789.", Chr(KeyAscii)) > 0 And GFKEY_LimiteValor(KeyAscii, txt_cot_crqAlqICMS(Index), 3, 2) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_cot_crqAlqICMS_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub txt_cot_crqAlqICMS_LostFocus(Index As Integer)
    txt_cot_crqAlqICMS(Index).BackColor = vbWhite
    
    If txt_cot_crqAlqICMS(Index).Text <> Empty Then
        txt_cot_crqAlqICMS(Index).Text = NumeroHelper.Formatar(txt_cot_crqAlqICMS(Index).Text, 2)
    End If
End Sub

Private Sub txt_cot_crqAlqIPI_GotFocus(Index As Integer)
    
    With txt_cot_crqAlqIPI(Index)
        .BackColor = vbYellow
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txt_cot_crqAlqIPI_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    Else
        If KeyAscii <> vbKeyBack And InStr("0123456789.,", UCase(Chr(KeyAscii))) <= 0 Then KeyAscii = 0
        
        If Chr(KeyAscii) = "," Then
            KeyAscii = IIf(InStr(txt_cot_crqAlqIPI(Index).Text, ".") > 0, 0, Asc("."))
        End If
        
        If InStr("0123456789.", Chr(KeyAscii)) > 0 And GFKEY_LimiteValor(KeyAscii, txt_cot_crqAlqIPI(Index), 3, 2) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_cot_crqAlqIPI_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub txt_cot_crqAlqIPI_LostFocus(Index As Integer)
    txt_cot_crqAlqIPI(Index).BackColor = vbWhite
    
    If txt_cot_crqAlqIPI(Index).Text <> Empty Then
        txt_cot_crqAlqIPI(Index).Text = NumeroHelper.Formatar(txt_cot_crqAlqIPI(Index).Text, 2)
    End If
End Sub

Private Sub txt_cot_crqCondPg_GotFocus(Index As Integer)
    txt_cot_crqCondPg(Index).BackColor = vbYellow
End Sub

Private Sub txt_cot_crqCondPg_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    End If
End Sub

Private Sub txt_cot_crqCondPg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub txt_cot_crqCondPg_LostFocus(Index As Integer)
    txt_cot_crqCondPg(Index).BackColor = vbWhite
End Sub

Private Sub txt_cot_crqPrazoE_GotFocus(Index As Integer)
    txt_cot_crqPrazoE(Index).BackColor = vbYellow
End Sub

Private Sub txt_cot_crqPrazoE_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    
    ElseIf KeyAscii <> vbKeyBack Then
        If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        
    End If
End Sub

Private Sub txt_cot_crqPrazoE_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub txt_cot_crqPrazoE_LostFocus(Index As Integer)
    txt_cot_crqPrazoE(Index).BackColor = vbWhite
End Sub

Private Sub txt_cot_crqPrUnit_GotFocus(Index As Integer)
    With txt_cot_crqPrUnit(Index)
        .BackColor = vbYellow
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txt_cot_crqPrUnit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    Else
        If KeyAscii <> vbKeyBack And InStr("0123456789.,", UCase(Chr(KeyAscii))) <= 0 Then KeyAscii = 0
        
        If Chr(KeyAscii) = "," Then
            KeyAscii = IIf(InStr(txt_cot_crqPrUnit(Index).Text, ".") > 0, 0, Asc("."))
        End If
        
        If InStr("0123456789.", Chr(KeyAscii)) > 0 And GFKEY_LimiteValor(KeyAscii, txt_cot_crqPrUnit(Index), 11, 4) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_cot_crqPrUnit_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub txt_cot_crqPrUnit_LostFocus(Index As Integer)
    txt_cot_crqPrUnit(Index).BackColor = vbWhite
    
    If txt_cot_crqPrUnit(Index).Text <> Empty Then
        txt_cot_crqPrUnit(Index).Text = NumeroHelper.Formatar(txt_cot_crqPrUnit(Index).Text, 4)
    End If
End Sub

Private Sub txt_cot_forCodigo_GotFocus(Index As Integer)
    txt_cot_forCodigo(Index).BackColor = vbYellow
    lbl_Titulo(eTecla.Fornecedores).Visible = True
    WMInt_Index = Index
    WMBol_Controle = True
End Sub

Private Sub txt_cot_forCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    Else
        If KeyAscii = vbKeyExecute Then
            KeyAscii = 0
            If Not WGObj_FormPsqFAlmox Is Nothing Then
                Beep
                GFKEY_MsgBox "Pesquisa de fornecedores já ativa em outro módulo!"
                DoEvents
            Else
                WMBol_Pesquisa = True
                txt_forCodigo.Tag = Format(Index, "#0")
                Set WGObj_FormPsqFAlmox = Me
                
                Me.Enabled = False
                PSQFORNECALMOX.Show vbModal
                
                txt_cot_forCodigo(Index).Text = txt_forCodigo.Text
                txt_cot_forCodigo(Index).SetFocus
                txt_forCodigo.Text = ""
                WMBol_Pesquisa = False
                
                DoEvents
            End If
        Else
            If KeyAscii <> vbKeyBack And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_cot_forCodigo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode, Index)
End Sub

Private Sub txt_cot_forCodigo_LostFocus(Index As Integer)
    Dim WLInt_IndexNovo As Integer
    
    txt_cot_forCodigo(Index).BackColor = vbWhite
    
    WLInt_IndexNovo = Index
    
    If txt_cot_forCodigo(Index).Text <> Empty Then
        txt_cot_forCodigo(Index).Text = Format(Val(Trim(txt_cot_forCodigo(Index).Text)), "000000")
        
        If Not MFcn_LerFornecedor(Index) Then
            Call MPrc_LimpaCamposCotacao("P", Index)
        End If
    Else
        If WMBol_Pesquisa = False Then
            cmd_Confirmar.SetFocus
        End If
    End If
    
    lbl_Titulo(eTecla.Fornecedores).Visible = False
End Sub

Private Sub txt_cot_forNomeCp_GotFocus(Index As Integer)
     txt_cot_forNomeCp(Index).BackColor = vbYellow
End Sub

Private Sub txt_cot_forNomeCp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    End If
End Sub

Private Sub txt_cot_forNomeCp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub txt_cot_forNomeCp_LostFocus(Index As Integer)
    txt_cot_forNomeCp(Index).BackColor = vbWhite
End Sub

Private Sub txt_rqsNumero_GotFocus()
    With txt_rqsNumero
        .BackColor = keyCAmarelo
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txt_rqsNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
    Else
        If KeyAscii = vbKeyExecute And InStr(Me.Caption, "- Incluir") <= 0 Then
            KeyAscii = 0
            If Not WGObj_FormPsqRqsPend Is Nothing Then
                Beep
                GFKEY_MsgBox "Pesquisa de requisições pendentes já ativa por outro módulo!"
                txt_rqsNumero.SetFocus
                DoEvents
            Else
                txt_rqsNumero.Text = ""
                Set WGObj_FormPsqRqsPend = Me
                Me.Enabled = False
                
                With PSQRQSPENDENTES
                    .Params.Status = "'N'"
                    .Show vbModal
                End With
                
                Set WGObj_FormPsqRqsPend = Nothing
                If txt_rqsNumero.Text <> Empty Then
                    Call MFcn_CarregaDadosReq
                End If
            End If
        Else
            If KeyAscii <> vbKeyBack And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_rqsNumero_KeyUp(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub txt_rqsNumero_LostFocus()
    txt_rqsNumero.BackColor = keyCBranco
    
    If txt_rqsNumero.Text <> Empty Then
        txt_rqsNumero.Text = Format(txt_rqsNumero.Text, "000000")
        
        Call MPrc_LimpaCamposCotacao("T")
        
        If Not MFcn_CarregaDadosReq Then
            txt_rqsNumero.Text = vbNullString
            txt_rqsNumero.SetFocus
        Else
            grd_Dados.SetFocus
        End If
    End If
End Sub
Private Sub MPrc_LimpaSelecao()
    Call MPrc_MontaGrid
    txt_rqsNumero.Text = vbNullString
    txt_rqsSolici.Text = vbNullString
    txt_rqsDataRf.Text = vbNullString
    txt_rqsStatus.Text = vbNullString
    txt_rqsGerent.Text = vbNullString
    
    txt_rqsStatus.BackColor = vbWhite
End Sub
Private Sub MPrc_LimpaCamposCotacao(PPStr_Tipo As String, Optional Index As Integer)
    Dim WLLng_Cont  As Long
    
    If PPStr_Tipo = "T" Then
        For WLLng_Cont = 0 To txt_cot_forCodigo.UBound
            txt_cot_forCodigo(WLLng_Cont).Text = vbNullString
            txt_cot_forNomeCp(WLLng_Cont).Text = vbNullString
            txt_cot_crqPrUnit(WLLng_Cont).Text = vbNullString
            txt_cot_crqAlqIPI(WLLng_Cont).Text = vbNullString
            txt_cot_crqAlqICMS(WLLng_Cont).Text = vbNullString
            txt_cot_crqCondPg(WLLng_Cont).Text = vbNullString
            txt_cot_crqPrazoE(WLLng_Cont).Text = vbNullString
            chk_creditaIPI(WLLng_Cont).Value = vbUnchecked
            chk_creditaICMS(WLLng_Cont).Value = vbUnchecked
            cbo_cot_cdpCodigo(WLLng_Cont).ListIndex = -1
        Next
    Else
        txt_cot_forCodigo(Index).Text = vbNullString
        txt_cot_forNomeCp(Index).Text = vbNullString
        txt_cot_crqPrUnit(Index).Text = vbNullString
        txt_cot_crqAlqIPI(Index).Text = vbNullString
        txt_cot_crqAlqICMS(Index).Text = vbNullString
        txt_cot_crqCondPg(Index).Text = vbNullString
        txt_cot_crqPrazoE(Index).Text = vbNullString
        chk_creditaIPI(Index).Value = vbUnchecked
        chk_creditaICMS(Index).Value = vbUnchecked
        cbo_cot_cdpCodigo(Index).ListIndex = -1
    End If
End Sub
Private Function MFcn_BuscaCotacao(PFLng_Row As Long) As Boolean
    Dim WLRst_Tabela    As ADODB.Recordset
    
    Dim WLLng_TotRow    As Long
    Dim WLLng_Cont      As Long
    Dim WLLng_Cont2     As Long
    
    Dim WLStr_Sql       As String
    Dim WLStr_EmpCod    As String
    Dim WLStr_MrcCod    As String
    Dim WLStr_ReqNum    As String
    Dim WLStr_GpxCod    As String
    Dim WLStr_PrxCod    As String
    Dim WLStr_DtNece    As String
    
    On Local Error GoTo Erro
    
    MFcn_BuscaCotacao = False
    
    WLStr_EmpCod = Left$(grd_Dados.TextMatrix(PFLng_Row, eCol.EmpMrc), 2)
    WLStr_MrcCod = Right$(grd_Dados.TextMatrix(PFLng_Row, eCol.EmpMrc), 2)
    WLStr_GpxCod = grd_Dados.TextMatrix(PFLng_Row, eCol.Grupo)
    WLStr_PrxCod = grd_Dados.TextMatrix(PFLng_Row, eCol.Produto)
    WLStr_DtNece = Format$(grd_Dados.TextMatrix(PFLng_Row, eCol.DtNecess), "yyyymmdd")
    
    WLStr_ReqNum = Trim$(txt_rqsNumero.Text)
    
    WLStr_Sql = "select " & _
            "a.rqs_Numero,a.gpx_Codigo,a.prx_Codigo,a.crq_DtNece," & _
            "a.for_Codigo,b.for_NomeCp,a.crq_PrUnit,a.crq_AlqIPI," & _
            "a.crq_PrazoE,a.crq_CondPg,a.cdp_Codigo,c.prf_PrUnit," & _
            "ifnull(c.prf_Descon,'') prf_Descon,ifnull(a.crq_AlqICMS,'') crq_AlqICMS," & _
            "a.crq_CrdIPI,a.crq_CrdICMS " & _
        "from " & _
            "ALMOX_CTREQ_COMPRAS a " & _
                "join " & _
            "FORNECEDORES b on b.for_Codigo = a.for_Codigo " & _
                "left join " & _
            "PROD_FORNECEDOR c on c.emp_Codigo = a.emp_Codigo and " & _
                "c.mrc_Codigo = a.mrc_Codigo and " & _
                "c.gpx_Codigo = a.gpx_Codigo and " & _
                "c.prx_Codigo = a.prx_Codigo and " & _
                "c.for_Codigo = a.for_Codigo " & _
        "where " & _
            "a.emp_Codigo = " & WLStr_EmpCod & " and " & _
            "a.mrc_Codigo = " & WLStr_MrcCod & " and " & _
            "a.rqs_Numero = " & WLStr_ReqNum & " and " & _
            "a.gpx_Codigo = " & WLStr_GpxCod & " and " & _
            "a.prx_Codigo = " & WLStr_PrxCod & " "
    
    If WGBol_CPRDtNecessid = True And Trim(grd_Dados.TextMatrix(PFLng_Row, eCol.DtNecess)) <> Empty Then
        WLStr_Sql = WLStr_Sql & "and a.crq_DtNece = '" & WLStr_DtNece & "' "
    End If
    
    WLStr_Sql = WLStr_Sql & "order by for_Codigo"
    
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
    
    If WLLng_TotRow > 0 Then
        WLRst_Tabela.MoveFirst
        WLLng_Cont = 0
        
        Call MPrc_LimpaCamposCotacao("T")
        
        Do While Not WLRst_Tabela.EOF
            
            txt_cot_forCodigo(WLLng_Cont).Text = Format(WLRst_Tabela!for_Codigo, "000000")
            txt_cot_forNomeCp(WLLng_Cont).Text = WLRst_Tabela!for_NomeCp
            txt_cot_crqPrUnit(WLLng_Cont).Text = NumeroHelper.Formatar(WLRst_Tabela!crq_PrUnit, 3)
            txt_cot_crqAlqIPI(WLLng_Cont).Text = NumeroHelper.Formatar(WLRst_Tabela!crq_AlqIPI, 2)
            txt_cot_crqAlqICMS(WLLng_Cont).Text = NumeroHelper.Formatar(WLRst_Tabela!crq_AlqICMS, 2)
            txt_cot_crqPrazoE(WLLng_Cont).Text = WLRst_Tabela!crq_PrazoE
            txt_cot_crqCondPg(WLLng_Cont).Text = WLRst_Tabela!crq_CondPg
            
            If WGBol_CPRCadCondPg Then
                For WLLng_Cont2 = 0 To cbo_cot_cdpCodigo(WLLng_Cont).ListCount - 1
                    If cbo_cot_cdpCodigo(WLLng_Cont).ItemData(WLLng_Cont2) = WLRst_Tabela!crq_CondPg Then
                        cbo_cot_cdpCodigo(WLLng_Cont).ListIndex = WLLng_Cont2
                        Exit For
                    End If
                Next
            End If
            
            If WLRst_Tabela!crq_CrdICMS = "S" Then
                chk_creditaIPI(WLLng_Cont).Value = 1
            End If
            
            If WLRst_Tabela!crq_CrdIPI = "S" Then
                chk_creditaICMS(WLLng_Cont).Value = 1
            End If
            
            WLLng_Cont = WLLng_Cont + 1
            WLRst_Tabela.MoveNext
        Loop
    Else
        Exit Function
    End If
    
    WLRst_Tabela.Close
    Set WLRst_Tabela = Nothing
    
    GoTo Fim
Erro:
    GFKEY_MsgBox "Erro ao buscar cotação!" & vbCrLf & vbCrLf & GFKEY_TrataErro(Err.Number, Err.Description)
    Exit Function
Fim:
    MFcn_BuscaCotacao = True
End Function
Private Sub MPrc_BuscaCondPgtos()
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLStr_Sql       As String
    
    Dim WLLng_TotRow    As Long
    Dim WLLng_Cont      As Long
    
    For WLLng_Cont = 0 To txt_cot_crqCondPg.UBound
        txt_cot_crqCondPg(WLLng_Cont).Visible = False
    Next
    
    For WLLng_Cont = 0 To cbo_cot_cdpCodigo.UBound
        cbo_cot_cdpCodigo(WLLng_Cont).Visible = True
        cbo_cot_cdpCodigo(WLLng_Cont).ListIndex = -1
        cbo_cot_cdpCodigo(WLLng_Cont).Clear
    Next
    
    WLStr_Sql = "select * from COND_PAGTOS order by cdp_Descri"
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
    
    If Not WLRst_Tabela Is Nothing Then
        If WLLng_TotRow > 0 Then
            WLRst_Tabela.MoveFirst
            
            Do While Not WLRst_Tabela.EOF
                cbo_cot_cdpCodigo(0).AddItem Trim(WLRst_Tabela!cdp_Descri)
                cbo_cot_cdpCodigo(0).ItemData(cbo_cot_cdpCodigo(0).newIndex) = WLRst_Tabela!cdp_Codigo
                
                cbo_cot_cdpCodigo(1).AddItem Trim(WLRst_Tabela!cdp_Descri)
                cbo_cot_cdpCodigo(1).ItemData(cbo_cot_cdpCodigo(1).newIndex) = WLRst_Tabela!cdp_Codigo
                
                cbo_cot_cdpCodigo(2).AddItem Trim(WLRst_Tabela!cdp_Descri)
                cbo_cot_cdpCodigo(2).ItemData(cbo_cot_cdpCodigo(2).newIndex) = WLRst_Tabela!cdp_Codigo
                
                cbo_cot_cdpCodigo(3).AddItem Trim(WLRst_Tabela!cdp_Descri)
                cbo_cot_cdpCodigo(3).ItemData(cbo_cot_cdpCodigo(3).newIndex) = WLRst_Tabela!cdp_Codigo
                
                cbo_cot_cdpCodigo(4).AddItem Trim(WLRst_Tabela!cdp_Descri)
                cbo_cot_cdpCodigo(4).ItemData(cbo_cot_cdpCodigo(4).newIndex) = WLRst_Tabela!cdp_Codigo
                
                WLRst_Tabela.MoveNext
            Loop
        End If
        WLRst_Tabela.Close
    End If
    Set WLRst_Tabela = Nothing
   
End Sub
Private Sub MPrc_Habilita(PPBol_Opcao As Boolean)
    lbl_Titulo(eTecla.Observacao).Visible = PPBol_Opcao
    lbl_Titulo(eTecla.MedCompras).Visible = PPBol_Opcao
    lbl_Titulo(eTecla.UltCompras).Visible = PPBol_Opcao
    lbl_Titulo(eTecla.Normas).Visible = PPBol_Opcao
    pnl_Cotacoes.Enabled = PPBol_Opcao
End Sub
Private Function MFcn_LerFornecedor(Index As Integer) As Boolean
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLObj_DtBloq    As DataHora
    
    Dim WLLng_TotRow    As Long
    Dim WLLng_Row       As Long
    
    Dim WLInt_Cont      As Integer

    Dim WLStr_Sql       As String
    Dim WLStr_EmpCod    As String
    Dim WLStr_MrcCod    As String
    Dim WLStr_ForCod    As String
    Dim WLStr_GpxCod    As String
    Dim WLStr_PrxCod    As String
    Dim WLStr_PrUnit    As String
    Dim WLStr_Descont   As String
    Dim WLStr_PrDscto   As String
    Dim WLStr_PrFinal   As String
    Dim WLStr_Cols      As String

    Dim WLBol_Existe    As Boolean

    On Local Error GoTo Erro
    
    MFcn_LerFornecedor = False
    
    If Not MFcn_VerificaSelecao(grd_Dados, WLLng_Row) Then Exit Function
    
    WLStr_EmpCod = Left$(Format(Trim(grd_Dados.TextMatrix(WLLng_Row, eCol.EmpMrc)), "0000"), 2)
    WLStr_MrcCod = Right$(Format(Trim(grd_Dados.TextMatrix(WLLng_Row, eCol.EmpMrc)), "0000"), 2)
    WLStr_ForCod = Trim$(txt_cot_forCodigo(Index).Text)
    WLStr_GpxCod = Trim$(grd_Dados.TextMatrix(WLLng_Row, eCol.Grupo))
    WLStr_PrxCod = Trim$(grd_Dados.TextMatrix(WLLng_Row, eCol.Produto))
    
    WLStr_Sql = "select " & _
            "a.for_NomeCP,a.for_DtBloq " & _
        "from " & _
            "FORNECEDORES a " & _
        "where " & _
            "a.for_Codigo = " & WLStr_ForCod
    
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
    
    If WLLng_TotRow > 0 Then
        WLRst_Tabela.MoveFirst
        txt_cot_forNomeCp(Index).Text = Trim(WLRst_Tabela!for_NomeCp)
        
        Set WLObj_DtBloq = New DataHora
        WLObj_DtBloq.Init WLRst_Tabela!for_DtBloq
        
        If Trim(WLRst_Tabela!for_DtBloq) <> Empty Then
            GFKEY_MsgBox "Fornecedor bloqueado desde " & WLObj_DtBloq.Data
            Exit Function
        End If
    Else
        GFKEY_MsgBox "Fornecedor não cadastrado!"
        Exit Function
    End If
    
    WLRst_Tabela.Close
    Set WLRst_Tabela = Nothing

    If WGStr_gerModSis <> "FAN" Then
        WLBol_Existe = False
        
        For WLInt_Cont = txt_cot_forCodigo.LBound To txt_cot_forCodigo.UBound
            If WLInt_Cont <> Index And WLStr_ForCod <> Empty And Trim(txt_cot_forCodigo(WLInt_Cont).Text) = WLStr_ForCod Then
                WLBol_Existe = True
                Exit For
            End If
        Next
        
        If WLBol_Existe = True Then
            GFKEY_MsgBox "Fornecedor já informado!"
            Exit Function
        End If
    End If

    WLLng_TotRow = 0
    
    WLStr_Sql = "select * " & _
        "from " & _
            "PROD_FORNECEDOR " & _
        "where " & _
            "emp_Codigo = '" & WLStr_EmpCod & "' and " & _
            "mrc_Codigo = '" & WLStr_MrcCod & "' and " & _
            "for_Codigo = " & WLStr_ForCod & " and " & _
            "gpx_Codigo = " & WLStr_GpxCod & " and " & _
            "prx_Codigo = '" & WLStr_PrxCod & "'"
    
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
    
    If WLLng_TotRow > 0 Then
        WLRst_Tabela.MoveFirst
        
        If Trim(WLRst_Tabela!prf_DtBloq) <> Empty Then
            GFKEY_MsgBox "Produto Bloqueado para este Fornecedor!"
            Exit Function
        Else
            If Val(WLRst_Tabela!prf_PrUnit) = 0 Then
                txt_cot_crqPrUnit(Index).Text = NumeroHelper.Formatar("0", 3)
            Else
                WLStr_PrUnit = WLRst_Tabela!prf_PrUnit
                WLStr_Descont = WLRst_Tabela!prf_Descon
                WLStr_PrDscto = Val(GFKEY_PreparaValor(WLStr_PrUnit)) * Val(GFKEY_PreparaValor(WLStr_Descont))
                WLStr_PrFinal = Val(GFKEY_PreparaValor(WLStr_PrUnit)) - (Val(GFKEY_PreparaValor(WLStr_PrDscto)) / 100)
                
                txt_cot_crqPrUnit(Index).Text = Format(WLStr_PrFinal, "###,###,##0.0000")
            End If
            
            If WLRst_Tabela!prf_CrdIPI = "S" Then
                chk_creditaIPI(Index).Value = 1
            End If

            If WLRst_Tabela!prf_CrdICMS = "S" Then
                chk_creditaICMS(Index).Value = 1
            End If
        End If
    Else
        If WMBol_Controle Then
            If GFKEY_MsgBox("Fornecedor não cadastrado para este produto!" & vbCrLf & vbCrLf & _
                "Deseja cadastrar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes _
            Then
                WMBol_Controle = False
                
                WLStr_Cols = "emp_Codigo,mrc_Codigo,gpx_Codigo,prx_Codigo,for_Codigo,prf_Descri"
                WLStr_Sql = "insert into PROD_FORNECEDOR (" & WLStr_Cols & _
                    ") values ('" & _
                        Left$(grd_Dados.TextMatrix(WLLng_Row, eCol.EmpMrc), 2) & "','" & _
                        Right$(grd_Dados.TextMatrix(WLLng_Row, eCol.EmpMrc), 2) & "'," & _
                        grd_Dados.TextMatrix(WLLng_Row, eCol.Grupo) & ",'" & _
                        grd_Dados.TextMatrix(WLLng_Row, eCol.Produto) & "'," & _
                        txt_cot_forCodigo(Index).Text & ",'" & _
                        txt_cot_forNomeCp(Index).Text & "')"
                 
                WGCnx_DBPrim.Execute WLStr_Sql
                
                If GFKEY_VerificaErro("A") Then
                    GFKEY_MsgBox "Erro ao cadastrar fornecedor! " & vbCrLf & _
                        vbCrLf & GFKEY_TrataErro(Err.Number, Err.Description)
                    Exit Function
                End If
            Else
                WMBol_Controle = True
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    WLRst_Tabela.Close
    Set WLRst_Tabela = Nothing
    
    GoTo Fim
Erro:
    GFKEY_MsgBox "Erro ao buscar fornecedor!" & vbCrLf & vbCrLf & GFKEY_TrataErro(Err.Number, Err.Description)
    Exit Function
Fim:
    MFcn_LerFornecedor = True
End Function

Private Function MFcn_GravarDados() As Boolean
    Dim WLStr_Sql           As String
    Dim WLStr_Cols          As String
    Dim WLStr_CdpCod        As String
    Dim WLStr_AlqIPI        As String
    Dim WLStr_AlqICM        As String
    
    Dim WLLng_Cont          As Long
    Dim WLBol_TemCotacao    As Boolean
    Dim WLBol_CotadoTotal   As Boolean
    
    On Local Error GoTo Erro
    
    MFcn_GravarDados = False
    '
    WLStr_Sql = "delete from ALMOX_CTREQ_COMPRAS " & _
        "where " & _
            "rqs_Numero = '" & txt_rqsNumero.Text & "' and " & _
            "prx_Codigo = " & grd_Dados.TextMatrix(grd_Dados.Row, eCol.Produto) & " and " & _
            "gpx_Codigo = " & grd_Dados.TextMatrix(grd_Dados.Row, eCol.Grupo)
    
    WGCnx_DBPrim.Execute WLStr_Sql
    If GFKEY_VerificaErro("A") Then GoTo Erro
    '
    WLStr_Sql = "update ALMOX_ITREQ_COMPRAS " & _
        "set irq_Status = 'C' " & _
        "where " & _
            "emp_Codigo = '" & Left$(grd_Dados.TextMatrix(grd_Dados.Row, eCol.EmpMrc), 2) & "' and " & _
            "mrc_Codigo = '" & Right$(grd_Dados.TextMatrix(grd_Dados.Row, eCol.EmpMrc), 2) & "' and " & _
            "rqs_Numero = " & txt_rqsNumero.Text & " and " & _
            "prx_Codigo = " & grd_Dados.TextMatrix(grd_Dados.Row, eCol.Produto) & " and " & _
            "gpx_Codigo = " & grd_Dados.TextMatrix(grd_Dados.Row, eCol.Grupo)
            
    
    WGCnx_DBPrim.Execute WLStr_Sql
    If GFKEY_VerificaErro("A") Then GoTo Erro
    
    WLStr_Cols = _
        "emp_Codigo,mrc_Codigo,rqs_Numero," & _
        "gpx_Codigo,prx_Codigo,crq_DtNece," & _
        "for_Codigo,crq_PrUnit,crq_PrazoE," & _
        "crq_CondPg,cdp_Codigo,crq_AlqIPI," & _
        "crq_AlqICMS,crq_CrdIPI,crq_CrdICMS"
    
    With grd_Dados
        For WLLng_Cont = 0 To txt_cot_forCodigo.UBound
            If txt_cot_forCodigo(WLLng_Cont).Text <> Empty Then
                WLBol_TemCotacao = True
                
                If WGBol_CPRCadCondPg Then
                    WLStr_CdpCod = cbo_cot_cdpCodigo(WLLng_Cont).ItemData(cbo_cot_cdpCodigo(WLLng_Cont).ListIndex)
                Else
                    WLStr_CdpCod = "0"
                End If
                
                WLStr_AlqIPI = GFKEY_PreparaValor(txt_cot_crqAlqIPI(WLLng_Cont).Text)
                WLStr_AlqICM = GFKEY_PreparaValor(txt_cot_crqAlqICMS(WLLng_Cont).Text)
                
                WLStr_Sql = "insert into ALMOX_CTREQ_COMPRAS (" & _
                        WLStr_Cols & _
                    ") values ('" & _
                        Left$(.TextMatrix(.Row, eCol.EmpMrc), 2) & "','" & _
                        Right$(.TextMatrix(.Row, eCol.EmpMrc), 2) & "'," & _
                        txt_rqsNumero.Text & "," & _
                        .TextMatrix(.Row, eCol.Grupo) & ",'" & _
                        .TextMatrix(.Row, eCol.Produto) & "','" & _
                        Format(.TextMatrix(.Row, eCol.DtNecess), "yyyymmdd") & "'," & _
                        txt_cot_forCodigo(WLLng_Cont).Text & ",'" & _
                        GFKEY_PreparaValor(txt_cot_crqPrUnit(WLLng_Cont).Text) & "','" & _
                        txt_cot_crqPrazoE(WLLng_Cont).Text & "','" & _
                        txt_cot_crqCondPg(WLLng_Cont).Text & "'," & _
                        Val(WLStr_CdpCod) & "," & _
                        WLStr_AlqIPI & "," & _
                        WLStr_AlqICM & "," & _
                        IIf(chk_creditaIPI(WLLng_Cont).Value = 0, "'N'", "'S'") & "," & _
                        IIf(chk_creditaICMS(WLLng_Cont).Value = 0, "'N'", "'S'") & ")"
                
                WGCnx_DBPrim.Execute WLStr_Sql
                If GFKEY_VerificaErro("A") Then GoTo Erro
            Else
                WLBol_CotadoTotal = False
            End If
        Next
    End With
    
    If WLBol_TemCotacao = False And WLBol_CotadoTotal = False Then
        Exit Function
    End If
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    MFcn_GravarDados = True

End Function
Private Function MFcn_VerificaSelecao(PFObj_Grid As VSFlexGrid, Optional PFLng_Linha As Long) As Boolean
    Dim WLInt_Cont      As Integer
    Dim WLBol_Selecao   As Boolean
    
    MFcn_VerificaSelecao = False
    WLBol_Selecao = False
    
    With PFObj_Grid
        For WLInt_Cont = 1 To .rows - 1
            If .TextMatrix(WLInt_Cont, 0) = True Then
                WLBol_Selecao = True
                PFLng_Linha = WLInt_Cont
            End If
        Next
    End With
    
    MFcn_VerificaSelecao = WLBol_Selecao
    
End Function
Private Sub MPrc_ChamaPesqNorma()
    Dim WLLng_Linha As Long
    
    If Not MFcn_VerificaSelecao(grd_Dados, WLLng_Linha) Then Exit Sub
    
    With PSQNORMATEC
        WMBol_Pesquisa = True
        .Params.MrcDescri = grd_Dados.TextMatrix(WLLng_Linha, eCol.DescMarca)
        .Params.GpxDescri = grd_Dados.TextMatrix(WLLng_Linha, eCol.DescGrupo)
        .Params.PrxCod = grd_Dados.TextMatrix(WLLng_Linha, eCol.Produto)
        .Params.PrxDescri = grd_Dados.TextMatrix(WLLng_Linha, eCol.DescProd)
        .Params.Norma = grd_Dados.TextMatrix(WLLng_Linha, eCol.Norma)
        
        .Show vbModal
        
        WMBol_Pesquisa = False
    End With
End Sub

Private Sub MPrc_ChamaPesqMedCompr()
    Dim WLLng_Linha As Long
    
    If Not MFcn_VerificaSelecao(grd_Dados, WLLng_Linha) Then Exit Sub
    
    With PSQMEDCOMPR
        WMBol_Pesquisa = True
        .Params.MrcDescri = grd_Dados.TextMatrix(WLLng_Linha, eCol.DescMarca)
        .Params.EmpCod = Left$(grd_Dados.TextMatrix(WLLng_Linha, eCol.EmpMrc), 2)
        .Params.MrcCod = Right$(grd_Dados.TextMatrix(WLLng_Linha, eCol.EmpMrc), 2)
        .Params.GpxCod = grd_Dados.TextMatrix(WLLng_Linha, eCol.Grupo)
        .Params.GpxDescri = grd_Dados.TextMatrix(WLLng_Linha, eCol.DescGrupo)
        .Params.PrxCod = grd_Dados.TextMatrix(WLLng_Linha, eCol.Produto)
        .Params.PrxDescri = grd_Dados.TextMatrix(WLLng_Linha, eCol.DescProd)
        
        .Show vbModal
        
        WMBol_Pesquisa = False
    End With
End Sub
Private Sub MPrc_ChamaPesqFornec(Index As Integer)
    Dim WLLng_Linha As Long
    Dim WLStr_Sql   As String
    Dim WLStr_Cols  As String
    
    On Local Error GoTo Erro
     
    If Not WGObj_FormPsqFor Is Nothing Then
        Beep
        GFKEY_MsgBox "Pesquisa de fornecedores já ativa em outro módulo!"
        DoEvents
    Else
        WMBol_Pesquisa = True
        Set WGObj_FormPsqFor = Me
        Me.Enabled = False
        PSQFORNECEDORES.Show vbModal
        
        txt_cot_forCodigo(Index).Text = txt_forCodigo.Text
        txt_cot_forNomeCp(Index).Text = txt_forNomeCp.Text
        Set WGObj_FormPsqFor = Nothing
        
        If txt_forCodigo.Text <> Empty Then
            If Not MFcn_LerFornecedor(Index) Then
                Call MPrc_LimpaCamposCotacao("P", Index)
                txt_cot_forCodigo(Index).SetFocus
                Exit Sub
            End If
        Else
            WMBol_Pesquisa = False
            Exit Sub
        End If
    End If
    
    WMBol_Pesquisa = False
    
    GoTo Fim
Erro:
    GFKEY_MsgBox "Erro ao cadastrar fornecedor! " & vbCrLf & vbCrLf & GFKEY_TrataErro(Err.Number, Err.Description)
    Exit Sub
Fim:
    GFKEY_MsgBox "Fornecedor cadastrado com sucesso!"
End Sub
Private Function MFcn_ValidaGravacao() As Boolean
    Dim WLLng_Cont  As Long
    
    MFcn_ValidaGravacao = False
    
    For WLLng_Cont = 0 To txt_cot_forCodigo.UBound
        If txt_cot_forCodigo(WLLng_Cont).Text <> Empty Then
            
            If txt_cot_crqPrUnit(WLLng_Cont).Text = Empty Then
                GFKEY_MsgBox "Informe o preço unitário!"
                txt_cot_crqPrUnit(WLLng_Cont).SetFocus
                Exit Function
            End If
            
            If txt_cot_crqAlqIPI(WLLng_Cont).Text = Empty Then
                GFKEY_MsgBox "Informe a alíquota de IPI!"
                txt_cot_crqAlqIPI(WLLng_Cont).SetFocus
                Exit Function
            End If
            
            If txt_cot_crqAlqICMS(WLLng_Cont).Text = Empty Then
                GFKEY_MsgBox "Informe a alíquota de ICMS!"
                txt_cot_crqAlqICMS(WLLng_Cont).SetFocus
                Exit Function
            End If
             
            If WGBol_CPRCadCondPg Then
                If cbo_cot_cdpCodigo(WLLng_Cont).ListIndex = -1 Then
                    GFKEY_MsgBox "Informe a Condição de pagamento!"
                    cbo_cot_cdpCodigo(WLLng_Cont).SetFocus
                    Exit Function
                End If
            Else
                If txt_cot_crqCondPg(WLLng_Cont).Text = Empty Then
                    GFKEY_MsgBox "Informe a Condição de pagamento!"
                    txt_cot_crqCondPg(WLLng_Cont).SetFocus
                    Exit Function
                End If
            End If
            
            If txt_cot_crqPrazoE(WLLng_Cont).Text = Empty Then
                GFKEY_MsgBox "Informe o prazo de entrega!"
                txt_cot_crqPrazoE(WLLng_Cont).SetFocus
                Exit Function
            End If
        End If
    Next
    
    MFcn_ValidaGravacao = True
    
End Function
