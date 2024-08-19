VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form KEYCPR01034Fani 
   BackColor       =   &H00FDDEC6&
   BorderStyle     =   0  'None
   Caption         =   "Fechamento da Requisição"
   ClientHeight    =   10335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14520
   Icon            =   "KEYCPR01034Fani.frx":0000
   LinkTopic       =   "KEYCPR01034Fani"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_Despesas 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Definir Grupo e Tipo de Despesa como Padrão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9300
      TabIndex        =   104
      Top             =   9240
      Width           =   4155
   End
   Begin VB.Frame fra_Selecao 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Seleção"
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
      Height          =   1455
      Left            =   60
      TabIndex        =   85
      Top             =   420
      Width           =   14355
      Begin VB.ComboBox cbo_empCodigo 
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
         Left            =   10260
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txt_rqsNumero 
         Height          =   315
         Left            =   1380
         TabIndex        =   95
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txt_rqsAprova 
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
         TabIndex        =   93
         Top             =   960
         Width           =   4635
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
         Left            =   10260
         MaxLength       =   10
         TabIndex        =   88
         Top             =   240
         Width           =   1455
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
         TabIndex        =   87
         Top             =   600
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
         Left            =   10260
         MaxLength       =   40
         TabIndex        =   86
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lbl_Titulo07 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Empresa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9000
         TabIndex        =   97
         Tag             =   "*"
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label lbl_Titulo14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Aprovador"
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
         Left            =   120
         TabIndex        =   94
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label lbl_Titulo01 
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
         Left            =   120
         TabIndex        =   92
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lbl_Titulo02 
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
         Left            =   9000
         TabIndex        =   91
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lbl_Titulo03 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Solicitante"
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
         Left            =   120
         TabIndex        =   90
         Top             =   660
         Width           =   900
      End
      Begin VB.Label lbl_Titulo04 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Gerente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   9000
         TabIndex        =   89
         Top             =   660
         Width           =   585
      End
   End
   Begin VB.CommandButton cmd_sair 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Sair"
      Height          =   675
      Left            =   13575
      Picture         =   "KEYCPR01034Fani.frx":1EAE2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9240
      Width           =   855
   End
   Begin VB.CommandButton cmd_Limpar 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Limpar"
      Height          =   675
      Left            =   900
      Picture         =   "KEYCPR01034Fani.frx":1EC2C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9240
      Width           =   855
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   60
      Picture         =   "KEYCPR01034Fani.frx":1ED76
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9240
      Width           =   855
   End
   Begin VB.PictureBox pic_observacao 
      BackColor       =   &H00C0E0FF&
      Height          =   3135
      Left            =   3420
      ScaleHeight     =   3075
      ScaleWidth      =   7215
      TabIndex        =   71
      Top             =   10800
      Visible         =   0   'False
      Width           =   7275
      Begin VB.Frame fra_Observ 
         BackColor       =   &H00C0E0FF&
         Height          =   2595
         Left            =   60
         TabIndex        =   74
         Top             =   240
         Width           =   7095
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   9
            Left            =   60
            TabIndex        =   84
            Top             =   2340
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   8
            Left            =   60
            TabIndex        =   83
            Top             =   2100
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   7
            Left            =   60
            TabIndex        =   82
            Top             =   1860
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   6
            Left            =   60
            TabIndex        =   81
            Top             =   1620
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   5
            Left            =   60
            TabIndex        =   80
            Top             =   1380
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   4
            Left            =   60
            TabIndex        =   79
            Top             =   1140
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   3
            Left            =   60
            TabIndex        =   78
            Top             =   900
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   2
            Left            =   60
            TabIndex        =   77
            Top             =   660
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   1
            Left            =   60
            TabIndex        =   76
            Top             =   420
            Width           =   6960
         End
         Begin VB.Label lbl_obs 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            Index           =   0
            Left            =   60
            TabIndex        =   75
            Top             =   180
            Width           =   6960
         End
      End
      Begin VB.Label lbl_TituloEscObs 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   3120
         TabIndex        =   73
         Top             =   2880
         Width           =   945
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   60
         TabIndex        =   72
         Top             =   0
         Width           =   6945
      End
   End
   Begin VB.PictureBox pic_ucp 
      BackColor       =   &H00FFC0C0&
      Height          =   3975
      Left            =   1860
      ScaleHeight     =   3915
      ScaleWidth      =   10155
      TabIndex        =   15
      Top             =   10500
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Label lbl_funcao06 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "<ESC> Sair"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4560
         TabIndex        =   69
         Top             =   3540
         Width           =   1050
      End
      Begin VB.Label lbl_Titulo39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   68
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lbl_ucp_empDescri 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   67
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbl_ucp_pcpDtPedi 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   66
         Top             =   3060
         Width           =   915
      End
      Begin VB.Label lbl_ucp_forNomeCp 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1200
         TabIndex        =   65
         Top             =   3060
         Width           =   3735
      End
      Begin VB.Label lbl_ucp_ipcPrUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   4980
         TabIndex        =   64
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_ipcQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6180
         TabIndex        =   63
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_pcpCondPg 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   7440
         TabIndex        =   62
         Top             =   3060
         Width           =   2535
      End
      Begin VB.Label lbl_ucp_pcpDtPedi 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   61
         Top             =   2820
         Width           =   915
      End
      Begin VB.Label lbl_ucp_forNomeCp 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1200
         TabIndex        =   60
         Top             =   2820
         Width           =   3735
      End
      Begin VB.Label lbl_ucp_ipcPrUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   4980
         TabIndex        =   59
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_ipcQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   6180
         TabIndex        =   58
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_pcpCondPg 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   7440
         TabIndex        =   57
         Top             =   2820
         Width           =   2535
      End
      Begin VB.Label lbl_ucp_pcpDtPedi 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   56
         Top             =   2580
         Width           =   915
      End
      Begin VB.Label lbl_ucp_forNomeCp 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   55
         Top             =   2580
         Width           =   3735
      End
      Begin VB.Label lbl_ucp_ipcPrUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   4980
         TabIndex        =   54
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_ipcQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6180
         TabIndex        =   53
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_pcpCondPg 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   7440
         TabIndex        =   52
         Top             =   2580
         Width           =   2535
      End
      Begin VB.Label lbl_ucp_pcpDtPedi 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   51
         Top             =   2340
         Width           =   915
      End
      Begin VB.Label lbl_ucp_forNomeCp 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   50
         Top             =   2340
         Width           =   3735
      End
      Begin VB.Label lbl_ucp_ipcPrUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4980
         TabIndex        =   49
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_ipcQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6180
         TabIndex        =   48
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_pcpCondPg 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   7440
         TabIndex        =   47
         Top             =   2340
         Width           =   2535
      End
      Begin VB.Label lbl_ucp_pcpDtPedi 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   46
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label lbl_ucp_forNomeCp 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   45
         Top             =   2100
         Width           =   3735
      End
      Begin VB.Label lbl_ucp_ipcPrUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4980
         TabIndex        =   44
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_ipcQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6180
         TabIndex        =   43
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_pcpCondPg 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7440
         TabIndex        =   42
         Top             =   2100
         Width           =   2535
      End
      Begin VB.Label lbl_ucp_pcpDtPedi 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   41
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label lbl_ucp_forNomeCp 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   40
         Top             =   1860
         Width           =   3735
      End
      Begin VB.Label lbl_ucp_ipcPrUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4980
         TabIndex        =   39
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_ipcQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6180
         TabIndex        =   38
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_pcpCondPg 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   7440
         TabIndex        =   37
         Top             =   1860
         Width           =   2535
      End
      Begin VB.Label lbl_ucp_pcpCondPg 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7440
         TabIndex        =   36
         Top             =   1620
         Width           =   2535
      End
      Begin VB.Label lbl_ucp_ipcQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6180
         TabIndex        =   35
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_ipcPrUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4980
         TabIndex        =   34
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label lbl_ucp_forNomeCp 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   33
         Top             =   1620
         Width           =   3735
      End
      Begin VB.Label lbl_ucp_pcpDtPedi 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   32
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8340
         TabIndex        =   31
         Top             =   900
         Width           =   675
      End
      Begin VB.Label lbl_ucp_prxSalEst 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9060
         TabIndex        =   30
         Top             =   900
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Últimas Compras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   60
         Width           =   9900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Ref."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label lbl_ucp_prxCodigo 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8400
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lbl_ucp_prxDescri 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   26
         Top             =   900
         Width           =   4200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Produto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6900
         TabIndex        =   25
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   900
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5760
         TabIndex        =   23
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lbl_ucp_irqQuanti 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6780
         TabIndex        =   22
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   1380
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5220
         TabIndex        =   20
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6360
         TabIndex        =   19
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cond.Pagto."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7440
         TabIndex        =   18
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Label lbl_ucp_gpxDescri 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4500
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   600
         Width           =   510
      End
      Begin VB.Line lin_linha02 
         X1              =   180
         X2              =   10020
         Y1              =   1260
         Y2              =   1260
      End
   End
   Begin VB.Frame pnl_Requisicoes 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Requisições"
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
      Height          =   3975
      Left            =   60
      TabIndex        =   14
      Top             =   1920
      Width           =   14355
      Begin VB.ComboBox cbo_TpDesp 
         Height          =   315
         Left            =   5220
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   180
         Width           =   2835
      End
      Begin VB.ComboBox cbo_GrDesp 
         Height          =   315
         Left            =   1020
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   180
         Width           =   2835
      End
      Begin VB.ComboBox cbo_EmpComp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10260
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   180
         Width           =   3735
      End
      Begin VSFlex8Ctl.VSFlexGrid grd_Requisicoes 
         Height          =   3255
         Left            =   60
         TabIndex        =   0
         Top             =   600
         Width           =   14235
         _cx             =   268788245
         _cy             =   268768877
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
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Tip Desp."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   101
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Gpo Desp."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lbl_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Empresa Compradora:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   8280
         TabIndex        =   99
         Top             =   240
         Width           =   1875
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
      Height          =   3195
      Left            =   60
      TabIndex        =   13
      Top             =   5940
      Width           =   14355
      Begin VSFlex8Ctl.VSFlexGrid grd_Cotacoes 
         Height          =   2835
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   14235
         _cx             =   933126677
         _cy             =   933106569
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
   Begin VB.PictureBox pic_titulo 
      Appearance      =   0  'Flat
      BackColor       =   &H009C832C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5205
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   5200
      Begin VB.Image img_fechar 
         Height          =   270
         Left            =   4740
         Picture         =   "KEYCPR01034Fani.frx":1F300
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image img_minimizar 
         Height          =   270
         Left            =   4440
         Picture         =   "KEYCPR01034Fani.frx":1F782
         ToolTipText     =   "Minimizar"
         Top             =   60
         Width           =   270
      End
      Begin VB.Label lbl_Tituloform 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fechamento da Requisição"
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
         TabIndex        =   11
         Top             =   60
         Width           =   2430
      End
      Begin VB.Label lbl_fundo 
         Appearance      =   0  'Flat
         BackColor       =   &H009C832C&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   12
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9960
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
         TabIndex        =   8
         Top             =   60
         Width           =   1680
      End
      Begin VB.Label lbl_fundo 
         BackColor       =   &H009C832C&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.Label lbl_FuncObsReq 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDEC6&
      Caption         =   "<F6> Observação"
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
      Left            =   8535
      TabIndex        =   6
      Top             =   9660
      Visible         =   0   'False
      Width           =   1620
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
      TabIndex        =   70
      Top             =   0
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label lbl_funcompras 
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
      Left            =   6060
      TabIndex        =   5
      Top             =   9660
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Shape shp_borda 
      BorderColor     =   &H80000012&
      Height          =   315
      Left            =   5340
      Top             =   10080
      Width           =   255
   End
End
Attribute VB_Name = "KEYCPR01034Fani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WMRst_Tabela        As ADODB.Recordset
Dim WMTyp_Perfil        As typ_Acesso

Dim WMStr_Log           As String
Dim WMStr_Auxi          As String
Dim WMStr_Chave         As String
Dim WMStr_EmpCod        As String
Dim WMStr_MrcCod        As String
Dim WMStr_EmpCodComp    As String
Dim WMStr_MrcCodComp    As String
Dim WMStr_GpxCod        As String
Dim WMStr_ReqNumero     As String

Dim WMBol_Inicio        As Boolean
Dim WMBol_ImprPedi      As Boolean

'Grid Requisicoes
Private Enum eColR
    Selecao = 0
    Grupo = 1
    GrpDesc = 2
    Codigo = 3
    CodDesc = 4
    DtNeces = 5
    Quantidade = 6
    CodFornec = 7
    NomeFornec = 8
    PrUnit = 9
    AlqIPI = 10
    DtPrev = 11
    CondPgto = 12
    Norma = 13
    PedCompra = 14
    Observacao = 15
    Exclui = 16
    EmpMrc = 17
    EmpMrcComp = 18
    Estoque = 19
    Cdp = 20
    ObservItem = 21
    Status = 22
    LastCol = 22
End Enum

'Grid Cotacoes
Private Enum eColC
    Selecao = 0
    CodFornec = 1
    DescFor = 2
    PrFornec = 3
    ValIPI = 4
    CredIPI = 5
    PrCIPI = 6
    ValTot = 7
    TotCIPI = 8
    CodCdp = 9
    CondPgto = 10
    DiasEntr = 11
    VlICMS = 12
    CredICMS = 13
    VlCustoReal = 14
    LastCol = 14
End Enum
Private Sub cbo_empCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
        SendKeysA vbKeyTab, False
    End If
End Sub

Private Sub cbo_EmpComp_GotFocus()
    cbo_EmpComp.BackColor = vbYellow
End Sub

Private Sub cbo_EmpComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
        SendKeysA vbKeyTab, False
    End If
End Sub

Private Sub cbo_EmpComp_LostFocus()
    cbo_EmpComp.BackColor = vbWhite
End Sub

Private Sub cbo_GrDesp_Click()
    If cbo_GrDesp.ListIndex >= 0 Then
        Call MFcn_CarregaTipoDespesas
    End If
End Sub

Private Sub cbo_GrDesp_GotFocus()
    cbo_GrDesp.BackColor = vbYellow
End Sub

Private Sub cbo_GrDesp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
        SendKeysA vbKeyTab, False
    End If
End Sub

Private Sub cbo_GrDesp_LostFocus()
    cbo_GrDesp.BackColor = vbWhite
End Sub

Private Sub cbo_TpDesp_GotFocus()
    cbo_TpDesp.BackColor = vbYellow
End Sub

Private Sub cbo_TpDesp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
        SendKeysA vbKeyTab, False
    End If
End Sub

Private Sub cbo_TpDesp_LostFocus()
    cbo_TpDesp.BackColor = vbWhite
End Sub

Private Sub cmd_confirmar_Click()
    Dim WLStr_Sql       As String
    Dim WLStr_Msg       As String
    Dim WLLng_QtdPed    As String
    Dim WLBol_Selecao   As Boolean
    Dim WLLng_Linha     As Long
    Dim WLStr_Erro      As String
    '
    On Local Error GoTo Erro
    
    If Not MFcn_ValidaRequisicao Then Exit Sub
    
    WGCnx_DBPrim.BeginTrans
    If Not MFcn_TrataDados Then GoTo Erro
   
    GoTo Fim
Erro:
    If Err.Number <> 0 Then
        WLStr_Erro = vbLf & vbLf & _
            "Erro: " & Err.Number & vbLf & _
            "Origem: " & Err.Source & vbLf & _
            "Descrição: " & Err.Description
    Else
        WLStr_Erro = ""
    End If
    WGCnx_DBPrim.RollbackTrans
    GFKEY_MsgBox "Erro na atualização da requisição de compras!" & vbCrLf & vbCrLf & "Entre em contato com o suporte!" & WLStr_Erro
    Exit Sub
Fim:
    GFKEY_MsgBox "Gravação Ok!"
    WGCnx_DBPrim.CommitTrans
    Call MPrc_LimpaCampos
    txt_rqsNumero.SetFocus
    
End Sub
Private Sub cmd_confirmar_GotFocus()
    cmd_confirmar.BackColor = keyCAmarelo
End Sub
Private Sub cmd_confirmar_LostFocus()
    cmd_confirmar.BackColor = &HFDDEC6
End Sub

Private Sub cmd_Limpar_Click()
    Call MPrc_LimpaCampos
End Sub

Private Sub cmd_Limpar_GotFocus()
    cmd_Limpar.BackColor = keyCAmarelo
End Sub

Private Sub cmd_Limpar_LostFocus()
    cmd_Limpar.BackColor = &HFDDEC6
End Sub

Private Sub cmd_sair_Click()
    Unload Me
End Sub

Private Sub cmd_sair_GotFocus()
    cmd_sair.BackColor = keyCAmarelo
End Sub

Private Sub cmd_sair_LostFocus()
    cmd_sair.BackColor = keyCBackColor
End Sub
Private Sub Form_Activate()
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLInt_TotRow    As Integer
    Dim WLInt_Cont      As Integer
    Dim WLInt_ContB      As Integer
    Dim WLStr_Sql      As String
    
    On Local Error GoTo Erro
    '
    If WMBol_Inicio = True Then
        WMBol_Inicio = False
        WMTyp_Perfil = GFKEY_VerAcessos(WGStr_Usuario, "P", Me.Name)
        Call GPKEY_CantoRedondo(Me, 25)
        
        If Not GFKEY_CarregaEmpresas(cbo_empCodigo, , True) Then GoTo Erro
        If Not GFKEY_CarregaEmpresas(cbo_EmpComp, , True) Then GoTo Erro
        
        cbo_empCodigo.Enabled = False
        
        If Not MFcn_CarregaGruposDespesas Then GoTo Erro
    End If

    txt_rqsAprova.Text = WGStr_CPR_Aprova
    txt_rqsNumero.SetFocus
    
    GoTo Fim
Erro:
    If Err.Number <> 0 Then
        GFKEY_MsgBox "Erro ao iniciar programa" & vbCrLf & vbCrLf & "Entre em contato com suporte"
    End If
    Unload Me
Fim:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If lbl_funcompras.Visible = True Then
        Call MPrc_VerTeclaAtalho(KeyCode)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If lbl_funcompras.Visible = True Then
        Call MPrc_VerTeclaAtalho(KeyAscii)
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If lbl_funcompras.Visible = True Then
        Call MPrc_VerTeclaAtalho(KeyCode)
    End If
End Sub

Private Sub Form_Load()
    WMBol_Inicio = True
End Sub

Private Sub grd_Cotacoes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim WLStr_Msg   As String
    With grd_Cotacoes
        If .TextMatrix(Row, eColC.Selecao) = True Then
            .Cell(flexcpBackColor, Row, eColC.CodFornec, Row, eColC.LastCol) = vbYellow
            
            WLStr_Msg = "Confirma a escolha da cotação?"
            
            If GFKEY_MsgBox(WLStr_Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call MPrc_GravaCotSel(Row)
            Else
                .TextMatrix(Row, eColR.Selecao) = False
                .Cell(flexcpBackColor, Row, eColC.CodFornec, Row, eColC.LastCol) = vbWhite
            End If
        Else
            .Cell(flexcpBackColor, Row, eColC.CodFornec, Row, eColC.LastCol) = vbWhite
        End If
    End With
End Sub

Private Sub grd_Cotacoes_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim WLLng_Linha As Long
    Dim WLLng_Cont  As Long
    
    With grd_Cotacoes
        If .rows > 1 Then
            If OldRow <> NewRow And OldRow > 0 Then
                If .TextMatrix(OldRow, eColR.Selecao) = True Then
                    .TextMatrix(OldRow, eColR.Selecao) = False
                End If
                .Cell(flexcpBackColor, OldRow, eColC.CodFornec, OldRow, eColC.LastCol) = vbWhite
                .Cell(flexcpBackColor, NewRow, eColC.CodFornec, NewRow, eColC.LastCol) = vbYellow
            End If
        End If
    End With
End Sub

Private Sub grd_Cotacoes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> eColC.Selecao Then
        Cancel = True
    End If
End Sub
Private Sub grd_Requisicoes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With grd_Requisicoes
        Select Case Col
            Case eColR.Selecao
                If .TextMatrix(Row, eColR.Selecao) = True Then
                    Call MPrc_HabilitaTeclasFunc(True)
                    Call MPrc_CarregaCotacoes(Row)
                    .Cell(flexcpBackColor, Row, eColR.Grupo, Row, eColR.LastCol) = vbYellow
                    
                    If UCase(.TextMatrix(Row, eColR.Status)) = "SIM" Then
                        Call MPrc_BuscaCotSel(Row)
                    End If
                    
                Else
                    Call MPrc_HabilitaTeclasFunc(False)
                    Call MPrc_MontaGridCot
                    .Cell(flexcpBackColor, Row, eColR.Grupo, Row, eColR.LastCol) = vbWhite
                End If
                
            Case eColR.Quantidade
                .TextMatrix(Row, eColR.Selecao) = True
                .Cell(flexcpBackColor, Row, eColR.Grupo, Row, eColR.LastCol) = vbYellow
                Call MPrc_CarregaCotacoes(Row)
        End Select
    End With
End Sub
Private Sub grd_Requisicoes_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With grd_Requisicoes
        If .rows > 1 Then
            If OldRow <> NewRow And OldRow > 0 Then
                If NewCol = eColR.Selecao Or NewCol = eColR.Quantidade Then
                    If .TextMatrix(OldRow, eColR.Selecao) = True Then
                        .TextMatrix(OldRow, eColR.Selecao) = False
                    End If
                    .Cell(flexcpBackColor, OldRow, eColR.Grupo, OldRow, eColR.LastCol) = keyCBranco
                    .Cell(flexcpBackColor, NewRow, eColR.Grupo, NewRow, eColR.LastCol) = keyCAmarelo
                End If
            End If
        End If
    End With
End Sub

Private Sub grd_Requisicoes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Select Case Col
        Case eColR.Selecao
            Cancel = False
        Case eColR.Quantidade
            Cancel = False
        Case Else
            Cancel = True
    End Select
End Sub

Private Sub grd_Requisicoes_KeyUp(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub lbl_funcao06_Click()
    Call MPrc_VerTeclaAtalho(vbKeyEscape)
End Sub

Private Sub lbl_FuncObsReq_Click()
    If lbl_FuncObsReq.Visible = True Then
        Call MPrc_VerTeclaAtalho(vbKeyF6)
    End If
End Sub

Private Sub lbl_funcompras_Click()
    If lbl_funcompras.Visible = True Then
        Call MPrc_VerTeclaAtalho(vbKeyF5)
    End If
End Sub

Private Sub lbl_Titulo01_Click()
    Call MPrc_PesquisaRequisicao
End Sub

Private Sub lbl_TituloEscObs_Click(Index As Integer)
    Call MPrc_VerTeclaAtalho(vbKeyEscape)
End Sub

Private Sub pic_observacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
End Sub

Private Sub pic_observacao_KeyPress(KeyAscii As Integer)
    Call MPrc_VerTeclaAtalho(KeyAscii)
End Sub

Private Sub pic_observacao_KeyUp(KeyCode As Integer, Shift As Integer)
    Call MPrc_VerTeclaAtalho(KeyCode)
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
Private Sub MPrc_MontaGridCot()
     Dim WLStr_TituloCols    As String
    
    On Local Error GoTo Erro
    
    WLStr_TituloCols = " |Codigo|Nome Fornecedor|Preço|Alq.IPI|Credita|Pr.c/IPI|Val.Tot|Tot c/IPI|" & _
        "Cod.CondPgto|Cond.Pgto|Data.Entrega|ICMS|Credita|Vl.Cst Real"
    
    With grd_Cotacoes
        .Clear
        .Refresh
        .rows = 1
        .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AllowUserResizing = flexResizeColumns
        .FormatString = WLStr_TituloCols
        .ScrollTrack = True
        .AllowBigSelection = False
        .FocusRect = flexFocusLight
        .Cols = eColC.LastCol + 1
        
        .ColWidth(eColC.Selecao) = 255
        .ColDataType(eColC.Selecao) = flexDTBoolean
        .FixedAlignment(eColC.Selecao) = flexAlignCenterCenter
        .ColAlignment(eColC.Selecao) = flexAlignCenterCenter
       
        .ColWidth(eColC.CodFornec) = 750
        .ColDataType(eColC.CodFornec) = flexDTString
        .Cell(flexcpAlignment, 0, eColC.CodFornec, .rows - 1, eColC.CodFornec) = flexAlignLeftTop
        .ColHidden(eColC.CodFornec) = True
        
        .ColWidth(eColC.DescFor) = 3810
        .ColDataType(eColC.DescFor) = flexDTString
        .Cell(flexcpAlignment, 0, eColC.DescFor, .rows - 1, eColC.DescFor) = flexAlignLeftTop
        
        .ColWidth(eColC.PrFornec) = 840
        .ColDataType(eColC.PrFornec) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColC.PrFornec, .rows - 1, eColC.PrFornec) = flexAlignLeftTop
        .ColFormat(eColC.PrFornec) = "#,###,##0.00"
        
        .ColWidth(eColC.ValIPI) = 735
        .ColDataType(eColC.ValIPI) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColC.ValIPI, .rows - 1, eColC.ValIPI) = flexAlignLeftTop
        .ColFormat(eColC.ValIPI) = "#,###,##0.00"
        
        .ColWidth(eColC.CredIPI) = 705
        .ColDataType(eColC.CredIPI) = flexDTBoolean
        .FixedAlignment(eColC.CredIPI) = flexAlignCenterCenter
        .ColAlignment(eColC.CredIPI) = flexAlignCenterCenter
        
        .ColWidth(eColC.PrCIPI) = 900
        .ColDataType(eColC.PrCIPI) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColC.PrCIPI, .rows - 1, eColC.PrCIPI) = flexAlignLeftTop
        .ColFormat(eColC.PrCIPI) = "#,###,##0.00"
        
        .ColWidth(eColC.ValTot) = 1065
        .ColDataType(eColC.ValTot) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColC.ValTot, .rows - 1, eColC.ValTot) = flexAlignLeftTop
        .ColFormat(eColC.ValTot) = "#,###,##0.00"
        
        .ColWidth(eColC.TotCIPI) = 1065
        .ColDataType(eColC.TotCIPI) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColC.TotCIPI, .rows - 1, eColC.TotCIPI) = flexAlignLeftTop
        .ColFormat(eColC.TotCIPI) = "#,###,##0.00"
        
        .ColWidth(eColC.CodCdp) = 1275
        .ColDataType(eColC.CodCdp) = flexDTString
        .Cell(flexcpAlignment, 0, eColC.CodCdp, .rows - 1, eColC.CodCdp) = flexAlignLeftTop
        .ColHidden(eColC.CodCdp) = True
        
        .ColWidth(eColC.CondPgto) = 990
        .ColDataType(eColC.CondPgto) = flexDTString
        .Cell(flexcpAlignment, 0, eColC.CondPgto, .rows - 1, eColC.CondPgto) = flexAlignLeftTop
        
        .ColWidth(eColC.DiasEntr) = 1365
        .ColDataType(eColC.DiasEntr) = flexDTString
        .Cell(flexcpAlignment, 0, eColC.DiasEntr, .rows - 1, eColC.DiasEntr) = flexAlignCenterTop
        
        .ColWidth(eColC.VlICMS) = 645
        .ColDataType(eColC.VlICMS) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColC.VlICMS, .rows - 1, eColC.VlICMS) = flexAlignLeftTop
        .ColFormat(eColC.VlICMS) = "#,###,##0.00"
        
        .ColWidth(eColC.CredICMS) = 705
        .ColDataType(eColC.CredICMS) = flexDTBoolean
        .FixedAlignment(eColC.CredICMS) = flexAlignCenterCenter
        .ColAlignment(eColC.CredICMS) = flexAlignCenterCenter
        
        .ColWidth(eColC.VlCustoReal) = 1035
        .ColDataType(eColC.VlCustoReal) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColC.VlCustoReal, .rows - 1, eColC.VlCustoReal) = flexAlignLeftTop
        .ColFormat(eColC.VlCustoReal) = "#,###,##0.00"
        
        .Cell(flexcpFontBold, .rows - 1, eColC.CodFornec, .rows - 1, eColC.LastCol) = True
        
    End With
    '
    GoTo Fim
Erro:
    If Err.Number <> 0 Then
        GFKEY_MsgBox "Erro ao montar grid " & vbCrLf & vbCrLf & "Entre em contato com o suporte"
        grd_Requisicoes.Clear
        Exit Sub
    End If
Fim:

End Sub
Private Sub MPrc_MontaGridReq()
    '
    Dim WLStr_TituloCols    As String
    
    On Local Error GoTo Erro
    
    WLStr_TituloCols = " |Grupo|Descrição Grupo|Produto|Descrição Produto|" & _
        "Dt.Neces|Quantidade|Cód Forn|Nome.Fornec|" & _
        "Pr.Unit|IPI|Dt.Prev|Cond.Pgto|Norma|Ped.Compra|Observação|Exclui?|" & _
        "EmpMrc|EmpMrcComp|Estoque|Cdp|Observacao item|Status"
        
    With grd_Requisicoes
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
        .Cols = eColR.LastCol + 1
        
        .ColWidth(eColC.Selecao) = 255
        .ColDataType(eColC.Selecao) = flexDTBoolean
        .FixedAlignment(eColC.Selecao) = flexAlignCenterCenter
        .ColAlignment(eColC.Selecao) = flexAlignCenterCenter
        
        .ColWidth(eColR.Grupo) = 1020
        .ColDataType(eColR.Grupo) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Grupo, .rows - 1, eColR.Grupo) = flexAlignCenterTop
        
        .ColWidth(eColR.GrpDesc) = 2985
        .ColDataType(eColR.GrpDesc) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.GrpDesc, .rows - 1, eColR.GrpDesc) = flexAlignCenterTop
        
        .ColWidth(eColR.Codigo) = 1320
        .ColDataType(eColR.Codigo) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Codigo, .rows - 1, eColR.Codigo) = flexAlignCenterTop
        
        .ColWidth(eColR.CodDesc) = 4500
        .ColDataType(eColR.CodDesc) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.CodDesc, .rows - 1, eColR.CodDesc) = flexAlignCenterTop
        
        .ColWidth(eColR.DtNeces) = 1395
        .ColDataType(eColR.DtNeces) = flexDTDate
        .Cell(flexcpAlignment, 0, eColR.DtNeces, .rows - 1, eColR.DtNeces) = flexAlignCenterTop
        .ColHidden(eColR.DtNeces) = IIf(WGBol_CPRDtNecessid = False, True, False)
        
        .ColWidth(eColR.Quantidade) = 1605
        .ColDataType(eColR.Quantidade) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Quantidade, .rows - 1, eColR.Quantidade) = flexAlignCenterTop
        
        .ColWidth(eColR.CodFornec) = 1200
        .ColDataType(eColR.CodFornec) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.CodFornec, .rows - 1, eColR.CodFornec) = flexAlignCenterTop
        .ColHidden(eColR.CodFornec) = True
        
        .ColWidth(eColR.NomeFornec) = 4000
        .ColDataType(eColR.NomeFornec) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.NomeFornec, .rows - 1, eColR.NomeFornec) = flexAlignCenterTop
        .ColHidden(eColR.NomeFornec) = True
        
        .ColWidth(eColR.PrUnit) = 1200
        .ColDataType(eColR.PrUnit) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColR.PrUnit, .rows - 1, eColR.PrUnit) = flexAlignCenterTop
        .ColHidden(eColR.PrUnit) = True
        .ColFormat(eColR.PrUnit) = "##,###,###.00"
        
        .ColWidth(eColR.AlqIPI) = 1200
        .ColDataType(eColR.AlqIPI) = flexDTDouble
        .Cell(flexcpAlignment, 0, eColR.AlqIPI, .rows - 1, eColR.AlqIPI) = flexAlignCenterTop
        .ColHidden(eColR.AlqIPI) = True
        .ColFormat(eColR.AlqIPI) = "##,###,###.00"
        
        .ColWidth(eColR.DtPrev) = 1200
        .ColDataType(eColR.DtPrev) = flexDTDate
        .Cell(flexcpAlignment, 0, eColR.DtPrev, .rows - 1, eColR.DtPrev) = flexAlignCenterTop
        .ColHidden(eColR.DtPrev) = True
        
        .ColWidth(eColR.CondPgto) = 1200
        .ColDataType(eColR.CondPgto) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.CondPgto, .rows - 1, eColR.CondPgto) = flexAlignCenterTop
        .ColHidden(eColR.CondPgto) = True
        
        .ColWidth(eColR.Norma) = 1200
        .ColDataType(eColR.Norma) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Norma, .rows - 1, eColR.Norma) = flexAlignCenterTop
        .ColHidden(eColR.Norma) = True
        
        .ColWidth(eColR.PedCompra) = 1200
        .ColDataType(eColR.PedCompra) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.PedCompra, .rows - 1, eColR.PedCompra) = flexAlignCenterTop
        .ColHidden(eColR.PedCompra) = True
        
        .ColWidth(eColR.Observacao) = 1200
        .ColDataType(eColR.Observacao) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Observacao, .rows - 1, eColR.Observacao) = flexAlignCenterTop
        .ColHidden(eColR.Observacao) = True
        
        .ColWidth(eColR.Exclui) = 1200
        .ColDataType(eColR.Exclui) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Exclui, .rows - 1, eColR.Exclui) = flexAlignCenterTop
        .ColHidden(eColR.Exclui) = True
        
        .ColWidth(eColR.EmpMrc) = 1200
        .ColDataType(eColR.EmpMrc) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.EmpMrc, .rows - 1, eColR.EmpMrc) = flexAlignCenterTop
        
        .ColWidth(eColR.EmpMrcComp) = 1200
        .ColDataType(eColR.EmpMrcComp) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.EmpMrcComp, .rows - 1, eColR.EmpMrcComp) = flexAlignCenterTop
        .ColHidden(eColR.EmpMrcComp) = True
        
        .ColWidth(eColR.Estoque) = 1200
        .ColDataType(eColR.Estoque) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Estoque, .rows - 1, eColR.Estoque) = flexAlignCenterTop
        
        .ColWidth(eColR.Cdp) = 1000
        .ColDataType(eColR.Cdp) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Cdp, .rows - 1, eColR.Cdp) = flexAlignCenterTop
        .ColHidden(eColR.Cdp) = True
        
        .ColWidth(eColR.ObservItem) = 1000
        .ColDataType(eColR.ObservItem) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.ObservItem, .rows - 1, eColR.ObservItem) = flexAlignCenterTop
        .ColHidden(eColR.ObservItem) = True
        
        .ColWidth(eColR.Status) = 1000
        .ColDataType(eColR.Status) = flexDTString
        .Cell(flexcpAlignment, 0, eColR.Status, .rows - 1, eColR.Status) = flexAlignCenterTop
        .ColHidden(eColR.Status) = True
        
        .Cell(flexcpFontBold, .rows - 1, eColR.Grupo, .rows - 1, eColR.LastCol) = True
        
    End With
    '
    GoTo Fim
Erro:
   Err.Raise ErrorCode.Padrao, Me.Name, "Erro ao montar o grid de informações!"
Fim:

End Sub

Private Function MFcn_MontaQuery() As String
    Dim WLStr_Sql       As String
    Dim WLStr_NotIn     As String
    
    On Local Error GoTo Erro
    
    MFcn_MontaQuery = ""
    
    If WGStr_gerModSis = "FAN" Then
        WLStr_NotIn = "('F','E','P') "
    Else
        WLStr_NotIn = "('F' , 'E') "
    End If
    
    WLStr_Sql = "select distinct " & _
            "a.rqs_Numero,a.ent_Codigo,a.rqs_DataRf,a.rqs_Solici," & _
            "a.rqs_Gerent,a.rqs_Obser1,a.rqs_Obser2,a.rqs_Obser3," & _
            "b.emp_Codigo,b.mrc_Codigo,b.gpx_Codigo,b.prx_Codigo," & _
            "b.nor_Codigo,b.irq_Quanti,b.irq_DtNece,b.irq_Status," & _
            "b.irq_PrUnit,b.irq_AlqIPI,b.for_Codigo,b.irq_CondPg," & _
            "b.cdp_Codigo,c.prx_Descri,c.prx_PrUnit,c.cfi_AlqIPI," & _
            "d.for_NomeCp,e.gpx_Descri,f.cdp_Descri,b.irq_Observ," & _
            "a.rqs_EmpCom,a.rqs_MrcCom " & _
        "from " & _
            "ALMOX_REQ_COMPRAS a " & _
            "left join ALMOX_ITREQ_COMPRAS b on b.rqs_Numero = a.rqs_Numero " & _
            "left join FORNECEDORES d ON d.for_Codigo = b.for_Codigo " & _
            "left join GRUPOS_ALMOXARIFADO e ON e.gpx_Codigo = b.gpx_Codigo " & _
            "left join COND_PAGTOS f ON f.cdp_Codigo = b.cdp_Codigo " & _
            "left join PROD_ALMOXARIFADO c on c.emp_Codigo = b.emp_Codigo " & _
                "and c.mrc_Codigo = b.mrc_Codigo and c.gpx_Codigo = b.gpx_Codigo " & _
                "and c.prx_Codigo = b.prx_Codigo " & _
        "where " & _
            "a.rqs_Numero = '" & Format(Val(Trim(txt_rqsNumero.Text)), "000000") & "' and " & _
            "b.irq_PCpNro <= 0 and " & _
            "b.irq_Status not in " & WLStr_NotIn & _
        "group by rqs_Numero,emp_Codigo,mrc_Codigo,gpx_Codigo,prx_Codigo,irq_DtNece " & _
        "order by gpx_Codigo,prx_Codigo,irq_DtNece "
    
    GoTo Fim
Erro:
    Exit Function
Fim:
    MFcn_MontaQuery = WLStr_Sql
End Function
Private Function MFcn_CarregaDadosReq(PFStr_Query As String) As Boolean
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLObj_DataRf    As DataHora
    Dim WLLng_TotRow    As Long
    Dim WLLng_Cont      As Long
    Dim WLStr_Observ    As String
    Dim WLStr_EmpMrc    As String
    Dim WLStr_Endere    As String
    Dim WLStr_EmpComp   As String
    Dim WLInt_Cont      As Integer
    
    On Local Error GoTo Erro
    
    MFcn_CarregaDadosReq = False
        
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(PFStr_Query, WLLng_TotRow, adCmdText)
    
    With grd_Requisicoes
        If WLLng_TotRow > 0 Then
            
            WLRst_Tabela.MoveFirst
            Do While Not WLRst_Tabela.EOF
                .rows = .rows + 1
                
                If Not MFcn_ValidaStatusRequisicao(WLRst_Tabela!emp_Codigo, WLRst_Tabela!mrc_Codigo, txt_rqsNumero.Text) Then
                    Call MPrc_LimpaCampos
                    fra_Selecao.Enabled = True
                    txt_rqsNumero.Text = Empty
                    txt_rqsNumero.SetFocus
                    WLRst_Tabela.Close
                    Set WLRst_Tabela = Nothing
                    Exit Function
                End If
                 
                .TextMatrix(.rows - 1, eColR.Selecao) = False
                .TextMatrix(.rows - 1, eColR.Status) = "NÃO"
                .TextMatrix(.rows - 1, eColR.Grupo) = WLRst_Tabela!gpx_Codigo
                .TextMatrix(.rows - 1, eColR.GrpDesc) = WLRst_Tabela!gpx_Descri
                .TextMatrix(.rows - 1, eColR.Codigo) = WLRst_Tabela!prx_Codigo
                .TextMatrix(.rows - 1, eColR.CodDesc) = WLRst_Tabela!prx_Descri
                .TextMatrix(.rows - 1, eColR.DtNeces) = WLRst_Tabela!irq_DtNece
                .TextMatrix(.rows - 1, eColR.Quantidade) = WLRst_Tabela!irq_Quanti
                .TextMatrix(.rows - 1, eColR.Norma) = Format(WLRst_Tabela!nor_Codigo, "00000")
                .TextMatrix(.rows - 1, eColR.EmpMrc) = WLRst_Tabela!emp_Codigo & WLRst_Tabela!mrc_Codigo
                .TextMatrix(.rows - 1, eColR.EmpMrcComp) = WLRst_Tabela!rqs_EmpCom & WLRst_Tabela!rqs_MrcCom
                
                 Call GFKey_PegaEndPrincipalProd(WLRst_Tabela!emp_Codigo, WLRst_Tabela!mrc_Codigo, _
                    WLRst_Tabela!gpx_Codigo, WLRst_Tabela!prx_Codigo, True, WLStr_Endere)
                    
                .TextMatrix(.rows - 1, eColR.Estoque) = GFKEY_CalcEstoque_N(WLRst_Tabela!emp_Codigo, _
                    WLRst_Tabela!mrc_Codigo, WLRst_Tabela!gpx_Codigo, WLRst_Tabela!prx_Codigo, , , WLStr_Endere)
                
                If Val(GFKEY_PreparaValor(WLRst_Tabela!irq_PrUnit)) > 0 Then
                    .TextMatrix(.rows - 1, eColR.PrUnit) = WLRst_Tabela!irq_PrUnit
                    .TextMatrix(.rows - 1, eColR.AlqIPI) = WLRst_Tabela!irq_AlqIPI
                End If
                
                If Val(Trim(WLRst_Tabela!for_Codigo)) > 0 Then
                    .TextMatrix(.rows - 1, eColR.CodFornec) = Format(WLRst_Tabela!for_Codigo, "000000")
                    .TextMatrix(.rows - 1, eColR.NomeFornec) = WLRst_Tabela!for_NomeCP
                End If
                
                If WLRst_Tabela!cdp_Codigo > 0 Then
                    .TextMatrix(.rows - 1, eColR.CondPgto) = WLRst_Tabela!cdp_Codigo
                ElseIf Trim(WLRst_Tabela!irq_CondPg) <> Empty Then
                    .TextMatrix(.rows - 1, eColR.Cdp) = WLRst_Tabela!irq_CondPg
                End If
                
                Set WLObj_DataRf = New DataHora
                WLObj_DataRf.Init WLRst_Tabela!rqs_DataRf
                txt_rqsDataRf.Text = WLObj_DataRf.Data
                
                txt_rqsSolici.Text = Trim(WLRst_Tabela!rqs_Solici)
                txt_rqsGerent.Text = Trim(WLRst_Tabela!rqs_Gerent)
                
                For WLLng_Cont = 0 To 9
                    lbl_obs(WLLng_Cont).Caption = Trim(Mid(Left(Trim(WLRst_Tabela!rqs_Obser1) & _
                        Space(220), 220) & Left(Trim(WLRst_Tabela!rqs_Obser2) & _
                        Space(220), 220) & Left(Trim(WLRst_Tabela!rqs_Obser3) & _
                        Space(220), 220), (WLLng_Cont * 66) + 1, 66))
                Next
                
                WLRst_Tabela.MoveNext
            Loop
            
            WLStr_EmpMrc = .TextMatrix(.rows - 1, eColR.EmpMrc)
            WLStr_EmpComp = .TextMatrix(.rows - 1, eColR.EmpMrcComp)
        Else
            GFKEY_MsgBox "Nenhum dado encontrado nas condições informadas! "
            fra_Selecao.Enabled = True
            txt_rqsNumero.Text = Empty
            Exit Function
        End If
    End With
    
    With cbo_empCodigo
        If .ListCount > 0 Then
            For WLInt_Cont = 0 To .ListCount - 1
                If Format(.ItemData(WLInt_Cont), "0000") = WLStr_EmpMrc Then
                    .ListIndex = WLInt_Cont
                    WMStr_EmpCod = Left(Format(cbo_empCodigo.ItemData(cbo_empCodigo.ListIndex), "0000"), 2)
                    WMStr_MrcCod = Right(Format(cbo_empCodigo.ItemData(cbo_empCodigo.ListIndex), "0000"), 2)
                    Exit For
                End If
            Next
        End If
    End With
    
    With cbo_EmpComp
        If .ListCount > 0 Then
            For WLInt_Cont = 0 To .ListCount - 1
                If Format(.ItemData(WLInt_Cont), "0000") = WLStr_EmpComp Then
                    .ListIndex = WLInt_Cont
                    Exit For
                End If
            Next
        End If
    End With
    
    pnl_Requisicoes.Enabled = True
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    MFcn_CarregaDadosReq = True
End Function
Private Sub MPrc_LimpaCampos()
    
    cbo_EmpComp.ListIndex = -1
    cbo_empCodigo.ListIndex = -1
    chk_Despesas.Value = 0
    
    txt_rqsNumero.Text = Empty
    txt_rqsSolici.Text = Empty
    txt_rqsAprova.Text = Empty
    txt_rqsDataRf.Text = Empty
    txt_rqsGerent.Text = Empty
    fra_Selecao.Enabled = True
    pnl_Cotacoes.Enabled = False
    pnl_Requisicoes.Enabled = False
    cmd_confirmar.Enabled = False
    
    Call MPrc_HabilitaTeclasFunc(False)
    Call MFcn_CarregaGruposDespesas
    Call MPrc_MontaGridReq
    Call MPrc_MontaGridCot
    
End Sub

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

Private Sub MPrc_GravaCotSel(Optional PPLng_Linha As Long)
    Dim WLLng_Linha     As Long
    Dim WLStr_Sql       As String
    Dim WLStr_CodFor    As String
    Dim WLStr_PrUnit    As String
    Dim WLStr_DtPrev    As String
    Dim WLStr_AlqIPI    As String
    Dim WLStr_CondPg    As String
    Dim WLStr_CodCpg    As String
    Dim WLStr_Msg       As String
    Dim WLStr_Titulo    As String
    Dim WLStr_Observ    As String
    
    On Local Error GoTo Erro
    
    With grd_Cotacoes
        WLStr_CodFor = .TextMatrix(PPLng_Linha, eColC.CodFornec)
        WLStr_PrUnit = .TextMatrix(PPLng_Linha, eColC.PrFornec)
        WLStr_DtPrev = Format(.TextMatrix(PPLng_Linha, eColC.DiasEntr), "yyyymmdd")
        WLStr_AlqIPI = .TextMatrix(PPLng_Linha, eColC.ValIPI)
        WLStr_CondPg = .TextMatrix(PPLng_Linha, eColC.CondPgto)
        WLStr_CodCpg = .TextMatrix(PPLng_Linha, eColC.CodCdp)
    End With
    
    Call MFcn_VerificaSelecao(grd_Requisicoes, WLLng_Linha)
    
    With grd_Requisicoes
        .TextMatrix(WLLng_Linha, eColR.CodFornec) = WLStr_CodFor
        .TextMatrix(WLLng_Linha, eColR.PrUnit) = WLStr_PrUnit
        .TextMatrix(WLLng_Linha, eColR.DtPrev) = WLStr_DtPrev
        .TextMatrix(WLLng_Linha, eColR.AlqIPI) = WLStr_AlqIPI
        .TextMatrix(WLLng_Linha, eColR.CondPgto) = WLStr_CondPg
        .TextMatrix(WLLng_Linha, eColR.Cdp) = WLStr_CodCpg
        .TextMatrix(WLLng_Linha, eColR.Status) = "SIM"
        
        WLStr_Msg = "Digite a observação para este item: "
        WLStr_Titulo = "Keysystems Informática - Requisição"
        WLStr_Observ = InputBox(WLStr_Msg, WLStr_Titulo, Trim(WLStr_Observ))
        
        .TextMatrix(WLLng_Linha, eColR.ObservItem) = Left(WLStr_Observ & Space(210), 210)
        
        .Cell(flexcpForeColor, WLLng_Linha, eColR.Grupo, WLLng_Linha, eColR.LastCol) = vbBlue
    End With

    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
Fim:

End Sub
Private Function MFcn_ValidaRequisicao() As Boolean
    Dim WLLng_Cont          As Long
    Dim WLBol_Selecionada   As Boolean
    
    MFcn_ValidaRequisicao = False
    
    With grd_Requisicoes
        WLBol_Selecionada = False
        For WLLng_Cont = 1 To .rows - 1
            If UCase(.TextMatrix(WLLng_Cont, eColR.Status)) = "SIM" Then
                WLBol_Selecionada = True
            End If
        Next
        
        If WLBol_Selecionada = False Then
            GFKEY_MsgBox "Selecione pelo menos uma cotação para um item! "
            grd_Requisicoes.SetFocus
            Exit Function
        End If
    End With
    
    If cbo_EmpComp.ListIndex = -1 Then
        GFKEY_MsgBox "Selecione uma empresa compradora! "
        cbo_EmpComp.SetFocus
        Exit Function
    End If
    
    If cbo_GrDesp.ListIndex = -1 Then
        GFKEY_MsgBox "Selecione o Grupo e o Tipo de Despesa! "
        cbo_GrDesp.SetFocus
        Exit Function
    End If
        
    
    If cbo_TpDesp.ListIndex = -1 Then
        GFKEY_MsgBox "Selecione o Tipo de Despesa! "
        cbo_TpDesp.SetFocus
        Exit Function
    End If
    
    With grd_Requisicoes
        For WLLng_Cont = 1 To .rows - 1
            If UCase(.TextMatrix(WLLng_Cont, eColR.Status)) = "NÃO" Then
                If GFKEY_MsgBox("Existem itens nesta requisição pendentes de finalização." & vbCrLf & vbCrLf & _
                    "Deseja continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes _
                Then
                    MFcn_ValidaRequisicao = True
                Else
                    MFcn_ValidaRequisicao = False
                End If
                
                Exit For
            Else
                MFcn_ValidaRequisicao = True
            End If
        Next
    End With
End Function
Private Sub MPrc_VerTeclaAtalho(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyF5
            If lbl_funcompras.Visible = True Then
                If MFcn_AtualizarUltimasCompras() Then
                    With pic_ucp
                        .Top = (Me.Height - pic_ucp.Height) / 2
                        .Left = (Me.Width - pic_ucp.Width) / 2
                        .Visible = True
                        .SetFocus
                    End With
                End If
            End If
            
        Case vbKeyEscape
            pic_ucp.Visible = False
            pic_observacao.Visible = False
            
        Case vbKeyF6
            If lbl_FuncObsReq.Visible = True Then
                lbl_obs_titulo01.Caption = "Observação da Requisição"
                pic_observacao.Top = (Me.Height - pic_observacao.Height) / 2
                pic_observacao.Left = (Me.Width - pic_observacao.Width) / 2
                pic_observacao.Visible = True
            End If

    End Select
End Sub
Private Function MFcn_AtualizarUltimasCompras() As Boolean
    Dim WLRst_Tabela     As ADODB.Recordset
    Dim WLObj_Data      As DataHora
    Dim WLInt_Cont      As Integer
    Dim WLint_Cont2     As Integer
    
    Dim WLLng_Row       As Long
    Dim WLLng_TotRow    As Long
    
    Dim WLStr_DataPed   As String
    Dim WLStr_Auxi      As String
    Dim WLStr_EmpCod    As String
    Dim WLStr_MrcCod    As String
    Dim WLStr_GpxCod    As String
    Dim WLStr_PrxDesc   As String
    Dim WLStr_Sql       As String
    Dim WLStr_PrxCod    As String
        
    MFcn_AtualizarUltimasCompras = False
    
    With grd_Requisicoes
        
        If Not MFcn_VerificaSelecao(grd_Requisicoes, WLLng_Row) Then Exit Function
        
        WLStr_EmpCod = Left(.TextMatrix(WLLng_Row, eColR.EmpMrc), 2)
        WLStr_MrcCod = Right(.TextMatrix(WLLng_Row, eColR.EmpMrc), 2)
        WLStr_GpxCod = Format(.TextMatrix(WLLng_Row, eColR.Grupo), "00")
        WLStr_PrxDesc = Trim(.TextMatrix(.Row, eColR.CodDesc))
        WLStr_PrxCod = .TextMatrix(WLLng_Row, eColR.Codigo)
        
        lbl_ucp_empDescri.Caption = Trim(cbo_empCodigo.Text)
        lbl_ucp_gpxDescri.Caption = Trim(.TextMatrix(WLLng_Row, eColR.GrpDesc))
        lbl_ucp_prxCodigo.Caption = Trim(.TextMatrix(WLLng_Row, eColR.Codigo))
        lbl_ucp_prxDescri.Caption = Trim(.TextMatrix(WLLng_Row, eColR.CodDesc))
        lbl_ucp_irqQuanti.Caption = Trim(.TextMatrix(WLLng_Row, eColR.Quantidade))
        
        WLStr_Auxi = GFKEY_PreparaValor(Format(0, "#0.000"))
        
        If Not WGCnx_DBSegu Is Nothing Then
            WLStr_Auxi = GFKEY_CalcEstoque_N(WLStr_EmpCod, WLStr_MrcCod, WLStr_GpxCod, WLStr_PrxDesc)
        End If
         
        WLStr_Sql = "select distinct " & _
                "a.pcp_Numero,a.gpx_Codigo,a.prx_Codigo," & _
                "a.mre_Quanti ipc_Quanti,a.mre_PrUnit ipc_PrUnit," & _
                "a.for_Codigo,if(b.for_Codigo is null,space(40),b.for_NomeCp) for_NomeCp," & _
                "a.mre_DtPedi pcp_DtPedi,a.mre_CondPg pcp_CondPg " & _
            "from " & _
                "ALMOX_MAT_RECEB a " & _
                    "left join " & _
                "FORNECEDORES b on b.for_Codigo = a.for_Codigo " & _
            "where " & _
                "a.emp_Codigo = '" & WLStr_EmpCod & "' and " & _
                "a.mrc_Codigo = '" & WLStr_MrcCod & "' and " & _
                "a.gpx_Codigo = " & WLStr_GpxCod & " and " & _
                "a.prx_Codigo = '" & WLStr_PrxCod & "' " & _
            "order by mre_DatNTF desc " & _
            "limit 0,7 "
        
        Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
        
        lbl_ucp_prxSalEst.Caption = Format(Val(GFKEY_PreparaValor(Trim(WLStr_Auxi))), _
            IIf(Val(Right(WLStr_Auxi, 3)) > 0, "###,###,##0.000", "###,###,##0"))
        
        WLInt_Cont = 1
        WLint_Cont2 = 0
        
        Do While Not WLRst_Tabela.EOF
            
            Set WLObj_Data = New DataHora
            WLObj_Data.Init WLRst_Tabela!pcp_DtPedi
            WLStr_DataPed = WLObj_Data.Data
            
            lbl_ucp_pcpDtPedi(WLInt_Cont - 1).Caption = WLStr_DataPed
            
            lbl_ucp_forNomeCp(WLInt_Cont - 1).Caption = Trim(WLRst_Tabela!for_NomeCp)
            lbl_ucp_ipcPrUnit(WLInt_Cont - 1).Caption = Format(WLRst_Tabela!ipc_PrUnit, "###,###,##0.0000")
            
            lbl_ucp_ipcQuanti(WLInt_Cont - 1).Caption = Format(WLRst_Tabela!ipc_Quanti, _
                IIf(Val(Right(Format(WLRst_Tabela!ipc_Quanti, "#0.000"), 3)) > 0, _
                "##,###,##0.000", "###,###,##0"))
                
            lbl_ucp_pcpCondPg(WLInt_Cont - 1).Caption = Trim(WLRst_Tabela!pcp_CondPg)
            
            WLInt_Cont = WLInt_Cont + 1
            WLint_Cont2 = WLint_Cont2 + 1
            WLRst_Tabela.MoveNext
            
            Set WLObj_Data = Nothing
        Loop
         
        WLRst_Tabela.Close
        Set WLRst_Tabela = Nothing
        
        For WLint_Cont2 = WLint_Cont2 To lbl_ucp_pcpDtPedi.UBound
            lbl_ucp_forNomeCp(WLint_Cont2).Caption = ""
            lbl_ucp_pcpDtPedi(WLint_Cont2).Caption = ""
            lbl_ucp_ipcPrUnit(WLint_Cont2).Caption = ""
            lbl_ucp_ipcQuanti(WLint_Cont2).Caption = ""
            lbl_ucp_pcpCondPg(WLint_Cont2).Caption = ""
        Next
        
        MFcn_AtualizarUltimasCompras = True
    End With
    
End Function
Private Function MFcn_GravaLog() As Boolean
    Dim WLLng_Cont      As Long
    Dim WLStr_NumPed    As String
    Dim WLStr_NumReq    As String
    Dim WLStr_Log       As String
    
    On Local Error GoTo Erro
    
    MFcn_GravaLog = False
    
    With grd_Requisicoes
        For WLLng_Cont = 1 To .rows - 1
            WLStr_Log = "Fechamento de Requisição de Compras: N° REQUISIÇÃO " & Format(txt_rqsNumero.Text, "000000")
            
            If Not GFKEY_GravaLog _
                ("CPR", Format(Date, "yyyymmdd"), Format(time, "hhmmss"), WGStr_Usuario, WLStr_Log, "A", Me.Name) _
            Then GoTo Erro
        
        Next
    End With
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    MFcn_GravaLog = True
End Function

Private Sub txt_rqsNumero_GotFocus()
    txt_rqsNumero.BackColor = vbYellow
End Sub

Private Sub txt_rqsNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysA vbKeyTab, True
        SendKeysA vbKeyTab, False
    Else
        If KeyAscii = vbKeyExecute Then
            KeyAscii = 0
            Call MPrc_PesquisaRequisicao
        Else
            If KeyAscii <> vbKeyBack And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_rqsNumero_LostFocus()
    Dim WLStr_Sql   As String
    
    On Local Error GoTo Erro
    
    txt_rqsNumero.BackColor = vbWhite
    
    If txt_rqsNumero.Text <> Empty Then
        txt_rqsNumero.Text = Format(Val(Trim(txt_rqsNumero.Text)), "000000")
        
        fra_Selecao.Enabled = False
        
        Call MPrc_MontaGridReq
        WLStr_Sql = MFcn_MontaQuery
        If MFcn_CarregaDadosReq(WLStr_Sql) Then cmd_confirmar.Enabled = True
    
    End If
    
    GoTo Fim
Erro:
    If Err.Number <> 0 Then
        GFKEY_MsgBox "Erro ao carregar informações! " & vbCrLf & vbCrLf & Err.Description
        Exit Sub
    End If
Fim:
    
End Sub
Private Sub MPrc_HabilitaTeclasFunc(PPBol_Habilita As Boolean)
    lbl_funcompras.Visible = PPBol_Habilita
    lbl_FuncObsReq.Visible = PPBol_Habilita
End Sub
Private Sub MPrc_CarregaCotacoes(PPLng_Row As Long)
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLObj_PrCIPI    As cPrecoUnitario
    
    Dim WLStr_Sql       As String
    Dim WLStr_DataPrev  As String
    Dim WLLng_TotRow    As Long
    Dim WLLng_Linha     As Long
    
    Dim WLDbl_Qtd       As Double
    Dim WLDbl_VlICMS    As Double
    Dim WLDbl_VlCstRl   As Double
    Dim WLDbl_AlqICMS   As Double
    Dim WLDbl_PrUnit    As Double
    
    Dim WLCur_VTotal    As Currency
    Dim WLCur_TotalCIPI As Currency
    
    On Local Error GoTo Erro
    
    Call MPrc_MontaGridCot
    
    With grd_Requisicoes
        WLStr_Sql = "select  " & _
               "a.rqs_Numero,a.crq_PrUnit,a.crq_PrazoE," & _
               "a.crq_CondPg,a.for_Codigo,b.for_NomeCP," & _
               "a.crq_AlqIPI,a.crq_AlqICMS,a.cdp_Codigo," & _
               "a.crq_CrdIPI,a.crq_CrdICMS " & _
           "from " & _
               "ALMOX_CTREQ_COMPRAS a " & _
                   "join " & _
               "FORNECEDORES b on b.for_Codigo = a.for_Codigo " & _
           "where " & _
               "a.emp_Codigo = '" & WMStr_EmpCod & "' and " & _
               "a.mrc_Codigo = '" & WMStr_MrcCod & "' and " & _
               "a.rqs_Numero = '" & Format(txt_rqsNumero.Text, "000000") & "' and " & _
               "a.prx_Codigo = '" & .TextMatrix(PPLng_Row, eColR.Codigo) & "' and " & _
               "a.gpx_Codigo = '" & .TextMatrix(PPLng_Row, eColR.Grupo) & "' "
                
        Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
    End With
    
    With grd_Cotacoes
        If Not WLRst_Tabela Is Nothing Then
            If WLLng_TotRow > 0 Then
                WLRst_Tabela.MoveFirst
                Do While Not WLRst_Tabela.EOF
                    
                    Set WLObj_PrCIPI = New cPrecoUnitario
                    WLObj_PrCIPI.Init WLRst_Tabela!crq_PrUnit, WLRst_Tabela!crq_AlqIPI
                    
                    WLDbl_Qtd = Val(grd_Requisicoes.TextMatrix(PPLng_Row, eColR.Quantidade))
                    WLCur_TotalCIPI = (WLObj_PrCIPI.ComIpi * WLDbl_Qtd)
                    
                    WLCur_VTotal = Val(grd_Requisicoes.TextMatrix(PPLng_Row, eColR.Quantidade)) * _
                        WLRst_Tabela!crq_PrUnit
                    
                    If WLRst_Tabela!crq_PrazoE <> "" Then
                        WLStr_DataPrev = DateAdd("d", WLRst_Tabela!crq_PrazoE, Format(Date, "dd/mm/yyyy"))
                    End If
                    
                    WLDbl_PrUnit = WLRst_Tabela!crq_PrUnit
                    WLDbl_AlqICMS = WLRst_Tabela!crq_AlqICMS
                    WLDbl_VlICMS = WLDbl_PrUnit * WLDbl_AlqICMS / 100
                    WLDbl_VlCstRl = WLDbl_PrUnit - WLDbl_VlICMS
                    
                    .rows = .rows + 1
                    
                    .TextMatrix(.rows - 1, eColC.Selecao) = False
                    .TextMatrix(.rows - 1, eColC.CodFornec) = WLRst_Tabela!for_Codigo
                    .TextMatrix(.rows - 1, eColC.DescFor) = WLRst_Tabela!for_NomeCp
                    .TextMatrix(.rows - 1, eColC.PrFornec) = WLRst_Tabela!crq_PrUnit
                    .TextMatrix(.rows - 1, eColC.ValTot) = WLCur_VTotal
                    .TextMatrix(.rows - 1, eColC.CondPgto) = WLRst_Tabela!crq_CondPg
                    .TextMatrix(.rows - 1, eColC.ValIPI) = WLRst_Tabela!crq_AlqIPI
                    .TextMatrix(.rows - 1, eColC.DiasEntr) = WLStr_DataPrev
                    .TextMatrix(.rows - 1, eColC.CodCdp) = WLRst_Tabela!cdp_Codigo
                    .TextMatrix(.rows - 1, eColC.VlICMS) = WLRst_Tabela!crq_AlqICMS
                    .TextMatrix(.rows - 1, eColC.CredIPI) = IIf(WLRst_Tabela!crq_CrdIPI = "S", True, False)
                    .TextMatrix(.rows - 1, eColC.CredICMS) = IIf(WLRst_Tabela!crq_CrdICMS = "S", True, False)
                    .TextMatrix(.rows - 1, eColC.PrCIPI) = WLObj_PrCIPI.ComIpi
                    .TextMatrix(.rows - 1, eColC.TotCIPI) = WLCur_TotalCIPI
                    .TextMatrix(.rows - 1, eColC.VlCustoReal) = Format(WLDbl_VlCstRl, "##,###,###.00")
                    
                    WLRst_Tabela.MoveNext
                Loop
            End If
        End If
    End With
    
    pnl_Cotacoes.Enabled = True
    
    WLRst_Tabela.Close
    Set WLRst_Tabela = Nothing
        
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
Fim:
End Sub
Private Function MFcn_TrataDados() As Boolean
    Dim WLStr_Sql           As String
    Dim WLStr_EmpComp       As String
    Dim WLStr_MrcComp       As String
    Dim WLStr_EmpCod        As String
    Dim WLStr_MrcCod        As String
    Dim WLStr_GrDesp        As String
    Dim WLStr_TpDesp        As String
    
    Dim WLBol_Finalizada    As Boolean
    Dim WLLng_Cont          As Long
    
    On Local Error GoTo Erro
    
    MFcn_TrataDados = False
    
    WLStr_EmpComp = Left(Format(cbo_EmpComp.ItemData(cbo_EmpComp.ListIndex), "0000"), 2)
    WLStr_MrcComp = Right(Format(cbo_EmpComp.ItemData(cbo_EmpComp.ListIndex), "0000"), 2)
    
    WLStr_EmpCod = Left(Format(cbo_empCodigo.ItemData(cbo_empCodigo.ListIndex), "0000"), 2)
    WLStr_MrcCod = Right(Format(cbo_empCodigo.ItemData(cbo_empCodigo.ListIndex), "0000"), 2)
    
    WLStr_GrDesp = cbo_GrDesp.ItemData(cbo_GrDesp.ListIndex)
    WLStr_TpDesp = cbo_TpDesp.ItemData(cbo_TpDesp.ListIndex)
    
    With grd_Requisicoes
        WLStr_Sql = "update ALMOX_REQ_COMPRAS " & _
            "set " & _
                "rqs_EmpCom = '" & WLStr_EmpComp & "'," & _
                "rqs_MrcCom = '" & WLStr_MrcComp & "', " & _
                "gds_Codigo = " & WLStr_GrDesp & ", " & _
                "tds_Codigo = " & WLStr_TpDesp & " " & _
            "where " & _
                "rqs_Numero = '" & Format(txt_rqsNumero.Text, "000000") & "'"
    
        If Not GFKEY_TrataRegistro("A", WLStr_Sql) Then GoTo Erro
    
        For WLLng_Cont = 1 To .rows - 1
            If UCase(.TextMatrix(WLLng_Cont, eColR.Status)) = "SIM" Then
                WLBol_Finalizada = True
            Else
                WLBol_Finalizada = False
            End If
            
            If WLBol_Finalizada Then
                WLStr_Sql = "update ALMOX_CTREQ_COMPRAS set " & _
                        "crq_CotSel = 'S' " & _
                    "where " & _
                        "emp_Codigo = '" & WLStr_EmpCod & "' and " & _
                        "mrc_Codigo = '" & WLStr_MrcCod & "' and " & _
                        "rqs_Numero = '" & Format(txt_rqsNumero.Text, "000000") & "' and " & _
                        "prx_Codigo = '" & .TextMatrix(WLLng_Cont, eColR.Codigo) & "' and " & _
                        "gpx_Codigo = '" & .TextMatrix(WLLng_Cont, eColR.Grupo) & "' and " & _
                        "for_Codigo = '" & .TextMatrix(WLLng_Cont, eColR.CodFornec) & "' and " & _
                        "crq_PrUnit = " & GFKEY_PreparaValor(.TextMatrix(WLLng_Cont, eColR.PrUnit)) & " and " & _
                        "crq_CondPg = '" & .TextMatrix(WLLng_Cont, eColR.CondPgto) & "' and " & _
                        "crq_AlqIPI = " & GFKEY_PreparaValor(.TextMatrix(WLLng_Cont, eColR.AlqIPI)) & " "
                        
                If Not GFKEY_TrataRegistro("A", WLStr_Sql) Then GoTo Erro
                
                WLStr_Sql = "update ALMOX_ITREQ_COMPRAS set " & _
                        "irq_Status = 'P', " & _
                        "for_Codigo = '" & .TextMatrix(WLLng_Cont, eColR.CodFornec) & "', " & _
                        "irq_PrUnit = '" & GFKEY_PreparaValor(.TextMatrix(WLLng_Cont, eColR.PrUnit)) & "', " & _
                        "irq_AlqIPI = '" & GFKEY_PreparaValor(.TextMatrix(WLLng_Cont, eColR.AlqIPI)) & "', " & _
                        "irq_CondPg = '" & .TextMatrix(WLLng_Cont, eColR.CondPgto) & "', " & _
                        "irq_DtPrev = '" & .TextMatrix(WLLng_Cont, eColR.DtPrev) & "', " & _
                        "irq_Observ = '" & .TextMatrix(WLLng_Cont, eColR.ObservItem) & "', " & _
                        "irq_Quanti = '" & .TextMatrix(WLLng_Cont, eColR.Quantidade) & "' " & _
                    "where emp_Codigo = '" & WLStr_EmpCod & "' and " & _
                        "mrc_Codigo = '" & WLStr_MrcCod & "' and " & _
                        "rqs_Numero = '" & Format(txt_rqsNumero.Text, "000000") & "' and " & _
                        "prx_Codigo = '" & .TextMatrix(WLLng_Cont, eColR.Codigo) & "' and " & _
                        "gpx_Codigo = '" & .TextMatrix(WLLng_Cont, eColR.Grupo) & "' "
                
                If Not GFKEY_TrataRegistro("A", WLStr_Sql) Then GoTo Erro
            End If
        Next
    End With
    
    If chk_Despesas.Value = 1 Then
        WLStr_Sql = "delete from TIPOS_DESPESAS_PADRAO"
        If Not GFKEY_TrataRegistro("A", WLStr_Sql) Then GoTo Erro

        WLStr_Sql = "insert into TIPOS_DESPESAS_PADRAO (" & _
                "gds_codigo, tds_codigo" & _
            ")values(" & _
                WLStr_GrDesp & ", " & WLStr_TpDesp & ") "
                
          If Not GFKEY_TrataRegistro("A", WLStr_Sql) Then GoTo Erro
    End If
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    MFcn_TrataDados = True
End Function
Private Sub MPrc_BuscaCotSel(PPLng_Row As Long)
    Dim WLLng_Cont      As Long
    Dim WLStr_PrUnit    As String
    Dim WLStr_AlqIPI    As String
    Dim WLStr_CondPg    As String
    Dim WLStr_ForCod    As String
    
    With grd_Requisicoes
        WLStr_ForCod = .TextMatrix(PPLng_Row, eColR.CodFornec)
        WLStr_PrUnit = .TextMatrix(PPLng_Row, eColR.PrUnit)
        WLStr_AlqIPI = .TextMatrix(PPLng_Row, eColR.AlqIPI)
        WLStr_CondPg = .TextMatrix(PPLng_Row, eColR.CondPgto)
    End With
    
    
    With grd_Cotacoes
        For WLLng_Cont = 1 To .rows - 1
            If WLStr_ForCod = .TextMatrix(WLLng_Cont, eColC.CodFornec) And _
                WLStr_PrUnit = .TextMatrix(WLLng_Cont, eColC.PrFornec) And _
                WLStr_AlqIPI = .TextMatrix(WLLng_Cont, eColC.ValIPI) And _
                WLStr_CondPg = .TextMatrix(WLLng_Cont, eColC.CondPgto) _
            Then
                .TextMatrix(WLLng_Cont, eColC.Selecao) = True
                .Cell(flexcpBackColor, WLLng_Cont, eColC.CodFornec, WLLng_Cont, eColC.LastCol) = vbYellow
            End If
        Next
    End With
End Sub
Private Sub MPrc_PesquisaRequisicao()
    
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
            .Params.Status = "'C'"
            .Show vbModal
            Set WGObj_FormPsqRqsPend = Nothing
        End With
        
        Call Form_Activate
    End If

End Sub
Private Function MFcn_CarregaGruposDespesas() As Boolean
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLLng_TotRow    As Long
    Dim WLLng_Cont      As Long
    Dim WLStr_Sql       As String
    
    On Local Error GoTo Erro
    
    MFcn_CarregaGruposDespesas = False
    
    cbo_GrDesp.Clear
    cbo_TpDesp.Clear
                
    WLStr_Sql = "select gds_Codigo,gds_Descri from GRUPOS_DESPESAS order by gds_Descri "
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
        
    If WLLng_TotRow > 0 Then
        WLRst_Tabela.MoveFirst
        
        Do While Not WLRst_Tabela.EOF
            cbo_GrDesp.AddItem WLRst_Tabela!gds_Descri
            cbo_GrDesp.ItemData(cbo_GrDesp.newIndex) = WLRst_Tabela!gds_Codigo
            WLRst_Tabela.MoveNext
        Loop
    End If
    
    WLRst_Tabela.Close
    Set WLRst_Tabela = Nothing
    
    WLLng_TotRow = 0
    WLStr_Sql = "select gds_Codigo,tds_Codigo from TIPOS_DESPESAS_PADRAO "
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
        
    If WLLng_TotRow > 0 Then
        WLRst_Tabela.MoveFirst
        
        For WLLng_Cont = 0 To cbo_GrDesp.ListCount - 1
            If cbo_GrDesp.ItemData(WLLng_Cont) = WLRst_Tabela!gds_Codigo Then
                cbo_GrDesp.ListIndex = WLLng_Cont
            End If
        Next
        
        For WLLng_Cont = 0 To cbo_TpDesp.ListCount - 1
            If cbo_TpDesp.ItemData(WLLng_Cont) = WLRst_Tabela!tds_Codigo Then
                cbo_TpDesp.ListIndex = WLLng_Cont
            End If
        Next
        
        WLRst_Tabela.Close
        Set WLRst_Tabela = Nothing
    End If
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    MFcn_CarregaGruposDespesas = True
End Function
Private Function MFcn_CarregaTipoDespesas() As Boolean
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLLng_TotRow    As Long
    Dim WLStr_Sql       As String
    Dim WLStr_GpDesp    As String

    On Local Error GoTo Erro

    MFcn_CarregaTipoDespesas = False
    
    WLStr_GpDesp = cbo_GrDesp.ItemData(cbo_GrDesp.ListIndex)
    
    cbo_TpDesp.Clear
    
    WLStr_Sql = "select " & _
            "tds_Codigo,tds_Descri " & _
        "from " & _
            "TIPOS_DESPESAS " & _
        "where " & _
            "gds_Codigo = " & WLStr_GpDesp & _
        " order by tds_Descri "
        
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)

    If WLLng_TotRow > 0 Then
        WLRst_Tabela.MoveFirst
        
        Do While Not WLRst_Tabela.EOF
            cbo_TpDesp.AddItem WLRst_Tabela!tds_Descri
            cbo_TpDesp.ItemData(cbo_TpDesp.newIndex) = WLRst_Tabela!tds_Codigo
            WLRst_Tabela.MoveNext
        Loop
    End If

    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    MFcn_CarregaTipoDespesas = True
End Function
Private Function MFcn_ValidaStatusRequisicao(PFStr_EmpCod As String, PFStr_MrcCod As String, PFStr_RqsNum As String) As Boolean
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLStr_Sql       As String
    Dim WLStr_Status    As String
    Dim WLLng_TotRow    As Long
    
    On Local Error GoTo Erro
    
    MFcn_ValidaStatusRequisicao = False
    
    WLStr_Status = "0"
    
    WLStr_Sql = "select distinct " & _
            "if(irq_Status = 'F',3," & _
                "if(irq_Status = 'C',2," & _
                    "if(irq_Status = 'P',2," & _
            "1))) irq_Status " & _
        "from " & _
            "ALMOX_ITREQ_COMPRAS " & _
        "where " & _
            "emp_Codigo = '" & PFStr_EmpCod & "' and " & _
            "mrc_Codigo = '" & PFStr_MrcCod & "' and " & _
            "rqs_Numero = '" & Format(PFStr_RqsNum, "000000") & "' " & _
        "order by irq_Status desc"
        
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
  
    If WLLng_TotRow > 0 Then
        WLRst_Tabela.MoveFirst
        Do While Not WLRst_Tabela.EOF
            
            Select Case WLRst_Tabela!irq_Status
                Case "3"
                    WLStr_Status = IIf(WLStr_Status = "0", "3", WLStr_Status)
                    
                Case "2"
                    WLStr_Status = IIf(WLStr_Status = "0" Or WLStr_Status = "3", "2", WLStr_Status)
                    
                Case Else
                    WLStr_Status = IIf(WLStr_Status = "0" Or WLStr_Status = "3", "0", IIf(WLStr_Status = "2", "1", WLStr_Status))
            End Select
            WLRst_Tabela.MoveNext
        Loop
    End If
    
    WLRst_Tabela.Close
    Set WLRst_Tabela = Nothing
    
    If WLStr_Status <> "2" Then
        Beep
        If WLStr_Status = "3" Then
            GFKEY_MsgBox "Requisição de compra já finalizada!"
        
        ElseIf WLStr_Status = "1" Then
            GFKEY_MsgBox "Requisição de compra com cotação em andamento!"
        Else
            GFKEY_MsgBox "Requisição de compra não cotada!"
        End If
        
        Exit Function
    End If
    
    GoTo Fim
Erro:
   Exit Function
Fim:
    MFcn_ValidaStatusRequisicao = True

End Function
