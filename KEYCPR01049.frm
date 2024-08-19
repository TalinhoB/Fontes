VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form KEYCPR01049 
   BackColor       =   &H00FDDEC6&
   BorderStyle     =   0  'None
   Caption         =   "Conferência do Pedido de Discos"
   ClientHeight    =   10350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17175
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   17175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame pnl_principal 
      BackColor       =   &H00FDDEC6&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   8715
      Left            =   120
      TabIndex        =   13
      Top             =   420
      Width           =   16995
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
         Height          =   315
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   4
         Top             =   660
         Width           =   4335
      End
      Begin VB.TextBox txt_pcpNumero 
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
         Height          =   315
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txt_forCodigo 
         Alignment       =   1  'Right Justify
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
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   3
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txt_pcpDtPedi 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid gvs_Dados 
         Height          =   7455
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   16755
         _cx             =   268792690
         _cy             =   268776286
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   0
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
      Begin VB.Label lbl_Titulo01 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Nº Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lbl_Titulo02 
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
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lbl_Titulo03 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDEC6&
         Caption         =   "Dt.Pedido Compra"
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
         Left            =   3540
         TabIndex        =   14
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmd_sair 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Sair"
      Height          =   675
      Left            =   960
      Picture         =   "KEYCPR01049.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9240
      Width           =   855
   End
   Begin VB.PictureBox pic_rodape 
      BackColor       =   &H009C832C&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5205
      TabIndex        =   10
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
         TabIndex        =   11
         Top             =   60
         Width           =   1680
      End
      Begin VB.Label lbl_fundo 
         BackColor       =   &H009C832C&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   255
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5200
      Begin VB.Label lbl_Tituloform 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conferência do Pedido de Discos"
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
         Width           =   3210
      End
      Begin VB.Image img_minimizar 
         Height          =   270
         Left            =   4440
         Picture         =   "KEYCPR01049.frx":014A
         ToolTipText     =   "Minimizar"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image img_fechar 
         Height          =   270
         Left            =   4740
         Picture         =   "KEYCPR01049.frx":05CC
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   270
      End
      Begin VB.Label lbl_fundo 
         Appearance      =   0  'Flat
         BackColor       =   &H009C832C&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   120
      Picture         =   "KEYCPR01049.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9240
      Width           =   855
   End
   Begin VB.Shape shp_borda 
      BorderColor     =   &H80000012&
      Height          =   315
      Left            =   5520
      Top             =   9960
      Width           =   255
   End
End
Attribute VB_Name = "KEYCPR01049"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WMObj_Param             As New cParamKEYCPR01049
Private WMBol_Inicio            As Boolean
Private WMBol_DesprezaSaldo     As Boolean

Private Enum eCols
    Empresa = 0
    Grupo = 1
    Produto = 2
    Descricao = 3
    Unidade = 4
    QtdPed = 5
    QtdRecebida = 6
    PesoTeorico = 7
    DiscosQuebra = 8
    PesoPrd = 9
    EmpCod = 10
    MrcCod = 11
    Sequencia = 12
    LastCol = 12
End Enum
'
Private Enum ePed
    Grupo13 = 13
    Grupo25 = 25
End Enum
Public Property Get Param() As cParamKEYCPR01049
    Set Param = WMObj_Param
End Property
Public Property Let Param(Value As cParamKEYCPR01049)
    Set WMObj_Param = Value
End Property

Private Sub MPrc_LimparTela()
    txt_pcpNumero.Text = vbNullString
    txt_pcpDtPedi.Text = vbNullString
    txt_forCodigo.Text = vbNullString
    txt_forNomeCp.Text = vbNullString
End Sub

Private Function MFcn_CarregarPedido() As Boolean
    Dim WLRst_Tabela    As ADODB.Recordset
    Dim WLLng_TotRow    As Long
    Dim WLLng_GrpPed    As Long
    Dim WLStr_Sql       As String
    Dim WLDbl_QtdConf   As Double
    Dim WLDbl_QtdEnt    As Double
    Dim WLDbl_QtdPed    As Double
    Dim WLBol_Conferido As Boolean
    
    On Local Error GoTo Erro
    
    MFcn_CarregarPedido = False
    
    WLStr_Sql = MFcn_MontaQuery(Trim(txt_pcpNumero.Text))
    Set WLRst_Tabela = WGCnx_DBPrim.Execute(WLStr_Sql, WLLng_TotRow, adCmdText)
    
    If Not WLRst_Tabela Is Nothing Then
        If WLLng_TotRow > 0 Then
            cmd_confirmar.Enabled = True
            WLRst_Tabela.MoveFirst
            
            WLDbl_QtdConf = WLRst_Tabela!ipc_QtdCnf
            WLDbl_QtdEnt = WLRst_Tabela!ipc_QtdEnt
            WLDbl_QtdPed = WLRst_Tabela!ipc_Quanti
            
            Do While Not WLRst_Tabela.EOF
                If WLRst_Tabela!ipc_Conferido = "SIM" Then
                    WLBol_Conferido = True
                Else
                    WLBol_Conferido = False
                    Exit Do
                End If
                WLRst_Tabela.MoveNext
            Loop
            
            WLRst_Tabela.MoveFirst
            
            If Not MFcn_ValidaPedido(WLDbl_QtdConf, WLDbl_QtdEnt, WLDbl_QtdPed, WLBol_Conferido) Then
                GoTo Erro
            Else
                txt_pcpNumero.Enabled = False
                txt_forCodigo.Text = WLRst_Tabela!for_Codigo
                txt_forNomeCp.Text = WLRst_Tabela!for_NomeCp
                
                Do While Not WLRst_Tabela.EOF
                    With gvs_Dados
                        
                        If WLRst_Tabela!ipc_Conferido = "NÃO" Then
                            .rows = .rows + 1
                            
                            WLLng_GrpPed = Val(WLRst_Tabela!gpx_Codigo)
                            
                            If WLLng_GrpPed <> ePed.Grupo13 And WLLng_GrpPed <> ePed.Grupo25 Then
                                GFKEY_MsgBox "Pedido contem grupo que não é permitido para esta conferência!"
                                Call MPrc_LimparTela
                                WLRst_Tabela.Close
                                Set WLRst_Tabela = Nothing
                                Exit Do
                                Unload Me
                            End If
                            
                            .TextMatrix(.rows - 1, eCols.Empresa) = WLRst_Tabela!mrc_Descri
                            .TextMatrix(.rows - 1, eCols.Grupo) = WLRst_Tabela!gpx_Codigo
                            .TextMatrix(.rows - 1, eCols.Produto) = WLRst_Tabela!prx_Codigo
                            .TextMatrix(.rows - 1, eCols.Descricao) = WLRst_Tabela!prx_Descri
                            .TextMatrix(.rows - 1, eCols.Unidade) = WLRst_Tabela!prx_Unidad
                            .TextMatrix(.rows - 1, eCols.QtdPed) = WLRst_Tabela!ipc_Quanti
                            .TextMatrix(.rows - 1, eCols.PesoPrd) = WLRst_Tabela!prx_PesoUn
                            .TextMatrix(.rows - 1, eCols.EmpCod) = WLRst_Tabela!emp_Codigo
                            .TextMatrix(.rows - 1, eCols.MrcCod) = WLRst_Tabela!mrc_Codigo
                            .TextMatrix(.rows - 1, eCols.Sequencia) = WLRst_Tabela!ipc_Sequen
                        End If
                    End With
                    WLRst_Tabela.MoveNext
                Loop
            End If
        End If
    Else
        GFKEY_MsgBox "Pedido não encontrado!"
        Call MPrc_LimparCampos
        Exit Function
    End If
    
    WLRst_Tabela.Close
    Set WLRst_Tabela = Nothing
    GoTo Fim
Erro:
    If Err.Number <> 0 Then
        GFKEY_MsgBox "Erro ao carregar pedido"
    End If
      
    Call MPrc_LimparCampos
    Exit Function
Fim:
    MFcn_CarregarPedido = True
End Function
Private Sub MPrc_PrepararGrid()
    Dim WLStr_TituloCols As String
    
    On Local Error GoTo Erro
    
    WLStr_TituloCols = "Empresa|Grupo|Produto|Descrição do Produto|Und|Qtd.Pedido|" & _
        "Qtd. Recebida(UN)|Qtd. Teorica(KG)|Qtd Discos Quebra|" & _
        "Peso Prd|EmpCod|MrcCod|Sequencia"
    
    With gvs_Dados
        .Clear
        .Refresh
        .AutoResize = True
        .AllowSelection = True
        .SelectionMode = flexSelectionByRow
        .BackColorFixed = &HFFC0C0
        .BackColorSel = vbYellow
        .ForeColorSel = vbBlack
        .AutoResize = True
        .FormatString = WLStr_TituloCols
        .ScrollTrack = True
        
        .ColWidth(eCols.Empresa) = 1597
        .ColDataType(eCols.Empresa) = flexDTString
        .Cell(flexcpAlignment, 0, eCols.Empresa, .rows - 1, eCols.Empresa) = flexAlignLeftTop
        
        .ColWidth(eCols.Grupo) = 610
        .ColDataType(eCols.Grupo) = flexDTString
        .Cell(flexcpAlignment, 0, eCols.Grupo, .rows - 1, eCols.Grupo) = flexAlignLeftTop
        
        .ColWidth(eCols.Produto) = 987
        .ColDataType(eCols.Produto) = flexDTString
        .Cell(flexcpAlignment, 0, eCols.Produto, .rows - 1, eCols.Produto) = flexAlignLeftTop
        
        .ColWidth(eCols.Descricao) = 4181
        .ColDataType(eCols.Descricao) = flexDTString
        .Cell(flexcpAlignment, 0, eCols.Descricao, .rows - 1, eCols.Descricao) = flexAlignLeftTop
        
        .ColWidth(eCols.Unidade) = 650
        .ColDataType(eCols.Unidade) = flexDTString
        .Cell(flexcpAlignment, 0, eCols.Unidade, .rows - 1, eCols.Unidade) = flexAlignLeftTop
        
        .ColWidth(eCols.QtdPed) = 1597
        .ColDataType(eCols.QtdPed) = flexDTDouble
        .Cell(flexcpAlignment, 0, eCols.QtdPed, .rows - 1, eCols.QtdPed) = flexAlignLeftTop
        .ColFormat(eCols.QtdPed) = "#,###,###"
        
        .ColWidth(eCols.QtdRecebida) = 1690
        .ColDataType(eCols.QtdRecebida) = flexDTDouble
        .Cell(flexcpAlignment, 0, eCols.QtdRecebida, .rows - 1, eCols.QtdRecebida) = flexAlignLeftTop
        .ColFormat(eCols.QtdRecebida) = "#,###,###"
        
        .ColWidth(eCols.PesoTeorico) = 1650
        .ColDataType(eCols.PesoTeorico) = flexDTDouble
        .Cell(flexcpAlignment, 0, eCols.PesoTeorico, .rows - 1, eCols.PesoTeorico) = flexAlignLeftTop
        .ColFormat(eCols.PesoTeorico) = "#,###,###.00"
        
        .ColWidth(eCols.DiscosQuebra) = 1650
        .ColDataType(eCols.DiscosQuebra) = flexDTDouble
        .Cell(flexcpAlignment, 0, eCols.DiscosQuebra, .rows - 1, eCols.DiscosQuebra) = flexAlignLeftTop
        
        .ColWidth(eCols.PesoPrd) = 1597
        .ColDataType(eCols.PesoPrd) = flexDTDouble
        .Cell(flexcpAlignment, 0, eCols.PesoPrd, .rows - 1, eCols.PesoPrd) = flexAlignLeftTop
        .ColHidden(eCols.PesoPrd) = True
        
        .ColWidth(eCols.EmpCod) = 1597
        .ColDataType(eCols.EmpCod) = flexDTDouble
        .Cell(flexcpAlignment, 0, eCols.EmpCod, .rows - 1, eCols.EmpCod) = flexAlignLeftTop
        .ColHidden(eCols.EmpCod) = True
        
        .ColWidth(eCols.MrcCod) = 1597
        .ColDataType(eCols.MrcCod) = flexDTDouble
        .Cell(flexcpAlignment, 0, eCols.MrcCod, .rows - 1, eCols.MrcCod) = flexAlignLeftTop
        .ColHidden(eCols.MrcCod) = True
        
        .ColWidth(eCols.Sequencia) = 1000
        .ColDataType(eCols.Sequencia) = flexDTString
        .Cell(flexcpAlignment, 0, eCols.Sequencia, .rows - 1, eCols.Sequencia) = flexAlignLeftTop
        .ColHidden(eCols.Sequencia) = True
        
        .Cell(flexcpFontBold, .rows - 1, eCols.Empresa, .rows - 1, eCols.LastCol) = True
    End With
    
    GoTo Fim
Erro:
    GFKEY_MsgBox "Erro ao montar o grid!" & vbCrLf & vbCrLf & "Entre em contato com o suporte"
    gvs_Dados.Clear
    gvs_Dados.rows = 0
    Exit Sub
Fim:
End Sub
    
Private Sub cmd_confirmar_Click()
    Dim WLStr_Log   As String
    Dim WLInt_Cont  As Integer
    
    If txt_pcpNumero.Text = Empty Then
        GFKEY_MsgBox "Para realizar a conferência, informe o pedido!"
        txt_pcpNumero.SetFocus
        Exit Sub
    End If
    
    If MFcn_GravaConferencia Then
        GFKEY_MsgBox "Conferência realizada com sucesso"
    
        With gvs_Dados
            For WLInt_Cont = 1 To .rows - 1
                WLStr_Log = "Conferencia do pedido: " & Trim(txt_pcpNumero.Text) & _
                    " realizada pelo usuário(a) " & WGStr_Usuario & " Item: " & _
                    .TextMatrix(WLInt_Cont, eCols.Produto) & " Grupo: " & _
                    .TextMatrix(WLInt_Cont, eCols.Grupo) & " Quantidade conferida: " & _
                    .TextMatrix(WLInt_Cont, eCols.QtdRecebida)
    
                Call GFKEY_GravaLog(WGStr_Modulo, Format(Date, "yyyymmdd"), _
                    Format(time, "hhmmss"), WGStr_Usuario, WLStr_Log, "A", Me.Name)
            Next
        End With
    Else
        GFKEY_MsgBox "Erro ao gravar conferência"
    End If
    Call MPrc_LimparCampos
End Sub

Private Sub cmd_confirmar_GotFocus()
    cmd_confirmar.BackColor = keyCAmarelo
End Sub

Private Sub cmd_confirmar_LostFocus()
    cmd_confirmar.BackColor = keyCBackColor
End Sub

Private Sub cmd_sair_Click()
    If txt_pcpNumero.Enabled Then
        Unload Me
    Else
        MPrc_LimparCampos
        txt_pcpNumero.Enabled = True
        txt_pcpNumero.SetFocus
    End If
End Sub

Private Sub MPrc_LimparCampos()
    txt_pcpNumero.Text = vbNullString
    txt_pcpDtPedi.Text = vbNullString
    txt_forCodigo.Text = vbNullString
    txt_forNomeCp.Text = vbNullString
    
    txt_pcpNumero.Enabled = True
    txt_pcpNumero.SetFocus
    gvs_Dados.rows = 1
End Sub

Private Sub cmd_sair_GotFocus()
    cmd_sair.BackColor = keyCAmarelo
End Sub

Private Sub cmd_sair_LostFocus()
    cmd_sair.BackColor = keyCBackColor
End Sub

Private Sub Form_Activate()
    If WMBol_Inicio = True Then
        Call GPKEY_CantoRedondo(Me, 25)
        
        Call MPrc_PrepararGrid
        Call MPrc_LimparCampos
        
        If WMObj_Param.PedidoNumero > 0 Then
            txt_pcpNumero.Text = CStr(WMObj_Param.PedidoNumero)
            If Not MFcn_CarregarPedido Then
                Unload Me
            End If
        End If
        
        If Me.Tag <> Empty Then
            txt_pcpNumero.Text = Me.Tag
            Call MFcn_CarregarPedido
        End If
        
        WMBol_Inicio = False
    End If
End Sub

Private Sub Form_Load()
    WMBol_Inicio = True
End Sub

Private Sub gvs_Dados_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim WLStr_QtdQbr    As String
    Dim WLDbl_Calculo   As Double
    Dim WLDbl_QtdPed    As Double
    Dim WLDbl_Peso      As Double
    Dim WLDbl_PesoReceb As Double
    Dim WLDbl_QtdPedido As Double
    Dim WLDbl_QtdReceb  As Double
    
    With gvs_Dados
        If .TextMatrix(Row, Col) <> Empty Then
            WLDbl_QtdPed = .TextMatrix(Row, eCols.QtdPed)
            WLDbl_Peso = .TextMatrix(Row, eCols.PesoPrd)
            
            Select Case Col
                Case eCols.QtdRecebida
                    WLDbl_QtdReceb = .TextMatrix(Row, eCols.QtdRecebida)
                    .TextMatrix(Row, eCols.PesoTeorico) = WLDbl_QtdReceb * WLDbl_Peso
                    
                Case eCols.PesoTeorico
                    WLDbl_PesoReceb = .TextMatrix(Row, eCols.PesoTeorico)
                    
                    If .TextMatrix(Row, eCols.PesoTeorico) = Empty Then
                        .TextMatrix(Row, eCols.QtdRecebida) = Empty
                    Else
                        .TextMatrix(Row, eCols.QtdRecebida) = WLDbl_PesoReceb / WLDbl_Peso
                    End If
                    
            End Select
            
            If Val(.TextMatrix(Row, eCols.Grupo)) = ePed.Grupo13 Then
                WLDbl_QtdReceb = .TextMatrix(Row, eCols.PesoTeorico)
            Else
                WLDbl_QtdReceb = .TextMatrix(Row, eCols.QtdRecebida)
            End If
            
            If WLDbl_QtdReceb < WLDbl_QtdPed Then
                WLStr_QtdQbr = WLDbl_QtdPed - WLDbl_QtdReceb
                .TextMatrix(Row, eCols.DiscosQuebra) = Format(WLStr_QtdQbr, "##,###,###.00")
            Else
                WLStr_QtdQbr = 0
                .TextMatrix(Row, eCols.DiscosQuebra) = Format(WLStr_QtdQbr, "#0")
            End If
        Else
            .TextMatrix(Row, eCols.QtdRecebida) = Empty
            .TextMatrix(Row, eCols.PesoTeorico) = Empty
            .TextMatrix(Row, eCols.DiscosQuebra) = Empty
            
        End If
        
        If Row < .rows - 1 Then
            .Row = Row + 1
            .EditCell
        End If
    End With

End Sub

Private Sub gvs_Dados_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim WLLng_Grupo As Long
    
    WLLng_Grupo = Val(gvs_Dados.TextMatrix(Row, eCols.Grupo))
    
    Select Case Col
        Case eCols.QtdRecebida
            If WLLng_Grupo = ePed.Grupo13 Then
                Cancel = True
            Else
                Cancel = False
            End If
            
        Case eCols.PesoTeorico
            If WLLng_Grupo = ePed.Grupo25 Then
                Cancel = True
            Else
                Cancel = False
            End If
            
        Case Else
            Cancel = True
    End Select
End Sub

Private Sub gvs_Dados_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With gvs_Dados
        If KeyCode = vbKeyReturn Then
            .FinishEditing False
            If Row = .rows - 1 Then
                 SendKeysA vbKeyTab, True
                 SendKeysA vbKeyTab, False
            End If
        End If
    End With
End Sub

Private Sub gvs_Dados_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If KeyAscii <> vbKeyBack Then
        Select Case Col
            Case eCols.PesoTeorico
                If Not IsNumeric(Chr(KeyAscii)) Then
                    If InStr(",", Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    End If
                End If
                
            Case Else
                If Not IsNumeric(Chr(KeyAscii)) Then
                    KeyAscii = 0
                End If
                
        End Select
    End If
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

Private Sub lbl_Titulo01_Click()
    Call MPrc_ChamaPesqPedCpr
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

Private Sub pnl_principal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call GPKEY_MoveTela(Me)
End Sub


Private Sub txt_pcpNumero_GotFocus()
    txt_pcpNumero.BackColor = keyCAmarelo
End Sub

Private Sub MPrc_ChamaPesqPedCpr()

    If Not WGObj_FormPsqPedCpr Is Nothing Then
        Beep
        GFKEY_MsgBox "Pesquisa de pedido de compras já ativa em outro módulo!"
        DoEvents
    Else
        Set WGObj_FormPsqPedCpr = Me
        PSQPEDCOMPRAS.Show vbModal
        Set WGObj_FormPsqPedCpr = Nothing
        DoEvents
        
        If Trim(txt_pcpNumero.Text) <> Empty Then
            txt_pcpNumero.SetFocus
            DoEvents
            SendKeysA vbKeyTab, True: SendKeysA vbKeyTab, False
        End If
    End If
End Sub

Private Sub txt_pcpNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyExecute Then
        KeyAscii = 0
        MPrc_ChamaPesqPedCpr
    Else
        GPKEY_KeyPress txt_pcpNumero, KeyAscii
    End If
End Sub

Private Sub txt_pcpNumero_LostFocus()
    txt_pcpNumero.BackColor = keyCBranco
    If txt_pcpNumero.Text <> Empty And gvs_Dados.rows <= 1 Then
        Call MFcn_CarregarPedido
    End If
End Sub

Private Sub txt_pcpNumero_Validate(Cancel As Boolean)
    Dim WLBol_CarregarPedido As Boolean
    
    If txt_pcpNumero.Text = vbNullString Then
        txt_pcpDtPedi.Text = vbNullString
        txt_forCodigo.Text = vbNullString
        txt_forNomeCp.Text = vbNullString
        Exit Sub
    End If
    
    If Not IsNumeric(txt_pcpNumero.Text) Then
        MsgBox "Pedido inválido", vbInformation, "Keysystems Informática"
        Cancel = True
    End If
End Sub
Private Function MFcn_MontaQuery(PFStr_pedNumero As String) As String
    Dim WLStr_Sql   As String
    Dim WLStr_Cols  As String
        
    On Local Error GoTo Erro
    
    WLStr_Cols = "if(c.ipc_QtdCnf > 0,'SIM','NÃO') ipc_Conferido," & _
        "c.ipc_Sequen,c.emp_Codigo,c.mrc_Codigo,c.gpx_Codigo,c.prx_Codigo," & _
        "c.ipc_Quanti,c.ipc_DtNece,c.ipc_DtPrev,a.pcp_Numero,a.pcp_DtPedi," & _
        "a.for_Codigo,ifnull(b.for_NomeCp, '') for_NomeCp," & _
        "ifnull(d.prx_Descri, '') prx_Descri,ifnull(d.prx_Unidad, '') prx_Unidad," & _
        "if(e.und_ADecim is null,'S',e.und_ADecim) und_ADecim," & _
        "ifnull(f.mrc_Descri, '') mrc_Descri," & _
        "c.ipc_QtdCnf,c.ipc_QtdEnt,d.prx_PesoUn,c.emp_Codigo,c.mrc_Codigo "
    
    WLStr_Sql = "select " & WLStr_Cols & _
        "from ALMOX_PED_COMPRAS a " & _
                "left join " & _
            "FORNECEDORES b ON b.for_Codigo = a.for_Codigo " & _
                "join " & _
            "ALMOX_ITPED_COMPRAS c ON c.pcp_Numero = a.pcp_Numero " & _
                "left join " & _
            "PROD_ALMOXARIFADO d ON d.emp_Codigo = c.emp_Codigo and " & _
                    "d.mrc_Codigo = c.mrc_Codigo and d.gpx_Codigo = c.gpx_Codigo and " & _
                    "d.prx_Codigo = c.prx_Codigo " & _
                "left join " & _
            "UNIDADES e ON e.und_Codigo = d.prx_Unidad " & _
                "left join " & _
            "MARCAS f ON f.emp_Codigo = c.emp_Codigo and " & _
                "f.mrc_Codigo = c.mrc_Codigo " & _
        "where " & _
            "a.pcp_Numero = '" & PFStr_pedNumero & "' " & _
        "order by pcp_Numero,gpx_Codigo,prx_Codigo,ipc_DtPrev "
    
    GoTo Fim
Erro:
    Exit Function
Fim:
    MFcn_MontaQuery = WLStr_Sql
End Function
Private Function MFcn_GravaConferencia() As Boolean
    Dim WLStr_Sql           As String
    Dim WLStr_Cols          As String
    Dim WLStr_QtdDiscos     As String
    Dim WLStr_PesoDiscos    As String
    Dim WLStr_QtdQuebra     As String
    
    Dim WLLng_Cont          As Long
    
    Dim WLDbl_QtdPedido     As Double
    
    
    On Local Error GoTo Erro
    
    MFcn_GravaConferencia = False
    
    WGCnx_DBPrim.BeginTrans
    
    WLStr_Cols = "pcp_Numero,ipc_Sequen," & _
        "con_QtdRec,con_QtdTeo,con_QtdQbr "
    
    With gvs_Dados
        For WLLng_Cont = 1 To .rows - 1
            
            WLStr_QtdDiscos = .TextMatrix(WLLng_Cont, eCols.QtdRecebida)
            WLStr_PesoDiscos = NumeroHelper.Formatar(.TextMatrix(WLLng_Cont, eCols.PesoTeorico), 2)
            WLStr_QtdQuebra = .TextMatrix(WLLng_Cont, eCols.DiscosQuebra)
            
            If Val(WLStr_QtdDiscos) > 0 Then
            
                WLStr_Sql = "insert into CONFERENCIA (" & WLStr_Cols & ") values ('" & _
                    Trim(txt_pcpNumero.Text) & "','" & _
                    .TextMatrix(WLLng_Cont, eCols.Sequencia) & "','" & _
                    GFKEY_PreparaValor(WLStr_QtdDiscos) & "','" & _
                    GFKEY_PreparaValor(WLStr_PesoDiscos) & "','" & _
                    GFKEY_PreparaValor(WLStr_QtdQuebra) & "')"
                
                WGCnx_DBPrim.Execute WLStr_Sql
                If GFKEY_VerificaErro("A") Then GoTo Erro
                
                WLStr_Sql = "update ALMOX_ITPED_COMPRAS set " & _
                        "ipc_QtdCnf = '" & GFKEY_PreparaValor(WLStr_QtdDiscos) & "' " & _
                    "where " & _
                        "emp_Codigo = '" & .TextMatrix(WLLng_Cont, eCols.EmpCod) & "' and " & _
                        "mrc_Codigo = '" & .TextMatrix(WLLng_Cont, eCols.MrcCod) & "' and " & _
                        "pcp_Numero = '" & Trim(txt_pcpNumero.Text) & "' and " & _
                        "prx_Codigo = '" & .TextMatrix(WLLng_Cont, eCols.Produto) & "' and " & _
                        "gpx_Codigo = '" & .TextMatrix(WLLng_Cont, eCols.Grupo) & "' "
                
                WGCnx_DBPrim.Execute WLStr_Sql
                If GFKEY_VerificaErro("A") Then GoTo Erro
            End If
        Next
    End With
    
    GoTo Fim
Erro:
    WGCnx_DBPrim.RollbackTrans
    Exit Function
Fim:
    MFcn_GravaConferencia = True
    WGCnx_DBPrim.CommitTrans
End Function


Private Function MFcn_ValidaPedido(PFDbl_QtdCnf As Double, _
    PFDbl_QtdEnt As Double, PFDbl_QtdPed As Double, PFBol_PedCnf As Boolean _
) As Boolean
    
    On Local Error GoTo Erro
    
    MFcn_ValidaPedido = False
    
    If PFDbl_QtdEnt = PFDbl_QtdPed Then
        GFKEY_MsgBox "Pedido já baixado!"
        GoTo Erro
    End If
    
    If PFDbl_QtdCnf > 0 And PFBol_PedCnf Then
        GFKEY_MsgBox "Pedido já conferido!"
        GoTo Erro
    End If
    
    GoTo Fim
Erro:
    Exit Function
Fim:
    MFcn_ValidaPedido = True
End Function
