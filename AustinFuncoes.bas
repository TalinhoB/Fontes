Attribute VB_Name = "AustinFuncoes"
Option Explicit

Private Enum eApenas
    Leitura = 1
End Enum

Private Enum eCampo
    Index = 0
    NomeConex = 1
    StringConexao1 = 2
    StringConexao2 = 3
    InfoUsuario = 4
End Enum
Public Function GFcn_MsgBox(Prompt, Optional Buttons As VbMsgBoxStyle = vbInformation, Optional Title = "Austin Assis") As VbMsgBoxResult
    Dim VLStr_Prefixo   As String
    Dim WLStr_Msg       As String
    
    VLStr_Prefixo = "Atenção!" & Chr(10) & Chr(10)
    WLStr_Msg = Prompt
    
    Beep
    GFcn_MsgBox = MsgBox(VLStr_Prefixo & WLStr_Msg, Buttons, Title)
    
End Function

Public Function GFcn_Criptografia(ByVal PFStr_Texto As String, ByVal PFBol_Decodificar As Boolean) As String
    Dim VLLng_Cont              As Long
    Dim VLLng_TamTexto          As Long
    Dim VLStr_ChaveCriptografia As String
    Dim VLStr_Resultado         As String
    Dim VLInt_Caracter          As Integer
    Dim VLStr_Temp              As String

    VLStr_ChaveCriptografia = "key65l91"
    VLStr_Resultado = ""

    If PFBol_Decodificar Then
        VLLng_TamTexto = Len(PFStr_Texto) \ 2
        
        For VLLng_Cont = 1 To VLLng_TamTexto
            VLInt_Caracter = CLng("&H" & Mid(PFStr_Texto, (VLLng_Cont - 1) * 2 + 1, 2))

            VLInt_Caracter = VLInt_Caracter Xor Asc(Mid(VLStr_ChaveCriptografia, (VLLng_Cont - 1) Mod Len(VLStr_ChaveCriptografia) + 1, 1))

            VLStr_Resultado = VLStr_Resultado & Chr(VLInt_Caracter)
        Next
    Else
        VLLng_TamTexto = Len(PFStr_Texto)
        For VLLng_Cont = 1 To VLLng_TamTexto

            VLInt_Caracter = Asc(Mid(PFStr_Texto, VLLng_Cont, 1)) Xor Asc(Mid(VLStr_ChaveCriptografia, (VLLng_Cont - 1) Mod Len(VLStr_ChaveCriptografia) + 1, 1))

            VLStr_Resultado = VLStr_Resultado & Right("0" & Hex(VLInt_Caracter), 2)
        Next
    End If

    GFcn_Criptografia = VLStr_Resultado
End Function

Public Function GFcn_PegaNomeUsuario() As String
    Dim VLStr_NomeUsr   As String
    Dim VLLng_Tamanho   As Long
    
    VLStr_NomeUsr = Space$(255)
    VLLng_Tamanho = Len(VLStr_NomeUsr)
    GetUserName VLStr_NomeUsr, VLLng_Tamanho
    GFcn_PegaNomeUsuario = Trim$(VLStr_NomeUsr)
End Function
Public Function GFcn_CarregaConexoes(PFObj_CboConexao As ComboBox, Optional PFBol_UsrAdmin As Boolean = False)
    Dim VLArr_ConexaoInfo       As Variant
    Dim VLObj_SystemObject      As Object
    Dim VLObj_Arquivo           As Object
    
    Dim VLStr_CaminhoArq        As String
    Dim VLStr_ConteudoArquivo   As String
    
    Dim VLLng_Cont              As Long
    
    On Local Error GoTo Erro
    
    ReDim VGTyp_Conexao(0)
    VLLng_Cont = 0
    
    Set VLObj_SystemObject = CreateObject("Scripting.FileSystemObject")
    
    VLStr_CaminhoArq = App.Path & "\ArqIni.txt"
    Set VLObj_Arquivo = VLObj_SystemObject.OpenTextFile(VLStr_CaminhoArq, eApenas.Leitura)
    
    Do While Not VLObj_Arquivo.AtEndOfStream
        VLStr_ConteudoArquivo = VLObj_Arquivo.ReadLine
        VLArr_ConexaoInfo = Split(VLStr_ConteudoArquivo, ",")
        
        If UBound(VLArr_ConexaoInfo) >= 4 Then
            With VGTyp_Conexao(VLLng_Cont)
                .Index = Trim(VLArr_ConexaoInfo(eCampo.Index))
                .NomeConexao = Trim(VLArr_ConexaoInfo(eCampo.NomeConex))
                .StringConexao1 = GFcn_Criptografia(Trim(VLArr_ConexaoInfo(eCampo.StringConexao1)), True)
                .StringConexao2 = GFcn_Criptografia(Trim(VLArr_ConexaoInfo(eCampo.StringConexao2)), True)

                PFObj_CboConexao.AddItem .NomeConexao
            End With
            
            If UCase(Trim(VLArr_ConexaoInfo(eCampo.InfoUsuario))) = "ADMINISTRADOR" Then
                PFBol_UsrAdmin = True
            End If

        End If
        
        VLLng_Cont = VLLng_Cont + 1
        ReDim Preserve VGTyp_Conexao(UBound(VGTyp_Conexao) + 1) As tTitulosConexao
    Loop
    
    VLObj_Arquivo.Close

    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    Set VLObj_SystemObject = Nothing
    Set VLObj_Arquivo = Nothing
End Function

Public Sub SendKeysTab()
    keybd_event KeyTab, 0, KeyDown, 0
    keybd_event KeyTab, 0, KeyUP, 0
End Sub

Public Function GFcn_CarregarBancos(PPObj_cboBancos As ComboBox) As Boolean
    Dim VLRst_Tabela    As ADODB.Recordset
    Dim VLStr_Sql       As String
    Dim VLLng_TotRow    As Long
    
    On Local Error GoTo Erro
    
    GFcn_CarregarBancos = False
    
    VLStr_Sql = "select Nome, CGC from Bancos where ISNUMERIC(CGC) = 1"
    VLLng_TotRow = GFcn_QtdRegistros(VLStr_Sql)
    
    VLStr_Sql = "select Nome, CGC from Bancos where ISNUMERIC(CGC) = 1 order by Nome, CGC;"
    Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
        
    If VLLng_TotRow > 0 Then
        With PPObj_cboBancos
            VLRst_Tabela.MoveFirst
            
            .Clear
            Do While Not VLRst_Tabela.EOF
               .AddItem VLRst_Tabela!NOME
               .itemData(.NewIndex) = VLRst_Tabela!CGC
               VLRst_Tabela.MoveNext
            Loop
        End With
    End If
    
    VLRst_Tabela.Close
    Set VLRst_Tabela = Nothing
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    GFcn_CarregarBancos = True
End Function

Public Function GFcn_QtdRegistros(PFStr_Query As String) As Long
    Dim VLRst_tabela    As ADODB.Recordset
    
    Dim VLLng_Posicao   As Long
    
    Dim VLStr_Sql       As String
    Dim VLStr_Result    As String
    
    On Local Error GoTo Erro
    
    VLLng_Posicao = InStr(1, PFStr_Query, "from", vbTextCompare)
    VLStr_Result = Mid(PFStr_Query, VLLng_Posicao)
    
    VLStr_Sql = "select count(*) as Total " & VLStr_Result
    Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
    
    If Not VLRst_tabela.EOF Then
        GFcn_QtdRegistros = VLRst_tabela!Total
    End If
    
    GoTo fim
Erro:
    GFcn_QtdRegistros = 0
    Exit Function
Fim:

End Function
Public Function GFcn_CarregaCombos(PFObj_Combo As ComboBox, PPLng_TipoCombo As Long) As Boolean
    Dim VLRst_Tabela    As ADODB.Recordset
    Dim VLStr_Sql       As String
    
    On Local Error GoTo Erro
    
    GFcn_CarregaCombos = False
    
    Select Case PPLng_TipoCombo
        Case eCombo.Estados
            
            VLStr_Sql = "Select UF, Nome from ESTADOS "
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                Do While Not VLRst_Tabela.EOF
                    PFObj_Combo.AddItem TextoHelper.TratarNulo(VLRst_Tabela!UF)
                    
                    VLRst_Tabela.MoveNext
                Loop
            End If
        
        Case eCombo.CodAtividade
            
            VLStr_Sql = "Select Código, Descrição from ATIVIDADES where Cad_Bancos = 'S' "
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                Do While Not VLRst_Tabela.EOF
                    PFObj_Combo.AddItem TextoHelper.TratarNulo(VLRst_Tabela!Código) & " - " & TextoHelper.TratarNulo(VLRst_Tabela!Descrição)
                    
                    VLRst_Tabela.MoveNext
                Loop
            End If
        
        Case eCombo.OrigemCapital
            
            VLStr_Sql = "Select OrCap, Descricao from ORIGEM_CAPITAL "
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                Do While Not VLRst_Tabela.EOF
                    PFObj_Combo.AddItem TextoHelper.TratarNulo(VLRst_Tabela!Orcap) & " - " & TextoHelper.TratarNulo(VLRst_Tabela!Descricao)
                    
                    VLRst_Tabela.MoveNext
                Loop
            End If
            
        Case eCombo.TipoBanco
        
            VLStr_Sql = "Select Descricao from TIPO_BANCO "
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                Do While Not VLRst_Tabela.EOF
                    PFObj_Combo.AddItem TextoHelper.TratarNulo(VLRst_Tabela!Descricao)
                    
                    VLRst_Tabela.MoveNext
                Loop
            End If
            
        Case eCombo.Segmento
        
            VLStr_Sql = "Select Segmento, Descricao from SEGMENTOS "
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                Do While Not VLRst_Tabela.EOF
                    PFObj_Combo.AddItem TextoHelper.TratarNulo(VLRst_Tabela!Segmento) & " - " & TextoHelper.TratarNulo(VLRst_Tabela!Descricao)
                    
                    VLRst_Tabela.MoveNext
                Loop
            End If
            
        Case eCombo.Rating
            
            VLStr_Sql = "Select Rating from RATING_X_SCORE "
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                Do While Not VLRst_Tabela.EOF
                    PFObj_Combo.AddItem TextoHelper.TratarNulo(VLRst_Tabela!Rating)
                    
                    VLRst_Tabela.MoveNext
                Loop
            End If
        
        Case eCombo.Porte
            With PFObj_Combo
                .AddItem "Pequeno"
                .AddItem "Médio"
                .AddItem "Grande"
            End With
            
        Case eCombo.Bancos
            VLStr_Sql = "Select Nome, CGC From Bancos where isnumeric(CGC) = 1 Order By Nome, CGC "
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                With PFObj_Combo
                    .Clear
                    
                    Do While Not VLRst_Tabela.EOF
                        .AddItem VLRst_Tabela!NOME
                        .itemData(.NewIndex) = VLRst_Tabela!Cgc
                        
                        VLRst_Tabela.MoveNext
                    Loop
                End With
            End If
            
        Case eCombo.Cargos
            VLStr_Sql = "Select COD, Nome From CARGOS order by COD;"
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                With PFObj_Combo
                    Do While Not VLRst_Tabela.EOF
                        .AddItem VLRst_Tabela!Cod & " - " & VLRst_Tabela!NOME
                        .itemData(.NewIndex) = VLRst_Tabela!Cod
                        
                        VLRst_Tabela.MoveNext
                     Loop
                End With
            End If
        
        Case eCombo.Nacionalidade
            VLStr_Sql = "select * from Nacionalidade"
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                VLRst_Tabela.MoveFirst
                
                With PFObj_Combo
                    Do While Not VLRst_Tabela.EOF
                        .AddItem VLRst_Tabela!Portugues
                        VLRst_Tabela.MoveNext
                    Loop
                End With
            End If
            
    End Select
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    GFcn_CarregaCombos = True
End Function

Public Function GFcn_CarregaListBox(PFObj_List As ListBox, PPLng_NomeLista As Long, Optional PPStr_Parametro As String) As Boolean
    Dim VLRst_Tabela    As ADODB.Recordset
    
    Dim VLStr_Sql       As String
    Dim VLStr_Where     As String
    
    On Local Error GoTo Erro
    
    GFcn_CarregaListBox = False
    
    Select Case PPLng_NomeLista
        Case eListBox.Bancos
        
            If PPStr_Parametro <> vbNullString Then
                VLStr_Sql = "Select CGC, NOMCIAL from BANCOS where NOMCIAL LIKE '%" & PPStr_Parametro & "%' order by NOMCIAL asc;"
            Else
                VLStr_Sql = "Select CGC, NOMCIAL from BANCOS order by NOMCIAL asc;"
            End If
        
            Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
            
            If Not VLRst_Tabela.EOF Then
                With PFObj_List
                    VLRst_Tabela.MoveFirst
                    
                    .Clear
                    
                    Do While Not VLRst_Tabela.EOF
                        .AddItem TextoHelper.TratarNulo(VLRst_Tabela!Cgc) & " - " & TextoHelper.TratarNulo(VLRst_Tabela!NomCial)
                        VLRst_Tabela.MoveNext
                    Loop
                    
                End With
            End If
    End Select
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
Fim:
    GFcn_CarregaListBox = True
End Function
