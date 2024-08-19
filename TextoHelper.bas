Attribute VB_Name = "TextoHelper"
Option Explicit

Public Function DividirTexto( _
    aTexto As String, aLimiteChars As Long, aCharDivisao As String, Optional aMaxIndices As Integer _
) As String()
    Dim textoOrigem            As String
    Dim textoResidual          As String
    Dim textoFinal()           As String
    Dim posicao                As Long
    Dim limiteQuebra           As Long
    Dim idx                    As Long
    Dim maxIndices             As Integer
    
    textoOrigem = aTexto
    textoResidual = aTexto
    limiteQuebra = IIf(aLimiteChars <= 0, Len(textoOrigem), aLimiteChars)
    maxIndices = IIf(aMaxIndices <= 0, 999, aMaxIndices)
    
    idx = 0
    
    Do
        ReDim Preserve textoFinal(idx) As String
    
        posicao = GFKEY_Obter_PosicaoTextoQuebrado(textoResidual, limiteQuebra, aCharDivisao)
        
        textoFinal(UBound(textoFinal)) = Left$(textoResidual, posicao)
        textoResidual = Mid$(textoResidual, posicao + 1)
        
        idx = idx + 1
    Loop While textoResidual <> Empty And idx < maxIndices
    
    DividirTexto = textoFinal
End Function

Public Function RemoverCaracteresEspeciais(aTexto As String) As String
    Dim TextoOriginal   As String
    Dim TextoNovo       As String
    
    TextoOriginal = aTexto
    TextoNovo = GFKEY_TiraAcentuacao(TextoOriginal)
    
    RemoverCaracteresEspeciais = TextoNovo
End Function

Public Function NormalizarTexto(aTexto As String) As String
    
    Dim TextoNovo       As String
    
    On Local Error GoTo TrataErro
    
    TextoNovo = aTexto
    
    If InStr(TextoNovo, Chr(34)) > 0 Then
        TextoNovo = Replace(TextoNovo, Chr(34), "¨")
    End If
    
    If InStr(TextoNovo, "'") > 0 Then
        TextoNovo = Replace(TextoNovo, "'", "`")
    End If
    
    If InStr(TextoNovo, "\") > 0 Then
        TextoNovo = Replace(TextoNovo, "\", "/")
    End If
    
    NormalizarTexto = TextoNovo
    
TrataErro:
    NormalizarTexto = ""

fim:
    
End Function

