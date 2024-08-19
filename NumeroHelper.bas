Attribute VB_Name = "NumeroHelper"
Option Explicit

Public Function ToNumber(aValor As Variant) As Double
    ToNumber = Val(GFKEY_PreparaValor(CStr(aValor)))
End Function

Public Function Formatar( _
    ByVal aValor As Variant, ByVal aQtdDecimais As Integer, Optional ByVal aDecimaisDinamica As Boolean = False _
) As String
    Dim lValFormatado           As String
    Dim lMascara                As String
    
    lMascara = "###,###,###,###,###,##0" & IIf(aQtdDecimais > 0, "." & GFKEY_Replica("0", aQtdDecimais), "")
    
    lValFormatado = Format(ToNumber(aValor), lMascara)
    
    If aDecimaisDinamica And aQtdDecimais > 0 Then
        If ToNumber(Right(lValFormatado, aQtdDecimais)) <= 0 Then
            lValFormatado = Left$(lValFormatado, Len(lValFormatado) - (aQtdDecimais + 1))
        End If
    End If
    
    Formatar = lValFormatado
End Function

' Não foi possível usar o nome "Round" pois conflita a função built-in do VB6
Public Function Arred(ByVal aValor As Variant, aQtdDecimais As Integer) As Double
    Arred = ToNumber(Formatar(aValor, aQtdDecimais))
End Function
Public Function LimparNumeros(aTexto As Variant) As String

     LimparNumeros = GFKEY_SoNumeros(CStr(aTexto))
    
End Function
