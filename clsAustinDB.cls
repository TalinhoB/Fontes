VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAustinDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ePropriedade
    Coluna = 0
    Valor = 1
End Enum

Private pConexao    As ADODB.Connection
Private pNomeTabela As String

Private Const MyName As String = "clsAustinDB"

Public Sub Initialize(ByVal aConexao As ADODB.Connection, ByVal aNomeTabela As String)
    Set pConexao = aConexao
    pNomeTabela = aNomeTabela
End Sub

Public Function Insert(aClasseOri As Object) As Boolean
    Dim VLStr_Sql As String
    
    On Error GoTo Erro
    
    VLStr_Sql = MontaQueryInsert(aClasseOri)
    
    pConexao.Execute VLStr_Sql
    
    Insert = True
    
    Exit Function
Erro:
    Insert = False
End Function

Public Function Update(aClasseOri As Object, ByVal aWhere As String) As Boolean
    Dim VLStr_Sql As String
    
    On Error GoTo Erro
    
    VLStr_Sql = MontarQueryUpdate(aClasseOri, aWhere)
    
    pConexao.Execute VLStr_Sql
    
    Update = True
    
    Exit Function
Erro:
    Update = False
End Function

Public Function Delete(ByVal aClausulaWhere As String) As Boolean
    Dim VLStr_Sql As String
    
    On Error GoTo Erro
    
    VLStr_Sql = "DELETE FROM " & pNomeTabela & " WHERE " & aClausulaWhere
    
    pConexao.Execute VLStr_Sql
    
    Delete = True
    
    Exit Function
Erro:
    Delete = False
End Function

Private Function MontaQueryInsert(aClasseOri As Object) As String
    Dim VLCol_Propriedades  As Collection
    Dim VLVar_ListaProp     As Variant
    Dim VLStr_Colunas       As String
    Dim VLStr_Valores       As String
    Dim VLStr_Sql           As String
    
    Set VLCol_Propriedades = PegarPropriedades(aClasseOri)
    
    For Each VLVar_ListaProp In VLCol_Propriedades
        VLStr_Colunas = VLStr_Colunas & VLVar_ListaProp(ePropriedade.Coluna) & ", "
        
        If VarType(VLVar_ListaProp(ePropriedade.Valor)) = vbString Then
            VLStr_Valores = VLStr_Valores & "'" & VLVar_ListaProp(ePropriedade.Valor) & "', "
        Else
            VLStr_Valores = VLStr_Valores & VLVar_ListaProp(ePropriedade.Valor) & ", "
            
        End If
        
    Next VLVar_ListaProp
    
    VLStr_Colunas = Left(VLStr_Colunas, Len(VLStr_Colunas) - 2)
    VLStr_Valores = Left(VLStr_Valores, Len(VLStr_Valores) - 2)
    
    VLStr_Sql = "INSERT INTO " & pNomeTabela & " (" & VLStr_Colunas & ") VALUES (" & VLStr_Valores & ")"
    
    MontaQueryInsert = VLStr_Sql
End Function

Private Function MontarQueryUpdate(aClasseOri As Object, ByVal aClausulaWhere As String) As String
    Dim VLCol_Propriedades  As Collection
    Dim VLStr_Listagem      As String
    Dim VLStr_Sql           As String
    
    Dim VLVar_ListaProp     As Variant
    
    Set VLCol_Propriedades = PegarPropriedades(aClasseOri)
    
    For Each VLVar_ListaProp In VLCol_Propriedades
        
        If VarType(VLVar_ListaProp(ePropriedade.Valor)) = vbString Then
        
            VLStr_Listagem = VLStr_Listagem & _
                VLVar_ListaProp(ePropriedade.Coluna) & " = '" & VLVar_ListaProp(ePropriedade.Valor) & "', "
        Else
            VLStr_Listagem = VLStr_Listagem & _
                VLVar_ListaProp(ePropriedade.Coluna) & " = " & VLVar_ListaProp(ePropriedade.Valor) & ", "
        End If
    Next VLVar_ListaProp
    
    VLStr_Listagem = Left(VLStr_Listagem, Len(VLStr_Listagem) - 2)
    
    VLStr_Sql = "UPDATE " & pNomeTabela & " SET " & VLStr_Listagem & " WHERE " & aClausulaWhere
    
    MontarQueryUpdate = VLStr_Sql
End Function

Private Function PegarPropriedades(aClasseOri As Object) As Collection
    Dim VLCol_Propriedades  As New Collection
    Dim VLObj_PropInfos     As Collection
    
    Dim VLStr_PropValores   As Variant
    Dim VLVar_Item          As Variant
    
    Dim VLStr_PropNome      As String
    
    Set VLObj_PropInfos = CallByName(aClasseOri, "Propriedades", VbMethod)
    
    For Each VLVar_Item In VLObj_PropInfos
        
        VLStr_PropNome = VLVar_Item(ePropriedade.Coluna)
        VLStr_PropValores = VLVar_Item(ePropriedade.Valor)
        
        VLCol_Propriedades.Add Array(VLStr_PropNome, VLStr_PropValores)
        
    Next VLVar_Item
    
    Set PegarPropriedades = VLCol_Propriedades
End Function

