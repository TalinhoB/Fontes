VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotasExplicativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MyName As String = "clsNotasExplicativas"
Private DataBase        As clsAustinDB

Private Enum eParam
    aCgc = 0
    aMes = 1
    aAno = 2
End Enum

Private pClasseInicializada As Boolean
Private pExisteRegistro     As Boolean

Private pCGC                As String
Private pAno                As String
Private pMes                As String
Private pNota               As String

Private Sub Class_Initialize()
    pCGC = ""
    pAno = ""
    pMes = ""
    pNota = ""
End Sub

Public Property Get Cgc() As String
    Cgc = pCGC
End Property

Public Property Let Cgc(Value As String)
    pCGC = Value
End Property

Public Property Get Ano() As String
    Ano = pAno
End Property

Public Property Let Ano(Value As String)
    pAno = Value
End Property

Public Property Get Mes() As String
    Mes = pMes
End Property

Public Property Let Mes(Value As String)
    pMes = Value
End Property

Public Property Get Nota() As String
    Nota = pNota
End Property

Public Property Let Nota(Value As String)
    pNota = Value
End Property

Private Property Get ClasseInicializada() As Boolean
    ClasseInicializada = pClasseInicializada
End Property

Private Property Let ClasseInicializada(Value As Boolean)
    pClasseInicializada = Value
End Property

Private Property Get ExisteRegistro() As Boolean
    ExisteRegistro = pExisteRegistro
End Property

Private Property Let ExisteRegistro(Value As Boolean)
    pExisteRegistro = Value
End Property

Public Sub Init(ByVal aChave As String)
    Dim VLArr_Parametros    As Variant
    
    On Local Error GoTo Erro
    
    ClasseInicializada = False
    
    VLArr_Parametros = Split(aChave, ",")
    
    If UBound(VLArr_Parametros) >= 2 Then
        Cgc = Trim(VLArr_Parametros(eParam.aCgc))
        Mes = Trim(VLArr_Parametros(eParam.aMes))
        Ano = Trim(VLArr_Parametros(eParam.aAno))
    Else
        GFcn_MsgBox "Erro na leitura da chave do registro!" & vbNewLine & "Verifique os parametros informados!"
        Exit Sub
    End If
    
    Set DataBase = New clsAustinDB
    DataBase.Initialize VGCnx_BrasilBank, "Notas"
    
    GoTo Fim
Erro:
    Err.Raise vbObjectError + 500, Err.Source, "Erro ao inicializar a classe: " & MyName & vbNewLine & "Verifique os parâmetros passados."
    Exit Sub
Fim:
  ClasseInicializada = True

End Sub

Public Function Propriedades() As Collection
    Dim VLCol_Props As New Collection
    
    VLCol_Props.Add Array("CGC", TextoHelper.PrepararTexto(Me.Cgc))
    VLCol_Props.Add Array("ANO", TextoHelper.PrepararTexto(Me.Ano))
    VLCol_Props.Add Array("MES", TextoHelper.PrepararTexto(Me.Mes))
    VLCol_Props.Add Array("NOTA", TextoHelper.PrepararTexto(Me.Nota))

    Set Propriedades = VLCol_Props
    
End Function
Public Function Carregar() As Boolean
    Dim VLStr_Sql       As String
    Dim VLRst_Tabela    As ADODB.Recordset
    
    On Error GoTo Erro
    
    Carregar = False
    
    If ClasseInicializada = False Then Exit Function
    
    VLStr_Sql = MontaQuery
    Set VLRst_Tabela = VGCnx_BrasilBank.Execute(VLStr_Sql)
    
    If Not VLRst_Tabela.EOF Then
        VLRst_Tabela.MoveFirst
        
        Cgc = TextoHelper.TratarNulo(VLRst_Tabela!Cgc)
        Ano = TextoHelper.TratarNulo(VLRst_Tabela!Ano)
        Mes = TextoHelper.TratarNulo(VLRst_Tabela!Mes)
        Nota = VLRst_Tabela!Nota
        
        Carregar = True
    Else
        Carregar = False
    End If
    
    ExisteRegistro = Carregar
    
    GoTo Fim
Erro:
    GFcn_MsgBox "Erro ao carregar dados: " & Err.Description
    Carregar = False
Fim:
    Set VLRst_Tabela = Nothing
End Function

Private Function MontaQuery() As String
    Dim VLStr_Sql   As String
    
    VLStr_Sql = "Select * From Notas Where " & Condicao
    
    MontaQuery = VLStr_Sql
End Function

Private Function Condicao() As String
    Dim VLStr_Sql   As String
    
    VLStr_Sql = "CGC = '" & Cgc & "' and ANO = '" & Ano & "' and MES = '" & Mes & "' "
    
    Condicao = VLStr_Sql
End Function

Public Function Gravar(Optional aOperacao As Integer) As Boolean
    
    On Local Error GoTo Erro
      
    Gravar = False
    
    If ClasseInicializada = False Then Exit Function
    
    VGCnx_BrasilBank.BeginTrans
    
    If aOperacao = eOperacao.Delete Then
        DataBase.Delete Condicao
    Else
        If ExisteRegistro Then
            DataBase.Update Me, Condicao
        Else
            DataBase.Insert Me
        End If
    End If
    
    GoTo Fim
Erro:
    VGCnx_BrasilBank.RollbackTrans
    Err.Raise vbObjectError + 501, Err.Source, "Erro ao gravar informações!"
    Exit Function
Fim:
    VGCnx_BrasilBank.CommitTrans
    Gravar = True
   
End Function
