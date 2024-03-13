VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCozinhasNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_CozinhasDados As New clsCozinhasDados      'Variável de acesso classe clsCozinhaDados

'Variáveis da classe
Private vsp_Query As String                'Variável usada nas consultas sql
Private vip_Count As Integer               'Variável utilizadas nos loops
Private vbp_RollBack As Boolean            'Define que uma conexão está aberta e necessita de um rollback

'Variáveis das Propriedades
Private mvarIdCozinhas As Long             'Identificador da Cozinhas
Private mvarDescricao As String            'Descrição da Cozinhas
Private mvarCapacidade As Integer          'Capacidade da Cozinhas




'Propriedades
'IdCozinhas
Public Property Let IdCozinhas(ByVal vData As Long)
    mvarIdCozinhas = vData
End Property

Public Property Get IdCozinhas() As Long
    IdCozinhas = mvarIdCozinhas
End Property

'Descricao
Public Property Let Descricao(ByVal vData As String)
    mvarDescricao = vData
End Property

Public Property Get Descricao() As String
    Descricao = mvarDescricao
End Property

'Capacidade
Public Property Let Capacidade(ByVal vData As Integer)
    mvarCapacidade = vData
End Property

Public Property Get Capacidade() As Integer
    Capacidade = mvarCapacidade
End Property



'Metodos
Public Function PesquisarCozinha(ByRef plvwCozinhas As ListView, ByVal pCozinha As String, ByVal pBusca As Integer) As Boolean

    PesquisarCozinha = False
    Screen.MousePointer = vbHourglass
      If vop_CozinhasDados.PesquisarCozinha(plvwCozinhas, pCozinha, pBusca) = True Then
         With plvwCozinhas
            IdCozinhas = Trim$(.ListItems(plvwCozinhas.ListItems.Count).text)
            Descricao = Trim$(.ListItems(plvwCozinhas.ListItems.Count).SubItems(1))
            Capacidade = Trim$(.ListItems(plvwCozinhas.ListItems.Count).SubItems(2))
         End With
         PesquisarCozinha = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarCozinha(ByRef padoCozinhas As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarCozinha = False
    Screen.MousePointer = vbHourglass
      If vop_CozinhasDados.LocalizarCozinha(padoCozinhas, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarCozinha = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridCozinhaRS(ByRef padoCozinhas As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridCozinhaRS = vop_CozinhasDados.CarregarGridCozinhaRS(padoCozinhas, pcmbLocalizar)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirCozinha() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      IncluirCozinha = vop_CozinhasDados.IncluirCozinha(Descricao, Capacidade)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarCozinha() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      AlterarCozinha = vop_CozinhasDados.AlterarCozinha(IdCozinhas, Descricao, Capacidade)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function ExcluirCozinha() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirCozinha = vop_CozinhasDados.ExcluirCozinha(IdCozinhas)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function














