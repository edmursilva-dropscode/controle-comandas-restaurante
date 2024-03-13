VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardapioNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_CardapioDados As New clsCardapioDados      'Variável de acesso classe clsCardapioDados

'Variáveis das Propriedades
Private mvarIdCardapio As Long             'Identificador do Cardapio
Private mvarTempoPreparo As Integer        'TempoPreparo do Cardapio
Private mvarIdCozinha As Integer           'Identificador da Cozinha
Private mvarPreco As Double                'Preço do item do Cardapio
Private mvarDescricao As String            'Descrição do Cardapio
Private mvarDescricaoCozinha As String     'Descrição da Cozinha



'Propriedades
'IdCardapio
Public Property Let IdCardapio(ByVal vData As Long)
    mvarIdCardapio = vData
End Property

Public Property Get IdCardapio() As Long
    IdCardapio = mvarIdCardapio
End Property

'IdCozinha
Public Property Let IdCozinha(ByVal vData As Long)
    mvarIdCozinha = vData
End Property

Public Property Get IdCozinha() As Long
    IdCozinha = mvarIdCozinha
End Property

'TempoPreparo
Public Property Let TempoPreparo(ByVal vData As Integer)
    mvarTempoPreparo = vData
End Property

Public Property Get TempoPreparo() As Integer
    TempoPreparo = mvarTempoPreparo
End Property

'Preco
Public Property Let Preco(ByVal vData As Double)
    mvarPreco = vData
End Property

Public Property Get Preco() As Double
    Preco = mvarPreco
End Property

'Descricao
Public Property Let Descricao(ByVal vData As String)
    mvarDescricao = vData
End Property

Public Property Get Descricao() As String
    Descricao = mvarDescricao
End Property

'DescricaoCozinha
Public Property Let DescricaoCozinha(ByVal vData As String)
    mvarDescricaoCozinha = vData
End Property

Public Property Get DescricaoCozinha() As String
    DescricaoCozinha = mvarDescricaoCozinha
End Property


'Metodos
Public Function PesquisarCardapio(ByRef plvwCardapio As ListView, ByVal pIdCardapio As Integer) As Boolean

    PesquisarCardapio = False
    Screen.MousePointer = vbHourglass
      If vop_CardapioDados.PesquisarCardapio(plvwCardapio, pIdCardapio) = True Then
         If plvwCardapio.ListItems.Count > 0 Then
            With plvwCardapio
               IdCardapio = Trim$(.ListItems(plvwCardapio.ListItems.Count).text)
               IdCozinha = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(1))
               TempoPreparo = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(2))
               Preco = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(3))
               Descricao = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(4))
               DescricaoCozinha = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(5))
            End With
            PesquisarCardapio = True
         End If
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function PesquisarCardapioItemComanda(ByRef plvwCardapio As ListView, ByVal pItem As String, ByVal pBusca As Integer) As Boolean

    PesquisarCardapioItemComanda = False
    Screen.MousePointer = vbHourglass
      If vop_CardapioDados.PesquisarCardapioItemComanda(plvwCardapio, pItem, pBusca) = True Then
         If plvwCardapio.ListItems.Count > 0 Then
            With plvwCardapio
               IdCardapio = Trim$(.ListItems(plvwCardapio.ListItems.Count).text)
               IdCozinha = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(1))
               TempoPreparo = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(2))
               Preco = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(3))
               Descricao = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(4))
               DescricaoCozinha = Trim$(.ListItems(plvwCardapio.ListItems.Count).SubItems(5))
            End With
            PesquisarCardapioItemComanda = True
         End If
      End If
    Screen.MousePointer = vbDefault
    
End Function



Public Function LocalizarCardapio(ByRef padoCardapio As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarCardapio = False
    Screen.MousePointer = vbHourglass
      If vop_CardapioDados.LocalizarCardapio(padoCardapio, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarCardapio = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridCardapioRS(ByRef padoCardapio As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridCardapioRS = vop_CardapioDados.CarregarGridCardapioRS(padoCardapio, pcmbLocalizar)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirCardapio() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      IncluirCardapio = vop_CardapioDados.IncluirCardapio(IdCozinha, Descricao, TempoPreparo, Preco)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarCardapio() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      AlterarCardapio = vop_CardapioDados.AlterarCardapio(IdCardapio, IdCozinha, Descricao, TempoPreparo, Preco)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function ExcluirCardapio() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirCardapio = vop_CardapioDados.ExcluirCardapio(IdCardapio)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function















