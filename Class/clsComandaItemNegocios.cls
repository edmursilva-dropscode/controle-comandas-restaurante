VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComandaItemNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_ItemComandaDados As New clsComandaItemDados      'Variável de acesso classe clsComandaItemDados

'Variáveis das Propriedades
Private mvarIdItemComanda As Long          'Identificador do Item da Comanda
Private mvarIdComanda As Integer           'Identificador da Comanda
Private mvarIdTipoComanda As Integer       'Identificador do tipo de Comanda
Private mvarIdCardapio As Integer          'Identificador do Cardapio
Private mvarDescricaoItem As String        'Descrição do item da Comanda
Private mvarQuantidade As Integer          'Quantidade do Item da Comanda
Private mvarPreco As Double                'Preço do item da Comanda
Private mvarTotalPreco As Double           'Total do item da Comanda
Private mvarDataHora As String             'Data e Hora do Item da Comanda
Private mvarDataHoraPrevista As String     'Data e Hora prevista de preparo do Item da Comanda
Private mvarDataHoraFinalizacao As String  'Data e Hora finalizacao do preparo do Item da Comanda
Private mvarStatusItem As Integer          'Status do Item da Comanda
Private mvarStatusItemDescricao As String  'Status do Item da Comanda




'Propriedades
'IdItemComanda
Public Property Let IdItemComanda(ByVal vData As Long)
    mvarIdItemComanda = vData
End Property

Public Property Get IdItemComanda() As Long
    IdItemComanda = mvarIdItemComanda
End Property

'IdComanda
Public Property Let IdComanda(ByVal vData As Long)
    mvarIdComanda = vData
End Property

Public Property Get IdComanda() As Long
    IdComanda = mvarIdComanda
End Property

'IdTipoComanda
Public Property Let IdTipoComanda(ByVal vData As Long)
    mvarIdTipoComanda = vData
End Property

Public Property Get IdTipoComanda() As Long
    IdTipoComanda = mvarIdTipoComanda
End Property

'IdCardapio
Public Property Let IdCardapio(ByVal vData As Long)
    mvarIdCardapio = vData
End Property

Public Property Get IdCardapio() As Long
    IdCardapio = mvarIdCardapio
End Property

'DescricaoItem
Public Property Let DescricaoItem(ByVal vData As String)
    mvarDescricaoItem = vData
End Property

Public Property Get DescricaoItem() As String
    DescricaoItem = mvarDescricaoItem
End Property

'Quantidade
Public Property Let Quantidade(ByVal vData As Integer)
    mvarQuantidade = vData
End Property

Public Property Get Quantidade() As Integer
    Quantidade = mvarQuantidade
End Property

'Preco
Public Property Let Preco(ByVal vData As Double)
    mvarPreco = vData
End Property

Public Property Get Preco() As Double
    Preco = mvarPreco
End Property

'TotalPreco
Public Property Let TotalPreco(ByVal vData As Double)
    mvarTotalPreco = vData
End Property

Public Property Get TotalPreco() As Double
    TotalPreco = mvarTotalPreco
End Property

'DataHora
Public Property Let DataHora(ByVal vData As String)
    mvarDataHora = vData
End Property

Public Property Get DataHora() As String
    DataHora = mvarDataHora
End Property

'DataHoraPrevista
Public Property Let DataHoraPrevista(ByVal vData As String)
    mvarDataHoraPrevista = vData
End Property

Public Property Get DataHoraPrevista() As String
    DataHoraPrevista = mvarDataHoraPrevista
End Property

'DataHoraFinalizacao
Public Property Let DataHoraFinalizacao(ByVal vData As String)
    mvarDataHoraFinalizacao = vData
End Property

Public Property Get DataHoraFinalizacao() As String
    DataHoraFinalizacao = mvarDataHoraFinalizacao
End Property

'StatusItem
Public Property Let StatusItem(ByVal vData As Long)
    mvarStatusItem = vData
End Property

Public Property Get StatusItem() As Long
    StatusItem = mvarStatusItem
End Property

'StatusItemDescricao
Public Property Let StatusItemDescricao(ByVal vData As String)
    mvarStatusItemDescricao = vData
End Property

Public Property Get StatusItemDescricao() As String
    StatusItemDescricao = mvarStatusItemDescricao
End Property

'Metodos
Public Function CancelarFinalizarItensComanda() As Boolean
     
    Screen.MousePointer = vbHourglass
      CancelarFinalizarItensComanda = vop_ItemComandaDados.CancelarFinalizarItensComanda(IdItemComanda, StatusItem)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirItensComanda(ByRef plvwItensComanda As ListView, ByVal plvwItemComanda As ListView) As Boolean
     
    Screen.MousePointer = vbHourglass
      IncluirItensComanda = vop_ItemComandaDados.IncluirItensComanda(plvwItensComanda, plvwItemComanda)
    Screen.MousePointer = vbDefault
    
End Function

Public Function AlterarItensComanda(ByRef plvwItensComanda As ListView, ByVal plvwItemComanda As ListView, ByVal pItemListaComanda As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      AlterarItensComanda = vop_ItemComandaDados.AlterarItensComanda(plvwItensComanda, plvwItemComanda, pItemListaComanda)
    Screen.MousePointer = vbDefault
    
End Function

Public Function PesquisarItemComanda(ByRef plvwItemComanda As ListView, ByVal pIdcomanda As Integer, ByVal pIdCardapio As Integer) As Boolean

    PesquisarItemComanda = False
    Screen.MousePointer = vbHourglass
      If vop_ItemComandaDados.PesquisarItemComanda(plvwItemComanda, pIdcomanda, pIdCardapio) = True Then
         If plvwItemComanda.ListItems.Count > 0 Then
            With plvwItemComanda
               IdItemComanda = Trim$(.ListItems(plvwItemComanda.ListItems.Count).text)
               IdComanda = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(1))
               IdTipoComanda = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(2))
               IdCardapio = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(3))
               DescricaoItem = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(4))
               Quantidade = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(5))
               Preco = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(6))
               TotalPreco = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(7))
               DataHora = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(8))
               DataHoraPrevista = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(9))
               DataHoraFinalizacao = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(10))
               StatusItem = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(11))
               StatusItemDescricao = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(12))
            End With
         End If
         PesquisarItemComanda = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CancelarItemComanda(ByRef plvwItemComanda As ListView, ByVal pIdItem As Integer) As Boolean

    CancelarItemComanda = False
    Screen.MousePointer = vbHourglass
      If vop_ItemComandaDados.CancelarItemComanda(plvwItemComanda, pIdItem) = True Then
         If plvwItemComanda.ListItems.Count > 0 Then
            With plvwItemComanda
               IdItemComanda = Trim$(.ListItems(plvwItemComanda.ListItems.Count).text)
               IdComanda = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(1))
               IdTipoComanda = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(2))
               IdCardapio = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(3))
               DescricaoItem = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(4))
               Quantidade = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(5))
               Preco = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(6))
               TotalPreco = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(7))
               DataHora = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(8))
               DataHoraPrevista = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(9))
               DataHoraFinalizacao = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(10))
               StatusItem = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(11))
               StatusItemDescricao = Trim$(.ListItems(plvwItemComanda.ListItems.Count).SubItems(12))
            End With
         End If
         CancelarItemComanda = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarItem(ByRef padoItensComanda As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarItem = False
    Screen.MousePointer = vbHourglass
      If vop_ItemComandaDados.LocalizarItem(padoItensComanda, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarItem = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarItemCozinha(ByRef padoCozinhaItens As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarItemCozinha = False
    Screen.MousePointer = vbHourglass
      If vop_ItemComandaDados.LocalizarItemCozinha(padoCozinhaItens, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarItemCozinha = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarItemEstatistica(ByRef padoEstatisticaItens As Adodc, ByRef padoEstatisticaItensDias As Adodc) As Boolean

    LocalizarItemEstatistica = False
    Screen.MousePointer = vbHourglass
      If vop_ItemComandaDados.LocalizarItemEstatistica(padoEstatisticaItens, padoEstatisticaItensDias) = True Then
         LocalizarItemEstatistica = True
      End If
    Screen.MousePointer = vbDefault
    
End Function












