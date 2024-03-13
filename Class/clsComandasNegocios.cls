VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComandasNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_ComandasDados As New clsComandasDados      'Variável de acesso classe clsComandasDados

'Variáveis da classe
Private vsp_Query As String                'Variável usada nas consultas sql
Private vip_Count As Integer               'Variável utilizadas nos loops
Private vbp_RollBack As Boolean            'Define que uma conexão está aberta e necessita de um rollback

'Variáveis das Propriedades
Private mvarIdComandas As Long                'Identificador da Comandas
Private mvarIdTipoComanda As Long             'Identificador do Tipo Comanda
Private mvarNumeroMesa As String              'NumeroMesa da Comanda
Private mvarQuantidadePessoa As Integer       'QuantidadePessoa da Comandas
Private mvarDataConfirmacaoPreparo As String
Private mvarDataPrevistaPreparo As String
Private mvarDataFinalizacaoPreparo As String
Private mvarStatusComanda As Integer


'Propriedades
'IdComandas
Public Property Let IdComandas(ByVal vData As Long)
    mvarIdComandas = vData
End Property

Public Property Get IdComandas() As Long
    IdComandas = mvarIdComandas
End Property

'IdComandas
Public Property Let IdTipoComanda(ByVal vData As Long)
    mvarIdTipoComanda = vData
End Property

Public Property Get IdTipoComanda() As Long
    IdTipoComanda = mvarIdTipoComanda
End Property

'NumeroMesa
Public Property Let NumeroMesa(ByVal vData As Integer)
    mvarNumeroMesa = vData
End Property

Public Property Get NumeroMesa() As Integer
    NumeroMesa = mvarNumeroMesa
End Property

'QuantidadePessoa
Public Property Let QuantidadePessoa(ByVal vData As Integer)
    mvarQuantidadePessoa = vData
End Property

Public Property Get QuantidadePessoa() As Integer
    QuantidadePessoa = mvarQuantidadePessoa
End Property

'DataConfirmacaoPreparo
Public Property Let DataConfirmacaoPreparo(ByVal vData As String)
    mvarDataConfirmacaoPreparo = vData
End Property

Public Property Get DataConfirmacaoPreparo() As String
    DataConfirmacaoPreparo = mvarDataConfirmacaoPreparo
End Property

'DataFinalizacaoPreparo
Public Property Let DataPrevistaPreparo(ByVal vData As String)
    mvarDataPrevistaPreparo = vData
End Property

Public Property Get DataPrevistaPreparo() As String
    DataPrevistaPreparo = mvarDataPrevistaPreparo
End Property

'DataFinalizacaoPreparo
Public Property Let DataFinalizacaoPreparo(ByVal vData As String)
    mvarDataFinalizacaoPreparo = vData
End Property

Public Property Get DataFinalizacaoPreparo() As String
    DataFinalizacaoPreparo = mvarDataFinalizacaoPreparo
End Property

'StatusComanda
Public Property Let StatusComanda(ByVal vData As Integer)
    mvarStatusComanda = vData
End Property

Public Property Get StatusComanda() As Integer
    StatusComanda = mvarStatusComanda
End Property


'Metodos
Public Function ImprimirEstatisticas(ByVal pcmbLocalizarItemEstatistica As Integer) As Boolean

    Screen.MousePointer = vbHourglass
      ImprimirEstatisticas = vop_ComandasDados.ImprimirEstatisticasDE(pcmbLocalizarItemEstatistica)
    Screen.MousePointer = vbDefault

End Function

Public Function CarregarGridComandaMainRS(ByRef padoComandas As Adodc, ByVal ptxtLocalizarComanda As String, ByVal pcmbLocalizarComanda As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridComandaMainRS = vop_ComandasDados.CarregarGridComandaMainRS(padoComandas, ptxtLocalizarComanda, pcmbLocalizarComanda)
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridComandaRS(ByRef padoComandas As Adodc, ByVal ptxtLocalizarComanda As String, ByVal pcmbLocalizarComanda As Integer, ByRef padoItensComanda As Adodc, ByVal ptxtLocalizarItem As String, ByVal pcmbLocalizarItem As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridComandaRS = vop_ComandasDados.CarregarGridComandaRS(padoComandas, ptxtLocalizarComanda, pcmbLocalizarComanda, padoItensComanda, ptxtLocalizarItem, pcmbLocalizarItem)
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridItensCozinhaRS(ByRef padoCozinhaItens As Adodc, ByVal ptxtLocalizarItemCozinha As String, ByVal pcmbLocalizarItemCozinha As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridItensCozinhaRS = vop_ComandasDados.CarregarGridItensCozinhaRS(padoCozinhaItens, ptxtLocalizarItemCozinha, pcmbLocalizarItemCozinha)
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridItensEstatisticaRS(ByRef padoEstatisticaItens As Adodc, ByRef padoEstatisticaItensDias As Adodc) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridItensEstatisticaRS = vop_ComandasDados.CarregarGridItensEstatisticaRS(padoEstatisticaItens, padoEstatisticaItensDias)
    Screen.MousePointer = vbDefault
    
End Function

Public Function FecharComanda() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      FecharComanda = vop_ComandasDados.FecharComanda(IdComandas, StatusComanda)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function IncluirComanda(ByVal plvwItensComanda As ListView, ByRef pIdComandas As Integer) As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      IncluirComanda = vop_ComandasDados.IncluirComanda(IdTipoComanda, NumeroMesa, QuantidadePessoa, StatusComanda, plvwItensComanda, pIdComandas)
      'Atualiza variavel de controle
      IdComandas = pIdComandas
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarComanda(ByVal plvwItensComanda As ListView, ByVal plvwIdItemComandaDeletado As ListView) As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      AlterarComanda = vop_ComandasDados.AlterarComanda(IdComandas, IdTipoComanda, NumeroMesa, QuantidadePessoa, StatusComanda, plvwItensComanda, plvwIdItemComandaDeletado)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function EnviarItensProcessamentoCozinha(ByVal plvwItensComanda As ListView) As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      EnviarItensProcessamentoCozinha = vop_ComandasDados.EnviarItensProcessamentoCozinha(IdTipoComanda, NumeroMesa, QuantidadePessoa, StatusComanda, plvwItensComanda)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function EnviarItensCozinha(ByRef plvwEnviarItensCozinha As ListView) As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      EnviarItensCozinha = vop_ComandasDados.AtualizarDataPrevistaPreparoItensCozinha(plvwEnviarItensCozinha)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function PesquisarComanda(ByRef plvwComandas As ListView, ByVal pComanda As String, ByVal pBusca As Integer) As Boolean

    PesquisarComanda = False
    Screen.MousePointer = vbHourglass
      If vop_ComandasDados.PesquisarComanda(plvwComandas, pComanda, pBusca) = True Then
         With plvwComandas
            IdComandas = Trim$(.ListItems(plvwComandas.ListItems.Count).text)
            IdTipoComanda = Trim$(.ListItems(plvwComandas.ListItems.Count).SubItems(1))
            NumeroMesa = Trim$(.ListItems(plvwComandas.ListItems.Count).SubItems(2))
            QuantidadePessoa = Trim$(.ListItems(plvwComandas.ListItems.Count).SubItems(3))
            DataConfirmacaoPreparo = Trim$(.ListItems(plvwComandas.ListItems.Count).SubItems(4))
            DataPrevistaPreparo = Trim$(.ListItems(plvwComandas.ListItems.Count).SubItems(5))
            DataFinalizacaoPreparo = Trim$(.ListItems(plvwComandas.ListItems.Count).SubItems(6))
            StatusComanda = Trim$(.ListItems(plvwComandas.ListItems.Count).SubItems(7))
         End With
         PesquisarComanda = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarComanda(ByRef padoComandas As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarComanda = False
    Screen.MousePointer = vbHourglass
      If vop_ComandasDados.LocalizarComanda(padoComandas, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarComanda = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function ExcluirComanda() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirComanda = vop_ComandasDados.ExcluirComanda(IdComandas)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function















