VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiposComandasNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_TiposComandasDados As New clsTiposComandasDados      'Variável de acesso classe clsTiposComandasDados

'Variáveis da classe
Private vsp_Query As String                'Variável usada nas consultas sql
Private vip_Count As Integer               'Variável utilizadas nos loops
Private vbp_RollBack As Boolean            'Define que uma conexão está aberta e necessita de um rollback

'Variáveis das Propriedades
Private mvarIdComanda As Long              'Identificador da Comandas
Private mvarDescricao As String            'Descrição da Comandas
Private mvarStatusComanda As Integer



'Propriedades
'IdComanda
Public Property Let IdComanda(ByVal vData As Long)
    mvarIdComanda = vData
End Property

Public Property Get IdComanda() As Long
    IdComanda = mvarIdComanda
End Property

'Descricao
Public Property Let Descricao(ByVal vData As String)
    mvarDescricao = vData
End Property

Public Property Get Descricao() As String
    Descricao = mvarDescricao
End Property

'StatusComanda
Public Property Let StatusComanda(ByVal vData As Integer)
    mvarStatusComanda = vData
End Property

Public Property Get StatusComanda() As Integer
    StatusComanda = mvarStatusComanda
End Property



'Metodos
Public Function PesquisarTipoComanda(ByRef plvwTiposComandas As ListView, ByVal pComanda As String, ByVal pBusca As Integer) As Boolean

    PesquisarTipoComanda = False
    Screen.MousePointer = vbHourglass
      If vop_TiposComandasDados.PesquisarTipoComanda(plvwTiposComandas, pComanda, pBusca) = True Then
         If plvwTiposComandas.ListItems.Count > 0 Then
            With plvwTiposComandas
               IdComanda = Trim$(.ListItems(plvwTiposComandas.ListItems.Count).text)
               Descricao = Trim$(.ListItems(plvwTiposComandas.ListItems.Count).SubItems(1))
               StatusComanda = Trim$(.ListItems(plvwTiposComandas.ListItems.Count).SubItems(2))
            End With
            PesquisarTipoComanda = True
         End If
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarTipoComanda(ByRef padoTiposComandas As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarTipoComanda = False
    Screen.MousePointer = vbHourglass
      If vop_TiposComandasDados.LocalizarTipoComanda(padoTiposComandas, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarTipoComanda = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridTipoComandaRS(ByRef padoTiposComandas As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridTipoComandaRS = vop_TiposComandasDados.CarregarGridTipoComandaRS(padoTiposComandas, pcmbLocalizar)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirTipoComanda() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      IncluirTipoComanda = vop_TiposComandasDados.IncluirTipoComanda(Descricao)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarTipoComanda() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      AlterarTipoComanda = vop_TiposComandasDados.AlterarTipoComanda(IdComanda, Descricao)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function ExcluirTipoComanda() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirTipoComanda = vop_TiposComandasDados.ExcluirTipoComanda(IdComanda)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function














