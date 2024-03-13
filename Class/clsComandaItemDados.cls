VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComandaItemDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variáveis do ADO
Private vol_Conexao As New clsConexao
Private vol_Command As New ADODB.Command



Public Property Set Conexao(ByVal vData As ADODB.Connection)
    Set cConexao = vData
End Property


'Metodos
Public Function CancelarFinalizarItensComanda(ByVal pIdItemComanda As Integer, ByVal pStatusItem As Integer) As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         CancelarFinalizarItensComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_U_AtualizarCancelarFinalizarItensComanda"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdItemComanda
           .Item(2).Value = pStatusItem
         End With
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       CancelarFinalizarItensComanda = False
       MsgBox "Não foi possível cancelar/finalizar informações do Item da Comanda !", vbExclamation
    End If
End Function



Public Function IncluirItensComanda(ByRef plvwItensComanda As ListView, ByVal plvwItemComanda As ListView) As Boolean
Dim vil_CountLista As Integer

On Error GoTo TrataErros
     
   IncluirItensComanda = True
            
   For vil_CountLista = 1 To plvwItemComanda.ListItems.Count
       plvwItensComanda.ListItems.Add , , plvwItemComanda.ListItems(vil_CountLista).text
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(1) = plvwItemComanda.ListItems(vil_CountLista).SubItems(1)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(2) = plvwItemComanda.ListItems(vil_CountLista).SubItems(2)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(3) = plvwItemComanda.ListItems(vil_CountLista).SubItems(3)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(4) = plvwItemComanda.ListItems(vil_CountLista).SubItems(4)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(5) = plvwItemComanda.ListItems(vil_CountLista).SubItems(5)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(6) = plvwItemComanda.ListItems(vil_CountLista).SubItems(6)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(7) = plvwItemComanda.ListItems(vil_CountLista).SubItems(7)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(8) = plvwItemComanda.ListItems(vil_CountLista).SubItems(8)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(9) = plvwItemComanda.ListItems(vil_CountLista).SubItems(9)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(10) = plvwItemComanda.ListItems(vil_CountLista).SubItems(10)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(11) = plvwItemComanda.ListItems(vil_CountLista).SubItems(11)
       plvwItensComanda.ListItems(plvwItensComanda.ListItems.Count).SubItems(12) = plvwItemComanda.ListItems(vil_CountLista).SubItems(12)
   Next vil_CountLista
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
        IncluirItensComanda = False
        MsgBox "Não foi possível incluir informações do Item da Comanda !", vbExclamation
    End If
End Function

Public Function AlterarItensComanda(ByRef plvwItensComanda As ListView, ByVal plvwItemComanda As ListView, ByVal pItemListaComanda As Integer) As Boolean
Dim vil_CountLista As Integer

On Error GoTo TrataErros
     
   AlterarItensComanda = True
            
   For vil_CountLista = 1 To plvwItensComanda.ListItems.Count
       'If plvwItensComanda.ListItems(pItemListaComanda).SubItems(3) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(3) Then
       If vil_CountLista = pItemListaComanda Then
          plvwItensComanda.ListItems(pItemListaComanda).text = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).text
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(1) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(1)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(2) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(2)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(3) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(3)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(4) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(4)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(5) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(5)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(6) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(6)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(7) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(7)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(8) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(8)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(9) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(9)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(10) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(10)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(11) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(11)
          plvwItensComanda.ListItems(pItemListaComanda).SubItems(12) = plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(12)
       End If
   Next vil_CountLista
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
        AlterarItensComanda = False
        MsgBox "Não foi possível incluir informações do Item da Comanda !", vbExclamation
    End If
End Function

Public Function PesquisarItemComanda(ByRef plvwItemComanda As ListView, ByVal pIdcomanda As Integer, ByVal pIdCardapio As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   PesquisarItemComanda = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, IdComanda, IdTipoComanda, IdCardapio, DescricaoItem, Quantidade, " _
                                              & "FORMAT(Preco, '#,0.00', 'pt-br') AS Preco, FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                              & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                              & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                              & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                              & "StatusItem, " _
                                              & "CASE " _
                                              & "    WHEN StatusItem = 1 THEN 'Item aguardando envio' " _
                                              & "    WHEN StatusItem = 2 THEN 'Item aguardando processamento' " _
                                              & "    WHEN StatusItem = 3 THEN 'Item sendo prepadada' " _
                                              & "    WHEN StatusItem = 4 THEN 'Item para entrega' " _
                                              & "    WHEN StatusItem = 5 THEN 'Item cancelada' " _
                                              & "END StatusItemDescricao " _
                                              & "FROM ComandasItem (NOLOCK) " _
                                              & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                              & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID " _
                                              & "" + IIf(pIdCardapio > 0, "WHERE IdComanda = '" + CStr(pIdcomanda) + "' AND IdCardapio = '" + CStr(pIdCardapio) + "' ", "WHERE IdComanda = '" + CStr(pIdcomanda) + "' ") + " " _
                                              & "ORDER BY Id ")
                                              
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwItemComanda.ListItems.Add , , .Fields(0).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(2) = .Fields(2).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(3) = .Fields(3).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(4) = .Fields(4).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(5) = .Fields(5).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(6) = .Fields(6).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(7) = .Fields(7).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(8) = IIf(IsNull(.Fields(8).Value), "????", .Fields(8).Value)
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(9) = IIf(IsNull(.Fields(9).Value), "????", .Fields(9).Value)
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(10) = IIf(IsNull(.Fields(10).Value), "????", .Fields(10).Value)
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(11) = .Fields(11).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(12) = IIf(IsNull(.Fields(12).Value), "????", .Fields(12).Value)
               .MoveNext
            Loop
         End If
      End With
      Set vol_Recordset = Nothing
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        PesquisarItemComanda = False
        MsgBox "Não foi possível encontrar informações do Item da Comanda !", vbExclamation
    End If
End Function

Public Function CancelarItemComanda(ByRef plvwItemComanda As ListView, ByVal pIdItem As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CancelarItemComanda = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, IdComanda, IdTipoComanda, IdCardapio, DescricaoItem, Quantidade, " _
                                              & "FORMAT(Preco, '#,0.00', 'pt-br') AS Preco, FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                              & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                              & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                              & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                              & "StatusItem, " _
                                              & "CASE " _
                                              & "    WHEN StatusItem = 1 THEN 'Item aguardando envio' " _
                                              & "    WHEN StatusItem = 2 THEN 'Item aguardando processamento' " _
                                              & "    WHEN StatusItem = 3 THEN 'Item sendo prepadada' " _
                                              & "    WHEN StatusItem = 4 THEN 'Item para entrega' " _
                                              & "    WHEN StatusItem = 5 THEN 'Item cancelada' " _
                                              & "END StatusItemDescricao " _
                                              & "FROM ComandasItem (NOLOCK) " _
                                              & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                              & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID " _
                                              & "WHERE Id = '" + CStr(pIdItem) + "' " _
                                              & "ORDER BY Id ")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwItemComanda.ListItems.Add , , .Fields(0).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(2) = .Fields(2).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(3) = .Fields(3).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(4) = .Fields(4).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(5) = .Fields(5).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(6) = .Fields(6).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(7) = .Fields(7).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(8) = IIf(IsNull(.Fields(8).Value), "????", .Fields(8).Value)
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(9) = IIf(IsNull(.Fields(9).Value), "????", .Fields(9).Value)
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(10) = IIf(IsNull(.Fields(10).Value), "????", .Fields(10).Value)
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(11) = .Fields(11).Value
               plvwItemComanda.ListItems(plvwItemComanda.ListItems.Count).SubItems(12) = IIf(IsNull(.Fields(12).Value), "????", .Fields(12).Value)
               .MoveNext
            Loop
         End If
      End With
      Set vol_Recordset = Nothing
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CancelarItemComanda = False
        MsgBox "Não foi possível encontrar informações do Item da Comanda !", vbExclamation
    End If
End Function

Public Function LocalizarItem(ByRef padoItensComanda As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarItem = True
   vol_Conexao.AbrirConexao
         padoItensComanda.RecordSource = "SELECT Id, IdComanda, IdTipoComanda, IdCardapio, DescricaoItem, Quantidade, " _
                                       & "FORMAT(Preco, '#,0.00', 'pt-br') AS Preco, FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                       & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                       & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                       & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                       & "StatusItem, " _
                                       & "StatusItemDescricao " _
                                       & "FROM ComandasItem (NOLOCK) " _
                                       & "INNER JOIN (SELECT Id AS ComandaIdItem, IIF(StatusItem=1,'Item aguardando envio',IIF(StatusItem=2,'Item aguardando processamento',IIF(StatusItem=3,'Item sendo prepadada',IIF(StatusItem=4,'Item para entrega','Item cancelada')))) As StatusItemDescricao FROM ComandasItem (NOLOCK)) As ComandasItemTipo ON ComandasItem.Id = ComandasItemTipo.ComandaIdItem " _
                                       & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                       & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID " _
                                       & "WHERE " + IIf(pcmbLocalizar = 0, "IdTipoComanda", IIf(pcmbLocalizar = 1, "DescricaoItem", "StatusItemDescricao")) + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "IdTipoComanda", IIf(pcmbLocalizar = 1, "DescricaoItem", "StatusItemDescricao")) + " "
      padoItensComanda.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarItem = False
        MsgBox "Não foi possível encontrar informações dos Itens das Comandas !", vbExclamation
    End If
End Function

Public Function LocalizarItemCozinha(ByRef padoCozinhaItens As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarItemCozinha = True
   vol_Conexao.AbrirConexao
      padoCozinhaItens.RecordSource = "SELECT IdCozinha, DescricaoCozinha, IdCardapio, DescricaoItem, Quantidade, Id, IdComanda, IdTipoComanda, " _
                                & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                & "FORMAT(Preco, '#,0.00', 'pt-br') AS Preco, FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                & "StatusItem, " _
                                & "StatusItemDescricao " _
                                & "FROM ComandasItem (NOLOCK) " _
                                & "INNER JOIN (SELECT Id AS ComandaId, IIF(StatusItem=1,'Item aguardando envio',IIF(StatusItem=2,'Item aguardando processamento',IIF(StatusItem=3,'Item sendo prepadada',IIF(StatusItem=4,'Item para entrega','Item cancelada')))) As StatusItemDescricao FROM ComandasItem (NOLOCK)) As ComandasItemStatus ON ComandasItem.Id = ComandasItemStatus.ComandaId " _
                                & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem, IdCozinha FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID " _
                                & "INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha, Capacidade FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaId " _
                                & "WHERE StatusItem IN(3, 4) " _
                                & "AND " + IIf(pcmbLocalizar = 0, "DescricaoCozinha", IIf(pcmbLocalizar = 1, "DescricaoItem", "StatusItemDescricao")) + " LIKE '%" + Trim$(ptxtLocalizar) + "%' " _
                                & "ORDER BY " + IIf(pcmbLocalizar = 0, "DescricaoCozinha, DataPrevistaPreparo DESC", IIf(pcmbLocalizar = 1, "DescricaoItem, DataPrevistaPreparo DESC", "StatusItemDescricao, DataPrevistaPreparo DESC")) + " "
                                '& "ORDER BY DescricaoCozinha, DataPrevistaPreparo "
         
      padoCozinhaItens.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarItemCozinha = False
        MsgBox "Não foi possível encontrar informações dos Itens das Comandas !", vbExclamation
    End If
End Function

Public Function LocalizarItemEstatistica(ByRef padoEstatisticaItens As Adodc, ByRef padoEstatisticaItensDias As Adodc) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarItemEstatistica = True
   vol_Conexao.AbrirConexao
      padoEstatisticaItens.RecordSource = "SELECT IdCozinha, DescricaoCozinha, IdCardapio, DescricaoItem, SUM(Quantidade) AS QUANTIDADE  " _
                                & "FROM ComandasItem (NOLOCK) " _
                                & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK) WHERE StatusComanda = 2) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId  " _
                                & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem, IdCozinha FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioId " _
                                & "INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha, Capacidade FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaId " _
                                & "WHERE StatusItem = 4 " _
                                & "AND ( DataFinalizacaoPreparo > DATEADD(day,-30,GETDATE()) AND DataFinalizacaoPreparo < DATEADD(day,1,GETDATE()) ) " _
                                & "GROUP BY IdCozinha, DescricaoCozinha, IdCardapio, DescricaoItem "
         
      padoEstatisticaItens.Refresh
      
      padoEstatisticaItensDias.RecordSource = "WITH DiaSemana AS ( " _
                                & "SELECT 1 AS idDSe, DATENAME(WEEKDAY, -1) AS nomeDSe " _
                                & "UNION ALL  " _
                                & "SELECT idDSe + 1, DATENAME(WEEKDAY, idDSe -1) " _
                                & "FROM DiaSemana " _
                                & " WHERE idDSe < 7 ) " _
                                & "SELECT idDSe, nomeDSe, ISNULL(QuantidadePessoa,0) as QuantidadePessoas FROM DiaSemana " _
                                & "LEFT JOIN (SELECT DiaSemana, SUM(QuantidadePessoa) AS QuantidadePessoa FROM Comandas (NOLOCK) INNER JOIN (SELECT Id As ComandaId, DATEPART(DW,DataFinalizacaoPreparo) as DiaSemana FROM Comandas (NOLOCK)) As ComandasData ON DiaSemana = ComandasData.ComandaId WHERE StatusComanda = 2 AND DataFinalizacaoPreparo IS NOT NULL AND ( DataFinalizacaoPreparo > DATEADD(day,-60,GETDATE()) AND DataFinalizacaoPreparo < DATEADD(day,-(DATEPART(DW,GETDATE())),GETDATE()) ) GROUP BY DiaSemana) As ComandaDias ON DiaSemana.idDSe = ComandaDias.DiaSemana "
      padoEstatisticaItensDias.Refresh
      
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarItemEstatistica = False
        MsgBox "Não foi possível encontrar informações dos Itens das Comandas !", vbExclamation
    End If
End Function



