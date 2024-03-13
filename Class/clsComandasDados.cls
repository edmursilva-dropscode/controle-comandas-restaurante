VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComandasDados"
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
Public Function ImprimirEstatisticasDE(ByVal pcmbLocalizarItemEstatistica As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   ImprimirEstatisticasDE = True
   vol_Conexao.AbrirConexao
      'Estatisticas
      If pcmbLocalizarItemEstatistica = 0 Then
         If DataEnvironment1.rsCommand1.State = adStateOpen Then
            DataEnvironment1.rsCommand1.Close
         End If
         DataEnvironment1.rsCommand1.Open "SELECT IdCozinha, DescricaoCozinha, IdCardapio, DescricaoItem, SUM(Quantidade) AS QUANTIDADE  " _
                                        + "FROM ComandasItem(NOLOCK) " _
                                        + "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK) WHERE StatusComanda = 2) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                        + "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem, IdCozinha FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioId " _
                                        + "INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha, Capacidade FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaId " _
                                        + "WHERE StatusItem = 4 " _
                                        + "AND ( DataFinalizacaoPreparo > DATEADD(day,-30,GETDATE()) AND DataFinalizacaoPreparo < DATEADD(day,1,GETDATE()) ) " _
                                        + "GROUP BY IdCozinha, DescricaoCozinha, IdCardapio, DescricaoItem "
      End If
      vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        ImprimirEstatisticasDE = False
        MsgBox "Não foi possível carregar informações da Estatística !", vbExclamation
    End If
End Function

Public Function CarregarGridComandaMainRS(ByRef padoComandas As Adodc, ByVal ptxtLocalizarComanda As String, ByVal pcmbLocalizarComanda As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CarregarGridComandaMainRS = True
   vol_Conexao.AbrirConexao
      'Comandas
      padoComandas.RecordSource = "SELECT Id, IdTipoComanda, DescricaoComanda, NumeroMesa, QuantidadePessoa, " _
                                & "FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                & "StatusComanda, " _
                                & "StatusComandaDescricao " _
                                & "FROM Comandas (NOLOCK) " _
                                & "INNER JOIN (SELECT Id AS ComandaId, IIF(StatusComanda=1,'Comanda aberta',IIF(StatusComanda=2,'Comanda fechada','Comanda cancelada')) As StatusComandaDescricao FROM Comandas (NOLOCK)) As ComandasTipo ON Comandas.Id = ComandasTipo.ComandaId " _
                                & "INNER JOIN (SELECT Id AS IdTiposComanda, Descricao AS DescricaoComanda FROM TiposComandas (NOLOCK)) As TiposComandas ON TiposComandas.IdTiposComanda = Comandas.IdTipoComanda " _
                                & "INNER JOIN (SELECT IdComanda, SUM(TotalPreco) AS TotalPreco FROM ComandasItem (NOLOCK) GROUP BY IdComanda) As ComandasItem ON ComandasItem.IdComanda = Comandas.Id " _
                                & "WHERE " + IIf(pcmbLocalizarComanda = 0, "IdTipoComanda", IIf(pcmbLocalizarComanda = 1, "NumeroMesa", "StatusComandaDescricao")) + " LIKE '%" + Trim$(ptxtLocalizarComanda) + "%' ORDER BY " + IIf(pcmbLocalizarComanda = 0, "IdTipoComanda", IIf(pcmbLocalizarComanda = 1, "NumeroMesa", "StatusComandaDescricao")) + " "
                                '& "ORDER BY IdTipoComanda"
      padoComandas.Refresh
      vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridComandaMainRS = False
        MsgBox "Não foi possível carregar informações da Comanda !", vbExclamation
    End If
End Function

Public Function CarregarGridComandaRS(ByRef padoComandas As Adodc, ByVal ptxtLocalizarComanda As String, ByVal pcmbLocalizarComanda As Integer, ByRef padoItensComanda As Adodc, ByVal ptxtLocalizarItem As String, ByVal pcmbLocalizarItem As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CarregarGridComandaRS = True
   vol_Conexao.AbrirConexao
      'Comandas
      padoComandas.RecordSource = "SELECT Id, IdTipoComanda, DescricaoComanda, NumeroMesa, QuantidadePessoa, " _
                                & "FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                & "StatusComanda, " _
                                & "StatusComandaDescricao " _
                                & "FROM Comandas (NOLOCK) " _
                                & "INNER JOIN (SELECT Id AS ComandaId, IIF(StatusComanda=1,'Comanda aberta',IIF(StatusComanda=2,'Comanda fechada','Comanda cancelada')) As StatusComandaDescricao FROM Comandas (NOLOCK)) As ComandasTipo ON Comandas.Id = ComandasTipo.ComandaId " _
                                & "INNER JOIN (SELECT Id AS IdTiposComanda, Descricao AS DescricaoComanda FROM TiposComandas (NOLOCK)) As TiposComandas ON TiposComandas.IdTiposComanda = Comandas.IdTipoComanda " _
                                & "INNER JOIN (SELECT IdComanda, SUM(TotalPreco) AS TotalPreco FROM ComandasItem (NOLOCK) GROUP BY IdComanda) As ComandasItem ON ComandasItem.IdComanda = Comandas.Id " _
                                & "WHERE " + IIf(pcmbLocalizarComanda = 0, "IdTipoComanda", IIf(pcmbLocalizarComanda = 1, "NumeroMesa", "StatusComandaDescricao")) + " LIKE '%" + Trim$(ptxtLocalizarComanda) + "%' ORDER BY " + IIf(pcmbLocalizarComanda = 0, "IdTipoComanda", IIf(pcmbLocalizarComanda = 1, "NumeroMesa", "StatusComandaDescricao")) + " "
                                '& "ORDER BY IdTipoComanda"
      padoComandas.Refresh
      'ItensComanda
      If IsNull(padoItensComanda) Then
         padoItensComanda.RecordSource = "SELECT Id, IdComanda, IdTipoComanda, IdCardapio, DescricaoItem, Quantidade, " _
                                   & "FORMAT(Preco, '#,0.00', 'pt-br') AS Preco, FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                   & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                   & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                   & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                   & "StatusItem, " _
                                   & "StatusItemDescricao " _
                                   & "FROM ComandasItem (NOLOCK) " _
                                   & "INNER JOIN (SELECT Id AS ComandaId, IIF(StatusItem=1,'Item aguardando envio',IIF(StatusItem=2,'Item aguardando processamento',IIF(StatusItem=3,'Item sendo prepadada',IIF(StatusItem=4,'Item para entrega','Item cancelada')))) As StatusItemDescricao FROM ComandasItem (NOLOCK)) As ComandasItemStatus ON ComandasItem.Id = ComandasItemStatus.ComandaId " _
                                   & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                   & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID " _
                                   & "WHERE " + IIf(pcmbLocalizarItem = 0, "IdTipoComanda", IIf(pcmbLocalizarItem = 1, "DescricaoItem", "StatusItemDescricao")) + " LIKE '%" + Trim$(ptxtLocalizarItem) + "%' ORDER BY " + IIf(pcmbLocalizarItem = 0, "IdTipoComanda", IIf(pcmbLocalizarItem = 1, "DescricaoItem", "StatusItemDescricao")) + " "
                                   '& "ORDER BY IdTipoComanda"
         padoItensComanda.Refresh
      End If
      vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridComandaRS = False
        MsgBox "Não foi possível carregar informações da Comanda !", vbExclamation
    End If
End Function

Public Function CarregarGridItensCozinhaRS(ByRef padoCozinhaItens As Adodc, ByVal ptxtLocalizarItemCozinha As String, ByVal pcmbLocalizarItemCozinha As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CarregarGridItensCozinhaRS = True
   vol_Conexao.AbrirConexao
      'ItensComanda
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
                                & "AND " + IIf(pcmbLocalizarItemCozinha = 0, "DescricaoCozinha", IIf(pcmbLocalizarItemCozinha = 1, "DescricaoItem", "StatusItemDescricao")) + " LIKE '%" + Trim$(ptxtLocalizarItemCozinha) + "%' " _
                                & "ORDER BY " + IIf(pcmbLocalizarItemCozinha = 0, "DescricaoCozinha, DataPrevistaPreparo", IIf(pcmbLocalizarItemCozinha = 1, "DescricaoItem", "StatusItemDescricao")) + " "
                                '& "ORDER BY DescricaoCozinha, DataPrevistaPreparo "
      padoCozinhaItens.Refresh
      vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridItensCozinhaRS = False
        MsgBox "Não foi possível carregar informações da Comanda !", vbExclamation
    End If
End Function

Public Function CarregarGridItensEstatisticaRS(ByRef padoEstatisticaItens As Adodc, ByRef padoEstatisticaItensDias As Adodc) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CarregarGridItensEstatisticaRS = True
   vol_Conexao.AbrirConexao
      'ItensComanda
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
        CarregarGridItensEstatisticaRS = False
        MsgBox "Não foi possível carregar informações dos Itens !", vbExclamation
    End If
End Function

Public Function FecharComanda(ByVal pIdcomanda As Integer, _
                                ByVal pStatusComanda As Integer) As Boolean
Dim vil_IdComanda As Integer
Dim vil_CountLista As Integer

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         FecharComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         'Comanda
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_U_ComandaFechar"
         vol_Command.Parameters.Refresh
         
         With vol_Command.Parameters
           .Item(1).Value = pIdcomanda
           .Item(2).Value = pStatusComanda
         End With
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       FecharComanda = False
       MsgBox "Não foi possível fechar informações da Comanda !", vbExclamation
    End If
End Function

Public Function IncluirComanda(ByVal pIdTipoComanda As Integer, _
                                ByVal pNumeroMesa As Integer, _
                                ByVal pQuantidadePessoa As Integer, _
                                ByVal pStatusComanda As Integer, _
                                ByVal plvwItensComanda As ListView, ByRef pIdcomanda As Integer) As Boolean
Dim vil_IdComanda As Integer
Dim vil_CountLista As Integer

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         'Comanda
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_I_Comanda"
         vol_Command.Parameters.Refresh
         
         With vol_Command.Parameters
           .Item(1).Value = pIdTipoComanda
           .Item(2).Value = pNumeroMesa
           .Item(3).Value = pQuantidadePessoa
           .Item(4).Value = pStatusComanda
         End With
         vol_Command.Execute
         
         'vil_IdComanda = vol_Command.Parameters("@IdComanda").Value
         pIdcomanda = vol_Command.Parameters("@IdComanda").Value
         
         'Itens da comanda
         For vil_CountLista = 1 To plvwItensComanda.ListItems.Count
         
             vol_Command.CommandType = adCmdStoredProc
             vol_Command.CommandText = "SP_Teste_I_ComandaItens"
             vol_Command.Parameters.Refresh
            
             With vol_Command.Parameters
               .Item(1).Value = pIdcomanda   'vil_IdComanda
               .Item(2).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(3)
               .Item(3).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(5)
               .Item(4).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(6)
               .Item(5).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(7)
               .Item(6).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(11)
             End With
             vol_Command.Execute
             
         Next vil_CountLista
         
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirComanda = False
       MsgBox "Não foi possível incluir informações da Comanda !", vbExclamation
    End If
End Function

Public Function AlterarComanda(ByVal pIdcomanda As Integer, _
                                ByVal pIdTipoComanda As Integer, _
                                ByVal pNumeroMesa As Integer, _
                                ByVal pQuantidadePessoa As Integer, _
                                ByVal pStatusComanda As Integer, _
                                ByVal plvwItensComanda As ListView, _
                                ByVal plvwIdItemComandaDeletado As ListView) As Boolean
Dim vil_CountLista As Integer

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         AlterarComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         'Comanda
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_U_Comanda"
         vol_Command.Parameters.Refresh
         
         With vol_Command.Parameters
           .Item(1).Value = pIdcomanda
           .Item(2).Value = pIdTipoComanda
           .Item(3).Value = pNumeroMesa
           .Item(4).Value = pQuantidadePessoa
           .Item(5).Value = pStatusComanda
         End With
         vol_Command.Execute
      
         'Deletar Itens da comanda
         For vil_CountLista = 1 To plvwIdItemComandaDeletado.ListItems.Count
         
             vol_Command.CommandType = adCmdStoredProc
             vol_Command.CommandText = "SP_Teste_D_ComandaItens"
             vol_Command.Parameters.Refresh
         
             With vol_Command.Parameters
               .Item(1).Value = Int(plvwIdItemComandaDeletado.ListItems(vil_CountLista).text)
             End With
             vol_Command.Execute
         
         Next vil_CountLista
         
         'Incluir Itens da comanda novos
         For vil_CountLista = 1 To plvwItensComanda.ListItems.Count
         
             If Int(plvwItensComanda.ListItems(vil_CountLista).text) = 0 Then
                vol_Command.CommandType = adCmdStoredProc
                vol_Command.CommandText = "SP_Teste_I_ComandaItens"
                vol_Command.Parameters.Refresh
            
                With vol_Command.Parameters
                  .Item(1).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(1)
                  .Item(2).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(3)
                  .Item(3).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(5)
                  .Item(4).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(6)
                  .Item(5).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(7)
                  .Item(6).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(11)
                End With
                vol_Command.Execute
             End If
         
         Next vil_CountLista
         
         'Alterar Itens da comanda
         For vil_CountLista = 1 To plvwItensComanda.ListItems.Count
         
             If Int(plvwItensComanda.ListItems(vil_CountLista).text) <> 0 Then
                vol_Command.CommandType = adCmdStoredProc
                vol_Command.CommandText = "SP_Teste_U_ComandaItens"
                vol_Command.Parameters.Refresh
            
                With vol_Command.Parameters
                  .Item(1).Value = plvwItensComanda.ListItems(vil_CountLista).text
                  .Item(2).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(1)
                  .Item(3).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(3)
                  .Item(4).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(5)
                  .Item(5).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(6)
                  .Item(6).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(7)
                  .Item(7).Value = CVDate(plvwItensComanda.ListItems(vil_CountLista).SubItems(8))
                  .Item(8).Value = plvwItensComanda.ListItems(vil_CountLista).SubItems(11)
                End With
                vol_Command.Execute
             End If
         
         Next vil_CountLista
         
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarComanda = False
       MsgBox "Não foi possível alterar informações da Comanda !", vbExclamation
    End If
End Function

Public Function EnviarItensProcessamentoCozinha(ByVal pIdTipoComanda As Integer, _
                                ByVal pNumeroMesa As Integer, _
                                ByVal pQuantidadePessoa As Integer, _
                                ByVal pStatusComanda As Integer, _
                                ByVal plvwItensComanda As ListView) As Boolean
Dim vil_IdComanda As Integer
Dim vil_CountLista As Integer

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         EnviarItensProcessamentoCozinha = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
           
         'Alterar Itens da comanda
         For vil_CountLista = 1 To plvwItensComanda.ListItems.Count
         
             If Int(plvwItensComanda.ListItems(vil_CountLista).text) <> 0 Then
                If Int(plvwItensComanda.ListItems(vil_CountLista).SubItems(11)) = 1 Then
                   vol_Command.CommandType = adCmdStoredProc
                   vol_Command.CommandText = "SP_Teste_U_EnviarItensProcessamentoCozinha"
                   vol_Command.Parameters.Refresh
               
                   With vol_Command.Parameters
                     .Item(1).Value = plvwItensComanda.ListItems(vil_CountLista).text
                     .Item(2).Value = 2
                   End With
                   vol_Command.Execute
                End If
             End If
         
         Next vil_CountLista
         
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       EnviarItensProcessamentoCozinha = False
       MsgBox "Não foi possível enviar item da Comanda para processamento na cozinha !", vbExclamation
    End If
End Function

'EnviarItensCozinha
Public Function AtualizarDataPrevistaPreparoItensCozinha(ByRef plvwEnviarItensCozinha As ListView) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim vdl_DataHoaAtual As String
Dim vdl_DataPrevistaPreparItemCozinha As String
Dim vdl_DataPrevistaPreparComanda As String
Dim vil_IdComanda As Integer
Dim vil_CountLista As Integer
Dim vbl_AtualizaDataPrevistaPreparoItem As Boolean

On Error GoTo TrataErros

   'Atualiza variaveis de controle
   AtualizarDataPrevistaPreparoItensCozinha = True
   
   'Limpa lista para atualização data prevista preparo
   plvwEnviarItensCozinha.ListItems.Clear

   'Atualiza variaveis de controle
   vdl_DataHoaAtual = VerDataHoraAtual()

   'Pesquisa itens para atualização da data prevista preparo
   vol_Conexao.AbrirConexao
         
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
           
         'Calcula hora prevista do item
         Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, IdComanda, IdTipoComanda, IdCardapio, DescricaoItem, Quantidade, QtdeItemCozinha, TempoPreparo, IdCozinha, DescricaoCozinha, Capacidade, " _
                                & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                & "IIF(UltimaDataPrevistaPreparo is null,'????', FORMAT (UltimaDataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS UltimaDataPrevistaPreparo, " _
                                & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                & "StatusItem, " _
                                & "CASE " _
                                & "    WHEN StatusItem = 1 THEN 'Item aguardando envio' " _
                                & "    WHEN StatusItem = 2 THEN 'Item aguardando processamente' " _
                                & "    WHEN StatusItem = 3 THEN 'Item sendo prepadada' " _
                                & "    WHEN StatusItem = 4 THEN 'Item para entrega' " _
                                & "    WHEN StatusItem = 5 THEN 'Item cancelado' " _
                                & "END StatusItemDescricao " _
                                & "FROM ComandasItem (NOLOCK) " _
                                & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem, TempoPreparo, IdCozinha FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID  " _
                                & "INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha, Capacidade FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaId " _
                                & "INNER JOIN (SELECT Id As CadapioId, ISNULL(QtdeItemCozinha,0) AS QtdeItemCozinha FROM Cardapio (NOLOCK) " _
                                & "         LEFT JOIN ( SELECT IdCardapio, COUNT(*) AS QtdeItemCozinha FROM ComandasItem WHERE StatusItem IN(2, 3) GROUP BY IdCardapio) AS ComandasItem ON Cardapio.Id = ComandasItem.IdCardapio ) As CardapioCozinha ON ComandasItem.IdCardapio = CardapioCozinha.CadapioId " _
                                & "INNER JOIN (SELECT IdCozinha As CozinhaId, MAX(DataPrevistaPreparo) UltimaDataPrevistaPreparo FROM ComandasItem (NOLOCK) " _
                                & "         INNER JOIN (SELECT Id As CardapioId, Idcozinha, TempoPreparo FROM Cardapio (NOLOCK)) AS Cardapio ON Cardapio.CardapioId = ComandasItem.IdCardapio " _
                                & "         INNER JOIN (SELECT Id As CardapioId, IdCozinha As CozinhaId, TempoPreparo FROM Cardapio (NOLOCK)) AS CardapioCozinha ON CardapioCozinha.CozinhaId = Cardapio.IdCozinha WHERE StatusItem IN(2, 3) GROUP BY IdCozinha ) AS CozinhaCardapio ON CozinhaCardapio.CozinhaId = Cardapio.IdCozinha " _
                                & "WHERE StatusItem = 2 " _
                                & "ORDER BY IdCozinha, TempoPreparo, IdCardapio, DataConfirmacaoPreparo")
         With vol_Recordset
            If .RecordCount <> 0 Then
               Do Until .EOF
                  plvwEnviarItensCozinha.ListItems.Add , , .Fields(0).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(1) = .Fields(1).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(2) = .Fields(2).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(3) = .Fields(3).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(4).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(5).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(6).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(7).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(8).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(9).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(10).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(11).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(12).Value
                  plvwEnviarItensCozinha.ListItems(plvwEnviarItensCozinha.ListItems.Count).SubItems(4) = .Fields(13).Value
                  .MoveNext
               Loop
            End If
         End With
         
   vol_Conexao.FecharConexao
   
   If plvwEnviarItensCozinha.ListItems.Count <> 0 Then
   
       vol_Conexao.AbrirConexao
          vol_Conexao.cConexao.BeginTrans
       
             'Atualiza variaveis de controle
             vbl_AtualizaDataPrevistaPreparoItem = False
          
             'Alterar Itens da comanda
             For vil_CountLista = 1 To plvwEnviarItensCozinha.ListItems.Count
             
                   Set vol_Command = New ADODB.Command
                   Set vol_Command.ActiveConnection = vol_Conexao.cConexao
                   If Int(plvwEnviarItensCozinha.ListItems(vil_CountLista).text) <> 0 Then
                                                                     
                      'Pesquisa UltimaDataPrevistaPreparo
                      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, IdComanda, IdTipoComanda, IdCardapio, DescricaoItem, Quantidade, QtdeItemCozinha, TempoPreparo, IdCozinha, DescricaoCozinha, Capacidade, " _
                                             & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                             & "IIF(UltimaDataPrevistaPreparo is null,'????', FORMAT (UltimaDataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS UltimaDataPrevistaPreparo, " _
                                             & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                             & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                             & "StatusItem, " _
                                             & "CASE " _
                                             & "    WHEN StatusItem = 1 THEN 'Item aguardando envio' " _
                                             & "    WHEN StatusItem = 2 THEN 'Item aguardando processamente' " _
                                             & "    WHEN StatusItem = 3 THEN 'Item sendo prepadada' " _
                                             & "    WHEN StatusItem = 4 THEN 'Item para entrega' " _
                                             & "    WHEN StatusItem = 5 THEN 'Item cancelado' " _
                                             & "END StatusItemDescricao " _
                                             & "FROM ComandasItem (NOLOCK) " _
                                             & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                             & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem, TempoPreparo, IdCozinha FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID  " _
                                             & "INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha, Capacidade FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaId " _
                                             & "INNER JOIN (SELECT Id As CadapioId, ISNULL(QtdeItemCozinha,0) AS QtdeItemCozinha FROM Cardapio (NOLOCK) " _
                                             & "         LEFT JOIN ( SELECT IdCardapio, COUNT(*) AS QtdeItemCozinha FROM ComandasItem WHERE StatusItem IN(2, 3) GROUP BY IdCardapio) AS ComandasItem ON Cardapio.Id = ComandasItem.IdCardapio ) As CardapioCozinha ON ComandasItem.IdCardapio = CardapioCozinha.CadapioId " _
                                             & "INNER JOIN (SELECT IdCozinha As CozinhaId, MAX(DataPrevistaPreparo) UltimaDataPrevistaPreparo FROM ComandasItem (NOLOCK) " _
                                             & "         INNER JOIN (SELECT Id As CardapioId, Idcozinha, TempoPreparo FROM Cardapio (NOLOCK)) AS Cardapio ON Cardapio.CardapioId = ComandasItem.IdCardapio " _
                                             & "         INNER JOIN (SELECT Id As CardapioId, IdCozinha As CozinhaId, TempoPreparo FROM Cardapio (NOLOCK)) AS CardapioCozinha ON CardapioCozinha.CozinhaId = Cardapio.IdCozinha WHERE StatusItem IN(2, 3) GROUP BY IdCozinha ) AS CozinhaCardapio ON CozinhaCardapio.CozinhaId = Cardapio.IdCozinha " _
                                             & "WHERE StatusItem = 2 " + "AND IdCardapio = '" + plvwEnviarItensCozinha.ListItems(vil_CountLista).SubItems(3) + "' " _
                                             & "ORDER BY IdCozinha, TempoPreparo, IdCardapio, DataConfirmacaoPreparo")
                      With vol_Recordset
                         If .RecordCount <> 0 Then
                            Do Until .EOF
                               
                               'UltimaDataPrevistaPreparo
                               vbl_AtualizaDataPrevistaPreparoItem = True
                               
                               If Trim$(.Fields(12).Value) = "????" Then
                                  vdl_DataPrevistaPreparItemCozinha = DateAdd("n", Int(Trim$(.Fields(7).Value)), CDate(Trim$(vdl_DataHoaAtual)))     '
                               Else
                                  'QuantidadeItemCozinha      'Capacidade
                                  If Int(.Fields(6).Value) < Int(.Fields(10).Value) Then
                                     'UltimaDataPrevistaPreparo
                                     vdl_DataPrevistaPreparItemCozinha = DateAdd("n", Int(Trim$(.Fields(7).Value)), CDate(Trim$(vdl_DataHoaAtual)))
                                  Else
                                     'TempoPreparo + UltimaDataPrevistaPreparo
                                     If CDate(Trim$(.Fields(12).Value)) < CDate(Trim$(vdl_DataHoaAtual)) Then
                                        vdl_DataPrevistaPreparItemCozinha = DateAdd("n", Int(Trim$(.Fields(7).Value)), CDate(Trim$(vdl_DataHoaAtual)))
                                     Else
                                        vdl_DataPrevistaPreparItemCozinha = DateAdd("n", Int(Trim$(.Fields(7).Value)), CDate(Trim$(.Fields(12).Value)))
                                     End If
                                  End If
                               End If
                               .MoveNext
                            Loop
                         End If
                      End With
                     
                      'Atualiza DataPrevistaPreparo Item
                      If vbl_AtualizaDataPrevistaPreparoItem = True Then
      
                         vol_Command.CommandType = adCmdStoredProc
                         vol_Command.CommandText = "SP_Teste_U_AtualizarDataPrevistaPreparoItensCozinha"
                         vol_Command.Parameters.Refresh
                      
                         With vol_Command.Parameters
                           .Item(1).Value = plvwEnviarItensCozinha.ListItems(vil_CountLista).text
                           .Item(2).Value = CDate(vdl_DataHoaAtual)
                           .Item(3).Value = CDate(vdl_DataPrevistaPreparItemCozinha)
                           .Item(4).Value = "3"
                         End With
                         vol_Command.Execute
                      End If
                         
                   End If
                                    
             Next vil_CountLista
          
             'Atualiza DataPrevistaPreparo Comanda
             If plvwEnviarItensCozinha.ListItems.Count <> 0 Then
             
                'Atualiza variaveis de controle
                vil_IdComanda = 0
             
                Set vol_Command = New ADODB.Command
                Set vol_Command.ActiveConnection = vol_Conexao.cConexao
                
                For vil_CountLista = 1 To plvwEnviarItensCozinha.ListItems.Count
                   
                    If Int(plvwEnviarItensCozinha.ListItems(vil_CountLista).text) <> 0 Then
                   
                       If vil_IdComanda <> Int(plvwEnviarItensCozinha.ListItems(vil_CountLista).SubItems(1)) Then
                       
                          vil_IdComanda = Int(plvwEnviarItensCozinha.ListItems(vil_CountLista).SubItems(1))
                          
                          Set vol_Recordset = vol_Conexao.RetornaRs("SELECT IdComanda, Max(DataPrevistaPreparo) As DataPrevistaPreparo " _
                                                 & "FROM ComandasItem (NOLOCK) " _
                                                 & "INNER JOIN (SELECT Id As ComandaId, IdTipoComanda FROM Comandas (NOLOCK)) As Comandas ON ComandasItem.IdComanda = Comandas.ComandaId " _
                                                 & "INNER JOIN (SELECT Id As CardapioId, Descricao As DescricaoItem, TempoPreparo, IdCozinha FROM Cardapio (NOLOCK)) As Cardapio ON ComandasItem.IdCardapio = Cardapio.CardapioID  " _
                                                 & "INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha, Capacidade FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaId " _
                                                 & "INNER JOIN (SELECT Id As CadapioId, ISNULL(QtdeItemCozinha,0) AS QtdeItemCozinha FROM Cardapio (NOLOCK) " _
                                                 & "         LEFT JOIN ( SELECT IdCardapio, COUNT(*) AS QtdeItemCozinha FROM ComandasItem WHERE StatusItem IN(2, 3) GROUP BY IdCardapio) AS ComandasItem ON Cardapio.Id = ComandasItem.IdCardapio ) As CardapioCozinha ON ComandasItem.IdCardapio = CardapioCozinha.CadapioId " _
                                                 & "INNER JOIN (SELECT IdCozinha As CozinhaId, MAX(DataPrevistaPreparo) UltimaDataPrevistaPreparo FROM ComandasItem (NOLOCK) " _
                                                 & "         INNER JOIN (SELECT Id As CardapioId, Idcozinha, TempoPreparo FROM Cardapio (NOLOCK)) AS Cardapio ON Cardapio.CardapioId = ComandasItem.IdCardapio " _
                                                 & "         INNER JOIN (SELECT Id As CardapioId, IdCozinha As CozinhaId, TempoPreparo FROM Cardapio (NOLOCK)) AS CardapioCozinha ON CardapioCozinha.CozinhaId = Cardapio.IdCozinha WHERE StatusItem IN(2, 3) GROUP BY IdCozinha ) AS CozinhaCardapio ON CozinhaCardapio.CozinhaId = Cardapio.IdCozinha " _
                                                 & "WHERE IdComanda = '" + CStr(vil_IdComanda) + "' " _
                                                 & "GROUP BY IdComanda")
                          With vol_Recordset
                            If .RecordCount <> 0 Then
                               Do Until .EOF
                                  vil_IdComanda = .Fields(0).Value
                                  vdl_DataPrevistaPreparComanda = CDate(.Fields(1).Value)
                                  .MoveNext
                               Loop
                            End If
                          End With
                          '
                          If vdl_DataPrevistaPreparComanda <> Empty Then
                             vol_Command.CommandType = adCmdStoredProc
                             vol_Command.CommandText = "SP_Teste_U_AtualizarDataPrevistaPreparoComanda"
                             vol_Command.Parameters.Refresh
                             With vol_Command.Parameters
                               .Item(1).Value = vil_IdComanda
                               .Item(2).Value = CDate(vdl_DataPrevistaPreparComanda)
                             End With
                             vol_Command.Execute
                          End If
                       
                       End If
                       
                    End If
                    
                Next vil_CountLista
                          
             End If
       
          vol_Conexao.cConexao.CommitTrans
      vol_Conexao.FecharConexao
  
  End If

   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AtualizarDataPrevistaPreparoItensCozinha = False
       MsgBox "Não foi possível enviar item da Comanda para processamento na cozinha !", vbExclamation
    End If
End Function

Public Function PesquisarComanda(ByRef plvwComandas As ListView, ByVal pComanda As String, ByVal pBusca As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   PesquisarComanda = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, IdTipoComanda, NumeroMesa, QuantidadePessoa, DataConfirmacaoPreparo, DataPrevistaPreparo, DataFinalizacaoPreparo, StatusComanda FROM Comandas (NOLOCK) WHERE " + IIf(pBusca = 0, " Id = ", " Descricao = ") + "'" + pComanda + "'")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwComandas.ListItems.Add , , .Fields(0).Value
               plvwComandas.ListItems(plvwComandas.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwComandas.ListItems(plvwComandas.ListItems.Count).SubItems(2) = .Fields(2).Value
               plvwComandas.ListItems(plvwComandas.ListItems.Count).SubItems(3) = .Fields(3).Value
               plvwComandas.ListItems(plvwComandas.ListItems.Count).SubItems(4) = .Fields(4).Value
               plvwComandas.ListItems(plvwComandas.ListItems.Count).SubItems(5) = IIf(IsNull(.Fields(5).Value), "????", .Fields(5).Value)
               plvwComandas.ListItems(plvwComandas.ListItems.Count).SubItems(6) = IIf(IsNull(.Fields(6).Value), "????", .Fields(6).Value)
               plvwComandas.ListItems(plvwComandas.ListItems.Count).SubItems(7) = IIf(IsNull(.Fields(7).Value), "????", .Fields(7).Value)
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
        PesquisarComanda = False
        MsgBox "Não foi possível encontrar informações da Comanda !", vbExclamation
    End If
End Function

Public Function LocalizarComanda(ByRef padoComandas As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarComanda = True
   vol_Conexao.AbrirConexao
         padoComandas.RecordSource = "SELECT Id, IdTipoComanda, DescricaoComanda, NumeroMesa, QuantidadePessoa, " _
                                   & "FORMAT(TotalPreco, '#,0.00','pt-br') AS TotalPreco, " _
                                   & "IIF(DataConfirmacaoPreparo is null,'????', FORMAT (DataConfirmacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataConfirmacaoPreparo, " _
                                   & "IIF(DataPrevistaPreparo is null,'????', FORMAT (DataPrevistaPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br') ) AS DataPrevistaPreparo, " _
                                   & "IIF(DataFinalizacaoPreparo is null,'????', FORMAT (DataFinalizacaoPreparo, 'dd-MM-yyyy HH:mm:ss', 'pt-br')  ) AS DataFinalizacaoPreparo, " _
                                   & "StatusComanda, " _
                                   & "StatusComandaDescricao " _
                                   & "FROM Comandas (NOLOCK) " _
                                   & "INNER JOIN (SELECT Id AS ComandaId, IIF(StatusComanda=1,'Comanda aberta',IIF(StatusComanda=2,'Comanda fechada','Comanda cancelada')) As StatusComandaDescricao FROM Comandas (NOLOCK)) As ComandasTipo ON Comandas.Id = ComandasTipo.ComandaId " _
                                   & "INNER JOIN (SELECT Id AS IdTiposComanda, Descricao AS DescricaoComanda FROM TiposComandas (NOLOCK)) As TiposComandas ON TiposComandas.IdTiposComanda = Comandas.IdTipoComanda " _
                                   & "INNER JOIN (SELECT IdComanda, SUM(TotalPreco) AS TotalPreco FROM ComandasItem (NOLOCK) GROUP BY IdComanda) As ComandasItem ON ComandasItem.IdComanda = Comandas.Id " _
                                   & "WHERE " + IIf(pcmbLocalizar = 0, "IdTipoComanda", IIf(pcmbLocalizar = 1, "NumeroMesa", "StatusComandaDescricao")) + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "IdTipoComanda", IIf(pcmbLocalizar = 1, "NumeroMesa", "StatusComandaDescricao")) + " "
      padoComandas.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarComanda = False
        MsgBox "Não foi possível encontrar informações da Comanda !", vbExclamation
    End If
End Function

Public Function ExcluirComanda(ByVal pIdcomanda As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_D_Comanda"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdcomanda
         End With
      
         vol_Command.Execute
   
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       ExcluirComanda = False
       MsgBox "Não foi possível excluir informações da Comanda !", vbExclamation
    End If
End Function























