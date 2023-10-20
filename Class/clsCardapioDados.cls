VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardapioDados"
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
Public Function CarregarGridCardapioRS(ByRef padoCardapio As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CarregarGridCardapioRS = True
   vol_Conexao.AbrirConexao
      padoCardapio.RecordSource = "SELECT Id, Descricao as DescricaoItem, IdCozinha, DescricaoCozinha, TempoPreparo, Preco FROM Cardapio (NOLOCK) INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaID ORDER BY " + IIf(pcmbLocalizar <= 0, "Id", IIf(pcmbLocalizar = 1, "DescricaoItem", "DescricaoCozinha")) + " "
      padoCardapio.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridCardapioRS = False
        MsgBox "Não foi possível carregar informações do Cardapio !", vbExclamation
    End If
End Function

Public Function PesquisarCardapio(ByRef plvwCardapio As ListView, ByVal pIdCardapio As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   PesquisarCardapio = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, IdCozinha, TempoPreparo, Preco, Descricao, DescricaoCozinha FROM Cardapio (NOLOCK) INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaID WHERE Id = " & pIdCardapio & " ")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwCardapio.ListItems.Add , , .Fields(0).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(2) = .Fields(2).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(3) = .Fields(3).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(4) = .Fields(4).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(5) = .Fields(5).Value
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
        PesquisarCardapio = False
        MsgBox "Não foi possível encontrar informações do Cardapio !", vbExclamation
    End If
End Function

Public Function PesquisarCardapioItemComanda(ByRef plvwCardapio As ListView, ByVal pItem As String, ByVal pBusca As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   PesquisarCardapioItemComanda = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, IdCozinha, TempoPreparo, Preco, Descricao, DescricaoCozinha FROM Cardapio (NOLOCK) INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaID WHERE " + IIf(pBusca = 0, " Id = ", " Descricao = ") + "'" + pItem + "'")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwCardapio.ListItems.Add , , .Fields(0).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(2) = .Fields(2).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(3) = .Fields(3).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(4) = .Fields(4).Value
               plvwCardapio.ListItems(plvwCardapio.ListItems.Count).SubItems(5) = .Fields(5).Value
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
        PesquisarCardapioItemComanda = False
        MsgBox "Não foi possível encontrar informações do Item da Comanda !", vbExclamation
    End If
End Function

Public Function LocalizarCardapio(ByRef padoCardapio As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarCardapio = True
   vol_Conexao.AbrirConexao
      If Trim$(ptxtLocalizar) = Empty Then
         padoCardapio.RecordSource = "SELECT Id, Descricao as DescricaoItem, IdCozinha, DescricaoCozinha, TempoPreparo, Preco FROM Cardapio (NOLOCK) INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaID ORDER BY " + IIf(pcmbLocalizar <= 0, "Id", IIf(pcmbLocalizar = 1, "DescricaoItem", "DescricaoCozinha")) + " "
      Else
         padoCardapio.RecordSource = "SELECT Id, Descricao as DescricaoItem, IdCozinha, DescricaoCozinha, TempoPreparo, Preco FROM Cardapio (NOLOCK) INNER JOIN (SELECT Id As CozinhaId, Descricao As DescricaoCozinha FROM Cozinhas (NOLOCK)) As Cozinhas ON Cardapio.IdCozinha = Cozinhas.CozinhaID WHERE " + IIf(pcmbLocalizar <= 0, "Id", IIf(pcmbLocalizar = 1, "Descricao", "descricaoCozinha")) + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "Id", IIf(pcmbLocalizar = 1, "Descricao", "DescricaoCozinha")) + " "
      End If
      padoCardapio.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarCardapio = False
        MsgBox "Não foi possível encontrar informações do Cardapio !", vbExclamation
    End If
End Function

Public Function IncluirCardapio(ByVal pIdCozinha As Integer, _
                                ByVal pDescricao As String, _
                                ByVal pTempoPreparo As Integer, _
                                ByVal pPreco As Double) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirCardapio = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_I_Cardapio"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdCozinha
           .Item(2).Value = pDescricao
           .Item(3).Value = pTempoPreparo
           .Item(4).Value = pPreco
         End With
   
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirCardapio = False
       MsgBox "Não foi possível incluir informações do Cardapio !", vbExclamation
    End If
End Function

Public Function AlterarCardapio(ByVal pIdCardapio As Integer, _
                                ByVal pIdCozinha As Integer, _
                                ByVal pDescricao As String, _
                                ByVal pTempoPreparo As Integer, _
                                ByVal pPreco As Double) As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         AlterarCardapio = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_U_Cardapio"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdCardapio
           .Item(2).Value = pIdCozinha
           .Item(3).Value = pDescricao
           .Item(4).Value = pTempoPreparo
           .Item(5).Value = pPreco
         End With
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarCardapio = False
       MsgBox "Não foi possível alterar informações do Cardapio !", vbExclamation
    End If
End Function

Public Function ExcluirCardapio(ByVal pIdCardapio As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirCardapio = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_D_Cardapio"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdCardapio
         End With
      
         vol_Command.Execute
   
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       ExcluirCardapio = False
       MsgBox "Não foi possível excluir informações do Cardapio !", vbExclamation
    End If
End Function























