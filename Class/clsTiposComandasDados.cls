VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiposComandasDados"
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
Public Function CarregarGridTipoComandaRS(ByRef padoTiposComandas As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CarregarGridTipoComandaRS = True
   vol_Conexao.AbrirConexao
      padoTiposComandas.RecordSource = "SELECT Id, Descricao FROM TiposComandas (NOLOCK) ORDER BY " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " "
      padoTiposComandas.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridTipoComandaRS = False
        MsgBox "Não foi possível carregar informações da Comanda !", vbExclamation
    End If
End Function

Public Function PesquisarTipoComanda(ByRef plvwTiposComandas As ListView, ByVal pComanda As String, ByVal pBusca As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   PesquisarTipoComanda = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, Descricao, ISNULL(StatusComanda,0) As StatusComanda FROM TiposComandas (NOLOCK) " _
                                              & "LEFT JOIN (SELECT IdTipoComanda, StatusComanda FROM Comandas (NOLOCK)) AS Comandas ON TiposComandas.Id = Comandas.IdTipoComanda " _
                                              & "WHERE " + IIf(pBusca = 0, " Id = ", " Descricao = ") + "'" + pComanda + "' ")
      With vol_Recordset
         plvwTiposComandas.ListItems.Clear
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwTiposComandas.ListItems.Add , , .Fields(0).Value
               plvwTiposComandas.ListItems(plvwTiposComandas.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwTiposComandas.ListItems(plvwTiposComandas.ListItems.Count).SubItems(2) = .Fields(2).Value
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
        PesquisarTipoComanda = False
        MsgBox "Não foi possível encontrar informações da Comanda !", vbExclamation
    End If
End Function

Public Function LocalizarTipoComanda(ByRef padoTiposComandas As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarTipoComanda = True
   vol_Conexao.AbrirConexao
      If Trim$(ptxtLocalizar) = Empty Then
         padoTiposComandas.RecordSource = "SELECT Id, Descricao FROM TiposComandas (NOLOCK) ORDER BY " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " "
      Else
         padoTiposComandas.RecordSource = "SELECT Id, Descricao FROM TiposComandas (NOLOCK) WHERE " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " "
      End If
      padoTiposComandas.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarTipoComanda = False
        MsgBox "Não foi possível encontrar informações da Comanda !", vbExclamation
    End If
End Function

Public Function IncluirTipoComanda(ByVal pDescricao As String) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirTipoComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_I_TiposComandas"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pDescricao
         End With
   
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirTipoComanda = False
       MsgBox "Não foi possível incluir informações da Comanda !", vbExclamation
    End If
End Function

Public Function AlterarTipoComanda(ByVal pIdcomanda As Integer, _
                               ByVal pDescricao As String) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         AlterarTipoComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_U_TiposComandas"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdcomanda
           .Item(2).Value = pDescricao
         End With
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarTipoComanda = False
       MsgBox "Não foi possível incluir informações da Comanda !", vbExclamation
    End If
End Function

Public Function ExcluirTipoComanda(ByVal pIdcomanda As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirTipoComanda = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_D_TiposComandas"
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
       ExcluirTipoComanda = False
       MsgBox "Não foi possível excluir informações da Comanda !", vbExclamation
    End If
End Function























