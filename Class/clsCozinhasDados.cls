VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCozinhasDados"
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
Public Function CarregarGridCozinhaRS(ByRef padoCozinhas As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   CarregarGridCozinhaRS = True
   vol_Conexao.AbrirConexao
      padoCozinhas.RecordSource = "SELECT Id, Descricao, Capacidade FROM Cozinhas (NOLOCK) ORDER BY " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " "
      padoCozinhas.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridCozinhaRS = False
        MsgBox "Não foi possível carregar informações da Cozinha !", vbExclamation
    End If
End Function

Public Function PesquisarCozinha(ByRef plvwCozinhas As ListView, ByVal pCozinha As String, ByVal pBusca As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   PesquisarCozinha = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT Id, Descricao, Capacidade FROM Cozinhas (NOLOCK) WHERE " + IIf(pBusca = 0, " Id = ", " Descricao = ") + "'" + pCozinha + "'")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwCozinhas.ListItems.Add , , .Fields(0).Value
               plvwCozinhas.ListItems(plvwCozinhas.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwCozinhas.ListItems(plvwCozinhas.ListItems.Count).SubItems(2) = .Fields(2).Value
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
        PesquisarCozinha = False
        MsgBox "Não foi possível encontrar informações da Cozinha !", vbExclamation
    End If
End Function

Public Function LocalizarCozinha(ByRef padoCozinhas As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarCozinha = True
   vol_Conexao.AbrirConexao
      If Trim$(ptxtLocalizar) = Empty Then
         padoCozinhas.RecordSource = "SELECT Id, Descricao, Capacidade FROM Cozinhas (NOLOCK) ORDER BY " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " "
      Else
         padoCozinhas.RecordSource = "SELECT Id, Descricao, Capacidade FROM Cozinhas (NOLOCK) WHERE " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "Id", "Descricao") + " "
      End If
      padoCozinhas.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarCozinha = False
        MsgBox "Não foi possível encontrar informações da Cozinha !", vbExclamation
    End If
End Function

Public Function IncluirCozinha(ByVal pDescricao As String, _
                                ByVal pCapacidade As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirCozinha = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_I_Cozinhas"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pDescricao
           .Item(2).Value = pCapacidade
         End With
   
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirCozinha = False
       MsgBox "Não foi possível incluir informações da Cozinha !", vbExclamation
    End If
End Function

Public Function AlterarCozinha(ByVal pIdCozinha As Integer, _
                               ByVal pDescricao As String, _
                               ByVal pCapacidade As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         AlterarCozinha = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_U_Cozinhas"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdCozinha
           .Item(2).Value = pDescricao
           .Item(3).Value = pCapacidade
         End With
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarCozinha = False
       MsgBox "Não foi possível incluir informações da Cozinha !", vbExclamation
    End If
End Function

Public Function ExcluirCozinha(ByVal pIdCozinha As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirCozinha = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_D_Cozinhas"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdCozinha
         End With
      
         vol_Command.Execute
   
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       ExcluirCozinha = False
       MsgBox "Não foi possível excluir informações da Cozinha !", vbExclamation
    End If
End Function





















