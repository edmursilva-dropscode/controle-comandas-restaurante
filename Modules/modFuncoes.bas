Attribute VB_Name = "modFuncoes"
Option Explicit

'Variáveis do ADO
Private vol_Conexao As New clsConexao
Private vol_System As New clsSystem

'Variaveis pricate
Private vcl_WinDir As String

'Constantes usadas para acessar o Registro
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&

'Declaração para acessar o Registro
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long



Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub

Public Function LimpaCampos(ByVal Tela As Form)
Dim vil_Contador As Integer
    
    vil_Contador = 0
    For vil_Contador = 0 To Tela.Controls.Count - 1
        If TypeOf Tela.Controls(vil_Contador) Is TextBox Then
            Tela.Controls(vil_Contador).text = Empty
            Tela.Controls(vil_Contador).ForeColor = vbBlack
        End If
        If TypeOf Tela.Controls(vil_Contador) Is Label Then
            If Tela.Controls(vil_Contador).BorderStyle = 1 Then
               Tela.Controls(vil_Contador).Caption = Empty
               Tela.Controls(vil_Contador).ForeColor = vbBlack
            End If
        End If
        If TypeOf Tela.Controls(vil_Contador) Is CheckBox Then
            Tela.Controls(vil_Contador).Value = 0
        End If
        If TypeOf Tela.Controls(vil_Contador) Is ComboBox Then
            If Tela.Controls(vil_Contador).Style <> 2 Then
                Tela.Controls(vil_Contador).text = Empty
                Tela.Controls(vil_Contador).ForeColor = vbBlack
            Else
                If Tela.Controls(vil_Contador).Tag <> Empty Then
                    Tela.Controls(vil_Contador).ForeColor = vbBlack
                End If
            End If
        End If
        If TypeOf Tela.Controls(vil_Contador) Is DTPicker Then
            If Tela.Controls(vil_Contador).Format = dtpShortDate Then
                Tela.Controls(vil_Contador).Value = Date
            End If
        End If
    Next vil_Contador

End Function

Public Function VerDataHoraAtual() As String
   
   VerDataHoraAtual = vol_System.DataHoraAtual()

End Function

'Public Function ComboBox(ByVal pCombo As Control, ByVal pTable$, ByVal pCampo$, Optional pLista As ListView)
'
'   Call vol_System.CarregaCombo(pCombo, pTable$, pCampo$, pLista)
'
'End Function

Public Function ComboBox(Combo As Object, Tabela As String, ID As String, Descricao As String, Condicao As String)

   Call vol_System.CarregaCombo(Combo, Tabela, ID, Descricao, Condicao)

End Function


Function Calculadora()
Dim vcl_Rum As String, vcl_Versao As String
Dim fso As Object, WinDir As Object
On Error GoTo TrataErros

    Set fso = CreateObject("Scripting.FileSystemObject")
        Set WinDir = fso.GetSpecialFolder(0)
            vcl_WinDir = WinDir
        Set WinDir = Nothing
    Set fso = Nothing

    vcl_Versao = Left(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "VersionNumber"), 1)
    If Trim$(vcl_Versao) = Empty Then
        vcl_Versao = Trim(Left(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "CurrentVersion"), 1))
    End If
    
    If vcl_Versao = "5" Or vcl_Versao = "6" Then
        vcl_Rum = Shell(Trim(vcl_WinDir) & "\System32\CALC.EXE", 1)
    Else
        vcl_Rum = Shell(Trim(vcl_WinDir) & "\CALC.EXE", 1)
    End If

TrataErros:
    If Err = 53 Then
        Exit Function
    End If

End Function

Public Function GetString(hKey As Long, strPath As String, strValue As String)
Dim strBuf As String, intZeroPos As Integer, r As String, lValueType
Dim keyhand As Long, datatype As Long, lResult As Long, lDataBufSize As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Function TotalizaComanda() As Boolean
Dim vil_CountLista As Integer
    
    'Totaliza valor total da comanda
    frmComandas.lblTotalComanda.Caption = "0,00"
    For vil_CountLista = 1 To frmComandas.lvwItensComanda.ListItems.Count
        frmComandas.lblTotalComanda.Caption = Format(CDbl(frmComandas.lblTotalComanda.Caption) + CDbl(frmComandas.lvwItensComanda.ListItems(vil_CountLista).SubItems(7)), "##,##0.00")
    Next vil_CountLista

End Function

Public Function StatusItemDescricao(IdStatusItem As Integer) As String

   If IdStatusItem = 1 Then
      StatusItemDescricao = "Item aguardando envio"
   ElseIf IdStatusItem = 2 Then
      StatusItemDescricao = "Item aguardando processamente"
   ElseIf IdStatusItem = 3 Then
      StatusItemDescricao = "Item sendo prepadada"
   ElseIf IdStatusItem = 4 Then
      StatusItemDescricao = "Item para entrega"
   ElseIf IdStatusItem = 5 Then
      StatusItemDescricao = "Item cancelado"
   End If

End Function




