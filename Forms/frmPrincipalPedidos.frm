VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de pedidos KDS software"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrEnviarItensCozinha 
      Interval        =   60000
      Left            =   9075
      Top             =   660
   End
   Begin MSComctlLib.ImageList imlBotões 
      Left            =   9840
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":02B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":0414
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":0570
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":06CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":07E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":093C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":0A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":0BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":1190
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalPedidos.frx":12EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Barra 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16200
      _ExtentX        =   28575
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlBotões"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoNovo"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoBarra01"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoImprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoEditar"
            Object.ToolTipText     =   "Editar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoExcluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoBarra02"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoAgrupar"
            Object.ToolTipText     =   "Agrupar"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoDetalhes"
            Object.ToolTipText     =   "Exibir Detalhes"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoAtualizar"
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoClassificar"
            Object.ToolTipText     =   "Classificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btoCalculadora"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoConfig"
            Object.ToolTipText     =   "Configurações"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoLogin"
            Object.ToolTipText     =   "Login"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoAjuda"
            Object.ToolTipText     =   "Ajuda"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwEnviarItensCozinha 
      Height          =   360
      Left            =   10755
      TabIndex        =   2
      Top             =   675
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IdComanda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IdTipoComanda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IdCardapio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DescricaoItem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "QuantidadeItemCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TempoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "IdCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "DescricaoCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Capacidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "DataConfirmacaoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "UltimaDataPrevistaPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "DataPrevistaPreparo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de pedidos KDS software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   2730
      Left            =   2865
      TabIndex        =   0
      Top             =   975
      Width           =   5865
   End
   Begin VB.Image Image1 
      DragMode        =   1  'Automatic
      Height          =   6000
      Left            =   -15
      Picture         =   "frmPrincipalPedidos.frx":1448
      Top             =   -15
      Width           =   16020
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu mnuCozinha 
         Caption         =   "Cozinhas"
      End
      Begin VB.Menu mnuCardapio 
         Caption         =   "Itens do Cardápio"
      End
      Begin VB.Menu mnuTiposComanda 
         Caption         =   "Comandas"
      End
   End
   Begin VB.Menu mnuMovimentacao 
      Caption         =   "Movimentação"
      Begin VB.Menu mnuComanda 
         Caption         =   "Comandas"
      End
      Begin VB.Menu mnuFinalizarItemcozinha 
         Caption         =   "Itens Cozinha"
      End
      Begin VB.Menu mnuEstatisticas 
         Caption         =   "Estatistica"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_ComandasNegocios As New clsComandasNegocios



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo TrataErros

    'Tecla de atalho da calculadora
    If KeyCode = vbKeyF7 Then
        KeyCode = 0
        Exit Sub
    End If
    'Tecla de sair do form
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmMain = Nothing
   End
   
End Sub

Private Sub mnuCardapio_Click()
   frmCardapioLista.Show vbModal
End Sub

Private Sub mnuComanda_Click()
   frmComandasLista.Show vbModal
End Sub

Private Sub mnuCozinha_Click()
   frmCozinhasLista.Show vbModal
End Sub

Private Sub mnuEstatisticas_Click()
   frmEstatisticaItensLista.Show vbModal
End Sub

Private Sub mnuFinalizarItemcozinha_Click()
   frmCozinhaItensLista.Show vbModal
End Sub

Private Sub mnuTiposComanda_Click()
   frmTiposComandasLista.Show vbModal
End Sub

Private Sub Barra_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "btoCalculadora"
            Call Calculadora
    End Select
End Sub

Private Sub tmrEnviarItensCozinha_Timer()
   Call EnviarItensCozinha
End Sub

Private Function EnviarItensCozinha() As Boolean

   'Envia item para cozinha e processamento das horas previstas
   Set vop_ComandasNegocios = New clsComandasNegocios
          If vop_ComandasNegocios.EnviarItensCozinha(lvwEnviarItensCozinha) = True Then
          'Atualiza dados entrada e saida
          frmComandasLista.CarregarGrid
       End If
   Set vop_ComandasNegocios = Nothing

End Function
