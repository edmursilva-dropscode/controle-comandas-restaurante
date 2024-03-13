VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "PRJXTAB.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCardapio 
   Caption         =   "Cadastro de Cadápio"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtIdCozinha 
      Height          =   285
      Left            =   1575
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   150
   End
   Begin prjXTab.XTab xtbCardapio 
      Height          =   2190
      Left            =   90
      TabIndex        =   2
      Top             =   765
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   3863
      TabCount        =   1
      TabCaption(0)   =   "  Cardápio "
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "fraTab1"
      Tab(0)ContCtrlCap(2)=   "lblCodigo"
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1665
         Index           =   1
         Left            =   165
         TabIndex        =   13
         Top             =   480
         Width           =   6765
         Begin VB.TextBox txtTempoDePreparo 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1740
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "0"
            Top             =   825
            Width           =   615
         End
         Begin VB.TextBox txtPreco 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1740
            MaxLength       =   9
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   1260
            Width           =   1500
         End
         Begin VB.ComboBox cmbLocalDePreparo 
            Height          =   315
            ItemData        =   "frmCardapio.frx":0000
            Left            =   1755
            List            =   "frmCardapio.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   420
            Width           =   2745
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   1755
            MaxLength       =   70
            TabIndex        =   0
            Tag             =   "0"
            Top             =   0
            Width           =   4935
         End
         Begin VB.Label lblPreco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   15
            TabIndex        =   9
            Top             =   1290
            Width           =   555
         End
         Begin VB.Label lblTempoDePreparo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tempo de preparo:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   7
            Top             =   900
            Width           =   1665
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   4
            Top             =   75
            Width           =   915
         End
         Begin VB.Label lblLocalPreparo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local de preparo:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   15
            TabIndex        =   5
            Top             =   480
            Width           =   1515
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   5080
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1020
         TabIndex        =   14
         Top             =   15
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin LVbuttons.LaVolpeButton lvbFechar 
      Height          =   360
      Left            =   6135
      TabIndex        =   12
      Top             =   3075
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Fechar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCardapio.frx":0004
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   0   'False
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton lblGravar 
      Height          =   360
      Left            =   5070
      TabIndex        =   11
      Top             =   3075
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Gravar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCardapio.frx":0020
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   0   'False
      BSTYLE          =   0
   End
   Begin MSComctlLib.ListView lvwCozinhas 
      Height          =   360
      Left            =   7755
      TabIndex        =   16
      Top             =   1080
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descricao"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Capacdade"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCardapio 
      Height          =   360
      Left            =   7755
      TabIndex        =   17
      Top             =   1515
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descricao"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IdCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DescricaoCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "TempoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Preco"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmCardapio.frx":003C
      Top             =   75
      Width           =   480
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1980
      Picture         =   "frmCardapio.frx":0906
      Top             =   675
      Width           =   10740
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17955
   End
End
Attribute VB_Name = "frmCardapio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_CardapioNegocios As New clsCardapioNegocios
Dim vop_CozinhasNegocios As New clsCozinhasNegocios
'Variaveis de controle do form
Dim vbp_Cardapio As Boolean                             'Verifica uma inclusao ou alteracao


'Eventos
Private Sub Form_Activate()

    Me.Refresh
   
End Sub

Public Sub Form_Load()

    'Verifica uma inclusao ou alteracao
    vbp_Cardapio = False
    
    'Inicializa entrada e saida
    Call InicializaEntradaSaida
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo TrataErros

    'Tecla de sair do form
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmCardapio = Nothing
   
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmbLocalDePreparo_Click()

    Set vop_CozinhasNegocios = New clsCozinhasNegocios
    
        If vop_CozinhasNegocios.PesquisarCozinha(lvwCozinhas, cmbLocalDePreparo.text, 1) = True Then
           txtIdCozinha.text = vop_CozinhasNegocios.IdCozinhas
        Else
           MsgBox "Não foi possível encontrar a Cozinha !", vbCritical, "Cardapio"
        End If
        
    Set vop_CozinhasNegocios = Nothing

End Sub

Private Sub cmbLocalDePreparo_LostFocus()
    
    Set vop_CozinhasNegocios = New clsCozinhasNegocios
    
        If vop_CozinhasNegocios.PesquisarCozinha(lvwCozinhas, cmbLocalDePreparo.text, 1) = True Then
           txtIdCozinha.text = vop_CozinhasNegocios.IdCozinhas
        Else
           MsgBox "Não foi possível encontrar a Cozinha !", vbCritical, "Cardapio"
        End If
        
    Set vop_CozinhasNegocios = Nothing
    
End Sub

Private Sub txtTempoDePreparo_Change()

   If VerNumeros = False Then Exit Sub

End Sub

Private Sub txtTempoDePreparo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtTempoDePreparo_KeyPress(KeyAscii As Integer)
    
    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
            
End Sub

Private Sub txtPreco_GotFocus()
    
    txtPreco.SelStart = 0
    txtPreco.SelLength = Len(txtPreco.text)
    
End Sub

Private Sub txtPreco_Change()
   
   If VerNumeros = False Then Exit Sub
   
End Sub

Private Sub txtPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)

    If VerNumeros = False Then Exit Sub
    
    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    'Permite apenas números e formato de Moedas
    If KeyAscii <> 13 Then
       If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub lblGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Cardapio = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_CardapioNegocios = New clsCardapioNegocios
          vop_CardapioNegocios.IdCardapio = lblCodigo.Caption
          vop_CardapioNegocios.Descricao = txtDescricao.text
          vop_CardapioNegocios.IdCozinha = txtIdCozinha.text
          vop_CardapioNegocios.TempoPreparo = txtTempoDePreparo.text
          vop_CardapioNegocios.Preco = txtPreco.text
          If vbp_Cardapio = False Then
             If vop_CardapioNegocios.IncluirCardapio() = True Then
                MsgBox "Cardapio cadastrado com sucesso !", vbExclamation, "Cardapio"
             End If
          Else
             If vop_CardapioNegocios.AlterarCardapio() = True Then
                MsgBox "Cardapio alterado com sucesso !", vbExclamation, "Cardapio"
             End If
          End If
      Set vop_CardapioNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmCardapioLista.CarregarGrid
  
    'Inicializa entrada e saida
    Call InicializaEntradaSaida
    
   'Valida entrada de dados
   If vbp_Cardapio = True Then
      Call lvbFechar_Click
   End If
   
End Sub

Private Sub lvbFechar_Click()
    Unload Me
End Sub

'Funcoes
Function Editar(ByVal pIdCardapio As Integer) As Boolean
Dim vcl_DescricaoLocalDePreparo As String

    'Verifica uma inclusao ou alteracao do cliente
    vbp_Cardapio = True
    'Controle de exibicao
    lblCodigo.Visible = True
    
    Set vop_CardapioNegocios = New clsCardapioNegocios
        
        If vop_CardapioNegocios.PesquisarCardapio(lvwCardapio, pIdCardapio) = True Then
           lblCodigo.Caption = pIdCardapio
           txtIdCozinha.text = vop_CardapioNegocios.IdCozinha
           txtTempoDePreparo.text = vop_CardapioNegocios.TempoPreparo
           txtPreco.text = Format(vop_CardapioNegocios.Preco, "##,##0.00")
           txtDescricao.text = vop_CardapioNegocios.Descricao
           vcl_DescricaoLocalDePreparo = vop_CardapioNegocios.DescricaoCozinha
        Else
            MsgBox "Não foi possível encontrar o Cardapio !", vbCritical, "Cardapio"
        End If
          
         'Default
         cmbLocalDePreparo.text = vcl_DescricaoLocalDePreparo
          
    Set vop_CardapioNegocios = Nothing
    
    Me.Show vbModal

End Function

Function VerCampos() As Boolean
    
    If Trim$(txtDescricao.text) = Empty Then
        MsgBox "Informe a descrição do Cardapio !", vbExclamation, "Cardapio"
        If txtDescricao.text <> Empty Then txtDescricao.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(cmbLocalDePreparo.text) = Empty Then
        MsgBox "Informe a cozinha do Cardápio !", vbExclamation, "Cardapio"
        If cmbLocalDePreparo.text <> Empty Then cmbLocalDePreparo.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtTempoDePreparo.text) = Empty Or Trim$(txtTempoDePreparo.text) = "0" Then
        MsgBox "Informe o Tempo de preparo do Cardapio !", vbExclamation, "Cardapio"
        If txtTempoDePreparo.text <> Empty Then txtTempoDePreparo.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtPreco.text) = Empty Or Trim$(txtPreco.text) = "0,00" Then
        MsgBox "Informe o Preço do Cardapio !", vbExclamation, "Cardapio"
        If txtPreco.text <> Empty Then txtPreco.SetFocus
        VerCampos = False
        Exit Function
    End If
    
    
    VerCampos = True

End Function

Private Function VerNumeros() As Boolean

    If IsNumeric(txtTempoDePreparo.text) = False Then
       If txtTempoDePreparo.text <> Empty Then txtTempoDePreparo.SetFocus
       VerNumeros = False
       Exit Function
    ElseIf IsNumeric(txtPreco.text) = False Then
       If txtPreco.text <> Empty Then txtPreco.SetFocus
       VerNumeros = False
       Exit Function
    End If
       
    VerNumeros = True

End Function

Private Function DefaultCampos() As BookmarkEnum

    txtTempoDePreparo.text = "0"
    txtPreco.text = "0,00"

End Function

Private Function InicializaEntradaSaida() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos
    
    'Carrega combobox
    cmbLocalDePreparo.Clear
    Call ComboBox(cmbLocalDePreparo, "Cozinhas", "Id", "Descricao", " ORDER BY Descricao")
    If cmbLocalDePreparo.ListCount > 0 Then
       cmbLocalDePreparo.ListIndex = (cmbLocalDePreparo.ListCount - cmbLocalDePreparo.ListCount) '+ 1
    End If

End Function




