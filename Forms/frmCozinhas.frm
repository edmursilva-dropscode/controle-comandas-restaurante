VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "PRJXTAB.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCozinhas 
   Caption         =   "Cadastro de Cozinhas"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin prjXTab.XTab xtbCozinhas 
      Height          =   1380
      Left            =   75
      TabIndex        =   2
      Top             =   765
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   2434
      TabCount        =   1
      TabCaption(0)   =   "  Cozinhas "
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
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   5080
      End
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   1
         Left            =   165
         TabIndex        =   9
         Top             =   450
         Width           =   6180
         Begin VB.TextBox txtCapacidade 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1170
            MaxLength       =   3
            TabIndex        =   6
            Tag             =   "0"
            Text            =   "0"
            Top             =   405
            Width           =   720
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   1185
            MaxLength       =   40
            TabIndex        =   0
            Tag             =   "0"
            Top             =   0
            Width           =   4935
         End
         Begin VB.Label lblCapacidade 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capacidade:"
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
            Width           =   1080
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
         TabIndex        =   3
         Top             =   15
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin LVbuttons.LaVolpeButton lvbFechar 
      Height          =   360
      Left            =   5490
      TabIndex        =   8
      Top             =   2265
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
      MICON           =   "frmCozinhas.frx":0000
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
      Left            =   6765
      TabIndex        =   11
      Top             =   1065
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
   Begin LVbuttons.LaVolpeButton lblGravar 
      Height          =   360
      Left            =   4425
      TabIndex        =   7
      Top             =   2265
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
      MICON           =   "frmCozinhas.frx":001C
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
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "frmCozinhas.frx":0038
      Top             =   75
      Width           =   480
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -675
      Picture         =   "frmCozinhas.frx":0E7A
      Top             =   675
      Width           =   10740
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17955
   End
End
Attribute VB_Name = "frmCozinhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_CozinhasNegocios As New clsCozinhasNegocios
'Variaveis de controle do form
Dim vbp_Cozinha As Boolean                             'Verifica uma inclusao ou alteracao


'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Public Sub Form_Load()

    'Verifica uma inclusao ou alteracao
    vbp_Cozinha = False
    
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
   
   Set frmCozinhas = Nothing
   
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

Private Sub txtCapacidade_Change()
    
    If VerNumeros = False Then Exit Sub
    
End Sub

Private Sub txtCapacidade_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtCapacidade_KeyPress(KeyAscii As Integer)
    
    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
            
End Sub

Private Sub lblGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Cozinha = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_CozinhasNegocios = New clsCozinhasNegocios
          vop_CozinhasNegocios.IdCozinhas = lblCodigo.Caption
          vop_CozinhasNegocios.Descricao = txtDescricao.text
          vop_CozinhasNegocios.Capacidade = txtCapacidade.text
          If vbp_Cozinha = False Then
             If vop_CozinhasNegocios.IncluirCozinha() = True Then
                MsgBox "Cozinha cadastrada com sucesso !", vbExclamation, "Cozinha"
             End If
          Else
             If vop_CozinhasNegocios.AlterarCozinha() = True Then
                MsgBox "Cozinha alterada com sucesso !", vbExclamation, "Cozinha"
             End If
          End If
      Set vop_CozinhasNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmCozinhasLista.CarregarGrid
   
   'Inicializa entrada e saida
   Call InicializaEntradaSaida
   
   'Valida entrada de dados
   If vbp_Cozinha = True Then
      Call lvbFechar_Click
   End If
   
End Sub

Private Sub lvbFechar_Click()
    Unload Me
End Sub

'Funcoes
Function Editar(ByVal pIdCozinha As Integer) As Boolean
    
    'Verifica uma inclusao ou alteracao do cliente
    vbp_Cozinha = True
    'Controle de exibicao
    lblCodigo.Visible = True
    
    Set vop_CozinhasNegocios = New clsCozinhasNegocios
    
        If vop_CozinhasNegocios.PesquisarCozinha(lvwCozinhas, pIdCozinha, 0) = True Then
           lblCodigo.Caption = pIdCozinha
           txtDescricao.text = vop_CozinhasNegocios.Descricao
           txtCapacidade.text = vop_CozinhasNegocios.Capacidade
        Else
            MsgBox "Não foi possível encontrar a Cozinha !", vbCritical, "Cozinha"
        End If
          
    Set vop_CozinhasNegocios = Nothing
    
    Me.Show vbModal

End Function

Function VerCampos() As Boolean
    
    If Trim$(txtDescricao.text) = Empty Then
        MsgBox "Informe a descrição da Cozinha !", vbExclamation, "Cozinha"
        'txtDescricao.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtCapacidade.text) = Empty Or Trim$(txtCapacidade.text) = "0" Then
        MsgBox "Informe a Capacidade da Cozinha !", vbExclamation, "Cozinha"
        'txtCapacidade.SetFocus
        VerCampos = False
        Exit Function
    End If
    
    VerCampos = True

End Function

Private Function VerNumeros() As Boolean

    If IsNumeric(txtCapacidade.text) = False Then
       'txtTempoDePreparo.SetFocus
       VerNumeros = False
       Exit Function
    End If

    VerNumeros = True

End Function

Private Function DefaultCampos() As BookmarkEnum

    txtCapacidade.text = "0"

End Function

Private Function InicializaEntradaSaida() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos

End Function
