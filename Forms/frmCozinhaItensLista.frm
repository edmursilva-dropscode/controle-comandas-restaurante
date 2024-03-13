VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "PRJXTAB.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCozinhaItensLista 
   Caption         =   "Itens da Cozinha"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   18615
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin prjXTab.XTab xtbDevolucao 
      Height          =   6945
      Left            =   75
      TabIndex        =   1
      Top             =   795
      Width           =   18480
      _ExtentX        =   32597
      _ExtentY        =   12250
      TabCount        =   1
      TabCaption(0)   =   "  Comandas "
      TabContCtrlCnt(0)=   5
      Tab(0)ContCtrlCap(1)=   "cmbLocalizarItemCozinha"
      Tab(0)ContCtrlCap(2)=   "txtLocalizarItemCozinha"
      Tab(0)ContCtrlCap(3)=   "dtgCozinhaItens"
      Tab(0)ContCtrlCap(4)=   "lblTitulo5"
      Tab(0)ContCtrlCap(5)=   "lblTitulo4"
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
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
      Begin VB.ComboBox cmbLocalizarItemCozinha 
         Height          =   315
         ItemData        =   "frmCozinhaItensLista.frx":0000
         Left            =   17295
         List            =   "frmCozinhaItensLista.frx":000D
         TabIndex        =   13
         Top             =   -15
         Width           =   1200
      End
      Begin VB.TextBox txtLocalizarItemCozinha 
         Height          =   315
         Left            =   14160
         MaxLength       =   50
         TabIndex        =   11
         Top             =   -15
         Width           =   2475
      End
      Begin MSDataGridLib.DataGrid dtgCozinhaItens 
         Bindings        =   "frmCozinhaItensLista.frx":0028
         Height          =   6435
         Left            =   105
         TabIndex        =   9
         Top             =   390
         Width           =   18255
         _ExtentX        =   32200
         _ExtentY        =   11351
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "IdCozinha"
            Caption         =   "IdCozinha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "DescricaoCozinha"
            Caption         =   "Cozinha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "IdCardapio"
            Caption         =   "IdCardapio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "DescricaoItem"
            Caption         =   "Descrição do item"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Quantidade"
            Caption         =   "   Qtde."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Id"
            Caption         =   "Id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "IdComanda"
            Caption         =   "IdComanda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "IdTipoComanda"
            Caption         =   "   Comanda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "DataConfirmacaoPreparo"
            Caption         =   "   Data/Hora de pedido do item"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "DataPrevistaPreparo"
            Caption         =   "   Data/Hora prevista do preparo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "DataFinalizacaoPreparo"
            Caption         =   "  Data/Hora do termino do preparo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Preco"
            Caption         =   "Preco"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "TotalPreco"
            Caption         =   "TotalPreco"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "StatusItem"
            Caption         =   "StatusItem"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "StatusItemDescricao"
            Caption         =   "                       Status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   2039,811
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   3435,024
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2505,26
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2550,047
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2640,189
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column14 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2745,071
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton lvbRemoverDevolucao 
         Height          =   360
         Left            =   -68070
         TabIndex        =   5
         Top             =   2955
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "&Remover Produto"
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
         MICON           =   "frmCozinhaItensLista.frx":0046
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
      Begin LVbuttons.LaVolpeButton lvbEditarDevolucao 
         Height          =   360
         Left            =   -69720
         TabIndex        =   4
         Top             =   2955
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "&Editar Devolução"
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
         MICON           =   "frmCozinhaItensLista.frx":0062
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
      Begin VB.TextBox Text1 
         Height          =   2905
         Left            =   -74920
         MaxLength       =   700
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   400
         Width           =   8505
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74280
         MaxLength       =   6
         TabIndex        =   2
         Top             =   760
         WhatsThisHelpID =   1067
         Width           =   990
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   5
         Left            =   16830
         TabIndex        =   12
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localizar:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   13320
         TabIndex        =   10
         Top             =   60
         Width           =   675
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Áreas:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   -74850
         TabIndex        =   7
         Top             =   810
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -73260
         TabIndex        =   6
         Top             =   760
         Width           =   4720
      End
   End
   Begin LVbuttons.LaVolpeButton lblFechar 
      Height          =   360
      Left            =   17535
      TabIndex        =   8
      Top             =   7830
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
      MICON           =   "frmCozinhaItensLista.frx":007E
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
   Begin MSAdodcLib.Adodc adoCozinhaItens 
      Height          =   330
      Left            =   18690
      Top             =   7260
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TESTE_VB6"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TESTE_VB6"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmCozinhaItensLista.frx":009A
      Caption         =   "adoCozinhaItens"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image imgIcone 
      Height          =   405
      Left            =   90
      Picture         =   "frmCozinhaItensLista.frx":0661
      Top             =   135
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   45
      Left            =   9210
      Picture         =   "frmCozinhaItensLista.frx":0E77
      Top             =   675
      Width           =   10740
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1200
      Picture         =   "frmCozinhaItensLista.frx":17FB
      Top             =   675
      Width           =   10740
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18930
   End
End
Attribute VB_Name = "frmCozinhaItensLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_ComandasNegocios As New clsComandasNegocios
Dim vop_ItemComandoNegocios As New clsComandaItemNegocios
'Variaveis de controle do form
Dim vil_IdComanda As Long              'Identificador da Comandas
Dim vip_ItemListaComanda As Long       'Identificador do Item da Comanda




'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Private Sub Form_Load()

    Call CarregarGrid
    cmbLocalizarItemCozinha.ListIndex = 0
    
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
   
   Set frmCozinhaItensLista = Nothing
   
End Sub

Private Sub lblFechar_Click()
   Unload Me
End Sub

Private Sub dtgCozinhaItens_DblClick()
On Error GoTo TrataErros

   Call frmComandaItens.Form_Load
   Call frmComandaItens.Editar(Me, vip_ItemListaComanda)

TrataErros:
    If Err.Number <> 0 Then Exit Sub
End Sub

Private Sub dtgCozinhaItens_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgCozinhaItens.SelBookmarks.Count > 0 Then
           dtgCozinhaItens.SelBookmarks.Remove 0
        End If
        adoCozinhaItens.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgCozinhaItens_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgCozinhaItens.SelBookmarks.Count > 0 Then
           dtgCozinhaItens.SelBookmarks.Remove 0
        End If
        adoCozinhaItens.Recordset.MovePrevious
        
    End If
   
End Sub

Private Sub dtgCozinhaItens_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If dtgCozinhaItens.Bookmark > 0 Then
      dtgCozinhaItens.SelBookmarks.Add dtgCozinhaItens.Bookmark
      vip_ItemListaComanda = dtgCozinhaItens.Columns(5).text
   End If
End Sub

Private Sub txtLocalizarItemCozinha_Change()

On Error GoTo TrataErros

   Set vop_ItemComandoNegocios = New clsComandaItemNegocios
       Call vop_ItemComandoNegocios.LocalizarItemCozinha(adoCozinhaItens, txtLocalizarItemCozinha.text, IIf(cmbLocalizarItemCozinha.ListIndex < 0, 0, cmbLocalizarItemCozinha.ListIndex))
   Set vop_ItemComandoNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ItemComandoNegocios = Nothing
       Exit Sub
    End If

End Sub

Private Sub cmbLocalizarItemCozinha_Click()

On Error GoTo TrataErros

   Set vop_ItemComandoNegocios = New clsComandaItemNegocios
       Call vop_ItemComandoNegocios.LocalizarItemCozinha(adoCozinhaItens, txtLocalizarItemCozinha.text, IIf(cmbLocalizarItemCozinha.ListIndex < 0, 0, cmbLocalizarItemCozinha.ListIndex))
   Set vop_ItemComandoNegocios = Nothing
    
   txtLocalizarItemCozinha.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ItemComandoNegocios = Nothing
       Exit Sub
    End If

End Sub






'Medotos
Public Sub CarregarGrid()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_ComandasNegocios = New clsComandasNegocios
        
        vbl_Carregar = vop_ComandasNegocios.CarregarGridItensCozinhaRS(adoCozinhaItens, txtLocalizarItemCozinha.text, IIf(cmbLocalizarItemCozinha.ListIndex < 0, 0, cmbLocalizarItemCozinha.ListIndex))
        If vbl_Carregar = True And dtgCozinhaItens.Bookmark = 0 Then
           dtgCozinhaItens.Refresh
        End If
        
    Set vop_ComandasNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
       Exit Sub
    End If

End Sub



