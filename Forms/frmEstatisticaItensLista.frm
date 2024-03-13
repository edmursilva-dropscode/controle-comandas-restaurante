VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmEstatisticaItensLista 
   Caption         =   "Estatisticas por Itens"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbLocalizarItemEstatistica 
      Height          =   315
      ItemData        =   "frmEstatisticaItensLista.frx":0000
      Left            =   4170
      List            =   "frmEstatisticaItensLista.frx":000D
      TabIndex        =   10
      Top             =   345
      Width           =   4995
   End
   Begin MSDataGridLib.DataGrid dtgEstatisticaItens 
      Bindings        =   "frmEstatisticaItensLista.frx":0097
      Height          =   6435
      Left            =   105
      TabIndex        =   8
      Top             =   765
      Visible         =   0   'False
      Width           =   9045
      _ExtentX        =   15954
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
      ColumnCount     =   5
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
         Caption         =   "Código"
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
         Caption         =   "Item"
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
         DataField       =   "QUANTIDADE"
         Caption         =   "      Qtde."
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
            ColumnWidth     =   2745,071
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   3704,882
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1094,74
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
      MICON           =   "frmEstatisticaItensLista.frx":00B9
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
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
      BTYPE           =   2
      TX              =   "LaVolpeButton"
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
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmEstatisticaItensLista.frx":00D5
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
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
   Begin MSAdodcLib.Adodc adoEstatisticaItens 
      Height          =   330
      Left            =   9690
      Top             =   7275
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
      RecordSource    =   $"frmEstatisticaItensLista.frx":00F1
      Caption         =   "adoEstatisticaItens"
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
   Begin MSAdodcLib.Adodc adoEstatisticaItensDias 
      Height          =   330
      Left            =   9675
      Top             =   6825
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
      RecordSource    =   $"frmEstatisticaItensLista.frx":041D
      Caption         =   "adoEstatisticaItensDias"
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
   Begin LVbuttons.LaVolpeButton lblImprimir 
      Height          =   360
      Left            =   7065
      TabIndex        =   11
      Top             =   7275
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmEstatisticaItensLista.frx":074B
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
   Begin LVbuttons.LaVolpeButton lblFechar 
      Height          =   360
      Left            =   8130
      TabIndex        =   12
      Top             =   7275
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
      MICON           =   "frmEstatisticaItensLista.frx":0767
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
   Begin MSDataGridLib.DataGrid dtgEstatisticaItensDias 
      Bindings        =   "frmEstatisticaItensLista.frx":0783
      Height          =   6435
      Left            =   105
      TabIndex        =   1
      Top             =   765
      Visible         =   0   'False
      Width           =   9045
      _ExtentX        =   15954
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "idDSe"
         Caption         =   "                           Dia "
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
         DataField       =   "nomeDSe"
         Caption         =   "                  Dia da Semana"
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
         DataField       =   "QuantidadePessoas"
         Caption         =   "       Qtde. de Pessoas Atendidas"
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
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   2910,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   2745,071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   2805,166
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   5
      Left            =   3630
      TabIndex        =   9
      Top             =   420
      Width           =   360
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
   Begin VB.Image Image2 
      Height          =   480
      Left            =   75
      Picture         =   "frmEstatisticaItensLista.frx":07A9
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   45
      Left            =   9210
      Picture         =   "frmEstatisticaItensLista.frx":1073
      Top             =   675
      Width           =   10740
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1200
      Picture         =   "frmEstatisticaItensLista.frx":19F7
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
Attribute VB_Name = "frmEstatisticaItensLista"
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

    'Exibe dados
    dtgEstatisticaItens.Visible = True
    dtgEstatisticaItensDias.Visible = False
    
    'Carrega grid
    Call CarregarGrid
    cmbLocalizarItemEstatistica.ListIndex = 0
    
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

Private Sub dtgEstatisticaItens_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If dtgEstatisticaItens.Bookmark > 0 Then
      dtgEstatisticaItens.SelBookmarks.Add dtgEstatisticaItens.Bookmark
      vip_ItemListaComanda = dtgEstatisticaItens.Columns(0).text
   End If
End Sub

Private Sub dtgEstatisticaItens_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgEstatisticaItens.SelBookmarks.Count > 0 Then
           dtgEstatisticaItens.SelBookmarks.Remove 0
        End If
        adoEstatisticaItens.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgEstatisticaItens_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgEstatisticaItens.SelBookmarks.Count > 0 Then
           dtgEstatisticaItens.SelBookmarks.Remove 0
        End If
        adoEstatisticaItens.Recordset.MovePrevious
        
    End If
   
End Sub


Private Sub dtgEstatisticaItensDias_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

   If dtgEstatisticaItensDias.Bookmark > 0 Then
      dtgEstatisticaItensDias.SelBookmarks.Add dtgEstatisticaItensDias.Bookmark
      vip_ItemListaComanda = dtgEstatisticaItensDias.Columns(0).text
   End If

End Sub

Private Sub dtgEstatisticaItensDias_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgEstatisticaItensDias.SelBookmarks.Count > 0 Then
           dtgEstatisticaItensDias.SelBookmarks.Remove 0
        End If
        adoEstatisticaItensDias.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgEstatisticaItensDias_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgEstatisticaItensDias.SelBookmarks.Count > 0 Then
           dtgEstatisticaItensDias.SelBookmarks.Remove 0
        End If
        adoEstatisticaItensDias.Recordset.MovePrevious
        
    End If
   
End Sub

Private Sub cmbLocalizarItemEstatistica_Click()

On Error GoTo TrataErros
   
   If cmbLocalizarItemEstatistica.ListIndex = 0 Then
      dtgEstatisticaItens.Visible = True
      dtgEstatisticaItensDias.Visible = False
   ElseIf cmbLocalizarItemEstatistica.ListIndex = 1 Then
      dtgEstatisticaItens.Visible = False
      dtgEstatisticaItensDias.Visible = True
   End If

   Set vop_ItemComandoNegocios = New clsComandaItemNegocios
       Call vop_ItemComandoNegocios.LocalizarItemEstatistica(adoEstatisticaItens, adoEstatisticaItensDias)
   Set vop_ItemComandoNegocios = Nothing

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
        
        vbl_Carregar = vop_ComandasNegocios.CarregarGridItensEstatisticaRS(adoEstatisticaItens, adoEstatisticaItensDias)
        If vbl_Carregar = True And dtgEstatisticaItens.Bookmark = 0 Then
           dtgEstatisticaItens.Refresh
        End If
        
    Set vop_ComandasNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
       Exit Sub
    End If

End Sub

Private Sub lblImprimir_Click()
   If cmbLocalizarItemEstatistica.ListIndex = 0 Then
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
      DataReport1.Show vbModal
   End If
End Sub

