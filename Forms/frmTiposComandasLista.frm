VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmTiposComandasLista 
   Caption         =   "Cadastro de Comandas"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbLocalizar 
      Height          =   315
      ItemData        =   "frmTiposComandasLista.frx":0000
      Left            =   3855
      List            =   "frmTiposComandasLista.frx":000A
      TabIndex        =   4
      Top             =   810
      Width           =   1245
   End
   Begin VB.TextBox txtLocalizar 
      Height          =   315
      Left            =   855
      MaxLength       =   50
      TabIndex        =   2
      Top             =   810
      Width           =   2475
   End
   Begin MSDataGridLib.DataGrid dtgTiposComandas 
      Bindings        =   "frmTiposComandasLista.frx":0021
      Height          =   3345
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   5900
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
         DataField       =   "Id"
         Caption         =   "  Código"
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
         DataField       =   "Descricao"
         Caption         =   "Descrição"
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
         DataField       =   "StatusComanda"
         Caption         =   "StatusComanda"
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
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3449,764
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1140,095
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton lblFechar 
      Height          =   360
      Left            =   4050
      TabIndex        =   6
      Top             =   4635
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
      MICON           =   "frmTiposComandasLista.frx":0040
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
   Begin LVbuttons.LaVolpeButton lblExcluir 
      Height          =   360
      Left            =   2985
      TabIndex        =   7
      Top             =   4635
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmTiposComandasLista.frx":005C
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
   Begin LVbuttons.LaVolpeButton lblNovo 
      Height          =   360
      Left            =   1905
      TabIndex        =   8
      Top             =   4635
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmTiposComandasLista.frx":0078
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
   Begin MSAdodcLib.Adodc adoTiposComandas 
      Height          =   330
      Left            =   5340
      Top             =   4185
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      RecordSource    =   $"frmTiposComandasLista.frx":0094
      Caption         =   "adoTiposComandas"
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
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   1
      Left            =   3405
      TabIndex        =   3
      Top             =   885
      Width           =   360
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   855
      Width           =   675
   End
   Begin VB.Image imgIcone 
      Height          =   480
      Left            =   105
      Picture         =   "frmTiposComandasLista.frx":0175
      Top             =   105
      Width           =   480
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -2010
      Picture         =   "frmTiposComandasLista.frx":0A3F
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
      Width           =   7515
   End
End
Attribute VB_Name = "frmTiposComandasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_TiposComandasNegocios As New clsTiposComandasNegocios
'Variaveis de controle do form
Dim vil_IdComanda As Long               'Identificador da Comanda



'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Private Sub Form_Load()

    Call CarregarGrid
    cmbLocalizar.ListIndex = 0
    
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
   
   Set frmTiposComandasLista = Nothing
   
End Sub

Private Sub lblExcluir_Click()
    If vil_IdComanda = 0 Then Exit Sub
    
    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_TiposComandasNegocios = New clsTiposComandasNegocios
          vop_TiposComandasNegocios.IdComanda = vil_IdComanda
          If vop_TiposComandasNegocios.ExcluirTipoComanda() = True Then
             txtLocalizar.text = Empty
             Call CarregarGrid
          End If
      Set vop_TiposComandasNegocios = Nothing
    End If
End Sub

Private Sub lblNovo_Click()
    frmTiposComandas.Show vbModal
End Sub

Private Sub lblFechar_Click()
   Unload Me
End Sub

Private Sub cmbLocalizar_Click()
   
On Error GoTo TrataErros

   Set vop_TiposComandasNegocios = New clsTiposComandasNegocios
       Call vop_TiposComandasNegocios.LocalizarTipoComanda(adoTiposComandas, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_TiposComandasNegocios = Nothing
    
   txtLocalizar.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_TiposComandasNegocios = Nothing
       Exit Sub
    End If
   
End Sub

Private Sub dtgTiposComandas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If dtgTiposComandas.Bookmark > 0 Then
      dtgTiposComandas.SelBookmarks.Add dtgTiposComandas.Bookmark
      vil_IdComanda = dtgTiposComandas.Columns(0).text
   End If
End Sub

Private Sub dtgTiposComandas_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgTiposComandas.SelBookmarks.Count > 0 Then
           dtgTiposComandas.SelBookmarks.Remove 0
        End If
        adoTiposComandas.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgTiposComandas_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgTiposComandas.SelBookmarks.Count > 0 Then
           dtgTiposComandas.SelBookmarks.Remove 0
        End If
        adoTiposComandas.Recordset.MovePrevious
        
    End If
   
End Sub

Private Sub dtgTiposComandas_DblClick()
On Error GoTo TrataErros

   Call frmTiposComandas.Form_Load
   Call frmTiposComandas.Editar(vil_IdComanda)

TrataErros:
    If Err.Number <> 0 Then Exit Sub
    
End Sub

Private Sub txtLocalizar_Change()
  
On Error GoTo TrataErros

   Set vop_TiposComandasNegocios = New clsTiposComandasNegocios
       Call vop_TiposComandasNegocios.LocalizarTipoComanda(adoTiposComandas, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_TiposComandasNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_TiposComandasNegocios = Nothing
       Exit Sub
    End If
    
End Sub

'Medotos
Public Sub CarregarGrid()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_TiposComandasNegocios = New clsTiposComandasNegocios
        
        vbl_Carregar = vop_TiposComandasNegocios.CarregarGridTipoComandaRS(adoTiposComandas, cmbLocalizar.ListIndex)
        If vbl_Carregar = True And adoTiposComandas.MaxRecords > 0 Then
           dtgTiposComandas.Refresh
           lblNovo.Left = 3000
           lblExcluir.Visible = False
        Else
           lblNovo.Left = 1950
           lblExcluir.Left = 3000
           lblExcluir.Visible = True
        End If
        
    Set vop_TiposComandasNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_TiposComandasNegocios = Nothing
       Exit Sub
    End If

End Sub




