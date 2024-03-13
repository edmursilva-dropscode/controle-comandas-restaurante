VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCozinhasLista 
   Caption         =   "Cadastro de Cozinhas"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbLocalizar 
      Height          =   315
      ItemData        =   "frmCozinhasLista.frx":0000
      Left            =   3855
      List            =   "frmCozinhasLista.frx":000A
      TabIndex        =   3
      Top             =   795
      Width           =   1245
   End
   Begin VB.TextBox txtLocalizar 
      Height          =   315
      Left            =   885
      MaxLength       =   50
      TabIndex        =   1
      Top             =   795
      Width           =   2475
   End
   Begin MSDataGridLib.DataGrid dtgCozinhas 
      Bindings        =   "frmCozinhasLista.frx":0021
      Height          =   3345
      Left            =   120
      TabIndex        =   4
      Top             =   1185
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
         Caption         =   "   Código"
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
         Caption         =   "Descricao"
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
         DataField       =   "Capacidade"
         Caption         =   "  Capacidade"
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
            ColumnWidth     =   2340,284
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1110,047
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton lblFechar 
      Height          =   360
      Left            =   4065
      TabIndex        =   7
      Top             =   4620
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
      MICON           =   "frmCozinhasLista.frx":003B
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
      Left            =   3000
      TabIndex        =   6
      Top             =   4620
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
      MICON           =   "frmCozinhasLista.frx":0057
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
      Left            =   1950
      TabIndex        =   5
      Top             =   4620
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
      MICON           =   "frmCozinhasLista.frx":0073
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
   Begin MSAdodcLib.Adodc adoCozinhas 
      Height          =   330
      Left            =   5340
      Top             =   4380
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
      RecordSource    =   "SELECT Id, Descricao, Capacidade FROM Cozinhas ORDER BY Id"
      Caption         =   "adoCozinhas"
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
      Left            =   3435
      TabIndex        =   2
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
      TabIndex        =   0
      Top             =   870
      Width           =   675
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1335
      Picture         =   "frmCozinhasLista.frx":008F
      Top             =   630
      Width           =   10740
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "frmCozinhasLista.frx":0A13
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7515
   End
End
Attribute VB_Name = "frmCozinhasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_CozinhasNegocios As New clsCozinhasNegocios
'Variaveis de controle do form
Dim vil_IdCozinha As Long               'Identificador da Cozinha



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

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmCozinhasLista = Nothing
   
End Sub

Private Sub lblExcluir_Click()
    If vil_IdCozinha = 0 Then Exit Sub
    
    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_CozinhasNegocios = New clsCozinhasNegocios
          vop_CozinhasNegocios.IdCozinhas = vil_IdCozinha
          If vop_CozinhasNegocios.ExcluirCozinha() = True Then
             txtLocalizar.text = Empty
             Call CarregarGrid
          End If
      Set vop_CozinhasNegocios = Nothing
    End If
End Sub

Private Sub lblNovo_Click()
    frmCozinhas.Show vbModal
End Sub

Private Sub lblFechar_Click()
   Unload Me
End Sub

Private Sub cmbLocalizar_Click()
   
On Error GoTo TrataErros

   Set vop_CozinhasNegocios = New clsCozinhasNegocios
       Call vop_CozinhasNegocios.LocalizarCozinha(adoCozinhas, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_CozinhasNegocios = Nothing
    
   txtLocalizar.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_CozinhasNegocios = Nothing
       Exit Sub
    End If
   
End Sub

Private Sub dtgCozinhas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If dtgCozinhas.Bookmark > 0 Then
      dtgCozinhas.SelBookmarks.Add dtgCozinhas.Bookmark
      vil_IdCozinha = dtgCozinhas.Columns(0).text
   End If
End Sub

Private Sub dtgCozinhas_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgCozinhas.SelBookmarks.Count > 0 Then
           dtgCozinhas.SelBookmarks.Remove 0
        End If
        adoCozinhas.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgCozinhas_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgCozinhas.SelBookmarks.Count > 0 Then
           dtgCozinhas.SelBookmarks.Remove 0
        End If
        adoCozinhas.Recordset.MovePrevious
        
    End If
   
End Sub

Private Sub dtgCozinhas_DblClick()
On Error GoTo TrataErros

   Call frmCozinhas.Form_Load
   Call frmCozinhas.Editar(vil_IdCozinha)

TrataErros:
    If Err.Number <> 0 Then Exit Sub
    
End Sub

Private Sub txtLocalizar_Change()
  
On Error GoTo TrataErros

   Set vop_CozinhasNegocios = New clsCozinhasNegocios
       Call vop_CozinhasNegocios.LocalizarCozinha(adoCozinhas, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_CozinhasNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_CozinhasNegocios = Nothing
       Exit Sub
    End If
    
End Sub

'Medotos
Public Sub CarregarGrid()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_CozinhasNegocios = New clsCozinhasNegocios
        
        vbl_Carregar = vop_CozinhasNegocios.CarregarGridCozinhaRS(adoCozinhas, cmbLocalizar.ListIndex)
        If vbl_Carregar = True And adoCozinhas.MaxRecords > 0 Then
           dtgCozinhas.Refresh
           lblNovo.Left = 3000
           lblExcluir.Visible = False
        Else
           lblNovo.Left = 1950
           lblExcluir.Left = 3000
           lblExcluir.Visible = True
        End If
        
    Set vop_CozinhasNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_CozinhasNegocios = Nothing
       Exit Sub
    End If

End Sub



