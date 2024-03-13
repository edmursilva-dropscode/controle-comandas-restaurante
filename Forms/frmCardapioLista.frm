VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCardapioLista 
   Caption         =   "Cadastro de Cadápio"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbLocalizar 
      Height          =   315
      ItemData        =   "frmCardapioLista.frx":0000
      Left            =   4440
      List            =   "frmCardapioLista.frx":000D
      TabIndex        =   4
      Top             =   825
      Width           =   1245
   End
   Begin VB.TextBox txtLocalizar 
      Height          =   315
      Left            =   855
      MaxLength       =   50
      TabIndex        =   2
      Top             =   825
      Width           =   3045
   End
   Begin MSDataGridLib.DataGrid dtgCardapio 
      Bindings        =   "frmCardapioLista.frx":0028
      Height          =   3345
      Left            =   135
      TabIndex        =   5
      Top             =   1185
      Width           =   8970
      _ExtentX        =   15822
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Id"
         Caption         =   "  Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "IdCozinha"
         Caption         =   "IdCozinha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "TempoPreparo"
         Caption         =   "  Tempo de preparo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Preco"
         Caption         =   "       Preço"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
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
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   3044,977
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1560,189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1244,976
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton lblNovo 
      Height          =   360
      Left            =   5970
      TabIndex        =   6
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
      MICON           =   "frmCardapioLista.frx":0042
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
      Left            =   7035
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
      MICON           =   "frmCardapioLista.frx":005E
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
      Left            =   8085
      TabIndex        =   8
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
      MICON           =   "frmCardapioLista.frx":007A
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
   Begin MSAdodcLib.Adodc adoCardapio 
      Height          =   330
      Left            =   9420
      Top             =   4200
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   $"frmCardapioLista.frx":0096
      Caption         =   "adoCardapio"
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
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmCardapioLista.frx":01A5
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   3
      Top             =   900
      Width           =   360
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   1
      Top             =   900
      Width           =   675
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -840
      Picture         =   "frmCardapioLista.frx":0A6F
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
      Width           =   11370
   End
End
Attribute VB_Name = "frmCardapioLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_CardapioNegocios As New clsCardapioNegocios
'Variaveis de controle do form
Dim vil_IdCardapio As Long              'Identificador do Cardapio




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
    
    Set frmCardapioLista = Nothing
    
End Sub

Private Sub lblExcluir_Click()
    If vil_IdCardapio = 0 Then Exit Sub
    
    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_CardapioNegocios = New clsCardapioNegocios
          vop_CardapioNegocios.IdCardapio = vil_IdCardapio
          If vop_CardapioNegocios.ExcluirCardapio() = True Then
             txtLocalizar.text = Empty
             Call CarregarGrid
          End If
      Set vop_CardapioNegocios = Nothing
    End If
End Sub

Private Sub lblNovo_Click()
    frmCardapio.Show vbModal
End Sub

Private Sub lblFechar_Click()
   Unload Me
End Sub

Private Sub cmbLocalizar_Click()
   
On Error GoTo TrataErros

   Set vop_CardapioNegocios = New clsCardapioNegocios
       Call vop_CardapioNegocios.LocalizarCardapio(adoCardapio, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_CardapioNegocios = Nothing
    
   txtLocalizar.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_CardapioNegocios = Nothing
       Exit Sub
    End If
   
End Sub

Private Sub dtgCardapio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If dtgCardapio.Bookmark > 0 Then
   
      dtgCardapio.SelBookmarks.Add dtgCardapio.Bookmark
      vil_IdCardapio = dtgCardapio.Columns(0).text
   
   End If
   
End Sub

Private Sub dtgCardapio_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgCardapio.SelBookmarks.Count > 0 Then
           dtgCardapio.SelBookmarks.Remove 0
        End If
        adoCardapio.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgCardapio_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgCardapio.SelBookmarks.Count > 0 Then
           dtgCardapio.SelBookmarks.Remove 0
        End If
        adoCardapio.Recordset.MovePrevious
        
    End If
   
End Sub

Private Sub dtgCardapio_DblClick()
On Error GoTo TrataErros

   Call frmCardapio.Form_Load
   Call frmCardapio.Editar(vil_IdCardapio)

TrataErros:
    If Err.Number <> 0 Then Exit Sub
    
End Sub

Private Sub txtLocalizar_Change()
  
On Error GoTo TrataErros

   Set vop_CardapioNegocios = New clsCardapioNegocios
       Call vop_CardapioNegocios.LocalizarCardapio(adoCardapio, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_CardapioNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_CardapioNegocios = Nothing
       Exit Sub
    End If
    
End Sub

'Metodos
Public Sub CarregarGrid()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_CardapioNegocios = New clsCardapioNegocios
        
        vbl_Carregar = vop_CardapioNegocios.CarregarGridCardapioRS(adoCardapio, cmbLocalizar.ListIndex)
        If vbl_Carregar = True And adoCardapio.MaxRecords > 0 Then
           dtgCardapio.Refresh
           lblNovo.Left = 7035
           lblExcluir.Visible = False
        Else
           lblNovo.Left = 5970
           lblExcluir.Left = 7035
           lblExcluir.Visible = True
        End If
        
    Set vop_CardapioNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_CardapioNegocios = Nothing
       Exit Sub
    End If

End Sub




