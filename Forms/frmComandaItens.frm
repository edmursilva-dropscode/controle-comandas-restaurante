VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "PRJXTAB.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmComandaItens 
   Caption         =   "Cadastro de Itens da Comanda"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtIdTipoComanda 
      Height          =   285
      Left            =   2535
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   195
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStatusItemDescricao 
      Height          =   285
      Left            =   3285
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStatusItem 
      Height          =   285
      Left            =   3015
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtIdCardapio 
      Height          =   285
      Left            =   2745
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtItemComanda 
      Height          =   285
      Left            =   2040
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtIdComanda 
      Height          =   285
      Left            =   2250
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin prjXTab.XTab xtbComandas 
      Height          =   3165
      Left            =   105
      TabIndex        =   9
      Top             =   780
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5583
      TabCount        =   1
      TabCaption(0)   =   "  Item  "
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
         Height          =   2805
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   6585
         Begin VB.ComboBox cmbItem 
            Height          =   315
            ItemData        =   "frmComandaItens.frx":0000
            Left            =   2025
            List            =   "frmComandaItens.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   4575
         End
         Begin VB.TextBox txtQuantidade 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2025
            MaxLength       =   3
            TabIndex        =   1
            Tag             =   "0"
            Text            =   "0"
            Top             =   660
            Width           =   720
         End
         Begin VB.Label lblPrecoItemComanda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2025
            TabIndex        =   24
            Top             =   1080
            Width           =   1605
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4410
            TabIndex        =   2
            Top             =   1080
            Width           =   1605
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
            Index           =   0
            Left            =   3780
            TabIndex        =   18
            Top             =   1140
            Width           =   555
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item:"
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
            Left            =   30
            TabIndex        =   17
            Top             =   315
            Width           =   480
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
            Index           =   6
            Left            =   45
            TabIndex        =   15
            Top             =   1140
            Width           =   555
         End
         Begin VB.Label lblQuantidade 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
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
            Left            =   30
            TabIndex        =   14
            Top             =   735
            Width           =   1050
         End
         Begin VB.Label lblDataHora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data/Hora:"
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
            Left            =   45
            TabIndex        =   13
            Top             =   1545
            Width           =   960
         End
         Begin VB.Label lblDataHora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data/Hora prevista:"
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
            Left            =   60
            TabIndex        =   12
            Top             =   1950
            Width           =   1710
         End
         Begin VB.Label lblDataHora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data/Hora finalização:"
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
            Index           =   2
            Left            =   45
            TabIndex        =   11
            Top             =   2400
            Width           =   1920
         End
         Begin VB.Label lblDataHoraPrevista 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2025
            TabIndex        =   4
            Top             =   1920
            Width           =   2025
         End
         Begin VB.Label lblDataHoraFinalizacao 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2025
            TabIndex        =   5
            Top             =   2340
            Width           =   2025
         End
         Begin VB.Label lblDataHoraComanda 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   1485
            Width           =   2025
         End
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda:  0"
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
         Left            =   870
         TabIndex        =   16
         Top             =   15
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin LVbuttons.LaVolpeButton lblGravar 
      Height          =   360
      Left            =   4965
      TabIndex        =   6
      Top             =   4050
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
      MICON           =   "frmComandaItens.frx":0004
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
   Begin LVbuttons.LaVolpeButton lvbFechar 
      Height          =   360
      Left            =   6015
      TabIndex        =   7
      Top             =   4050
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
      MICON           =   "frmComandaItens.frx":0020
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
   Begin MSComctlLib.ListView lvwItemComandaa 
      Height          =   360
      Left            =   7245
      TabIndex        =   21
      Top             =   1140
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
      NumItems        =   12
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
         Text            =   "IdCardapio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DescricaoItem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Preco"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "TotalPreco"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "DataConfirmacaoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "DataPrevistaPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "DataFinalizacaoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "StatusItem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "StatusItemDescricao"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCardapio 
      Height          =   360
      Left            =   7245
      TabIndex        =   23
      Top             =   1635
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
   Begin MSComctlLib.ListView lvwItemComanda 
      Height          =   375
      Left            =   8640
      TabIndex        =   27
      Top             =   1125
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IdComanda"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IdTipoComanda"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IdCardapio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Item"
         Object.Width           =   5134
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Qtde."
         Object.Width           =   1219
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Preço"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Data/Hora"
         Object.Width           =   3070
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Data/Hora Prevista"
         Object.Width           =   3070
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Data/Hora Finalização"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "StatusItem"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Status"
         Object.Width           =   4286
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton lblCancelar 
      Height          =   360
      Left            =   3915
      TabIndex        =   29
      Top             =   4050
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmComandaItens.frx":003C
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
   Begin LVbuttons.LaVolpeButton lblFinalizar 
      Height          =   360
      Left            =   2865
      TabIndex        =   30
      Top             =   4050
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Finalizar"
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
      MICON           =   "frmComandaItens.frx":0058
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
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1785
      Picture         =   "frmComandaItens.frx":0074
      Top             =   675
      Width           =   10740
   End
   Begin VB.Image imgIcone 
      Height          =   480
      Left            =   90
      Picture         =   "frmComandaItens.frx":09F8
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   17955
   End
End
Attribute VB_Name = "frmComandaItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_ItemComandoNegocios As New clsComandaItemNegocios
Dim vop_CardapioNegocios As New clsCardapioNegocios
'Variaveis de controle do form
Dim vbp_ItemComanda As Boolean                             'Verifica uma inclusao ou alteracao
Dim vip_ItemListaComanda As Integer


'Eventos
Private Sub Form_Activate()

   Me.Refresh
   
End Sub

Public Sub Form_Load()

    'Verifica uma inclusao ou alteracao
    vbp_ItemComanda = False
    
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

   Set frmComandaItens = Nothing
   
End Sub

Private Sub lblCancelar_Click()

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   If MsgBox("Confirma o cancelamento ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_ItemComandoNegocios = New clsComandaItemNegocios
          vop_ItemComandoNegocios.IdItemComanda = txtItemComanda.text
          vop_ItemComandoNegocios.StatusItem = 5
          If vop_ItemComandoNegocios.CancelarFinalizarItensComanda = True Then
             frmComandasLista.CarregarGrid
             MsgBox "Item da Comanda cancelado com sucesso !", vbExclamation, "Item da Comanda"
          End If
      Set vop_ItemComandoNegocios = Nothing
      
      'Inicializa entrada e saida
      Call InicializaEntradaSaida
      
      'Fechar form
      Call lvbFechar_Click
      
   End If
      
End Sub

Private Sub lblFinalizar_Click()

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   If Trim$(lblDataHoraPrevista.Caption) = "????" Then
      MsgBox "Não há data prevista de preparo do Item !", vbExclamation, "Item da Comanda"
      Exit Sub
   End If
   
   If MsgBox("Confirma a finalização ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_ItemComandoNegocios = New clsComandaItemNegocios
          vop_ItemComandoNegocios.IdItemComanda = txtItemComanda.text
          vop_ItemComandoNegocios.StatusItem = 4
          If vop_ItemComandoNegocios.CancelarFinalizarItensComanda = True Then
             frmCozinhaItensLista.CarregarGrid
             MsgBox "Item da Comanda finalizado com sucesso !", vbExclamation, "Item da Comanda"
          End If
      Set vop_ItemComandoNegocios = Nothing
      
      'Inicializa entrada e saida
      Call InicializaEntradaSaida
      
      'Fechar form
      Call lvbFechar_Click
      
   End If

End Sub

Private Sub lblGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_ItemComanda = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      
      If vbp_ItemComanda = False Then
         lvwItemComanda.ListItems.Add , , txtItemComanda.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(1) = txtIdComanda.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(2) = txtIdTipoComanda.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(3) = txtIdCardapio.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(4) = cmbItem.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(5) = txtQuantidade.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(6) = lblPrecoItemComanda.Caption
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(7) = lblTotal.Caption
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(8) = lblDataHoraComanda.Caption
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(9) = lblDataHoraPrevista.Caption
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(10) = lblDataHoraFinalizacao
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(11) = txtStatusItem.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(12) = txtStatusItemDescricao.text
         If IncluirItensComanda = True Then
            MsgBox "Item da Comanda cadastrado com sucesso !", vbExclamation, "Item da Comanda"
         End If
      Else
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).text = txtItemComanda.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(1) = txtIdComanda.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(2) = txtIdTipoComanda.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(3) = txtIdCardapio.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(4) = cmbItem.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(5) = txtQuantidade.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(6) = lblPrecoItemComanda.Caption
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(7) = lblTotal.Caption
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(8) = VerDataHoraAtual()
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(9) = lblDataHoraPrevista.Caption
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(10) = lblDataHoraFinalizacao
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(11) = txtStatusItem.text
         lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(12) = txtStatusItemDescricao.text
         If AlterarItensComanda = True Then
            MsgBox "Item da Comanda alterado com sucesso !", vbExclamation, "Item da Comanda"
         End If
      End If
      
      'Inicializa entrada e saida
      Call InicializaEntradaSaida
      
      'Valida processo de inclusao e alteração
      If vbp_ItemComanda = True Then
         Call lvbFechar_Click
      End If
      
   End If

End Sub

Private Sub lvbFechar_Click()
   Unload Me
End Sub

Private Sub cmbItem_Click()

    Set vop_CardapioNegocios = New clsCardapioNegocios
    
        If vop_CardapioNegocios.PesquisarCardapioItemComanda(lvwCardapio, cmbItem.text, 1) = True Then
           txtIdCardapio.text = vop_CardapioNegocios.IdCardapio
           lblPrecoItemComanda.Caption = Format(vop_CardapioNegocios.Preco, "##0.00")
        Else
           MsgBox "Não foi possível encontrar o Item da Comanda !", vbCritical, "Item da Comanda"
        End If
        
    Set vop_CardapioNegocios = Nothing

End Sub

Private Sub cmbItem_LostFocus()
    
    Set vop_CardapioNegocios = New clsCardapioNegocios
    
        If vop_CardapioNegocios.PesquisarCardapioItemComanda(lvwCardapio, cmbItem.text, 1) = True Then
           txtIdCardapio.text = vop_CardapioNegocios.IdCardapio
           lblPrecoItemComanda.Caption = Format(vop_CardapioNegocios.Preco, "##0.00")
        Else
           MsgBox "Não foi possível encontrar o Item da Comanda !", vbCritical, "Item da Comanda"
        End If
                
    Set vop_CardapioNegocios = Nothing
    
End Sub

Private Sub txtQuantidade_Change()

   If VerNumeros = False Then Exit Sub

End Sub

Private Sub txtQuantidade_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    
    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
            
End Sub

Private Sub txtQuantidade_LostFocus()

    If lblPrecoItemComanda.Caption <> Empty And txtQuantidade.text <> Empty Then
       lblTotal.Caption = Format(CDbl(lblPrecoItemComanda.Caption) * Int(txtQuantidade.text), "##0.00")
       End If

End Sub



'Metodos




'Funcoes
Function VerCampos() As Boolean
Dim vil_CountLista As Integer
Dim vbl_ItemExistente As Boolean
    
    If Trim$(cmbItem.text) = Empty Then
        MsgBox "Informe o Item da Comanda !", vbExclamation, "Item da Comanda"
        If cmbItem.text <> Empty Then cmbItem.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtQuantidade.text) = Empty Or Trim$(txtQuantidade.text) = "0" Then
        MsgBox "Informe a Quantidade do Item da Comanda !", vbExclamation, "Item da Comanda"
        If txtQuantidade.text <> Empty Then txtQuantidade.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtIdCardapio.text) = Empty Or Trim$(txtIdCardapio.text) = "0" Then
        MsgBox "Informe o Itema da Comanda !", vbExclamation, "Item da Comanda"
        If cmbItem.text <> Empty Then cmbItem.SetFocus
        VerCampos = False
        Exit Function
    Else
       If vbp_ItemComanda = False Then
          vbl_ItemExistente = False
          For vil_CountLista = 1 To frmComandas.lvwItensComanda.ListItems.Count
              If frmComandas.lvwItensComanda.ListItems(vil_CountLista).SubItems(3) = txtIdCardapio.text Then
                 If Int(frmComandas.lvwItensComanda.ListItems(vil_CountLista).SubItems(11)) <> 5 Then
                 vbl_ItemExistente = True
                 Exit For
                 End If
              End If
          Next vil_CountLista
          If vbl_ItemExistente = True Then
             MsgBox "Este item já existe na lista da Comanda !", vbExclamation, "Item da Comanda"
             If cmbItem.text <> Empty Then cmbItem.SetFocus
             VerCampos = False
             Exit Function
          End If
       End If
    End If
       
    VerCampos = True

End Function

Private Function VerNumeros() As Boolean

    If IsNumeric(txtQuantidade.text) = False Then
       If txtQuantidade.text <> Empty Then txtQuantidade.SetFocus
       VerNumeros = False
       Exit Function
    End If
       
    VerNumeros = True

End Function

Private Function DefaultCampos() As BookmarkEnum
    
    lblCodigo.Caption = Empty
    txtItemComanda.text = "0"
    txtIdComanda.text = frmComandas.lblIdComanda.Caption
    txtIdTipoComanda.text = frmComandas.txtIdTipoComanda.text
    txtIdCardapio.text = "0"
    txtQuantidade.text = "0"
    lblPrecoItemComanda.Caption = "0,00"
    lblTotal.Caption = "0,00"
    lblDataHoraComanda.Caption = VerDataHoraAtual()
    lblDataHoraPrevista.Caption = "????"
    lblDataHoraFinalizacao.Caption = "????"
    txtStatusItem.text = "1"
    txtStatusItemDescricao.text = StatusItemDescricao(txtStatusItem.text)
    lvwItemComanda.ListItems.Clear

End Function

Private Function InicializaEntradaSaida() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos
    
    'Carrega combobox
    cmbItem.Clear
    Call ComboBox(cmbItem, "Cardapio", "Id", "Descricao", " ORDER BY Descricao")
    If cmbItem.ListCount > 0 Then
       cmbItem.ListIndex = (cmbItem.ListCount - cmbItem.ListCount) '+ 1
    End If

    'Totaliza comanda
    Call TotalizaComanda

End Function

Function IncluirItensComanda() As Boolean
Dim vbl_Carregar As Boolean
Dim vil_CountLista As Integer
    
    Set vop_ItemComandoNegocios = New clsComandaItemNegocios
        vbl_Carregar = vop_ItemComandoNegocios.IncluirItensComanda(frmComandas.lvwItensComanda, lvwItemComanda)
        frmComandas.lvwItensComanda.Refresh
    Set vop_ItemComandoNegocios = Nothing

End Function

Function AlterarItensComanda() As Boolean
Dim vbl_Carregar As Boolean
Dim vil_CountLista As Integer
    
    Set vop_ItemComandoNegocios = New clsComandaItemNegocios
        vbl_Carregar = vop_ItemComandoNegocios.AlterarItensComanda(frmComandas.lvwItensComanda, lvwItemComanda, vip_ItemListaComanda)
        frmComandas.lvwItensComanda.Refresh
    Set vop_ItemComandoNegocios = Nothing

End Function

Function Editar(ByVal pFrm As Form, ByVal pIdItemListaComanda As Integer) As Boolean
Dim vcl_DescricaoItem As String

    'Verifica uma inclusao ou alteracao do cliente
    vbp_ItemComanda = True
    
    'Controle de exibicao
    lblCodigo.Visible = True
       
    'Item da lista de comanda
    vip_ItemListaComanda = pIdItemListaComanda
    
    If pFrm.Name = "frmComandas" Then
    
       'Itens comanda
       Set vop_ItemComandoNegocios = New clsComandaItemNegocios
      
          If Int(frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(1)) > 0 Then
             If vop_ItemComandoNegocios.PesquisarItemComanda(lvwItemComanda, frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(1), frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(3)) = False Then
                 MsgBox "Não foi possível encontrar o Item da Comanda !", vbCritical, "Item da Comanda"
             Else
                 txtItemComanda.text = vop_ItemComandoNegocios.IdItemComanda
                 txtIdComanda.text = vop_ItemComandoNegocios.IdComanda
                 txtIdTipoComanda.text = vop_ItemComandoNegocios.IdTipoComanda
                 txtIdCardapio.text = vop_ItemComandoNegocios.IdCardapio
                 vcl_DescricaoItem = Trim$(vop_ItemComandoNegocios.DescricaoItem)
                 txtQuantidade.text = vop_ItemComandoNegocios.Quantidade
                 lblPrecoItemComanda.Caption = Format(vop_ItemComandoNegocios.Preco, "##0.00")
                 lblTotal.Caption = Format(vop_ItemComandoNegocios.TotalPreco, "##0.00")
                 lblDataHoraComanda = vop_ItemComandoNegocios.DataHora
                 lblDataHoraPrevista = vop_ItemComandoNegocios.DataHoraPrevista
                 lblDataHoraFinalizacao = vop_ItemComandoNegocios.DataHoraFinalizacao
                 txtStatusItem.text = IIf(vop_ItemComandoNegocios.StatusItem = 0, "1", vop_ItemComandoNegocios.StatusItem)
                 txtStatusItemDescricao.text = StatusItemDescricao(txtStatusItem.text)
              End If
          Else
              'Atualiza entrada/saida
              txtItemComanda.text = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).text
              txtIdComanda.text = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(1)
              txtIdTipoComanda.text = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(2)
              txtIdCardapio.text = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(3)
              vcl_DescricaoItem = Trim$(frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(4))
              txtQuantidade.text = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(5)
              lblPrecoItemComanda.Caption = Format(frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(6), "##0.00")
              lblTotal.Caption = Format(frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(7), "##0.00")
              lblDataHoraComanda.Caption = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(8)
              lblDataHoraPrevista.Caption = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(9)
              lblDataHoraFinalizacao.Caption = frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(10)
              txtStatusItem.text = IIf(Int(frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(11)) = 0, "1", frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(11))
              txtStatusItemDescricao.text = StatusItemDescricao(frmComandas.lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(11))
              'Adiciona na lista para alteração
              lvwItemComanda.ListItems.Add , , txtItemComanda.text
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(1) = txtIdComanda.text
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(2) = txtIdTipoComanda.text
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(3) = txtIdCardapio.text
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(4) = vcl_DescricaoItem
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(5) = txtQuantidade.text
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(6) = lblPrecoItemComanda.Caption
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(7) = lblTotal.Caption
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(8) = IIf(IsNull(lblDataHoraComanda.Caption), "????", lblDataHoraComanda.Caption)
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(9) = IIf(IsNull(lblDataHoraPrevista.Caption), "????", lblDataHoraPrevista.Caption)
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(10) = IIf(IsNull(lblDataHoraFinalizacao.Caption), "????", lblDataHoraFinalizacao.Caption)
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(11) = txtStatusItem.text
              lvwItemComanda.ListItems(lvwItemComanda.ListItems.Count).SubItems(12) = IIf(IsNull(txtStatusItemDescricao.text), "????", txtStatusItemDescricao.text)
          End If
                    
          'Valida botao
          If Int(txtStatusItem.text) = 1 Then
             lblCancelar.Visible = True
          End If
                    
       Set vop_ItemComandoNegocios = Nothing
       
       'Default
       cmbItem.text = vcl_DescricaoItem
          
       'Valida Status do Item
       If Int(txtStatusItem.text) > 1 Then
          'Desabilida entrada e saida
          Call DesabilitaEntradaSaida
       End If
   
    Else
       
       If vop_ItemComandoNegocios.CancelarItemComanda(lvwItemComanda, vip_ItemListaComanda) = False Then
           MsgBox "Não foi possível encontrar o Item da Comanda !", vbCritical, "Item da Comanda"
       Else
           txtItemComanda.text = vop_ItemComandoNegocios.IdItemComanda
           txtIdComanda.text = vop_ItemComandoNegocios.IdComanda
           txtIdTipoComanda.text = vop_ItemComandoNegocios.IdTipoComanda
           txtIdCardapio.text = vop_ItemComandoNegocios.IdCardapio
           vcl_DescricaoItem = Trim$(vop_ItemComandoNegocios.DescricaoItem)
           txtQuantidade.text = vop_ItemComandoNegocios.Quantidade
           lblPrecoItemComanda.Caption = Format(vop_ItemComandoNegocios.Preco, "##0.00")
           lblTotal.Caption = Format(vop_ItemComandoNegocios.TotalPreco, "##0.00")
           lblDataHoraComanda = vop_ItemComandoNegocios.DataHora
           lblDataHoraPrevista = vop_ItemComandoNegocios.DataHoraPrevista
           lblDataHoraFinalizacao = vop_ItemComandoNegocios.DataHoraFinalizacao
           txtStatusItem.text = IIf(vop_ItemComandoNegocios.StatusItem = 0, "1", vop_ItemComandoNegocios.StatusItem)
           txtStatusItemDescricao.text = StatusItemDescricao(txtStatusItem.text)
                      
        End If
        
       'Valida visualizar
       If pFrm.Name = "frmCozinhaItensLista" Or pFrm.Name = "frmMain" Then
          lblFinalizar.Visible = True
          lblFinalizar.Left = lblGravar.Left
          lblGravar.Visible = False
          lblCancelar.Visible = True
       ElseIf pFrm.Name = "frmComandasLista" Then
          lblCancelar.Visible = True
          lblCancelar.Left = lblGravar.Left
          lblGravar.Visible = False
          lblFinalizar.Visible = False
       End If
        
       'Default
       cmbItem.text = vcl_DescricaoItem
          
       'Valida botao
       If Int(txtStatusItem.text) = 5 Then
          lblCancelar.Visible = False
       End If
       If Int(txtStatusItem.text) = 4 Then
          lblFinalizar.Visible = False
          lblCancelar.Left = lblFinalizar.Left
       End If
       
       'desabilta entrada e saida de dados
       Call DesabilitaEntradaSaida
    
    End If
        
    Me.Show vbModal

End Function

Private Function DesabilitaEntradaSaida() As Boolean

    cmbItem.Enabled = False
    txtQuantidade.Enabled = False
    lblGravar.Visible = False

End Function
