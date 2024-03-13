VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "PRJXTAB.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmComandas 
   Caption         =   "Cadastro de Comandas"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   14370
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtStatusComanda 
      Height          =   285
      Left            =   1515
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   165
      Visible         =   0   'False
      Width           =   150
   End
   Begin prjXTab.XTab xtbComandas 
      Height          =   5700
      Left            =   90
      TabIndex        =   9
      Top             =   795
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   10054
      TabCount        =   1
      TabCaption(0)   =   "  Comandas "
      TabContCtrlCnt(0)=   4
      Tab(0)ContCtrlCap(1)=   "lblTipoComanda"
      Tab(0)ContCtrlCap(2)=   "txtIdTipoComanda"
      Tab(0)ContCtrlCap(3)=   "fraTab1"
      Tab(0)ContCtrlCap(4)=   "lblIdComanda"
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
      Begin LVbuttons.LaVolpeButton lblTipoComanda 
         Height          =   285
         Left            =   2595
         TabIndex        =   1
         Top             =   420
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmComandas.frx":0000
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
      Begin VB.TextBox txtIdTipoComanda 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "CONTATO"
         Top             =   405
         Width           =   720
      End
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   5355
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   13905
         Begin VB.Frame fraLinha 
            Height          =   30
            Index           =   2
            Left            =   -315
            TabIndex        =   30
            Top             =   480
            Width           =   9300
         End
         Begin VB.Frame fraLinha 
            Height          =   1710
            Index           =   1
            Left            =   8970
            TabIndex        =   21
            Top             =   -135
            Width           =   20
         End
         Begin VB.CommandButton cmdDeletarItem 
            Height          =   570
            Left            =   9780
            Picture         =   "frmComandas.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   510
            Width           =   570
         End
         Begin VB.CommandButton cmdIncluirItem 
            Height          =   570
            Left            =   9135
            Picture         =   "frmComandas.frx":06A2
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   510
            Width           =   570
         End
         Begin VB.Frame fraLinha 
            Height          =   1185
            Index           =   0
            Left            =   2475
            TabIndex        =   18
            Top             =   405
            Width           =   20
         End
         Begin VB.Frame Frame1 
            Caption         =   " Itens da Comanda "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3810
            Left            =   30
            TabIndex        =   17
            Top             =   1455
            Width           =   13875
            Begin MSComctlLib.ListView lvwItensComanda 
               Height          =   3495
               Left            =   150
               TabIndex        =   23
               Top             =   195
               Width           =   13590
               _ExtentX        =   23971
               _ExtentY        =   6165
               View            =   3
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               HoverSelection  =   -1  'True
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
         End
         Begin VB.TextBox txtMesa 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   2
            Tag             =   "0"
            Text            =   "0"
            Top             =   600
            Width           =   720
         End
         Begin VB.TextBox txtQtdePessoas 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   3
            Tag             =   "0"
            Text            =   "0"
            Top             =   1035
            Width           =   720
         End
         Begin VB.Label lblMesa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
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
            Left            =   75
            TabIndex        =   29
            Top             =   180
            Width           =   675
         End
         Begin VB.Label lblTiposComanda 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2970
            TabIndex        =   28
            Top             =   90
            Width           =   5295
         End
         Begin VB.Label lblTotalComanda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   22.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   10455
            TabIndex        =   27
            Top             =   540
            Width           =   3420
         End
         Begin VB.Label lblDataHora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Index           =   4
            Left            =   10440
            TabIndex        =   26
            Top             =   225
            Width           =   495
         End
         Begin VB.Label lblDataHora 
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
            Index           =   3
            Left            =   9150
            TabIndex        =   22
            Top             =   195
            Width           =   480
         End
         Begin VB.Label lblDataHoraComanda 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2625
            TabIndex        =   4
            Top             =   1035
            Width           =   1800
         End
         Begin VB.Label lblDataHoraFinalizacao 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6810
            TabIndex        =   6
            Top             =   1035
            Width           =   2025
         End
         Begin VB.Label lblDataHoraPrevista 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4755
            TabIndex        =   5
            Top             =   1035
            Width           =   1800
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
            Left            =   6810
            TabIndex        =   16
            Top             =   735
            Width           =   1920
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
            Left            =   4740
            TabIndex        =   15
            Top             =   720
            Width           =   1710
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
            Left            =   2625
            TabIndex        =   14
            Top             =   705
            Width           =   960
         End
         Begin VB.Label lblMesa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mesa:"
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
            TabIndex        =   12
            Top             =   675
            Width           =   510
         End
         Begin VB.Label lblQtdePessoas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. de pessoas:"
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
            Left            =   30
            TabIndex        =   11
            Top             =   1095
            Width           =   1560
         End
      End
      Begin VB.Label lblIdComanda 
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
         Left            =   1170
         TabIndex        =   13
         Top             =   15
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView lvwComanda 
      Height          =   360
      Left            =   14460
      TabIndex        =   20
      Top             =   1215
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IdTipoComanda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NumeroMesa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "QuantidadePessoa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DataConfirmacaoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "DataPrevistaPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "DataFinalizacaoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "StatusComanda"
         Object.Width           =   2540
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton lblGravar 
      Height          =   360
      Left            =   12180
      TabIndex        =   24
      Top             =   6585
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
      MICON           =   "frmComandas.frx":0DA7
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
      Left            =   13245
      TabIndex        =   25
      Top             =   6585
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
      MICON           =   "frmComandas.frx":0DC3
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
   Begin MSComctlLib.ListView lvwTiposComandas 
      Height          =   360
      Left            =   14460
      TabIndex        =   31
      Top             =   1725
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
         Text            =   "StatusComanda"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwIdItemComandaDeletado 
      Height          =   360
      Left            =   14475
      TabIndex        =   33
      Top             =   2250
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   2540
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton lblEnviarCozinha 
      Height          =   360
      Left            =   10305
      TabIndex        =   34
      Top             =   6585
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Enviar Itens Cozinha"
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
      MICON           =   "frmComandas.frx":0DDF
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
   Begin LVbuttons.LaVolpeButton lblFecharComanda 
      Height          =   360
      Left            =   8565
      TabIndex        =   35
      Top             =   6570
      Visible         =   0   'False
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Fechar Comanda"
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
      MICON           =   "frmComandas.frx":0DFB
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
   Begin VB.Image Image1 
      Height          =   45
      Left            =   11040
      Picture         =   "frmComandas.frx":0E17
      Top             =   720
      Width           =   10740
   End
   Begin VB.Image imgIcone 
      Height          =   480
      Left            =   90
      Picture         =   "frmComandas.frx":179B
      Top             =   135
      Width           =   480
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   17955
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   120
      Picture         =   "frmComandas.frx":2465
      Top             =   720
      Width           =   10740
   End
End
Attribute VB_Name = "frmComandas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_System As New clsSystem
Dim vop_ComandasNegocios As New clsComandasNegocios
Dim vop_ItemComandaNegocios As New clsComandaItemNegocios
Dim vop_TiposComandasNegocios As New clsTiposComandasNegocios
'Variaveis de controle do form
Dim vbp_Comanda As Boolean                             'Verifica uma inclusao ou alteracao
Dim vip_IdComanda As Integer
Dim vip_ItemListaComanda As Integer



'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Public Sub Form_Load()

    'Verifica uma inclusao ou alteracao
    vbp_Comanda = False
    
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
   
   Set frmComandas = Nothing
   
End Sub

Private Sub lblEnviarCozinha_Click()

    'Valida Itens na lista de comanda
    If ValidaEnviaCozinha = False Then Exit Sub
    
   'Envia item para processamento na cozinha
   If MsgBox("Confirma envio para Cozinha ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_ComandasNegocios = New clsComandasNegocios
          vop_ComandasNegocios.IdComandas = lblIdComanda.Caption
          vop_ComandasNegocios.IdTipoComanda = txtIdTipoComanda.text
          vop_ComandasNegocios.NumeroMesa = txtMesa.text
          vop_ComandasNegocios.QuantidadePessoa = txtQtdePessoas.text
          vop_ComandasNegocios.DataConfirmacaoPreparo = lblDataHoraComanda.Caption
          vop_ComandasNegocios.DataPrevistaPreparo = lblDataHoraPrevista.Caption
          vop_ComandasNegocios.DataFinalizacaoPreparo = lblDataHoraFinalizacao.Caption
          vop_ComandasNegocios.StatusComanda = txtStatusComanda.text
          If vop_ComandasNegocios.EnviarItensProcessamentoCozinha(lvwItensComanda) = True Then
             MsgBox "Itens da Comanda enviada para cozinha com sucesso !", vbExclamation, "Comanda"
          End If
      Set vop_ComandasNegocios = Nothing
      
   Else
   
      Exit Sub
      
   End If
   
   'Atualiza grid
   Call frmComandasLista.CarregarGrid
   
   'Editar dados da comanda
   'Call LocalizarEditar(Int(lblIdComanda.Caption))
   
   'Valida metodo Editar
   Call Editar(Int(lblIdComanda.Caption))

End Sub

Private Sub lblFecharComanda_Click()

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   If VerItensFecharComanda = False Then Exit Sub
   
   'Grava entrada de dados
   If MsgBox("Confirma fechar Comanda", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_ComandasNegocios = New clsComandasNegocios
          vop_ComandasNegocios.IdComandas = lblIdComanda.Caption
          vop_ComandasNegocios.StatusComanda = 2
          If vop_ComandasNegocios.FecharComanda() = True Then
             MsgBox "Comanda fechada com sucesso !", vbExclamation, "Comanda"
          End If
          
          'Valida metodo Editar
          Call Editar(vop_ComandasNegocios.IdComandas)
          
      Set vop_ComandasNegocios = Nothing
   
   Else
   
      Exit Sub
      
   End If
   
   'Atualiza grid
   Call frmComandasLista.CarregarGrid
   
   'Atualiza variaveis de controle
   vbp_Comanda = ValidaIncluirAlterarComanda
   
   'Desabilida entrada e saida
   Call DesabilitaEntradaSaidaFechar


End Sub

Private Sub lblGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Comanda = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   'Grava entrada de dados
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_ComandasNegocios = New clsComandasNegocios
          vop_ComandasNegocios.IdComandas = lblIdComanda.Caption
          vop_ComandasNegocios.IdTipoComanda = txtIdTipoComanda.text
          vop_ComandasNegocios.NumeroMesa = txtMesa.text
          vop_ComandasNegocios.QuantidadePessoa = txtQtdePessoas.text
          vop_ComandasNegocios.DataConfirmacaoPreparo = lblDataHoraComanda.Caption
          vop_ComandasNegocios.DataPrevistaPreparo = lblDataHoraPrevista.Caption
          vop_ComandasNegocios.DataFinalizacaoPreparo = lblDataHoraFinalizacao.Caption
          vop_ComandasNegocios.StatusComanda = txtStatusComanda.text
          If vbp_Comanda = False Then
             If vop_ComandasNegocios.IncluirComanda(lvwItensComanda, vop_ComandasNegocios.IdComandas) = True Then
                MsgBox "Comanda cadastrada com sucesso !", vbExclamation, "Comanda"
             End If
          Else
             If vop_ComandasNegocios.AlterarComanda(lvwItensComanda, lvwIdItemComandaDeletado) = True Then
                MsgBox "Comanda alterada com sucesso !", vbExclamation, "Comanda"
             End If
          End If
          
          'Valida metodo Editar
          Call Editar(vop_ComandasNegocios.IdComandas)
          
      Set vop_ComandasNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmComandasLista.CarregarGrid
   
   'Atualiza variaveis de controle
   vbp_Comanda = ValidaIncluirAlterarComanda
   
   'Desabilida entrada e saida
   Call DesabilitaEntradaSaida
      
End Sub

Private Sub lblTipoComanda_Click()

On Error GoTo TrataErros

   Set vop_System = New clsSystem
       vop_System.Coluna_01 = "Código:"
       vop_System.Coluna_02 = "Áreas:"
       vop_System.FormataColuna_01 = "000"
       If vop_System.FindLista("TiposComandas", "Id", "Descricao", frmFindLista.lvwLista) = True Then
          frmFindLista.Caption = "Área"
          frmFindLista.optSort01.Caption = "Código"
          frmFindLista.optSort02.Caption = "Descricao"
          frmFindLista.Show vbModal
       Else
          txtIdTipoComanda.SetFocus
          Set vop_System = Nothing
          Exit Sub
       End If
    Set vop_System = Nothing
    
    If Trim(frmFindLista.lvwLista.SelectedItem.text) <> Empty Then
       txtIdTipoComanda.text = Trim(frmFindLista.lvwLista.SelectedItem)
       lblTiposComanda.Caption = frmFindLista.lvwLista.ListItems(frmFindLista.lvwLista.SelectedItem.Index).SubItems(1)
       Unload frmFindLista

    Else
       Unload frmFindLista
       txtIdTipoComanda.SetFocus
    End If

TrataErros:
    If Err.Number <> 0 Then
       MsgBox "Não foi possível listar as Áreas !", vbCritical
       Err.Clear
    End If
End Sub

Private Sub lvbFechar_Click()
   Unload Me
End Sub

Private Sub lvwItensComanda_Click()

    If lvwItensComanda.ListItems.Count <> 0 Then
       vip_ItemListaComanda = lvwItensComanda.SelectedItem.Index
    End If

End Sub

Private Sub lvwItensComanda_DblClick()
On Error GoTo TrataErros

   Call frmComandaItens.Form_Load
   Call frmComandaItens.Editar(Me, vip_ItemListaComanda)
   
   'Itens comanda
   Set vop_ItemComandaNegocios = New clsComandaItemNegocios
       lvwItensComanda.ListItems.Clear
       If vop_ItemComandaNegocios.PesquisarItemComanda(lvwItensComanda, lblIdComanda.Caption, 0) = False Then
           MsgBox "Não foi possível encontrar o Item da Comanda !", vbCritical, "Item da Comanda"
       End If
   Set vop_ItemComandaNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then Exit Sub
End Sub

Private Sub txtIdTipoComanda_Change()
    
    If VerNumeros = False Then Exit Sub
    
End Sub

Private Sub txtIdTipoComanda_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If

End Sub

Private Sub txtIdTipoComanda_KeyPress(KeyAscii As Integer)

    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtIdTipoComanda_LostFocus()

    If VerTipoComanda() = False Then Exit Sub

End Sub

Private Sub txtMesa_Change()

   If VerNumeros = False Then Exit Sub
   
End Sub

Private Sub txtMesa_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtMesa_KeyPress(KeyAscii As Integer)

    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    
End Sub

Private Sub txtQtdePessoas_Change()

   If VerNumeros = False Then Exit Sub

End Sub

Private Sub txtQtdePessoas_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtQtdePessoas_KeyPress(KeyAscii As Integer)

    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    
End Sub

Private Sub cmdIncluirItem_Click()

   'Valida entrada de dados
   If VerCamposComanda = False Then Exit Sub
   
   'Chama cadastro de itens da comanda
   frmComandaItens.Show vbModal

End Sub

Private Sub cmdDeletarItem_Click()
Dim vsp_Mensagem As String
Dim vil_CountLista As Integer

   If lvwItensComanda.ListItems.Count = 0 Then Exit Sub
   If vip_ItemListaComanda = 0 Then Exit Sub
   
   If Int(lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(11)) <= 1 Then
   
      vsp_Mensagem = Trim$("Para o item " + vbCrLf & "" & Trim$(UCase(lvwItensComanda.ListItems(vip_ItemListaComanda).SubItems(4))) & "" + vbCrLf & "confirma a exclusão ?")
      If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
         
         'Adiciona na lista item deletado da comanda
         If Int(lvwItensComanda.ListItems(vip_ItemListaComanda).text) > 0 Then
            lvwIdItemComandaDeletado.ListItems.Add , , lvwItensComanda.ListItems(vip_ItemListaComanda).text
         End If
         
         'Exclui o item da comanda
         lvwItensComanda.ListItems.Remove (vip_ItemListaComanda)
         lvwItensComanda.Refresh
         
         'Totaliza comanda
         Call TotalizaComanda
                  
      End If
      
   Else
   
      MsgBox "Impossível excluir item com este Status na Comanda !", vbExclamation, "Comanda"
   
   End If


End Sub



'Metodos











'Funcoes
Function ValidaIncluirAlterarComanda() As Boolean
Dim vbl_Retorno As Boolean
Dim vil_CountLista As Integer
    
    vbl_Retorno = False
    If lvwItensComanda.ListItems.Count <> 0 Then
    
        For vil_CountLista = 1 To lvwItensComanda.ListItems.Count
            If Int(lvwItensComanda.ListItems(vil_CountLista).text) <> 0 Then
               vbl_Retorno = True
               Exit For
            End If
        Next vil_CountLista
    End If
    
    ValidaIncluirAlterarComanda = vbl_Retorno
        
End Function

Function Editar(ByVal pIdcomanda As Integer) As Boolean

    'Editar dados da comanda
    Call LocalizarEditar(pIdcomanda)
    
    'Exibe form
    If Me.Visible = False Then
       Me.Show vbModal
    End If

End Function

Private Function LocalizarEditar(ByVal pIdcomanda As Integer) As Boolean

    'Verifica uma inclusao ou alteracao do cliente
    vbp_Comanda = True
    
    'Controle de exibicao
    lblIdComanda.Visible = False
    lblEnviarCozinha.Visible = True
    lblFecharComanda.Visible = True
    
    'Default
    lvwComanda.ListItems.Clear
    lvwItensComanda.ListItems.Clear
    
    'Valida metodo Editar
    Set vop_ComandasNegocios = New clsComandasNegocios
        'Comanda
        If vop_ComandasNegocios.PesquisarComanda(lvwComanda, pIdcomanda, 0) = True Then
           lblIdComanda.Caption = pIdcomanda
           txtIdTipoComanda.text = vop_ComandasNegocios.IdTipoComanda
           txtMesa.text = vop_ComandasNegocios.NumeroMesa
           txtQtdePessoas.text = vop_ComandasNegocios.QuantidadePessoa
           lblDataHoraComanda.Caption = vop_ComandasNegocios.DataConfirmacaoPreparo
           lblDataHoraPrevista.Caption = vop_ComandasNegocios.DataPrevistaPreparo
           lblDataHoraFinalizacao.Caption = vop_ComandasNegocios.DataFinalizacaoPreparo
           txtStatusComanda.text = vop_ComandasNegocios.StatusComanda
           
           'Default
           Call txtIdTipoComanda_LostFocus
           
           'Itens comanda
           Set vop_ItemComandaNegocios = New clsComandaItemNegocios
               lvwItensComanda.ListItems.Clear
               If vop_ItemComandaNegocios.PesquisarItemComanda(lvwItensComanda, lblIdComanda.Caption, 0) = False Then
                   MsgBox "Não foi possível encontrar o Item da Comanda !", vbCritical, "Item da Comanda"
               End If
           Set vop_ItemComandaNegocios = Nothing
        Else
            MsgBox "Não foi possível encontrar a Comanda !", vbCritical, "Comanda"
        End If
          
    Set vop_ComandasNegocios = Nothing
    
    'Desabilida entrada e saida
    If Int(txtStatusComanda.text) <> 2 Then
       Call DesabilitaEntradaSaida
    Else
       Call DesabilitaEntradaSaidaFechar
    End If

End Function

Private Function VerCampos() As Boolean
    
    If Trim$(txtMesa.text) = Empty Or Int(txtMesa.text) = 0 Then
        MsgBox "Informe a mesa da Comanda !", vbExclamation, "Comanda"
        If txtMesa.text <> Empty Then txtMesa.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtQtdePessoas.text) = Empty Or Int(txtQtdePessoas.text) = 0 Then
        MsgBox "Informe a quantidade de pessoas da Comanda !", vbExclamation, "Comanda"
        If txtQtdePessoas.text <> Empty Then txtQtdePessoas.SetFocus
        VerCampos = False
        Exit Function
    End If
    If lvwItensComanda.ListItems.Count <= 0 Then
        MsgBox "Informe algum item(ns) para a Comanda !", vbExclamation, "Comanda"
        lvwItensComanda.SetFocus
        VerCampos = False
        Exit Function
    End If
    
    VerCampos = True

End Function

Private Function VerItensFecharComanda() As Boolean
Dim vbl_ValidaItensLista As Boolean
Dim vil_CountLista As Integer

    vbl_ValidaItensLista = True
    VerItensFecharComanda = True
    For vil_CountLista = 1 To lvwItensComanda.ListItems.Count
        If Int(lvwItensComanda.ListItems(vil_CountLista).SubItems(11)) < 3 Then
           vbl_ValidaItensLista = False
           Exit For
        End If
    Next
    If vbl_ValidaItensLista = False Then
       MsgBox "Há item(ns) com Status a ser resolvido na Comanda !", vbExclamation, "Comanda"
       VerItensFecharComanda = False
       lblFecharComanda.SetFocus
    End If
    
End Function


Private Function VerTipoComanda() As Boolean

    If Trim$(txtIdTipoComanda.text) = "0" Then
       MsgBox "Comanda não existe !", vbCritical, "Comanda"
       lblTiposComanda.Caption = Empty
       txtIdTipoComanda.SetFocus
       VerTipoComanda = False
       Exit Function
    ElseIf Trim$(txtIdTipoComanda.text) <> Empty Then
       Set vop_TiposComandasNegocios = New clsTiposComandasNegocios
          If vop_TiposComandasNegocios.PesquisarTipoComanda(lvwTiposComandas, txtIdTipoComanda, 0) = True Then
             lblTiposComanda.Caption = vop_TiposComandasNegocios.Descricao
             If vbp_Comanda = False And vop_TiposComandasNegocios.StatusComanda = 1 Then
                MsgBox "Comanda já esta em aberta !", vbCritical, "Comanda"
                txtIdTipoComanda.SetFocus
                VerTipoComanda = False
                Exit Function
             End If
          Else
             MsgBox "Comanda não existe !", vbCritical, "Comanda"
             lblTiposComanda.Caption = Empty
             txtIdTipoComanda.SetFocus
             VerTipoComanda = False
             Exit Function
          End If
       Set vop_TiposComandasNegocios = Nothing
    Else
       lblTiposComanda.Caption = Empty
    End If

    VerTipoComanda = True

End Function


Private Function VerNumeros() As Boolean

    If IsNumeric(txtIdTipoComanda.text) = False Then
       If txtIdTipoComanda.text <> Empty Then txtIdTipoComanda.SetFocus
       VerNumeros = False
       Exit Function
    ElseIf IsNumeric(txtMesa.text) = False Then
       If txtMesa.text <> Empty Then txtMesa.SetFocus
       VerNumeros = False
       Exit Function
    ElseIf IsNumeric(txtQtdePessoas.text) = False Then
       If txtQtdePessoas.text <> Empty Then txtQtdePessoas.SetFocus
       VerNumeros = False
       Exit Function
    End If
       
    VerNumeros = True

End Function

Private Function DefaultCampos() As BookmarkEnum

    'txtIdTipoComanda.text = "0"
    txtMesa.text = "0"
    txtQtdePessoas.text = "0"
    lblTotalComanda.Caption = "0,00"
    lblDataHoraComanda.Caption = VerDataHoraAtual()
    lblDataHoraPrevista.Caption = "????"
    lblDataHoraFinalizacao.Caption = "????"
    txtStatusComanda.text = "1"
    vip_ItemListaComanda = 0
    
End Function

Private Function DesabilitaEntradaSaidaFechar() As Boolean

    txtIdTipoComanda.Enabled = False
    lblTipoComanda.Enabled = False
    txtMesa.Enabled = False
    txtQtdePessoas.Enabled = False
    cmdIncluirItem.Enabled = False
    cmdDeletarItem.Enabled = False
    lblGravar.Visible = False
    lblEnviarCozinha.Visible = False
    lblFecharComanda.Visible = False

    'Totaliza comanda
    Call TotalizaComanda

End Function

Private Function DesabilitaEntradaSaida() As Boolean

    txtIdTipoComanda.Enabled = False
    lblTipoComanda.Enabled = False
    txtMesa.Enabled = False
    txtQtdePessoas.Enabled = False

    'Totaliza comanda
    Call TotalizaComanda

End Function

Private Function InicializaEntradaSaida() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos

End Function

Private Function VerCamposComanda() As Boolean
    
    If Trim$(txtIdTipoComanda.text) = Empty Or Trim$(txtIdTipoComanda.text) = "0" Then
        MsgBox "Informe o numero da Comanda !", vbExclamation, "Comanda"
        If txtIdTipoComanda.text <> Empty Then txtIdTipoComanda.SetFocus
        VerCamposComanda = False
        Exit Function
    End If
    If Trim$(txtMesa.text) = Empty Or Trim$(txtMesa.text) = "0" Then
        MsgBox "Informe a mess da Comanda !", vbExclamation, "Comanda"
        If txtMesa.text <> Empty Then txtMesa.SetFocus
        VerCamposComanda = False
        Exit Function
    End If
    If Trim$(txtQtdePessoas.text) = Empty Or Trim$(txtQtdePessoas.text) = "0" Then
        MsgBox "Informe a quantidade de pessoas da Comanda !", vbExclamation, "Comanda"
        If txtQtdePessoas.text <> Empty Then txtQtdePessoas.SetFocus
        VerCamposComanda = False
        Exit Function
    End If
    
    'Valida tipo de comanda
    If VerTipoComanda() = False Then Exit Function
    
    VerCamposComanda = True

End Function

Private Function ValidaEnviaCozinha() As Boolean
Dim vil_CountLista As Integer
Dim vbl_ValidaItensLista As Boolean

   ValidaEnviaCozinha = True
   If lvwItensComanda.ListItems.Count > 0 Then
       vbl_ValidaItensLista = False
       For vil_CountLista = 1 To lvwItensComanda.ListItems.Count
           If Int(lvwItensComanda.ListItems(vil_CountLista).text) = 0 Then
              vbl_ValidaItensLista = True
              Exit For
           End If
       Next
       If vbl_ValidaItensLista = True Then
          MsgBox "Existe Item a ser gravado antes de enviar para Cozinha !", vbExclamation, "Comanda"
          ValidaEnviaCozinha = False
          lblEnviarCozinha.SetFocus
       End If
   End If
   If lvwItensComanda.ListItems.Count = 0 Then
       MsgBox "Comanda sem Item(ns) para enviar a Cozinha !", vbExclamation, "Comanda"
       ValidaEnviaCozinha = False
       lblEnviarCozinha.SetFocus
   End If
   If lvwItensComanda.ListItems.Count > 0 Then
       vbl_ValidaItensLista = True
       For vil_CountLista = 1 To lvwItensComanda.ListItems.Count
           If Int(lvwItensComanda.ListItems(vil_CountLista).SubItems(11)) < 3 Then
              vbl_ValidaItensLista = False
              Exit For
           End If
       Next
       If vbl_ValidaItensLista = True Then
          MsgBox "Não há item(ns) para enviar a Cozinha !", vbExclamation, "Comanda"
          ValidaEnviaCozinha = False
          lblEnviarCozinha.SetFocus
       End If
   End If

End Function
