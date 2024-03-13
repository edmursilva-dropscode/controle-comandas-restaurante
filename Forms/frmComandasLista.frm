VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "PRJXTAB.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmComandasLista 
   Caption         =   "Cadastro de Comanda"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   16590
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin prjXTab.XTab xtbDevolucao 
      Height          =   6945
      Left            =   75
      TabIndex        =   1
      Top             =   795
      Width           =   16425
      _ExtentX        =   28972
      _ExtentY        =   12250
      TabCount        =   2
      TabCaption(0)   =   "  Comandas "
      TabContCtrlCnt(0)=   5
      Tab(0)ContCtrlCap(1)=   "txtLocalizarComanda"
      Tab(0)ContCtrlCap(2)=   "cmbLocalizarComanda"
      Tab(0)ContCtrlCap(3)=   "dtgComandas"
      Tab(0)ContCtrlCap(4)=   "lblTitulo3"
      Tab(0)ContCtrlCap(5)=   "lblTitulo2"
      TabCaption(1)   =   "  Itens da comanda  "
      TabContCtrlCnt(1)=   5
      Tab(1)ContCtrlCap(1)=   "cmbLocalizarItem"
      Tab(1)ContCtrlCap(2)=   "txtLocalizarItem"
      Tab(1)ContCtrlCap(3)=   "dtgItensComanda"
      Tab(1)ContCtrlCap(4)=   "lblTitulo5"
      Tab(1)ContCtrlCap(5)=   "lblTitulo4"
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
      Begin VB.ComboBox cmbLocalizarItem 
         Height          =   315
         ItemData        =   "frmComandasLista.frx":0000
         Left            =   -59745
         List            =   "frmComandasLista.frx":000D
         TabIndex        =   15
         Top             =   -15
         Width           =   1200
      End
      Begin VB.TextBox txtLocalizarItem 
         Height          =   315
         Left            =   -62895
         MaxLength       =   50
         TabIndex        =   14
         Top             =   -15
         Width           =   2475
      End
      Begin VB.TextBox txtLocalizarComanda 
         Height          =   315
         Left            =   12105
         MaxLength       =   50
         TabIndex        =   12
         Top             =   -15
         Width           =   2475
      End
      Begin VB.ComboBox cmbLocalizarComanda 
         Height          =   315
         ItemData        =   "frmComandasLista.frx":0028
         Left            =   15255
         List            =   "frmComandasLista.frx":0035
         TabIndex        =   13
         Top             =   -15
         Width           =   1200
      End
      Begin MSDataGridLib.DataGrid dtgItensComanda 
         Bindings        =   "frmComandasLista.frx":0050
         Height          =   6435
         Left            =   -74895
         TabIndex        =   16
         Top             =   405
         Width           =   16215
         _ExtentX        =   28601
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
         ColumnCount     =   13
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "IdTipoComanda"
            Caption         =   "  Comanda"
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "Quantidade"
            Caption         =   "  Qtde."
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
            DataField       =   "Preco"
            Caption         =   "        Preço"
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
            DataField       =   "TotalPreco"
            Caption         =   "          Total"
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
            Caption         =   "          Data/Hora"
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
            Caption         =   "  Data/Hora Prevista"
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
            Caption         =   "  Data/Hora Finalização"
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
         BeginProperty Column12 
            DataField       =   "StatusItemDescricao"
            Caption         =   "                   Status do Item"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   3525,166
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1904,882
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   3000,189
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgComandas 
         Bindings        =   "frmComandasLista.frx":006E
         Height          =   6435
         Left            =   105
         TabIndex        =   8
         Top             =   405
         Width           =   16215
         _ExtentX        =   28601
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
         ColumnCount     =   11
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "IdTipoComanda"
            Caption         =   "  Comanda"
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
            DataField       =   "DescricaoComanda"
            Caption         =   "DescricaoComanda"
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
            DataField       =   "NumeroMesa"
            Caption         =   "  Mesa"
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
            DataField       =   "QuantidadePessoa"
            Caption         =   "  Qtde. de Pessoas"
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
            DataField       =   "TotalPreco"
            Caption         =   "       Total da Comanda"
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
            DataField       =   "DataConfirmacaoPreparo"
            Caption         =   "                  Data/Hora"
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
            DataField       =   "DataPrevistaPreparo"
            Caption         =   "            Data/Hora Prevista"
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
            DataField       =   "DataFinalizacaoPreparo"
            Caption         =   "          Data/Hora Finbalização"
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
         BeginProperty Column10 
            DataField       =   "StatusComandaDescricao"
            Caption         =   "          Status da Comanda"
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
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   705,26
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2505,26
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2505,26
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2610,142
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   3000,189
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
         MICON           =   "frmComandasLista.frx":0088
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
         MICON           =   "frmComandasLista.frx":00A4
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
         Left            =   -60240
         TabIndex        =   20
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
         Left            =   -63750
         TabIndex        =   19
         Top             =   60
         Width           =   675
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localizar:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   11250
         TabIndex        =   18
         Top             =   60
         Width           =   675
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   2
         Left            =   14760
         TabIndex        =   17
         Top             =   60
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
   End
   Begin LVbuttons.LaVolpeButton lblNovo 
      Height          =   360
      Left            =   13380
      TabIndex        =   9
      Top             =   7830
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
      MICON           =   "frmComandasLista.frx":00C0
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
   Begin LVbuttons.LaVolpeButton lblCancelar 
      Height          =   360
      Left            =   14430
      TabIndex        =   10
      Top             =   7830
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
      MICON           =   "frmComandasLista.frx":00DC
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
      Left            =   15480
      TabIndex        =   11
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
      MICON           =   "frmComandasLista.frx":00F8
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
   Begin MSAdodcLib.Adodc adoComandas 
      Height          =   330
      Left            =   17265
      Top             =   6615
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
      RecordSource    =   $"frmComandasLista.frx":0114
      Caption         =   "adoComandas"
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
   Begin MSAdodcLib.Adodc adoItensComanda 
      Height          =   330
      Left            =   17235
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
      RecordSource    =   $"frmComandasLista.frx":0573
      Caption         =   "adoItensComanda"
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
   Begin VB.Image Image1 
      Height          =   45
      Left            =   9210
      Picture         =   "frmComandasLista.frx":0A99
      Top             =   675
      Width           =   10740
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1200
      Picture         =   "frmComandasLista.frx":141D
      Top             =   675
      Width           =   10740
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   90
      Picture         =   "frmComandasLista.frx":1DA1
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16740
   End
End
Attribute VB_Name = "frmComandasLista"
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
    cmbLocalizarComanda.ListIndex = 0
    cmbLocalizarItem.ListIndex = 0
    
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
   
   Set frmComandasLista = Nothing
   
End Sub

Private Sub lblCancelar_Click()
    If vil_IdCliente = 0 Then Exit Sub
    
    If MsgBox("Confirma o Cancelamento ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_ComandasNegocios = New clsComandasNegocios
          vop_ComandasNegocios.IdComandas = vil_IdComanda
          If vop_ComandasNegocios.ExcluirComanda() = True Then
             txtLocalizar.text = Empty
             Call CarregarGrid
          End If
      Set vop_ComandasNegocios = Nothing
    End If
End Sub

Private Sub lblNovo_Click()
    frmComandas.Show vbModal
End Sub

Private Sub lblFechar_Click()
   Unload Me
End Sub

Private Sub dtgComandas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If dtgComandas.Bookmark > 0 Then
      dtgComandas.SelBookmarks.Add dtgComandas.Bookmark
      vil_IdComanda = dtgComandas.Columns(0).text
   End If
End Sub

Private Sub dtgComandas_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgComandas.SelBookmarks.Count > 0 Then
           dtgComandas.SelBookmarks.Remove 0
        End If
        adoComandas.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgComandas_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgComandas.SelBookmarks.Count > 0 Then
           dtgComandas.SelBookmarks.Remove 0
        End If
        adoComandas.Recordset.MovePrevious
        
    End If
   
End Sub

Private Sub dtgComandas_DblClick()

On Error GoTo TrataErros

   Call frmComandas.Form_Load
   Call frmComandas.Editar(vil_IdComanda)

TrataErros:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub dtgItensComanda_DblClick()
On Error GoTo TrataErros

   Call frmComandaItens.Form_Load
   Call frmComandaItens.Editar(Me, vip_ItemListaComanda)

TrataErros:
    If Err.Number <> 0 Then Exit Sub
End Sub

Private Sub dtgItensComanda_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgItensComanda.SelBookmarks.Count > 0 Then
           dtgItensComanda.SelBookmarks.Remove 0
        End If
        adoItensComanda.Recordset.MoveNext
        
    End If

End Sub

Private Sub dtgItensComanda_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgItensComanda.SelBookmarks.Count > 0 Then
           dtgItensComanda.SelBookmarks.Remove 0
        End If
        adoItensComanda.Recordset.MovePrevious
        
    End If
   
End Sub

Private Sub dtgItensComanda_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If dtgItensComanda.Bookmark > 0 Then
      dtgItensComanda.SelBookmarks.Add dtgItensComanda.Bookmark
      vip_ItemListaComanda = dtgItensComanda.Columns(0).text
   End If
End Sub

Private Sub txtLocalizarComanda_Change()
  
On Error GoTo TrataErros

   Set vop_ComandasNegocios = New clsComandasNegocios
       Call vop_ComandasNegocios.LocalizarComanda(adoComandas, txtLocalizarComanda.text, cmbLocalizarComanda.ListIndex)
   Set vop_ComandasNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
       Exit Sub
    End If
    
End Sub

Private Sub cmbLocalizarComanda_Click()

On Error GoTo TrataErros

   Set vop_ComandasNegocios = New clsComandasNegocios
       Call vop_ComandasNegocios.LocalizarComanda(adoComandas, txtLocalizarComanda.text, cmbLocalizarComanda.ListIndex)
   Set vop_ComandasNegocios = Nothing
    
   txtLocalizarComanda.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
       Exit Sub
    End If
    
End Sub

Private Sub txtLocalizarItem_Change()

On Error GoTo TrataErros

   Set vop_ItemComandoNegocios = New clsComandaItemNegocios
       Call vop_ItemComandoNegocios.LocalizarItem(adoItensComanda, txtLocalizarItem.text, cmbLocalizarItem.ListIndex)
   Set vop_ItemComandoNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ItemComandoNegocios = Nothing
       Exit Sub
    End If

End Sub

Private Sub cmbLocalizarItem_Click()

On Error GoTo TrataErros

   Set vop_ItemComandoNegocios = New clsComandaItemNegocios
       Call vop_ItemComandoNegocios.LocalizarItem(adoItensComanda, txtLocalizarItem.text, cmbLocalizarItem.ListIndex)
   Set vop_ItemComandoNegocios = Nothing
    
   txtLocalizarItem.text = Empty
   
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
        
        vbl_Carregar = vop_ComandasNegocios.CarregarGridComandaRS(adoComandas, txtLocalizarComanda.text, IIf(cmbLocalizarComanda.ListIndex < 0, 0, cmbLocalizarComanda.ListIndex), adoItensComanda, txtLocalizarItem.text, IIf(cmbLocalizarItem.ListIndex < 0, 0, cmbLocalizarItem.ListIndex))
        If vbl_Carregar = True And dtgComandas.Bookmark = 0 Then
           dtgComandas.Refresh
           lblNovo.Left = 14430
           lblCancelar.Visible = False
        Else
           lblNovo.Left = 13380
           lblCancelar.Left = 14430
           lblCancelar.Visible = True
        End If
        
    Set vop_ComandasNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
       Exit Sub
    End If

End Sub



