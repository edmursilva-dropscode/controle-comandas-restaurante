VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de pedidos KDS software"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   10170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   20400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox mainBorder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   15
      ScaleHeight     =   8655
      ScaleWidth      =   19965
      TabIndex        =   3
      Top             =   360
      Width           =   19965
      Begin VB.PictureBox MenuItem 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   3
         Left            =   11355
         ScaleHeight     =   7395
         ScaleWidth      =   2850
         TabIndex        =   15
         Top             =   15
         Visible         =   0   'False
         Width           =   2910
         Begin VB.Frame frmEstatistica 
            Height          =   600
            Left            =   30
            TabIndex        =   26
            Top             =   270
            Width           =   16230
            Begin VB.ComboBox cmbLocalizarItemEstatistica 
               Height          =   315
               ItemData        =   "frmMain.frx":0000
               Left            =   540
               List            =   "frmMain.frx":000D
               TabIndex        =   28
               Top             =   180
               Width           =   4995
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   105
               TabIndex        =   27
               Top             =   255
               Width           =   360
            End
         End
         Begin MSAdodcLib.Adodc adoEstatisticaItensDias 
            Height          =   330
            Left            =   9450
            Top             =   5475
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
            RecordSource    =   $"frmMain.frx":0097
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
         Begin MSAdodcLib.Adodc adoEstatisticaItens 
            Height          =   330
            Left            =   9480
            Top             =   6060
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
            RecordSource    =   $"frmMain.frx":03C5
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
         Begin MSDataGridLib.DataGrid dtgEstatisticaItens 
            Bindings        =   "frmMain.frx":06F1
            Height          =   6435
            Left            =   30
            TabIndex        =   30
            Top             =   900
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
         Begin MSDataGridLib.DataGrid dtgEstatisticaItensDias 
            Bindings        =   "frmMain.frx":0713
            Height          =   6435
            Left            =   30
            TabIndex        =   29
            Top             =   900
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
         Begin VB.Image imgItem 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   3
            Left            =   1230
            Picture         =   "frmMain.frx":0739
            Top             =   120
            Width           =   210
         End
         Begin VB.Label lblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatisticas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   60
            Width           =   1005
         End
      End
      Begin VB.PictureBox MenuItem 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   2
         Left            =   7965
         ScaleHeight     =   7395
         ScaleWidth      =   2850
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2910
         Begin VB.Frame frmItensCozinha 
            Height          =   600
            Left            =   30
            TabIndex        =   19
            Top             =   270
            Width           =   16230
            Begin VB.TextBox txtLocalizarItemCozinha 
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   24
               Top             =   180
               Width           =   2475
            End
            Begin VB.ComboBox cmbLocalizarItemCozinha 
               Height          =   315
               ItemData        =   "frmMain.frx":0A8B
               Left            =   3915
               List            =   "frmMain.frx":0A98
               TabIndex        =   23
               Top             =   180
               Width           =   1200
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   3465
               TabIndex        =   21
               Top             =   255
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Localizar:"
               Height          =   195
               Left            =   105
               TabIndex        =   20
               Top             =   255
               Width           =   675
            End
         End
         Begin MSDataGridLib.DataGrid dtgCozinhaItens 
            Bindings        =   "frmMain.frx":0AB3
            Height          =   6435
            Left            =   30
            TabIndex        =   22
            Top             =   900
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
         Begin MSAdodcLib.Adodc adoCozinhaItens 
            Height          =   330
            Left            =   18555
            Top             =   6225
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
            RecordSource    =   $"frmMain.frx":0AD1
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
         Begin VB.Image imgItem 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   2
            Left            =   1350
            Picture         =   "frmMain.frx":1098
            Top             =   120
            Width           =   210
         End
         Begin VB.Label lblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Itens Cozinha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   60
            Width           =   1230
         End
      End
      Begin VB.PictureBox MenuItem 
         Height          =   7455
         Index           =   1
         Left            =   4590
         ScaleHeight     =   7395
         ScaleWidth      =   2850
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2910
         Begin VB.Frame frmComanda 
            Height          =   600
            Left            =   30
            TabIndex        =   11
            Top             =   270
            Width           =   16230
            Begin VB.TextBox txtLocalizarComanda 
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   13
               Top             =   180
               Width           =   2475
            End
            Begin VB.ComboBox cmbLocalizarComanda 
               Height          =   315
               ItemData        =   "frmMain.frx":13EA
               Left            =   3915
               List            =   "frmMain.frx":13F7
               TabIndex        =   12
               Top             =   180
               Width           =   1200
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Localizar:"
               Height          =   195
               Left            =   105
               TabIndex        =   17
               Top             =   255
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   3465
               TabIndex        =   16
               Top             =   255
               Width           =   360
            End
         End
         Begin MSDataGridLib.DataGrid dtgComandas 
            Bindings        =   "frmMain.frx":1412
            Height          =   6435
            Left            =   30
            TabIndex        =   10
            Top             =   900
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
         Begin MSAdodcLib.Adodc adoComandas 
            Height          =   330
            Left            =   16470
            Top             =   1290
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
            RecordSource    =   $"frmMain.frx":142C
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
         Begin VB.Label lblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comandas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   60
            Width           =   915
         End
         Begin VB.Image imgItem 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   1
            Left            =   1110
            Picture         =   "frmMain.frx":188B
            Top             =   120
            Width           =   210
         End
      End
      Begin VB.CommandButton cmdEstatisticas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estatisticas"
         Height          =   1290
         Left            =   120
         Picture         =   "frmMain.frx":1BDD
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3180
         Width           =   1275
      End
      Begin VB.CommandButton cmdItensCozinha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Itens Cozinha"
         Height          =   1290
         Left            =   120
         Picture         =   "frmMain.frx":22E0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1665
         Width           =   1275
      End
      Begin VB.CommandButton cmdComandas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comandas"
         Height          =   1290
         Left            =   120
         Picture         =   "frmMain.frx":2AAE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   165
         Width           =   1275
      End
      Begin VB.PictureBox MenuItem 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   0
         Left            =   1560
         ScaleHeight     =   7395
         ScaleWidth      =   2505
         TabIndex        =   14
         Top             =   0
         Width           =   2565
      End
   End
   Begin VB.Timer tmrEnviarItensCozinha 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   13785
      Top             =   9705
   End
   Begin MSComctlLib.ImageList imlBotões 
      Left            =   14805
      Top             =   9555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3230
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":338C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3644
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":451C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Barra 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   21200
      _ExtentX        =   37386
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlBotões"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoNovo"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoBarra01"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoImprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoEditar"
            Object.ToolTipText     =   "Editar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoExcluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoBarra02"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoAgrupar"
            Object.ToolTipText     =   "Agrupar"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoDetalhes"
            Object.ToolTipText     =   "Exibir Detalhes"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btoAtualizar"
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoClassificar"
            Object.ToolTipText     =   "Classificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btoCalculadora"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoConfig"
            Object.ToolTipText     =   "Configurações"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoLogin"
            Object.ToolTipText     =   "Login"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoAjuda"
            Object.ToolTipText     =   "Ajuda"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwEnviarItensCozinha 
      Height          =   360
      Left            =   15975
      TabIndex        =   1
      Top             =   9765
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   635
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
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
         Text            =   "IdTipoComanda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IdCardapio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DescricaoItem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "QuantidadeItemCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TempoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "IdCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "DescricaoCozinha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Capacidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "DataConfirmacaoPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "UltimaDataPrevistaPreparo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "DataPrevistaPreparo"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   10515
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   27755
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "13/03/2024"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "12:32"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuFileSair 
         Caption         =   "&Sair do Sistema"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Tabelas"
      Begin VB.Menu mnuCozinha 
         Caption         =   "&Cozinhas"
      End
      Begin VB.Menu mnuCardapio 
         Caption         =   "C&ardápio"
      End
      Begin VB.Menu mnuTiposComanda 
         Caption         =   "C&omandas"
      End
   End
   Begin VB.Menu mnuMovimentacao 
      Caption         =   "&Movimentação"
      Visible         =   0   'False
      Begin VB.Menu mnuComanda 
         Caption         =   "Comandas"
      End
      Begin VB.Menu mnuFinalizarItemcozinha 
         Caption         =   "Itens Cozinha"
      End
      Begin VB.Menu mnuEstatisticas 
         Caption         =   "Estatistica"
      End
   End
   Begin VB.Menu mnuPopComandas 
      Caption         =   "mnuPopComandas"
      Visible         =   0   'False
      Begin VB.Menu mnuNovaComanda 
         Caption         =   "&Nova comanda"
      End
      Begin VB.Menu mnuCancelarComanda 
         Caption         =   "&Cancelar comanda"
      End
   End
   Begin VB.Menu mnuPopEstatisticas 
      Caption         =   "mnuPopEstatisticas"
      Visible         =   0   'False
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_ComandasNegocios As New clsComandasNegocios
Dim vop_ItemComandoNegocios As New clsComandaItemNegocios

'Variaveis de controle do form
Dim vil_IdComanda As Long
Dim vip_ItemListaComanda As Long





'Form
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

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo TrataErros
    If MsgBox(MSG01, Style10, Title01) = vbYes Then
       Set frmMain = Nothing
       End
    Else
      Cancel = 1
    End If
TrataErros:
    If Err.Number = 3420 Then End
   
End Sub

Private Sub Form_Resize()

   Call HabilitaMenuItem
    
End Sub

Private Sub imgItem_Click(Index As Integer)
    Select Case Index
        Case 1
            PopupMenu mnuPopComandas, , Screen.TwipsPerPixelX + 1600, Screen.TwipsPerPixelY + 740
        Case 3
            PopupMenu mnuPopEstatisticas, , Screen.TwipsPerPixelX + 1600, Screen.TwipsPerPixelY + 740
    End Select
End Sub

Private Sub mnuNovaComanda_Click()
On Error GoTo TrataErro
    DoEvents
    Load frmComandas
    DoEvents
    frmComandas.Show vbModal
    '
    txtLocalizarComanda.text = Empty
    cmbLocalizarComanda.ListIndex = 0
    'Carregar Grid
    Call CarregarGridComandaMain
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub

Private Sub mnuCancelarComanda_Click()
    If vil_IdComanda = 0 Then Exit Sub
    
    If MsgBox("Confirma o cancelamento da Comanda No. " & vil_IdComanda & " ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_ComandasNegocios = New clsComandasNegocios
          vop_ComandasNegocios.IdComandas = vil_IdComanda
          If vop_ComandasNegocios.ExcluirComanda() = True Then
             txtLocalizarComanda.text = Empty
             cmbLocalizarComanda.ListIndex = 0
             'Carregar Grid
             Call CarregarGridComandaMain
          End If
      Set vop_ComandasNegocios = Nothing
    End If

End Sub

Private Sub mnuImprimir_Click()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_ComandasNegocios = New clsComandasNegocios
        
        vbl_Carregar = vop_ComandasNegocios.ImprimirEstatisticas(cmbLocalizarItemEstatistica.ListIndex)
        If vbl_Carregar = True Then
           DataReport1.Show vbModal
        End If
        
    Set vop_ComandasNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
    End If
    
End Sub

Private Sub mnuComanda_Click()
On Error GoTo TrataErros

    DoEvents
    Load frmComandasLista
    DoEvents
    frmComandasLista.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
    End If
End Sub

Private Sub mnuCozinha_Click()
On Error GoTo TrataErros

    DoEvents
    Load frmCozinhasLista
    DoEvents
    frmCozinhasLista.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
    End If
End Sub

Private Sub mnuCardapio_Click()
On Error GoTo TrataErros

    DoEvents
    Load frmCardapioLista
    DoEvents
    frmCardapioLista.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
    End If
End Sub

Private Sub mnuTiposComanda_Click()
On Error GoTo TrataErros

    DoEvents
    Load frmTiposComandasLista
    DoEvents
    frmTiposComandasLista.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
    End If
End Sub

Private Sub mnuEstatisticas_Click()
On Error GoTo TrataErros

    DoEvents
    Load frmEstatisticaItensLista
    DoEvents
    frmEstatisticaItensLista.Show vbModal
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
    End If
End Sub

Private Sub mnuFileSair_Click()
On Error GoTo TrataErros
    If MsgBox(MSG01, Style10, Title01) = vbYes Then
       Set frmMain = Nothing
       End
    End If
TrataErros:
    If Err.Number = 3420 Then End
End Sub

Private Sub mnuFinalizarItemcozinha_Click()
   frmCozinhaItensLista.Show vbModal
End Sub

Private Sub Barra_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
         Case "btoAtualizar"
            If MenuItem(1).Visible = True Then
               Call CarregarGridComanda
            ElseIf MenuItem(2).Visible = True Then
               Call CarregarGridItensCozinha
            ElseIf MenuItem(2).Visible = True Then
               Call CarregarGridEstatisticas
            End If
        Case "btoCalculadora"
            Call Calculadora
    End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Screen.MousePointer = vbNoDrop
   'Screen.MousePointer = vbDefault
End Sub

Private Sub tmrEnviarItensCozinha_Timer()
   Call EnviarItensCozinha
End Sub






'Comandas
Private Sub cmdComandas_Click()
   '
   Call CarregarGridComanda
   cmbLocalizarComanda.ListIndex = 0
   'Comanda
   Call HabilitaMenuItem
   MenuItem(0).Visible = False
   MenuItem(1).Visible = True
   frmComanda.Width = mainBorder.Width - 1630 - 90
   dtgComandas.Width = mainBorder.Width - 1630 - 110
   dtgComandas.Height = mainBorder.Height - 960
   '
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

Private Sub cmbLocalizarComanda_Change()
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
    '
    txtLocalizarComanda.text = Empty
    cmbLocalizarComanda.ListIndex = 0
    'Carregar Grid
    Call CarregarGridComanda

TrataErros:
    If Err.Number <> 0 Then Exit Sub

End Sub








'Itens Cozinha
Private Sub cmdItensCozinha_Click()
   '
   Call CarregarGridItensCozinha
   cmbLocalizarItemCozinha.ListIndex = 0
   'Comanda
   Call HabilitaMenuItem
   MenuItem(0).Visible = False
   MenuItem(2).Visible = True
   frmItensCozinha.Width = mainBorder.Width - 1630 - 90
   dtgCozinhaItens.Width = mainBorder.Width - 1630 - 110
   dtgCozinhaItens.Height = mainBorder.Height - 960
   '
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

Private Sub dtgCozinhaItens_DblClick()
On Error GoTo TrataErros

   Call frmComandaItens.Form_Load
   Call frmComandaItens.Editar(Me, vip_ItemListaComanda)
    '
    txtLocalizarItemCozinha.text = Empty
    cmbLocalizarItemCozinha.ListIndex = 0
    'Carregar Grid
    Call CarregarGridItensCozinha

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









'Estatiscicas
Private Sub cmdEstatisticas_Click()
   '
   Call CarregarGridEstatisticas
   cmbLocalizarItemEstatistica.ListIndex = 0
   'Comanda
   Call HabilitaMenuItem
   MenuItem(0).Visible = False
   MenuItem(3).Visible = True
   frmEstatistica.Width = mainBorder.Width - 1630 - 90
   dtgEstatisticaItensDias.Width = mainBorder.Width - 1630 - 110
   dtgEstatisticaItensDias.Height = mainBorder.Height - 960
   '
   dtgEstatisticaItens.Width = mainBorder.Width - 1630 - 110
   dtgEstatisticaItens.Height = mainBorder.Height - 960
   '
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



















'Function
Private Function CarregarGridComanda()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_ComandasNegocios = New clsComandasNegocios
        
        vbl_Carregar = vop_ComandasNegocios.CarregarGridComandaMainRS(adoComandas, txtLocalizarComanda.text, IIf(cmbLocalizarComanda.ListIndex < 0, 0, cmbLocalizarComanda.ListIndex))
        If vbl_Carregar = True Then
           dtgComandas.Refresh
        End If
        
    Set vop_ComandasNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ComandasNegocios = Nothing
       Exit Function
    End If

End Function

Private Function CarregarGridItensCozinha()
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
       Exit Function
    End If

End Function

Private Function CarregarGridEstatisticas()
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
       Exit Function
    End If

End Function

Private Function EnviarItensCozinha() As Boolean

   'Envia item para cozinha e processamento das horas previstas
   Set vop_ComandasNegocios = New clsComandasNegocios
          If vop_ComandasNegocios.EnviarItensCozinha(lvwEnviarItensCozinha) = True Then
          'Atualiza dados entrada e saida
          frmComandasLista.CarregarGrid
       End If
   Set vop_ComandasNegocios = Nothing

End Function

Private Function HabilitaMenuItem()
Dim vil_Count As Integer

    If Me.WindowState = vbMaximized Then
        Barra.Refresh
        mainBorder.Height = Me.Height - 1330
        mainBorder.Width = Me.Width
        Barra.Width = Me.Width
        For vil_Count = 0 To MenuItem.Count - 1
            MenuItem(vil_Count).Visible = False
            MenuItem(vil_Count).Left = 1560
            MenuItem(vil_Count).Top = 0
            MenuItem(vil_Count).Height = mainBorder.Height
            MenuItem(vil_Count).Width = mainBorder.Width - 1630
        Next
        MenuItem(0).Visible = True
    End If

End Function




