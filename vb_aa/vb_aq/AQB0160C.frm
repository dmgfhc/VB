VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQB0160C 
   Caption         =   "�����淶��ƽ����ѯ - AQB0160C"
   ClientHeight    =   9165
   ClientLeft      =   210
   ClientTop       =   1395
   ClientWidth     =   15345
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   15345
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel7 
      Align           =   1  'Align Top
      Height          =   2610
      Left            =   0
      TabIndex        =   133
      Top             =   3645
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   4604
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin TabDlg.SSTab SSTab1 
         Height          =   2460
         Left            =   0
         TabIndex        =   134
         Top             =   120
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   4339
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "�к�������"
         TabPicture(0)   =   "AQB0160C.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "SSFrame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "�а�����"
         TabPicture(1)   =   "AQB0160C.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "SSFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "���������"
         TabPicture(2)   =   "AQB0160C.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSFrame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin Threed.SSFrame SSFrame1 
            Height          =   1755
            Left            =   -74880
            TabIndex        =   135
            Top             =   360
            Width           =   15045
            _ExtentX        =   26538
            _ExtentY        =   3096
            _Version        =   196609
            Begin VB.TextBox TXT_SHEAR_C1 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   13545
               TabIndex        =   235
               Top             =   480
               Width           =   915
            End
            Begin VB.TextBox txt_STEAM_TEMP_MIN 
               Height          =   300
               Left            =   6390
               TabIndex        =   231
               Top             =   1335
               Width           =   525
            End
            Begin VB.TextBox txt_STEAM_TEMP_MAX 
               Height          =   300
               Left            =   6930
               TabIndex        =   230
               Top             =   1335
               Width           =   525
            End
            Begin VB.TextBox txt_STEAM_TEMP_TGT 
               Height          =   300
               Left            =   5790
               TabIndex        =   229
               Top             =   1335
               Width           =   585
            End
            Begin VB.TextBox txt_MILL_RATET2 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2880
               MaxLength       =   3
               TabIndex        =   228
               Top             =   1335
               Width           =   645
            End
            Begin VB.TextBox txt_MILL_TMP_TGT 
               Height          =   300
               Left            =   5790
               MaxLength       =   4
               TabIndex        =   155
               Top             =   75
               Width           =   585
            End
            Begin VB.TextBox txt_CHG_TMP_DEF_SC 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5790
               MaxLength       =   4
               TabIndex        =   154
               Top             =   390
               Width           =   1305
            End
            Begin VB.TextBox txt_COOL_TMP_RATE 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   9690
               TabIndex        =   153
               Top             =   705
               Width           =   915
            End
            Begin VB.TextBox txt_MILL_RATET1 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2880
               MaxLength       =   3
               TabIndex        =   152
               Top             =   1020
               Width           =   645
            End
            Begin VB.TextBox txt_MILL_TMPT2 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   151
               Top             =   1335
               Width           =   1035
            End
            Begin VB.TextBox txt_MILL_TMPT1 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   150
               Top             =   1020
               Width           =   1035
            End
            Begin VB.TextBox txt_MILL_TIME 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   149
               Top             =   75
               Width           =   705
            End
            Begin VB.TextBox txt_CHG_TMP_DEF_TAPE 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5790
               MaxLength       =   4
               TabIndex        =   148
               Top             =   705
               Width           =   1305
            End
            Begin VB.TextBox txt_CHG_TMP_TGT 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   4
               TabIndex        =   147
               Top             =   390
               Width           =   705
            End
            Begin VB.TextBox txt_HOT_USE 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   9690
               MaxLength       =   1
               TabIndex        =   146
               Top             =   1335
               Width           =   465
            End
            Begin VB.TextBox txt_COOL_CTL_TYP 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   5790
               MaxLength       =   1
               TabIndex        =   145
               Top             =   1020
               Width           =   465
            End
            Begin VB.TextBox txt_COOL_CTL_NAME 
               Enabled         =   0   'False
               Height          =   300
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   144
               Top             =   1020
               Width           =   1215
            End
            Begin VB.TextBox txt_COOL_WAY 
               Height          =   300
               Left            =   9690
               MaxLength       =   1
               TabIndex        =   143
               Top             =   75
               Width           =   465
            End
            Begin VB.TextBox txt_COOL_WAY_NAME 
               Enabled         =   0   'False
               Height          =   300
               Left            =   10170
               Locked          =   -1  'True
               TabIndex        =   142
               Top             =   75
               Width           =   1215
            End
            Begin VB.TextBox txt_CR_CD 
               Height          =   300
               Left            =   1830
               MaxLength       =   1
               TabIndex        =   141
               Top             =   705
               Width           =   465
            End
            Begin VB.TextBox txt_CR_NAME 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2310
               Locked          =   -1  'True
               TabIndex        =   140
               Top             =   705
               Width           =   1215
            End
            Begin VB.TextBox txt_CHG_TMP_MIN 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2550
               TabIndex        =   139
               Top             =   390
               Width           =   495
            End
            Begin VB.TextBox txt_CHG_TMP_MAX 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   3050
               TabIndex        =   138
               Top             =   390
               Width           =   495
            End
            Begin VB.TextBox txt_STEAM_FL 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   13530
               TabIndex        =   137
               Top             =   75
               Width           =   465
            End
            Begin VB.TextBox txt_STEAM_RATE 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   14010
               TabIndex        =   136
               Top             =   75
               Width           =   465
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_TGT 
               Height          =   300
               Left            =   9690
               TabIndex        =   156
               Top             =   390
               Width           =   495
               _Version        =   262145
               _ExtentX        =   873
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MAX 
               Height          =   300
               Left            =   10695
               TabIndex        =   157
               Top             =   390
               Width           =   495
               _Version        =   262145
               _ExtentX        =   873
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   3
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MIN 
               Height          =   300
               Left            =   10215
               TabIndex        =   158
               Top             =   390
               Width           =   495
               _Version        =   262145
               _ExtentX        =   873
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   3
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MAX 
               Height          =   300
               Left            =   6930
               TabIndex        =   159
               Top             =   75
               Width           =   540
               _Version        =   262145
               _ExtentX        =   952
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MIN 
               Height          =   300
               Left            =   6375
               TabIndex        =   160
               Top             =   75
               Width           =   540
               _Version        =   262145
               _ExtentX        =   952
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   2
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtThousands    =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   2
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   22
               Left            =   30
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "���Ƽ����S��"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   23
               Left            =   30
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "ƽ����¯�¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   24
               Left            =   30
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "һ�׶��¶�/��ȱ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   25
               Left            =   30
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "���׶��¶�/��ȱ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   26
               Left            =   3990
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "����ͷβ�²�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   27
               Left            =   3990
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "����Ŀ���¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   28
               Left            =   3990
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��������/�����²�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   29
               Left            =   7905
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴĿ���¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   30
               Left            =   7905
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴ����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   31
               Left            =   3990
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "������ȴ"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_BED_TMP_TGT 
               Height          =   300
               Left            =   9690
               TabIndex        =   161
               Top             =   1020
               Width           =   735
               _Version        =   262145
               _ExtentX        =   1296
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   32
               Left            =   7905
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�䴲Ŀ���¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   33
               Left            =   7905
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "ʹ���Ƚ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   34
               Left            =   30
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��������"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   35
               Left            =   7905
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴ����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   97
               Left            =   11745
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�Ƿ�������ȴ/����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   96
               Left            =   3990
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "������ȴ�¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   119
               Left            =   11760
               Top             =   480
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�����¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   1725
            Left            =   120
            TabIndex        =   162
            Top             =   360
            Width           =   15045
            _ExtentX        =   26538
            _ExtentY        =   3043
            _Version        =   196609
            Begin VB.TextBox TXT_SHEAR_C3 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   13545
               TabIndex        =   237
               Top             =   390
               Width           =   915
            End
            Begin VB.TextBox txt_CR_NAME_Z 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2310
               Locked          =   -1  'True
               TabIndex        =   186
               Top             =   705
               Width           =   1215
            End
            Begin VB.TextBox txt_CR_CD_Z 
               Height          =   300
               Left            =   1830
               MaxLength       =   1
               TabIndex        =   185
               Top             =   705
               Width           =   465
            End
            Begin VB.TextBox txt_COOL_WAY_NAME_Z 
               Enabled         =   0   'False
               Height          =   300
               Left            =   10170
               Locked          =   -1  'True
               TabIndex        =   184
               Top             =   75
               Width           =   1215
            End
            Begin VB.TextBox txt_COOL_WAY_Z 
               Height          =   300
               Left            =   9690
               MaxLength       =   1
               TabIndex        =   183
               Top             =   75
               Width           =   465
            End
            Begin VB.TextBox txt_COOL_CTL_NAME_Z 
               Enabled         =   0   'False
               Height          =   300
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   182
               Top             =   1020
               Width           =   1215
            End
            Begin VB.TextBox txt_COOL_CTL_TYP_Z 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   5790
               MaxLength       =   1
               TabIndex        =   181
               Top             =   1020
               Width           =   465
            End
            Begin VB.TextBox txt_HOT_USE_Z 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   9690
               MaxLength       =   1
               TabIndex        =   180
               Top             =   1335
               Width           =   465
            End
            Begin VB.TextBox txt_CHG_TMP_TGT_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   4
               TabIndex        =   179
               Top             =   390
               Width           =   705
            End
            Begin VB.TextBox txt_CHG_TMP_DEF_TAPE_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5790
               MaxLength       =   4
               TabIndex        =   178
               Top             =   705
               Width           =   1305
            End
            Begin VB.TextBox txt_MILL_TIME_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   177
               Top             =   75
               Width           =   705
            End
            Begin VB.TextBox txt_MILL_TMPT1_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   176
               Top             =   1020
               Width           =   1035
            End
            Begin VB.TextBox txt_MILL_TMPT2_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   175
               Top             =   1335
               Width           =   1035
            End
            Begin VB.TextBox txt_MILL_RATET1_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2880
               MaxLength       =   3
               TabIndex        =   174
               Top             =   1020
               Width           =   645
            End
            Begin VB.TextBox txt_MILL_RATET2_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2880
               MaxLength       =   3
               TabIndex        =   173
               Top             =   1335
               Width           =   645
            End
            Begin VB.TextBox txt_COOL_TMP_RATE_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   9690
               TabIndex        =   172
               Top             =   705
               Width           =   915
            End
            Begin VB.TextBox txt_CHG_TMP_DEF_SC_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5790
               MaxLength       =   4
               TabIndex        =   171
               Top             =   390
               Width           =   1305
            End
            Begin VB.TextBox txt_MILL_TMP_TGT_Z 
               Height          =   300
               Left            =   5790
               MaxLength       =   4
               TabIndex        =   170
               Top             =   75
               Width           =   585
            End
            Begin VB.TextBox txt_CHG_TMP_MIN_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2550
               TabIndex        =   169
               Top             =   390
               Width           =   495
            End
            Begin VB.TextBox txt_CHG_TMP_MAX_Z 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   3050
               TabIndex        =   168
               Top             =   390
               Width           =   495
            End
            Begin VB.TextBox txt_STEAM_TEMP_MIN_Z 
               Height          =   300
               Left            =   6390
               TabIndex        =   167
               Top             =   1335
               Width           =   525
            End
            Begin VB.TextBox txt_STEAM_TEMP_MAX_Z 
               Height          =   300
               Left            =   6930
               TabIndex        =   166
               Top             =   1335
               Width           =   525
            End
            Begin VB.TextBox txt_STEAM_RATE_Z 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   14010
               TabIndex        =   165
               Top             =   75
               Width           =   465
            End
            Begin VB.TextBox txt_STEAM_FL_Z 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   13530
               TabIndex        =   164
               Top             =   75
               Width           =   465
            End
            Begin VB.TextBox txt_STEAM_TEMP_TGT_Z 
               Height          =   300
               Left            =   5790
               TabIndex        =   163
               Top             =   1335
               Width           =   585
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_TGT_Z 
               Height          =   300
               Left            =   9690
               TabIndex        =   187
               Top             =   390
               Width           =   495
               _Version        =   262145
               _ExtentX        =   873
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MAX_Z 
               Height          =   300
               Left            =   10695
               TabIndex        =   188
               Top             =   390
               Width           =   495
               _Version        =   262145
               _ExtentX        =   873
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   3
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MIN_Z 
               Height          =   300
               Left            =   10215
               TabIndex        =   189
               Top             =   390
               Width           =   495
               _Version        =   262145
               _ExtentX        =   873
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   3
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MAX_Z 
               Height          =   300
               Left            =   6930
               TabIndex        =   190
               Top             =   75
               Width           =   540
               _Version        =   262145
               _ExtentX        =   952
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MIN_Z 
               Height          =   300
               Left            =   6375
               TabIndex        =   191
               Top             =   75
               Width           =   540
               _Version        =   262145
               _ExtentX        =   952
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   2
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtThousands    =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   2
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   40
               Left            =   30
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "���Ƽ����S��"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   42
               Left            =   30
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "ƽ����¯�¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   49
               Left            =   30
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "һ�׶��¶�/��ȱ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   50
               Left            =   30
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "���׶��¶�/��ȱ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   51
               Left            =   3990
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "����ͷβ�²�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   52
               Left            =   3990
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "����Ŀ���¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   53
               Left            =   3990
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��������/�����²�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   54
               Left            =   7905
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴĿ���¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   55
               Left            =   7905
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴ����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   56
               Left            =   3990
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "������ȴ"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_BED_TMP_TGT_Z 
               Height          =   300
               Left            =   9690
               TabIndex        =   192
               Top             =   1020
               Width           =   735
               _Version        =   262145
               _ExtentX        =   1296
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   57
               Left            =   7905
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�䴲Ŀ���¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   58
               Left            =   7905
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "ʹ���Ƚ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   59
               Left            =   30
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��������"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   60
               Left            =   7905
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴ����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   98
               Left            =   3990
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "������ȴ�¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   99
               Left            =   11745
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�Ƿ�������ȴ/����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   121
               Left            =   11760
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�����¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   1995
            Left            =   -74880
            TabIndex        =   193
            Top             =   360
            Width           =   15045
            _ExtentX        =   26538
            _ExtentY        =   3519
            _Version        =   196609
            Begin VB.TextBox TXT_SHEAR_C2 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   13665
               TabIndex        =   236
               Top             =   1320
               Width           =   915
            End
            Begin VB.TextBox txt_PP_SLOW_COOL_TIME 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   13680
               MaxLength       =   4
               TabIndex        =   234
               Top             =   1010
               Width           =   1305
            End
            Begin VB.TextBox txt_SL_SLOW_COOL_TIME 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   13680
               MaxLength       =   4
               TabIndex        =   233
               Top             =   690
               Width           =   1305
            End
            Begin VB.TextBox txt_COOLING_RATE_DQ_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   13660
               MaxLength       =   4
               TabIndex        =   227
               Top             =   375
               Width           =   1305
            End
            Begin VB.TextBox txt_AIM_COOL_TEMP_DQ_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   13650
               MaxLength       =   4
               TabIndex        =   226
               Top             =   60
               Width           =   1305
            End
            Begin VB.TextBox txt_SLOW_COOL_TEMP_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   10050
               MaxLength       =   4
               TabIndex        =   225
               Top             =   1665
               Width           =   1305
            End
            Begin VB.TextBox txt_COOL_STR_TEMP_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   10050
               MaxLength       =   4
               TabIndex        =   224
               Top             =   1350
               Width           =   1305
            End
            Begin VB.TextBox txt_STEAM_TEMP_TGT_K 
               Height          =   300
               Left            =   5890
               TabIndex        =   217
               Top             =   1020
               Width           =   585
            End
            Begin VB.TextBox txt_STEAM_FL_K 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   10050
               TabIndex        =   216
               Top             =   1035
               Width           =   465
            End
            Begin VB.TextBox txt_STEAM_RATE_K 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   10530
               TabIndex        =   215
               Top             =   1035
               Width           =   465
            End
            Begin VB.TextBox txt_STEAM_TEMP_MAX_K 
               Height          =   300
               Left            =   7000
               TabIndex        =   214
               Top             =   1020
               Width           =   585
            End
            Begin VB.TextBox txt_STEAM_TEMP_MIN_K 
               Height          =   300
               Left            =   6480
               TabIndex        =   213
               Top             =   1020
               Width           =   525
            End
            Begin VB.TextBox txt_CHG_TMP_MAX_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   3050
               TabIndex        =   212
               Top             =   390
               Width           =   495
            End
            Begin VB.TextBox txt_CHG_TMP_MIN_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2550
               TabIndex        =   211
               Top             =   390
               Width           =   495
            End
            Begin VB.TextBox txt_MILL_TMP_TGT_K 
               Height          =   300
               Left            =   1830
               MaxLength       =   4
               TabIndex        =   210
               Top             =   1635
               Width           =   585
            End
            Begin VB.TextBox txt_CHG_TMP_DEF_SC_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5890
               MaxLength       =   4
               TabIndex        =   209
               Top             =   75
               Width           =   1665
            End
            Begin VB.TextBox txt_COOL_TMP_RATE_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   10050
               TabIndex        =   208
               Top             =   75
               Width           =   915
            End
            Begin VB.TextBox txt_MILL_RATET2_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2880
               MaxLength       =   3
               TabIndex        =   207
               Top             =   1335
               Width           =   645
            End
            Begin VB.TextBox txt_MILL_RATET1_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2880
               MaxLength       =   3
               TabIndex        =   206
               Top             =   1020
               Width           =   645
            End
            Begin VB.TextBox txt_MILL_TMPT2_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   205
               Top             =   1335
               Width           =   1035
            End
            Begin VB.TextBox txt_MILL_TMPT1_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   204
               Top             =   1020
               Width           =   1035
            End
            Begin VB.TextBox txt_MILL_TIME_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   11
               TabIndex        =   203
               Top             =   75
               Width           =   705
            End
            Begin VB.TextBox txt_CHG_TMP_DEF_TAPE_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5890
               MaxLength       =   4
               TabIndex        =   202
               Top             =   390
               Width           =   1665
            End
            Begin VB.TextBox txt_CHG_TMP_TGT_K 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1830
               MaxLength       =   4
               TabIndex        =   201
               Top             =   390
               Width           =   705
            End
            Begin VB.TextBox txt_HOT_USE_K 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   10050
               MaxLength       =   1
               TabIndex        =   200
               Top             =   705
               Width           =   945
            End
            Begin VB.TextBox txt_COOL_CTL_TYP_K 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   5890
               MaxLength       =   1
               TabIndex        =   199
               Top             =   705
               Width           =   465
            End
            Begin VB.TextBox txt_COOL_CTL_NAME_K 
               Enabled         =   0   'False
               Height          =   300
               Left            =   6355
               Locked          =   -1  'True
               TabIndex        =   198
               Top             =   705
               Width           =   1215
            End
            Begin VB.TextBox txt_COOL_WAY_K 
               Height          =   300
               Left            =   5880
               MaxLength       =   1
               TabIndex        =   197
               Top             =   1335
               Width           =   465
            End
            Begin VB.TextBox txt_COOL_WAY_NAME_K 
               Enabled         =   0   'False
               Height          =   300
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   196
               Top             =   1335
               Width           =   1215
            End
            Begin VB.TextBox txt_CR_CD_K 
               Height          =   300
               Left            =   1830
               MaxLength       =   1
               TabIndex        =   195
               Top             =   705
               Width           =   465
            End
            Begin VB.TextBox txt_CR_NAME_K 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2310
               Locked          =   -1  'True
               TabIndex        =   194
               Top             =   705
               Width           =   1215
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_TGT_K 
               Height          =   300
               Left            =   5865
               TabIndex        =   218
               Top             =   1635
               Width           =   570
               _Version        =   262145
               _ExtentX        =   1005
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MAX_K 
               Height          =   300
               Left            =   7000
               TabIndex        =   219
               Top             =   1635
               Width           =   570
               _Version        =   262145
               _ExtentX        =   1005
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   3
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MIN_K 
               Height          =   300
               Left            =   6430
               TabIndex        =   220
               Top             =   1635
               Width           =   570
               _Version        =   262145
               _ExtentX        =   1005
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   3
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MAX_K 
               Height          =   300
               Left            =   2970
               TabIndex        =   221
               Top             =   1635
               Width           =   540
               _Version        =   262145
               _ExtentX        =   952
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MIN_K 
               Height          =   300
               Left            =   2415
               TabIndex        =   222
               Top             =   1635
               Width           =   540
               _Version        =   262145
               _ExtentX        =   952
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   2
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtThousands    =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   2
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   36
               Left            =   30
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "���Ƽ����S��"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   61
               Left            =   30
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "ƽ����¯�¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   100
               Left            =   30
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "һ�׶��¶�/��ȱ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   101
               Left            =   30
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "���׶��¶�/��ȱ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   102
               Left            =   4110
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "����ͷβ�²�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   103
               Left            =   30
               Top             =   1635
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "����Ŀ���¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   104
               Left            =   4110
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��������/�����²�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   105
               Left            =   4110
               Top             =   1635
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴĿ���¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   106
               Left            =   8265
               Top             =   75
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴ����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   107
               Left            =   4110
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "������ȴ"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin CSTextLibCtl.sidbEdit txt_COOL_BED_TMP_TGT_K 
               Height          =   300
               Left            =   10050
               TabIndex        =   223
               Top             =   390
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   529
               _StockProps     =   125
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoScroll      =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   ""
               StartText.x     =   3
               StartText.y     =   2
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   108
               Left            =   8265
               Top             =   390
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�䴲Ŀ���¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   109
               Left            =   8265
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "ʹ���Ƚ�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   110
               Left            =   30
               Top             =   705
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��������"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   3990
               Left            =   4110
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ȴ����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   112
               Left            =   4110
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "������ȴ�¶�/���"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   113
               Left            =   8265
               Top             =   1020
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�Ƿ�������ȴ/����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   114
               Left            =   8265
               Top             =   1335
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��ԥ�¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   115
               Left            =   8265
               Top             =   1680
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��Ʒ�����¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   116
               Left            =   11865
               Top             =   375
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "DQĿ����ȴ����"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   117
               Left            =   11865
               Top             =   60
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "DQĿ����ȴ�¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   111
               Left            =   11865
               Top             =   690
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��������ʱ��"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   118
               Left            =   11865
               Top             =   1010
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "��Ʒ����ʱ��"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   120
               Left            =   11880
               Top             =   1320
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               Caption         =   "�����¶�"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Align           =   1  'Align Top
      Height          =   1545
      Left            =   0
      TabIndex        =   66
      Top             =   7335
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   2725
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_DRW_TEMP_AIM3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_AIM2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_AIM1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_MIN3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_MIN2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   121
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_MIN1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   120
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_MIN3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_MIN2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_MIN1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_AIM3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_AIM2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   115
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_AIM1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_MAX3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_MAX2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_2F_MAX1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_AIM3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_AIM2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_AIM1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_MAX3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_MAX2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_MAX1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_SHOT_BLAST_NAME 
         Height          =   600
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   104
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox txt_SHOT_BLAST 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   60
         MaxLength       =   2
         TabIndex        =   103
         Top             =   585
         Width           =   1545
      End
      Begin VB.TextBox txt_HTM_COOL_TMP3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   12780
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   1230
         Width           =   1035
      End
      Begin VB.TextBox txt_HTM_COOL_TMP2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   12780
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   915
         Width           =   1035
      End
      Begin VB.TextBox txt_HTM_COOL_TMP1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   12780
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   585
         Width           =   1035
      End
      Begin VB.TextBox txt_HTM_COOL_TYP3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12210
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_COOL_TYP2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12210
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_COOL_TYP1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12210
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_MAX3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11070
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_MAX2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11070
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_MAX1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11070
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_MIN3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10500
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_MIN2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10500
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_DRW_TEMP_MIN1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10500
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_AIM3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9930
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_AIM2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9930
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_AIM1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9930
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_MAX3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_MAX2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_COOL_TIME_MAX1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_MIN3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5370
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_MIN2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5370
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TIME_1F_MIN1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5370
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_TGT3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_TGT2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_TGT1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_MAX3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_MAX2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_MAX1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_MIN3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_MIN2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_TEMP_MIN1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txt_MTH_COND3 
         Height          =   315
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   72
         Top             =   1230
         Width           =   1035
      End
      Begin VB.TextBox txt_MTH_COND2 
         Height          =   315
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   71
         Top             =   915
         Width           =   1035
      End
      Begin VB.TextBox txt_MTH_COND1 
         Height          =   315
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   70
         Top             =   585
         Width           =   1035
      End
      Begin VB.TextBox txt_HTM_METH3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2070
         MaxLength       =   1
         TabIndex        =   69
         Top             =   1230
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_METH2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2070
         MaxLength       =   1
         TabIndex        =   68
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txt_HTM_METH1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2070
         MaxLength       =   1
         TabIndex        =   67
         Top             =   585
         Width           =   555
      End
      Begin InDate.ULabel ULabel1 
         Height          =   555
         Index           =   63
         Left            =   2070
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel1 
         Height          =   555
         Index           =   64
         Left            =   2640
         Top             =   30
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   979
         Caption         =   "�ȴ�������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   66
         Left            =   1620
         Top             =   585
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         Caption         =   "1"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.26
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   67
         Left            =   1620
         Top             =   915
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         Caption         =   "2"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.26
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   68
         Left            =   1620
         Top             =   1230
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         Caption         =   "3"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.26
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   69
         Left            =   3660
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "���ȶ��¶�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   70
         Left            =   8790
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "��¯��ȴʱ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   71
         Left            =   10500
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "��¯�ְ��¶�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   73
         Left            =   3660
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "��С"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   74
         Left            =   4230
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   75
         Left            =   4800
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "Ŀ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   76
         Left            =   5370
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "��С"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   72
         Left            =   9360
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   77
         Left            =   9930
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "Ŀ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   78
         Left            =   10500
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "��С"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   79
         Left            =   11070
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   86
         Left            =   12210
         Top             =   30
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         Caption         =   "��ȴ����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   87
         Left            =   12210
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   88
         Left            =   12780
         Top             =   360
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         Caption         =   "���䴲�¶�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   555
         Index           =   62
         Left            =   60
         Top             =   30
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   979
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel1 
         Height          =   555
         Index           =   65
         Left            =   1620
         Top             =   30
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   979
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   80
         Left            =   5370
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "1#����פ��ʱ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   81
         Left            =   5940
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   82
         Left            =   6510
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "Ŀ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   83
         Left            =   7080
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "2#����פ��ʱ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   84
         Left            =   7650
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   85
         Left            =   8220
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "Ŀ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   89
         Left            =   7080
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "��С"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   90
         Left            =   8790
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "��С"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   255
         Index           =   91
         Left            =   11640
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         Caption         =   "Ŀ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   56
      Top             =   6255
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   1905
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_MILL_PLT 
         Height          =   315
         Left            =   14640
         TabIndex        =   131
         Top             =   690
         Visible         =   0   'False
         Width           =   540
      End
      Begin Threed.SSOption SSOp_C1 
         Height          =   315
         Left            =   11280
         TabIndex        =   128
         Top             =   75
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14804173
         Caption         =   "�к���"
      End
      Begin VB.TextBox txt_HCR_KND_NAME_1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   6930
         MaxLength       =   11
         TabIndex        =   127
         Top             =   90
         Width           =   2400
      End
      Begin VB.TextBox txt_HCR_KND_1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   6540
         MaxLength       =   1
         TabIndex        =   126
         Top             =   90
         Width           =   390
      End
      Begin VB.TextBox txt_UST_FL 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   11280
         MaxLength       =   4
         TabIndex        =   59
         Top             =   390
         Width           =   615
      End
      Begin VB.TextBox txt_UST_FL_NAME 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   11880
         TabIndex        =   58
         Top             =   390
         Width           =   2595
      End
      Begin VB.TextBox txt_MILL_STD_EDT_NO 
         Height          =   300
         Left            =   1890
         MaxLength       =   80
         TabIndex        =   57
         Top             =   720
         Width           =   12735
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   37
         Left            =   90
         Top             =   390
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "���ƺ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   38
         Left            =   4710
         Top             =   390
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "���ƿ���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   39
         Left            =   9480
         Top             =   390
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "UST����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin CSTextLibCtl.sidbEdit sdb_MILL_TGT_THK 
         Height          =   315
         Left            =   1890
         TabIndex        =   60
         Top             =   390
         Width           =   870
         _Version        =   262145
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_MILL_TGT_WID 
         Height          =   315
         Left            =   6540
         TabIndex        =   61
         Top             =   390
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_MILL_TGT_THK_MAX 
         Height          =   315
         Left            =   2760
         TabIndex        =   62
         Top             =   390
         Width           =   870
         _Version        =   262145
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_MILL_TGT_THK_MIN 
         Height          =   315
         Left            =   3660
         TabIndex        =   63
         Top             =   390
         Width           =   870
         _Version        =   262145
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_MILL_TGT_WID_MAX 
         Height          =   315
         Left            =   7560
         TabIndex        =   64
         Top             =   390
         Width           =   870
         _Version        =   262145
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_MILL_TGT_WID_MIN 
         Height          =   315
         Left            =   8460
         TabIndex        =   65
         Top             =   390
         Width           =   870
         _Version        =   262145
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel3 
         Height          =   300
         Index           =   0
         Left            =   90
         Top             =   720
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         Caption         =   "���ֹ淶�༭��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   300
         Index           =   44
         Left            =   90
         Top             =   75
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ul_STLGRD 
         Height          =   315
         Left            =   1890
         Top             =   75
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   300
         Index           =   92
         Left            =   4710
         Top             =   90
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         Caption         =   "������ʽ "
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   93
         Left            =   9480
         Top             =   75
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "���ù���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin Threed.SSOption SSOp_C2 
         Height          =   315
         Left            =   12360
         TabIndex        =   129
         Top             =   75
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14804173
         Caption         =   "�а峧"
      End
      Begin Threed.SSOption SSOp_All 
         Height          =   315
         Left            =   14160
         TabIndex        =   130
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14804173
         Caption         =   "�����Կ�"
      End
      Begin Threed.SSOption SSOp_C3 
         Height          =   315
         Left            =   13280
         TabIndex        =   232
         Top             =   75
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14804173
         Caption         =   "�����"
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Align           =   1  'Align Top
      Height          =   2250
      Left            =   0
      TabIndex        =   10
      Top             =   1395
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   3969
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_MLT_STD_EDT_NO 
         Height          =   300
         Left            =   7860
         MaxLength       =   80
         TabIndex        =   55
         Top             =   1890
         Width           =   7290
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   90
         TabIndex        =   35
         Top             =   375
         Width           =   7530
         Begin VB.ComboBox cob_MLT_PROC_CD_3 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "AQB0160C.frx":0054
            Left            =   2655
            List            =   "AQB0160C.frx":005E
            TabIndex        =   39
            Top             =   255
            Width           =   765
         End
         Begin VB.ComboBox cob_MLT_PROC_CD_2 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "AQB0160C.frx":006A
            Left            =   1875
            List            =   "AQB0160C.frx":0077
            TabIndex        =   38
            Top             =   255
            Width           =   765
         End
         Begin VB.ComboBox cob_MLT_PROC_CD_1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "AQB0160C.frx":0087
            Left            =   1095
            List            =   "AQB0160C.frx":0097
            TabIndex        =   37
            Top             =   255
            Width           =   765
         End
         Begin VB.TextBox txt_MLT_PROC_CD 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   36
            Top             =   720
            Width           =   2325
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   16
            Left            =   60
            Top             =   255
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            Caption         =   "��������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CSTextLibCtl.sidbEdit sdb_MLT_TMP_MIN 
            Height          =   315
            Left            =   4860
            TabIndex        =   40
            Top             =   195
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_MLT_TMP_MAX 
            Height          =   315
            Left            =   5715
            TabIndex        =   41
            Top             =   195
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_MLT_TMP_TGT 
            Height          =   315
            Left            =   6570
            TabIndex        =   42
            Top             =   195
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_LF_TEMP_MIN 
            Height          =   315
            Left            =   4860
            TabIndex        =   43
            Top             =   510
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_LF_TEMP_MAX 
            Height          =   315
            Left            =   5715
            TabIndex        =   44
            Top             =   510
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_LF_TEMP_TGT 
            Height          =   315
            Left            =   6570
            TabIndex        =   45
            Top             =   510
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_VD_TMP_MIN 
            Height          =   315
            Left            =   4860
            TabIndex        =   46
            Top             =   840
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_VD_TMP_MAX 
            Height          =   315
            Left            =   5715
            TabIndex        =   47
            Top             =   840
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_VD_TMP_TGT 
            Height          =   315
            Left            =   6570
            TabIndex        =   48
            Top             =   840
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_TD_TMP_MIN 
            Height          =   315
            Left            =   4860
            TabIndex        =   49
            Top             =   1155
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_TD_TMP_MAX 
            Height          =   315
            Left            =   5715
            TabIndex        =   50
            Top             =   1155
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_TD_TMP_TGT 
            Height          =   315
            Left            =   6570
            TabIndex        =   51
            Top             =   1155
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   18
            Left            =   3735
            Top             =   195
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "�����¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   19
            Left            =   3735
            Top             =   510
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "LF����¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   20
            Left            =   3735
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "VD����¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   21
            Left            =   3735
            Top             =   1155
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "�м���¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��С"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   5205
            TabIndex        =   54
            Top             =   15
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ŀ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   6885
            TabIndex        =   53
            Top             =   15
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   15
            Left            =   6045
            TabIndex        =   52
            Top             =   15
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1530
         Left            =   7665
         TabIndex        =   15
         Top             =   375
         Width           =   7530
         Begin VB.TextBox txt_MLT_PROC_CD2 
            DragMode        =   1  'Automatic
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   19
            Top             =   660
            Width           =   2325
         End
         Begin VB.ComboBox cob_MLT_PROC_CD2_1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "AQB0160C.frx":00AC
            Left            =   1095
            List            =   "AQB0160C.frx":00BC
            TabIndex        =   18
            Top             =   225
            Width           =   765
         End
         Begin VB.ComboBox cob_MLT_PROC_CD2_2 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "AQB0160C.frx":00D1
            Left            =   1875
            List            =   "AQB0160C.frx":00DE
            TabIndex        =   17
            Top             =   225
            Width           =   765
         End
         Begin VB.ComboBox cob_MLT_PROC_CD2_3 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "AQB0160C.frx":00F0
            Left            =   2640
            List            =   "AQB0160C.frx":00FA
            TabIndex        =   16
            Top             =   225
            Width           =   765
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   43
            Left            =   60
            Top             =   225
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            Caption         =   "��������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   45
            Left            =   3735
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "�����¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   46
            Left            =   3735
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "LF����¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   47
            Left            =   3735
            Top             =   510
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "RH����¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   48
            Left            =   3735
            Top             =   1170
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "�м���¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CSTextLibCtl.sidbEdit sdb_MLT_TMP_MIN2 
            Height          =   315
            Left            =   4860
            TabIndex        =   20
            Top             =   180
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_MLT_TMP_MAX2 
            Height          =   315
            Left            =   5715
            TabIndex        =   21
            Top             =   180
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_MLT_TMP_TGT2 
            Height          =   315
            Left            =   6570
            TabIndex        =   22
            Top             =   180
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_LF_TEMP_MIN2 
            Height          =   315
            Left            =   4860
            TabIndex        =   23
            Top             =   840
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_LF_TEMP_MAX2 
            Height          =   315
            Left            =   5715
            TabIndex        =   24
            Top             =   840
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_LF_TEMP_TGT2 
            Height          =   315
            Left            =   6570
            TabIndex        =   25
            Top             =   840
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_RH_TMP_MIN2 
            Height          =   315
            Left            =   4860
            TabIndex        =   26
            Top             =   510
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_RH_TMP_MAX2 
            Height          =   315
            Left            =   5715
            TabIndex        =   27
            Top             =   510
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_RH_TMP_TGT2 
            Height          =   315
            Left            =   6570
            TabIndex        =   28
            Top             =   510
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_TD_TMP_MIN2 
            Height          =   315
            Left            =   4860
            TabIndex        =   29
            Top             =   1170
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_TD_TMP_MAX2 
            Height          =   315
            Left            =   5715
            TabIndex        =   30
            Top             =   1170
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_TD_TMP_TGT2 
            Height          =   315
            Left            =   6570
            TabIndex        =   31
            Top             =   1170
            Width           =   885
            _Version        =   262145
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   6045
            TabIndex        =   34
            Top             =   0
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ŀ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   6885
            TabIndex        =   33
            Top             =   0
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��С"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   5205
            TabIndex        =   32
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.TextBox txt_HCR_KND 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3825
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1890
         Width           =   390
      End
      Begin VB.TextBox txt_HCR_KND_NAME 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   4275
         MaxLength       =   11
         TabIndex        =   13
         Top             =   1890
         Width           =   1905
      End
      Begin VB.TextBox txt_MLT_PROC_LINE 
         Height          =   310
         Left            =   1575
         MaxLength       =   1
         TabIndex        =   12
         Tag             =   "����·��"
         Top             =   30
         Width           =   315
      End
      Begin VB.TextBox txt_MLT_PROC_NAME 
         Enabled         =   0   'False
         Height          =   310
         Left            =   1890
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   30
         Width           =   5385
      End
      Begin InDate.ULabel ULabel3 
         Height          =   300
         Index           =   11
         Left            =   6330
         Top             =   1890
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   529
         Caption         =   "���ֹ淶�༭��"
         Alignment       =   0
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   13
         Left            =   105
         Top             =   1890
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel txt_STLGRD 
         Height          =   315
         Left            =   1155
         Top             =   1890
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   17
         Left            =   2790
         Top             =   1890
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "������ʽ "
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   18
         Left            =   90
         Top             =   30
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "����·��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1020
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   661
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_SLAB_WAIT_TIME 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   12165
         TabIndex        =   132
         Top             =   35
         Width           =   1410
      End
      Begin VB.TextBox txt_MILL_STD_NO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   8805
         MaxLength       =   6
         TabIndex        =   9
         Top             =   35
         Width           =   1410
      End
      Begin VB.TextBox txt_MLT_STD_NO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5445
         MaxLength       =   6
         TabIndex        =   8
         Top             =   35
         Width           =   1410
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   12
         Left            =   90
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   14
         Left            =   3600
         Top             =   30
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Caption         =   "���������淶���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   15
         Left            =   6960
         Top             =   30
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Caption         =   "���������淶���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel txt_STEEL_GRD 
         Height          =   315
         Left            =   1890
         Top             =   30
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   95
         Left            =   10320
         Top             =   30
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Caption         =   "����ʱ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   375
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   1138
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   1
         Left            =   60
         Top             =   0
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         Caption         =   "��׼����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   2
         Left            =   2190
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "�������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   3
         Left            =   3090
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         Caption         =   "Ʒ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   4
         Left            =   4320
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   5
         Left            =   5130
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   6
         Left            =   6150
         Top             =   0
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   7
         Left            =   9720
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   8
         Left            =   10920
         Top             =   0
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "�ͻ�����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   9
         Left            =   8610
         Top             =   0
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   10
         Left            =   12150
         Top             =   0
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "����Ҫ�����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   11
         Left            =   13380
         Top             =   0
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "������;"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_STDSPEC 
         Height          =   345
         Left            =   45
         Top             =   300
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_STDSPEC_YY 
         Height          =   345
         Left            =   2190
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_PROD_CD 
         Height          =   345
         Left            =   3090
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_ORD_THK 
         Height          =   345
         Left            =   4320
         Top             =   300
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_ORD_WID 
         Height          =   345
         Left            =   5130
         Top             =   300
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_ORD_LEN 
         Height          =   345
         Left            =   6150
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_DEL_TO_DATE 
         Height          =   345
         Left            =   9720
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_CUST_CD 
         Height          =   345
         Left            =   10920
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_UNIT_WGT 
         Height          =   345
         Left            =   8625
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_CUST_SPEC_NO 
         Height          =   345
         Left            =   12150
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_ENDUSE_CD 
         Height          =   345
         Left            =   13380
         Top             =   300
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   41
         Left            =   7380
         Top             =   0
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "����Ŀ����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel txt_THK_TGT 
         Height          =   345
         Left            =   7380
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         Caption         =   ""
         Alignment       =   1
         BackColor       =   15529975
         BackgroundStyle =   1
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   661
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_ORD_ITEM 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5250
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "���к�"
         Top             =   30
         Width           =   1125
      End
      Begin VB.TextBox txt_ORD_NO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1470
         MaxLength       =   11
         TabIndex        =   4
         Tag             =   "������"
         Top             =   30
         Width           =   2265
      End
      Begin VB.TextBox txt_ins_emp 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   7860
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "INS_EMP"
         Top             =   45
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txt_Design_STS 
         Height          =   315
         Left            =   6690
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   795
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   0
         Left            =   60
         Top             =   30
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   3810
         Top             =   30
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "���к�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox txt_KND 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   13905
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      Top             =   135
      Visible         =   0   'False
      Width           =   210
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   94
      Left            =   8970
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Caption         =   "���������淶���"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
End
Attribute VB_Name = "AQB0160C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       ��������
'-- Sub_System Name   �������
'-- Program Name      �淶��ƽ���޸ļ���ѯ
'-- Program ID        AQB0160C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.08.21
'-- Description       �淶��ƽ���޸ļ���ѯ
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE        EDITOR       DESCRIPTION
'   1.1   2005.01.25  HJD
'   1.2   2007.04.06  KIM.SUNG.HO
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"
    
'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'TOP
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
                Call Gp_Ms_Collection(txt_ORD_NO, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_ORD_ITEM, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_KND, "p", "n", " ", "i", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
            Call Gp_Ms_Collection(txt_Design_STS, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STDSPEC_YY, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_PROD_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_THK_TGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_DEL_TO_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CUST_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_UNIT_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_CUST_SPEC_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
             Call Gp_Ms_Collection(txt_STEEL_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_MLT_STD_NO, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_MILL_STD_NO, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 '20090525 SUN BIN START
        Call Gp_Ms_Collection(txt_SLAB_WAIT_TIME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 '20090525 SUN BIN END
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Body
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call Gp_Ms_Collection(txt_MLT_PROC_LINE, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_MLT_PROC_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_MLT_PROC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MLT_PROC_CD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_MLT_TMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_MLT_TMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_MLT_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_LF_TEMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_LF_TEMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_LF_TEMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_VD_TMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_VD_TMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_VD_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_TD_TMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_TD_TMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_TD_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
          Call Gp_Ms_Collection(sdb_MLT_TMP_MIN2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_MLT_TMP_MAX2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_MLT_TMP_TGT2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_LF_TEMP_MIN2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_LF_TEMP_MAX2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_LF_TEMP_TGT2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RH_TMP_MIN2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RH_TMP_MAX2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RH_TMP_TGT2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_TD_TMP_MIN2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_TD_TMP_MAX2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_TD_TMP_TGT2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'---------------------------------------------------------------------------------------------------
                Call Gp_Ms_Collection(txt_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_HCR_KND, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HCR_KND_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MLT_STD_EDT_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------�к���--------------------------------------
             Call Gp_Ms_Collection(txt_MILL_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_CHG_TMP_DEF_SC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_CHG_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 sun bin start
           Call Gp_Ms_Collection(txt_CHG_TMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_CHG_TMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 sun bin end
      Call Gp_Ms_Collection(txt_CHG_TMP_DEF_TAPE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_CR_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CR_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_MILL_TMPT1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_MILL_RATET1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_MILL_TMPT2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_MILL_RATET2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MILL_TMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MILL_TMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MILL_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_COOL_CTL_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_COOL_CTL_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_COOL_WAY, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_COOL_WAY_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_COOL_TMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_COOL_TMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_COOL_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_COOL_TMP_RATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_COOL_BED_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_HOT_USE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 SUN BIN START
        Call Gp_Ms_Collection(txt_STEAM_TEMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STEAM_TEMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STEAM_TEMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_STEAM_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STEAM_RATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 SUN BIN END
'-----------------------------------------�а�-------------------------------------------------
           Call Gp_Ms_Collection(txt_MILL_TIME_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CHG_TMP_DEF_SC_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_CHG_TMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 sun bin start
         Call Gp_Ms_Collection(txt_CHG_TMP_MIN_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_CHG_TMP_MAX_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 sun bin end
    
    Call Gp_Ms_Collection(txt_CHG_TMP_DEF_TAPE_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CR_CD_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_CR_NAME_Z, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MILL_TMPT1_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_MILL_RATET1_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MILL_TMPT2_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_MILL_RATET2_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_MIN_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_MAX_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_CTL_TYP_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_CTL_NAME_Z, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_COOL_WAY_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_WAY_NAME_Z, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_MIN_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_MAX_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_TMP_RATE_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_COOL_BED_TMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HOT_USE_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 SUN BIN START
      Call Gp_Ms_Collection(txt_STEAM_TEMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_STEAM_TEMP_MIN_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_STEAM_TEMP_MAX_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STEAM_FL_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_STEAM_RATE_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090525 SUN BIN END


'-----------------------------------------�����  ���� 2012.11.16 -------------------------------------------------
           Call Gp_Ms_Collection(txt_MILL_TIME_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CHG_TMP_DEF_SC_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_CHG_TMP_TGT_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_CHG_TMP_MIN_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_CHG_TMP_MAX_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_CHG_TMP_DEF_TAPE_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CR_CD_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_CR_NAME_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MILL_TMPT1_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_MILL_RATET1_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_MILL_TMPT2_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_MILL_RATET2_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_MIN_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_MAX_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_TGT_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_CTL_TYP_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_CTL_NAME_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_COOL_WAY_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_WAY_NAME_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_MIN_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_MAX_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_TGT_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_TMP_RATE_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_COOL_BED_TMP_TGT_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HOT_USE_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_STEAM_TEMP_TGT_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_STEAM_TEMP_MIN_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_STEAM_TEMP_MAX_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STEAM_FL_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_STEAM_RATE_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
       Call Gp_Ms_Collection(txt_COOL_STR_TEMP_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SLOW_COOL_TEMP_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_AIM_COOL_TEMP_DQ_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_COOLING_RATE_DQ_K, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_SL_SLOW_COOL_TIME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_PP_SLOW_COOL_TIME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'-----------------------------------------����-------------------------------------------------
          Call Gp_Ms_Collection(sdb_MILL_TGT_THK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_MILL_TGT_THK_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_MILL_TGT_THK_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_MILL_TGT_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_MILL_TGT_WID_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_MILL_TGT_WID_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_UST_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_UST_FL_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_MILL_STD_EDT_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(ul_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HCR_KND_1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HCR_KND_NAME_1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_MILL_PLT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'================================================================================================================================================================
            Call Gp_Ms_Collection(txt_SHOT_BLAST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_SHOT_BLAST_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
             Call Gp_Ms_Collection(txt_HTM_METH1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_MTH_COND1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

             Call Gp_Ms_Collection(txt_HTM_METH2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_MTH_COND2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

             Call Gp_Ms_Collection(txt_HTM_METH3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_MTH_COND3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'================================================================================================================================================================
              Call Gp_Ms_Collection(TXT_SHEAR_C1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_SHEAR_C2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_SHEAR_C3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)


'Insert Emp
'----------------------------------------------------------------------------------------------------------------------------------------------------------------

        Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
    
    'MASTER Collection
     Mc1.Add Item:="AQB0160C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQB0160C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
'        Case "txt_MLT_STD_NO"           '���������淶���
'            sCode = "MLT_STD_NO"
'
'        Case "txt_MILL_STD_NO"          '���������淶���
'            sCode = "MILL_STD_NO"
            
        Case "txt_HCR_KND"              '������ʽ
            sCode = "C0005"
            Set oCodeName = txt_HCR_KND_NAME
        Case "txt_HCR_KND_1"            'HCR����
            sCode = "C0005"
            Set oCodeName = txt_HCR_KND_NAME_1
                
        Case "txt_CR_CD"                '��������
            sCode = "Q0035"
            Set oCodeName = txt_CR_NAME
                
        Case "txt_COOL_WAY"             '��ȴ����
            sCode = "Q0036"
            Set oCodeName = txt_COOL_WAY_NAME
            
        Case "txt_COOL_CTL_TYP"         '������ȴ����
            sCode = "Q0037"
            Set oCodeName = txt_COOL_CTL_NAME
            
        Case "txt_HOT_LVL_USE"          'ʹ���Ƚô���
            sCode = "Q0038"
            
'HYS INSERT START
        Case "txt_CR_CD_Z"                '��������
            sCode = "Q0035"
            Set oCodeName = txt_CR_NAME_Z
                
        Case "txt_COOL_WAY_Z"             '��ȴ����
            sCode = "Q0036"
            Set oCodeName = txt_COOL_WAY_NAME_Z
            
        Case "txt_COOL_CTL_TYP_Z"         '������ȴ����
            sCode = "Q0037"
            Set oCodeName = txt_COOL_CTL_NAME_Z
            
        Case "txt_HOT_LVL_USE_Z"          'ʹ���Ƚô���
            sCode = "Q0038"
'HYS INSERT END
                
        Case "txt_UST_FL"               'UST�c��
            sCode = "Q0046"
            Set oCodeName = txt_UST_FL_NAME
            
        Case "txt_SHOT_BLAST"            '�������
            sCode = "Q0074"
            Set oCodeName = txt_SHOT_BLAST_NAME
        
        Case "txt_HTM_METH1"            '�ȴ�������
            sCode = "Q0073"
        
        Case "txt_HTM_METH2"            '�ȴ�������
            sCode = "Q0073"
        
        Case "txt_HTM_METH3"            '�ȴ�������
            sCode = "Q0073"
    
        Case "txt_MTH_COND1"            '�ȴ�������
            sCode = "HTM_COND_CD"
        
        Case "txt_MTH_COND2"            '�ȴ�������
            sCode = "HTM_COND_CD"
        
        Case "txt_MTH_COND3"            '�ȴ�������
            sCode = "HTM_COND_CD"
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
    Call Form_Define

    'Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    txt_ORD_NO.Text = sOrderNo
    txt_ORD_ITEM.Text = sOrderItem

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    pControl(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

'    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

'    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()
        
    Dim sMesg As String
            
        If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gf_subMasterLock(Mc1, Trim(txt_Design_STS.Text))
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            If Trim(txt_Design_STS.Text) = "A" Or Trim(txt_Design_STS.Text) = "a" Or Trim(txt_Design_STS.Text) = "*" Then
                SSOp_C1.Enabled = False
                SSOp_C2.Enabled = False
                SSOp_C3.Enabled = False
                SSOp_All.Enabled = False
                cob_MLT_PROC_CD_1.Enabled = False
                cob_MLT_PROC_CD_2.Enabled = False
                cob_MLT_PROC_CD_3.Enabled = False
                cob_MLT_PROC_CD2_1.Enabled = False
                cob_MLT_PROC_CD2_2.Enabled = False
                cob_MLT_PROC_CD2_3.Enabled = False
            End If
            
            Call MLT_PROC_LINE_CHECK
            
        End If
    
End Sub

Public Sub Form_Pro()

    '------  Midify  07.04.07  Kim sung ho
    
    Dim sMessg As String
           
    If txt_MLT_PROC_CD.Enabled = False And txt_MLT_PROC_CD.Enabled = False Then Exit Sub

    If Trim(txt_MLT_PROC_LINE.Text) = "1" Or Trim(txt_MLT_PROC_LINE.Text) = "A" Or Trim(txt_MLT_PROC_LINE.Text) = "B" Then
        If Trim(sdb_MLT_TMP_MIN.Text) = "" Or _
        Trim(sdb_MLT_TMP_MAX.Text) = "" Or _
        Trim(sdb_MLT_TMP_TGT.Text) = "" Then
            Call Gp_MsgBoxDisplay("�����¶���Ϣ������������", "I")
            Exit Sub
        End If

        If Trim(sdb_TD_TMP_MIN.Text) = "" Or _
            Trim(sdb_TD_TMP_MAX.Text) = "" Or _
            Trim(sdb_TD_TMP_TGT.Text) = "" Then
            Call Gp_MsgBoxDisplay("�м���¶���Ϣ������������", "I")
            Exit Sub
        End If
    End If
    
    If Trim(txt_MLT_PROC_LINE.Text) = "2" Or Trim(txt_MLT_PROC_LINE.Text) = "A" Or Trim(txt_MLT_PROC_LINE.Text) = "B" Then
        If Trim(sdb_MLT_TMP_MIN2.Text) = "" Or _
           Trim(sdb_MLT_TMP_MAX2.Text) = "" Or _
           Trim(sdb_MLT_TMP_TGT2.Text) = "" Then
            Call Gp_MsgBoxDisplay("�����¶���Ϣ������������", "I")
            Exit Sub
        End If

        If Trim(sdb_TD_TMP_MIN2.Text) = "" Or _
            Trim(sdb_TD_TMP_MAX2.Text) = "" Or _
            Trim(sdb_TD_TMP_TGT2.Text) = "" Then
            Call Gp_MsgBoxDisplay("�м���¶���Ϣ������������", "I")
            Exit Sub
        End If

    End If
    
    sMessg = Gf_Ms_NeceCheck(Mc1.Item("nControl"))
    
    If Trim(sMessg) <> "OK" Then
        Call Gp_MsgBoxDisplay(Trim(sMessg) + "��������", "I")
        Exit Sub
    End If
    
    If Trim(txt_MLT_PROC_LINE.Text) = "1" Then
        sdb_MLT_TMP_MIN2.Text = ""
        sdb_MLT_TMP_MAX2.Text = ""
        sdb_MLT_TMP_TGT2.Text = ""
        
        sdb_LF_TEMP_MIN2.Text = ""
        sdb_LF_TEMP_MAX2.Text = ""
        sdb_LF_TEMP_TGT2.Text = ""
        
        sdb_RH_TMP_MIN2.Text = ""
        sdb_RH_TMP_MAX2.Text = ""
        sdb_RH_TMP_TGT2.Text = ""
        
        sdb_TD_TMP_MIN2.Text = ""
        sdb_TD_TMP_MAX2.Text = ""
        sdb_TD_TMP_TGT2.Text = ""
        
        cob_MLT_PROC_CD2_1.ListIndex = 0
        cob_MLT_PROC_CD2_2.ListIndex = 0
        cob_MLT_PROC_CD2_3.ListIndex = 0
        
    End If
    
    If Trim(txt_MLT_PROC_LINE.Text) = "2" Then
        sdb_MLT_TMP_MIN.Text = ""
        sdb_MLT_TMP_MAX.Text = ""
        sdb_MLT_TMP_TGT.Text = ""
        
        sdb_LF_TEMP_MIN.Text = ""
        sdb_LF_TEMP_MAX.Text = ""
        sdb_LF_TEMP_TGT.Text = ""
        
        sdb_VD_TMP_MIN.Text = ""
        sdb_VD_TMP_MAX.Text = ""
        sdb_VD_TMP_TGT.Text = ""
        
        sdb_TD_TMP_MIN.Text = ""
        sdb_TD_TMP_MAX.Text = ""
        sdb_TD_TMP_TGT.Text = ""
        
        cob_MLT_PROC_CD_1.ListIndex = 0
        cob_MLT_PROC_CD_2.ListIndex = 0
        cob_MLT_PROC_CD_3.ListIndex = 0
    
    End If
    
    If proc_Value_Check = False Then Exit Sub
        If Sp_AllUse_NecessaryCheck() = False Then Exit Sub
        If subMinMaxValueCheck = False Then Exit Sub
        txt_ins_emp.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    'End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub


'����ֵ , ����ֵ Check
Private Function subMinMaxValueCheck() As Boolean
    
    '------ Modify 07.04.07 Kim sung ho
    
    If Trim(sdb_MLT_TMP_MIN.Text) <> "" Or Trim(sdb_MLT_TMP_MAX.Text) <> "" Or Trim(sdb_MLT_TMP_TGT.Text) <> "" Then
        If Gf_subValueCheck(sdb_MLT_TMP_MIN, sdb_MLT_TMP_MAX, sdb_MLT_TMP_TGT) = False Then Exit Function
    End If
    
    If Trim(sdb_LF_TEMP_MIN.Text) <> "" Or Trim(sdb_LF_TEMP_MAX.Text) <> "" Or Trim(sdb_LF_TEMP_TGT.Text) <> "" Then
        If Gf_subValueCheck(sdb_LF_TEMP_MIN, sdb_LF_TEMP_MAX, sdb_LF_TEMP_TGT) = False Then Exit Function
    End If
    
    If Trim(sdb_VD_TMP_MIN.Text) <> "" Or Trim(sdb_VD_TMP_MAX.Text) <> "" Or Trim(sdb_VD_TMP_TGT.Text) <> "" Then
        If Gf_subValueCheck(sdb_VD_TMP_MIN, sdb_VD_TMP_MAX, sdb_VD_TMP_TGT) = False Then Exit Function
    End If
    
    If Trim(sdb_TD_TMP_MIN.Text) <> "" Or Trim(sdb_TD_TMP_MAX.Text) <> "" Or Trim(sdb_TD_TMP_TGT.Text) <> "" Then
        If Gf_subValueCheck(sdb_TD_TMP_MIN, sdb_TD_TMP_MAX, sdb_TD_TMP_TGT) = False Then Exit Function
    End If
    
    If Trim(sdb_MLT_TMP_MIN2.Text) <> "" Or Trim(sdb_MLT_TMP_MAX2.Text) <> "" Or Trim(sdb_MLT_TMP_TGT2.Text) <> "" Then
        If Gf_subValueCheck(sdb_MLT_TMP_MIN2, sdb_MLT_TMP_MAX2, sdb_MLT_TMP_TGT2) = False Then Exit Function
    End If
    
    If Trim(sdb_LF_TEMP_MIN2.Text) <> "" Or Trim(sdb_LF_TEMP_MAX2.Text) <> "" Or Trim(sdb_LF_TEMP_TGT2.Text) <> "" Then
        If Gf_subValueCheck(sdb_LF_TEMP_MIN2, sdb_LF_TEMP_MAX2, sdb_LF_TEMP_TGT2) = False Then Exit Function
    End If
    
    If Trim(sdb_RH_TMP_MIN2.Text) <> "" Or Trim(sdb_RH_TMP_MAX2.Text) <> "" Or Trim(sdb_RH_TMP_TGT2.Text) <> "" Then
        If Gf_subValueCheck(sdb_RH_TMP_MIN2, sdb_RH_TMP_MAX2, sdb_RH_TMP_TGT2) = False Then Exit Function
    End If
    
    If Trim(sdb_TD_TMP_MIN2.Text) <> "" Or Trim(sdb_TD_TMP_MAX2.Text) <> "" Or Trim(sdb_TD_TMP_TGT2.Text) <> "" Then
        If Gf_subValueCheck(sdb_TD_TMP_MIN2, sdb_TD_TMP_MAX2, sdb_TD_TMP_TGT2) = False Then Exit Function
    End If
    
    If Gf_subValueCheck(txt_MILL_TMP_MIN, txt_MILL_TMP_MAX) = False Then Exit Function
    If Gf_subValueCheck(txt_COOL_TMP_MIN, txt_COOL_TMP_MAX) = False Then Exit Function
        
    subMinMaxValueCheck = True

End Function

Private Sub cob_MLT_PROC_CD_1_Click()
    Dim CD As String
    Dim sText As String
    
    sText = txt_MLT_PROC_CD.Text
    
    With cob_MLT_PROC_CD_1
    Select Case .ListIndex
        Case 0
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD)
            
            cob_MLT_PROC_CD_2.Clear
            Call cob_MLT_PROC_CD_2.AddItem("**", 0)
            cob_MLT_PROC_CD_2.ListIndex = 0
        Case 1
            CD = "BG"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD)
            
            cob_MLT_PROC_CD_2.Clear
            Call cob_MLT_PROC_CD_2.AddItem("**", 0)
            Call cob_MLT_PROC_CD_2.AddItem("LF", 1)
            Call cob_MLT_PROC_CD_2.AddItem("VD", 2)
            cob_MLT_PROC_CD_2.ListIndex = 0
        Case 2
            CD = "BD"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD)
            
            cob_MLT_PROC_CD_2.Clear
            Call cob_MLT_PROC_CD_2.AddItem("**", 0)
            Call cob_MLT_PROC_CD_2.AddItem("VD", 1)
            cob_MLT_PROC_CD_2.ListIndex = 0
        Case 3
            CD = "BE"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD)
            
            cob_MLT_PROC_CD_2.Clear
            Call cob_MLT_PROC_CD_2.AddItem("**", 0)
            Call cob_MLT_PROC_CD_2.AddItem("LF", 1)
            cob_MLT_PROC_CD_2.ListIndex = 0
        Case Else
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD)
            
            cob_MLT_PROC_CD_2.Clear
            Call cob_MLT_PROC_CD_2.AddItem("**", 0)
            cob_MLT_PROC_CD_2.ListIndex = 0
    End Select
    End With
End Sub

Private Sub cob_MLT_PROC_CD2_1_Click()
    Dim CD As String
    Dim sText As String
    
    sText = txt_MLT_PROC_CD2.Text
    
    With cob_MLT_PROC_CD2_1
    Select Case .ListIndex
        Case 0
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_2.Clear
            Call cob_MLT_PROC_CD2_2.AddItem("**", 0)
            cob_MLT_PROC_CD2_2.ListIndex = 0
         Case 1
            CD = "BG"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_2.Clear
            Call cob_MLT_PROC_CD2_2.AddItem("**", 0)
            Call cob_MLT_PROC_CD2_2.AddItem("LF", 1)
            Call cob_MLT_PROC_CD2_2.AddItem("RH", 2)
            cob_MLT_PROC_CD2_2.ListIndex = 0
        Case 2
            CD = "BD"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_2.Clear
            Call cob_MLT_PROC_CD2_2.AddItem("**", 0)
            Call cob_MLT_PROC_CD2_2.AddItem("RH", 1)
            cob_MLT_PROC_CD2_2.ListIndex = 0
        Case 3
            CD = "BH"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_2.Clear
            Call cob_MLT_PROC_CD2_2.AddItem("**", 0)
            Call cob_MLT_PROC_CD2_2.AddItem("LF", 1)
            cob_MLT_PROC_CD2_2.ListIndex = 0
        Case Else
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 1, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_2.Clear
            Call cob_MLT_PROC_CD2_2.AddItem("**", 0)
            cob_MLT_PROC_CD2_2.ListIndex = 0
    End Select
    End With

End Sub

Private Sub cob_MLT_PROC_CD_2_Click()
    Dim CD As String
    Dim sText As String
    
    sText = txt_MLT_PROC_CD.Text
    
    With cob_MLT_PROC_CD_2
    Select Case .Text
        Case "**"
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD)
            
            cob_MLT_PROC_CD_3.Clear
            Call cob_MLT_PROC_CD_3.AddItem("**", 0)
            cob_MLT_PROC_CD_3.ListIndex = 0
        Case "LF"
            CD = "BD"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD)
            
            If cob_MLT_PROC_CD_1.Text = "CAS" Then
               cob_MLT_PROC_CD_3.Clear
               Call cob_MLT_PROC_CD_3.AddItem("**", 0)
               Call cob_MLT_PROC_CD_3.AddItem("VD", 1)
               cob_MLT_PROC_CD_3.ListIndex = 0
            Else
                cob_MLT_PROC_CD_3.Clear
                Call cob_MLT_PROC_CD_3.AddItem("**", 0)
                cob_MLT_PROC_CD_3.ListIndex = 0
            End If
        Case "VD"
            CD = "BE"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD)
            
            If cob_MLT_PROC_CD_1.Text = "CAS" Then
                cob_MLT_PROC_CD_3.Clear
                Call cob_MLT_PROC_CD_3.AddItem("**", 0)
                Call cob_MLT_PROC_CD_3.AddItem("LF", 1)
                cob_MLT_PROC_CD_3.ListIndex = 0
            Else
                cob_MLT_PROC_CD_3.Clear
                Call cob_MLT_PROC_CD_3.AddItem("**", 0)
                cob_MLT_PROC_CD_3.ListIndex = 0
            End If
        Case Else
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD)
            
            cob_MLT_PROC_CD_3.Clear
            Call cob_MLT_PROC_CD_3.AddItem("**", 0)
            cob_MLT_PROC_CD_3.ListIndex = 0
    End Select
    End With

End Sub

Private Sub cob_MLT_PROC_CD2_2_Click()
    Dim CD As String
    Dim sText As String
    
    sText = txt_MLT_PROC_CD2.Text
    
    With cob_MLT_PROC_CD2_2
    Select Case .Text
        Case "**"
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_3.Clear
            Call cob_MLT_PROC_CD2_3.AddItem("**", 0)
            cob_MLT_PROC_CD2_3.ListIndex = 0
        Case "LF"
            CD = "BD"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD2)
            
            If cob_MLT_PROC_CD2_1.Text = "CAS" Then
               cob_MLT_PROC_CD2_3.Clear
               Call cob_MLT_PROC_CD2_3.AddItem("**", 0)
               Call cob_MLT_PROC_CD2_3.AddItem("RH", 1)
               cob_MLT_PROC_CD2_3.ListIndex = 0
            Else
                cob_MLT_PROC_CD2_3.Clear
                Call cob_MLT_PROC_CD2_3.AddItem("**", 0)
                cob_MLT_PROC_CD2_3.ListIndex = 0
            End If
        Case "RH"
            CD = "BH"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_3.Clear
            If cob_MLT_PROC_CD2_1.Text = "CAS" Then
               cob_MLT_PROC_CD2_3.Clear
               Call cob_MLT_PROC_CD2_3.AddItem("**", 0)
               Call cob_MLT_PROC_CD2_3.AddItem("LF", 1)
               cob_MLT_PROC_CD2_3.ListIndex = 0
            Else
                cob_MLT_PROC_CD2_3.Clear
                Call cob_MLT_PROC_CD2_3.AddItem("**", 0)
                cob_MLT_PROC_CD2_3.ListIndex = 0
            End If
        Case Else
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 3, txt_MLT_PROC_CD2)
            
            cob_MLT_PROC_CD2_3.Clear
            Call cob_MLT_PROC_CD2_3.AddItem("**", 0)
            cob_MLT_PROC_CD2_3.ListIndex = 0
    End Select
    End With

End Sub

Private Sub cob_MLT_PROC_CD_3_Click()
    Dim CD As String
    Dim sText As String
    
    sText = txt_MLT_PROC_CD.Text
    
    With cob_MLT_PROC_CD_3
    Select Case .Text
        Case "**"
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD)
            
        Case "LF"
            CD = "BD"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD)
            
        Case "VD"
            CD = "BE"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD)
            
        Case Else
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD)
            
    End Select
    End With

End Sub

Private Sub cob_MLT_PROC_CD2_3_Click()
    Dim CD As String
    Dim sText As String
    
    sText = txt_MLT_PROC_CD2.Text
    
    With cob_MLT_PROC_CD2_3
    Select Case .Text
        Case "**"
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD2)
            
        Case "LF"
            CD = "BD"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD2)
            
        Case "RH"
            CD = "BH"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD2)
            
        Case Else
            CD = "**"
            
            Call Change_PROC_CD(sText, CD, 5, txt_MLT_PROC_CD2)
            
    End Select
    End With

End Sub



Private Sub SSOp_All_Click(Value As Integer)
    If SSOp_All.Value = True Then
       txt_MILL_PLT.Text = "**"
    
' C1 COLOR SET : YELLOW
       txt_CHG_TMP_TGT.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_SC.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_TAPE.BackColor = &HC0FFFF
       txt_CR_CD.BackColor = &HC0FFFF
       txt_COOL_WAY.BackColor = &HC0FFFF
       txt_COOL_CTL_TYP.BackColor = &HC0FFFF
       txt_HOT_USE.BackColor = &HC0FFFF
' C3 COLOR SET : YELLOW
       txt_CHG_TMP_TGT_Z.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_SC_Z.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_TAPE_Z.BackColor = &HC0FFFF
       txt_CR_CD_Z.BackColor = &HC0FFFF
       txt_COOL_WAY_Z.BackColor = &HC0FFFF
       txt_COOL_CTL_TYP_Z.BackColor = &HC0FFFF
       txt_HOT_USE_Z.BackColor = &HC0FFFF
    End If

End Sub

Private Sub SSOp_C1_Click(Value As Integer)
    If SSOp_C1.Value = True Then
       txt_MILL_PLT.Text = "C1"
' C1 COLOR SET : YELLOW
       txt_CHG_TMP_TGT.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_SC.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_TAPE.BackColor = &HC0FFFF
       txt_CR_CD.BackColor = &HC0FFFF
       txt_COOL_WAY.BackColor = &HC0FFFF
       txt_COOL_CTL_TYP.BackColor = &HC0FFFF
       txt_HOT_USE.BackColor = &HC0FFFF
          
' C3 COLOR CLEAR
       txt_CHG_TMP_TGT_Z.BackColor = &H80000005
       txt_CHG_TMP_DEF_SC_Z.BackColor = &H80000005
       txt_CHG_TMP_DEF_TAPE_Z.BackColor = &H80000005
       txt_CR_CD_Z.BackColor = &H80000005
       txt_COOL_WAY_Z.BackColor = &H80000005
       txt_COOL_CTL_TYP_Z.BackColor = &H80000005
       txt_HOT_USE_Z.BackColor = &H80000005
    End If
End Sub

Private Sub SSOp_C2_Click(Value As Integer)
    If SSOp_C2.Value = True Then
       txt_MILL_PLT.Text = "C3"
' C3 COLOR SET : YELLOW
       txt_CHG_TMP_TGT_Z.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_SC_Z.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_TAPE_Z.BackColor = &HC0FFFF
       txt_CR_CD_Z.BackColor = &HC0FFFF
       txt_COOL_WAY_Z.BackColor = &HC0FFFF
       txt_COOL_CTL_TYP_Z.BackColor = &HC0FFFF
       txt_HOT_USE_Z.BackColor = &HC0FFFF
         
' C1 COLOR CLEAR : WHITE
       txt_CHG_TMP_TGT.BackColor = &H80000005
       txt_CHG_TMP_DEF_SC.BackColor = &H80000005
       txt_CHG_TMP_DEF_TAPE.BackColor = &H80000005
       txt_CR_CD.BackColor = &H80000005
       txt_COOL_WAY.BackColor = &H80000005
       txt_COOL_CTL_TYP.BackColor = &H80000005
       txt_HOT_USE.BackColor = &H80000005
    End If
End Sub



Private Sub SSOption1_Click(Value As Integer)

End Sub

Private Sub SSOp_C3_Click(Value As Integer)
     If SSOp_C3.Value = True Then
       txt_MILL_PLT.Text = "C2"
' C3 COLOR SET : YELLOW
       txt_CHG_TMP_TGT_Z.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_SC_Z.BackColor = &HC0FFFF
       txt_CHG_TMP_DEF_TAPE_Z.BackColor = &HC0FFFF
       txt_CR_CD_Z.BackColor = &HC0FFFF
       txt_COOL_WAY_Z.BackColor = &HC0FFFF
       txt_COOL_CTL_TYP_Z.BackColor = &HC0FFFF
       txt_HOT_USE_Z.BackColor = &HC0FFFF
         
' C1 COLOR CLEAR : WHITE
       txt_CHG_TMP_TGT.BackColor = &H80000005
       txt_CHG_TMP_DEF_SC.BackColor = &H80000005
       txt_CHG_TMP_DEF_TAPE.BackColor = &H80000005
       txt_CR_CD.BackColor = &H80000005
       txt_COOL_WAY.BackColor = &H80000005
       txt_COOL_CTL_TYP.BackColor = &H80000005
       txt_HOT_USE.BackColor = &H80000005
    End If
End Sub

Private Sub txt_MILL_PLT_Change()
    Select Case Trim(txt_MILL_PLT.Text)
        Case "C1"
            SSOp_C1.Value = True
        Case "C3"
            SSOp_C2.Value = True
         Case "C2"
            SSOp_C3.Value = True
        Case "**"
            SSOp_All.Value = True
    End Select
End Sub

Private Sub txt_MLT_PROC_CD_Change()
    With txt_MLT_PROC_CD
        If Len(txt_MLT_PROC_CD) = 0 Then
            txt_MLT_PROC_CD = "******"
        End If
        If Mid(.Text, 1, 2) = "**" Or Mid(.Text, 1, 2) = "BD" Or Mid(.Text, 1, 2) = "BE" Or Mid(.Text, 1, 2) = "BG" Then
            Select Case Mid(.Text, 1, 2)
                Case "**"
                    cob_MLT_PROC_CD_1.Text = "**"
                Case "BG"
                    cob_MLT_PROC_CD_1.Text = "CAS"
                Case "BD"
                    cob_MLT_PROC_CD_1.Text = "LF"
                Case "BE"
                    cob_MLT_PROC_CD_1.Text = "VD"
                Case Else
                    cob_MLT_PROC_CD_1.Text = "**"
            End Select
        End If
        If Mid(.Text, 3, 2) = "**" Or Mid(.Text, 3, 2) = "BE" Or Mid(.Text, 3, 2) = "BD" Then
            Select Case Mid(.Text, 3, 2)
                Case "**"
                    cob_MLT_PROC_CD_2.Text = "**"
                Case "BD"
                    cob_MLT_PROC_CD_2.Text = "LF"
                Case "BE"
                    cob_MLT_PROC_CD_2.Text = "VD"
                Case Else
                    cob_MLT_PROC_CD_2.Text = "**"
            End Select
        End If
        If Mid(.Text, 5, 2) = "**" Or Mid(.Text, 5, 2) = "BD" Or Mid(.Text, 5, 2) = "BE" Then
            Select Case Mid(.Text, 5, 2)
                Case "**"
                    cob_MLT_PROC_CD_3.Text = "**"
                Case "BD"
                    cob_MLT_PROC_CD_3.Text = "LF"
                Case "BE"
                    cob_MLT_PROC_CD_3.Text = "VD"
                Case Else
                    cob_MLT_PROC_CD_3.Text = "**"
            End Select
        End If
    End With

End Sub

Private Sub txt_MLT_PROC_CD2_Change()
    With txt_MLT_PROC_CD2
        If Len(txt_MLT_PROC_CD2) = 0 Then
            txt_MLT_PROC_CD2 = "******"
        End If
        
        
        If Mid(.Text, 1, 2) = "**" Or Mid(.Text, 1, 2) = "BD" Or Mid(.Text, 1, 2) = "BH" Or Mid(.Text, 1, 2) = "BG" Then
            Select Case Mid(.Text, 1, 2)
                Case "**"
                    cob_MLT_PROC_CD2_1.Text = "**"

                Case "BG"
                    cob_MLT_PROC_CD2_1.Text = "CAS"

                Case "BD"
                    cob_MLT_PROC_CD2_1.Text = "LF"

                Case "BH"
                    cob_MLT_PROC_CD2_1.Text = "RH"

                Case Else
                    cob_MLT_PROC_CD2_1.Text = "**"
            End Select
        End If
        If Mid(.Text, 3, 2) = "**" Or Mid(.Text, 3, 2) = "BH" Or Mid(.Text, 3, 2) = "BD" Then
            Select Case Mid(.Text, 3, 2)
                Case "**"
                    cob_MLT_PROC_CD2_2.Text = "**"

                Case "BD"
                    cob_MLT_PROC_CD2_2.Text = "LF"

                Case "BH"
                    cob_MLT_PROC_CD2_2.Text = "RH"

                Case Else
                    cob_MLT_PROC_CD2_2.Text = "**"
            End Select
        End If
        If Mid(.Text, 5, 2) = "**" Or Mid(.Text, 5, 2) = "BD" Or Mid(.Text, 5, 2) = "BH" Then
            Select Case Mid(.Text, 5, 2)
                Case "**"
                    cob_MLT_PROC_CD2_3.Text = "**"
                Case "BD"
                    cob_MLT_PROC_CD2_3.Text = "LF"
                Case "BH"
                    cob_MLT_PROC_CD2_3.Text = "RH"
                Case Else
                    cob_MLT_PROC_CD2_3.Text = "**"
            End Select
        End If
            
    End With

End Sub

Private Function proc_Value_Check() As Boolean

    '--------  Modify 07.04.07 Kim sung ho
    
    proc_Value_Check = False
    
    '��������·�����Ƿ����LF,�����������
    If InStr(1, Trim(txt_MLT_PROC_CD.Text), "BD") > 0 Then
        If Trim(sdb_LF_TEMP_MIN.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� LF ����¶������Ƿ�����!", "I")
           sdb_LF_TEMP_MIN.SetFocus
           
           Exit Function
        End If
        If Trim(sdb_LF_TEMP_MAX.Text) = "" Then
         Call Gp_MsgBoxDisplay("����У�����-���� LF ����¶������Ƿ�����!", "I")
           sdb_LF_TEMP_MAX.SetFocus
           Exit Function
        End If
        If Trim(sdb_LF_TEMP_TGT.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� LF Ŀ���¶������Ƿ�����!", "I")
           sdb_LF_TEMP_TGT.SetFocus
           Exit Function
        End If
    End If
        
    If InStr(1, Trim(txt_MLT_PROC_CD2.Text), "BD") > 0 Then
        If Trim(sdb_LF_TEMP_MIN2.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� LF ����¶������Ƿ�����!", "I")
           sdb_LF_TEMP_MIN2.SetFocus
           Exit Function
        End If
        If Trim(sdb_LF_TEMP_MAX2.Text) = "" Then
         Call Gp_MsgBoxDisplay("����У�����-���� LF ����¶������Ƿ�����!", "I")
           sdb_LF_TEMP_MAX2.SetFocus
           Exit Function
        End If
        If Trim(sdb_LF_TEMP_TGT2.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� LF Ŀ���¶������Ƿ�����!", "I")
           sdb_LF_TEMP_TGT2.SetFocus
           Exit Function
        End If

    End If
    
    '��������·�����Ƿ����VD,�����������
    If InStr(1, Trim(txt_MLT_PROC_CD.Text), "BE") > 0 Then
        If Trim(sdb_VD_TMP_MIN.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� VD ����¶������Ƿ�����!", "I")
           sdb_VD_TMP_MIN.SetFocus
           Exit Function
        End If
        If Trim(sdb_VD_TMP_MAX.Text) = "" Then
         Call Gp_MsgBoxDisplay("����У�����-���� VD ����¶������Ƿ�����!", "I")
           sdb_VD_TMP_MAX.SetFocus
           Exit Function
        End If
        If Trim(sdb_VD_TMP_TGT.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� VD Ŀ���¶������Ƿ�����!", "I")
           sdb_VD_TMP_TGT.SetFocus
           Exit Function
        End If

    End If
    
    '��������·�����Ƿ����RH,�����������
    If InStr(1, Trim(txt_MLT_PROC_CD2.Text), "BH") > 0 Then
        If Trim(sdb_RH_TMP_MIN2.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� RH ����¶������Ƿ�����!", "I")
           sdb_RH_TMP_MIN2.SetFocus
           Exit Function
        End If
        If Trim(sdb_RH_TMP_MAX2.Text) = "" Then
         Call Gp_MsgBoxDisplay("����У�����-���� RH ����¶������Ƿ�����!", "I")
           sdb_RH_TMP_MAX2.SetFocus
           Exit Function
        End If
        If Trim(sdb_RH_TMP_TGT2.Text) = "" Then
           Call Gp_MsgBoxDisplay("����У�����-���� RH Ŀ���¶������Ƿ�����!", "I")
           sdb_RH_TMP_TGT2.SetFocus
           Exit Function
        End If
    End If
    
    proc_Value_Check = True
    
End Function


Private Sub txt_MLT_PROC_LINE_KeyUp(KeyCode As Integer, Shift As Integer)

'Modify 07.04.05 Kim sung ho
'------------
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0070"
        DD.rControl.Add Item:=txt_MLT_PROC_LINE
        DD.rControl.Add Item:=txt_MLT_PROC_NAME

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(txt_MLT_PROC_LINE)) = txt_MLT_PROC_LINE.MaxLength Then
        txt_MLT_PROC_NAME.Text = Gf_ComnNameFind(M_CN1, "Q0070", Trim(txt_MLT_PROC_LINE.Text), 2)
    Else
        txt_MLT_PROC_NAME.Text = ""
    End If
    
    Call MLT_PROC_LINE_CHECK
    
End Sub

Private Sub MLT_PROC_LINE_CHECK()

    If txt_MLT_PROC_LINE = "1" Then
' PROC#1 COLOR SET : YELLOW
       sdb_MLT_TMP_MIN.BackColor = &HC0FFFF
       sdb_MLT_TMP_MAX.BackColor = &HC0FFFF
       sdb_MLT_TMP_TGT.BackColor = &HC0FFFF
       sdb_TD_TMP_MIN.BackColor = &HC0FFFF
       sdb_TD_TMP_MAX.BackColor = &HC0FFFF
       sdb_TD_TMP_TGT.BackColor = &HC0FFFF
' PROC#2 COLOR SET : WHITE
       sdb_MLT_TMP_MIN2.BackColor = &H80000005
       sdb_MLT_TMP_MAX2.BackColor = &H80000005
       sdb_MLT_TMP_TGT2.BackColor = &H80000005
       sdb_TD_TMP_MIN2.BackColor = &H80000005
       sdb_TD_TMP_MAX2.BackColor = &H80000005
       sdb_TD_TMP_TGT2.BackColor = &H80000005
       
    ElseIf txt_MLT_PROC_LINE = "2" Then
' PROC#2 COLOR SET : YELLOW
       sdb_MLT_TMP_MIN2.BackColor = &HC0FFFF
       sdb_MLT_TMP_MAX2.BackColor = &HC0FFFF
       sdb_MLT_TMP_TGT2.BackColor = &HC0FFFF
       sdb_TD_TMP_MIN2.BackColor = &HC0FFFF
       sdb_TD_TMP_MAX2.BackColor = &HC0FFFF
       sdb_TD_TMP_TGT2.BackColor = &HC0FFFF
       
' PROC#1 COLOR SET : WHITE
       sdb_MLT_TMP_MIN.BackColor = &H80000005
       sdb_MLT_TMP_MAX.BackColor = &H80000005
       sdb_MLT_TMP_TGT.BackColor = &H80000005
       sdb_TD_TMP_MIN.BackColor = &H80000005
       sdb_TD_TMP_MAX.BackColor = &H80000005
       sdb_TD_TMP_TGT.BackColor = &H80000005
       
    ElseIf txt_MLT_PROC_LINE = "A" Or txt_MLT_PROC_LINE = "B" Then
    
' PROC#1 COLOR SET : YELLOW : MANDANTORY INPUT
       sdb_MLT_TMP_MIN.BackColor = &HC0FFFF
       sdb_MLT_TMP_MAX.BackColor = &HC0FFFF
       sdb_MLT_TMP_TGT.BackColor = &HC0FFFF
       sdb_TD_TMP_MIN.BackColor = &HC0FFFF
       sdb_TD_TMP_MAX.BackColor = &HC0FFFF
       
' PROC#2 COLOR SET : YELLOW : MANDANTORY INPUT
       sdb_MLT_TMP_MIN2.BackColor = &HC0FFFF
       sdb_MLT_TMP_MAX2.BackColor = &HC0FFFF
       sdb_MLT_TMP_TGT2.BackColor = &HC0FFFF
       sdb_TD_TMP_MIN2.BackColor = &HC0FFFF
       sdb_TD_TMP_MAX2.BackColor = &HC0FFFF
       sdb_TD_TMP_TGT2.BackColor = &HC0FFFF
       
    Else
    
' PROC#1 COLOR SET : WHITE
       sdb_MLT_TMP_MIN.BackColor = &H80000005
       sdb_MLT_TMP_MAX.BackColor = &H80000005
       sdb_MLT_TMP_TGT.BackColor = &H80000005
       sdb_TD_TMP_MIN.BackColor = &H80000005
       sdb_TD_TMP_MAX.BackColor = &H80000005
       sdb_TD_TMP_TGT.BackColor = &H80000005
       
' PROC#2 COLOR SET : WHITE
       sdb_MLT_TMP_MIN2.BackColor = &H80000005
       sdb_MLT_TMP_MAX2.BackColor = &H80000005
       sdb_MLT_TMP_TGT2.BackColor = &H80000005
       sdb_TD_TMP_MIN2.BackColor = &H80000005
       sdb_TD_TMP_MAX2.BackColor = &H80000005
       sdb_TD_TMP_TGT2.BackColor = &H80000005
    End If

    
End Sub

Private Sub Change_PROC_CD(ByVal sText As String, ByVal sCD As String, ByVal iLOC As Integer, ByVal oTEXT_OBJ As Object)
               
        If Len(sText) >= 6 Then
            Mid(sText, iLOC, 2) = sCD
        Else
            If iLOC = 1 Then
                sText = sCD
            Else
                sText = Mid(sText, 1, iLOC + 1) + sCD
            End If
        End If
        
        If TypeName(oTEXT_OBJ) = "TextBox" Then
            oTEXT_OBJ.Text = sText
        End If
End Sub

Private Function Change_Mth_COND(ByVal iNO As Integer) As Boolean
    Dim sQuery      As String
    Dim AdoRs       As adodb.Recordset
    Dim sCOND_CODE  As String
        
        Select Case iNO
            Case 1
                sCOND_CODE = Trim(txt_MTH_COND1.Text)
            Case 2
                sCOND_CODE = Trim(txt_MTH_COND2.Text)
            Case 3
                sCOND_CODE = Trim(txt_MTH_COND3.Text)
            Case Else
                AdoRs.Close
                Set AdoRs = Nothing
                Change_Mth_COND = False
                Exit Function
        End Select
    
        sQuery = "Select      NVL(HTM_TEMP_MIN,0),NVL(HTM_TEMP_MAX,0),NVL(HTM_TEMP_TGT,0),"
        sQuery = sQuery + "   NVL(HTM_TIME_1F_MIN,0) ,NVL(HTM_TIME_1F_MAX,0) ,NVL(HTM_TIME_1F_AIM,0) ,"
        sQuery = sQuery + "   NVL(HTM_TIME_2F_MIN,0) ,NVL(HTM_TIME_2F_MAX,0) ,NVL(HTM_TIME_2F_AIM,0) ,"
        sQuery = sQuery + "   NVL(COOL_TIME_MIN,0),NVL(COOL_TIME_MAX,0),NVL(COOL_TIME_AIM,0),"
        sQuery = sQuery + "   NVL(DRW_TEMP_MIN,0),NVL(DRW_TEMP_MAX,0),NVL(DRW_TEMP_AIM,0),"
        sQuery = sQuery + "   NVL(HTM_COOL_TYP,''),NVL(HTM_COOL_TMP,0) "
        sQuery = sQuery + " From         QP_HEAT_COND"
        sQuery = sQuery + " Where            HTM_COND = " + "'" + sCOND_CODE + "'"

        Set AdoRs = New adodb.Recordset
    
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.EOF Or AdoRs.MaxRecords > 1 Then
            AdoRs.Close
            Set AdoRs = Nothing
            Change_Mth_COND = False
            Exit Function
        End If
        
        Select Case iNO
            Case 1
                   txt_HTM_TEMP_MIN1.Text = AdoRs.Fields(0).Value
                   txt_HTM_TEMP_MAX1.Text = AdoRs.Fields(1).Value
                   txt_HTM_TEMP_TGT1.Text = AdoRs.Fields(2).Value
                txt_HTM_TIME_1F_MIN1.Text = AdoRs.Fields(3).Value
                txt_HTM_TIME_1F_MAX1.Text = AdoRs.Fields(4).Value
                txt_HTM_TIME_1F_AIM1.Text = AdoRs.Fields(5).Value
                txt_HTM_TIME_2F_MIN1.Text = AdoRs.Fields(6).Value
                txt_HTM_TIME_2F_MAX1.Text = AdoRs.Fields(7).Value
                txt_HTM_TIME_2F_AIM1.Text = AdoRs.Fields(8).Value
                  txt_COOL_TIME_MIN1.Text = AdoRs.Fields(9).Value
                  txt_COOL_TIME_MAX1.Text = AdoRs.Fields(10).Value
                  txt_COOL_TIME_AIM1.Text = AdoRs.Fields(11).Value
                   txt_DRW_TEMP_MIN1.Text = AdoRs.Fields(12).Value
                   txt_DRW_TEMP_MAX1.Text = AdoRs.Fields(13).Value
                   txt_DRW_TEMP_AIM1.Text = AdoRs.Fields(14).Value
                   txt_HTM_COOL_TYP1.Text = AdoRs.Fields(15).Value
                   txt_HTM_COOL_TMP1.Text = AdoRs.Fields(16).Value
            Case 2
                   txt_HTM_TEMP_MIN2.Text = AdoRs.Fields(0).Value
                   txt_HTM_TEMP_MAX2.Text = AdoRs.Fields(1).Value
                   txt_HTM_TEMP_TGT2.Text = AdoRs.Fields(2).Value
                txt_HTM_TIME_1F_MIN2.Text = AdoRs.Fields(3).Value
                txt_HTM_TIME_1F_MAX2.Text = AdoRs.Fields(4).Value
                txt_HTM_TIME_1F_AIM2.Text = AdoRs.Fields(5).Value
                txt_HTM_TIME_2F_MIN2.Text = AdoRs.Fields(6).Value
                txt_HTM_TIME_2F_MAX2.Text = AdoRs.Fields(7).Value
                txt_HTM_TIME_2F_AIM2.Text = AdoRs.Fields(8).Value
                  txt_COOL_TIME_MIN2.Text = AdoRs.Fields(9).Value
                  txt_COOL_TIME_MAX2.Text = AdoRs.Fields(10).Value
                  txt_COOL_TIME_AIM2.Text = AdoRs.Fields(11).Value
                   txt_DRW_TEMP_MIN2.Text = AdoRs.Fields(12).Value
                   txt_DRW_TEMP_MAX2.Text = AdoRs.Fields(13).Value
                   txt_DRW_TEMP_AIM2.Text = AdoRs.Fields(14).Value
                   txt_HTM_COOL_TYP2.Text = AdoRs.Fields(15).Value
                   txt_HTM_COOL_TMP2.Text = AdoRs.Fields(16).Value
            Case 3
                   txt_HTM_TEMP_MIN3.Text = AdoRs.Fields(0).Value
                   txt_HTM_TEMP_MAX3.Text = AdoRs.Fields(1).Value
                   txt_HTM_TEMP_TGT3.Text = AdoRs.Fields(2).Value
                txt_HTM_TIME_1F_MIN3.Text = AdoRs.Fields(3).Value
                txt_HTM_TIME_1F_MAX3.Text = AdoRs.Fields(4).Value
                txt_HTM_TIME_1F_AIM3.Text = AdoRs.Fields(5).Value
                txt_HTM_TIME_2F_MIN3.Text = AdoRs.Fields(6).Value
                txt_HTM_TIME_2F_MAX3.Text = AdoRs.Fields(7).Value
                txt_HTM_TIME_2F_AIM3.Text = AdoRs.Fields(8).Value
                  txt_COOL_TIME_MIN3.Text = AdoRs.Fields(9).Value
                  txt_COOL_TIME_MAX3.Text = AdoRs.Fields(10).Value
                  txt_COOL_TIME_AIM3.Text = AdoRs.Fields(11).Value
                   txt_DRW_TEMP_MIN3.Text = AdoRs.Fields(12).Value
                   txt_DRW_TEMP_MAX3.Text = AdoRs.Fields(13).Value
                   txt_DRW_TEMP_AIM3.Text = AdoRs.Fields(14).Value
                   txt_HTM_COOL_TYP3.Text = AdoRs.Fields(15).Value
                   txt_HTM_COOL_TMP3.Text = AdoRs.Fields(16).Value
            Case Else
                AdoRs.Close
                Set AdoRs = Nothing
                Change_Mth_COND = False
                Exit Function
        End Select
    
            AdoRs.Close
            Set AdoRs = Nothing
            Change_Mth_COND = True

End Function

Private Sub txt_MTH_COND1_Change()
    If Len(Trim(txt_MTH_COND1.Text)) < 4 Then Exit Sub
    
    If Change_Mth_COND(1) = False Then
        Call MsgBox("��Ӧ�ȴ��������������!", vbOKOnly, "ϵͳ��ʾ")
    End If
End Sub

Private Sub txt_MTH_COND2_Change()
    If Len(Trim(txt_MTH_COND1.Text)) < 4 Then Exit Sub
    
    If Change_Mth_COND(2) = False Then
        Call MsgBox("��Ӧ�ȴ��������������!", vbOKOnly, "ϵͳ��ʾ")
    End If

End Sub

Private Sub txt_MTH_COND3_Change()
    If Len(Trim(txt_MTH_COND1.Text)) < 4 Then Exit Sub
    
    If Change_Mth_COND(3) = False Then
        Call MsgBox("��Ӧ�ȴ��������������!", vbOKOnly, "ϵͳ��ʾ")
    End If

End Sub

Private Function Sp_AllUse_NecessaryCheck() As Boolean
Dim sPLT_CD As String
            sPLT_CD = txt_MILL_PLT.Text

'------------------------------------------------------ ��ͬ��Ŀ ---------------------------------------------------------

'����ȥ��
    If GF_Necessary_Value_Check(txt_HCR_KND_1, ULabel1(4).Caption, txt_HCR_KND_1) = False Then Exit Function
'��������·��
    If GF_Necessary_Value_Check(txt_MILL_PLT, "��ѡ�����ֳ���") = False Then Exit Function
    Select Case sPLT_CD
           Case "C1"
            If Sp_C1_Item_NecessaryCheck() = False Then Exit Function
           Case "C3"
            If Sp_C2_Item_NecessaryCheck() = False Then Exit Function
           Case "**"
            If Sp_C1_Item_NecessaryCheck() = False Then Exit Function
            If Sp_C2_Item_NecessaryCheck() = False Then Exit Function
    End Select
    
    Sp_AllUse_NecessaryCheck = True
End Function

Private Function Sp_C1_Item_NecessaryCheck() As Boolean
'ƽ����¯�²�
    If GF_Necessary_Value_Check(txt_CHG_TMP_TGT, "ƽ����¯�²�", txt_CHG_TMP_TGT) = False Then Exit Function
    
'��������/�����²�
    If GF_Necessary_Value_Check(txt_CHG_TMP_DEF_SC, "��������/�����²�", txt_CHG_TMP_DEF_SC) = False Then Exit Function
    
'����ͷβ�²�
    If GF_Necessary_Value_Check(txt_CHG_TMP_DEF_TAPE, "����ͷβ�²�", txt_CHG_TMP_DEF_TAPE) = False Then Exit Function
    
'��������
    If GF_Necessary_Value_Check(txt_CR_CD, "��������", txt_CR_CD) = False Then Exit Function
    
'T1�¶�/ѹ����&�����¶�
        If txt_CR_CD.Text = "Y" Or txt_CR_CD.Text = "y" Then
            If GF_Necessary_Value_Check(txt_MILL_TMPT1, "T1�¶�", txt_MILL_TMPT1) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_RATET1, "T1�¶�/ѹ����", txt_MILL_RATET1) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_TMP_MIN, "�����¶��������", txt_MILL_TMP_MIN) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_TMP_MAX, "�����¶��������", txt_MILL_TMP_MAX) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_TMP_TGT, "�����¶�Ŀ��ֵ", txt_MILL_TMP_TGT) = False Then Exit Function
        End If


'��ȴ��������
    If GF_Necessary_Value_Check(txt_COOL_WAY, "��ȴ��������", txt_COOL_WAY) = False Then Exit Function
    
'��ȴ�¶�/��ȴ����

        If txt_COOL_WAY.Text = "W" Or txt_COOL_WAY.Text = "w" Then
            If GF_Necessary_Value_Check(txt_COOL_TMP_MIN, "��ȴ�¶��������", txt_COOL_TMP_MIN) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_COOL_TMP_MAX, "��ȴ�¶��������", txt_COOL_TMP_MAX) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_COOL_TMP_TGT, "��ȴ�¶�Ŀ��ֵ", txt_COOL_TMP_TGT) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_COOL_TMP_RATE, "��ȴ����", txt_COOL_TMP_RATE) = False Then Exit Function
        End If
    
'�������
    If GF_Necessary_Value_Check(txt_COOL_CTL_TYP, "�������", txt_COOL_CTL_TYP) = False Then Exit Function
'�Ƚ�ֱ����
    If GF_Necessary_Value_Check(txt_HOT_USE, "�Ƚ�ֱ����", txt_HOT_USE) = False Then Exit Function
    
    Sp_C1_Item_NecessaryCheck = True
    
End Function

Private Function Sp_C2_Item_NecessaryCheck() As Boolean
'ƽ����¯�²�
    If GF_Necessary_Value_Check(txt_CHG_TMP_TGT_Z, "ƽ����¯�²�", txt_CHG_TMP_TGT_Z) = False Then Exit Function
    
'��������/�����²�
    If GF_Necessary_Value_Check(txt_CHG_TMP_DEF_SC_Z, "��������/�����²�", txt_CHG_TMP_DEF_SC_Z) = False Then Exit Function
    
'����ͷβ�²�
    If GF_Necessary_Value_Check(txt_CHG_TMP_DEF_TAPE_Z, "����ͷβ�²�", txt_CHG_TMP_DEF_TAPE_Z) = False Then Exit Function
    
'��������
    If GF_Necessary_Value_Check(txt_CR_CD_Z, "��������", txt_CR_CD_Z) = False Then Exit Function
    
'T1�¶�/ѹ����&�����¶�
 
        If txt_CR_CD_Z.Text = "Y" Or txt_CR_CD_Z.Text = "y" Then
            If GF_Necessary_Value_Check(txt_MILL_TMPT1_Z, "T1�¶�", txt_MILL_TMPT1_Z) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_RATET1_Z, "T1�¶�/ѹ����", txt_MILL_RATET1_Z) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_TMP_MIN_Z, "�����¶��������", txt_MILL_TMP_MIN_Z) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_TMP_MAX_Z, "�����¶��������", txt_MILL_TMP_MAX_Z) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_MILL_TMP_TGT_Z, "�����¶�Ŀ��ֵ", txt_MILL_TMP_TGT_Z) = False Then Exit Function
        End If

'��ȴ��������
    If GF_Necessary_Value_Check(txt_COOL_WAY_Z, "��ȴ��������", txt_COOL_WAY_Z) = False Then Exit Function
    
'��ȴ�¶�/��ȴ����

        If txt_COOL_WAY_Z.Text = "W" Or txt_COOL_WAY_Z.Text = "w" Then
            If GF_Necessary_Value_Check(txt_COOL_TMP_MIN_Z, "��ȴ�¶��������", txt_COOL_TMP_MIN_Z) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_COOL_TMP_MAX_Z, "��ȴ�¶��������", txt_COOL_TMP_MAX_Z) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_COOL_TMP_TGT_Z, "��ȴ�¶�Ŀ��ֵ", txt_COOL_TMP_TGT_Z) = False Then Exit Function
            If GF_Necessary_Value_Check(txt_COOL_TMP_RATE_Z, "��ȴ����", txt_COOL_TMP_RATE_Z) = False Then Exit Function
        End If

'�������
    If GF_Necessary_Value_Check(txt_COOL_CTL_TYP_Z, "�������", txt_COOL_CTL_TYP_Z) = False Then Exit Function
'�Ƚ�ֱ����
    If GF_Necessary_Value_Check(txt_HOT_USE_Z, "�Ƚ�ֱ����", txt_HOT_USE_Z) = False Then Exit Function
    
    Sp_C2_Item_NecessaryCheck = True
    
End Function

