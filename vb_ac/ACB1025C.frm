VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB1025C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "����ת����ҵָʾ¼��_ACB1025C"
   ClientHeight    =   9225
   ClientLeft      =   345
   ClientTop       =   1575
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_heat_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      MaxLength       =   8
      TabIndex        =   29
      Tag             =   "¯��"
      Top             =   870
      Width           =   975
   End
   Begin VB.TextBox Text_size_knd 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   10590
      MaxLength       =   2
      TabIndex        =   28
      Tag             =   "����"
      Top             =   870
      Width           =   465
   End
   Begin VB.TextBox Text_size_knd_name 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   11070
      TabIndex        =   27
      Tag             =   "����"
      Top             =   870
      Width           =   1275
   End
   Begin VB.TextBox text_cur_inv 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   23
      Top             =   870
      Width           =   1050
   End
   Begin VB.TextBox text_cur_inv_code 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   2
      TabIndex        =   22
      Tag             =   "�ֿ�"
      Top             =   870
      Width           =   375
   End
   Begin VB.TextBox txt_emp 
      Height          =   255
      Left            =   13605
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Text_STLGRD 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      MaxLength       =   11
      TabIndex        =   18
      Top             =   120
      Width           =   1485
   End
   Begin VB.TextBox txt_PLT 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10590
      MaxLength       =   2
      TabIndex        =   16
      Tag             =   "�� ��"
      Top             =   120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_PLT_NAME 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11070
      TabIndex        =   15
      Tag             =   "�� ��"
      Top             =   120
      Visible         =   0   'False
      Width           =   1710
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   7935
      Left            =   60
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1260
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   13996
      _StockProps     =   64
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   18
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACB1025C.frx":0000
   End
   Begin VB.TextBox Text_ORD_ITEM 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "09"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   310
      Left            =   11940
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2655
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   310
      Left            =   12345
      Max             =   1
      Min             =   99
      TabIndex        =   4
      Top             =   2655
      Value           =   1
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox Text_PROC_CD_Name 
      Height          =   270
      Left            =   13065
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text_PROD_CD_Name 
      Height          =   270
      Left            =   12705
      TabIndex        =   2
      Top             =   2820
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox Text_PROD_CD 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "��Ʒ"
      Text            =   "SL"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text_REC_STS_Name 
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
      Left            =   12375
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "CD_MANA_NO"
      Top             =   2265
      Visible         =   0   'False
      Width           =   645
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   165
      Top             =   495
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "��������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   165
      Top             =   120
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "��Ʒ"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin CSTextLibCtl.sidbEdit sdb_thk_fr 
      Height          =   315
      Left            =   5955
      TabIndex        =   6
      Top             =   120
      Width           =   1305
      _Version        =   262145
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      RawData         =   "0.00"
      Text            =   " 0.00"
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
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_thk_to 
      Height          =   315
      Left            =   7500
      TabIndex        =   7
      Top             =   120
      Width           =   1335
      _Version        =   262145
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      RawData         =   "0.00"
      Text            =   " 0.00"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.UDate dtp_ins_date_PROD_DATE1 
      Height          =   315
      Left            =   1170
      TabIndex        =   8
      Tag             =   "INS_DATE"
      Top             =   495
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.UDate dtp_ins_date_PROD_DATE2 
      Height          =   315
      Left            =   2805
      TabIndex        =   9
      Tag             =   "INS_DATE"
      Top             =   495
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   4890
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "���"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   4890
      Top             =   495
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
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_wid_fr 
      Height          =   315
      Left            =   5955
      TabIndex        =   10
      Top             =   495
      Width           =   1305
      _Version        =   262145
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Text            =   " 0"
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
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_wid_to 
      Height          =   315
      Left            =   7500
      TabIndex        =   11
      Top             =   495
      Width           =   1335
      _Version        =   262145
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Text            =   " 0"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   9405
      Tag             =   "�� �� �� ��"
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "Ŀ���"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   1755
      Top             =   120
      Width           =   990
      _ExtentX        =   1746
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   9405
      Top             =   495
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "��ѡ������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_num 
      Height          =   315
      Left            =   10590
      TabIndex        =   20
      Top             =   495
      Width           =   960
      _Version        =   262145
      _ExtentX        =   1693
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Text            =   " 0"
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
      NumIntDigits    =   7
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt 
      Height          =   315
      Left            =   12795
      TabIndex        =   21
      Top             =   495
      Width           =   1260
      _Version        =   262145
      _ExtentX        =   2222
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   11610
      Top             =   495
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "��ѡ��������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   12720
      Top             =   870
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "���ͻ���"
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
   Begin VB.TextBox txt_PRC_line 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13920
      MaxLength       =   1
      TabIndex        =   17
      Tag             =   "����"
      Text            =   "1"
      Top             =   870
      Visible         =   0   'False
      Width           =   465
   End
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   4890
      Top             =   870
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
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_len_fr 
      Height          =   315
      Left            =   5955
      TabIndex        =   24
      Top             =   870
      Width           =   1305
      _Version        =   262145
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Text            =   " 0"
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
      NumIntDigits    =   7
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_len_to 
      Height          =   315
      Left            =   7500
      TabIndex        =   25
      Top             =   870
      Width           =   1335
      _Version        =   262145
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Text            =   " 0"
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   165
      Top             =   870
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "�ѷŲֿ�"
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   9405
      Top             =   870
      Width           =   1170
      _ExtentX        =   2064
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   2805
      Top             =   870
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   556
      Caption         =   "¯��"
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
      ForeColor       =   -2147483641
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   225
      Left            =   7335
      TabIndex        =   26
      Top             =   1020
      Width           =   150
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2760
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   315
      Left            =   7335
      TabIndex        =   13
      Top             =   495
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   315
      Left            =   7335
      TabIndex        =   12
      Top             =   120
      Width           =   150
   End
End
Attribute VB_Name = "ACB1025C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACB1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public STR1 As String
Public BASE As String
Public AIMNO As String
Dim sQuery As String

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

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection


'Dim pColumn2 As New Collection      'Spread Primary Key Collection
'Dim nColumn2 As New Collection      'Spread necessary Column Collection
'Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
'Dim iColumn2 As New Collection      'Spread Insert Column Collection
'Dim aColumn2 As New Collection      'Master -> Spread Column Collection
'Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer
Dim iCol As Integer
Dim iRow As Integer


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
             Call Gp_Ms_Collection(text_prod_cd, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(text_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(dtp_ins_date_prod_date1, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(dtp_ins_date_prod_date2, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          'Call Gp_Ms_Collection(txt_PRC_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                             
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
                
  
    'Spread_Collection
     Call Gp_Sp_Collection1(ss2, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 17, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss2, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    sc1.Add Item:=ss2, Key:="Spread"
    sc1.Add Item:="ACB1025C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss2.MaxCols, Key:="Last"

    
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
       
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Call Gp_Sp_ColHidden(ss2, 15, True)
    Call Gp_Sp_ColHidden(ss2, 16, True)
    Call Gp_Sp_ColHidden(ss2, 17, True)
    Call Gp_Sp_ColHidden(ss2, 18, True)
    
    text_cur_inv_code.Text = "00"
End Sub


Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
      '  Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
      '  Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 4)
    End If
    
End Sub


Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub


Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim iIdx    As Integer
    Dim sFlag   As String
    If Row < 2 Then Exit Sub
    
    sFlag = ""
    ss2.Col = 0
    For iIdx = 1 To ss2.MaxRows
        ss2.Row = iIdx
        If ss2.Text = "Input" Or ss2.Text = "Update" Or ss2.Text = "Delete" Then
            sFlag = "Y"
            iIdx = ss2.MaxRows
        End If
    Next iIdx
    
    If sFlag = "Y" Then
        If vbNo = MsgBox("�Ѿ����޸ĵ���Ϣ....����(Sorting)��?", vbYesNo + vbQuestion, "ȷ��!!") Then Exit Sub
    End If
    
    Set Active_Spread = Me.ss2
    PopupMenu MDIMain.PopUp_Spread
    
    For iIdx = 1 To ss2.MaxRows
        ss2.Col = 0
        ss2.Row = iIdx
        ss2.Text = ""
        ss2.Col = -1
        ss2.BackColor = &HFFFFFF
    Next iIdx
End Sub

Private Sub text_cur_inv_code_DblClick()
    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub Text_PROD_CD_Change()
   
    Select Case text_prod_cd.Text
        Case "S", "s", "SL"
            text_prod_cd.Text = "SL"
'        Case "P", "p", "PP"
'            Text_PROD_CD.Text = "PP"
'        Case "H", "h", "HC"
'            Text_PROD_CD.Text = "HC"
        Case Else
            text_prod_cd.Text = ""
            Call MsgBox("��Ʒ�������" & Chr(10) & "�����Ϲ淶! �������", vbExclamation + vbOKOnly, "����")
    End Select

End Sub

Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)

   Text_PROD_CD_Name.Text = ""
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=text_prod_cd
        DD.rControl.Add Item:=Text_PROD_CD_Name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(text_prod_cd.Text)) = text_prod_cd.MaxLength Then
        Text_PROD_CD_Name.Text = Gf_ComnNameFind(M_CN1, "B0005", text_prod_cd.Text, 2)
    Else
        Text_PROD_CD_Name.Text = ""
    End If
    
End Sub

Private Sub Form_Activate()
    
    Call FormMenuSetting1(Me, FormType, Toolbar_St, sAuthority)
  
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
     sAuthority = Gf_Pgm_Authority(Me.Name)
    
     Call Form_Define
     Call FormMenuSetting1(Me, FormType, "FS", sAuthority)
     Call Gp_Ms_NeceColor(Mc1("nControl"))

'     Call Gp_Ms_Cls(Mc1("rControl"))

     Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)

'
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    txt_emp = sUserID
    Screen.MousePointer = vbDefault
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    Set iSumCol = Nothing
    
    Call FormMenuSetting1(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call FormMenuSetting1(Me, FormType, "CLS", sAuthority)
  
    End If
    
    text_prod_cd.Text = "SL"
    text_stlgrd.Text = ""
    dtp_ins_date_prod_date1.RawData = ""
    dtp_ins_date_prod_date2.RawData = ""
    sdb_wid_to.Value = 0
    sdb_thk_to.Value = 0
    sdb_len_to.Value = 0
    sdb_wid_fr.Value = 0
    sdb_thk_fr.Value = 0
    sdb_len_fr.Value = 0
    txt_plt = ""
    txt_plt_name = ""
    txt_prc_line = "1"
    sdb_slab_num.Value = 0
    sdb_slab_wgt.Value = 0
    ULabel5.Visible = False
    ULabel6.Visible = False
    txt_plt.Visible = False
    txt_plt_name.Visible = False
    txt_prc_line.Visible = False
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

     Dim SMESG As String
     Dim S As String
     Dim i As Long
     
    txt_plt = ""
    txt_plt_name = ""
    txt_prc_line = "1"
    sdb_slab_num.Value = 0
    sdb_slab_wgt.Value = 0
    ULabel5.Visible = False
    ULabel6.Visible = False
    txt_plt.Visible = False
    txt_plt_name.Visible = False
    txt_prc_line.Visible = False
  
     
    If sdb_wid_to.Value = 0 Then sdb_wid_to.Value = 9999.99
    If sdb_thk_to.Value = 0 Then sdb_thk_to.Value = 9999.99
    If sdb_len_to.Value = 0 Then sdb_len_to.Value = 9999999.9
    If dtp_ins_date_prod_date1.RawData = "" Then
       dtp_ins_date_prod_date1.RawData = Format(Date, "YYYYMM") + "01"
    End If
    If dtp_ins_date_prod_date2.RawData = "" Then
       dtp_ins_date_prod_date2.RawData = Format(Date, "YYYYMMDD")
    End If
     
 
   If text_prod_cd.Text = "SL" Then
       sQuery = "Select A.SLAB_NO,A.PROC_CD,A.APLY_STDSPEC,A.STLGRD,A.THK,A.WID,A.LEN,A.WGT,A.QUALITY_UPD_GRD,A.APLY_ENDUSE_CD,A.ORD_FL,TO_DATE(A.PROD_DATE,'YYYY-MM-DD'),"
        sQuery = sQuery + " A.LOC,A.CUST_CD,'','','',A.PROD_CD "
        sQuery = sQuery + "  From FP_SLAB A "
        sQuery = sQuery + "  WHERE A.PROC_CD IN('CAC','XAA', 'XAC') "
'    ElseIf Text_PROD_CD.Text = "PP" Then
'       sQuery = "Select A.PLATE_NO,A.PROC_CD,A.APLY_STDSPEC,A.STLGRD,A.THK,A.WID,A.LEN,A.WGT,A.QUALITY_UPD_GRD,A.APLY_ENDUSE_CD,A.ORD_FL,TO_DATE(A.PROD_DATE,'YYYY-MM-DD'),"
'        sQuery = sQuery + " A.LOC,A.CUST_CD,'','','',A.PROD_CD "
'        sQuery = sQuery + "  From GP_PLATE A "
'        sQuery = sQuery + "  WHERE A.PROC_CD LIKE 'XA%'"
'    ElseIf Text_PROD_CD.Text = "HC" Then
'       sQuery = "Select A.COIL_NO,A.PROC_CD,A.APLY_STDSPEC,A.STLGRD,A.THK,A.WID,A.LEN,A.WGT,A.QUALITY_UPD_GRD,A.APLY_ENDUSE_CD,A.ORD_FL,TO_DATE(A.PROD_DATE,'YYYY-MM-DD'),"
'        sQuery = sQuery + " A.LOC,A.CUST_CD,'','','',A.PROD_CD "
'        sQuery = sQuery + "  From GP_COIL A "
'        sQuery = sQuery + "  WHERE A.PROC_CD LIKE 'XA%'"
    Else
       Call MsgBox("��Ʒ���಻��Ϊ�գ�" & Chr(10) & "�������", vbExclamation + vbOKOnly, "����")
       Exit Sub
    End If
         
    sQuery = sQuery + "   AND A.REC_STS = '2'"
    sQuery = sQuery + "   AND NVL(A.OUT_PLT_CD,'9') <> '0' "
    sQuery = sQuery + "   AND NVL(A.STLGRD,' ')    Like '" + Trim(text_stlgrd.Text) + "%' "
    sQuery = sQuery + "   AND NVL(A.PROD_DATE,' ') BETWEEN '" + Trim(dtp_ins_date_prod_date1.RawData) + "' AND "
    sQuery = sQuery + "                                  '" + Trim(dtp_ins_date_prod_date2.RawData) + "' "
    sQuery = sQuery + "   AND A.WID BETWEEN " & sdb_wid_fr.Value & " AND " & sdb_wid_to.Value
    sQuery = sQuery + "   AND A.THK BETWEEN " & sdb_thk_fr.Value & " AND " & sdb_thk_to.Value
    sQuery = sQuery + "   AND A.LEN BETWEEN " & sdb_len_fr.Value & " AND " & sdb_len_to.Value
    sQuery = sQuery + "   AND NVL(CUR_INV,'*') = '" + Trim(text_cur_inv_code.Text) + "'"
    sQuery = sQuery + "   AND NVL(SIZE_KND,' ') LIKE '" + Trim(Text_size_knd.Text) + "%'"
    sQuery = sQuery + "   AND A.SLAB_NO  LIKE  '" + Trim(txt_heat_no.Text) + "%'"

    
    If text_prod_cd.Text = "SL" Then
       sQuery = sQuery + "ORDER BY A.slab_no"
'    ElseIf Text_PROD_CD.Text = "PP" Then
'       sQuery = sQuery + "ORDER BY A.plate_no"
'    ElseIf Text_PROD_CD.Text = "HC" Then
'       sQuery = sQuery + "ORDER BY A.coil_no"
    End If
       
    SMESG = Gf_Ms_NeceCheck(nControl)
    If SMESG = "OK" Then
    
        SMESG = Gf_Ms_NeceCheck2(mControl)
        If SMESG = "OK" Then
        
            If Gf_Sp_Display(M_CN1, sc1.Item("Spread"), sQuery, sc1.Item("pColumn"), True) Then
                Call FormMenuSetting1(Me, FormType, "RE", sAuthority)
            End If
        Else
            SMESG = SMESG + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(SMESG)
        End If
    Else
        SMESG = SMESG + " Must input necessarily"
        Call Gp_MsgBoxDisplay(SMESG)
    End If
    

End Sub
Public Sub Form_Pro()
Dim i As Long

If txt_plt.Text = "" Then
Gp_MsgBoxDisplay ("���Ϳ����ѡ��")
Exit Sub
End If


 With ss2

    For i = 1 To .MaxRows
        .Col = 0
        .Row = i
        If .Text = "Update" Then

            
           .Col = 15
           .Text = Trim(txt_plt)
           .Col = 16
           .Text = Trim(txt_prc_line)
           .Col = 17
           .Text = Trim(txt_emp)
           .Col = 18
           .Text = Trim(text_prod_cd)
        End If
    Next i
 End With

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub
Private Sub text_cur_inv_code_Change()
If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
      text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
      Exit Sub
Else
      text_cur_inv.Text = ""
End If
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Exit Sub
    Else
        text_cur_inv.Text = ""
End If
    End If
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub
Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
Dim PRE As Long

    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

 If Row < 1 Then Exit Sub
    If ss2.MaxRows < 1 Then Exit Sub
    ss2.Row = Row
    ss2.Col = 0
    
    If ss2.Text <> "Update" Then
       ss2.Col = 0
       ss2.Text = "Update"
       ss2.Col = 7
       sdb_slab_num.Value = sdb_slab_num.Value + 1
       sdb_slab_wgt.Value = sdb_slab_wgt.Value + ss2.Value
       Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row, , &HFFFF80)
   Else
       ss2.Col = 0
       ss2.Text = " "
       ss2.Col = 7
       sdb_slab_num.Value = sdb_slab_num.Value - 1
       sdb_slab_wgt.Value = sdb_slab_wgt.Value - ss2.Value
       Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row)
       PRE = Row
       ss2.Row = PRE - 1
       ss2.Col = 0
       If PRE <> 0 Then
          ss2.Row = Row
          ss2.Text = Trim(Str(Row))
       Else
          ss2.Row = Row
          ss2.Text = "1"
       End If
   
   End If


    ULabel5.Visible = True
    'ULabel6.Visible = True
    txt_plt.Visible = True
    txt_plt_name.Visible = True
    'txt_PRC_line.Visible = True

End Sub

Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub Text_ORD_ITEM_Change()

    If Text_ORD_ITEM.Text <> "" Then
        If Val(Text_ORD_ITEM.Text) > iCount Or Val(Text_ORD_ITEM.Text) < 0 Or Text_ORD_ITEM.Text = "00" Then
            Call MsgBox("����������벻��ȷ!" & Chr(10) & "�����ԡ�", vbExclamation + vbOKOnly, "����")
            Text_ORD_ITEM.Text = ""
        End If
    End If

End Sub

Private Sub Text_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0004"
        DD.nameType = "2"
     
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub
        
    End If
End Sub

Private Sub text_rec_sts_KeyUp(KeyCode As Integer, Shift As Integer)
  
    Text_REC_STS_Name = ""
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "Z0005"

  
        DD.nameType = "2"
    
        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub
        
    End If

End Sub

Private Sub Text_size_knd_DblClick()
    Call Text_size_knd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub text_stlgrd_DblClick()
    Call text_stlgrd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub text_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

         DD.sWitch = "MS"
         DD.rControl.Add Item:=text_stlgrd
        
         DD.nameType = "1"
         Call Gf_Stlgrd_DD(M_CN1, KeyCode)
         Exit Sub

    End If
        
End Sub

Private Sub Text_size_knd_Change()
    If Len(Trim(Text_size_knd.Text)) = Text_size_knd.MaxLength Then
        Text_size_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0043", Text_size_knd.Text, 2)
        Exit Sub
    Else
        Text_size_knd_name.Text = ""
    End If
End Sub

Private Sub Text_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=Text_size_knd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_PLT_Change()
    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
          txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_plt.Text, 2)
          Exit Sub
    Else
          txt_plt_name.Text = ""
    End If
End Sub

Private Sub txt_plt_DblClick()
    Call txt_plt_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_plt
       ' DD.rControl.Add Item:=txt_PLT_NAME
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

'    If Len(Trim(txt_PLT.Text)) = txt_PLT.MaxLength Then
'        txt_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_PLT.Text), 2)
'    Else
'        txt_PLT_NAME.Text = ""
'    End If

End Sub

Private Sub txt_stdgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    'If KeyCode = vbKeyF4 Then
        
        'DD.nameType = "1"
        'DD.sWitch = "MS"
        
        'DD.rControl.Add Item:=TxT_stdgrd
        'Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    'End If
    
End Sub

 
  
Private Function Gf_Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

'On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim SMESG As String
    Dim sTemp As String
    Dim ProcessChk As String
    Dim DelYN As Boolean
    Dim Msg_Count As Integer
    Dim Msg_Yes As String
    
    Dim adoCmd As ADODB.Command

    Gf_Sp_Process = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
        Gf_Sp_Process = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(iCount, Sc, MC)
                End If
                
                'Maxlength Check
                SMESG = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
                        
                If Trim(SMESG) = "OK" Then
                    
                ElseIf Mid(SMESG, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    SMESG = Mid(SMESG, 6, Len(SMESG))
                    SMESG = SMESG + "���Ȳ���ȷ"
                    Call Gp_MsgBoxDisplay(SMESG)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    SMESG = SMESG + "��������"
                    Call Gp_MsgBoxDisplay(SMESG)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Sp_Process = False: Screen.MousePointer = vbDefault: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To Sc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    Msg_Count = 1
    For iCount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        DelYN = False
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input"
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).Value = "D"
                If Msg_Count = 1 Then
                   DelYN = Gf_MessConfirm("��ȷ��Ҫɾ��״̬Ϊ[Delete]��������", "Q")
                   If DelYN Then Msg_Yes = "yes"
                   Msg_Count = Msg_Count + 1
                End If
                If Msg_Yes = "yes" Then DelYN = True
        End Select
          
        If ProcessChk = "YES" Or DelYN Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempFloat)
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempInt)
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Text = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
                        
                End Select
                
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Sp_Process = False
                Exit Function
        
             End If
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input", "Update"
                Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                
            Case "Delete"
                If DelYN Then
                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), iCount)
                   iCount = iCount - 1
                End If
        End Select
        
    Next iCount
    
    Sc.Item("Spread").ReDraw = True
    
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            If RefChek = False Then Call Form_Ref
                                                    
        Else
            If RefChek = False Then Screen.MousePointer = vbDefault: Exit Function
        End If
        
        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ���ɹ�������" & iProcessCount & "����¼"
        'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
        
    End If
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Gf_Sp_Process = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_Process = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

Private Sub Gp_Sp_Collection1(sPname As Variant, Num As Integer, pcol As String, ncol As String, mcol As String, _
                                                               iCol As String, acol As String, lCol As String, _
                            pColumn As Collection, nColumn As Collection, mColumn As Collection, iColumn As Collection, _
                            aColumn As Collection, lColumn As Collection)
   
    If LCase(Trim(pcol)) = "p" Then       'PK Column
        pColumn.Add Item:=Num
    End If
    
    If LCase(Trim(ncol)) = "n" Then       'Necessary Column
        nColumn.Add Item:=Num
        'Call Gp_Sp_ColColor(SpName, Num, , &H80FF80)
    End If
    
    If LCase(Trim(mcol)) = "m" Then       'Spread Maxlength check Column
        mColumn.Add Item:=Num
    End If
    
    If LCase(Trim(iCol)) = "i" Then       'Spread Insert Column
        iColumn.Add Item:=Num
        
    End If
    
    If LCase(Trim(acol)) = "a" Then       'Master -> Spread Column
        aColumn.Add Item:=Num
        Call Gp_Sp_ColHidden(sPname, Num, True)
    End If
    
    If LCase(Trim(lCol)) = "l" Then       'Spread Lock Column
        lColumn.Add Item:=Num
        Call Gp_Sp_ColLock(sPname, Num, True)
    End If

    
End Sub
Public Sub FormMenuSetting1(Fm As Variant, FormType As String, ButtonType As String, sAuthority As String)



On Error Resume Next
    
    With MDIMain.MenuTool
    
        Select Case FormType
              
               Case "Start"
                    .Buttons(1).Enabled = False                 'Screen Clear
                    .Buttons(2).Enabled = False                 'Refer
                    .Buttons(3).Enabled = False                 'Separator
                    .Buttons(4).Enabled = False                 'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = False                 'Separator
                    .Buttons(7).Enabled = False                 'Row Insert
                    .Buttons(8).Enabled = False                 'Row Delete
                    .Buttons(9).Enabled = False                 'Row Cancel
                    .Buttons(10).Enabled = False                'Separator
                    .Buttons(11).Enabled = False                'Copy
                    .Buttons(12).Enabled = False                'Paste
                    .Buttons(13).Enabled = False                'Separator
                    .Buttons(14).Enabled = False                'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = False                'Separator
                    .Buttons(17).Visible = True                 'Exit
                    
                  Case "Msheet"
                    .Buttons(1).Enabled = True                  'Screen Clear
                    .Buttons(2).Enabled = True                  'Refer
                    .Buttons(3).Enabled = True                  'Separator
                    .Buttons(4).Enabled = True                  'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = True                  'Separator
                    .Buttons(7).Enabled = False                 'Row Insert
                    .Buttons(8).Enabled = False                 'Row Delete
                    .Buttons(9).Enabled = False                 'Row Cancel
                    .Buttons(10).Enabled = True                 'Separator
                    
                    .Buttons(11).Enabled = False                'Copy
                    .Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
                    .Buttons(11).ButtonMenus(2).Enabled = False 'Master Copy
                    .Buttons(11).ButtonMenus(3).Enabled = True  'Spread Copy
                    
                    .Buttons(12).Enabled = False                 'Paste
                    .Buttons(12).ButtonMenus(1).Enabled = False 'All Paste
                    .Buttons(12).ButtonMenus(2).Enabled = False 'Master Paste
                    .Buttons(12).ButtonMenus(3).Enabled = False 'Spread Paste
                    
                    .Buttons(13).Enabled = True                 'Separator
                    .Buttons(14).Enabled = True                 'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = True                 'Separator
                    .Buttons(17).Enabled = True                 'Exit
                
        End Select
        
        Fm.Toolbar_St = ButtonType
        
        Select Case ButtonType
                 'Save, Refer
            Case "SE", "RE"
                
                Select Case FormType
                                        
                    Case "Msheet"
                        .Buttons(7).Enabled = False              'Row Insert
                        .Buttons(8).Enabled = False              'Row Delete
                        .Buttons(9).Enabled = False             'Row Cancel
                        .Buttons(14).Enabled = True             'Excel
                     End Select
                
                 'Form Start, Screen Clear
            Case "FS", "CLS"
                
                Select Case FormType

                    Case "Msheet"
                        .Buttons(7).Enabled = False              'Row Insert
                        .Buttons(8).Enabled = False             'Row Delete
                        .Buttons(9).Enabled = False              'Row Cancel
                        .Buttons(14).Enabled = False            'Excel
                                        
                End Select
                
            Case "Acopy"
            
                .Buttons(12).ButtonMenus(1).Enabled = True      'All Paste
                .Buttons(12).ButtonMenus(2).Enabled = False     'Master Paste
                .Buttons(12).ButtonMenus(3).Enabled = False     'Spread Paste
                
            Case "Mcopy"
            
                .Buttons(12).ButtonMenus(1).Enabled = False     'All Paste
                .Buttons(12).ButtonMenus(2).Enabled = True      'Master Paste
                .Buttons(12).ButtonMenus(3).Enabled = False     'Spread Paste
                
            Case "Scopy"
            
                .Buttons(12).ButtonMenus(1).Enabled = False     'All Paste
                .Buttons(12).ButtonMenus(2).Enabled = False     'Master Paste
                .Buttons(12).ButtonMenus(3).Enabled = True      'Spread Paste
                
        End Select
        
        'Autority Inquiry Check
        If Mid(sAuthority, 1, 1) = "0" Then
            .Buttons(2).Enabled = False                         'Refer
        End If
        
        Select Case Mid(sAuthority, 2, 3) 'Insert, Update, Delete
        
            Case "000"      'No Authority
                .Buttons(4).Enabled = False                     'Save
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(8).Enabled = False                     'Row Delete
                .Buttons(9).Enabled = False                     'Row Cancel
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste
            
            Case "001"      'Delete Authority
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste
            
            Case "010"      'Update Authority
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(8).Enabled = False                     'Row Delete
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste
            
            Case "011"      'Update, Delete Authority
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste
            
            Case "100"      'Insert Authority
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(8).Enabled = False                     'Row Delete
            
            Case "101"      'Insert, Delete Authority
            
            Case "110"      'Insert, Update Authority
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(8).Enabled = False                     'Row Delete
            
            Case "111"      'Insert, Update, Delete Authority
        
        End Select
        
        .Wrappable = True
        
    End With

End Sub