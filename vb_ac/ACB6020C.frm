VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB6020C 
   Caption         =   "����ת��ƻ�ʵ��¼��_ACB6020C"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin InDate.UDate dtp_ins_date_prod_date2 
      Height          =   315
      Left            =   8590
      TabIndex        =   10
      Tag             =   "INS_DATE"
      Top             =   480
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
   Begin VB.TextBox txt_mill_plt_name 
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
      Left            =   13410
      TabIndex        =   30
      Top             =   885
      Width           =   1560
   End
   Begin VB.TextBox txt_mill_plt 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
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
      Left            =   12930
      MaxLength       =   2
      TabIndex        =   29
      Tag             =   "����"
      Top             =   885
      Width           =   450
   End
   Begin VB.TextBox txt_ORD_ITEM 
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
      Left            =   14385
      MaxLength       =   2
      TabIndex        =   28
      Top             =   480
      Width           =   345
   End
   Begin VB.TextBox txt_ORD_NO 
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
      Left            =   12930
      MaxLength       =   11
      TabIndex        =   27
      Top             =   480
      Width           =   1410
   End
   Begin VB.TextBox Text_PROC_CD_Name 
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
      Left            =   13920
      MaxLength       =   11
      TabIndex        =   26
      Top             =   135
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox text_PROC_CD 
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
      Left            =   12930
      MaxLength       =   11
      TabIndex        =   23
      Top             =   135
      Width           =   945
   End
   Begin VB.TextBox text_stlgrd 
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
      Left            =   1470
      MaxLength       =   11
      TabIndex        =   8
      Top             =   510
      Width           =   1755
   End
   Begin VB.TextBox txt_slab_no 
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
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "������"
      Top             =   510
      Width           =   1320
   End
   Begin VB.ComboBox cbo_hcr 
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
      ItemData        =   "ACB6020C.frx":0000
      Left            =   6525
      List            =   "ACB6020C.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   510
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox text_prod_cd 
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
      Left            =   1470
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "��Ʒ"
      Text            =   "SL"
      Top             =   135
      Width           =   375
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
      Left            =   5160
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "�ֿ�"
      Top             =   135
      Width           =   375
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
      Left            =   5535
      TabIndex        =   3
      Top             =   135
      Width           =   1560
   End
   Begin VB.TextBox txt_prod_plt 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
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
      Left            =   8580
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "����"
      Top             =   135
      Width           =   450
   End
   Begin VB.TextBox txt_prod_plt_nm 
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
      Left            =   9030
      TabIndex        =   1
      Top             =   135
      Width           =   1560
   End
   Begin VB.TextBox txt_plan_plt_nm 
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
      Left            =   2100
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   1560
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   1
      Left            =   240
      Top             =   135
      Width           =   1170
      _ExtentX        =   2064
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   3945
      Top             =   135
      Width           =   1170
      _ExtentX        =   2064
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   7365
      Top             =   135
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   7365
      Top             =   495
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
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.UDate dtp_ins_date_prod_date1 
      Height          =   315
      Left            =   10200
      TabIndex        =   9
      Tag             =   "INS_DATE"
      Top             =   480
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   240
      Top             =   510
      Width           =   1170
      _ExtentX        =   2064
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   3945
      Tag             =   "������"
      Top             =   510
      Width           =   1170
      _ExtentX        =   2064
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
      ForeColor       =   -2147483641
   End
   Begin CSTextLibCtl.sidbEdit sdb_thk_fr 
      Height          =   315
      Left            =   1470
      TabIndex        =   11
      Top             =   885
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      Left            =   2445
      TabIndex        =   12
      Top             =   885
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   240
      Top             =   885
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   3945
      Top             =   885
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   5160
      TabIndex        =   13
      Top             =   885
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      Left            =   6135
      TabIndex        =   14
      Top             =   885
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   7380
      Top             =   885
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   8580
      TabIndex        =   15
      Top             =   885
      Width           =   1110
      _Version        =   262145
      _ExtentX        =   1958
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
      Left            =   9690
      TabIndex        =   16
      Top             =   885
      Width           =   1110
      _Version        =   262145
      _ExtentX        =   1958
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   120
      TabIndex        =   17
      Top             =   1275
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   979
      _Version        =   196609
      BackColor       =   14737918
      BorderWidth     =   1
      BevelOuter      =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_priority 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12840
         MaxLength       =   1
         TabIndex        =   33
         Tag             =   "ת�����ȼ�"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txt_priority_name 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   13320
         TabIndex        =   32
         Tag             =   "���ȼ�"
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txt_plt 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   19
         Tag             =   "Ŀ���"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txt_plt_name 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Tag             =   "�� ��"
         Top             =   120
         Width           =   1575
      End
      Begin InDate.ULabel ULabel16 
         Height          =   375
         Left            =   120
         Tag             =   "Ŀ���"
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Ŀ���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel14 
         Height          =   375
         Left            =   3720
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "ת��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.UDate udt_due_move_date 
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
      End
      Begin CSTextLibCtl.sitxEdit stx_due_move_date_time 
         Height          =   375
         Left            =   6360
         TabIndex        =   22
         Top             =   120
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   125
         Text            =   "__:__:__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "__:__:__"
         StartText.x     =   3
         StartText.y     =   5
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
         Mask            =   "%%:%%:%%"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   6
      End
      Begin InDate.ULabel ULabel5 
         Height          =   375
         Left            =   7800
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "��ѡ����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_num 
         Height          =   375
         Left            =   9000
         TabIndex        =   24
         Top             =   120
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   255
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
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
         StartText.y     =   5
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
         Height          =   375
         Left            =   9840
         TabIndex        =   25
         Top             =   120
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
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
         StartText.y     =   5
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
      Begin Threed.SSPanel SSP90 
         Height          =   375
         Left            =   14640
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���º�ͬ"
         FloodColor      =   65535
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin InDate.ULabel ULabe21 
         Height          =   375
         Left            =   11640
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "ת�����ȼ�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   6990
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Width           =   15150
      _Version        =   393216
      _ExtentX        =   26723
      _ExtentY        =   12330
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
      MaxCols         =   31
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACB6020C.frx":001D
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   11730
      Top             =   135
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "���̴���"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   11730
      Top             =   480
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "������"
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   11730
      Top             =   885
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "ʹ�ù���"
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
End
Attribute VB_Name = "ACB6020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name   ����ת��ƻ�¼��
'-- Program Name
'-- Program ID        ACB6020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          YIDUJUN
'-- Coder
'-- Date              2011.3.9
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
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

Dim Mc1 As New Collection           'Master Collection
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
        Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(text_cur_inv, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_prod_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_prod_plt_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(Text_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(text_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(dtp_ins_date_prod_date1, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(dtp_ins_date_prod_date2, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_mill_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_plt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(udt_due_move_date, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_slab_num, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_slab_wgt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_priority, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_priority_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
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
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '�ص��ͬ
    Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    
    
    sc1.Add Item:=ss2, Key:="Spread"
    sc1.Add Item:="ACB6020C.P_SREFER", Key:="P-R"
    sc1.Add Item:="ACB6020C.P_MODIFY", Key:="P-M"
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
 
    Call Gp_Sp_ColHidden(ss2, 3, True)
    Call Gp_Sp_ColHidden(ss2, 10, True)
    Call Gp_Sp_ColHidden(ss2, 20, True)
    Call Gp_Sp_ColHidden(ss2, 21, True)
    Call Gp_Sp_ColHidden(ss2, 22, True)   'caolei priority ת�����ȼ�  20130305
    Call Gp_Sp_ColHidden(ss2, 23, True)
    
End Sub



Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    Dim iIdx    As Integer
    Dim sFlag   As String
    If ROW < 2 Then Exit Sub
    
    sFlag = ""
    ss2.Col = 0
    For iIdx = 1 To ss2.MaxRows
        ss2.ROW = iIdx
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
        ss2.ROW = iIdx
        ss2.Text = ""
        ss2.Col = -1
        ss2.BackColor = &HFFFFFF
    Next iIdx
    
End Sub

'Private Sub stx_due_move_date_time_DblClick()
'
'    stx_due_move_date_time.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'HH24MISS') FROM DUAL")
'
'End Sub

'Private Sub stx_plan_move_time_DblClick()
'
'    stx_plan_move_time.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'HH24MISS') FROM DUAL")
'
'End Sub

Private Sub text_cur_inv_code_DblClick()
    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
End Sub

'Private Sub txt_priority_DblClick()        '22222222222
'    Call txt_priority_KeyUp(vbKeyF4, 0)
'End Sub
'
'Private Sub txt_priority_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'         DD.sWitch = "MS"
'         DD.rControl.Add Item:=text_priority_cd
'
'         DD.nameType = "1"
'         Call Gf_Stlgrd_DD(M_CN1, KeyCode)
'         Exit Sub
'
'    End If
'
'End Sub

Private Sub Text_PROC_CD_DblClick()
    Call Text_PROC_CD_KeyUp(vbKeyF4, 0)
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

   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=text_prod_cd
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
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
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)

    
    text_cur_inv_code.Text = "00"
    Call text_cur_inv_code_KeyUp(0, 0)
    txt_plt.Text = "ZB"
    txt_plt_name.Text = "�а�"
    txt_mill_plt.Text = "C3"
    txt_mill_plt_name.Text = "�а峧"
'    txt_priority.Text = "2"
'    txt_priority_name.Text = "δָʾ"
'
    
'    dtp_ins_date_prod_date1.Text = Format(DateAdd("d", -2, CDate(dtp_ins_date_prod_date2.Text)), "YYYY-MM-DD")
    udt_due_move_date.Text = Format(DateAdd("d", 1, CDate(Now)), "YYYY-MM-DD")


    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
     
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    Set iSumCol = Nothing
    
    Call FormMenuSetting1(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call FormMenuSetting1(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        text_prod_cd.Text = "SL"
        text_cur_inv_code.Text = "00"
        Call text_cur_inv_code_KeyUp(0, 0)
        txt_plt.Text = "ZB"
        txt_plt_name.Text = "�а�"
        txt_mill_plt.Text = "C3"
        txt_mill_plt_name.Text = "�а峧"
'        txt_priority.Text = "2"
'        txt_priority_name.Text = "δָʾ"
'        Call txt_plt_KeyUp(0, 0)
'        dtp_ins_date_prod_date1.Text = Format(DateAdd("d", -2, CDate(dtp_ins_date_prod_date2.Text)), "YYYY-MM-DD")
        udt_due_move_date.Text = Format(DateAdd("d", 1, CDate(Now)), "YYYY-MM-DD")
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

 Dim ForCnt As Integer
    Dim tmWgt As Long
    Dim tmLen As Long
    Dim lRow As Long
    Dim sBlockSeq As String
    Dim iRow As Integer
    Dim i As Integer
    Dim TIME As String

    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
    
    sdb_slab_num.Value = 0
    sdb_slab_wgt.Value = 0
        
    If txt_mill_plt.Text <> "" And txt_mill_plt.Text <> "C1" And txt_mill_plt.Text <> "C2" And txt_mill_plt.Text <> "C3" Then
        Call Gp_MsgBoxDisplay("ʹ�ù�������Ϊ'C1'����'C2'����'C3'")
        Exit Sub
    End If
    
     
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss2.OperationMode = OperationModeNormal
        Call FormMenuSetting1(Me, FormType, "RE", sAuthority)
    End If
    
'       ��Ҫ��ͬ���
    Call SS2_CHANGE_COLOR
    
'      TIME = Format(Now, "YYYY-MM")
'     For iRow = 1 To ss2.MaxRows
'
'      ss2.Row = iRow
'      ss2.Col = 26
'          If ss2.Text <> "" Then
'
''      ���º�ͬ��ʾ��ɫ
'        If Mid(ss2.Text, 1, 7) < TIME Then
'          For i = 1 To ss2.MaxCols
'               ss2.Col = i
'               ss2.BackColor = &HC0C0FF
'          Next
'
'
'       End If
'    End If
'
'    Next iRow
'
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    
    
End Sub

Public Sub Form_Pro()

    Dim MvNo        As String
    Dim TransNo     As String
    Dim iRow        As Integer
    
    If Trim(txt_plt.Text) = "" Or Trim(txt_plt_name.Text) = "" Then
        Call Gp_MsgBoxDisplay("����Ŀ���...")
        Exit Sub
    End If

    
    If Len(udt_due_move_date.RawData) <> 8 Then
        Call Gp_MsgBoxDisplay("����ת��������...")
        Exit Sub
    End If
    
    If Trim(text_cur_inv_code.Text) = Trim(txt_plt.Text) Then
        Call Gp_MsgBoxDisplay("��ʼ�� = Ŀ���")
        Exit Sub
    End If
    
    If UCase(Trim(txt_plt.Text)) = "ZZ" Then
        Call Gp_MsgBoxDisplay("����Ŀ���")
        Exit Sub
    End If

            
    iCount = 0
    For iRow = 1 To ss2.MaxRows
        ss2.ROW = iRow
        ss2.Col = 0
        If ss2.Text = "Update" Then
                        
            ss2.Col = 20
            ss2.Text = txt_plt.Text
            
            ss2.Col = 21
            ss2.Text = udt_due_move_date.RawData & Left(stx_due_move_date_time.RawData, 2) & "0000"
            
            ss2.Col = 22
            ss2.Text = txt_priority.Text      'add at 20130305
            
            ss2.Col = 23
            ss2.Text = sUserID
            
            
            
        End If
        
        
    Next iRow

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call FormMenuSetting1(Me, FormType, "RE", sAuthority)

End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Private Sub SS2_CHANGE_COLOR()

    With ss2
      
        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .ROW = iCount
            
            '�ص��ͬ
            ss2.ROW = .ROW:       ss2.Col = 27
            If ss2.Text = "Y" Then

                 Call Gp_Sp_RowColor(ss2, .ROW, , &HFF&)
            End If
   
        Next iCount

    End With
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        
        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    Else
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
           text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
           Exit Sub
        Else
           text_cur_inv.Text = ""
        End If
        
    End If
    
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim i As Integer
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

    Dim Row1 As Long
    Dim Row2 As Long
    Dim Col As Long
    
    Dim str_ord_fl As String
    Dim str_rec_sts As String
    
    If Not Gf_Sc_Authority(sAuthority, "U") Then
       Exit Sub
    End If
    
    Col = BlockCol
    Row1 = BlockRow
    Row2 = BlockRow2
  
    If Col = -1 Then

     For i = BlockRow To BlockRow2
        Call ss2_row_Click(1, i)
'         Call ss2_Click(1, i)
     Next
     
   End If

'   Call ss2.SetActiveCell(1, Row2)

End Sub
Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)

 Dim ForCnt As Integer
    Dim tmWgt As Long
    Dim tmLen As Long
    
    Dim lRow As Long
    Dim sBlockSeq As String
    Dim iRow As Integer
    Dim i As Integer
    Dim TIME As String
    
    
'  TIME = Format(Now, "YYYY-MM")
'
'
'     For iRow = 1 To ss2.MaxRows
'
'      ss2.Row = iRow
'      ss2.Col = 26
'
'      If ss2.Text <> "" Then
'
'        If Mid(ss2.Text, 1, 7) < TIME Then
'          For i = 1 To ss2.MaxCols
'               ss2.Col = i
'               ss2.BackColor = &HFFC0CF
'          Next
'
'       End If
'      End If
'
'    Next iRow
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)


End Sub

Private Sub ss2_row_Click(ByVal Col As Long, ByVal ROW As Long)

    Dim PRE As Long

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0

    lBlkrow2 = 0

    If ROW < 1 Then Exit Sub

    If ss2.MaxRows < 1 Then Exit Sub

    ss2.ROW = ROW
    ss2.Col = 0

    If ss2.Text <> "Update" Then

'        ss2.Col = 10
'        If Trim(ss2.Text) = "N" Or Trim(ss2.Text) = "S" Then Exit Sub

        ss2.Col = 0
        ss2.Text = "Update"
        ss2.Col = 9
        sdb_slab_num.Value = sdb_slab_num.Value + 1
        sdb_slab_wgt.Value = sdb_slab_wgt.Value + ss2.Value
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ROW, ROW, , &HFFFF80)

    Else           '  #################################################################################################

       ss2.Col = 0
       ss2.Text = " "
       ss2.Col = 9
       sdb_slab_num.Value = sdb_slab_num.Value - 1
       sdb_slab_wgt.Value = sdb_slab_wgt.Value - ss2.Value

       Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ROW, ROW)
       

       ss2.Col = 27
     If ss2.Text = "Y" Then
       Call Gp_Sp_RowColor(ss2, ROW, , &HFF&)
     End If
       PRE = ROW
       ss2.ROW = PRE - 1
       ss2.Col = 0

       If PRE <> 0 Then
          ss2.ROW = ROW
          ss2.Text = Trim(Str(ROW))
       Else
          ss2.ROW = ROW
          ss2.Text = "1"
       End If

    End If

End Sub

Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub Text_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0004"

        DD.rControl.Add Item:=Text_PROC_CD
        DD.rControl.Add Item:=Text_PROC_CD_Name
   
        DD.nameType = "2"
        'DD.nameType="1" ���������Ʋ�ѯ
        'DD.nameType="2" ��Ӣ�����Ʋ�ѯ
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() ���ڿͻ�����
        Exit Sub
        
    End If

    If Len(Trim(Text_PROC_CD.Text)) = Text_PROC_CD.MaxLength Then
       '  Gf_ComnNAME_Find( �����ַ���, DD.sKEy���� ,DD.nameType)
       ' Gf_CustNameFind( �����ַ���, �ͻ���������,DD.nameType)
        Text_PROC_CD_Name.Text = Gf_ComnNameFind(M_CN1, "C0004", Text_PROC_CD.Text, 2)
    Else
        Text_PROC_CD_Name.Text = ""
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
        DD.rControl.Add Item:=txt_plt_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    Else

        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_plt.Text, 2)
            Exit Sub
        Else
            txt_plt_name.Text = ""
        End If
        
    End If
        
        
End Sub

Private Function Gf_Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

'On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim sMesg As String
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
                sMesg = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "���Ȳ���ȷ"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = sMesg + "��������"
                    Call Gp_MsgBoxDisplay(sMesg)
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

Private Sub txt_priority_DblClick()

    Call txt_priority_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_priority_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0022"
        DD.rControl.Add Item:=txt_priority
        DD.rControl.Add Item:=txt_priority_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    Else

        If Len(Trim(txt_priority.Text)) = txt_priority.MaxLength Then
            txt_priority_name.Text = Gf_ComnNameFind(M_CN1, "C0022", Trim(txt_priority.Text), 2)
            Exit Sub
        Else
            txt_priority_name.Text = ""
        End If
        
    End If
        
        
End Sub

Private Sub txt_prod_plt_DblClick()

    Call txt_prod_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_prod_plt
        DD.rControl.Add Item:=txt_prod_plt_nm

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(txt_prod_plt)) = txt_prod_plt.MaxLength Then
        txt_prod_plt_nm.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_prod_plt.Text), 2)
    Else
        txt_prod_plt_nm.Text = ""
    End If
    
End Sub
Private Sub txt_mill_plt_DblClick()

    Call txt_mill_plt_KeyUp(vbKeyF4, 0)
    
End Sub
Private Sub txt_mill_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_mill_plt
        DD.rControl.Add Item:=txt_mill_plt_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(txt_mill_plt)) = txt_mill_plt.MaxLength Then
        txt_mill_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_mill_plt.Text), 2)
    Else
        txt_mill_plt_name.Text = ""
    End If
    
End Sub
