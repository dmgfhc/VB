VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACE1065C 
   Caption         =   "物料替代"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   38831.59
   ScaleMode       =   0  'User
   ScaleWidth      =   41689.23
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ord_ord_item 
      Height          =   300
      Left            =   7125
      TabIndex        =   31
      Top             =   90
      Width           =   735
   End
   Begin VB.TextBox ord_ord_no 
      Height          =   315
      Left            =   5655
      MaxLength       =   11
      TabIndex        =   30
      Top             =   75
      Width           =   1470
   End
   Begin CSTextLibCtl.sidbEdit ord_slab_wid_max 
      Height          =   390
      Left            =   5070
      TabIndex        =   29
      Top             =   6540
      Visible         =   0   'False
      Width           =   1260
      _Version        =   262145
      _ExtentX        =   2222
      _ExtentY        =   688
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   2
      StartText.y     =   6
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      FmtThousands    =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit ord_slab_wid_min 
      Height          =   360
      Left            =   3450
      TabIndex        =   28
      Top             =   6555
      Visible         =   0   'False
      Width           =   1320
      _Version        =   262145
      _ExtentX        =   2328
      _ExtentY        =   635
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   2
      StartText.y     =   5
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      FmtThousands    =   0
      Undo            =   0
      Data            =   0
   End
   Begin VB.TextBox prod_no 
      Height          =   315
      Left            =   2430
      TabIndex        =   27
      Top             =   4230
      Width           =   1260
   End
   Begin VB.TextBox prod_txt_prod_cd 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1035
      MaxLength       =   2
      TabIndex        =   26
      Top             =   4230
      Width           =   585
   End
   Begin VB.TextBox ord_txt_prod_cd 
      BackColor       =   &H00C0FFFF&
      Height          =   310
      Left            =   1140
      MaxLength       =   2
      TabIndex        =   25
      Top             =   59
      Width           =   615
   End
   Begin VB.TextBox prod_ord_no 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8700
      MaxLength       =   11
      TabIndex        =   16
      Tag             =   "订单号"
      Top             =   4230
      Width           =   1380
   End
   Begin VB.TextBox prod_txt_stlgrd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4935
      MaxLength       =   12
      TabIndex        =   15
      Tag             =   "钢种"
      Top             =   4230
      Width           =   2535
   End
   Begin VB.TextBox prod_loc 
      Height          =   315
      Left            =   12630
      TabIndex        =   14
      Top             =   4230
      Width           =   2415
   End
   Begin VB.ComboBox prod_ord_itm 
      Enabled         =   0   'False
      Height          =   300
      Left            =   10080
      TabIndex        =   11
      Top             =   4230
      Width           =   735
   End
   Begin VB.TextBox ord_TxT_STLGRD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2610
      MaxLength       =   11
      TabIndex        =   0
      Top             =   62
      Width           =   1500
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   45
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "产品"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Index           =   0
      Left            =   9150
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "订单输入日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.UDate UDate_DEL_TO_b 
      Height          =   315
      Left            =   10485
      TabIndex        =   1
      Tag             =   "INS_DATE"
      Top             =   60
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   1830
      Top             =   60
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      Caption         =   "钢种"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Index           =   0
      Left            =   45
      Top             =   435
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "产品厚度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Index           =   0
      Left            =   4545
      Top             =   435
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "产品宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   1
      Left            =   9150
      Top             =   435
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "产品长度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit ord_prod_thk_fr 
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   435
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
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
      StartText.y     =   4
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit ord_prod_thk_to 
      Height          =   315
      Left            =   2625
      TabIndex        =   3
      Top             =   435
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "9999.99"
      Text            =   " 9,999.99"
      StartText.x     =   3
      StartText.y     =   4
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   9999.99
   End
   Begin CSTextLibCtl.sidbEdit ord_prod_wid_fr 
      Height          =   315
      Left            =   5655
      TabIndex        =   4
      Top             =   435
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit ord_prod_wid_to 
      Height          =   315
      Left            =   7125
      TabIndex        =   5
      Top             =   435
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "99999"
      Text            =   " 99,999"
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   99999
   End
   Begin CSTextLibCtl.sidbEdit ord_prod_len_fr 
      Height          =   315
      Left            =   10260
      TabIndex        =   6
      Top             =   435
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit ord_prod_len_to 
      Height          =   315
      Left            =   11730
      TabIndex        =   7
      Top             =   435
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "9999999"
      Text            =   " 9,999,999"
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   9999999
   End
   Begin Threed.SSCommand cmd_confirm 
      Height          =   390
      Left            =   13590
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   420
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   688
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定处理"
   End
   Begin Threed.SSCommand Command_REP 
      Height          =   390
      Left            =   13590
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   688
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "替代处理"
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_wgt_to 
      Height          =   315
      Left            =   13830
      TabIndex        =   12
      Top             =   4590
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "99999.99"
      Text            =   " 0.00"
      StartText.x     =   3
      StartText.y     =   4
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      Undo            =   0
      Data            =   99999.99
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_wgt_fr 
      Height          =   315
      Left            =   12630
      TabIndex        =   13
      Top             =   4590
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
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
      StartText.y     =   4
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      Undo            =   0
      Data            =   0
   End
   Begin FPSpread.vaSpread prod_ss 
      Height          =   4320
      Left            =   30
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4995
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   7620
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
      MaxCols         =   15
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACE1040C.frx":0000
      VisibleCols     =   1
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   1
      Left            =   45
      Top             =   4230
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "产品"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Index           =   1
      Left            =   3945
      Top             =   4230
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "钢种"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   7725
      Top             =   4230
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "订单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_thk_fr 
      Height          =   315
      Left            =   1035
      TabIndex        =   18
      Tag             =   "产品厚度（MIN）"
      Top             =   4590
      Width           =   1395
      _Version        =   262145
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Index           =   1
      Left            =   45
      Top             =   4590
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "产品厚度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   3945
      Top             =   4590
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "产品宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   0
      Left            =   7725
      Top             =   4590
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "产品长度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_thk_to 
      Height          =   315
      Left            =   2430
      TabIndex        =   19
      Tag             =   "产品厚度（MAX）"
      Top             =   4590
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "9999.99"
      Text            =   " 9,999.99"
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
      Data            =   9999.99
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_len_fr 
      Height          =   315
      Left            =   8685
      TabIndex        =   20
      Tag             =   "产品长度（MIN）"
      Top             =   4590
      Width           =   1395
      _Version        =   262145
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
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
   Begin CSTextLibCtl.sidbEdit prod_prod_len_to 
      Height          =   315
      Left            =   10080
      TabIndex        =   21
      Tag             =   "产品长度（MIN）"
      Top             =   4590
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
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
      RawData         =   "9999999"
      Text            =   " 9,999,999"
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
      Data            =   9999999
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_wid_fr 
      Height          =   315
      Left            =   4920
      TabIndex        =   22
      Tag             =   "产品宽度（MIN）"
      Top             =   4590
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
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
   Begin CSTextLibCtl.sidbEdit prod_prod_wid_to 
      Height          =   315
      Left            =   6195
      TabIndex        =   23
      Tag             =   "产品宽度（MAX）"
      Top             =   4590
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "999999"
      Text            =   " 999,999"
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
      Data            =   999999
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Index           =   1
      Left            =   11655
      Top             =   4230
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "物料位置"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   2
      Left            =   11670
      Top             =   4590
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "产品重量"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpread.vaSpread ord_ss 
      Height          =   2820
      Left            =   45
      TabIndex        =   9
      Top             =   855
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   4974
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   25
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACE1040C.frx":1CBB
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Index           =   2
      Left            =   1620
      Top             =   4230
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   556
      Caption         =   "物料号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Left            =   4545
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "订单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line3 
      X1              =   32.826
      X2              =   32273.59
      Y1              =   10082.59
      Y2              =   10082.59
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   0
      X2              =   32273.59
      Y1              =   10005.73
      Y2              =   10044.16
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   0
      X2              =   32273.59
      Y1              =   9967.298
      Y2              =   9967.298
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   25183.14
      Y1              =   8664.178
      Y2              =   8664.178
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   25183.14
      Y1              =   8608.28
      Y2              =   8608.28
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   24.62
      X2              =   25210.49
      Y1              =   8548.889
      Y2              =   8548.889
   End
   Begin VB.Line Line2 
      X1              =   120.363
      X2              =   7394.093
      Y1              =   5037.8
      Y2              =   5037.8
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "——"
      Height          =   255
      Left            =   9945
      TabIndex        =   24
      Top             =   4260
      Width           =   255
   End
End
Attribute VB_Name = "ACE1065C"
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
'-- Program ID        ACE1010C
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
Dim t_ord_no As String
Dim t_ord_item As String
Dim sdel As String

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public Active_CForm As String       'Form Active

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column1 Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column1 Collection
Dim iColumn1 As New Collection      'Spread Insert Column1 Collection
Dim aColumn1 As New Collection      'Master -> Spread Column1 Collection
Dim lColumn1 As New Collection      'Spread Lock Column1 Collection

Dim Mc1 As New Collection           'Master Collection
Dim ord_sc As New Collection           'order Spread Collection
Dim prod_sc As New Collection          'product spread collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column1

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
  'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
'    Call Gp_Ms_Collection(ord_txt_prod_cd, "p", "n ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_TxT_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'    Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(ord_prod_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(ord_prod_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(ord_prod_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(UDate_DEL_TO_b, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(ord_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(ord_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)


    
  '  Call Gp_Ms_Collection(prod_txt_prod_cd, "p", "n ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(prod_txt_stlgrd, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(prod_ord_no, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'    Call Gp_Ms_Collection(Text_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(prod_ord_itm, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(prod_prod_thk_fr, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(prod_prod_thk_to, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(prod_prod_wid_fr, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(prod_prod_wid_to, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(prod_prod_len_fr, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(prod_prod_len_to, "p", " ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         
         
   Call Gp_Ms_Collection(prod_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(prod_prod_wgt_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(prod_prod_wgt_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'    Call Gp_Ms_Collection(prod_combo, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
    
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Call Spread_Collection("Column1_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection1(ord_ss, 1, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 2, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(ord_ss, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 23, " ", " ", " ", "", " ", " l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ord_ss, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   Call Gp_Sp_Collection1(prod_ss, 1, " ", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 2, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection1(prod_ss, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(prod_ss, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(prod_ss, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(prod_ss, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(prod_ss, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(prod_ss, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(prod_ss, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection1(prod_ss, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection1(prod_ss, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection1(prod_ss, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection1(prod_ss, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection1(prod_ss, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'
   
   
    'Spread_Collection
    ord_sc.Add Item:=ord_ss, Key:="Spread"
    prod_sc.Add Item:=prod_ss, Key:="Spread"
    ord_sc.Add Item:="ACE1010C.P_MODIFY", Key:="P-M"
    ord_sc.Add Item:=pColumn1, Key:="pColumn"
    ord_sc.Add Item:=nColumn1, Key:="nColumn"
    ord_sc.Add Item:=aColumn1, Key:="aColumn"
    ord_sc.Add Item:=mColumn1, Key:="mColumn"
    ord_sc.Add Item:=iColumn1, Key:="iColumn"
    ord_sc.Add Item:=lColumn1, Key:="lColumn"
    ord_sc.Add Item:=1, Key:="First"
    ord_sc.Add Item:=ord_ss.MaxCols, Key:="ord_Last"
    


    prod_sc.Add Item:=pColumn1, Key:="pColumn"
    prod_sc.Add Item:=nColumn1, Key:="nColumn"
    prod_sc.Add Item:=aColumn1, Key:="aColumn"
    prod_sc.Add Item:=mColumn1, Key:="mColumn"
    prod_sc.Add Item:=iColumn1, Key:="iColumn"
    prod_sc.Add Item:=lColumn1, Key:="lColumn"
    prod_sc.Add Item:=1, Key:="First"
    prod_sc.Add Item:=prod_ss.MaxCols, Key:="prod_Last"




   ' Proc_Sc.Add Item:=sc1, Key:="Sc"
    Proc_Sc.Add Item:=ord_sc, Key:="oSc"
     Proc_Sc.Add Item:=prod_sc, Key:="pSc"
    

    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub cmd_confirm_Click()

'On Error GoTo Process_Exec_ERROR
'
'    Dim OutParam(1, 4) As Variant
'    Dim ret_Result_ErrMsg As String
'    Dim squery As String
'    Dim iCount As Integer
'
'    Dim adoCmd As ADODB.Command
'
'    'If ss1.MaxRows = 0 Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'    'Return Error Messsage Parameter
'    OutParam(1, 1) = "arg_e_msg"
'    OutParam(1, 2) = adVarChar
'    OutParam(1, 3) = adParamOutput
'    OutParam(1, 4) = 256
'
'    squery = "{call ACE1210P ('C1','','" + sUserID + "',?)}"
'
'    'Ado Setting
'    M_CN1.CursorLocation = adUseServer
'    Set adoCmd = New ADODB.Command
'
'    adoCmd.CommandType = adCmdText
'    Set adoCmd.ActiveConnection = M_CN1
'
'    adoCmd.CommandText = squery
'
'    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
'
'    adoCmd.Execute , , adExecuteNoRecords
'
'    'Process Error Check
'    If adoCmd("arg_e_msg") <> "" Then
'        ret_Result_ErrMsg = adoCmd("arg_e_msg")
'        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
'        Call Gp_MsgBoxDisplay(sErrMessg)
'    Else
'        Call Gp_MsgBoxDisplay("确定处理完了..!!", "I")
'    End If
'
'    Set adoCmd = Nothing
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'Process_Exec_ERROR:
'
'    Set adoCmd = Nothing
'    Screen.MousePointer = vbDefault
'    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub



Private Sub Command_REP_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim squery As String
    Dim iCount As Integer

    Dim adoCmd As ADODB.Command

Dim str_ord_prod_cd As String
Dim str_prod_prod_cd As String
Dim str_ord_no As String
Dim str_ord_item As String
Dim str_prod_loc As String
Dim str_prod_stlgrd As String
Dim str_prod_no As String

    Screen.MousePointer = vbHourglass

    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256


str_ord_prod_cd = ord_txt_prod_cd.Text
str_prod_prod_cd = prod_txt_prod_cd.Text
str_prod_stlgrd = prod_txt_stlgrd.Text
str_ord_no = prod_ord_no.Text
str_ord_item = prod_ord_itm.Text
str_prod_loc = prod_loc.Text
str_prod_no = prod_no.Text

If Len(str_prod_no) > 10 Then
       Call MsgBox("物料号输入错误" & Chr(10) & "请重新输入", vbExclamation + vbOKOnly, "警告")
       Screen.MousePointer = vbDefault
       Exit Sub
'ElseIf Len(str_prod_no) < 7 Then
'       str_prod_no = str_prod_no + "%"
End If
       

If str_ord_prod_cd <> "" Then

    If str_prod_prod_cd <> "" Then
    
        If str_ord_prod_cd = "PP" And str_prod_prod_cd = "PP" Then
        
            squery = "{call ACE1080P('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + str_prod_loc + "','" + str_prod_stlgrd + "',?)}"
        ElseIf str_ord_prod_cd = "HC" And str_prod_prod_cd = "HC" Then
            squery = "{call ACE1070P('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + str_prod_loc + "','" + str_prod_stlgrd + "',?)}"
        ElseIf str_ord_prod_cd = "SL" And str_prod_prod_cd = "SL" Then
            squery = "{call ACE1090P('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + str_prod_loc + "','" + str_prod_stlgrd + "',?)}"
        ElseIf (str_ord_prod_cd = "HC" Or str_ord_prod_cd = "PP") And str_prod_prod_cd = "SL" Then
            squery = "{call ACE1100P ('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + str_prod_loc + "','" + str_prod_stlgrd + "',?)}"
        Else
            Call MsgBox("订单与物料产品代码输入错误" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
  
                        '    'Ado Setting
            M_CN1.CursorLocation = adUseServer
            Set adoCmd = New ADODB.Command
        
            adoCmd.CommandType = adCmdText
            Set adoCmd.ActiveConnection = M_CN1
        
            adoCmd.CommandText = squery
        
             adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
 
            adoCmd.Execute , , adExecuteNoRecords
        
            'Process Error Check
            If adoCmd("arg_e_msg") <> "" Then
                ret_Result_ErrMsg = adoCmd("arg_e_msg")
                sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
                Call Gp_MsgBoxDisplay(sErrMessg)
            Else
                
                Call Gp_MsgBoxDisplay("替代处理完了..!!", "I")
                Call Form_Cls
                Call Form_Ref
            End If
        
            Set adoCmd = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
            
    Else
       Call MsgBox("产品分类代码不能为空！" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
       prod_txt_prod_cd.Text = ""
       prod_txt_prod_cd.SetFocus
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
Else
    Call MsgBox("产品分类代码不能为空！" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
    ord_txt_prod_cd.Text = ""
    ord_txt_prod_cd.SetFocus
       Screen.MousePointer = vbDefault
    Exit Sub
End If



Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
'On Error GoTo Process_Exec_ERROR
'
'    Dim OutParam(1, 4) As Variant
'    Dim ret_Result_ErrMsg As String
'    Dim squery As String
'    Dim iCount As Integer
'
'    Dim adoCmd As ADODB.Command
'
'    'If ss1.MaxRows = 0 Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'    'Return Error Messsage Parameter
'    OutParam(1, 1) = "arg_e_msg"
'    OutParam(1, 2) = adVarChar
'    OutParam(1, 3) = adParamOutput
'    OutParam(1, 4) = 256
'
'    squery = "{call ACE1070P (?)}"
'
'    'Ado Setting
'    M_CN1.CursorLocation = adUseServer
'    Set adoCmd = New ADODB.Command
'
'    adoCmd.CommandType = adCmdText
'    Set adoCmd.ActiveConnection = M_CN1
'
'    adoCmd.CommandText = squery
'
'    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
'
'    adoCmd.Execute , , adExecuteNoRecords
'
'    'Process Error Check
'    If adoCmd("arg_e_msg") <> "" Then
'        ret_Result_ErrMsg = adoCmd("arg_e_msg")
'        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
'        Call Gp_MsgBoxDisplay(sErrMessg)
'    Else
'        Call Gp_MsgBoxDisplay("替代处理完了..!!", "I")
'        Call Form_Ref
'    End If
'
'    Set adoCmd = Nothing
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'Process_Exec_ERROR:
'
'    Set adoCmd = Nothing
'    Screen.MousePointer = vbDefault
'    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub


Private Sub Form_Activate()

    Call FormMenuSetting1(Me, FormType, Toolbar_St, sAuthority)
    With MDIMain.MenuTool
        .Buttons(4).Enabled = True                 'Save
        .Buttons(9).Enabled = True                 'Delete
    End With
  

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
    'sAuthority = "0001"
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("oSc")("Spread"), False)
     Call Gp_Sp_Setting(Proc_Sc("pSc")("Spread"), False)
   ' Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call FormMenuSetting1(Me, FormType, "FS", sAuthority)
    Call Gf_Sp_Cls(Proc_Sc("oSc"))
    Call Gp_Sp_ColGet(Proc_Sc("oSc")("Spread"), "C-System.INI", Me.Name)
       Call Gf_Sp_Cls(Proc_Sc("pSc"))
    Call Gp_Sp_ColGet(Proc_Sc("pSc")("Spread"), "C-System.INI", Me.Name)
    
'     If Mid(sAuthority, 3, 1) <> "1" Then
''             Command_ALLSELECT.Enabled = False
'             Command_REP.Enabled = False
'             cmd_confirm.Enabled = False
'     End If

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("OSc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("PSc")("Spread"), "C-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set ord_sc = Nothing
    Set prod_sc = Nothing
    Set Proc_Sc = Nothing
    
    Call FormMenuSetting1(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("OSc")) And Gf_Sp_Cls(Proc_Sc("PSc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call FormMenuSetting1(Me, FormType, "CLS", sAuthority)
    End If
    
    UDate_DEL_TO_b.RawData = ""
    ord_TxT_STLGRD.Text = ""
   
    ord_prod_thk_fr = 0
    ord_prod_thk_to = 9999.99
    ord_prod_wid_fr.Text = 0
    ord_prod_wid_to.Text = 99999
    ord_prod_len_fr.Text = 0
    ord_prod_len_to.Text = 9999999
    'ord_txt_prod_cd.Text = ""
ord_TxT_STLGRD.Text = ""
'txt_cust_cd.Text = ""
UDate_DEL_TO_b.Text = ""

'prod_txt_prod_cd.Text = ""
prod_txt_stlgrd.Text = ""
prod_ord_no.Text = ""
prod_ord_itm.Text = ""
prod_loc.Text = ""
ord_slab_wid_min.Text = 0
ord_slab_wid_max.Text = 0
prod_no.Text = ""
ord_ord_no.Text = ""
ord_ord_item.Text = ""
    
    
      prod_prod_thk_fr.Text = 0
    prod_prod_thk_to.Text = 9999.99
    prod_prod_wid_fr.Text = 0
    prod_prod_wid_to.Text = 99999
    prod_prod_len_fr.Text = 0
    prod_prod_len_to.Text = 9999999
      prod_prod_wgt_fr.Text = 0
    prod_prod_wgt_to.Text = 9999999
  
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("ord_Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

Call ord_sel
Call prod_sel


End Sub
Public Sub Form_Pro()
  
End Sub

Public Sub Spread_Column1sSort()

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


Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("ord_Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0



End Sub


Private Sub ord_ss_Click(ByVal Col As Long, ByVal ROW As Long)
 Call Gp_Sp_Sort(Proc_Sc("osc")("Spread"), Col, ROW)
  lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ord_ss_DblClick(ByVal Col As Long, ByVal ROW As Long)
Dim sMesg As String
    If ord_ss.MaxRows < 1 Or ROW < 1 Then Exit Sub
  
        ord_ss.ROW = ROW
        
        ord_ss.Col = 1
        prod_ord_no.Text = ord_ss.Text
        
        ord_ss.Col = 2
        prod_ord_itm.Text = Trim(ord_ss.Value)
        
        
        ord_ss.Col = 14
        ord_slab_wid_min.Text = Trim(ord_ss.Value)
        
        ord_ss.Col = 15
        ord_slab_wid_max.Text = Trim(ord_ss.Value)

  

End Sub

Private Sub ord_ss_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
   If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("oSC")("spread"), Mode)
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If
End Sub

Private Sub ord_ss_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ord_ss_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ROW > 0 Then
        Set Active_Spread = Me.ord_ss
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub



Private Sub prod_ss_Click(ByVal Col As Long, ByVal ROW As Long)
 Call Gp_Sp_Sort(Proc_Sc("PSc")("Spread"), Col, ROW)
  lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub prod_ss_DblClick(ByVal Col As Long, ByVal ROW As Long)
     If prod_ss.MaxRows < 1 Or ROW < 1 Then Exit Sub
  
        prod_ss.ROW = ROW
        
        prod_ss.Col = 2
        prod_no.Text = prod_ss.Text
        
End Sub

Private Sub prod_ss_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
   If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("pSC")("spread"), Mode)
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If
End Sub

Private Sub prod_ss_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub prod_ss_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ROW > 0 Then
        Set Active_Spread = Me.prod_ss
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub ord_txt_prod_cd_Change()
  
    Select Case ord_txt_prod_cd.Text
           Case "S", "s", "SL"
               ord_txt_prod_cd.Text = "SL"
           Case "P", "p", "PP"
               ord_txt_prod_cd.Text = "PP"
           Case "H", "h", "HC"
               ord_txt_prod_cd.Text = "HC"
           Case "", "**"
               ord_txt_prod_cd.Text = ""
           Case Else
               ord_txt_prod_cd.Text = ""
               Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
     End Select
     
End Sub

Private Sub ord_txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then  '          '          '                 '                 ' ''''''''''
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=ord_txt_prod_cd
'        DD.rControl.Add Item:=Text_PROD_CD_mate
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

'    If Len(Trim(ord_txt_prod_cd.Text)) = ord_txt_prod_cd.MaxLength Then
'
'        Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", ord_txt_prod_cd.Text, 2)
'    Else
'        Text_PROD_CD_mate.Text = ""
'    End If
    
End Sub

Private Sub ord_TxT_STLGRD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
            
        DD.nameType = "1"
        DD.sWitch = "MS"
        DD.rControl.Add Item:=ord_TxT_STLGRD
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
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

Public Sub Gp_Sp_Collection1(sPname As Variant, Num As Integer, pcol As String, ncol As String, mcol As String, _
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
        Call Gp_Sp_ColColor(sPname, Num, , &HC0FFFF)
    End If
    
    If LCase(Trim(acol)) = "a" Then       'Master -> Spread Column
        aColumn.Add Item:=Num
        Call Gp_Sp_ColHidden(sPname, Num, True)
    End If
    
    If LCase(Trim(lCol)) = "l" Then       'Spread Lock Column
        lColumn.Add Item:=Num
        Call Gp_Sp_ColColor(sPname, Num, , &H80000005)
        Call Gp_Sp_ColLock(sPname, Num, True)
    End If

    
End Sub

Private Sub prod_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim squery As String
    
    If Len(Trim(prod_ord_no.Text)) = prod_ord_no.MaxLength Then
    
        If prod_ord_itm.Text <> "" Then Exit Sub
        
        prod_ord_no.Text = StrConv(prod_ord_no.Text, vbUpperCase)
        
        squery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(prod_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, prod_ord_itm, squery)
        
       ' If combo_ord_item.ListCount <> 0 Then
       '       combo_ord_item.ListIndex = 0
       ' End If
    Else
        prod_ord_itm.Clear
    End If

End Sub





Private Sub ord_ss_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

Dim i As Integer
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    Dim ROW1 As Long
    Dim row2 As Long
    Dim Col As Long
    Col = BlockCol
    ROW1 = BlockRow
    row2 = BlockRow2
    If Col = -1 Then
     For i = BlockRow To BlockRow2
        ord_ss.ROW = i
        ord_ss.Col = 0
        If ord_ss.Text = "Delete" Then
           ord_ss.Text = ""
            Call Gp_Sp_BlockColor(ord_ss, 1, ord_ss.MaxCols, ROW1, row2)
        Else: ord_ss.Text = "Delete"
         Call Gp_Sp_BlockColor(ord_ss, 1, ord_ss.MaxCols, ROW1, row2, , &HFFFF80)
        End If
        
     Next
   End If
     

End Sub
Private Sub prod_ss_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

Dim i As Integer
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    Dim ROW1 As Long
    Dim row2 As Long
    Dim Col As Long
    Col = BlockCol
    ROW1 = BlockRow
    row2 = BlockRow2
'    If Col = -1 Then
'     For i = BlockRow To BlockRow2
'        prod_ss.ROW = i
'        prod_ss.Col = 0
'        If prod_ss.Text = "Delete" Then
'           prod_ss.Text = ""
'            Call Gp_Sp_BlockColor(prod_ss, 1, prod_ss.MaxCols, ROW1, row2)
'        Else: prod_ss.Text = "Delete"
'         Call Gp_Sp_BlockColor(prod_ss, 1, prod_ss.MaxCols, ROW1, row2, , &HFFFF80)
'        End If
'
'     Next
'   End If
     

End Sub


'-------------------------------------------------------------------------------------
'--------------------------------------物料-------------------------------------------
'-------------------------------------------------------------------------------------

Private Sub prod_ord_itm_KeyPress(KeyAscii As Integer)
    'KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub prod_ord_itm_LostFocus()

    Dim S As String
  
    If Len(prod_ord_itm.Text) = 1 Then
        S = prod_ord_itm.Text
        prod_ord_itm.Text = "0" + S
    End If
    
End Sub


Private Sub prod_txt_prod_cd_Change()

    Select Case prod_txt_prod_cd.Text
           Case "S", "s", "SL"
               prod_txt_prod_cd.Text = "SL"
           Case "P", "p", "PP"
               prod_txt_prod_cd.Text = "PP"
           Case "H", "h", "HC"
               prod_txt_prod_cd.Text = "HC"
           Case ""
               prod_txt_prod_cd.Text = ""
           Case Else
               prod_txt_prod_cd.Text = ""
               Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select

End Sub

Private Sub prod_txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=prod_txt_prod_cd
'        DD.rControl.Add Item:=Text_PROD_CD_mate

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub

    End If
'
'    If Len(Trim(prod_txt_prod_cd.Text)) = prod_txt_prod_cd.MaxLength Then
'
'        Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", prod_txt_prod_cd.Text, 2)
'    Else
'        Text_PROD_CD_mate.Text = ""
'    End If


End Sub

Private Sub prod_Txt_PROD_CD_LostFocus()

    If prod_txt_prod_cd.Text <> "" Then
        If (Len(prod_txt_prod_cd.Text) < prod_txt_prod_cd.MaxLength) Then
            Call Gp_MsgBoxDisplay("产品分类不符合规范！")
            'Text_PROD_CD.Text = ""
            prod_txt_prod_cd.SetFocus
        End If
    End If

End Sub

Private Sub prod_txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
            
        DD.nameType = "1"
        DD.sWitch = "MS"
        DD.rControl.Add Item:=prod_txt_stlgrd
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    End If
        
End Sub

Private Function prod_sel() As Boolean


  Dim squery As String
    Dim sMesg As String
    Dim sProduct As String
    Dim S As String
      Dim minSIZEthk As Single     '--thick
    Dim maxSIZEthk As Single
    Dim minSIZEwid As Single     '--wide
    Dim maxSIZEwid As Single
    Dim minSIZElen As Single     '--lenth
    Dim maxSIZElen As Single
    Dim minSIZEwgt As Single     '--wight
    Dim maxSIZEwgt As Single
    Dim stlgrd     As String
    Dim str_prod_loc        As String
    Dim str_ord_no      As String
    Dim str_ord_item    As String
    Dim str_prod_no As String
    
      '----物料查询的条件：钢种，物料位置，产品重量，长，宽，厚，订单号
        minSIZEthk = prod_prod_thk_fr.Value
        maxSIZEthk = prod_prod_thk_to.Value
        minSIZEwid = prod_prod_wid_fr.Value
        maxSIZEwid = prod_prod_wid_to.Value
        minSIZElen = prod_prod_len_fr.Value
        maxSIZElen = prod_prod_len_to.Value
        minSIZEwgt = prod_prod_wgt_fr.Value
        maxSIZEwgt = prod_prod_wgt_to.Value
        stlgrd = prod_txt_stlgrd.Text
        str_prod_loc = prod_loc.Text
        str_ord_no = prod_ord_no.Text
        str_ord_item = prod_ord_itm.Text
        
     str_prod_no = prod_no.Text

If Len(str_prod_no) > 10 Then
       Call MsgBox("物料号输入错误" & Chr(10) & "请重新输入", vbExclamation + vbOKOnly, "警告")
       Screen.MousePointer = vbDefault
       Exit Function
'ElseIf Len(str_prod_no) < 7 Then
'       str_prod_no = str_prod_no + "%"
End If
        
        
        
If Len(str_prod_loc) > 7 Then
       Call MsgBox("物料位置输入错误！" & Chr(10) & "请重新输入。", vbExclamation + vbOKOnly, "警告")
       Exit Function
End If
If Len(str_prod_loc) > 0 Then
       If Left(str_prod_loc, 1) <> Left(prod_txt_prod_cd.Text, 1) Then
           Call MsgBox("产品代码与物料位置不符！" & Chr(10) & "请检查后输入。", vbExclamation + vbOKOnly, "警告")
           Exit Function
       End If
End If
    
    If prod_ord_itm.Text <> "" Then
        If Len(prod_ord_itm.Text) = 1 Then
            S = prod_ord_itm.Text
            prod_ord_itm.Text = "0" + S
        End If
    End If
    
    '---------------------根据产品的类型-------------------------------
  
    
    If prod_txt_prod_cd.Text <> "" Then
    
       If prod_txt_prod_cd.Text = "PP" And prod_ord_no.Text = "" Then
           
           squery = "Select  'PP',  A.PLATE_NO,A.LOC, "
           squery = squery + " A.STLGRD, A.PROD_THK, A.PROD_WID, A.PROD_LEN, A.WGT, A.PROD_OUTDIA, "
'           squery = squery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
           squery = squery + " A.TRIM_FL, A.UST_FL, TO_DATE(A.PROD_DATE,'YYYYMMDDHH24MISS'), Gf_ComnNameFind('C0008',A.WOO_RSN), "
           squery = squery + " A.CR_CD,  A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM "
           squery = squery + " From CP_REP_PLATE A, CP_REP_ORD B "
           squery = squery + "  Where NVL(A.PROD_CD,' ')  =    '" + Trim(prod_txt_prod_cd.Text) + "' "
           squery = squery + "    AND NVL(A.STLGRD,' ')   Like '" + Trim(prod_txt_stlgrd.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(prod_ord_no.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(prod_ord_itm.Text) + "%' "
           squery = squery + "    AND A.PROD_THK  BETWEEN " & minSIZEthk & " AND " & maxSIZEthk
           squery = squery + "    AND A.PROD_WID  BETWEEN " & minSIZEwid & " AND " & maxSIZEwid
           squery = squery + "    AND A.PROD_LEN  BETWEEN " & minSIZElen & " AND " & maxSIZElen
           squery = squery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
           squery = squery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
           squery = squery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt  '重量范围              'GENG
           squery = squery + "    AND NVL(A.LOC,' ')  LIKE '" + str_prod_loc + "%'"             '产品位置
           squery = squery + "    AND NVL(A.PLATE_NO,'') LIKE '" + str_prod_no + "%'"               '物料号
        
           squery = squery + "  ORDER BY A.LOC ASC "
            
       ' ElseIf prod_txt_prod_cd.Text = "HC" And prod_ord_no.Text = "" Then
        ElseIf prod_txt_prod_cd.Text = "HC" Then
           
           squery = " Select  'HC',   A.COIL_NO, A.LOC,"
           squery = squery + " A.STLGRD, A.PROD_THK, A.PROD_WID, A.PROD_LEN, A.WGT, A.PROD_OUTDIA, "
'           squery = squery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
           squery = squery + " A.TRIM_FL, A.UST_FL, TO_DATE(A.PROD_DATE,'YYYYMMDDHH24MISS'), Gf_ComnNameFind('C0008',A.WOO_RSN), "
           squery = squery + " A.CR_CD,  A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM "
           squery = squery + " From CP_REP_COIL A, CP_REP_ORD B  "
           squery = squery + "  Where NVL(A.PROD_CD,' ')  =    '" + Trim(prod_txt_prod_cd.Text) + "' "
           squery = squery + "    AND NVL(A.STLGRD,' ')   Like '" + Trim(prod_txt_stlgrd.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(prod_ord_no.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(prod_ord_itm.Text) + "%' "
           squery = squery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt  '重量范围              'GENG
           squery = squery + "    AND A.PROD_THK  BETWEEN " & minSIZEthk & " AND " & maxSIZEthk
           squery = squery + "    AND A.PROD_WID  BETWEEN " & minSIZEwid & " AND " & maxSIZEwid
           squery = squery + "    AND A.PROD_LEN  BETWEEN " & minSIZElen & " AND " & maxSIZElen
           squery = squery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
           squery = squery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
            squery = squery + "    AND NVL(A.LOC,' ')  LIKE '" + str_prod_loc + "%'"              '产品位置
            squery = squery + " AND NVL(A.COIL_NO,'') LIKE '" + str_prod_no + "%'"
           squery = squery + "  ORDER BY A.LOC ASC "
       
        Else
        ' ElseIf prod_txt_prod_cd.Text = "SL" And prod_ord_no.Text = "" Then
       
           squery = " Select  'SL',  A.SLAB_NO,A.LOC, "
           squery = squery + " A.STLGRD, A.THK, A.WID, A.LEN, A.WGT, A.PROD_OUTDIA, "
'           squery = squery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
          squery = squery + " A.TRIM_FL, A.UST_FL, TO_DATE(A.PROD_DATE,'YYYYMMDDHH24MISS'), Gf_ComnNameFind('C0008',A.WOO_RSN), "
           squery = squery + " A.CR_CD,  A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM "
           squery = squery + " From CP_REP_SLAB A, CP_REP_ORD B "
           squery = squery + "  Where NVL(A.STLGRD,' ')   Like '" + Trim(prod_txt_stlgrd.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(prod_ord_no.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(prod_ord_itm.Text) + "%' "
           squery = squery + "    AND A.PROD_THK  BETWEEN " & minSIZEthk & " AND " & maxSIZEthk
           squery = squery + "    AND A.PROD_WID  BETWEEN " & minSIZEwid & " AND " & maxSIZEwid
           squery = squery + "    AND A.PROD_LEN  BETWEEN " & minSIZElen & " AND " & maxSIZElen
           squery = squery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
           squery = squery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
           squery = squery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt                         'GENG 重量范围
           squery = squery + "    AND NVL(A.LOC,' ')  LIKE '" + str_prod_loc + "%'"                              '根据位置查询
            squery = squery + " AND NVL(A.SLAB_NO,'') LIKE '" + str_prod_no + "%'"
          If ord_slab_wid_min <> 0 And ord_slab_wid_max <> 0 Then
            squery = squery + "    AND A.WID BETWEEN " & ord_slab_wid_min.Value & " and " & ord_slab_wid_max.Value
          End If
           squery = squery + "  ORDER BY A.LOC ASC "
           
       End If
       ''''''''DANGZERONG    仅显示余材，以下替代的板坯查询取消
     
'     If prod_ord_no.Text <> "" Then
'
'          '''''' PP
'
'           squery = "Select 'PP',   A.PLATE_NO, "
'           squery = squery + " A.STLGRD, A.PROD_THK, A.PROD_WID, A.PROD_LEN, A.WGT, A.PROD_OUTDIA, "
'           squery = squery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
'           squery = squery + " A.TRIM_FL, A.UST_FL, TO_DATE(A.PROD_DATE,'YYYYMMDDHH24MISS'), Gf_ComnNameFind('C0008',A.WOO_RSN), "
'           squery = squery + " A.CR_CD, A.LOC, A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM "
'           squery = squery + " From CP_REP_PLATE A, CP_REP_ORD B "
'
'           squery = squery + "   WHERE NVL(A.STLGRD,' ')   Like '" + Trim(prod_txt_stlgrd.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(prod_ord_no.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(prod_ord_itm.Text) + "%' "
'           squery = squery + "    AND A.PROD_THK  BETWEEN " & minSIZEthk & " AND " & maxSIZEthk
'           squery = squery + "    AND A.PROD_WID  BETWEEN " & minSIZEwid & " AND " & maxSIZEwid
'           squery = squery + "    AND A.PROD_LEN  BETWEEN " & minSIZElen & " AND " & maxSIZElen
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
'           squery = squery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt                         'GENG 重量范围
'           squery = squery + "    AND NVL(A.LOC,' ')  LIKE '" + str_prod_loc + "%'"                              '根据位置查询
'         '  sQuery = sQuery + "  ORDER BY A.LOC ASC "
'           squery = squery + " UNION ALL "
'
'           '''' HC
'
'           squery = squery + " Select  'HC' , A.COIL_NO, "
'           squery = squery + " A.STLGRD, A.PROD_THK, A.PROD_WID, A.PROD_LEN, A.WGT, A.PROD_OUTDIA, "
'           squery = squery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
'           squery = squery + " A.TRIM_FL, A.UST_FL, TO_DATE(A.PROD_DATE,'YYYYMMDDHH24MISS'), Gf_ComnNameFind('C0008',A.WOO_RSN), "
'           squery = squery + " A.CR_CD, A.LOC, A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM "
'           squery = squery + " From CP_REP_COIL A, CP_REP_ORD B  "
'
'           squery = squery + "   WHERE NVL(A.STLGRD,' ')   Like '" + Trim(prod_txt_stlgrd.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(prod_ord_no.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(prod_ord_itm.Text) + "%' "
'           squery = squery + "    AND A.PROD_THK  BETWEEN " & minSIZEthk & " AND " & maxSIZEthk
'           squery = squery + "    AND A.PROD_WID  BETWEEN " & minSIZEwid & " AND " & maxSIZEwid
'           squery = squery + "    AND A.PROD_LEN  BETWEEN " & minSIZElen & " AND " & maxSIZElen
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
'           squery = squery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt                         'GENG 重量范围
'           squery = squery + "    AND NVL(A.LOC,' ')  LIKE '" + str_prod_loc + "%'"                              '根据位置查询
'         ' sQuery = sQuery + "  ORDER BY A.LOC ASC "
'           squery = squery + " UNION  ALL"
'
'           ''''' SL
'           squery = squery + " Select 'SL',   A.SLAB_NO, "
'           squery = squery + " A.STLGRD, A.THK, A.WID, A.LEN, A.WGT, A.PROD_OUTDIA, "
'           squery = squery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
'           squery = squery + " A.TRIM_FL, A.UST_FL, TO_DATE(A.PROD_DATE,'YYYYMMDDHH24MISS'),Gf_ComnNameFind('C0008',A.WOO_RSN), "
'           squery = squery + " A.CR_CD, A.LOC, A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM "
'           squery = squery + " From CP_REP_SLAB A, CP_REP_ORD B  "
'           squery = squery + "  Where NVL(A.STLGRD,' ')   Like '" + Trim(prod_txt_stlgrd.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(prod_ord_no.Text) + "%' "
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(prod_ord_itm.Text) + "%' "
'           squery = squery + "    AND A.PROD_THK  BETWEEN " & minSIZEthk & " AND " & maxSIZEthk
'           squery = squery + "    AND A.PROD_WID  BETWEEN " & minSIZEwid & " AND " & maxSIZEwid
'           squery = squery + "    AND A.PROD_LEN  BETWEEN " & minSIZElen & " AND " & maxSIZElen
'           squery = squery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
'           squery = squery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
'           squery = squery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt                         'GENG 重量范围
'           squery = squery + "    AND NVL(A.LOC,' ')  LIKE '" + str_prod_loc + "%'"                              '根据位置查询
'           squery = squery + "  ORDER BY LOC ASC "
'
'       End If

       
       sMesg = Gf_Ms_NeceCheck(nControl)
       If sMesg = "OK" Then

           sMesg = Gf_Ms_NeceCheck2(mControl)
           If sMesg = "OK" Then
                If Sp_Display(M_CN1, prod_sc.Item("Spread"), squery, prod_sc.Item("pColumn"), True) Then

'               If Gf_Only_Display(M_CN1, Proc_Sc("psc"), squery, , , False) Then
                prod_sel = True
                   prod_ss.OperationMode = OperationModeNormal
                   Call Gp_Ms_ControlLock(Mc1("lControl"), True)
                   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
               
           Else
               sMesg = sMesg + " Must input according to length of item"
               Call Gp_MsgBoxDisplay(sMesg)
           End If

       Else
           sMesg = sMesg + " Must input necessarily"
           Call Gp_MsgBoxDisplay(sMesg)
       End If
'
    Else
    
       Call MsgBox("产品分类代码不能为空！" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
       prod_txt_prod_cd.Text = ""
       prod_txt_prod_cd.SetFocus
       
    End If

End Function
Private Function ord_sel() As Boolean
 Dim squery As String
    Dim sMesg As String
    Dim maxDATE As String
    Dim minSIZEthk As Single
    Dim maxSIZEthk As Single
    Dim minSIZEwid As Single
    Dim maxSIZEwid As Single
    Dim minSIZElen As Single
    Dim maxSIZElen As Single


        minSIZEthk = ord_prod_thk_fr.Value
        maxSIZEthk = ord_prod_thk_to.Value
        minSIZEwid = ord_prod_wid_fr.Value
        maxSIZEwid = ord_prod_wid_to.Value
        minSIZElen = ord_prod_len_fr.Value
        maxSIZElen = ord_prod_len_to.Value

    If UDate_DEL_TO_b.RawData = "" Then
        maxDATE = "99991231"
    Else
        maxDATE = UDate_DEL_TO_b.RawData
    End If

    If maxSIZEthk >= minSIZEthk Then
        If maxSIZEwid >= minSIZEwid Then
            If maxSIZElen >= minSIZElen Then

                squery = "Select A.ORD_NO,A.ORD_ITEM,A.URGNT_FL,A.CUST_CD,TO_DATE(A.DEL_TO,'YYYY-MM-DD'),A.PROD_CD,A.STLGRD,A.ENDUSE_CD,"
                squery = squery + " A.PROD_THK,A.PROD_WID,A.PROD_LEN,A.PROD_OUTDIA ,A.SLAB_THK,A.SLAB_WID_MIN,A.SLAB_WID_MAX,A.ORD_WGT,A.DEL_TOL_MIN,A.DEL_TOL_MAX,"
                squery = squery + " A.CR_CD,A.UST_FL,A.ORD_TRIM_FL,A.ORD_HCR_FL,A.DESIGN_TOT_WGT,A.DESIGN_END_WGT,A.DESIGN_REM_WGT"
                squery = squery + "  From CP_REP_ORD A,BP_ORDER_ITEM B "

                squery = squery + " Where NVL(A.PROD_CD,' ')    Like '" + Trim(ord_txt_prod_cd.Text) + "%' "
                squery = squery + "   AND NVL(A.STLGRD,' ')     Like '" + Trim(ord_TxT_STLGRD.Text) + "%' "

               ' squery = squery + "   AND NVL(urgnt_fl,' ')   Like '" + Trim(txt_urgnt_fl.Text) + "%' "              订单紧急选项已不存在
              '  squery = squery + "   AND A.DEL_TO   <= " + maxDATE
                squery = squery + "   AND B.ORD_ACCP_DATE LIKE '" + Trim(UDate_DEL_TO_b.RawData) + "%' "
                squery = squery + "   AND A.PROD_THK <= " + Str$(maxSIZEthk)
                squery = squery + "   AND A.PROD_THK >= " + Str$(minSIZEthk)
                squery = squery + "   AND A.PROD_WID <= " + Str$(maxSIZEwid)
                squery = squery + "   AND A.PROD_WID >= " + Str$(minSIZEwid)
                squery = squery + "   AND A.PROD_LEN <= " + Str$(maxSIZElen)
                squery = squery + "   AND A.PROD_LEN >= " + Str$(minSIZElen)
                squery = squery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(ord_ord_no.Text) + "%' "
                squery = squery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(ord_ord_item.Text) + "%' "
                 squery = squery + "   AND A.ORD_NO = B.ORD_NO AND A.ORD_ITEM = B.ORD_ITEM"
                squery = squery + "   ORDER BY A.PROD_CD, A.STLGRD, A.PROD_THK, A.PROD_WID, A.PROD_LEN "

'                If sidbEdit_Slab_WID.Value <> 0 Then
'                    squery = squery + "   AND SLAB_WID = " + Str$(sidbEdit_Slab_WID.Value)
'                End If                       板坯宽度范围选择已删除

                sMesg = Gf_Ms_NeceCheck(nControl)
                If sMesg = "OK" Then

                    sMesg = Gf_Ms_NeceCheck2(mControl)
                    If sMesg = "OK" Then

                        If Sp_Display(M_CN1, ord_sc.Item("Spread"), squery, ord_sc.Item("pColumn"), True) Then
                            ord_sel = True
                            
                            Call FormMenuSetting1(Me, FormType, "RE", sAuthority)
                        End If

                    Else
                        sMesg = sMesg + " Must input according to length of item"
                        Call Gp_MsgBoxDisplay(sMesg)
                    End If

                Else
                    sMesg = sMesg + " Must input necessarily"
                    Call Gp_MsgBoxDisplay(sMesg)
                End If

            Else
                Call MsgBox("长度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
            End If
        Else
            Call MsgBox("宽度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        End If
    Else
        Call MsgBox("厚度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
    End If



End Function

Public Function Sp_Display(Conn As ADODB.Connection, sPname As vaSpread, squery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim ssname As String


    If sPname.Name = "ord_ss" Then
      ssname = "订单"
    End If
    If sPname.Name = "prod_ss" Then
      ssname = "物料"
    End If
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Display = True
        
        .ReDraw = False
        .MaxRows = 0: iCount = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open squery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("" + ssname + "无相关记录", "I")
                
            Sp_Display = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        .MaxRows = UBound(ArrayRecords, 2) + 1
    
        For iRowCount = 0 To .MaxRows - 1
        
            .ROW = iRowCount + 1
            
            For iColcount = 0 To .MaxCols - 1
            
                .Col = iColcount + 1
                
                Select Case .CellType
                
                    Case SS_CELL_TYPE_CHECKBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                            .Value = 0
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Value = ""
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case Else
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                End Select
                
            Next iColcount
            
        Next iRowCount
            
        If Not lColumn Is Nothing Then

            'lControl Lock
            For iCount = 1 To lColumn.Count

                .Protect = True
                .Col = lColumn(iCount): .Col2 = lColumn(iCount)
                .ROW = 1:               .row2 = .MaxRows
                .BlockMode = True: .Lock = True
                .BlockMode = False

            Next iCount

        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display = False
    Call Gp_MsgBoxDisplay("Sp_Display Error : " + ssname + squery)
    Screen.MousePointer = vbDefault

End Function


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub


Private Sub ord_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim squery As String
    
    If Len(Trim(ord_ord_no.Text)) = ord_ord_no.MaxLength Then
    
        If ord_ord_item.Text <> "" Then Exit Sub
        
        ord_ord_no.Text = StrConv(ord_ord_no.Text, vbUpperCase)
        
        squery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(ord_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, ord_ord_item, squery)
        
       ' If combo_ord_item.ListCount <> 0 Then
       '       combo_ord_item.ListIndex = 0
       ' End If
    Else
        ord_ord_item.Clear
    End If

End Sub
