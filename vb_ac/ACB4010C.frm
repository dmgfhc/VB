VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB4010C 
   Caption         =   "转库作业指示录入_ACB4010C"
   ClientHeight    =   8505
   ClientLeft      =   645
   ClientTop       =   1785
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_Trim_FL_S 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   7545
      MaxLength       =   1
      TabIndex        =   53
      Tag             =   "钢种"
      Top             =   30
      Width           =   450
   End
   Begin VB.TextBox txt_Trim_NAME_S 
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
      Height          =   310
      Left            =   8010
      TabIndex        =   52
      Tag             =   "钢种"
      Top             =   30
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2445
      Left            =   -15
      TabIndex        =   26
      Top             =   6750
      Width           =   15075
      Begin VB.TextBox txt_Trim_NAME 
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
         Height          =   310
         Left            =   5070
         TabIndex        =   51
         Tag             =   "钢种"
         Top             =   975
         Width           =   1530
      End
      Begin VB.TextBox txt_Trim_FL 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   4635
         MaxLength       =   1
         TabIndex        =   50
         Tag             =   "钢种"
         Top             =   975
         Width           =   420
      End
      Begin VB.TextBox Txt_STLGRD_SE 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8940
         MaxLength       =   20
         TabIndex        =   49
         Tag             =   "CD_MANA_NO"
         Top             =   210
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.ComboBox cbo_prod_grd 
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
         Left            =   7875
         TabIndex        =   48
         Top             =   975
         Width           =   1965
      End
      Begin VB.TextBox Text_CUST_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1455
         MaxLength       =   20
         TabIndex        =   47
         Tag             =   "CD_MANA_NO"
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox Text_INS_REMARK 
         Height          =   540
         Left            =   1455
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   1725
         Width           =   11160
      End
      Begin VB.TextBox text_cur_inv_code_T 
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
         Left            =   7860
         MaxLength       =   2
         TabIndex        =   44
         Top             =   1350
         Width           =   420
      End
      Begin VB.TextBox text_cur_inv_T 
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
         Left            =   8310
         TabIndex        =   43
         Top             =   1350
         Width           =   1530
      End
      Begin VB.TextBox TXT_ENDUSE_SE 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7860
         MaxLength       =   3
         TabIndex        =   42
         Tag             =   "CD_MANA_NO"
         Top             =   225
         Width           =   735
      End
      Begin VB.TextBox Txt_STLGRD_Detail 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4635
         MaxLength       =   20
         TabIndex        =   41
         Tag             =   "CD_MANA_NO"
         Top             =   225
         Width           =   1965
      End
      Begin VB.TextBox txt_prod_grd_IN 
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
         Left            =   9990
         MaxLength       =   1
         TabIndex        =   31
         Top             =   975
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox Text_size_knd_IN 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   30
         Tag             =   "钢种"
         Top             =   975
         Width           =   345
      End
      Begin VB.TextBox Text_size_knd_name_IN 
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
         Height          =   310
         Left            =   1800
         TabIndex        =   29
         Tag             =   "钢种"
         Top             =   975
         Width           =   1125
      End
      Begin VB.TextBox text_cur_inv_code_F 
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
         Left            =   4635
         MaxLength       =   2
         TabIndex        =   28
         Top             =   1350
         Width           =   420
      End
      Begin VB.TextBox text_cur_inv_F 
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
         Left            =   5070
         TabIndex        =   27
         Top             =   1350
         Width           =   1530
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   375
         Top             =   600
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "厚度"
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   3540
         Top             =   600
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "宽度"
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   6780
         Top             =   600
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "长度"
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
      Begin CSTextLibCtl.sidbEdit sidbEdit_size_Athk_SE 
         Height          =   315
         Left            =   1455
         TabIndex        =   32
         Top             =   600
         Width           =   885
         _Version        =   262145
         _ExtentX        =   1561
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
      Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bthk_SE 
         Height          =   315
         Left            =   2490
         TabIndex        =   33
         Top             =   600
         Width           =   885
         _Version        =   262145
         _ExtentX        =   1561
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
      Begin CSTextLibCtl.sidbEdit sidbEdit_size_Awid_SE 
         Height          =   315
         Left            =   4635
         TabIndex        =   34
         Top             =   600
         Width           =   885
         _Version        =   262145
         _ExtentX        =   1561
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
      Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bwid_SE 
         Height          =   315
         Left            =   5715
         TabIndex        =   35
         Top             =   600
         Width           =   885
         _Version        =   262145
         _ExtentX        =   1561
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
      Begin CSTextLibCtl.sidbEdit sidbEdit_size_Alen_SE 
         Height          =   315
         Left            =   7860
         TabIndex        =   36
         Top             =   600
         Width           =   885
         _Version        =   262145
         _ExtentX        =   1561
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
      Begin CSTextLibCtl.sidbEdit sidbEdit_size_Blen_SE 
         Height          =   315
         Left            =   8940
         TabIndex        =   37
         Top             =   600
         Width           =   885
         _Version        =   262145
         _ExtentX        =   1561
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   3540
         Top             =   1350
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "起始库"
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   375
         Top             =   975
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "定尺区分"
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   6780
         Top             =   975
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "等级"
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
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   3540
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "标准号"
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   6780
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "用途"
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   6780
         Top             =   1350
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "目标库"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   375
         Top             =   1350
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "转库重量"
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
      Begin CSTextLibCtl.sidbEdit sidbEdit_WGT_IN 
         Height          =   315
         Left            =   1455
         TabIndex        =   45
         Top             =   1350
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
      Begin InDate.ULabel ULabel21 
         Height          =   540
         Left            =   375
         Top             =   1725
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   953
         Caption         =   "说明"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   375
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "客户代码"
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   3540
         Top             =   975
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "切边"
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
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   180
         Left            =   8820
         TabIndex        =   40
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   180
         Left            =   5580
         TabIndex        =   39
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   180
         Left            =   2355
         TabIndex        =   38
         Top             =   690
         Width           =   90
      End
   End
   Begin VB.TextBox Text_PROD_CD_mate 
      Height          =   270
      Left            =   10740
      TabIndex        =   11
      Top             =   465
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox Text_PROD_CD 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1110
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "PP"
      Top             =   45
      Width           =   465
   End
   Begin VB.TextBox Text_PROC_CD_mate 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9945
      TabIndex        =   9
      Top             =   750
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text_REC_STS_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11430
      TabIndex        =   8
      Top             =   540
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text_STLGRD_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11550
      TabIndex        =   7
      Top             =   165
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text_STLGRD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1110
      MaxLength       =   20
      TabIndex        =   6
      Tag             =   "CD_MANA_NO"
      Top             =   390
      Width           =   1950
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   4770
      TabIndex        =   5
      Top             =   390
      Width           =   1515
   End
   Begin VB.TextBox text_cur_inv_code 
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
      Left            =   4350
      MaxLength       =   2
      TabIndex        =   4
      Top             =   390
      Width           =   390
   End
   Begin VB.TextBox Text_size_knd_name 
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
      Height          =   310
      Left            =   8010
      TabIndex        =   3
      Tag             =   "钢种"
      Top             =   390
      Width           =   1500
   End
   Begin VB.TextBox Text_size_knd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   7545
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "钢种"
      Top             =   390
      Width           =   450
   End
   Begin VB.TextBox txt_prod_grd_name 
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
      Height          =   310
      Left            =   4695
      TabIndex        =   1
      Tag             =   "钢种"
      Top             =   30
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txt_prod_grd 
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
      Left            =   4335
      MaxLength       =   1
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   345
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   5520
      Left            =   15
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1215
      Width           =   15150
      _Version        =   393216
      _ExtentX        =   26723
      _ExtentY        =   9737
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
      MaxRows         =   1
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACB4010C.frx":0000
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   30
      Top             =   45
      Width           =   1065
      _ExtentX        =   1879
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   30
      Top             =   390
      Width           =   1065
      _ExtentX        =   1879
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   12270
      Top             =   390
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "数量合计"
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
      Left            =   12285
      Top             =   735
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "重量合计"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   30
      Top             =   735
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3255
      Top             =   735
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "宽度"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   6465
      Top             =   735
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "长度"
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
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Athk 
      Height          =   315
      Left            =   1110
      TabIndex        =   13
      Top             =   735
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bthk 
      Height          =   315
      Left            =   2175
      TabIndex        =   14
      Top             =   735
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Awid 
      Height          =   315
      Left            =   4335
      TabIndex        =   15
      Top             =   735
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bwid 
      Height          =   315
      Left            =   5400
      TabIndex        =   16
      Top             =   735
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Alen 
      Height          =   315
      Left            =   7545
      TabIndex        =   17
      Top             =   735
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Blen 
      Height          =   315
      Left            =   8625
      TabIndex        =   18
      Top             =   735
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
   Begin CSTextLibCtl.sidbEdit Text_TOT_WGT 
      Height          =   315
      Left            =   13335
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   735
      Width           =   1485
      _Version        =   262145
      _ExtentX        =   2619
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Insert          =   0   'False
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit Text_TOT_SHEETS 
      Height          =   315
      Left            =   13320
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   390
      Width           =   1515
      _Version        =   262145
      _ExtentX        =   2672
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Insert          =   0   'False
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3255
      Top             =   390
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "堆放仓库"
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   6465
      Top             =   390
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "定尺区分"
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   3255
      Top             =   30
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "等级"
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
   Begin InDate.ULabel ULabel24 
      Height          =   315
      Left            =   6465
      Top             =   30
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "切边"
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
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   15
      X2              =   15150
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   30
      X2              =   15165
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14865
      TabIndex        =   25
      Top             =   495
      Width           =   195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "吨"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14850
      TabIndex        =   24
      Top             =   855
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   180
      Left            =   2040
      TabIndex        =   23
      Top             =   855
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   180
      Left            =   5265
      TabIndex        =   22
      Top             =   855
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   180
      Left            =   8505
      TabIndex        =   21
      Top             =   855
      Width           =   90
   End
End
Attribute VB_Name = "ACB4010C"
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
'-- Program ID        ACB4010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHENG WEN
'-- Coder             ZHENG WEN
'-- Date              2005.8.19
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim localize_iSumCnt As Integer
Dim localize_iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(Text_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 
        Call Gp_Ms_Collection(Text_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
   Call Gp_Ms_Collection(sidbEdit_size_Athk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sidbEdit_size_Bthk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sidbEdit_size_Awid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sidbEdit_size_Bwid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sidbEdit_size_Alen, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sidbEdit_size_Blen, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
         Call Gp_Ms_Collection(Text_CUST_NO, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Txt_STLGRD_SE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(TXT_ENDUSE_SE, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(sidbEdit_size_Athk_SE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(sidbEdit_size_Bthk_SE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(sidbEdit_size_Awid_SE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(sidbEdit_size_Bwid_SE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(sidbEdit_size_Alen_SE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(sidbEdit_size_Blen_SE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(Text_size_knd_IN, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(cbo_prod_grd, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_prod_grd_IN, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(sidbEdit_WGT_IN, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(text_cur_inv_code_F, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(text_cur_inv_F, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(text_cur_inv_code_T, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(text_cur_inv_T, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(Text_INS_REMARK, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
    Mc2.Add Item:=pControl1, Key:="pControl"
    Mc2.Add Item:=nControl1, Key:="nControl"
    Mc2.Add Item:=mControl1, Key:="mControl"
    Mc2.Add Item:=iControl1, Key:="iControl"
    Mc2.Add Item:=rControl1, Key:="rControl"
    Mc2.Add Item:=cControl1, Key:="cControl"
    Mc2.Add Item:=aControl1, Key:="aControl"
    Mc2.Add Item:=lControl1, Key:="lControl"
         
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Duplicate Count
    iDupCnt = 1
    'Sum Column Count
    localize_iSumCnt = 9
    
   ' Sum Column Setting
    localize_iSumCol.Add Item:=6
    localize_iSumCol.Add Item:=7
    localize_iSumCol.Add Item:=8
    localize_iSumCol.Add Item:=9
    localize_iSumCol.Add Item:=10
    localize_iSumCol.Add Item:=11
    localize_iSumCol.Add Item:=12
    localize_iSumCol.Add Item:=13
    localize_iSumCol.Add Item:=14
        
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Form_Button_Edit
    
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
    
'    cbo_prod_grd.AddItem "1:正品"
'    cbo_prod_grd.AddItem "2:订单外一级"
'    cbo_prod_grd.AddItem "3:订单外二级"
'    cbo_prod_grd.AddItem "5:等外品"
    Call AC_ComboAdd(M_CN1, cbo_prod_grd, "Q0034")
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
'    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    Call Form_Button_Edit

    Screen.MousePointer = vbDefault
    ss1.Row = -1
    ss1.Col = -1
    ss1.Lock = True
    Text_PROD_CD.Text = "PP"
    text_cur_inv_code_F.Text = "00"
    text_cur_inv_F.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code_F.Text, 2)
    txt_trim_fl.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set localize_iSumCol = Nothing
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Form_Button_Edit
        Txt_STLGRD_Detail.Text = ""
        txt_trim_fl.Text = ""
    End If
    
    text_tot_sheets.Value = 0
    text_tot_wgt.Value = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sQuery      As String
    Dim SMESG       As String
    Dim sTableName  As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    Dim TotalWeight As Single
    Dim TotalSheets As Single
    
    Dim minSIZEthk  As Single
    Dim maxSIZEthk  As Single
    Dim minSIZEwid  As Single
    Dim maxSIZEwid  As Single
    Dim minSIZElen  As Single
    Dim maxSIZElen  As Single
    
    TotalWeight = 0
    TotalSheets = 0
            
    If sidbEdit_size_Athk.Value = 0 Then
        minSIZEthk = 0
    Else
        minSIZEthk = sidbEdit_size_Athk.Value
    End If
    
    If sidbEdit_size_Bthk.Value = 0 Then
        maxSIZEthk = 9999.99
    Else
        maxSIZEthk = sidbEdit_size_Bthk.Value
    End If
      
    If sidbEdit_size_Awid.Value = 0 Then
        minSIZEwid = 0
    Else
        minSIZEwid = sidbEdit_size_Awid.Value
    End If
    
    If sidbEdit_size_Bwid.Value = 0 Then
        maxSIZEwid = 9999.99
    Else
        maxSIZEwid = sidbEdit_size_Bwid.Value
    End If
      
    If sidbEdit_size_Alen.Value = 0 Then
        minSIZElen = 0
    Else
        minSIZElen = sidbEdit_size_Alen.Value
    End If
     
    If sidbEdit_size_Blen.Value = 0 Then
        maxSIZElen = 9999999.9
    Else
        maxSIZElen = sidbEdit_size_Blen.Value
    End If
      
    Select Case Text_PROD_CD.Text
        Case "SL"
            sTableName = "FP_SLAB"
        Case "PP"
            sTableName = "GP_PLATE"
        Case "HC"
            sTableName = "GP_COIL"
        Case Else
            Call MsgBox("产品分类代码为空" & Chr(10) & "或不规范!请重试。", vbExclamation + vbOKOnly, "警告")
            Text_PROD_CD.Text = ""
            Text_PROD_CD.SetFocus
    End Select
  
    If maxSIZEthk < minSIZEthk Then
        Call MsgBox("长度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        Exit Sub
    ElseIf maxSIZEwid < minSIZEwid Then
        Call MsgBox("宽度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        Exit Sub
    ElseIf maxSIZElen < minSIZElen Then
        Call MsgBox("厚度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        Exit Sub
    End If
    
    If Text_PROD_CD.Text = "SL" Then
        sQuery = "Select Gf_Stlgrd_Detail(STLGRD),'',"
        sQuery = sQuery + "NVL(THK,0),NVL(TRUNC(WID/10)*10,0),NVL(TRUNC(LEN /100)* 100,0),COUNT(*),SUM(WGT),SUM(NVL(LOAD_WGT,0)),"
        sQuery = sQuery + "SUM(CASE WHEN CUR_INV ='00'  THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE WHEN CUR_INV<>'00'  THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD = '0' THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD = '1' THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD = '2' THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD IN ('3','4','5') THEN WGT ELSE 0 END), STLGRD "
    Else
        sQuery = "Select APLY_STDSPEC,APLY_ENDUSE_CD||' '||Gf_Enduse_Name(SUBSTR(PROD_CD,1,1),APLY_ENDUSE_CD),"
        sQuery = sQuery + "NVL(ORD_THK,0),NVL(ORD_WID,0),NVL(ORD_LEN,0),COUNT(*),SUM(WGT),SUM(NVL(LOAD_WGT,0)),"
        sQuery = sQuery + "SUM(CASE WHEN CUR_INV =  '00' THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE WHEN CUR_INV <> '00' THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD = '1'  THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD = '2'  THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD = '3'  THEN WGT ELSE 0 END),"
        sQuery = sQuery + "SUM(CASE when PROD_GRD = '5'  THEN WGT ELSE 0 END),  APLY_STDSPEC"
    End If
    
    sQuery = sQuery + "    From  " & sTableName
    sQuery = sQuery + "   Where  REC_STS = '2' "
    
    If Text_PROD_CD.Text = "SL" Then
        'sQuery = sQuery + "   AND PROC_CD   NOT IN('CAA', 'CAB') "
        sQuery = sQuery + "   AND PROC_CD       IN('XAA', 'XAC') "
        sQuery = sQuery + "   AND NVL(STLGRD,' ') Like '" + Trim(Text_STLGRD.Text) + "%' "
        sQuery = sQuery + "   AND NVL(THK,0) >= " + Str$(minSIZEthk)
        sQuery = sQuery + "   AND NVL(THK,0) <= " + Str$(maxSIZEthk)
        sQuery = sQuery + "   AND NVL(TRUNC(WID/10)*10,0) >= " + Str$(minSIZEwid)
        sQuery = sQuery + "   AND NVL(TRUNC(WID/10)*10,0) <= " + Str$(maxSIZEwid)
        sQuery = sQuery + "   AND NVL(TRUNC(LEN /100)* 100,0) >= " + Str$(minSIZElen)
        sQuery = sQuery + "   AND NVL(TRUNC(LEN /100)* 100,0) <= " + Str$(maxSIZElen)
    Else
        sQuery = sQuery + "   AND PROC_CD   LIKE 'X%' "
        sQuery = sQuery + "   AND NVL(APLY_STDSPEC,' ') Like '" + Trim(Text_STLGRD.Text) + "%' "
        sQuery = sQuery + "   AND NVL(ORD_THK,0) >= " + Str$(minSIZEthk)
        sQuery = sQuery + "   AND NVL(ORD_THK,0) <= " + Str$(maxSIZEthk)
        sQuery = sQuery + "   AND NVL(ORD_WID,0) >= " + Str$(minSIZEwid)
        sQuery = sQuery + "   AND NVL(ORD_WID,0) <= " + Str$(maxSIZEwid)
        sQuery = sQuery + "   AND NVL(ORD_LEN,0) >= " + Str$(minSIZElen)
        sQuery = sQuery + "   AND NVL(ORD_LEN,0) <= " + Str$(maxSIZElen)
    End If
        
    sQuery = sQuery + " AND NVL(CUR_INV,' ')  LIKE '" + Trim(text_cur_inv_code.Text) + "%'"
    sQuery = sQuery + " AND NVL(SIZE_KND,' ') LIKE '" + Trim(Text_size_knd.Text) + "%'"
    If Text_PROD_CD.Text = "PP" Then
        sQuery = sQuery + " AND NVL(TRIM_FL,'N')  LIKE '" + Trim(txt_Trim_FL_S.Text) + "%'"
    End If
    
    If Text_PROD_CD.Text = "SL" Then
        sQuery = sQuery + "   Group By STLGRD,THK, TRUNC(WID/10)*10, TRUNC(LEN /100)* 100 "
        sQuery = sQuery + "   Order By STLGRD,THK, TRUNC(WID/10)*10, TRUNC(LEN /100)* 100 "
    Else
        sQuery = sQuery + "   Group By APLY_STDSPEC,APLY_ENDUSE_CD,PROD_CD,ORD_THK,ORD_WID,ORD_LEN "
        sQuery = sQuery + "   Order By 1,APLY_ENDUSE_CD,PROD_CD,ORD_THK,ORD_WID,ORD_LEN "
    End If
    
    SMESG = Gf_Ms_NeceCheck(nControl)
    If SMESG = "OK" Then
    
        SMESG = Gf_Ms_NeceCheck2(mControl)
        If SMESG = "OK" Then

            If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, localize_iSumCnt, localize_iSumCol) Then
'            If Gf_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, localize_iSumCnt, localize_iSumCol) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            End If
    
        Else
            SMESG = SMESG + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(SMESG)
        End If
    
     Else
        SMESG = SMESG + " Must input necessarily"
        Call Gp_MsgBoxDisplay(SMESG)
     End If

     With ss1
         If .MaxRows = 0 Then
            text_tot_sheets.Text = "0"
            text_tot_wgt.Value = 0
         Else
            For iRow = 1 To .MaxRows - 1
                .Row = iRow
                .Col = 6: TotalSheets = .Value + TotalSheets
                .Col = 7: TotalWeight = .Value + TotalWeight
            Next iRow
            text_tot_sheets.Text = Str$(TotalSheets)
            text_tot_wgt.Text = Str$(TotalWeight)
         End If
     End With
     
    Call Form_Button_Edit
    Call Gp_Ms_Cls(Mc2("rControl"))
    Txt_STLGRD_Detail.Text = ""
    txt_trim_fl.Text = ""
    text_cur_inv_code_F.Text = "00"
    text_cur_inv_F.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code_F.Text, 2)

'
'    Text_CUST_NO = ""
'    Txt_STLGRD_SE = ""
'    TXT_ENDUSE_SE = ""
'    sidbEdit_size_Athk_SE.Value = 0
'    sidbEdit_size_Bthk_SE.Value = 0
'    sidbEdit_size_Awid_SE.Value = 0
'    sidbEdit_size_Bwid_SE.Value = 0
'    sidbEdit_size_Alen_SE.Value = 0
'    sidbEdit_size_Blen_SE.Value = 0
'    Text_size_knd_IN = ""
'    txt_prod_grd_IN = ""
'    sidbEdit_WGT_IN.Value = 0
'    text_cur_inv_code_F = ""
'    text_cur_inv_code_T = ""
'    Text_INS_REMARK = ""
    
End Sub

Public Sub Form_Pro()
On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim minSIZEthk As Single
    Dim maxSIZEthk As Single
    Dim minSIZEwid As Single
    Dim maxSIZEwid As Single
    Dim minSIZElen As Single
    Dim maxSIZElen As Single
   
    Dim adoCmd As ADODB.Command
    
    
    If sidbEdit_size_Athk_SE.Value = 0 Then
        minSIZEthk = 0
    Else
        minSIZEthk = sidbEdit_size_Athk_SE.Value
    End If
    
    If sidbEdit_size_Bthk_SE.Value = 0 Then
        maxSIZEthk = 9999.99
    Else
        maxSIZEthk = sidbEdit_size_Bthk_SE.Value
    End If
      
    If sidbEdit_size_Awid_SE.Value = 0 Then
        minSIZEwid = 0
    Else
        minSIZEwid = sidbEdit_size_Awid_SE.Value
    End If
    
    If sidbEdit_size_Bwid_SE.Value = 0 Then
        maxSIZEwid = 9999.99
    Else
        maxSIZEwid = sidbEdit_size_Bwid_SE.Value
    End If
      
    If sidbEdit_size_Alen_SE.Value = 0 Then
        minSIZElen = 0
    Else
        minSIZElen = sidbEdit_size_Alen_SE.Value
    End If
     
    If sidbEdit_size_Blen_SE.Value = 0 Then
        maxSIZElen = 9999999.9
    Else
        maxSIZElen = sidbEdit_size_Blen_SE.Value
    End If

    If sidbEdit_WGT_IN.Value = 0 Then
       Call Gp_MsgBoxDisplay("请输入移送重量！", "I")
       Exit Sub
    End If
    
    If text_cur_inv_code_F = "" Then
       Call Gp_MsgBoxDisplay("请输入原仓库代码！", "I")
       Exit Sub
    End If
    
    If text_cur_inv_code_T = "" Then
       Call Gp_MsgBoxDisplay("请输入移送仓库代码！", "I")
       Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
'    If Trim(Text_CUST_NO) = "" And Trim(Txt_STLGRD_Detail) = "" Then
'       Call Gp_MsgBoxDisplay("移送钢种或客户不能同时空", "I")
'       Exit Sub
'    End If
    If Text_PROD_CD = "SL" Then
       txt_trim_fl = ""
       sQuery = "{call ACB4010P ('" + sUserID + "','SL','" + Text_CUST_NO + "','','" + Txt_STLGRD_SE.Text + "','" + TXT_ENDUSE_SE + _
                "'," & minSIZEthk & "," & maxSIZEthk & "," & minSIZEwid & "," & maxSIZEwid & "," & _
                minSIZElen & "," & maxSIZElen & ",'" + Text_size_knd_IN + "','" + txt_trim_fl + "','" + txt_prod_grd_IN + "'," & _
                sidbEdit_WGT_IN.Value & ",'" + Trim(Text_INS_REMARK) + "','" + text_cur_inv_code_F + "','" + _
                text_cur_inv_code_T + "',?)}"
       
    Else
       sQuery = "{call ACB4010P ('" + sUserID + "','" + Text_PROD_CD + "','" + Text_CUST_NO + "','" + Txt_STLGRD_SE.Text + "','','" + _
                TXT_ENDUSE_SE + "'," & minSIZEthk & "," & maxSIZEthk & "," & minSIZEwid & "," & maxSIZEwid & "," & _
                minSIZElen & "," & maxSIZElen & ",'" + Text_size_knd_IN + "','" + txt_trim_fl + "','" + txt_prod_grd_IN + "'," & _
                sidbEdit_WGT_IN.Value & ",'" + Trim(Text_INS_REMARK) + "','" + text_cur_inv_code_F + "','" + _
                text_cur_inv_code_T + "',?)}"
    
    End If

    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1

    adoCmd.CommandText = sQuery

    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))

    adoCmd.Execute , , adExecuteNoRecords

'    Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("确定处理完了..!!", "I")
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Form_Button_Edit
        Call Gp_Ms_Cls(Mc2("rControl"))
        Txt_STLGRD_Detail.Text = ""

        Text_CUST_NO = ""
        Txt_STLGRD_SE = ""
        TXT_ENDUSE_SE = ""
        sidbEdit_size_Athk_SE.Value = 0
        sidbEdit_size_Bthk_SE.Value = 0
        sidbEdit_size_Awid_SE.Value = 0
        sidbEdit_size_Bwid_SE.Value = 0
        sidbEdit_size_Alen_SE.Value = 0
        sidbEdit_size_Blen_SE.Value = 0
        Text_size_knd_IN = ""
        txt_prod_grd_IN = ""
        sidbEdit_WGT_IN.Value = 0
        text_cur_inv_code_F = ""
        text_cur_inv_code_T = ""
        Text_INS_REMARK = ""
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
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

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim iRowCount As Long
  Dim MaxRow As Long
  Dim iRow As Integer
  Dim grd  As String

    If ss1.MaxRows < 1 Or Row = 0 Or ss1.MaxRows = Row Then Exit Sub
    With ss1
        .Row = .ActiveRow
'''        .Col = 1
'''        If .Text <> "" Then
'''            Txt_STLGRD_Detail = .Text
'''        Else
'''            For iRow = .ActiveRow To 1 Step -1
'''                .Row = iRow
'''                If .Text <> "" Then
'''                   Txt_STLGRD_Detail = .Text
'''                   Exit For
'''                End If
'''            Next iRow
'''            .Row = .ActiveRow
'''        End If
'''
'''        .Col = .MaxCols
'''        If .Text <> "" Then
'''            Txt_STLGRD_SE = .Text
'''        Else
'''            For iRow = .ActiveRow To 1 Step -1
'''                .Row = iRow
'''                If .Text <> "" Then
'''                   Txt_STLGRD_SE = .Text
'''                   Exit For
'''                End If
'''            Next iRow
'''            .Row = .ActiveRow
'''        End If
'''
'''        .Col = 2
'''        TXT_ENDUSE_SE = Left(.Text, 3)
        .Col = 3
        sidbEdit_size_Athk_SE.Value = .Text
        sidbEdit_size_Bthk_SE.Value = .Text
        .Col = 4
        sidbEdit_size_Awid_SE.Value = .Text
        sidbEdit_size_Bwid_SE.Value = .Value + 9
        .Col = 5
        sidbEdit_size_Alen_SE.Value = .Text
        sidbEdit_size_Blen_SE.Value = .Value + 99
        .Col = 8
        sidbEdit_WGT_IN.Value = .Text
    End With

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub text_cur_inv_code_Change()
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
          Exit Sub
    Else
          text_cur_inv.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_F_DblClick()

    Call text_cur_inv_code_F_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub text_cur_inv_code_F_Change()
    If Len(Trim(text_cur_inv_code_F.Text)) = text_cur_inv_code_F.MaxLength Then
        text_cur_inv_F.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code_F.Text, 2)
        Exit Sub
    Else
        text_cur_inv_F.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_F_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code_F
        DD.rControl.Add Item:=text_cur_inv_F
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub text_cur_inv_code_T_Change()
       
    If Len(Trim(text_cur_inv_code_T.Text)) = text_cur_inv_code_T.MaxLength Then
        text_cur_inv_T.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code_T.Text, 2)
        Exit Sub
    Else
        text_cur_inv_T.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_T_DblClick()

    Call text_cur_inv_code_T_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_T_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code_T
        DD.rControl.Add Item:=text_cur_inv_T
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If
End Sub

Private Sub Text_CUST_NO_DblClick()

    Call Text_CUST_NO_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_CUST_NO_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=Text_CUST_NO
        DD.nameType = "2"

        Call Gf_Customer_DD(M_CN1, KeyCode)

       ' Exit Sub

    End If

End Sub

Private Sub Text_PROD_CD_Change()

    ULabel2.Caption = "钢种"
    Select Case Text_PROD_CD.Text
'        Case "S", "s", "SL"
'            Text_PROD_CD.Text = "SL"
'            ULabel2.Caption = "钢种"
'            ULabel17.Caption = "钢种"
        Case "P", "p", "PP"
            Text_PROD_CD.Text = "PP"
            ULabel2.Caption = "标准号"
            ULabel17.Caption = "标准号"
        Case "H", "h", "HC"
            Text_PROD_CD.Text = "HC"
            ULabel2.Caption = "标准号"
            ULabel17.Caption = "标准号"
        Case ""
            Text_PROD_CD.Text = ""
        Case Else
            Text_PROD_CD.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范，只能为“PP”或“HC”！请更正。", vbExclamation + vbOKOnly, "警告")
    End Select
  
    cbo_prod_grd.Clear
    Select Case Text_PROD_CD.Text
        Case "S", "s", "SL"
            cbo_prod_grd.AddItem "0:合格"
            cbo_prod_grd.AddItem "1:表面不合格"
            cbo_prod_grd.AddItem "2:内部缺陷"
            cbo_prod_grd.AddItem "3:内外缺陷"
            cbo_prod_grd.AddItem "4:操作员变更"
            cbo_prod_grd.AddItem "5:长度不合格"
            
            ss1.Row = 0
            ss1.Col = 10: ss1.Text = "合格品重量"
            ss1.Col = 11: ss1.Text = "表面不合格品重量"
            ss1.Col = 12: ss1.Text = "内部缺陷品重量"
            
            Call Gp_Sp_ColHidden(ss1, 2, True)
        Case Else
'            cbo_prod_grd.AddItem "1:正品"
'            cbo_prod_grd.AddItem "2:订单外一级"
'            cbo_prod_grd.AddItem "3:订单外二级"
'            cbo_prod_grd.AddItem "5:等外品"
            Call AC_ComboAdd(M_CN1, cbo_prod_grd, "Q0034")
            ss1.Row = 0
            ss1.Col = 10: ss1.Text = "正品重量"
            ss1.Col = 11: ss1.Text = "订单外一级(改判)重量"
            ss1.Col = 12: ss1.Text = "订单外二级(协议品)重量"
            Call Gp_Sp_ColHidden(ss1, 2, False)
    End Select
    
'    Call Form_Ref
      
End Sub

Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=Text_PROD_CD
        DD.rControl.Add Item:=Text_PROD_CD_mate

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(Text_PROD_CD.Text)) = Text_PROD_CD.MaxLength Then
        Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_PROD_CD.Text, 2)
    Else
        Text_PROD_CD_mate.Text = ""
    End If
    
End Sub

Private Sub Text_size_knd_DblClick()

    Call Text_size_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_size_knd_IN_DblClick()

    Call Text_size_knd_IN_KeyUp(vbKeyF4, 0)
End Sub

Private Sub text_stlgrd_DblClick()

    Call text_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
       
        If Text_PROD_CD.Text = "SL" Then
            DD.sWitch = "MS"
            DD.rControl.Add Item:=Text_STLGRD
            
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            DD.sWitch = "MS"
            DD.rControl.Add Item:=Text_STLGRD
    
            Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        End If
        
    End If

End Sub

Private Sub TXT_ENDUSE_SE_DblClick()

    Call TXT_ENDUSE_SE_KeyDown(vbKeyF4, 0)
    
End Sub

Private Sub TXT_ENDUSE_SE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        If Text_PROD_CD.Text = "SL" Then
            DD.sKey = "S"
        Else
            DD.sKey = "P"
        End If
        
        DD.rControl.Add Item:=TXT_ENDUSE_SE
        
        Call Gf_Usage_DD(M_CN1, KeyCode)
    End If
    
End Sub

'Private Sub Text_STLGRD_SE_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        If Text_PROD_CD.Text = "SL" Then
'            DD.sWitch = "MS"
'            DD.rControl.Add Item:=Text_STLGRD_SE
'
'            DD.nameType = "1"
'            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
'        Else
'            DD.sWitch = "MS"
'            DD.rControl.Add Item:=Text_STLGRD_SE
'
'            Call Gf_StdSPEC_DD(M_CN1, KeyCode)
'        End If
'
'    End If
'
'End Sub

Private Sub txt_prod_grd_Change()
    If Len(Trim(txt_PROD_GRD.Text)) = txt_PROD_GRD.MaxLength Then
        txt_prod_grd_name.Text = Gf_ComnNameFind(M_CN1, "Q0034", txt_PROD_GRD.Text, 1)
        Exit Sub
    Else
        txt_prod_grd_name.Text = ""
    End If
End Sub

Private Sub txt_prod_grd_DblClick()

    Call txt_prod_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_grd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0034"

        DD.rControl.Add Item:=txt_PROD_GRD

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub cbo_prod_grd_Click()
    If Trim(cbo_prod_grd.Text) <> "" Then
        txt_prod_grd_IN.Text = Left(cbo_prod_grd.Text, 1)
    Else
        txt_prod_grd_IN.Text = ""
    End If
End Sub

Private Sub cbo_prod_grd_Change()
    If Trim(cbo_prod_grd.Text) <> "" Then
        txt_prod_grd_IN.Text = Left(cbo_prod_grd.Text, 1)
    Else
        txt_prod_grd_IN.Text = ""
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

Private Sub Text_size_knd_IN_Change()
    If Len(Trim(Text_size_knd_IN.Text)) = Text_size_knd_IN.MaxLength Then
        text_size_knd_name_in.Text = Gf_ComnNameFind(M_CN1, "B0043", Text_size_knd_IN.Text, 2)
        Exit Sub
    Else
        text_size_knd_name_in.Text = ""
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

Private Sub Text_size_knd_IN_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=Text_size_knd_IN

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub Form_Button_Edit()
    MDIMain.MenuTool.Buttons(7).Enabled = False              'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False              'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False              'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False             'Row Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False             'Row Paste
End Sub

Private Sub Txt_STLGRD_Detail_DblClick()

    Call Txt_STLGRD_Detail_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_Change()
    If Len(Trim(txt_trim_fl.Text)) = txt_trim_fl.MaxLength Then
        txt_trim_name.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_trim_fl.Text, 2)
        txt_trim_fl.Text = Trim(txt_trim_fl.Text)
        Exit Sub
    Else
        txt_trim_name.Text = ""
        txt_trim_fl.Text = ""
    End If

End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=txt_trim_fl

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub txt_TRIM_FL_S_Change()
    If Len(Trim(txt_Trim_FL_S.Text)) = txt_Trim_FL_S.MaxLength Then
        txt_Trim_NAME_S.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_Trim_FL_S.Text, 2)
        txt_Trim_FL_S.Text = Trim(txt_Trim_FL_S.Text)
        Exit Sub
    Else
        txt_Trim_NAME_S.Text = ""
        txt_Trim_FL_S.Text = ""
    End If

End Sub

Private Sub txt_Trim_FL_S_DblClick()

    Call txt_Trim_FL_S_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_Trim_FL_S_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=txt_Trim_FL_S

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub Txt_STLGRD_Detail_Change()
    Txt_STLGRD_SE.Text = Txt_STLGRD_Detail.Text
End Sub

Private Sub Txt_STLGRD_Detail_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
   
        If Text_PROD_CD.Text = "SL" Then
            DD.sWitch = "MS"
            DD.rControl.Add Item:=Txt_STLGRD_Detail
            
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            DD.sWitch = "MS"
            DD.rControl.Add Item:=Txt_STLGRD_Detail
    
            Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        End If
    End If
End Sub

Private Function AC_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPrc As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String
    Dim intCount As Integer
    intCount = 1
    Dim AdoRs As ADODB.Recordset
    
    If Trim(sPrc) = "" Then
        AC_ComboAdd = False: Exit Function
    End If
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then AC_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT CD_NAME FROM ZP_CD Where CD_MANA_NO = '" + Trim(sPrc) + "'"

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                If intCount = 6 Then intCount = 7
                Cbo.AddItem Trim(Str(intCount)) + ":" + AdoRs.Fields(0)
                intCount = intCount + 1
            End If
            AdoRs.MoveNext
            
        Wend
        AC_ComboAdd = True
    Else
        AC_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    AC_ComboAdd = False

End Function

