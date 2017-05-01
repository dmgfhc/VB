VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form CEC_Slab_Request 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "申请紧急坯"
   ClientHeight    =   4155
   ClientLeft      =   3930
   ClientTop       =   4170
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   10350
   Begin Threed.SSPanel pnl_first 
      Height          =   1905
      Left            =   30
      TabIndex        =   21
      Top             =   30
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   3360
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cbo_prod_cnt 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8025
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "轧件内产品数"
         Top             =   585
         Width           =   780
      End
      Begin VB.TextBox txt_hcr_fl 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   9645
         MaxLength       =   1
         TabIndex        =   9
         Tag             =   "H/C"
         Top             =   150
         Width           =   405
      End
      Begin VB.TextBox txt_ord_item 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   2910
         MaxLength       =   11
         TabIndex        =   2
         Tag             =   "订单号"
         Top             =   155
         Width           =   405
      End
      Begin VB.TextBox txt_ord_no 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1605
         MaxLength       =   11
         TabIndex        =   1
         Tag             =   "订单号"
         Top             =   155
         Width           =   1305
      End
      Begin VB.TextBox txt_stlgrd_name 
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
         Left            =   6210
         TabIndex        =   4
         Tag             =   "钢种"
         Top             =   155
         Width           =   1755
      End
      Begin VB.TextBox txt_stlgrd 
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
         Left            =   4935
         MaxLength       =   11
         TabIndex        =   3
         Tag             =   "钢种"
         Top             =   155
         Width           =   1275
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Index           =   0
         Left            =   270
         Top             =   1005
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "板坯厚度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         Left            =   3600
         Top             =   1005
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "板坯宽度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_thk 
         Height          =   315
         Left            =   1605
         TabIndex        =   5
         Tag             =   "板坯厚度"
         Top             =   1005
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
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
         FocusSelect     =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit sdb_slab_wid 
         Height          =   315
         Left            =   4935
         TabIndex        =   6
         Tag             =   "板坯宽度"
         Top             =   1005
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
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
         FocusSelect     =   -1  'True
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   3600
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "钢种"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         Left            =   270
         Top             =   1440
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "板坯重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_wgt 
         Height          =   315
         Left            =   1605
         TabIndex        =   8
         Tag             =   "板坯重量"
         Top             =   1440
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   -1  'True
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
         NumIntDigits    =   12
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   270
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "订单号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Index           =   1
         Left            =   270
         Top             =   585
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "订单欠重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord_rem_wgt 
         Height          =   315
         Left            =   1605
         TabIndex        =   10
         Top             =   585
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
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
         Enabled         =   0   'False
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
         NumIntDigits    =   12
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Index           =   2
         Left            =   3600
         Top             =   585
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "所需产品数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_rem_cnt 
         Height          =   315
         Left            =   4935
         TabIndex        =   11
         Top             =   585
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
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
         Enabled         =   0   'False
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
         NumIntDigits    =   10
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   6690
         Top             =   1020
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "板坯长度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_len 
         Height          =   315
         Left            =   8025
         TabIndex        =   7
         Tag             =   "板坯长度"
         Top             =   1020
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   8310
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "H/C"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Index           =   5
         Left            =   6690
         Top             =   585
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "轧件内产品数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Index           =   6
         Left            =   3600
         Top             =   1440
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "所需板坯数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_rem_cnt 
         Height          =   315
         Left            =   4935
         TabIndex        =   25
         Top             =   1440
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
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
         Enabled         =   0   'False
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
         NumIntDigits    =   10
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   180
         TabIndex        =   22
         Top             =   135
         Width           =   105
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   465
      Left            =   4050
      TabIndex        =   20
      Top             =   3585
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "申请確認"
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   465
      Left            =   5445
      TabIndex        =   23
      Top             =   3585
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1515
      Left            =   30
      TabIndex        =   24
      Top             =   1950
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   2672
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_plt_name 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   2070
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "申请炼钢厂"
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox txt_plt 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   15
         Tag             =   "申请炼钢厂"
         Top             =   660
         Width           =   465
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   270
         Top             =   660
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "申请炼钢厂"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         Index           =   3
         Left            =   270
         Top             =   1080
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "切割数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_unit_cnt 
         Height          =   315
         Left            =   1605
         TabIndex        =   17
         Tag             =   "切割数"
         Top             =   1080
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   12583104
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
         FocusSelect     =   -1  'True
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
         NumIntDigits    =   10
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Index           =   4
         Left            =   3600
         Top             =   1080
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "申请板坯数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_req_cnt 
         Height          =   315
         Left            =   4935
         TabIndex        =   18
         Tag             =   "申请板坯数"
         Top             =   1080
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   12583104
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
         FocusSelect     =   -1  'True
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
         NumIntDigits    =   10
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   3600
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "申请板坯重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_long_wgt 
         Height          =   315
         Left            =   4935
         TabIndex        =   13
         Tag             =   "申请板坯重量"
         Top             =   150
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   255
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
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
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
         NumIntDigits    =   12
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   270
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "申请板坯长度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_long_len 
         Height          =   315
         Left            =   1605
         TabIndex        =   12
         Tag             =   "申请板坯长度"
         Top             =   150
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   255
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
         ReadOnly        =   -1  'True
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
      Begin InDate.UDate udt_req_date 
         Height          =   315
         Left            =   4935
         TabIndex        =   16
         Tag             =   "计划使用日期"
         Top             =   660
         Width           =   1500
         _ExtentX        =   2646
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   3600
         Top             =   660
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "计划使用日期"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   6690
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "订单欠余重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_rem_wgt 
         Height          =   315
         Left            =   8025
         TabIndex        =   14
         Tag             =   "申请板坯重量"
         Top             =   150
         Width           =   1110
         _Version        =   262145
         _ExtentX        =   1958
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   255
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
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
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
         NumIntDigits    =   12
         Undo            =   0
         Data            =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   10110
         Y1              =   570
         Y2              =   570
      End
   End
End
Attribute VB_Name = "CEC_Slab_Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name
'-- Program Name      SLAB URGENT REQUEST
'-- Program ID        Slab_Request
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2007.11.01
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public Ord_Fl As String
Public lProd_Wgt As Double

Private Sub cbo_prod_cnt_Click()
    
    sdb_slab_wgt.Value = Gf_FloatFind(M_CN1, "SELECT DISTINCT SLAB_WGT FROM NISCO.CP_ORD_SL_D WHERE ORD_NO = '" & CEC_Slab_Request.txt_ord_no.Text & "' AND ORD_ITEM = '" & CEC_Slab_Request.txt_ord_item.Text & "' AND ASROLL_PROD_CNT = " & cbo_prod_cnt.Text & " AND REP_STD_FL = 'Y'")
    sdb_slab_len.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & ",0," & sdb_slab_wgt.Value & ") FROM DUAL")
    sdb_slab_rem_cnt.Value = Gf_FloatFind(M_CN1, "SELECT TRUNC(" & sdb_rem_cnt.Value & " / " & Val(cbo_prod_cnt.Text) & ") FROM DUAL ")
    sdb_long_len.Value = 0
    sdb_long_wgt.Value = 0
    sdb_slab_rem_wgt.Value = sdb_ord_rem_wgt.Value
    sdb_unit_cnt.Value = 0
    sdb_req_cnt.Value = 0
End Sub

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()
    
    If Ord_Fl = "1" Then
        If txt_ord_no.Text = "" Or txt_ord_item.Text = "" Then
            Call Gp_MsgBoxDisplay(txt_ord_no.Tag + "必须输入", "I")
            Exit Sub
        End If
        
        If cbo_prod_cnt.Text = "" Then
            Call Gp_MsgBoxDisplay(cbo_prod_cnt.Tag + "必须输入", "I")
            Exit Sub
        End If
    End If
    
    If txt_stlgrd.Text = "" Or txt_stlgrd_name.Text = "" Then
        Call Gp_MsgBoxDisplay(txt_stlgrd.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If sdb_slab_thk.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_slab_thk.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If sdb_slab_wid.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_slab_wid.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If sdb_slab_len.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_slab_len.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If sdb_slab_wgt.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_slab_wgt.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If txt_hcr_fl.Text = "" Or (UCase(txt_hcr_fl.Text) <> "C" And UCase(txt_hcr_fl.Text) <> "H") Then
        Call Gp_MsgBoxDisplay(txt_hcr_fl.Tag + "必须输入, 'H','C' ", "I")
        Exit Sub
    End If
    
    If txt_plt.Text = "" Or txt_plt_name.Text = "" Or (UCase(txt_plt.Text) <> "B1" And UCase(txt_plt.Text) <> "B3") Then
        Call Gp_MsgBoxDisplay(txt_plt.Tag + "必须输入, 'B1','B3' ", "I")
        Exit Sub
    End If
    
    If udt_req_date.RawData = "" Or Len(udt_req_date.RawData) <> 8 Then
        Call Gp_MsgBoxDisplay(udt_req_date.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If udt_req_date.RawData < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL") Then
        Call Gp_MsgBoxDisplay(udt_req_date.Tag + " < 现在日期", "I")
        Exit Sub
    End If
    
    If sdb_unit_cnt.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_unit_cnt.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If sdb_req_cnt.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_req_cnt.Tag + "必须输入", "I")
        Exit Sub
    End If
    
    If Ord_Fl = "1" Then
        If sdb_rem_cnt.Value < sdb_unit_cnt.Value * sdb_req_cnt.Value Then
            If Not Gf_MessConfirm("所需板坯数 < 申请板坯数", "I") Then
                Exit Sub
            End If
        End If
    End If
    
    Call Gp_Process_Exec
    
End Sub

Private Sub Form_Activate()

    Call txt_stlgrd_KeyUp(0, 0)
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    If Ord_Fl = "2" Then
        txt_stlgrd.Enabled = True
        txt_stlgrd_name.Enabled = True
        txt_hcr_fl.Enabled = True
        sdb_slab_thk.Enabled = True
        sdb_slab_wid.Enabled = True
        sdb_slab_len.Enabled = True
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Call Gp_FormCenter(Me)
    Me.BackColor = &HE0E0E0

End Sub

Private Sub sdb_req_cnt_KeyUp(KeyCode As Integer, Shift As Integer)

    If Ord_Fl = "2" Then Exit Sub
    If sdb_long_len.Value = 0 Then Exit Sub
    
    sdb_slab_rem_wgt.Value = sdb_ord_rem_wgt - (lProd_Wgt * Val(cbo_prod_cnt.Text) * sdb_unit_cnt.Value * sdb_req_cnt.Value)
    
End Sub

Private Sub sdb_slab_len_KeyUp(KeyCode As Integer, Shift As Integer)

    sdb_slab_wgt.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & "," & sdb_slab_len.Value & ",0) FROM DUAL")
    
End Sub

Private Sub sdb_slab_thk_KeyUp(KeyCode As Integer, Shift As Integer)

    If Ord_Fl = "1" Then
        sdb_slab_len.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & ",0," & sdb_slab_wgt.Value & ") FROM DUAL")
    Else
        sdb_slab_wgt.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & "," & sdb_slab_len.Value & ",0) FROM DUAL")
    End If
    
    sdb_long_len.Value = 0
    sdb_long_wgt.Value = 0
    sdb_slab_rem_wgt.Value = sdb_ord_rem_wgt.Value
    sdb_unit_cnt.Value = 0
    sdb_req_cnt.Value = 0
    
End Sub

Private Sub sdb_slab_wgt_KeyUp(KeyCode As Integer, Shift As Integer)

    sdb_long_wgt.Value = sdb_unit_cnt.Value * sdb_slab_wgt.Value
    
End Sub

Private Sub sdb_slab_wid_KeyUp(KeyCode As Integer, Shift As Integer)

    If Ord_Fl = "1" Then
        sdb_slab_len.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & ",0," & sdb_slab_wgt.Value & ") FROM DUAL")
    Else
        sdb_slab_wgt.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & "," & sdb_slab_len.Value & ",0) FROM DUAL")
    End If
    
    sdb_long_len.Value = 0
    sdb_long_wgt.Value = 0
    sdb_slab_rem_wgt.Value = sdb_ord_rem_wgt.Value
    sdb_unit_cnt.Value = 0
    sdb_req_cnt.Value = 0
    
End Sub

Private Sub sdb_unit_cnt_KeyUp(KeyCode As Integer, Shift As Integer)

    If sdb_slab_len.Value = 0 Then Exit Sub
    If sdb_unit_cnt.Value = 0 Then
        sdb_long_len.Value = 0
        sdb_long_wgt.Value = 0
        sdb_slab_rem_wgt.Value = sdb_slab_wgt.Value
        Exit Sub
    End If
    
    sdb_long_len.Value = sdb_unit_cnt.Value * sdb_slab_len.Value + ((sdb_unit_cnt.Value - 1) * 10)
    sdb_long_wgt.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & "," & sdb_long_len.Value & ",0) FROM DUAL")
    sdb_slab_rem_wgt.Value = sdb_ord_rem_wgt - (lProd_Wgt * Val(cbo_prod_cnt.Text) * sdb_unit_cnt.Value * sdb_req_cnt.Value)
    
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        
        DD.nameType = "1"
        DD.sWitch = "MS"
        
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_name
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
        If txt_stlgrd.Text <> "" And Ord_Fl = "2" Then
            sdb_slab_wgt.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & "," & sdb_slab_len.Value & ",0) FROM DUAL")
        End If
        
    Else
    
        If Len(Trim(txt_stlgrd.Text)) = txt_stlgrd.MaxLength Then
            txt_stlgrd_name.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
            If Ord_Fl = "2" Then
                sdb_slab_wgt.Value = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & "," & sdb_slab_len.Value & ",0) FROM DUAL")
            End If
        Else
            txt_stlgrd_name.Text = ""
        End If
        
    End If
    
End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub

Private Sub Gp_Process_Exec()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
'    Exit Sub '-------------------------
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call CEC2140C.P_SLAB_ADD_REQUEST ('" & Ord_Fl & "','C3','" & txt_ord_no.Text & "','" & txt_ord_item.Text & "','" _
                                                    & txt_stlgrd.Text & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & "," _
                                                    & sdb_slab_len.Value & "," & sdb_slab_wgt.Value & "," & sdb_slab_rem_cnt.Value & "," _
                                                    & Val(cbo_prod_cnt.Text) & ",'" & UCase(txt_hcr_fl.Text) & "','" _
                                                    & UCase(txt_plt.Text) & "'," & sdb_unit_cnt.Value & "," & sdb_req_cnt.Value & ",'" _
                                                    & udt_req_date.RawData & "','" & sUserID & "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
 
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        Call Gp_MsgBoxDisplay("申请紧急坯完了..!!", "I")
        CEG1050C.Complete = True
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Unload Me
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub
