VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ABA1020C 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "订单序列录入_ABA1020C"
   ClientHeight    =   10470
   ClientLeft      =   240
   ClientTop       =   780
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   5280
      Top             =   780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "订单序列状态"
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   15720
      TabIndex        =   86
      Top             =   0
      Width           =   15780
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   600
         Left            =   0
         TabIndex        =   87
         Top             =   0
         Width           =   15420
         _ExtentX        =   27199
         _ExtentY        =   1058
         BandCount       =   1
         _CBWidth        =   15420
         _CBHeight       =   600
         _Version        =   "6.7.8988"
         Child1          =   "MenuTool"
         MinHeight1      =   540
         Width1          =   15360
         NewRow1         =   0   'False
         BandStyle1      =   1
         Begin MSComctlLib.Toolbar MenuTool 
            Height          =   540
            Left            =   30
            TabIndex        =   88
            Top             =   30
            Width           =   15360
            _ExtentX        =   27093
            _ExtentY        =   953
            ButtonWidth     =   1244
            ButtonHeight    =   953
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList2"
            HotImageList    =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Clear"
                  Object.ToolTipText     =   "空界面"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line1"
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Save"
                  Object.ToolTipText     =   "保存"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Delete"
                  Object.ToolTipText     =   "删除"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line2"
                  Style           =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Copy"
                  Object.ToolTipText     =   "复制"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Paste"
                  Object.ToolTipText     =   "粘贴"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line4"
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Exit"
                  Object.ToolTipText     =   "退出"
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1365
      Index           =   0
      Left            =   60
      TabIndex        =   30
      Top             =   4755
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   2408
      _Version        =   196609
      ForeColor       =   16711680
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "尺寸"
      Begin VB.TextBox TXT_SIZE_KND 
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
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   96
         Tag             =   "尺寸分类"
         Top             =   225
         Width           =   495
      End
      Begin VB.TextBox TXT_SIZE_NM 
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
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   95
         Tag             =   "Can_Fl_Name"
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt_ord_size 
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
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   35
         Tag             =   "Ord_Size"
         Top             =   240
         Width           =   3975
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   18
         Left            =   7800
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "订单尺寸"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord_thk 
         Height          =   315
         Left            =   1680
         TabIndex        =   36
         Tag             =   "订单厚度"
         Top             =   585
         Width           =   1575
         _Version        =   262145
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   19
         Left            =   120
         Top             =   585
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "厚度"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord_wid 
         Height          =   315
         Left            =   1680
         TabIndex        =   37
         Tag             =   "订单宽度"
         Top             =   960
         Width           =   1575
         _Version        =   262145
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   20
         Left            =   135
         Top             =   945
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         Caption         =   "宽度"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord_len 
         Height          =   315
         Left            =   9360
         TabIndex        =   38
         Tag             =   "Ord_Len"
         Top             =   600
         Width           =   1395
         _Version        =   262145
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   1
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   21
         Left            =   7800
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "长度"
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
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   3360
         Top             =   585
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "轧制目标厚度"
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
      Begin CSTextLibCtl.sidbEdit sdb_mb_thk 
         Height          =   315
         Left            =   4920
         TabIndex        =   94
         Tag             =   "轧制目标厚度"
         Top             =   600
         Width           =   1395
         _Version        =   262145
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Index           =   2
         Left            =   120
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "尺寸分类"
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Index           =   3
         Left            =   3360
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "长度范围"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord_LEN_MIN 
         Height          =   315
         Left            =   4920
         TabIndex        =   98
         Tag             =   "Tot_Wgt"
         Top             =   240
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord_LEN_MAX 
         Height          =   315
         Left            =   6480
         TabIndex        =   99
         Tag             =   "Tot_Wgt"
         Top             =   240
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   47
         Left            =   3360
         Top             =   960
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "LP板厚度1/2/3"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   48
         Left            =   7800
         Top             =   960
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "LP板长度1/2/3/4/5"
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
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK1 
         Height          =   315
         Left            =   5175
         TabIndex        =   119
         Tag             =   "订单宽度"
         Top             =   960
         Width           =   775
         _Version        =   262145
         _ExtentX        =   1367
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK3 
         Height          =   315
         Left            =   6670
         TabIndex        =   120
         Tag             =   "订单宽度"
         Top             =   960
         Width           =   775
         _Version        =   262145
         _ExtentX        =   1367
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK2 
         Height          =   315
         Left            =   5950
         TabIndex        =   121
         Tag             =   "订单宽度"
         Top             =   960
         Width           =   775
         _Version        =   262145
         _ExtentX        =   1367
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN1 
         Height          =   315
         Left            =   9705
         TabIndex        =   122
         Tag             =   "Tot_Wgt"
         Top             =   960
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN2 
         Height          =   315
         Left            =   10760
         TabIndex        =   123
         Tag             =   "Tot_Wgt"
         Top             =   960
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN3 
         Height          =   315
         Left            =   11810
         TabIndex        =   124
         Tag             =   "Tot_Wgt"
         Top             =   960
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN4 
         Height          =   315
         Left            =   12870
         TabIndex        =   125
         Tag             =   "Tot_Wgt"
         Top             =   960
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN5 
         Height          =   315
         Left            =   13925
         TabIndex        =   126
         Tag             =   "Tot_Wgt"
         Top             =   960
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         Height          =   195
         Left            =   6240
         TabIndex        =   97
         Top             =   285
         Width           =   165
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2085
      Left            =   60
      TabIndex        =   6
      Top             =   1200
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   3678
      _Version        =   196609
      ForeColor       =   16711680
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "共用"
      Begin VB.ComboBox cbo_imp_cont 
         Height          =   300
         ItemData        =   "ABA1020C.frx":0000
         Left            =   11610
         List            =   "ABA1020C.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txt_extra_fl_1 
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
         Left            =   14880
         MaxLength       =   11
         TabIndex        =   127
         Tag             =   "Ord_No"
         Top             =   990
         Width           =   690
      End
      Begin VB.TextBox txt_urgnt_fl_name 
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
         Left            =   12210
         MaxLength       =   40
         TabIndex        =   117
         Tag             =   "urgnt_fl"
         Top             =   990
         Width           =   945
      End
      Begin VB.TextBox txt_cust_req_plant_name 
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
         Left            =   9480
         MaxLength       =   40
         TabIndex        =   76
         Tag             =   "Cust_Rea_Plant"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt_cust_req_plant 
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
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   75
         Tag             =   "Cust_Rea_Plant"
         Top             =   1710
         Width           =   585
      End
      Begin VB.TextBox txt_prod_dgr_name 
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
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   29
         Tag             =   "Prod_Dgr"
         Top             =   630
         Width           =   2500
      End
      Begin VB.TextBox txt_prod_dgr 
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
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   28
         Tag             =   "产品等级"
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox txt_hold_fl 
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
         Left            =   11610
         MaxLength       =   1
         TabIndex        =   27
         Tag             =   "Hold_Fl"
         Top             =   1350
         Width           =   600
      End
      Begin VB.TextBox txt_hold_fl_name 
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
         Left            =   12210
         MaxLength       =   40
         TabIndex        =   26
         Tag             =   "Hold_Fl"
         Top             =   1350
         Width           =   945
      End
      Begin VB.TextBox txt_sale_emp_id_name 
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
         Left            =   7515
         MaxLength       =   40
         TabIndex        =   25
         Tag             =   "Sale_Emp_ID"
         Top             =   1350
         Width           =   2200
      End
      Begin VB.TextBox txt_sale_emp_id 
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
         Left            =   6600
         MaxLength       =   7
         TabIndex        =   24
         Tag             =   "Sale_Emp_ID"
         Top             =   1350
         Width           =   900
      End
      Begin VB.TextBox txt_dest_cd_name 
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
         Left            =   2550
         MaxLength       =   40
         TabIndex        =   23
         Tag             =   "Dest_Cd"
         Top             =   1350
         Width           =   2230
      End
      Begin VB.TextBox txt_dest_cd 
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   22
         Tag             =   "Dest_Cd"
         Top             =   1350
         Width           =   870
      End
      Begin VB.TextBox txt_urgnt_fl 
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
         Left            =   11610
         MaxLength       =   1
         TabIndex        =   21
         Tag             =   "Urgnt_Fl"
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txt_end_cust_cd_name 
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
         Left            =   7515
         MaxLength       =   40
         TabIndex        =   20
         Tag             =   "End_Cust_Cd"
         Top             =   990
         Width           =   2200
      End
      Begin VB.TextBox txt_end_cust_cd 
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
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "End_Cust_Cd"
         Top             =   990
         Width           =   900
      End
      Begin VB.TextBox txt_ord_knd 
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
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   18
         Tag             =   "订单种类"
         Top             =   990
         Width           =   600
      End
      Begin VB.TextBox txt_ord_knd_name 
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
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   17
         Tag             =   "Ord_Knd"
         Top             =   990
         Width           =   2500
      End
      Begin VB.TextBox txt_dept_cd_name 
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
         Left            =   12210
         MaxLength       =   40
         TabIndex        =   16
         Tag             =   "Dept_Cd"
         Top             =   630
         Width           =   2500
      End
      Begin VB.TextBox txt_dept_cd 
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
         Left            =   11610
         MaxLength       =   3
         TabIndex        =   15
         Tag             =   "Dept_Cd"
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox txt_ord_cust_cd_name 
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
         Left            =   7515
         MaxLength       =   40
         TabIndex        =   14
         Tag             =   "Ord_Cust_Cd"
         Top             =   630
         Width           =   2200
      End
      Begin VB.TextBox txt_ord_cust_cd 
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
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   13
         Tag             =   "Ord_Cust_Cd"
         Top             =   630
         Width           =   900
      End
      Begin VB.TextBox txt_sale_way 
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
         Left            =   11610
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "Sale_Way"
         Top             =   270
         Width           =   600
      End
      Begin VB.TextBox txt_sale_way_name 
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
         Left            =   12210
         MaxLength       =   40
         TabIndex        =   11
         Tag             =   "Sale_Way"
         Top             =   270
         Width           =   2500
      End
      Begin VB.TextBox txt_cust_cd 
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
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   10
         Tag             =   "客户代码"
         Top             =   270
         Width           =   900
      End
      Begin VB.TextBox txt_cust_cd_name 
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
         Left            =   7515
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Cust_Cd"
         Top             =   270
         Width           =   2200
      End
      Begin VB.TextBox txt_prod_cd_name 
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
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Prod_Cd"
         Top             =   270
         Width           =   2500
      End
      Begin VB.TextBox txt_prod_cd 
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
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "产品代码"
         Top             =   270
         Width           =   600
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   3
         Left            =   120
         Top             =   270
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "产品"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   4
         Left            =   5040
         Top             =   270
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "客户"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   5
         Left            =   10050
         Top             =   270
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "销售方式"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   13
         Left            =   5040
         Top             =   630
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "订单客户"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   6
         Left            =   10050
         Top             =   630
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "部门"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   16
         Left            =   120
         Top             =   990
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "订单种类"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   12
         Left            =   5040
         Top             =   990
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "最终客户"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   7
         Left            =   10050
         Top             =   990
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "是否紧急订单"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   15
         Left            =   120
         Top             =   1350
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "目的地"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   11
         Left            =   5040
         Top             =   1350
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "销售负责人"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   8
         Left            =   10050
         Top             =   1350
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "是否订单保留"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   17
         Left            =   120
         Top             =   630
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "产品等级"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   14
         Left            =   7800
         Top             =   1710
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "客户指定工厂"
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
      Begin CSTextLibCtl.sidbEdit sdb_prod_prc 
         Height          =   315
         Left            =   1680
         TabIndex        =   77
         Tag             =   "Prod_Prc"
         Top             =   1710
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   10
         Left            =   120
         Top             =   1710
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "产品单价"
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
      Begin CSTextLibCtl.sidbEdit sdb_trans_prc 
         Height          =   315
         Left            =   5880
         TabIndex        =   78
         Tag             =   "Trans_Prc"
         Top             =   1710
         Width           =   1860
         _Version        =   262145
         _ExtentX        =   3281
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   9
         Left            =   5040
         Top             =   1710
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "运费"
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
      Begin Threed.SSCommand cmd_upd_rush 
         Height          =   495
         Left            =   13320
         TabIndex        =   118
         Top             =   1440
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   873
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "订单标记/解锁"
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   63
         Left            =   13320
         Top             =   990
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "流向"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   60
         Left            =   10050
         Top             =   1680
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "是否重点合同"
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
   End
   Begin VB.TextBox txt_ord_sts 
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
      Left            =   6840
      MaxLength       =   1
      TabIndex        =   4
      Tag             =   "Ord_Sts"
      Top             =   780
      Width           =   600
   End
   Begin VB.TextBox txt_ord_sts_name 
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
      Left            =   7485
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "Ord_St_name"
      Top             =   780
      Width           =   2500
   End
   Begin VB.TextBox txt_ord_item 
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
      Left            =   3300
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "Ord_Item"
      Top             =   780
      Width           =   465
   End
   Begin VB.TextBox txt_ord_no 
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
      Left            =   1755
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "Ord_No"
      Top             =   780
      Width           =   1530
   End
   Begin InDate.UDate dtp_ord_accp_date 
      Height          =   315
      Left            =   11670
      TabIndex        =   2
      Tag             =   "Ord_Accp_Date"
      Top             =   780
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
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   0
      Left            =   180
      Top             =   780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "订单号/序列号"
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
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   2
      Left            =   10110
      Top             =   780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "订单接受日期"
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
   Begin VB.TextBox txt_dest_seq 
      Appearance      =   0  'Flat
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
      Left            =   14880
      TabIndex        =   5
      Text            =   "F4"
      Top             =   8280
      Visible         =   0   'False
      Width           =   285
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2340
      Index           =   1
      Left            =   60
      TabIndex        =   31
      Top             =   8055
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   4128
      _Version        =   196609
      ForeColor       =   16711680
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "修改"
      Begin VB.TextBox txt_can_date 
         Height          =   315
         Left            =   1680
         TabIndex        =   84
         Top             =   1320
         Width           =   1545
      End
      Begin VB.TextBox txt_mod_date 
         Height          =   315
         Left            =   1680
         TabIndex        =   83
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txt_can_emp_id 
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
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   67
         Tag             =   "Can_Emp_ID"
         Top             =   1680
         Width           =   900
      End
      Begin VB.TextBox txt_can_emp_id_name 
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
         Left            =   2580
         MaxLength       =   40
         TabIndex        =   66
         Tag             =   "Can_Emp_ID"
         Top             =   1680
         Width           =   2200
      End
      Begin VB.TextBox txt_can_fl_name 
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
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   65
         Tag             =   "Can_Fl"
         Top             =   960
         Width           =   2520
      End
      Begin VB.TextBox txt_can_fl 
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
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   64
         Tag             =   "Can_Fl"
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txt_mod_time 
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
         Height          =   315
         Left            =   3225
         MaxLength       =   8
         TabIndex        =   63
         Tag             =   "Mod_Time"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txt_mod_fl 
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
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   62
         Tag             =   "Can_Fl"
         Top             =   250
         Width           =   600
      End
      Begin VB.TextBox txt_mod_fl_name 
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
         Left            =   2295
         MaxLength       =   40
         TabIndex        =   61
         Tag             =   "Can_Fl"
         Top             =   250
         Width           =   2505
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   44
         Left            =   120
         Top             =   250
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "修改分类"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   45
         Left            =   120
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         Caption         =   "修改日期、时间"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   46
         Left            =   120
         Top             =   960
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "取消分类"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   47
         Left            =   120
         Top             =   1320
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "取消日期"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   48
         Left            =   120
         Top             =   1680
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "取消负责人"
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
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1830
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   6150
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   3228
      _Version        =   196609
      ForeColor       =   16711680
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "质量"
      Begin VB.TextBox TXT_HTM_MET2 
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
         Left            =   9465
         MaxLength       =   1
         TabIndex        =   116
         Tag             =   "结算方式"
         Top             =   1425
         Width           =   600
      End
      Begin VB.TextBox TXT_HTM_MET3 
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
         Left            =   12120
         MaxLength       =   1
         TabIndex        =   115
         Tag             =   "结算方式"
         Top             =   1440
         Width           =   600
      End
      Begin VB.TextBox TXT_HTM_MET1 
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
         Left            =   6840
         MaxLength       =   1
         TabIndex        =   114
         Tag             =   "结算方式"
         Top             =   1425
         Width           =   600
      End
      Begin VB.TextBox TXT_shot_blast 
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
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   113
         Tag             =   "价格性质"
         Top             =   1440
         Width           =   600
      End
      Begin VB.TextBox TXT_HTM_MET3_nm 
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
         Height          =   315
         Left            =   12840
         MaxLength       =   60
         TabIndex        =   112
         Tag             =   "Mod_Fl_name"
         Top             =   1440
         Width           =   1950
      End
      Begin VB.TextBox TXT_shot_blast_nm 
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
         Left            =   2265
         MaxLength       =   60
         TabIndex        =   111
         Tag             =   "价格性质"
         Top             =   1440
         Width           =   2760
      End
      Begin VB.TextBox TXT_HTM_MET2_nm 
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
         Height          =   315
         Left            =   10095
         MaxLength       =   60
         TabIndex        =   110
         Tag             =   "Mod_Fl_name"
         Top             =   1440
         Width           =   1950
      End
      Begin VB.TextBox TXT_HTM_MET1_nm 
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
         Left            =   7470
         MaxLength       =   60
         TabIndex        =   109
         Tag             =   "结算方式"
         Top             =   1425
         Width           =   1950
      End
      Begin VB.TextBox TXT_MATR_FL 
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
         Left            =   6840
         MaxLength       =   1
         TabIndex        =   105
         Tag             =   "力学性能"
         Top             =   675
         Width           =   600
      End
      Begin VB.TextBox TXT_MATR_FL_NM 
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
         Height          =   315
         Left            =   7485
         MaxLength       =   40
         TabIndex        =   104
         Tag             =   "Mod_Fl_name"
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox txt_UST_FL 
         Height          =   315
         Left            =   11475
         MaxLength       =   4
         TabIndex        =   103
         Tag             =   "UST"
         Top             =   690
         Width           =   645
      End
      Begin VB.TextBox Txt_ust_fl_name 
         Height          =   315
         Left            =   12135
         TabIndex        =   102
         Top             =   690
         Width           =   3000
      End
      Begin VB.TextBox txt_stlgrd 
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
         Left            =   11460
         MaxLength       =   11
         TabIndex        =   81
         Tag             =   "钢种"
         Top             =   1080
         Width           =   1400
      End
      Begin VB.TextBox txt_stlgrd_name 
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
         Left            =   12870
         MaxLength       =   60
         TabIndex        =   80
         Tag             =   "STLGRD"
         Top             =   1080
         Width           =   2265
      End
      Begin VB.TextBox txt_insp_cd 
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
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   60
         Tag             =   "Test Method"
         Top             =   1065
         Width           =   600
      End
      Begin VB.TextBox txt_insp_cd_name 
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
         Left            =   7470
         MaxLength       =   40
         TabIndex        =   59
         Tag             =   "Test Method"
         Top             =   1065
         Width           =   1815
      End
      Begin VB.TextBox txt_cust_spec_no 
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
         Left            =   11460
         MaxLength       =   15
         TabIndex        =   58
         Tag             =   "Cust_Spec_No"
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_cust_spec_no_det 
         Height          =   310
         Left            =   12615
         TabIndex        =   57
         Top             =   300
         Width           =   2520
      End
      Begin VB.ComboBox cbo_india 
         Height          =   300
         ItemData        =   "ABA1020C.frx":0014
         Left            =   6285
         List            =   "ABA1020C.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txt_enduse_cd_name 
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
         Left            =   2260
         MaxLength       =   40
         TabIndex        =   54
         Tag             =   "Enduse_Cd"
         Top             =   1050
         Width           =   2775
      End
      Begin VB.TextBox txt_enduse_cd 
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
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   53
         Tag             =   "订单用途"
         Top             =   1050
         Width           =   600
      End
      Begin VB.TextBox txt_stdspec_yy 
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
         Left            =   4020
         MaxLength       =   4
         TabIndex        =   52
         Tag             =   "标准年度"
         Top             =   300
         Width           =   1005
      End
      Begin VB.TextBox txt_stdspec 
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
         Left            =   1665
         MaxLength       =   18
         TabIndex        =   51
         Tag             =   "标准代码"
         Top             =   300
         Width           =   2355
      End
      Begin VB.TextBox txt_stdspec_name 
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
         Left            =   1650
         MaxLength       =   40
         TabIndex        =   50
         Tag             =   "STDSPEC"
         Top             =   675
         Width           =   3390
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   41
         Left            =   120
         Top             =   298
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "标准/年度"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   43
         Left            =   120
         Top             =   1050
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "订单用途"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   38
         Left            =   5280
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "内径"
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
      Begin CSTextLibCtl.sidbEdit sdb_outdia 
         Height          =   315
         Left            =   8535
         TabIndex        =   56
         Tag             =   "Ord_Len"
         Top             =   285
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   39
         Left            =   7470
         Top             =   285
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         Caption         =   "外径"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   35
         Left            =   9900
         Top             =   300
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "客户特殊要求"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   40
         Left            =   5280
         Top             =   1065
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "检查机关"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   37
         Left            =   9900
         Top             =   1080
         Width           =   1500
         _ExtentX        =   2646
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   34
         Left            =   9900
         Top             =   690
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "UST"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   5280
         Top             =   690
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "力学性能"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   120
         Top             =   1440
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "抛丸"
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   5280
         Top             =   1425
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "热处理方式"
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
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2385
      Index           =   3
      Left            =   5220
      TabIndex        =   33
      Top             =   8040
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   4207
      _Version        =   196609
      ForeColor       =   16711680
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "交货"
      Begin VB.OptionButton opt_modify_knd 
         Caption         =   "整定单修改"
         Height          =   180
         Index           =   1
         Left            =   6240
         TabIndex        =   137
         Top             =   960
         Width           =   1335
      End
      Begin Threed.SSCommand cmd_upd_cust_date 
         Height          =   405
         Left            =   7680
         TabIndex        =   134
         Top             =   920
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   714
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "交货期修改"
      End
      Begin VB.OptionButton opt_modify_knd 
         Caption         =   "项次修改"
         Height          =   180
         Index           =   0
         Left            =   5160
         TabIndex        =   136
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txt_jit_stringa 
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
         Left            =   6240
         MaxLength       =   30
         TabIndex        =   133
         Tag             =   "ship_No"
         Top             =   1320
         Width           =   1395
      End
      Begin VB.TextBox txt_jitid 
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
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   132
         Tag             =   "ship_No"
         Top             =   1680
         Width           =   1395
      End
      Begin VB.TextBox txt_jit_stringc 
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
         Left            =   6240
         MaxLength       =   30
         TabIndex        =   131
         Tag             =   "ship_No"
         Top             =   1680
         Width           =   1395
      End
      Begin VB.TextBox txt_jit_stringb 
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
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   130
         Tag             =   "ship_No"
         Top             =   1320
         Width           =   1395
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   33
         Left            =   7680
         Top             =   1320
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "批次号"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   32
         Left            =   7680
         Top             =   560
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "定制配送"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   1
         Left            =   120
         Top             =   1320
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "加喷"
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
      Begin VB.TextBox txt_jit_flag 
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
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   129
         Tag             =   "ship_No"
         Top             =   560
         Width           =   1395
      End
      Begin VB.TextBox txt_stdmark 
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
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   128
         Tag             =   "Marking_Way"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txt_color 
         Height          =   315
         Left            =   1680
         TabIndex        =   108
         Tag             =   "color_bz"
         Top             =   2040
         Width           =   8625
      End
      Begin VB.TextBox TXT_cust_REQ_NO 
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
         Left            =   6240
         MaxLength       =   30
         TabIndex        =   107
         Tag             =   "Cust_REQ_No"
         Top             =   560
         Width           =   1395
      End
      Begin VB.TextBox txt_ship_no 
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
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   106
         Tag             =   "ship_No"
         Top             =   1320
         Width           =   3195
      End
      Begin VB.TextBox txt_dest_detail 
         Height          =   345
         Left            =   7440
         TabIndex        =   82
         Top             =   840
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.TextBox txt_payment_fl 
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
         Left            =   8880
         MaxLength       =   1
         TabIndex        =   73
         Tag             =   "Payment_Fl"
         Top             =   600
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txt_payment_fl_name 
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
         Left            =   8400
         MaxLength       =   40
         TabIndex        =   72
         Tag             =   "urgnt_fl"
         Top             =   600
         Visible         =   0   'False
         Width           =   2500
      End
      Begin VB.TextBox txt_stamp_name 
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
         Left            =   6840
         MaxLength       =   40
         TabIndex        =   71
         Tag             =   "Stamp"
         Top             =   200
         Width           =   3435
      End
      Begin VB.TextBox txt_stamp 
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
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   70
         Tag             =   "Stamp"
         Top             =   200
         Width           =   600
      End
      Begin VB.TextBox txt_marking_way_name 
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
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   69
         Tag             =   "Marking_Way"
         Top             =   200
         Width           =   2530
      End
      Begin VB.TextBox txt_marking_way 
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
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   68
         Tag             =   "标识方式"
         Top             =   200
         Width           =   600
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   51
         Left            =   120
         Top             =   200
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "标识方式"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   52
         Left            =   5040
         Top             =   200
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "色标方式"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   58
         Left            =   5040
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "资金入帐"
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
      Begin CSTextLibCtl.sidbEdit sdb_discon_prc 
         Height          =   315
         Left            =   8040
         TabIndex        =   74
         Tag             =   "Discon_Prc"
         Top             =   600
         Visible         =   0   'False
         Width           =   1905
         _Version        =   262145
         _ExtentX        =   3360
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   56
         Left            =   5040
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "折扣金额"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   53
         Left            =   120
         Top             =   560
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "系统计算交货期"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   54
         Left            =   120
         Top             =   920
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "客户要求交货期"
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
      Begin InDate.UDate dtp_cust_del_to_date 
         Height          =   300
         Left            =   3480
         TabIndex        =   90
         Top             =   920
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
         MaxLength       =   10
      End
      Begin InDate.UDate dtp_cust_del_fr_date 
         Height          =   300
         Left            =   1680
         TabIndex        =   91
         Top             =   920
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
         MaxLength       =   10
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Index           =   0
         Left            =   3120
         Top             =   920
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   "至"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Index           =   1
         Left            =   3120
         Top             =   560
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   "至"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
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
      Begin InDate.UDate dtp_del_to_date 
         Height          =   300
         Left            =   3480
         TabIndex        =   92
         Top             =   560
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
         MaxLength       =   10
      End
      Begin InDate.UDate dtp_del_fr_date 
         Height          =   300
         Left            =   1680
         TabIndex        =   93
         Top             =   560
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
         MaxLength       =   10
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   42
         Left            =   5040
         Top             =   560
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "客户要求号"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   62
         Left            =   120
         Top             =   2040
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "色标及备注"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   31
         Left            =   120
         Top             =   1680
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "侧喷加印"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   50
         Left            =   5040
         Top             =   1680
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "分段号"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   57
         Left            =   7680
         Top             =   1680
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "交付编号"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   59
         Left            =   5040
         Top             =   1320
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "船号"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   55
         Left            =   6240
         Top             =   840
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "详细目的地"
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
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1365
      Index           =   4
      Left            =   60
      TabIndex        =   34
      Top             =   3315
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   2408
      _Version        =   196609
      ForeColor       =   16711680
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "重量"
      Begin VB.TextBox txt_trim_fl 
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
         Left            =   9495
         MaxLength       =   1
         TabIndex        =   101
         Tag             =   "切边"
         Top             =   960
         Width           =   420
      End
      Begin VB.TextBox txt_trim_fl_name 
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
         Left            =   9915
         MaxLength       =   40
         TabIndex        =   100
         Tag             =   "Trim_Fl"
         Top             =   960
         Width           =   1200
      End
      Begin VB.TextBox txt_wgt_grp_name 
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
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   49
         Tag             =   "Wgt_Grp"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txt_wgt_grp 
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
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   48
         Tag             =   "交货重量"
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txt_del_tol_unit_name 
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
         Left            =   13860
         MaxLength       =   40
         TabIndex        =   47
         Tag             =   "Del_Tol_Unit"
         Top             =   600
         Width           =   885
      End
      Begin VB.TextBox txt_del_tol_unit 
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
         Left            =   13230
         MaxLength       =   1
         TabIndex        =   46
         Tag             =   "交付公差单位"
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox txt_wgt_unit_name 
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
         Left            =   10080
         MaxLength       =   40
         TabIndex        =   42
         Tag             =   "Sale_Way"
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txt_wgt_unit 
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
         Left            =   9480
         MaxLength       =   1
         TabIndex        =   41
         Tag             =   "重量单位"
         Top             =   240
         Width           =   600
      End
      Begin CSTextLibCtl.sidbEdit sdb_tot_wgt 
         Height          =   315
         Left            =   1680
         TabIndex        =   39
         Tag             =   "重量"
         Top             =   240
         Width           =   1650
         _Version        =   262145
         _ExtentX        =   2910
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
         RawData         =   "0.000"
         Text            =   " 0.000"
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   22
         Left            =   120
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "重量"
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
      Begin CSTextLibCtl.sidbEdit sdb_prod_wgt 
         Height          =   315
         Left            =   5550
         TabIndex        =   40
         Tag             =   "产品单重"
         Top             =   240
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   30
         Left            =   4080
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "产品单重"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   23
         Left            =   7920
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "重量单位"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   61
         Left            =   11760
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "产品单重下限"
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
      Begin CSTextLibCtl.sidbEdit sdb_del_tol_min 
         Height          =   315
         Left            =   1680
         TabIndex        =   43
         Tag             =   "Del_Tol_Min"
         Top             =   600
         Width           =   1680
         _Version        =   262145
         _ExtentX        =   2963
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   2
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   24
         Left            =   120
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "交付公差下限"
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
      Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_max 
         Height          =   315
         Left            =   5550
         TabIndex        =   44
         Tag             =   "产品单重上限"
         Top             =   600
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   2
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   29
         Left            =   4080
         Top             =   600
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "产品单重上限"
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
      Begin CSTextLibCtl.sidbEdit sdb_del_tol_max 
         Height          =   315
         Left            =   9480
         TabIndex        =   45
         Tag             =   "交付公差上限"
         Top             =   600
         Width           =   1680
         _Version        =   262145
         _ExtentX        =   2963
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   2
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   25
         Left            =   7920
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "交付公差上限"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   28
         Left            =   11760
         Top             =   600
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "交付公差单位"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   26
         Left            =   120
         Top             =   960
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "交货重量"
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
      Begin CSTextLibCtl.sidbEdit sdb_num_prod 
         Height          =   315
         Left            =   5550
         TabIndex        =   79
         Tag             =   "Num_Prod"
         Top             =   960
         Width           =   1530
         _Version        =   262145
         _ExtentX        =   2699
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   27
         Left            =   4080
         Top             =   960
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "产品数量"
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
      Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_min 
         Height          =   315
         Left            =   13230
         TabIndex        =   85
         Tag             =   "产品单重"
         Top             =   240
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   36
         Left            =   7920
         Top             =   960
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "切边"
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4065
      Top             =   585
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":002A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":04E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":0803
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":0AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":0DC5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4665
      Top             =   585
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1277
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1577
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1657
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1860
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1998
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1BD3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSCommand SCmd_COPY 
      Height          =   375
      Left            =   13560
      TabIndex        =   89
      Top             =   720
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单主要标准"
   End
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   49
      Left            =   0
      Top             =   0
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "定制配送"
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   15225
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "ABA1020C"
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
'-- Program Name      Order-Detail
'-- Program ID        ABA1021C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.19
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
Public Dis_sw As Boolean            'Display sw Boolean
Public MSG As Boolean               'Display sw Boolean

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "PopMaster"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary )", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
                 Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ord_item, "p", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_prod_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_prod_dgr, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_prod_dgr_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_stdspec, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stdspec_yy, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stdspec_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stlgrd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ord_cust_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ord_cust_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_cust_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_end_cust_cd, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_end_cust_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_thk, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_wid, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_len, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(cbo_india, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_outdia, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_wgt_grp, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_wgt_grp_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'               Call Gp_Ms_Collection(txt_del_cond, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(txt_del_cond_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_num_prod, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_del_fr_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_del_to_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_cust_del_fr_date, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_cust_del_to_date, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(dtp_ord_accp_date, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'             Call Gp_Ms_Collection(txt_transp_way, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'        Call Gp_Ms_Collection(txt_transp_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'           Call Gp_Ms_Collection(txt_payment_cond, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'      Call Gp_Ms_Collection(txt_payment_cond_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_sale_way, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_sale_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_del_tol_unit, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_del_tol_unit_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_del_tol_max, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_del_tol_min, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_dept_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_dept_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_sale_emp_id, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_sale_emp_id_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_wgt_unit, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wgt_unit_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_urgnt_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_urgnt_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_prod_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_prod_wgt_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_prod_wgt_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_enduse_cd, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_enduse_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_trim_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_trim_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_spec_no, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cust_spec_no_det, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_cust_req_plant, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cust_req_plant_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_prod_prc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'               Call Gp_Ms_Collection(txt_extra_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(txt_extra_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_discon_prc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_trans_prc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_payment_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_payment_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_marking_way, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_marking_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_stamp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stamp_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'               Call Gp_Ms_Collection(txt_pack_way, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(txt_pack_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'           Call Gp_Ms_Collection(sdb_pack_wgt_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'           Call Gp_Ms_Collection(sdb_pack_wgt_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ord_knd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_knd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_hold_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_hold_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_mod_fl, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_mod_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_mod_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_mod_time, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_can_fl, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_can_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_can_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_can_emp_id, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_can_emp_id_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_dest_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_dest_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'            Call Gp_Ms_Collection(txt_dest_detail, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ord_sts, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_sts_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_insp_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_insp_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'               Call Gp_Ms_Collection(txt_currency, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(txt_currency_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_tot_wgt, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ord_size, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_UST_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(Txt_ust_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_mb_thk, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_SIZE_KND, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_SIZE_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_ord_LEN_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_ord_LEN_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_MATR_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_MATR_FL_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ship_no, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_cust_REQ_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_color, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_shot_blast, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_shot_blast_nm, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_HTM_MET1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_HTM_MET1_nm, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_HTM_MET2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_HTM_MET2_nm, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_HTM_MET3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_HTM_MET3_nm, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_THK1, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_THK2, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_THK3, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_LEN1, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_LEN2, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_LEN3, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_LEN4, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_LP_LEN5, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_extra_fl_1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_stdmark, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
               Call Gp_Ms_Collection(txt_jit_flag, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_jit_stringa, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_jit_stringb, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_jit_stringc, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_jitid, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  
               Call Gp_Ms_Collection(cbo_imp_cont, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             

            
     
     txt_ord_no.BackColor = &HE0E0E0
     txt_ord_item.BackColor = &HE0E0E0
     txt_prod_cd.BackColor = &HE0E0E0
     txt_prod_cd_name.BackColor = &HE0E0E0
     txt_prod_dgr.BackColor = &HE0E0E0
     txt_prod_dgr_name.BackColor = &HE0E0E0
     txt_cust_cd.BackColor = &HE0E0E0
     txt_cust_cd_name.BackColor = &HE0E0E0
     sdb_ord_thk.BackColor = &HE0E0E0
     sdb_ord_wid.BackColor = &HE0E0E0
     sdb_ord_len.BackColor = &HE0E0E0
     sdb_num_prod.BackColor = &HE0E0E0
     dtp_del_fr_date.BackColor = &HE0E0E0
     dtp_del_to_date.BackColor = &HE0E0E0
     dtp_ord_accp_date.BackColor = &HE0E0E0
    
     txt_sale_way.BackColor = &HE0E0E0
     txt_sale_way_name.BackColor = &HE0E0E0
     txt_dept_cd.BackColor = &HE0E0E0
     txt_dept_cd_name.BackColor = &HE0E0E0
     txt_ord_knd.BackColor = &HE0E0E0
     txt_ord_knd_name.BackColor = &HE0E0E0
     txt_mod_fl.BackColor = &HE0E0E0
     txt_mod_fl_name.BackColor = &HE0E0E0
     txt_mod_date.BackColor = &HE0E0E0
     txt_mod_time.BackColor = &HE0E0E0
     txt_can_fl.BackColor = &HE0E0E0
     txt_can_fl_name.BackColor = &HE0E0E0
     txt_can_date.BackColor = &HE0E0E0
     txt_can_emp_id.BackColor = &HE0E0E0
     txt_can_emp_id_name.BackColor = &HE0E0E0
     txt_ord_sts.BackColor = &HE0E0E0
     txt_ord_sts_name.BackColor = &HE0E0E0
     
     MSG = True
    
    'MASTER Collection
    
'     Mc1.Add Item:="ABA1020C.P_UPDATE_ORD_RUSH", Key:="P-M"
     Mc1.Add Item:="ABA1020C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="ABA1020C.P_REFER", Key:="P-R"
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

Private Sub cmd_upd_cust_date_Click()
On Error GoTo Gp_Call_upd_cust_date_Error

    Dim v_fr_date As String
    Dim v_to_date As String
    
    
    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim modify_knd As String
    
    Dim adoCmd As ADODB.Command
    
    If opt_modify_knd(0).Value = True Then
      modify_knd = "0"
    Else
      modify_knd = "1"
    End If
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    v_fr_date = Mid(dtp_cust_del_fr_date, 1, 4) + Mid(dtp_cust_del_fr_date, 6, 2) + Mid(dtp_cust_del_fr_date, 9, 2)
    v_to_date = Mid(dtp_cust_del_to_date, 1, 4) + Mid(dtp_cust_del_to_date, 6, 2) + Mid(dtp_cust_del_to_date, 9, 2)
    
    sQuery = "{call ABA1020C.P_UPDATE_CUST_DEL_DATE ('" + txt_ord_no.Text + "','" + txt_ord_item.Text + "','" + v_fr_date + "','" + v_to_date + "','" + sUserID + "','" + modify_knd + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'OS Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay(txt_ord_no.Text + "-" + txt_ord_item.Text + ".订单信息修改成功。", "I")
        Call Form_QueryUnload(True, 0)
        Call Form_Exit
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Gp_Call_upd_cust_date_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Gp_Call_upd_cust_date_Error : " & Error)
End Sub

Private Sub cmd_upd_rush_Click()

On Error GoTo Gp_Callupd_rush_Error

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ABA1020C.P_UPDATE_ORD_RUSH ('" + txt_ord_no.Text + "','" + txt_ord_item.Text + "','" + txt_urgnt_fl.Text + "','" + txt_extra_fl_1.Text + "','" + txt_hold_fl.Text + "','" + cbo_imp_cont.Text + "','" + sUserID + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'OS Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay(txt_ord_no.Text + "-" + txt_ord_item.Text + ".订单信息修改成功。", "I")
        Call Form_QueryUnload(True, 0)
        Call Form_Exit
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Gp_Callupd_rush_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Gp_Callupd_rush_Error : " & Error)
End Sub

Private Sub Form_Activate()

    If Dis_sw = False Then
        Exit Sub
'    Else
'        Dis_sw = False
    End If
    
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    
    If Mc1("pControl").Item(1).Text = "" Then
        pControl(1).Text = ABA1010C.txt_ord_no.Text
        pControl(1).Enabled = True
        pControl(2).Enabled = True
            
        txt_cust_cd.Text = ABA1010C.txt_cust_cd.Text
        txt_cust_cd_name.Text = ABA1010C.txt_cust_cd_name.Text
        txt_prod_cd.Text = ABA1010C.txt_prod_cd.Text
        txt_prod_cd_name.Text = ABA1010C.txt_prod_cd_name.Text
        txt_prod_dgr.Text = ABA1010C.txt_prod_dgr.Text
        txt_prod_dgr_name.Text = ABA1010C.txt_prod_dgr_name.Text
        txt_dept_cd.Text = ABA1010C.txt_dept_cd.Text
        txt_dept_cd_name.Text = ABA1010C.txt_dept_cd_name.Text
        txt_sale_way.Text = ABA1010C.txt_sale_way.Text
        txt_sale_way_name.Text = ABA1010C.txt_sale_way_name.Text
        txt_ord_knd.Text = ABA1010C.txt_ord_knd.Text
        txt_ord_knd_name.Text = ABA1010C.txt_ord_knd_name.Text
        txt_dest_cd.Text = ABA1010C.TXT_DEST.Text
        txt_dest_cd_name.Text = ABA1010C.TXT_DEST_NM.Text
'        txt_transp_way.Text = ABA1010C.TXT_YSFS.Text
'        txt_transp_way_name.Text = ABA1010C.TXT_YSFS_NM.Text
'        txt_currency.Text = "RMB"
'        txt_currency_name.Text = "人民币"
'        txt_trim_fl.Text = "Y"
'        txt_trim_fl_name.Text = "切边"
'        txt_UST_FL.Text = "X"
'        Txt_ust_fl_name.Text = "NO UST"
        txt_ship_no.Text = ABA1010C.txt_ship_no.Text
        txt_end_cust_cd.Text = ABA1010C.txt_cust_cd.Text
        txt_end_cust_cd_name.Text = ABA1010C.txt_cust_cd_name.Text
        
        dtp_del_fr_date.Text = ""
        dtp_del_to_date.Text = ""
        dtp_cust_del_fr_date.Text = ""
        dtp_cust_del_to_date.Text = ""
        
      
        
        pControl(2).SetFocus
        
    Else
        Call Gf_Ms_Refer(M_CN1, Mc1)
         If txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text = "02" Then
            sdb_ord_LEN_MIN.Enabled = True
            sdb_ord_LEN_MAX.Enabled = True
         Else
        '    TXT_SIZE_NM.Text = "定尺"
            sdb_ord_LEN_MIN.Enabled = False
            sdb_ord_LEN_MAX.Enabled = False
         End If
        
    End If
    
    If txt_ord_knd.Text = "A" Or txt_ord_knd.Text = "S" Then
       sdb_prod_prc.BackColor = &HC0FFFF
    Else
       sdb_prod_prc.BackColor = &H80000005
    End If

    
    If txt_ord_sts.Text = "A" Or txt_ord_sts.Text = "" Then
'       Call Gp_Ms_ControlLock(Mc1("iControl"), False)
       TXT_SIZE_KND.Enabled = True
       txt_cust_req_plant.Enabled = True
       
       Call Gp_Ms_ControlLock(Mc1("lControl"), True)
       If txt_ord_sts.Text = "" Then
          pControl(1).Enabled = True
          pControl(2).Enabled = True
       End If
       MenuTool.Buttons.Item(1).Enabled = True
       MenuTool.Buttons.Item(3).Enabled = False
       MenuTool.Buttons.Item(4).Enabled = False
    Else
       Call Gp_Ms_ControlLock(Mc1("iControl"), True)
       MenuTool.Buttons.Item(1).Enabled = True
       MenuTool.Buttons.Item(3).Enabled = False
       MenuTool.Buttons.Item(4).Enabled = False
    End If

    If txt_ord_sts.Text = "A" Or txt_ord_sts.Text = "" Then
    
        Select Case txt_prod_cd.Text
    
               Case "HC"
               
                    '订单内径/订单外径/检查机关/切边代码/UST
                    cbo_india.Enabled = True
                    sdb_outdia.Enabled = True
                    txt_insp_cd.Enabled = False
                    txt_trim_fl.Enabled = False
                    txt_UST_FL.Enabled = False
                    TXT_SIZE_KND.Enabled = False
                    TXT_SIZE_KND.Text = "02"
                    TXT_SIZE_NM = "单定尺"
                    txt_cust_req_plant.Enabled = False
                    txt_cust_req_plant.Text = "C1"
                    txt_cust_req_plant_name = "#1 轧钢"
                    If txt_prod_dgr.Text <> "1" And txt_prod_dgr.Text <> "2" Then
                        TXT_MATR_FL.Enabled = False
                        TXT_MATR_FL.Text = "N"
                        TXT_MATR_FL_NM = "不保证产品力学性能"
                    Else
                    TXT_MATR_FL.Enabled = True
                    End If
                    '重量单位
                    txt_wgt_unit.Enabled = True
                    
                    '产品单重下限/产品单重上限/产品单重
                    sdb_prod_wgt_min.Enabled = True
                    sdb_prod_wgt_max.Enabled = True
                    sdb_prod_wgt.Enabled = True
                    
                    '包装重量下限/包装重量上限/包装方法
'                    sdb_pack_wgt_min.Enabled = True
'                    sdb_pack_wgt_max.Enabled = True
'                    txt_pack_way.Enabled = True
                    
                    '交付公差单位
                    txt_del_tol_unit.Enabled = False
                    txt_del_tol_unit.Text = "W"
    
               Case "SL"
               
                    '订单内径/订单外径/检查机关/切边代码/UST
                    cbo_india.Enabled = False
                    sdb_outdia.Enabled = False
                    txt_insp_cd.Enabled = False
                    txt_trim_fl.Enabled = False
                    txt_UST_FL.Enabled = False
                    TXT_MATR_FL.Enabled = False
                    TXT_MATR_FL.Text = "N"
                    TXT_MATR_FL_NM = "不保证产品力学性能"
                    '重量单位
                    txt_wgt_unit.Enabled = True
                    
                    '产品单重下限/产品单重上限/产品单重
                    sdb_prod_wgt_min.Enabled = False
                    sdb_prod_wgt_max.Enabled = False
                    sdb_prod_wgt.Enabled = False
                    
                    '包装重量下限/包装重量上限/包装方法
'                    sdb_pack_wgt_min.Enabled = False
'                    sdb_pack_wgt_max.Enabled = False
'                    txt_pack_way.Enabled = False
                    '交付公差单位
                    txt_del_tol_unit.Enabled = False
                    txt_del_tol_unit.Text = "W"
    
                    TXT_SIZE_KND.Text = "01"
                    TXT_SIZE_NM.Text = "定尺"
                    
               Case "PP"
                       
                    '订单内径/订单外径/检查机关/切边代码/UST
                    cbo_india.Enabled = False
                    sdb_outdia.Enabled = False
                    txt_insp_cd.Enabled = True
                    txt_trim_fl.Enabled = True
                    txt_UST_FL.Enabled = True
                    If txt_prod_dgr.Text <> "1" And txt_prod_dgr.Text <> "2" Then
                        TXT_MATR_FL.Enabled = False
                        TXT_MATR_FL.Text = "N"
                        TXT_MATR_FL_NM = "不保证产品力学性能"
                    Else
                    TXT_MATR_FL.Enabled = True
                    End If
                    '重量单位
                    txt_wgt_unit.Enabled = True
                    
                    '产品单重下限/产品单重上限/产品单重
                    sdb_prod_wgt_min.Enabled = False
                    sdb_prod_wgt_max.Enabled = False
                    sdb_prod_wgt.Enabled = False
                    
                    '包装重量下限/包装重量上限/包装方法
'                    sdb_pack_wgt_min.Enabled = False
'                    sdb_pack_wgt_max.Enabled = False
'                    txt_pack_way.Enabled = True
                    
                    '交付公差单位
                    txt_del_tol_unit.Enabled = False
                    txt_del_tol_unit.Text = "W"
    
        End Select
        
        If txt_prod_cd.Text = "HC" Then
            TXT_SIZE_KND.Enabled = False
            TXT_SIZE_KND.Text = "02"
            TXT_SIZE_NM = "单定尺"
            txt_cust_req_plant.Enabled = False
            txt_cust_req_plant.Text = "C1"
            txt_cust_req_plant_name = "#1 轧钢"
            cbo_india.BackColor = &HC0E0FF
        '    sdb_outdia.BackColor = &HC0E0FF
        '    txt_wgt_unit.BackColor = &HC0E0FF
            sdb_prod_wgt_min.BackColor = &HC0E0FF
            sdb_prod_wgt_max.BackColor = &HC0E0FF
            sdb_prod_wgt.BackColor = &HC0E0FF
'            sdb_pack_wgt_min.BackColor = &HC0E0FF
'            sdb_pack_wgt_max.BackColor = &HC0E0FF
'            txt_pack_way.BackColor = &HC0E0FF
'            txt_pack_way_name.BackColor = &HC0E0FF
        Else
            
            '颜色
            cbo_india.BackColor = &HE0E0E0
            sdb_outdia.BackColor = &HE0E0E0
            txt_wgt_unit.BackColor = &HE0E0E0
            txt_wgt_unit_name.BackColor = &HE0E0E0
            sdb_prod_wgt_min.BackColor = &HE0E0E0
            sdb_prod_wgt_max.BackColor = &HE0E0E0
            sdb_prod_wgt.BackColor = &HE0E0E0
'            sdb_pack_wgt_min.BackColor = &HE0E0E0
'            sdb_pack_wgt_max.BackColor = &HE0E0E0
'            txt_pack_way.BackColor = &HE0E0E0
'            txt_pack_way_name.BackColor = &HE0E0E0
        End If
    
        If txt_prod_cd.Text <> "PP" Then
            txt_del_tol_unit.Enabled = False
            txt_del_tol_unit_name.Enabled = False
            txt_del_tol_unit.Text = "W"
            txt_insp_cd.Enabled = False
            txt_insp_cd_name.Enabled = False
            txt_del_tol_unit.BackColor = &HE0E0E0
            txt_del_tol_unit_name.BackColor = &HE0E0E0
            txt_insp_cd.BackColor = &HE0E0E0
            txt_insp_cd_name.BackColor = &HE0E0E0
        End If
    
    End If
    
'    Call txt_extra_fl_LostFocus
    
    If txt_ord_knd.Text = "P" Then
       txt_dest_cd.Text = "NG0001"
       txt_dest_cd_name.Text = Gf_DestNameFind(M_CN1, Trim(txt_dest_cd.Text), 1)
       txt_dest_cd.Enabled = False
       txt_dest_cd_name.Enabled = False
    Else
       txt_dest_cd.Enabled = True
       txt_dest_cd_name.Enabled = True

    End If
    
    If txt_prod_cd.Text = "PP" Then
        txt_UST_FL.BackColor = &HC0FFFF
        txt_trim_fl.BackColor = &HC0FFFF
        TXT_MATR_FL.BackColor = &HC0FFFF
    Else
        txt_UST_FL.BackColor = &HE0E0E0
        txt_trim_fl.BackColor = &HE0E0E0

    End If
    
    If txt_prod_cd.Text = "SL" Then
       TXT_MATR_FL.BackColor = &HE0E0E0
    Else
       TXT_MATR_FL.BackColor = &HC0FFFF
    End If


'    dtp_del_fr_date.Text = ""
'    dtp_del_to_date.Text = ""
    
'   txt_can_date.Text = ""
'   txt_mod_date.Text = ""
    
    
    'dtp_del_fr_date.Enabled = True
    'dtp_del_to_date.Enabled = True
    dtp_ord_accp_date.Enabled = False
'    TXT_SIZE_KND.Text = "01"
'    TXT_SIZE_NM.Text = "定尺"
'    sdb_ord_LEN_MIN.Enabled = False
'    sdb_ord_LEN_MAX.Enabled = False

    Select Case Mid(sAuthority, 3, 1) ' Update
    
    Case "1"      'No Authority
             cmd_upd_rush.Enabled = True
             cmd_upd_cust_date.Enabled = True
             dtp_cust_del_fr_date.Enabled = True
             dtp_cust_del_to_date.Enabled = True
             cbo_imp_cont.Enabled = True
    End Select


    txt_urgnt_fl.Locked = False   '特殊放开权利，给修改紧急订单
    txt_urgnt_fl.Enabled = True
    txt_extra_fl_1.Locked = False
    txt_extra_fl_1.Enabled = True
    txt_hold_fl.Locked = False
    txt_hold_fl.Enabled = True
    Dis_sw = False


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority("ABA1010C") ', True)

    Call Popup_Menu_Setting

    Call Form_Define

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_ControlLock(Mc1("pControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_FormCenter(Me)
   
    Screen.MousePointer = vbDefault
    
'    TXT_SIZE_KND.Text = "01"
'    TXT_SIZE_NM.Text = "定尺"
'    sdb_ord_LEN_MIN.Enabled = False
'    sdb_ord_LEN_MAX.Enabled = False
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing

    Call ABA1010C.Form_Ref
    

End Sub

Public Sub Form_Exit()

    Unload Me
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Call Gp_Ms_ControlLock(Mc1("iControl"), False)
    
    MenuTool.Buttons(4).Enabled = False    'Delete
    MenuTool.Buttons(6).Enabled = False    'Copy
    MenuTool.Buttons(7).Enabled = False    'Paste
    
    txt_ord_no.Text = ABA1010C.txt_ord_no.Text
    txt_cust_cd.Text = ABA1010C.txt_cust_cd.Text
    txt_cust_cd_name.Text = ABA1010C.txt_cust_cd_name.Text
    txt_prod_cd.Text = ABA1010C.txt_prod_cd.Text
    txt_prod_cd_name.Text = ABA1010C.txt_prod_cd_name.Text
    txt_prod_dgr.Text = ABA1010C.txt_prod_dgr.Text
    txt_prod_dgr_name.Text = ABA1010C.txt_prod_dgr_name.Text
    txt_dept_cd.Text = ABA1010C.txt_dept_cd.Text
    txt_dept_cd_name.Text = ABA1010C.txt_dept_cd_name.Text
    txt_sale_way.Text = ABA1010C.txt_sale_way.Text
    txt_sale_way_name.Text = ABA1010C.txt_sale_way_name.Text
    txt_ord_knd.Text = ABA1010C.txt_ord_knd.Text
    txt_ord_knd_name.Text = ABA1010C.txt_ord_knd_name.Text
'    txt_currency.Text = "RMB"
'    txt_currency_name.Text = "人民币"
    
'    txt_currency.Enabled = False


    
    dtp_del_fr_date.Text = ""
    dtp_del_to_date.Text = ""
    dtp_cust_del_fr_date.Text = ""
    dtp_cust_del_to_date.Text = ""

    If txt_prod_cd.Text = "PP" Then
        If txt_prod_dgr.Text <> "1" And txt_prod_dgr.Text <> "2" Then
            TXT_MATR_FL.Enabled = False
            TXT_MATR_FL.Text = "N"
            TXT_MATR_FL_NM = "不保证产品力学性能"
        Else
        TXT_MATR_FL.Enabled = True
        End If

        txt_del_tol_unit.Text = ""
    ElseIf txt_prod_cd.Text = "HC" Then
            TXT_SIZE_KND.Enabled = False
            TXT_SIZE_KND.Text = "02"
            TXT_SIZE_NM = "单定尺"
            txt_cust_req_plant.Enabled = False
            txt_cust_req_plant.Text = "C1"
            txt_cust_req_plant_name = "#1 轧钢"
            If txt_prod_dgr.Text <> "1" And txt_prod_dgr.Text <> "2" Then
                TXT_MATR_FL.Enabled = False
                TXT_MATR_FL.Text = "N"
                TXT_MATR_FL_NM = "不保证产品力学性能"
            Else
                TXT_MATR_FL.Enabled = True
            End If
    
           txt_del_tol_unit.Text = "W"
    ElseIf txt_prod_cd.Text = "SL" Then
            TXT_MATR_FL.Text = "N"
            TXT_MATR_FL_NM = "不保证产品力学性能"
            txt_del_tol_unit.Text = "W"

            TXT_SIZE_KND.Text = "01"
            TXT_SIZE_NM.Text = "定尺"
    End If

End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then MenuTool.Buttons(4).Enabled = False   'Delete
    
End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    Dim sQuery As String
    
    
    
    Select Case Me.ActiveControl.Name
           Case "txt_ord_size"
'                 Call txt_ord_size_LostFocus
                 If MSG = False Then Exit Sub
           Case "sdb_prod_wgt"
'                 Call sdb_prod_wgt_LostFocus
           Case "sdb_tot_wgt"
'                 Call sdb_tot_wgt_LostFocus
'           Case "txt_pack_way"
'                 Call txt_pack_way_LostFocus
           Case "TXT_MATR_FL"
                 Call TXT_MATR_FL_LostFocus
           Case "txt_UST_FL"
                 Call txt_UST_FL_LostFocus
      
    End Select
    
'    Call txt_ord_size_LostFocus
'    If MSG = False Then Exit Sub
    
'    If ABA1010C.txt_prod_cd.Text = "PP" Then
'
'        If txt_UST_FL.Text = "" Then
'           Call Gp_MsgBoxDisplay("探伤必须输入", "I")
'           Exit Sub
'        End If
'
'        If txt_trim_fl.Text = "" Then
'           Call Gp_MsgBoxDisplay("切边必须输入", "I")
'           Exit Sub
'        End If
'
'        If TXT_MATR_FL.Text = "" Then
'           Call Gp_MsgBoxDisplay("力学性能必须输入", "I")
'           Exit Sub
'        End If
'
'        If Mid(txt_cust_req_plant.Text, 1, 1) = "B" Then
'           Call Gp_MsgBoxDisplay("工厂分配错误", "I")
'           Exit Sub
'        End If
'
'    End If
'
'    If ABA1010C.txt_prod_cd.Text = "SL" Then
'        If Mid(txt_cust_req_plant.Text, 1, 1) = "C" Then
'           Call Gp_MsgBoxDisplay("工厂分配错误", "I")
'           Exit Sub
'        End If
'    End If
'
'    If ABA1010C.txt_prod_cd.Text = "HC" Then
'
'        If TXT_MATR_FL.Text = "" Then
'           Call Gp_MsgBoxDisplay("力学性能必须输入", "I")
'           Exit Sub
'        End If
'
'        If Trim(txt_cust_req_plant.Text) <> "C1" Then
'           Call Gp_MsgBoxDisplay("工厂分配错误", "I")
'           Exit Sub
'        End If
'
'    End If
'
'    If ABA1010C.txt_prod_cd.Text = "HC" Then
'
'        If cbo_india.Text = "" Or cbo_india.Text = 0 Then
'           Call Gp_MsgBoxDisplay("内径必须输入", "I")
'           Exit Sub
'        End If
'
''        If txt_wgt_unit.Text = "" Or sdb_prod_wgt_min.Value = 0 Or sdb_prod_wgt_max.Value = 0 Or sdb_prod_wgt.Value = 0 Then
''           Call Gp_MsgBoxDisplay("单重，重量上下限，重量单位必须输入", "I")
''           Exit Sub
''        End If
'        If sdb_prod_wgt_min.Value = 0 Or sdb_prod_wgt_max.Value = 0 Or sdb_prod_wgt.Value = 0 Then
'           Call Gp_MsgBoxDisplay("单重，重量上下限必须输入", "I")
'           Exit Sub
'        End If
'
'        If sdb_pack_wgt_min.Value = 0 Or sdb_pack_wgt_max.Value = 0 Or txt_pack_way.Text = "" Then
'           Call Gp_MsgBoxDisplay("包装方式，包装重量必须输入", "I")
'           Exit Sub
'        End If
'
'        If sdb_prod_wgt.Value < sdb_prod_wgt_min.Value Or sdb_prod_wgt.Value > sdb_prod_wgt_max.Value Then
'           Call Gp_MsgBoxDisplay("产品单重必须在上下限之间", "I")
'           Exit Sub
'        End If
'
'    End If
'
'    If txt_prod_cd.Text = "HC" Then
'       If sdb_prod_wgt_min.Value >= sdb_prod_wgt_max.Value Then Call Gp_MsgBoxDisplay("产品单重上限不可小于下限", "I"): Exit Sub
'
'       If sdb_pack_wgt_min.Value >= sdb_pack_wgt_max.Value Then Call Gp_MsgBoxDisplay("包装重量上限不可小于下限", "I"): Exit Sub
'
'       If sdb_prod_wgt_min.Value > sdb_pack_wgt_min.Value Or sdb_prod_wgt_max.Value > sdb_pack_wgt_max.Value Then
'          Call Gp_MsgBoxDisplay("包装重量上下限必须大于等于产品单重上限", "I")
'          Exit Sub
'       End If
'       If txt_pack_way.Text = "NO" Then
'          If sdb_pack_wgt_min.Value <> sdb_prod_wgt_min.Value Or sdb_pack_wgt_max.Value <> sdb_prod_wgt_max.Value Then
'             Call Gp_MsgBoxDisplay("未包装时，包装重量上下限必须等于产品单重上下限", "I")
'             Exit Sub
'          End If
'
'       End If
'
'
'    End If
'
'    If Left(dtp_cust_del_fr_date.RawData, 8) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL") Or Left(dtp_cust_del_fr_date.RawData, 8) > Left(dtp_cust_del_to_date.RawData, 8) Then
'        Call Gp_MsgBoxDisplay("客户要求交货期有错误", "I")
'        Exit Sub
'    End If
'
'    If txt_transp_way.Text = "2" And sdb_trans_prc.Value = 0 Then
'        Call Gp_MsgBoxDisplay("运输方式为铁运，必须输入运费", "I")
'        Exit Sub
'    End If
'
'    If txt_ord_knd.Text = "A" Or txt_ord_knd.Text = "S" Then
'       If sdb_prod_prc.Value = 0 Then
'          Call Gp_MsgBoxDisplay("产品单价必须输入", "I")
'          Exit Sub
'       End If
'    End If
'
'    If txt_ord_knd.Text = "P" Or txt_ord_knd.Text = "S" Then
'       If txt_cust_req_plant.Text = "**" Then
'          Call Gp_MsgBoxDisplay("计划订单，库存订单 必须输入工厂", "I")
'          Exit Sub
'       End If
'    End If


'    If sdb_del_tol_min.Value > sdb_del_tol_max.Value Then Call Gp_MsgBoxDisplay("交付公差上限不可小于下限", "I"): Exit Sub
'
'    Dim sQuery1, sQuery2 As String
'    If txt_prod_dgr <= "5" And TXT_SIZE_KND <> "08" And TXT_SIZE_KND <> "06" Then
'        sQuery1 = "{CALL ABA1020C.P_MASTER_CHECK ('" + txt_prod_cd.Text + "','" + txt_stdspec.Text + "','" + Trim(sdb_ord_thk.Value) + "','"
'
'        sQuery1 = sQuery1 + Trim(sdb_ord_wid.Value) + "','" + Trim(sdb_ord_len.Value) + "','" + Trim(sdb_tot_wgt.Value) + "','" + Trim(sdb_prod_wgt.Value) + "',?,?,?)}"
'
'        Dim OutParam(3, 4) As Variant
'
'        OutParam(1, 1) = "arg_e_wgt"
'        OutParam(1, 2) = adVarNumeric
'        OutParam(1, 3) = adParamOutput
'        OutParam(1, 4) = 256
'
'        'Return Error Code Parameter
'        OutParam(2, 1) = "arg_e_code"
'        OutParam(2, 2) = adInteger
'        OutParam(2, 3) = adParamOutput
'        OutParam(2, 4) = 1
'
'        'Return Error Messsage Parameter
'        OutParam(3, 1) = "arg_e_msg"
'        OutParam(3, 2) = adVarChar
'        OutParam(3, 3) = adParamOutput
'        OutParam(3, 4) = 256
'
'        Dim ret_Result_ErrCode As Integer
'        Dim ret_Result_ErrMsg As String
'        Dim adoCmd As ADODB.Command
'
'        'Ado Setting
'        M_CN1.CursorLocation = adUseServer
'        Set adoCmd = New ADODB.Command
'
'        adoCmd.CommandType = adCmdText
'        Set adoCmd.ActiveConnection = M_CN1
'
'        adoCmd.CommandText = sQuery1
'
'        adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
'        adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
'        adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(3, 1), OutParam(3, 2), OutParam(3, 3), OutParam(3, 4))
'
'        adoCmd.Execute , , adExecuteNoRecords
'
'        'Process Error Check
'        If adoCmd("arg_e_code") <> "0" Then
'
'            ret_Result_ErrCode = adoCmd("arg_e_code")
'            ret_Result_ErrMsg = adoCmd("arg_e_msg")
'
'            sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
'
'            Call Gp_MsgBoxDisplay(sErrMessg)
'
'            Set adoCmd = Nothing
'
'
'            Exit Sub
'        Else
'
'            sdb_prod_wgt.Value = adoCmd("arg_e_wgt")
'
'        End If
'
'    End If
'
'    Set adoCmd = Nothing
'
'
'    If Gf_Mc_Authority(sAuthority, Mc1) Then
'        txt_sale_emp_id.Text = sUserID
'
'        If Mc1.Item("pControl")(1).Enabled Then
'
'            sQuery = "{call ABA1020C.P_ORD_SEQ ( '" + txt_ord_no.Text + "' )}"
'
'            txt_ord_item.Text = Gf_CodeFind(M_CN1, sQuery)
'            txt_ord_sts.Text = "A"
'            Call txt_ord_sts_KeyUp(0, 0)
'        Else
'            If txt_ord_sts.Text <> "A" Then
'                sMesg = "不能修改状态不是'A'的订单"
'                Call Gp_MsgBoxDisplay(sMesg)
'                Exit Sub
'            End If
'        End If
'
'        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
'            Call Popup_Menu_Setting
'        End If
'
'    End If
    
End Sub

Public Sub Form_Del()

    Dim sMesg As String
    
    If txt_ord_sts.Text <> "A" Then
        sMesg = "不能删除状态不是'A'的订单"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If txt_ord_item.Text <> Gf_CodeFind(M_CN1, "SELECT MAX(ORD_ITEM) FROM BP_ORDER_ITEM WHERE ORD_NO = '" + txt_ord_no.Text + "'") Then
        sMesg = "Can delete last order item"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If Not Gf_Ms_Del(M_CN1, Mc1) Then
        Call Popup_Menu_Setting
    Else
        'Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If
    
End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "Clear"              'Clear
            Call Form_Cls
        Case "Save"               'Process
            Call Form_Pro
        Case "Delete"             'Delete
            Call Form_Del
        Case "Copy"               'Copy
            Call Master_Cpy
        Case "Paste"              'Paste
            Call Master_Pst
        Case "Exit"               'Exit
            Call Form_Exit
    End Select

End Sub

Public Sub Popup_Menu_Setting()

    Select Case Mid(sAuthority, 2, 3)

        Case "000"      'No Authority
            MenuTool.Buttons(3).Enabled = False                     'Save
            MenuTool.Buttons(4).Enabled = False                     'Delete
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste
        
        Case "001"      'Delete Authority
            MenuTool.Buttons(3).Enabled = False                     'Save
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste

        Case "010"      'Update Authority
            MenuTool.Buttons(4).Enabled = False                     'Delete
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste

        Case "011"      'Update, Delete Authority
            MenuTool.Buttons(6).Enabled = False                     'Copy
            MenuTool.Buttons(7).Enabled = False                     'Paste

        Case "100"      'Insert Authority
            MenuTool.Buttons(4).Enabled = False                     'Delete

        Case "101"      'Insert, Delete Authority

        Case "110"      'Insert, Update Authority
            MenuTool.Buttons(4).Enabled = False                     'Delete

        Case "111"      'Insert, Update, Delete Authority

    End Select
    
End Sub


Private Sub SCmd_COPY_Click()

   Load ABZ1010C
   ABZ1010C.txt_prod_cd.Text = txt_prod_cd.Text
   ABZ1010C.Show 1

End Sub

Private Sub sdb_ord_len_LostFocus()

   Dim sMesg As String
   
   If txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text = "02" Then
   
      If sdb_ord_LEN_MIN.Value <> 0 Then
      
         If sdb_ord_len.Value < sdb_ord_LEN_MIN.Value Then
                  sMesg = "长度范围错误：不能小于" + str(sdb_ord_LEN_MIN.Value)
                  Call Gp_MsgBoxDisplay(sMesg)
                  sdb_ord_len.Value = sdb_ord_LEN_MIN.Value
         End If
         
      End If
      
      If sdb_ord_LEN_MAX.Value <> 0 Then
      
         If sdb_ord_len.Value > sdb_ord_LEN_MAX.Value Then
         
                  sMesg = "长度范围错误：不能大于" + str(sdb_ord_LEN_MAX.Value)
                  Call Gp_MsgBoxDisplay(sMesg)
                  sdb_ord_len.Value = sdb_ord_LEN_MAX.Value
         End If
         
      End If
      
   End If
   
End Sub


Private Sub sdb_ord_LEN_MAX_Change()

'If txt_ord_size.Text <> "" Then
'    Call txt_ord_size_LostFocus
'End If

End Sub

Private Sub sdb_ord_LEN_MAX_LostFocus()

    Dim sMesg As String
    Dim ORD_LEN_MIN_T As Integer
    
    If MSG = False Then
       Exit Sub
    End If

    If TXT_SIZE_KND.Text = "01" And sdb_ord_LEN_MAX.Value > 0 Then
          sMesg = "定尺没有长度上限 "
          Call Gp_MsgBoxDisplay(sMesg)
          sdb_ord_LEN_MAX.SetFocus
    End If
    
    If txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text = "02" And sdb_ord_LEN_MAX.Value > 22000 Then
          sMesg = "单定尺长度在22000以下"
          Call Gp_MsgBoxDisplay(sMesg)
          sdb_ord_LEN_MAX.SetFocus
    End If
    
    
'    If sdb_ord_LEN_MIN.Value <> 0 And sdb_ord_LEN_MAX.Value <> 0 Then
'       If sdb_ord_LEN_MAX.Value - sdb_ord_LEN_MIN.Value < 1000 Then
'          sMesg = "长度范围错误：不能小于1000"
'          Call Gp_MsgBoxDisplay(sMesg)
'          sdb_ord_LEN_MAX.SetFocus
'       End If
'    End If

    
    If TXT_SIZE_KND.Text = "02" Then
    
       If sdb_ord_LEN_MAX.Value <> 0 Then
       
          If sdb_ord_LEN_MIN.Value = 4000 And sdb_ord_LEN_MAX.Value = 22000 Then
               If Trim(sdb_ord_len.Value) = "" Or sdb_ord_len.Value = 0 Then sdb_ord_len.Value = 12000
          Else
               ORD_LEN_MIN_T = Int((sdb_ord_LEN_MIN.Value + 999) / 1000)
               sdb_ord_len.Value = ORD_LEN_MIN_T * 1000
          End If
          
          If sdb_ord_len.Value <> 0 Then
            If sdb_ord_LEN_MAX.Value < sdb_ord_len.Value Then
                     sMesg = "长度范围错误：不能小于" + str(sdb_ord_len.Value)
                     Call Gp_MsgBoxDisplay(sMesg)
                     sdb_ord_LEN_MAX.SetFocus
                     sdb_ord_LEN_MAX.Value = sdb_ord_len.Value
                     Exit Sub
            End If
          End If
       End If
    End If

' Call txt_ord_size_LostFocus

End Sub

Private Sub sdb_ord_LEN_MIN_Change()

'If txt_ord_size.Text <> "" Then
'    Call txt_ord_size_LostFocus
'End If

End Sub

Private Sub sdb_ord_LEN_MIN_LostFocus()

    Dim sMesg As String
    Dim ORD_LEN_MIN_T As Integer
    
    If MSG = False Then
       Exit Sub
    End If
    
    If TXT_SIZE_KND.Text = "01" And sdb_ord_LEN_MIN.Value > 0 Then
          sMesg = "定尺没有长度下限 "
          Call Gp_MsgBoxDisplay(sMesg)
          sdb_ord_LEN_MIN.SetFocus
    End If


    If txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text = "02" And sdb_ord_LEN_MIN.Value < 4000 Then
          sMesg = "单定尺长度在4000以上"
          Call Gp_MsgBoxDisplay(sMesg)
    '      Exit Sub
          sdb_ord_LEN_MIN.SetFocus
    End If
    
    
'    If sdb_ord_LEN_MIN.Value <> 0 And sdb_ord_LEN_MAX.Value <> 0 Then
'       If sdb_ord_LEN_MAX.Value - sdb_ord_LEN_MIN.Value < 1000 Then
'          sMesg = "长度范围错误：不能小于1000"
'          Call Gp_MsgBoxDisplay(sMesg)
'          sdb_ord_LEN_MIN.SetFocus
'       End If
'    End If
    
    
    If TXT_SIZE_KND.Text = "02" Then
    
       If sdb_ord_LEN_MIN.Value <> 0 Then
       
          If sdb_ord_LEN_MIN.Value = 4000 And sdb_ord_LEN_MAX.Value = 22000 Then
               If Trim(sdb_ord_len.Value) = "" Or sdb_ord_len.Value = 0 Then sdb_ord_len.Value = 12000
          Else
               ORD_LEN_MIN_T = Int((sdb_ord_LEN_MIN.Value + 999) / 1000)
               sdb_ord_len.Value = ORD_LEN_MIN_T * 1000
          End If
       
          If sdb_ord_len.Value <> 0 Then
            If sdb_ord_LEN_MIN.Value > sdb_ord_len.Value Then
                     sMesg = "长度范围错误：不能大于" + str(sdb_ord_len.Value)
                     Call Gp_MsgBoxDisplay(sMesg)
                     sdb_ord_LEN_MIN.SetFocus
                     sdb_ord_LEN_MIN.Value = sdb_ord_len.Value
                     Exit Sub
            End If
          End If
       End If
    End If

' Call txt_ord_size_LostFocus
 
End Sub

Private Sub sdb_ord_thk_Change()

    sdb_mb_thk.Value = sdb_ord_thk.Value
    
End Sub

Private Sub sdb_prod_wgt_max_LostFocus()
'    If txt_pack_way.Text = "NO" Then
'       sdb_pack_wgt_max.Value = sdb_prod_wgt_max.Value
'    End If
    If sdb_prod_wgt_max.Value < sdb_prod_wgt.Value Then
       Call Gp_MsgBoxDisplay("产品单重上限值必须大于等于产品单重")
    End If
       
'    If txt_pack_way.Text <> "" Then
'       Call txt_pack_way_LostFocus
'    End If
    
End Sub


Private Sub sdb_prod_wgt_min_LostFocus()

'    If txt_pack_way.Text = "NO" Then
'       sdb_pack_wgt_min.Value = sdb_prod_wgt_min.Value
'    End If

End Sub


Private Sub SSCommand1_Click()

End Sub

'Private Sub txt_currency_DblClick()
'
'    Call txt_currency_KeyUp(vbKeyF4, 0)
'
'End Sub

Private Sub txt_cust_req_plant_DblClick()

    Call txt_cust_req_plant_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_cust_req_plant_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_cust_req_plant
        DD.rControl.Add Item:=txt_cust_req_plant_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_req_plant)) = txt_cust_req_plant.MaxLength Then
        txt_cust_req_plant_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_cust_req_plant.Text), 2)
    Else
        txt_cust_req_plant_name.Text = ""
    End If

End Sub

Private Sub txt_cust_spec_no_DblClick()

    Call txt_cust_spec_no_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_cust_spec_no_KeyUp(KeyCode As Integer, Shift As Integer)

'    If KeyCode = vbKeyF4 Then
'
'       Load ABX1110C
'
'       ABX1110C.txt_form_nm.Text = "ABA1020C"
'       ABX1110C.txt_cust_cd.Text = txt_cust_cd.Text
'       ABX1110C.txt_prod_cd.Text = txt_prod_cd.Text
'
'       ABX1110C.Show 1
'
'    End If
    
End Sub

'Private Sub txt_del_cond_DblClick()
'
'    Call txt_del_cond_KeyUp(vbKeyF4, 0)
'
'End Sub

Private Sub txt_del_tol_unit_DblClick()

    Call txt_del_tol_unit_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_dest_detail_KeyUp(KeyCode As Integer, Shift As Integer)

' If KeyCode = vbKeyF4 Then
'
'       Load ABX1120C
'
'       ABX1120C.txt_form_nm.Text = "ABA1021C"
'       ABX1120C.txt_dest_cd.Text = txt_dest_cd.Text
'
'       ABX1120C.Show 1
'
'    End If

End Sub


Private Sub txt_end_cust_cd_DblClick()

    Call txt_end_cust_cd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_enduse_cd_DblClick()

    Call txt_enduse_cd_KeyUp(vbKeyF4, 0)

End Sub





Private Sub TXT_HTM_MET1_DblClick()

    Call TXT_HTM_MET1_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_HTM_MET1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=TXT_HTM_MET1
        DD.rControl.Add Item:=TXT_HTM_MET1_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_HTM_MET1)) = TXT_HTM_MET1.MaxLength Then
        TXT_HTM_MET1_nm.Text = Gf_ComnNameFind(M_CN1, "Q0073", Trim(TXT_HTM_MET1.Text), 2)
    Else
        TXT_HTM_MET1_nm.Text = ""
    End If
    
End Sub
Private Sub TXT_HTM_MET2_DblClick()

    Call TXT_HTM_MET2_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_HTM_MET2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=TXT_HTM_MET2
        DD.rControl.Add Item:=TXT_HTM_MET2_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_HTM_MET2)) = TXT_HTM_MET2.MaxLength Then
        TXT_HTM_MET2_nm.Text = Gf_ComnNameFind(M_CN1, "Q0073", Trim(TXT_HTM_MET2.Text), 2)
    Else
        TXT_HTM_MET2_nm.Text = ""
    End If
    
End Sub
Private Sub TXT_HTM_MET3_DblClick()

    Call TXT_HTM_MET3_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_HTM_MET3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=TXT_HTM_MET3
        DD.rControl.Add Item:=TXT_HTM_MET3_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_HTM_MET3)) = TXT_HTM_MET3.MaxLength Then
        TXT_HTM_MET3_nm.Text = Gf_ComnNameFind(M_CN1, "Q0073", Trim(TXT_HTM_MET3.Text), 2)
    Else
        TXT_HTM_MET3_nm.Text = ""
    End If
    
End Sub

Private Sub txt_insp_cd_DblClick()

    Call txt_insp_cd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_marking_way_DblClick()

    Call txt_marking_way_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_ord_cust_cd_DblClick()

    Call txt_ord_cust_cd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_ord_size_LostFocus()

    Dim i As Integer
    Dim T As Double
    Dim W As Double
    Dim l As Double
    Dim N1 As Long
    Dim N2 As Long
    Dim N3 As Long
    Dim Num As Long
    Dim Tot_size As Long
    Dim ORD_LEN_MIN As Long
    Dim ORD_LEN_MAX As Long
    Dim ORD_LEN_TGT As Long
    Dim ORD_LEN_MIN_T As Integer

    Dim T_VALUE  As String
    Dim W_VALUE  As String
    Dim L_VALUE  As String
    Dim T_SIZE As Integer
    Dim W_SIZE As Integer
    Dim L_SIZE As Integer

    MSG = True

    If txt_ord_size.Text <> "" Then
       txt_ord_size.Text = Replace(txt_ord_size.Text, " ", "")
       N1 = InStr(1, txt_ord_size.Text, "*")
       N2 = InStr(N1 + 1, txt_ord_size.Text, "*")
       N3 = InStr(N2 + 1, txt_ord_size.Text, "*")
       Tot_size = Len(Trim(txt_ord_size.Text))

       If N1 = 0 Or N2 = 0 Then
          Call Gp_MsgBoxDisplay("尺寸不完整", "I")
          MSG = False
          txt_ord_size.SetFocus
          Exit Sub
       End If

       If N3 > 0 Then
          Call Gp_MsgBoxDisplay("输入错误, '*' 两个 以上 不可能 ", "I")
          MSG = False
          txt_ord_size.SetFocus
          Exit Sub
       End If

       For i = 1 To Tot_size
            If (Mid(txt_ord_size.Text, i, 1) >= "0" And Mid(Trim(txt_ord_size.Text), i, 1) <= "9") _
                Or (Mid(txt_ord_size.Text, i, 1) = "*") _
                Or (Mid(txt_ord_size.Text, i, 1) = ".") _
                Or (Mid(txt_ord_size.Text, i, 1) = "C") _
                Or (Mid(txt_ord_size.Text, i, 1) = "L") _
                Or (Mid(txt_ord_size.Text, i, 1) = "W") _
                Or (Mid(txt_ord_size.Text, i, 1) = Chr(34)) _
                Or (Mid(txt_ord_size.Text, i, 1) = Chr(39)) Then
            Else
                Call Gp_MsgBoxDisplay(" Number, . , * , C , ' ,""""  以外 不可能 ", "I")
                MSG = False
                txt_ord_size.SetFocus
                Exit Sub
            End If
       Next i

       T_VALUE = Mid(Trim(txt_ord_size.Text), 1, N1 - 1)
       T_SIZE = Len(Trim(T_VALUE))
       W_VALUE = Mid(Trim(txt_ord_size.Text), N1 + 1, N2 - N1 - 1)
       W_SIZE = Len(Trim(W_VALUE))
       L_VALUE = Mid(Trim(txt_ord_size.Text), N2 + 1, Tot_size - N2)
       L_SIZE = Len(Trim(L_VALUE))

       If txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text = "02" Then
           If sdb_ord_LEN_MIN.Value = 0 Then
               sdb_ord_LEN_MIN.Value = 4000
           End If

           If sdb_ord_LEN_MAX.Value = 0 Then
               sdb_ord_LEN_MAX.Value = 22000
           End If

           ORD_LEN_MIN = sdb_ord_LEN_MIN.Value
           ORD_LEN_MAX = sdb_ord_LEN_MAX.Value
'2007.11.18 HYS INSERT START
       ElseIf txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text = "06" Then
           If sdb_ord_LEN_MIN.Value = 0 Then
               sdb_ord_LEN_MIN.Value = 4000
           End If

           If sdb_ord_LEN_MAX.Value = 0 Then
               sdb_ord_LEN_MAX.Value = 5999.9
           End If

           ORD_LEN_MIN = sdb_ord_LEN_MIN.Value
           ORD_LEN_MAX = sdb_ord_LEN_MAX.Value
'2007.11.18 HYS INSERT END
       End If

       If TXT_SIZE_KND.Text = "08" Then
               sdb_ord_LEN_MIN.Value = 0
               sdb_ord_LEN_MAX.Value = 3999.9

           ORD_LEN_MIN = sdb_ord_LEN_MIN.Value
           ORD_LEN_MAX = sdb_ord_LEN_MAX.Value

       End If

' THICKNESS VALUE CHECK
       For i = 1 To T_SIZE
            If (Mid(T_VALUE, i, 1) >= "0" And Mid(Trim(T_VALUE), i, 1) <= "9") _
                Or (Mid(T_VALUE, i, 1) = ".") _
                Or (Mid(T_VALUE, i, 1) = Chr(34)) _
                Or (Mid(T_VALUE, i, 1) = Chr(39)) Then
            Else
                Call Gp_MsgBoxDisplay("THICKNESS DATA : 0~9 , . , '，""""  以外 不可能 ", "I")
                MSG = False
                txt_ord_size.SetFocus
                Exit Sub
            End If
       Next i

' WIDTH VALUE CHECK
       If TXT_SIZE_KND.Text <> "08" And TXT_SIZE_KND.Text <> "06" Then

           For i = 1 To W_SIZE
                If (Mid(W_VALUE, i, 1) >= "0" And Mid(Trim(W_VALUE), i, 1) <= "9") _
                    Or (Mid(W_VALUE, i, 1) = ".") _
                    Or (Mid(L_VALUE, i, 1) = "W") _
                    Or (Mid(W_VALUE, i, 1) = Chr(34)) _
                    Or (Mid(W_VALUE, i, 1) = Chr(39)) Then
                Else
                    Call Gp_MsgBoxDisplay(" WIDTH DATA : 0~9 , . , ,',"""" 以外 不可能  ", "I")
                    MSG = False
                    txt_ord_size.SetFocus
                    Exit Sub
                End If
           Next i
       End If

' LENGTH VALUE CHECK
       If txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text <> "02" And TXT_SIZE_KND.Text <> "08" And TXT_SIZE_KND.Text <> "06" Then
            For i = 1 To L_SIZE
                If (Mid(L_VALUE, i, 1) >= "0" And Mid(Trim(L_VALUE), i, 1) <= "9") _
                    Or (Mid(L_VALUE, i, 1) = ".") _
                    Or (Mid(L_VALUE, i, 1) = "C") _
                    Or (Mid(L_VALUE, i, 1) = "L") _
                    Or (Mid(L_VALUE, i, 1) = Chr(34)) _
                    Or (Mid(L_VALUE, i, 1) = Chr(39)) Then
                Else
                     Call Gp_MsgBoxDisplay(" LENGTH DATA : 0~9 , . , '  以外 不可能  ", "I")
                     MSG = False
                     txt_ord_size.SetFocus
                     Exit Sub
                End If
            Next i
       End If

' THICKNESS
       If InStr(1, T_VALUE, Chr(34)) = 0 And InStr(1, T_VALUE, Chr(39)) = 0 Then
          If IsNumeric(T_VALUE) = False Then
             Call Gp_MsgBoxDisplay(" THICKNESS DATA : NUMBER 以外 不可能  ", "I")
             MSG = False
             txt_ord_size.SetFocus
             Exit Sub
          Else
             T = Val(T_VALUE)
          End If

       Else
          If IsNumeric(Mid(T_VALUE, 1, T_SIZE - 1)) = False Then
             Call Gp_MsgBoxDisplay(" THICKNESS DATA : NUMBER 以外 不可能  ", "I")
             MSG = False
             txt_ord_size.SetFocus
             Exit Sub
          End If

          If Mid(T_VALUE, T_SIZE, 1) = Chr(39) Then
             T = Val(Mid(T_VALUE, 1, T_SIZE - 1)) * 2.54
          Else
             T = Val(Mid(T_VALUE, 1, T_SIZE - 1)) * 304.8
          End If
       End If

' WIDTH

       If TXT_SIZE_KND = "08" Then 'Or TXT_SIZE_KND.Text = "06" Then
             If W_VALUE = "W" Then
                W = 0
             Else
                Call Gp_MsgBoxDisplay("SHORT SIZE : WIDTH DATA ERROR", "I")
                MSG = False
                txt_ord_size.SetFocus
                Exit Sub
             End If
       Else
           If InStr(1, W_VALUE, Chr(34)) = 0 And InStr(1, W_VALUE, Chr(39)) = 0 Then
              If IsNumeric(W_VALUE) = False Then
                 Call Gp_MsgBoxDisplay(" WIDTH DATA : NUMBER 以外 不可能  ", "I")
                 MSG = False
                 txt_ord_size.SetFocus
                 Exit Sub
              Else
                 W = Val(W_VALUE)
              End If
           Else
              If IsNumeric(Mid(W_VALUE, 1, W_SIZE - 1)) = False Then
                 Call Gp_MsgBoxDisplay(" WIDTH DATA : NUMBER 以外 不可能  ", "I")
                 MSG = False
                 txt_ord_size.SetFocus
                 Exit Sub
              End If

              If Mid(W_VALUE, W_SIZE, 1) = Chr(39) Then
                 W = Val(Mid(W_VALUE, 1, W_SIZE - 1)) * 2.54
              Else
                 W = Val(Mid(W_VALUE, 1, W_SIZE - 1)) * 304.8
              End If
           End If
       End If



' LENGTH
       If txt_prod_cd.Text = "HC" Then
             If L_VALUE = "C" Then
                l = 0
             Else
                Call Gp_MsgBoxDisplay("输入错误,钢卷没有长度，长度用C表示", "I")
                MSG = False
                txt_ord_size.SetFocus
                Exit Sub
             End If
       ElseIf txt_prod_cd.Text = "PP" And (TXT_SIZE_KND.Text = "02" Or TXT_SIZE_KND.Text = "06" Or TXT_SIZE_KND.Text = "08") Then
             If TXT_SIZE_KND.Text = "02" Then
                  If L_VALUE = "L" Then
                     If ORD_LEN_MIN = 4000 And ORD_LEN_MAX = 22000 Then
                        If Trim(sdb_ord_len.Value) = "" Or sdb_ord_len.Value = 0 Then sdb_ord_len.Value = 12000
                        l = sdb_ord_len.Value
                     Else
                        ORD_LEN_MIN_T = Int((ORD_LEN_MIN + 999) / 1000)
                        If Trim(sdb_ord_len.Value) = "" Or sdb_ord_len.Value = 0 Then sdb_ord_len.Value = ORD_LEN_MIN_T * 1000
                        l = sdb_ord_len.Value
                     End If
'                        sdb_ord_len.Enabled = True
                  Else
                     Call Gp_MsgBoxDisplay("输入错误,单定尺长度为L", "I")
                     MSG = False
                     txt_ord_size.SetFocus
                     Exit Sub
                  End If
'2007.11.18 HYS INSERT START
             ElseIf TXT_SIZE_KND.Text = "06" Then
                  If L_VALUE = "L" Then
                      ORD_LEN_MIN_T = Int((ORD_LEN_MIN) / 1000)
                      If Trim(sdb_ord_len.Value) = "" Or sdb_ord_len.Value = 0 Then sdb_ord_len.Value = ORD_LEN_MIN_T * 1000
                      l = sdb_ord_len.Value
'                        sdb_ord_len.Enabled = True
                  Else
                     Call Gp_MsgBoxDisplay("输入错误,非尺长度为L", "I")
                     MSG = False
                     txt_ord_size.SetFocus
                     Exit Sub
                  End If
'2007.11.18 HYS INSERT END
             ElseIf TXT_SIZE_KND.Text = "08" Then
                  If L_VALUE = "L" Then
                     If Trim(sdb_ord_len.Value) = "" Or sdb_ord_len.Value = 0 Then sdb_ord_len.Value = ORD_LEN_MAX
                     l = sdb_ord_len.Value
                  Else
                     Call Gp_MsgBoxDisplay("输入错误", "I")
                     MSG = False
                     txt_ord_size.SetFocus
                     Exit Sub
                  End If
             End If
       Else
            If InStr(1, L_VALUE, Chr(34)) = 0 And InStr(1, L_VALUE, Chr(39)) = 0 Then
               If IsNumeric(L_VALUE) = False Then
                  Call Gp_MsgBoxDisplay(" LENGTH DATA : NUMBER 以外 不可能 ", "I")
                  MSG = False
                  txt_ord_size.SetFocus
                  Exit Sub
               Else
                  l = Val(L_VALUE)
               End If
            Else
               If IsNumeric(Mid(L_VALUE, 1, L_SIZE - 1)) = False Then
                  Call Gp_MsgBoxDisplay(" LENGTH DATA : NUMBER 以外 不可能 ", "I")
                  MSG = False
                  txt_ord_size.SetFocus
                  Exit Sub
               End If

               If Mid(L_VALUE, L_SIZE, 1) = Chr(39) Then
                  l = Val(Mid(L_VALUE, 1, L_SIZE - 1)) * 2.54
               Else
                  l = Val(Mid(L_VALUE, 1, L_SIZE - 1)) * 304.8
               End If
            End If
        End If
    Else
        Exit Sub
    End If

    sdb_ord_thk.Value = T
    sdb_ord_wid.Value = W

    If TXT_SIZE_KND = "08" Then
         sdb_ord_wid.Value = 0
         sdb_ord_len.Value = sdb_ord_LEN_MAX.Value
    End If

    If TXT_SIZE_KND = "06" Then
     '    sdb_ord_wid.Value = 0
         sdb_ord_len.Value = 5000
    End If

    If Trim(sdb_ord_len.Value) = "" Or sdb_ord_len.Value = 0 Or TXT_SIZE_KND = "01" Then sdb_ord_len.Value = l

'    If TXT_SIZE_KND = "02" Then
'       If ORD_LEN_MIN = 4000 And ORD_LEN_MAX = 22000 Then
'           sdb_ord_len.Value = 12000
'       Else
'           sdb_ord_len.Value = ORD_LEN_MIN
'       End If
'    ElseIf TXT_SIZE_KND = "08" Then
'           sdb_ord_len.Value = sdb_ord_LEN_MAX.Value
'    End If
End Sub


Private Sub txt_ord_sts_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0011"
        DD.rControl.Add Item:=txt_ord_sts
        DD.rControl.Add Item:=txt_ord_sts_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    
    If Len(Trim(txt_ord_sts)) = txt_ord_sts.MaxLength Then
        txt_ord_sts_name.Text = Gf_ComnNameFind(M_CN1, "B0011", Trim(txt_ord_sts.Text), 2)
    Else
        txt_ord_sts_name.Text = ""
    End If


End Sub
'
'Private Sub txt_pack_way_DblClick()
'
'    Call txt_pack_way_KeyUp(vbKeyF4, 0)
'
'End Sub

Private Sub txt_pack_way_GotFocus()

    If (sdb_prod_wgt_min = 0) Or (sdb_prod_wgt_min = 0) Then
    
       Call Gp_MsgBoxDisplay("请输入产品单重上下限", "I")
       
    End If

End Sub

'Private Sub txt_pack_way_LostFocus()
'
'    Dim sQuery As String
'    Dim AdoRs As ADODB.Recordset
'
'    If Len(txt_pack_way.Text) <> 0 Then
'
'        sQuery = "select CD from NISCO.ZP_CD  where CD_MANA_NO= 'B0025' "
'        sQuery = sQuery + " AND  CD='" + txt_pack_way.Text + "'"
'
'        Set AdoRs = New ADODB.Recordset
'        AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'        If AdoRs.BOF And AdoRs.EOF Then
'
'            Call Gp_MsgBoxDisplay("包装方式代码不存在.......")
'
'            txt_pack_way.Text = ""
'            txt_pack_way.SetFocus
'
'       End If
'
'    End If


'    If txt_prod_cd.Text = "HC" Then
'
'        If txt_pack_way.Text = "NO" Then
''           sdb_pack_wgt_min.Enabled = False
''           sdb_pack_wgt_max.Enabled = False
''           sdb_pack_wgt_min.BackColor = &HE0E0E0
''           sdb_pack_wgt_max.BackColor = &HE0E0E0
''           sdb_pack_wgt_min.Value = sdb_prod_wgt_min.Value
''           sdb_pack_wgt_max.Value = sdb_prod_wgt_max.Value
'        Else
'           sdb_pack_wgt_min.Enabled = True
'           sdb_pack_wgt_max.Enabled = True
'        End If
'    End If
'

'End Sub

'Private Sub txt_payment_cond_DblClick()
'
'    Call txt_payment_cond_KeyUp(vbKeyF4, 0)
'
'End Sub

Private Sub txt_sale_emp_id_DblClick()

    Call txt_sale_emp_id_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_sale_emp_id_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_sale_emp_id
        DD.rControl.Add Item:=txt_sale_emp_id_name

        DD.nameType = "1"

        Call Gf_EmpID_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_stamp)) = txt_stamp.MaxLength Then
        txt_sale_emp_id_name.Text = Gf_EmpNameFind(M_CN1, Trim(txt_sale_emp_id.Text))
    Else
        txt_sale_emp_id_name.Text = ""
    End If
    
End Sub


Private Sub TXT_shot_blast_DblClick()

    Call TXT_shot_blast_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_shot_blast_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0074"
        DD.rControl.Add Item:=TXT_shot_blast
        DD.rControl.Add Item:=TXT_shot_blast_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_shot_blast)) = TXT_shot_blast.MaxLength Then
        TXT_shot_blast_nm.Text = Gf_ComnNameFind(M_CN1, "Q0074", Trim(TXT_shot_blast.Text), 2)
    Else
        TXT_shot_blast_nm.Text = ""
    End If
    
End Sub

Private Sub TXT_SIZE_KND_Change()

    If Len(TXT_SIZE_KND.Text) = 2 Then
    
        If Dis_sw = False Then
          txt_ord_size.Text = ""
          sdb_ord_LEN_MIN.Value = 0
          sdb_ord_LEN_MAX.Value = 0
          sdb_ord_len.Value = 0
          sdb_ord_thk.Value = 0
          sdb_ord_wid.Value = 0
          sdb_mb_thk.Value = 0
        End If
       
       If txt_prod_cd.Text <> "HC" And TXT_SIZE_KND.Text = "02" Then
          sdb_ord_LEN_MAX.Enabled = True
          sdb_ord_LEN_MIN.Enabled = True
          sdb_ord_len.Enabled = True
       
       Else
          sdb_ord_LEN_MAX.Enabled = False
          sdb_ord_LEN_MIN.Enabled = False
          sdb_ord_len.Enabled = False
       End If
    
    
    End If

'Call txt_ord_size_LostFocus

End Sub

Private Sub TXT_SIZE_KND_DblClick()

    Call txt_size_knd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_SIZE_KND_LostFocus()
    
    If txt_prod_dgr.Text <> "5" And TXT_SIZE_KND.Text = "08" Then
    
       Call Gp_MsgBoxDisplay("产品等级不是次品时，不可能是短尺", "I")
       Exit Sub
       
    End If

End Sub

Private Sub txt_stamp_DblClick()

    Call txt_stamp_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_stamp_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0050"
        DD.rControl.Add Item:=txt_stamp
        DD.rControl.Add Item:=txt_stamp_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_stamp)) = txt_stamp.MaxLength Then
        txt_stamp_name.Text = Gf_ComnNameFind(M_CN1, "B0050", Trim(txt_stamp.Text), 2)
    Else
        txt_stamp_name.Text = ""
    End If

End Sub



Private Sub txt_ord_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_ord_cust_cd
        DD.rControl.Add Item:=txt_ord_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_ord_cust_cd)) = txt_ord_cust_cd.MaxLength Then
        txt_ord_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_ord_cust_cd.Text), 1)
    Else
        txt_ord_cust_cd_name.Text = ""
    End If

End Sub

Private Sub txt_end_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_end_cust_cd
        DD.rControl.Add Item:=txt_end_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_end_cust_cd)) = txt_end_cust_cd.MaxLength Then
        txt_end_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_end_cust_cd.Text), 1)
    Else
        txt_end_cust_cd_name.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_Validate(Cancel As Boolean)

    If txt_prod_cd.Text <> "HC" Then
    
       cbo_india.Enabled = False
       sdb_outdia.Enabled = False
       cbo_india.BackColor = &HE0E0E0
       sdb_outdia.BackColor = &HE0E0E0
    
    End If

End Sub

Private Sub txt_stdspec_Change()

    txt_stdspec_yy.Text = ""
    txt_enduse_cd.Text = ""

End Sub

Private Sub txt_stdspec_DblClick()

    Call txt_stdspec_KeyUp(vbKeyF4, 0)

End Sub

'Private Sub txt_stlgrd_DblClick()
'Call txt_stlgrd_KeyUp(vbKeyF4, 0)
'
'End Sub

'Private Sub txt_transp_way_Change()
'
'    If txt_transp_way.Text = "2" Then
'       sdb_trans_prc.BackColor = &HC0FFFF
'    Else
'       sdb_trans_prc.BackColor = &H80000005
'    End If
'
'End Sub

'Private Sub txt_transp_way_DblClick()
'
'    Call txt_transp_way_KeyUp(vbKeyF4, 0)
'
'End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_urgnt_fl_DblClick()

    Call txt_urgnt_fl_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_UST_FL_DblClick()

    Call txt_UST_FL_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_UST_FL_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0046"
        DD.rControl.Add Item:=txt_UST_FL
        DD.rControl.Add Item:=Txt_ust_fl_name
        

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_UST_FL)) = txt_UST_FL.MaxLength Then
        Txt_ust_fl_name.Text = Gf_ComnNameFind(M_CN1, "Q0046", Trim(txt_UST_FL.Text), 2)
    Else
        Txt_ust_fl_name.Text = ""
    End If
    
End Sub

Private Sub TXT_MATR_FL_DblClick()

    Call TXT_MATR_FL_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_MATR_FL_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0059"
        DD.rControl.Add Item:=TXT_MATR_FL
        DD.rControl.Add Item:=TXT_MATR_FL_NM
        

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_MATR_FL)) = TXT_MATR_FL.MaxLength Then
        TXT_MATR_FL_NM.Text = Gf_ComnNameFind(M_CN1, "Q0059", Trim(TXT_MATR_FL.Text), 2)
    Else
        TXT_MATR_FL_NM.Text = ""
    End If
    
End Sub

Private Sub txt_wgt_grp_DblClick()

    Call txt_wgt_grp_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_wgt_grp_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0018"
        DD.rControl.Add Item:=txt_wgt_grp
        DD.rControl.Add Item:=txt_wgt_grp_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_wgt_grp)) = txt_wgt_grp.MaxLength Then
        txt_wgt_grp_name.Text = Gf_ComnNameFind(M_CN1, "B0018", Trim(txt_wgt_grp.Text), 2)
    Else
        txt_wgt_grp_name.Text = ""
    End If

End Sub


'Private Sub txt_transp_way_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "B0020"
'        DD.rControl.Add Item:=txt_transp_way
'        DD.rControl.Add Item:=txt_transp_way_name
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'    If Len(Trim(txt_transp_way)) = txt_transp_way.MaxLength Then
'        txt_transp_way_name.Text = Gf_ComnNameFind(M_CN1, "B0020", Trim(txt_transp_way.Text), 2)
'    Else
'        txt_transp_way_name.Text = ""
'    End If
'
'End Sub

Private Sub txt_del_tol_unit_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0028"
        DD.rControl.Add Item:=txt_del_tol_unit
        DD.rControl.Add Item:=txt_del_tol_unit_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_del_tol_unit)) = txt_del_tol_unit.MaxLength Then
        txt_del_tol_unit_name.Text = Gf_ComnNameFind(M_CN1, "B0028", Trim(txt_del_tol_unit.Text), 2)
    Else
        txt_del_tol_unit_name.Text = ""
    End If

End Sub

Private Sub txt_wgt_unit_DblClick()

    Call txt_wgt_unit_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_wgt_unit_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0017"
        DD.rControl.Add Item:=txt_wgt_unit
        DD.rControl.Add Item:=txt_wgt_unit_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_wgt_unit)) = txt_wgt_unit.MaxLength Then
        txt_wgt_unit_name.Text = Gf_ComnNameFind(M_CN1, "B0017", Trim(txt_wgt_unit.Text), 2)
    Else
        txt_wgt_unit_name.Text = ""
    End If

End Sub



Private Sub txt_marking_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0027"
        DD.rControl.Add Item:=txt_marking_way
        DD.rControl.Add Item:=txt_marking_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_marking_way)) = txt_marking_way.MaxLength Then
        txt_marking_way_name.Text = Gf_ComnNameFind(M_CN1, "B0027", Trim(txt_marking_way.Text), 2)
    Else
        txt_marking_way_name.Text = ""
    End If

End Sub

'Private Sub txt_pack_way_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "B0025"
'        DD.rControl.Add Item:=txt_pack_way
'        DD.rControl.Add Item:=txt_pack_way_name
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'    If Len(Trim(txt_pack_way)) = txt_pack_way.MaxLength Then
'        txt_pack_way_name.Text = Gf_ComnNameFind(M_CN1, "B0025", Trim(txt_pack_way.Text), 2)
'    Else
'        txt_pack_way_name.Text = ""
'    End If
'
'End Sub


Private Sub txt_insp_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0026"
        DD.rControl.Add Item:=txt_insp_cd
        DD.rControl.Add Item:=txt_insp_cd_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_insp_cd)) = txt_insp_cd.MaxLength Then
        txt_insp_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0026", Trim(txt_insp_cd.Text), 2)
    Else
        txt_insp_cd_name.Text = ""
    End If

End Sub

Private Sub txt_dest_cd_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then

            DD.sWitch = "MS"
            DD.rControl.Add Item:=txt_dest_cd
            DD.rControl.Add Item:=txt_dest_cd_name

            DD.nameType = "1"

            Call Gf_Destination_DD(M_CN1, KeyCode)

            Exit Sub

    End If

    If Len(Trim(txt_dest_cd)) = txt_dest_cd.MaxLength Then
        txt_dest_cd_name.Text = Gf_DestNameFind(M_CN1, Trim(txt_dest_cd.Text), 1)
    Else
        txt_dest_cd_name.Text = ""
    End If
        
End Sub

'Private Sub txt_del_cond_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "B0016"
'        DD.rControl.Add Item:=txt_del_cond
'        DD.rControl.Add Item:=txt_del_cond_name
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'    If Len(Trim(txt_del_cond)) = txt_del_cond.MaxLength Then
'        txt_del_cond_name.Text = Gf_ComnNameFind(M_CN1, "B0016", Trim(txt_del_cond.Text), 2)
'    Else
'        txt_del_cond_name.Text = ""
'    End If

'End Sub
'Private Sub txt_payment_cond_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "B0015"
'        DD.rControl.Add Item:=txt_payment_cond
'        DD.rControl.Add Item:=txt_payment_cond_name
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'    If Len(Trim(txt_payment_cond)) = txt_payment_cond.MaxLength Then
'        txt_payment_cond_name.Text = Gf_ComnNameFind(M_CN1, "B0015", Trim(txt_payment_cond.Text), 2)
'    Else
'        txt_payment_cond_name.Text = ""
'    End If
'
'End Sub

Private Sub txt_urgnt_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0022"
        DD.rControl.Add Item:=txt_urgnt_fl
        DD.rControl.Add Item:=txt_urgnt_fl_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_urgnt_fl)) = txt_urgnt_fl.MaxLength Then
        txt_urgnt_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0022", Trim(txt_urgnt_fl.Text), 2)
    Else
        txt_urgnt_fl_name.Text = ""
    End If

End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"
        DD.rControl.Add Item:=txt_trim_fl
        DD.rControl.Add Item:=txt_trim_fl_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_trim_fl)) = txt_trim_fl.MaxLength Then
        txt_trim_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0021", Trim(txt_trim_fl.Text), 2)
    Else
        txt_trim_fl_name.Text = ""
    End If

End Sub
Private Sub txt_hold_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0012"
        DD.rControl.Add Item:=txt_hold_fl
        DD.rControl.Add Item:=txt_hold_fl_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_hold_fl)) = txt_hold_fl.MaxLength Then
        txt_hold_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0012", Trim(txt_hold_fl.Text), 2)
    Else
        txt_hold_fl_name.Text = ""
    End If

End Sub
Private Sub txt_payment_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0024"
        DD.rControl.Add Item:=txt_payment_fl
        DD.rControl.Add Item:=txt_payment_fl_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_payment_fl)) = txt_payment_fl.MaxLength Then
        txt_payment_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0024", Trim(txt_payment_fl.Text), 2)
    Else
        txt_payment_fl_name.Text = ""
    End If

End Sub



'Private Sub txt_currency_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "B0013"
'        DD.rControl.Add Item:=txt_currency
'        DD.rControl.Add Item:=txt_currency_name
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'    If Len(Trim(txt_currency)) = txt_currency.MaxLength Then
'        txt_currency_name.Text = Gf_ComnNameFind(M_CN1, "B0013", Trim(txt_currency.Text), 2)
'    Else
'        txt_currency_name.Text = ""
'    End If
'
'End Sub

'Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.rControl.Add Item:=txt_stlgrd
'        DD.rControl.Add Item:=txt_stlgrd_name
'
'        DD.nameType = "2"
'
'        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
'       txt_stlgrd.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
'    Else
'       txt_stlgrd_name.Text = ""
'    End If
'
'End Sub


Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec
        DD.rControl.Add Item:=txt_stdspec_yy
        DD.rControl.Add Item:=txt_stdspec_name

        Call Gf_StdSPEC_DD1(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Public Function Gf_StdSPEC_DD1(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    Dim STDSPEC_FL As String
    
    Dim icount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "T"        'StdSPEC Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If txt_cust_req_plant = "C1" Then
       STDSPEC_FL = "1"
    ElseIf txt_cust_req_plant = "C3" Then
       STDSPEC_FL = "2"
    Else
       STDSPEC_FL = "%"
    End If
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND STDSPEC_CHR_CD <> 'N' "
        DD.sWhere = DD.sWhere + "   AND (STDSPEC_CHR_CD = 'Y' OR STDSPEC_CHR_CD LIKE '" & STDSPEC_FL & "' ) "
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND STDSPEC_CHR_CD <> 'N' "
        DD.sWhere = DD.sWhere + "   AND (STDSPEC_CHR_CD = 'Y' OR STDSPEC_CHR_CD LIKE '" & STDSPEC_FL & "' ) "
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
   
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                sNew_Name = DD.sPname.Text
            End If
            
            DD.sPname.TabStop = True
            DD.sPname.SetFocus
            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
            DD.sPname.EditMode = True
            DD.sPname.TabStop = False
            
            If DD.sSelect Then
                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
            End If
            
        End If
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function



Private Sub txt_enduse_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

             ABX1050C.Show 1
             Exit Sub
    Else
             txt_enduse_cd.Text = ""
             Exit Sub
    End If
    
   If Trim(txt_enduse_cd.Text) <> "" Then
   
      txt_enduse_cd_name = Gf_UsageNameFind(M_CN1, Mid(Trim(txt_prod_cd), 1, 1), txt_enduse_cd.Text)
      
   End If
   
End Sub



Private Sub txt_dept_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Z0002"
        DD.rControl.Add Item:=txt_dept_cd
        DD.rControl.Add Item:=txt_dept_cd_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_dept_cd)) = txt_dept_cd.MaxLength Then
        txt_dept_cd_name.Text = Gf_ComnNameFind(M_CN1, "Z0002", Trim(txt_dept_cd.Text), 2)
    Else
        txt_dept_cd_name.Text = ""
    End If
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_cd_name.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_prod_cd)) = txt_prod_cd.MaxLength Then
        txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
    Else
        txt_prod_cd_name.Text = ""
    End If

End Sub

Private Sub txt_ord_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0009"
        DD.rControl.Add Item:=txt_ord_knd
        DD.rControl.Add Item:=txt_ord_knd_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_ord_knd)) = txt_ord_knd.MaxLength Then
        txt_ord_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0009", Trim(txt_ord_knd.Text), 2)
    Else
        txt_ord_knd_name.Text = ""
    End If

End Sub

Private Sub txt_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"
        DD.rControl.Add Item:=TXT_SIZE_KND
        DD.rControl.Add Item:=TXT_SIZE_NM

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_SIZE_KND)) = TXT_SIZE_KND.MaxLength Then
        TXT_SIZE_NM.Text = Gf_ComnNameFind(M_CN1, "B0043", Trim(TXT_SIZE_KND.Text), 2)
    Else
        TXT_SIZE_NM.Text = ""
    End If

End Sub



Private Sub txt_prod_dgr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0034"
        DD.rControl.Add Item:=txt_prod_dgr
        DD.rControl.Add Item:=txt_prod_dgr_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_prod_dgr)) = txt_prod_dgr.MaxLength Then
        txt_prod_dgr_name.Text = Gf_ComnNameFind(M_CN1, "Q0034", Trim(txt_prod_dgr.Text), 2)
    Else
        txt_prod_dgr_name.Text = ""
    End If

End Sub

Private Sub txt_sale_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0010"
        DD.rControl.Add Item:=txt_sale_way
        DD.rControl.Add Item:=txt_sale_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_sale_way)) = txt_sale_way.MaxLength Then
        txt_sale_way_name.Text = Gf_ComnNameFind(M_CN1, "B0010", Trim(txt_sale_way.Text), 2)
    Else
        txt_sale_way_name.Text = ""
    End If

End Sub

Private Sub txt_UST_FL_LostFocus()

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    If Len(txt_UST_FL.Text) <> 0 Then
        
        sQuery = "select CD from NISCO.ZP_CD  where CD_MANA_NO= 'Q0046' "
        sQuery = sQuery + " AND  CD='" + txt_UST_FL.Text + "'"
        
        Set AdoRs = New ADODB.Recordset
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        If AdoRs.BOF And AdoRs.EOF Then
       
            Call Gp_MsgBoxDisplay("探伤代码不存在.......")
            
            txt_UST_FL.Text = ""
            txt_UST_FL.SetFocus
            
       End If
       
    End If
    
End Sub

Private Sub TXT_MATR_FL_LostFocus()

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    If Len(TXT_MATR_FL.Text) <> 0 Then
        
        sQuery = "select CD from NISCO.ZP_CD  where CD_MANA_NO= 'Q0059' "
        sQuery = sQuery + " AND  CD='" + TXT_MATR_FL.Text + "'"
        
        Set AdoRs = New ADODB.Recordset
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        If AdoRs.BOF And AdoRs.EOF Then
       
            Call Gp_MsgBoxDisplay("力学性能代码不存在.......")
            
            TXT_MATR_FL.Text = ""
            TXT_MATR_FL.SetFocus
            
       End If
       
    End If
    
End Sub

Private Sub txt_trim_fl_LostFocus()

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    If Len(txt_trim_fl.Text) <> 0 Then
        
        sQuery = "select CD from NISCO.ZP_CD  where CD_MANA_NO= 'B0021' "
        sQuery = sQuery + " AND  CD='" + txt_trim_fl.Text + "'"
        
        Set AdoRs = New ADODB.Recordset
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        If AdoRs.BOF And AdoRs.EOF Then
       
            Call Gp_MsgBoxDisplay("切边代码不存在.......")
            
            txt_trim_fl.Text = ""
            txt_trim_fl.SetFocus
            
       End If
       
    End If
    
End Sub

