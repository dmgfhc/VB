VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACA1031C 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "订单进程详细查询_ACA1031C"
   ClientHeight    =   10890
   ClientLeft      =   225
   ClientTop       =   585
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   5250
      Top             =   795
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
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2115
      Index           =   0
      Left            =   135
      TabIndex        =   30
      Top             =   3885
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   3731
      _Version        =   196609
      Font3D          =   1
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
      ShadowStyle     =   1
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
         Left            =   3825
         TabIndex        =   111
         Tag             =   "钢种"
         Top             =   1710
         Width           =   1050
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
         Left            =   1695
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   35
         Tag             =   "Ord_Size"
         Top             =   255
         Width           =   3000
      End
      Begin VB.TextBox txt_TRIM_NAME 
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
         Left            =   1665
         TabIndex        =   110
         Tag             =   "钢种"
         Top             =   1710
         Width           =   825
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   18
         Left            =   120
         Top             =   240
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
            Size            =   9.76
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
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   2
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   19
         Left            =   120
         Top             =   600
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
            Size            =   9.76
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
         Width           =   1515
         _Version        =   262145
         _ExtentX        =   2672
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
         ReadOnly        =   -1  'True
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
         Index           =   20
         Left            =   120
         Top             =   960
         Width           =   1500
         _ExtentX        =   2646
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
      Begin CSTextLibCtl.sidbEdit sdb_ord_len 
         Height          =   315
         Left            =   1680
         TabIndex        =   38
         Tag             =   "Ord_Len"
         Top             =   1320
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
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   5
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   21
         Left            =   120
         Top             =   1320
         Width           =   1500
         _ExtentX        =   2646
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   120
         Top             =   1710
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   2595
         Top             =   1710
         Width           =   1170
         _ExtentX        =   2064
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
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2190
      Left            =   135
      TabIndex        =   5
      Top             =   1515
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   3863
      _Version        =   196609
      Font3D          =   1
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
      ShadowStyle     =   1
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
         Left            =   2325
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   91
         Tag             =   "Cust_Rea_Plant"
         Top             =   1710
         Width           =   2500
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   90
         Tag             =   "Cust_Rea_Plant"
         Top             =   1710
         Width           =   600
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   26
         Tag             =   "Hold_Fl"
         Top             =   1350
         Width           =   2500
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
         Left            =   7750
         Locked          =   -1  'True
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
         Left            =   6840
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   22
         Tag             =   "Dest_Cd"
         Top             =   1350
         Width           =   870
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   21
         Tag             =   "urgnt_fl"
         Top             =   990
         Width           =   2500
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
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   20
         Tag             =   "Urgnt_Fl"
         Top             =   990
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
         Left            =   7750
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   19
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
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   18
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
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   17
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   16
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
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
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   14
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
         Left            =   7750
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   13
         Tag             =   "Ord_Cust_Cd"
         Top             =   630
         Width           =   2175
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
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   12
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
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   11
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   10
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
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   9
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
         Left            =   7750
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Cust_Cd"
         Top             =   270
         Width           =   2175
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   7
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
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   4
         Left            =   5280
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   13
         Left            =   5280
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   12
         Left            =   5280
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   11
         Left            =   5280
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   14
         Left            =   120
         Top             =   1710
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "客户指定工厂"
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
         ForeColor       =   -2147483641
      End
      Begin CSTextLibCtl.sidbEdit sdb_prod_prc 
         Height          =   315
         Left            =   6840
         TabIndex        =   92
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
         NumIntDigits    =   8
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   10
         Left            =   5280
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_trans_prc 
         Height          =   315
         Left            =   11610
         TabIndex        =   93
         Tag             =   "Trans_Prc"
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
         NumIntDigits    =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   9
         Left            =   10050
         Top             =   1710
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "运费"
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
      Left            =   6810
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   3
      Tag             =   "Ord_Sts"
      Top             =   795
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
      Left            =   7455
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Ord_St_name"
      Top             =   795
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
      Left            =   3045
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "Ord_Item"
      Top             =   795
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
      Left            =   1725
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "Ord_No"
      Top             =   795
      Width           =   1275
   End
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   0
      Left            =   150
      Top             =   795
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
         Size            =   9.76
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
      Left            =   10080
      Top             =   795
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
         Size            =   9.76
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
      Left            =   9900
      TabIndex        =   4
      Text            =   "F4"
      Top             =   9945
      Visible         =   0   'False
      Width           =   285
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2805
      Index           =   1
      Left            =   90
      TabIndex        =   31
      Top             =   7770
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   4948
      _Version        =   196609
      Font3D          =   1
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
      ShadowStyle     =   1
      Begin VB.TextBox txt_can_date 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   1500
         Width           =   1545
      End
      Begin VB.TextBox txt_mod_date 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   720
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
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   72
         Tag             =   "Can_Emp_ID"
         Top             =   1890
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   71
         Tag             =   "Can_Emp_ID"
         Top             =   1890
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   70
         Tag             =   "Can_Fl"
         Top             =   1140
         Width           =   2500
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
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   69
         Tag             =   "Can_Fl"
         Top             =   1140
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
         Height          =   285
         Left            =   3225
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   68
         Tag             =   "Mod_Time"
         Top             =   735
         Width           =   1410
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   67
         Tag             =   "Can_Fl"
         Top             =   300
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
         Left            =   2310
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   66
         Tag             =   "Can_Fl"
         Top             =   300
         Width           =   2500
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   44
         Left            =   120
         Top             =   300
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   45
         Left            =   120
         Top             =   720
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
            Size            =   9.76
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
         Top             =   1140
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   47
         Left            =   120
         Top             =   1515
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
         Top             =   1890
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
      Height          =   1515
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   6120
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   2672
      _Version        =   196609
      Font3D          =   1
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
      ShadowStyle     =   1
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
         Left            =   11310
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   96
         Tag             =   "钢种"
         Top             =   1110
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
         Left            =   12705
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   95
         Tag             =   "STLGRD"
         Top             =   1110
         Width           =   2100
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
         Left            =   6825
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   65
         Tag             =   "Test Method"
         Top             =   1020
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   64
         Tag             =   "Test Method"
         Top             =   1020
         Width           =   1815
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
         Left            =   11910
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   63
         Tag             =   "Trim_Fl"
         Top             =   690
         Width           =   2500
      End
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
         Left            =   11310
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   62
         Tag             =   "Trim_Fl"
         Top             =   690
         Width           =   600
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
         Left            =   11310
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   61
         Tag             =   "Cust_Spec_No"
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_cust_spec_no_det 
         Height          =   310
         Left            =   12525
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   300
         Width           =   2340
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
         Left            =   2220
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   58
         Tag             =   "Enduse_Cd"
         Top             =   1050
         Width           =   2500
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
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   57
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
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   56
         Tag             =   "标准年度"
         Top             =   690
         Width           =   1500
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
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   55
         Tag             =   "标准代码"
         Top             =   300
         Width           =   1400
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
         Left            =   3045
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   54
         Tag             =   "STDSPEC"
         Top             =   300
         Width           =   2100
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   41
         Left            =   120
         Top             =   300
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "标准"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   42
         Left            =   120
         Top             =   690
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "标准年度"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   43
         Left            =   135
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   38
         Left            =   5280
         Top             =   300
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "内径"
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
      Begin CSTextLibCtl.sidbEdit sdb_outdia 
         Height          =   315
         Left            =   6840
         TabIndex        =   59
         Tag             =   "Ord_Len"
         Top             =   660
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
         ReadOnly        =   -1  'True
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
         Left            =   5280
         Top             =   660
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "外径"
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   35
         Left            =   9750
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   36
         Left            =   9750
         Top             =   690
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   40
         Left            =   5280
         Top             =   1020
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
            Size            =   9.76
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
         Left            =   9750
         Top             =   1110
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin CSTextLibCtl.sidbEdit cbo_india 
         Height          =   315
         Left            =   6840
         TabIndex        =   106
         Tag             =   "Ord_Len"
         Top             =   300
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
         ReadOnly        =   -1  'True
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
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2775
      Index           =   3
      Left            =   5190
      TabIndex        =   33
      Top             =   7785
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   4895
      _Version        =   196609
      Font3D          =   1
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
      ShadowStyle     =   1
      Begin VB.TextBox txt__CUST_del_to_date 
         Height          =   315
         Left            =   3295
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox txt_CUST_del_fr_date 
         Height          =   315
         Left            =   3295
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   1800
         Width           =   1545
      End
      Begin VB.TextBox txt_del_to_date 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox txt_del_fr_date 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   1800
         Width           =   1545
      End
      Begin VB.TextBox txt_dest_detail 
         Height          =   345
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   2310
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.TextBox txt_extra_fl 
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
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   88
         Tag             =   "Extra_Fl"
         Top             =   1380
         Width           =   600
      End
      Begin VB.TextBox txt_extra_fl_name 
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
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   87
         Tag             =   "Extra_Fl"
         Top             =   1380
         Width           =   2500
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
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   86
         Tag             =   "Payment_Fl"
         Top             =   990
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
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   85
         Tag             =   "urgnt_fl"
         Top             =   990
         Width           =   2500
      End
      Begin VB.TextBox txt_payment_cond 
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
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   84
         Tag             =   "Payment_Cond"
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox txt_payment_cond_name 
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
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   83
         Tag             =   "Payment_Cond"
         Top             =   630
         Width           =   2500
      End
      Begin VB.TextBox txt_currency_name 
         Height          =   310
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   240
         Width           =   2500
      End
      Begin VB.TextBox txt_currency 
         Height          =   310
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   81
         Top             =   240
         Width           =   600
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
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   80
         Tag             =   "Stamp"
         Top             =   1440
         Width           =   2500
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   79
         Tag             =   "Stamp"
         Top             =   1440
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
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   78
         Tag             =   "Marking_Way"
         Top             =   1050
         Width           =   2500
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   77
         Tag             =   "标识方式"
         Top             =   1050
         Width           =   600
      End
      Begin VB.TextBox txt_transp_way_name 
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
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   76
         Tag             =   "Transp_Way"
         Top             =   660
         Width           =   2500
      End
      Begin VB.TextBox txt_transp_way 
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   75
         Tag             =   "Transp_Way"
         Top             =   660
         Width           =   600
      End
      Begin VB.TextBox txt_del_cond_name 
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
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   74
         Tag             =   "Del_Cond"
         Top             =   270
         Width           =   2500
      End
      Begin VB.TextBox txt_del_cond 
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   73
         Tag             =   "Del_Cond"
         Top             =   270
         Width           =   600
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   49
         Left            =   120
         Top             =   270
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "交货条件"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   50
         Left            =   120
         Top             =   660
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "运输方式"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   51
         Left            =   120
         Top             =   1050
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   52
         Left            =   120
         Top             =   1440
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "是否标识"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   60
         Left            =   4980
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "货币种类"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   59
         Left            =   4980
         Top             =   630
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "结算条件"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   58
         Left            =   4980
         Top             =   990
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "资金入帐"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   57
         Left            =   4980
         Top             =   1380
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "基价/浮动价分类"
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
         ForeColor       =   -2147483641
      End
      Begin CSTextLibCtl.sidbEdit sdb_discon_prc 
         Height          =   315
         Left            =   6540
         TabIndex        =   89
         Tag             =   "Discon_Prc"
         Top             =   1830
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
         NumIntDigits    =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   56
         Left            =   4980
         Top             =   1830
         Width           =   1500
         _ExtentX        =   2646
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
         Top             =   1785
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "交货期开始"
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
         Top             =   2160
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "交货期结束"
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
         Index           =   55
         Left            =   4980
         Top             =   2325
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
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
      Height          =   2085
      Index           =   4
      Left            =   5205
      TabIndex        =   34
      Top             =   3900
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   3678
      _Version        =   196609
      Font3D          =   1
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
      ShadowStyle     =   1
      Begin VB.TextBox Txt_ust_fl_name 
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txt_UST_FL 
         Height          =   315
         Left            =   8220
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   98
         Top             =   1680
         Width           =   705
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   53
         Tag             =   "Wgt_Grp"
         Top             =   1680
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
         Left            =   1695
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   52
         Tag             =   "交货重量"
         Top             =   1680
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
         Left            =   5580
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   51
         Tag             =   "Del_Tol_Unit"
         Top             =   1320
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
         Left            =   4950
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   50
         Tag             =   "交付公差单位"
         Top             =   1320
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   44
         Tag             =   "Sale_Way"
         Top             =   600
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   43
         Tag             =   "重量单位"
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox txt_pack_way_name 
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
         Left            =   8370
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   42
         Tag             =   "Pack_Way"
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txt_pack_way 
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
         Left            =   7770
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   41
         Tag             =   "包装方法"
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
         ReadOnly        =   -1  'True
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   22
         Left            =   135
         Top             =   270
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_prod_wgt 
         Height          =   315
         Left            =   4950
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
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   30
         Left            =   3480
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   31
         Left            =   6630
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "包装方法"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   23
         Left            =   120
         Top             =   600
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   61
         Left            =   3480
         Top             =   600
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_pack_wgt_min 
         Height          =   315
         Left            =   8190
         TabIndex        =   45
         Tag             =   "包装重量下限"
         Top             =   585
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
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   2
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   32
         Left            =   6630
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "包装重量下限"
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
      Begin CSTextLibCtl.sidbEdit sdb_del_tol_min 
         Height          =   315
         Left            =   1680
         TabIndex        =   46
         Tag             =   "Del_Tol_Min"
         Top             =   960
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
         ReadOnly        =   -1  'True
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
         Top             =   960
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_max 
         Height          =   315
         Left            =   4950
         TabIndex        =   47
         Tag             =   "产品单重上限"
         Top             =   960
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
         NumIntDigits    =   2
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   29
         Left            =   3480
         Top             =   960
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_pack_wgt_max 
         Height          =   315
         Left            =   8190
         TabIndex        =   48
         Tag             =   "包装重量上限"
         Top             =   960
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
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   2
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   33
         Left            =   6630
         Top             =   960
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "包装重量上限"
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
      Begin CSTextLibCtl.sidbEdit sdb_del_tol_max 
         Height          =   315
         Left            =   1680
         TabIndex        =   49
         Tag             =   "交付公差上限"
         Top             =   1320
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
         ReadOnly        =   -1  'True
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
         Left            =   120
         Top             =   1320
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
            Size            =   9.76
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
         Left            =   3480
         Top             =   1320
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   26
         Left            =   120
         Top             =   1680
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
      Begin CSTextLibCtl.sidbEdit sdb_num_prod 
         Height          =   315
         Left            =   4950
         TabIndex        =   94
         Tag             =   "Num_Prod"
         Top             =   1680
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
         ReadOnly        =   -1  'True
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
         Left            =   3480
         Top             =   1680
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   34
         Left            =   6630
         Top             =   1680
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         Caption         =   "UST"
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
      Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_min 
         Height          =   315
         Left            =   4950
         TabIndex        =   103
         Tag             =   "产品单重"
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
   End
   Begin CSTextLibCtl.sitxEdit dtp_ord_accp_date 
      Height          =   315
      Left            =   11640
      TabIndex        =   105
      Top             =   795
      Width           =   1335
      _Version        =   262145
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__"
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-__-__"
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
      Mask            =   "____-__-__"
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Left            =   13455
      TabIndex        =   107
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "退出"
      BevelWidth      =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   150
      X2              =   14775
      Y1              =   1275
      Y2              =   1275
   End
End
Attribute VB_Name = "ACA1031C"
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
'-- Program ID        ACA1030C
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
                 Call Gp_Ms_Collection(txt_ORD_NO, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_ITEM, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_prod_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_prod_dgr, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_prod_dgr_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_stdspec, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stdspec_yy, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stdspec_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STLGRD_Name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ord_cust_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ord_cust_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_cust_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_end_cust_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_end_cust_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_thk, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_wid, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_len, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(cbo_india, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_outdia, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_wgt_grp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_wgt_grp_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_del_cond, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_del_cond_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_num_prod, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_del_fr_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_DEL_TO_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_CUST_del_fr_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt__CUST_del_to_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(dtp_ord_accp_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_transp_way, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_transp_way_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_payment_cond, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_payment_cond_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_sale_way, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_sale_way_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_del_tol_unit, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_del_tol_unit_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_del_tol_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_del_tol_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_dept_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_dept_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_sale_emp_id, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_sale_emp_id_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_wgt_unit, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wgt_unit_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_urgnt_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_urgnt_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_prod_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_prod_wgt_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_prod_wgt_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_enduse_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_enduse_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_TRIM_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_trim_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_spec_no, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cust_spec_no_det, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_cust_req_plant, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cust_req_plant_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_prod_prc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_extra_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_extra_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_discon_prc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_trans_prc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_payment_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_payment_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_marking_way, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_marking_way_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_stamp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stamp_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_pack_way, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_pack_way_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_pack_wgt_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_pack_wgt_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ord_knd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_knd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_hold_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_hold_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_mod_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_mod_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_mod_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_mod_time, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_can_fl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_can_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_can_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_can_emp_id, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_can_emp_id_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_dest_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_dest_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  '          Call Gp_Ms_Collection(txt_dest_detail, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ord_sts, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_sts_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_insp_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_insp_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_currency, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_currency_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_tot_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ord_size, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(TXT_UST_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(Txt_ust_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_TRIM_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Text_size_knd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
    'MASTER Collection
     Mc1.Add Item:="ACA1031C.P_REFER", Key:="P-R"
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

Private Sub Form_Activate()

    Select Case txt_prod_cd.Text

           Case "HC"
                '交付公差单位
                txt_del_tol_unit.Text = "W"

           Case "SL"
                '交付公差单位
                txt_del_tol_unit.Text = "W"

           Case "PP"
                   
                '交付公差单位
                txt_del_tol_unit.Text = "W"


    End Select
    
    If txt_prod_cd.Text <> "PP" Then
        txt_del_tol_unit.Enabled = False
        txt_del_tol_unit_name.Enabled = False
        txt_del_tol_unit.Text = "W"
        txt_insp_cd.Enabled = False
        txt_insp_cd_name.Enabled = False
    End If

    
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority("ABA1010C")

    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_FormCenter(Me)
    
    Screen.MousePointer = vbDefault
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

    

End Sub

Private Sub SSCommand1_Click()
    Unload Me

End Sub
Public Sub Form_Ref()

    Dim sQuery As String

     If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl")) Then
         Call Gp_Ms_ControlLock(Mc1("pControl"), True)
     End If
            
End Sub

