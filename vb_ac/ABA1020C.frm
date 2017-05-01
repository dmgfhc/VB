VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACA1031C 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABA1020C"
   ClientHeight    =   10140
   ClientLeft      =   240
   ClientTop       =   780
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   5250
      Top             =   105
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   2115
      Index           =   0
      Left            =   30
      TabIndex        =   31
      Top             =   2715
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   3731
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
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   36
         Tag             =   "Ord_Size"
         Top             =   240
         Width           =   3255
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
         TabIndex        =   37
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
         TabIndex        =   38
         Tag             =   "订单宽度"
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
         TabIndex        =   39
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
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2145
      Left            =   30
      TabIndex        =   6
      Top             =   555
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   3784
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
         MaxLength       =   40
         TabIndex        =   93
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
         MaxLength       =   1
         TabIndex        =   92
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
         MaxLength       =   40
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         MaxLength       =   40
         TabIndex        =   26
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
         MaxLength       =   7
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         MaxLength       =   40
         TabIndex        =   22
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
         MaxLength       =   1
         TabIndex        =   21
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
         Left            =   6840
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
         Left            =   7750
         MaxLength       =   40
         TabIndex        =   14
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
         Left            =   6840
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
         Left            =   7750
         MaxLength       =   40
         TabIndex        =   9
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
         TabIndex        =   94
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
         Left            =   11610
         TabIndex        =   95
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
            Size            =   9.75
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
      MaxLength       =   1
      TabIndex        =   4
      Tag             =   "Ord_Sts"
      Top             =   105
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
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "Ord_St_name"
      Top             =   105
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
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "Ord_Item"
      Top             =   105
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
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "Ord_No"
      Top             =   105
      Width           =   1275
   End
   Begin InDate.UDate dtp_ord_accp_date 
      Height          =   315
      Left            =   11640
      TabIndex        =   2
      Tag             =   "Ord_Accp_Date"
      Top             =   105
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
      Left            =   150
      Top             =   105
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
      Top             =   105
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
      Left            =   9915
      TabIndex        =   5
      Text            =   "F4"
      Top             =   8685
      Visible         =   0   'False
      Width           =   285
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2805
      Index           =   1
      Left            =   30
      TabIndex        =   32
      Top             =   6435
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   4948
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
         TabIndex        =   102
         Top             =   1500
         Width           =   1545
      End
      Begin VB.TextBox txt_mod_date 
         Height          =   315
         Left            =   1680
         TabIndex        =   101
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
         MaxLength       =   7
         TabIndex        =   74
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
         MaxLength       =   40
         TabIndex        =   73
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
         MaxLength       =   40
         TabIndex        =   72
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
         MaxLength       =   1
         TabIndex        =   71
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
         MaxLength       =   8
         TabIndex        =   70
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
         MaxLength       =   1
         TabIndex        =   69
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
         MaxLength       =   40
         TabIndex        =   68
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
      Left            =   30
      TabIndex        =   33
      Top             =   4845
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   2672
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
         MaxLength       =   11
         TabIndex        =   98
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
         MaxLength       =   60
         TabIndex        =   97
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         MaxLength       =   40
         TabIndex        =   65
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
         MaxLength       =   1
         TabIndex        =   64
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
         MaxLength       =   9
         TabIndex        =   63
         Tag             =   "Cust_Spec_No"
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_cust_spec_no_det 
         Height          =   310
         Left            =   12525
         TabIndex        =   62
         Top             =   300
         Width           =   2340
      End
      Begin VB.ComboBox cbo_india 
         Height          =   300
         ItemData        =   "ABA1020C.frx":0000
         Left            =   6840
         List            =   "ABA1020C.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   300
         Width           =   1500
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
         MaxLength       =   40
         TabIndex        =   59
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
         MaxLength       =   4
         TabIndex        =   58
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
         MaxLength       =   4
         TabIndex        =   57
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
         MaxLength       =   11
         TabIndex        =   56
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
         MaxLength       =   40
         TabIndex        =   55
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
         Left            =   6840
         TabIndex        =   61
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483641
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2775
      Index           =   3
      Left            =   5190
      TabIndex        =   34
      Top             =   6435
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   4895
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
      Begin VB.TextBox txt_del_to_date 
         Height          =   315
         Left            =   1710
         TabIndex        =   104
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox txt_del_fr_date 
         Height          =   315
         Left            =   1710
         TabIndex        =   103
         Top             =   1800
         Width           =   1545
      End
      Begin VB.TextBox txt_dest_detail 
         Height          =   345
         Left            =   6510
         TabIndex        =   99
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
         MaxLength       =   1
         TabIndex        =   90
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
         MaxLength       =   40
         TabIndex        =   89
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
         MaxLength       =   1
         TabIndex        =   88
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
         MaxLength       =   40
         TabIndex        =   87
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
         MaxLength       =   4
         TabIndex        =   86
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
         MaxLength       =   40
         TabIndex        =   85
         Tag             =   "Payment_Cond"
         Top             =   630
         Width           =   2500
      End
      Begin VB.TextBox txt_currency_name 
         Height          =   310
         Left            =   7200
         TabIndex        =   84
         Top             =   240
         Width           =   2500
      End
      Begin VB.TextBox txt_currency 
         Height          =   310
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   83
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
         MaxLength       =   40
         TabIndex        =   82
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
         MaxLength       =   1
         TabIndex        =   81
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
         Left            =   2370
         MaxLength       =   40
         TabIndex        =   80
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
         MaxLength       =   1
         TabIndex        =   79
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
         MaxLength       =   40
         TabIndex        =   78
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
         MaxLength       =   1
         TabIndex        =   77
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
         Left            =   2310
         MaxLength       =   40
         TabIndex        =   76
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
         MaxLength       =   2
         TabIndex        =   75
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
         TabIndex        =   91
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
      Left            =   5190
      TabIndex        =   35
      Top             =   2715
      Width           =   10005
      _ExtentX        =   17648
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
      Caption         =   "重量"
      Begin VB.TextBox Txt_ust_fl_name 
         Height          =   315
         Left            =   9000
         TabIndex        =   106
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txt_UST_FL 
         Height          =   315
         Left            =   8220
         MaxLength       =   1
         TabIndex        =   100
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
         MaxLength       =   40
         TabIndex        =   54
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
         TabIndex        =   53
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
         MaxLength       =   40
         TabIndex        =   52
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
         MaxLength       =   1
         TabIndex        =   51
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
         MaxLength       =   40
         TabIndex        =   45
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
         TabIndex        =   44
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
         MaxLength       =   40
         TabIndex        =   43
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
         MaxLength       =   2
         TabIndex        =   42
         Tag             =   "包装方法"
         Top             =   240
         Width           =   600
      End
      Begin CSTextLibCtl.sidbEdit sdb_tot_wgt 
         Height          =   315
         Left            =   1680
         TabIndex        =   40
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
         Insert          =   0   'False
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
         Left            =   4950
         TabIndex        =   41
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
            Size            =   9.75
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
         TabIndex        =   46
         Tag             =   "包装重量下限"
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
         TabIndex        =   47
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
         Left            =   4950
         TabIndex        =   48
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
            Size            =   9.75
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
         TabIndex        =   49
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
         Left            =   1680
         TabIndex        =   50
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
            Size            =   9.75
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
         TabIndex        =   96
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
         Left            =   4950
         TabIndex        =   105
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4035
      Top             =   -90
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
            Picture         =   "ABA1020C.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":04D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":07F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":09DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":0AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":0DB7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4635
      Top             =   -90
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
            Picture         =   "ABA1020C.frx":1269
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1569
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1649
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1852
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":198A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ABA1020C.frx":1BC5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   105
      X2              =   14730
      Y1              =   510
      Y2              =   510
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
'-- Program ID        ABA1020C
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
                 Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ord_item, "p", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_prod_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_prod_dgr, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_prod_dgr_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_stdspec, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stdspec_yy, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stdspec_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stlgrd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ord_cust_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ord_cust_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_cust_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_end_cust_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_end_cust_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_thk, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_wid, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_ord_len, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(cbo_india, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_outdia, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_wgt_grp, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_wgt_grp_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_del_cond, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_del_cond_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_num_prod, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_del_fr_date, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_del_to_date, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(dtp_ord_accp_date, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_transp_way, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_transp_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_payment_cond, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_payment_cond_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_sale_way, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_sale_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_del_tol_unit, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_del_tol_unit_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_del_tol_max, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_del_tol_min, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_dept_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_dept_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_sale_emp_id, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_sale_emp_id_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_wgt_unit, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wgt_unit_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_urgnt_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_urgnt_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_prod_wgt, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_prod_wgt_min, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_prod_wgt_max, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_enduse_cd, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_enduse_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_trim_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_trim_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_spec_no, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cust_spec_no_det, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_cust_req_plant, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cust_req_plant_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_prod_prc, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_extra_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_extra_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_discon_prc, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_trans_prc, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_payment_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_payment_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_marking_way, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_marking_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_stamp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stamp_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_pack_way, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_pack_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_pack_wgt_max, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_pack_wgt_min, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ord_knd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_knd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_hold_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_hold_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_mod_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_mod_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_mod_date, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_mod_time, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_can_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_can_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_can_date, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_can_emp_id, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_can_emp_id_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_dest_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_dest_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  '          Call Gp_Ms_Collection(txt_dest_detail, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ord_sts, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_sts_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_insp_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_insp_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_currency, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_currency_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_tot_wgt, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ord_size, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_UST_FL, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(Txt_ust_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
'     txt_ord_no.BackColor = &HE0E0E0
'     txt_ord_item.BackColor = &HE0E0E0
'     txt_prod_cd.BackColor = &HE0E0E0
'     txt_prod_cd_name.BackColor = &HE0E0E0
'     txt_prod_dgr.BackColor = &HE0E0E0
'     txt_prod_dgr_name.BackColor = &HE0E0E0
'     txt_cust_cd.BackColor = &HE0E0E0
'     txt_cust_cd_name.BackColor = &HE0E0E0
'     sdb_ord_thk.BackColor = &HE0E0E0
'     sdb_ord_wid.BackColor = &HE0E0E0
'     sdb_ord_len.BackColor = &HE0E0E0
'     sdb_num_prod.BackColor = &HE0E0E0
'     txt_del_fr_date.BackColor = &HE0E0E0
'     txt_del_to_date.BackColor = &HE0E0E0
'     dtp_ord_accp_date.BackColor = &HE0E0E0
'
'     txt_sale_way.BackColor = &HE0E0E0
'     txt_sale_way_name.BackColor = &HE0E0E0
'     txt_dept_cd.BackColor = &HE0E0E0
'     txt_dept_cd_name.BackColor = &HE0E0E0
'     txt_ord_knd.BackColor = &HE0E0E0
'     txt_ord_knd_name.BackColor = &HE0E0E0
'     txt_mod_fl.BackColor = &HE0E0E0
'     txt_mod_fl_name.BackColor = &HE0E0E0
'     txt_mod_date.BackColor = &HE0E0E0
'     txt_mod_time.BackColor = &HE0E0E0
'     txt_can_fl.BackColor = &HE0E0E0
'     txt_can_fl_name.BackColor = &HE0E0E0
'     txt_can_date.BackColor = &HE0E0E0
'     txt_can_emp_id.BackColor = &HE0E0E0
'     txt_can_emp_id_name.BackColor = &HE0E0E0
'     txt_ord_sts.BackColor = &HE0E0E0
'     txt_ord_sts_name.BackColor = &HE0E0E0
'
    'MASTER Collection
   ' Mc1.Add Item:="ABA1020C.P_MODIFY", Key:="P-M"
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
'     Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Activate()
'
'    If Dis_sw = False Then
'        Exit Sub
'    Else
'        Dis_sw = False
'    End If
'
'    If Mc1("pControl").Item(1).Text = "" Then
''        pControl(1).Text = ACA1030C.txt_ord_no.Text
''        pControl(1).Enabled = True
''        pControl(2).Enabled = True
''
''        txt_cust_cd.Text = ACA1030C.txt_cust_cd.Text
''        txt_cust_cd_name.Text = ABA1010C.txt_cust_cd_name.Text
''        txt_prod_cd.Text = ABA1010C.txt_prod_cd.Text
''        txt_prod_cd_name.Text = ABA1010C.txt_prod_cd_name.Text
''        txt_prod_dgr.Text = ABA1010C.txt_prod_dgr.Text
''        txt_prod_dgr_name.Text = ABA1010C.txt_prod_dgr_name.Text
''        txt_dept_cd.Text = ABA1010C.txt_dept_cd.Text
''        txt_dept_cd_name.Text = ABA1010C.txt_dept_cd_name.Text
''        txt_sale_way.Text = ABA1010C.txt_sale_way.Text
''        txt_sale_way_name.Text = ABA1010C.txt_sale_way_name.Text
''        txt_ord_knd.Text = ABA1010C.txt_ord_knd.Text
''        txt_ord_knd_name.Text = ABA1010C.txt_ord_knd_name.Text
''
''        pControl(2).SetFocus
'
'    Else
'
        Call Gf_Ms_Refer(M_CN1, Mc1)
        
    'End If

    Select Case txt_prod_cd.Text

           Case "HC"
           
                '订单内径/订单外径/检查机关/切边代码/UST
                cbo_india.Enabled = True
'                sdb_outdia.Enabled = True
'                txt_insp_cd.Enabled = False
'                txt_trim_fl.Enabled = False
'                txt_UST_FL.Enabled = False
'
                '重量单位
'               txt_wgt_unit.Enabled = True
                
                '产品单重下限/产品单重上限/产品单重
'                sdb_prod_wgt_min.Enabled = True
'                sdb_prod_wgt_max.Enabled = True
'                sdb_prod_wgt.Enabled = True
'
                '包装重量下限/包装重量上限/包装方法
'                sdb_pack_wgt_min.Enabled = True
'                sdb_pack_wgt_max.Enabled = True
'                txt_pack_way.Enabled = True
                
                '交付公差单位
'                txt_del_tol_unit.Enabled = False
                txt_del_tol_unit.Text = "W"

           Case "SL"
           
                '订单内径/订单外径/检查机关/切边代码/UST
                cbo_india.Enabled = False
'                sdb_outdia.Enabled = False
'                txt_insp_cd.Enabled = False
'                txt_trim_fl.Enabled = False
'                txt_UST_FL.Enabled = False
'
                '重量单位
'                txt_wgt_unit.Enabled = True
                
                '产品单重下限/产品单重上限/产品单重
'                sdb_prod_wgt_min.Enabled = False
'                sdb_prod_wgt_max.Enabled = False
'                sdb_prod_wgt.Enabled = False
                
                '包装重量下限/包装重量上限/包装方法
'                sdb_pack_wgt_min.Enabled = False
'                sdb_pack_wgt_max.Enabled = False
'                txt_pack_way.Enabled = False
                '交付公差单位
'                txt_del_tol_unit.Enabled = False
                txt_del_tol_unit.Text = "W"

           Case "PP"
                   
                '订单内径/订单外径/检查机关/切边代码/UST
'                cbo_india.Enabled = False
'                sdb_outdia.Enabled = False
'                txt_insp_cd.Enabled = True
'                txt_trim_fl.Enabled = True
'                txt_UST_FL.Enabled = True
                
                '重量单位
             '   txt_wgt_unit.Enabled = True
                
                '产品单重下限/产品单重上限/产品单重
'                sdb_prod_wgt_min.Enabled = False
'                sdb_prod_wgt_max.Enabled = False
'                sdb_prod_wgt.Enabled = False
                
                '包装重量下限/包装重量上限/包装方法
'                sdb_pack_wgt_min.Enabled = False
'                sdb_pack_wgt_max.Enabled = False
'                txt_pack_way.Enabled = True
                
                '交付公差单位
'                txt_del_tol_unit.Enabled = False
                txt_del_tol_unit.Text = "W"


    End Select
    
   ' If txt_prod_cd.Text = "HC" Then
    
'        cbo_india.BackColor = &HC0E0FF
'    '    sdb_outdia.BackColor = &HC0E0FF
'    '    txt_wgt_unit.BackColor = &HC0E0FF
'        sdb_prod_wgt_min.BackColor = &HC0E0FF
'        sdb_prod_wgt_max.BackColor = &HC0E0FF
'        sdb_prod_wgt.BackColor = &HC0E0FF
'        sdb_pack_wgt_min.BackColor = &HC0E0FF
'        sdb_pack_wgt_max.BackColor = &HC0E0FF
'        txt_pack_way.BackColor = &HC0E0FF
'        txt_pack_way_name.BackColor = &HC0E0FF
'    Else
'
'        '颜色
'        cbo_india.BackColor = &HE0E0E0
'        sdb_outdia.BackColor = &HE0E0E0
'        txt_wgt_unit.BackColor = &HE0E0E0
'        txt_wgt_unit_name.BackColor = &HE0E0E0
'        sdb_prod_wgt_min.BackColor = &HE0E0E0
'        sdb_prod_wgt_max.BackColor = &HE0E0E0
'        sdb_prod_wgt.BackColor = &HE0E0E0
'        sdb_pack_wgt_min.BackColor = &HE0E0E0
'        sdb_pack_wgt_max.BackColor = &HE0E0E0
'        txt_pack_way.BackColor = &HE0E0E0
'        txt_pack_way_name.BackColor = &HE0E0E0
'    End If

    If txt_prod_cd.Text <> "PP" Then
        txt_del_tol_unit.Enabled = False
        txt_del_tol_unit_name.Enabled = False
        txt_del_tol_unit.Text = "W"
        txt_insp_cd.Enabled = False
        txt_insp_cd_name.Enabled = False
'        txt_del_tol_unit.BackColor = &HE0E0E0
'        txt_del_tol_unit_name.BackColor = &HE0E0E0
'        txt_insp_cd.BackColor = &HE0E0E0
'        txt_insp_cd_name.BackColor = &HE0E0E0
    End If

'
'   txt_del_fr_date.Text = ""
'   txt_del_to_date.Text = ""
'   txt_can_date.Text = ""
'   txt_mod_date.Text = ""
    

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

  '  Call Popup_Menu_Setting

    Call Form_Define

   ' Call Gp_Ms_Cls(Mc1("rControl"))

 '   Call Gp_Ms_ControlLock(Mc1("lControl"), True)

  '  Call Gp_Ms_ControlLock(Mc1("pControl"), True)

  '  Call Gp_Ms_NeceColor(Mc1("nControl"))
    
   ' Call Gp_FormCenter(Me)
   
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

  '  Call ABA1010C.Form_Ref
    

End Sub

Public Sub Form_Exit()

    Unload Me
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub



Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub



Public Sub Form_Pro()

    Dim sMesg As String
    Dim sQuery As String
    
    Select Case Me.ActiveControl.Name
           Case "txt_ord_size"
                 Call txt_ord_size_LostFocus
           Case "sdb_prod_wgt"
'                 Call sdb_prod_wgt_LostFocus
           Case "sdb_tot_wgt"
'                 Call sdb_tot_wgt_LostFocus
           Case "txt_extra_fl"
                 Call txt_extra_fl_LostFocus
           Case "txt_pack_way"
                 Call txt_pack_way_LostFocus
    End Select
    
    If ACA1030C.txt_prod_cd.Text = "HC" Then
    
        If cbo_india.Text = "" Then
           Call Gp_MsgBoxDisplay("内径必须输入", "I")
           Exit Sub
        End If
        
        If txt_wgt_unit.Text = "" Or sdb_prod_wgt_min.Value = 0 Or sdb_prod_wgt_max.Value = 0 Or sdb_prod_wgt.Value = 0 Then
           Call Gp_MsgBoxDisplay("单重，重量上下限，重量单位必须输入", "I")
           Exit Sub
        End If
        
        If sdb_pack_wgt_min.Value = 0 Or sdb_pack_wgt_max.Value = 0 Or txt_pack_way.Text = "" Then
           Call Gp_MsgBoxDisplay("包装方式，包装重量必须输入", "I")
           Exit Sub
        End If
         
        If sdb_prod_wgt.Value < sdb_prod_wgt_min.Value Or sdb_prod_wgt.Value > sdb_prod_wgt_max.Value Then
           Call Gp_MsgBoxDisplay("产品单重必须在上下限之间", "I")
           Exit Sub
        End If
        
    End If
    
    If txt_prod_cd.Text = "HC" Then
       If sdb_prod_wgt_min.Value >= sdb_prod_wgt_max.Value Then Call Gp_MsgBoxDisplay("产品单重上限不可小于下限", "I"): Exit Sub
    
       If sdb_pack_wgt_min.Value >= sdb_pack_wgt_max.Value Then Call Gp_MsgBoxDisplay("包装重量上限不可小于下限", "I"): Exit Sub
    
       If sdb_prod_wgt_min.Value > sdb_pack_wgt_min.Value Or sdb_prod_wgt_max.Value > sdb_pack_wgt_max.Value Then
          Call Gp_MsgBoxDisplay("包装重量上下限必须大于等于产品单重上限", "I")
          Exit Sub
       End If
       If txt_pack_way.Text = "NO" Then
          If sdb_pack_wgt_min.Value <> sdb_prod_wgt_min.Value Or sdb_pack_wgt_max.Value <> sdb_prod_wgt_max.Value Then
             Call Gp_MsgBoxDisplay("未包装时，包装重量上下限必须等于产品单重上下限", "I")
          End If
          Exit Sub
    End If
       
       
    End If
    
    If sdb_del_tol_min.Value > sdb_del_tol_max.Value Then Call Gp_MsgBoxDisplay("交付公差上限不可小于下限", "I"): Exit Sub
    
    Dim sQuery1, sQuery2 As String
   
    sQuery1 = "{CALL ABA1020C.P_MASTER_CHECK ('" + txt_prod_cd.Text + "','" + txt_stdspec.Text + "','" + Trim(sdb_ord_thk.Value) + "','"
    sQuery1 = sQuery1 + Trim(sdb_ord_wid.Value) + "','" + Trim(sdb_ord_len.Value) + "','" + Trim(sdb_tot_wgt.Value) + "','" + Trim(sdb_prod_wgt.Value) + "',?,?,?)}"
     
    Dim OutParam(3, 4) As Variant
    
    OutParam(1, 1) = "arg_e_wgt"
    OutParam(1, 2) = adVarNumeric
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    'Return Error Code Parameter
    OutParam(2, 1) = "arg_e_code"
    OutParam(2, 2) = adInteger
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 1

    'Return Error Messsage Parameter
    OutParam(3, 1) = "arg_e_msg"
    OutParam(3, 2) = adVarChar
    OutParam(3, 3) = adParamOutput
    OutParam(3, 4) = 256
    
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim adoCmd As ADODB.Command
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery1
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(3, 1), OutParam(3, 2), OutParam(3, 3), OutParam(3, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "0" Then

        ret_Result_ErrCode = adoCmd("arg_e_code")
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
        
        Call Gp_MsgBoxDisplay(sErrMessg)
        
        Set adoCmd = Nothing
    
    
        Exit Sub
    Else
    
        sdb_prod_wgt.Value = adoCmd("arg_e_wgt")
    
    End If
    

    Set adoCmd = Nothing
    
    
     If Gf_Mc_Authority(sAuthority, Mc1) Then
    
        If Mc1.Item("pControl")(1).Enabled Then
            sQuery = "{call ABA1020C.P_ORD_SEQ ( '" + txt_ord_no.Text + "' )}"
            txt_ord_item.Text = Gf_CodeFind(M_CN1, sQuery)
            txt_ord_sts.Text = "A"
            Call txt_ord_sts_KeyUp(0, 0)
        Else
            If txt_ord_sts.Text <> "A" Then
                sMesg = "不能修改状态不是'A'的订单"
                Call Gp_MsgBoxDisplay(sMesg)
                Exit Sub
            End If
        End If
    
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
           ' Call Popup_Menu_Setting
        End If
    
    End If
    
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
        'Call Popup_Menu_Setting
    Else
        'Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If
    
End Sub





Private Sub sdb_prod_wgt_max_LostFocus()
'    If txt_pack_way.Text = "NO" Then
'       sdb_pack_wgt_max.Value = sdb_prod_wgt_max.Value
'    End If
End Sub


Private Sub sdb_prod_wgt_min_LostFocus()

'    If txt_pack_way.Text = "NO" Then
'       sdb_pack_wgt_min.Value = sdb_prod_wgt_min.Value
'    End If

End Sub


Private Sub txt_cust_req_plant_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0031"
        DD.rControl.Add Item:=txt_cust_req_plant
        DD.rControl.Add Item:=txt_cust_req_plant_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_req_plant)) = txt_cust_req_plant.MaxLength Then
        txt_cust_req_plant_name.Text = Gf_ComnNameFind(M_CN1, "B0031", Trim(txt_cust_req_plant.Text), 2)
    Else
        txt_cust_req_plant_name.Text = ""
    End If

End Sub

Private Sub txt_cust_spec_no_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF4 Then
       
'       Load ABX1110C
'
'       ABX1110C.txt_form_nm.Text = "ABA1020C"
'       ABX1110C.txt_cust_cd.Text = txt_cust_cd.Text
'       ABX1110C.txt_prod_cd.Text = txt_prod_cd.Text
'
'       ABX1110C.Show 1
       
    End If
    
End Sub

Private Sub txt_dest_detail_KeyUp(KeyCode As Integer, Shift As Integer)

' If KeyCode = vbKeyF4 Then
'
'       Load ABX1120C
'
'       ABX1120C.txt_form_nm.Text = "ABA1020C"
'       ABX1120C.txt_dest_cd.Text = txt_dest_cd.Text
'
'       ABX1120C.Show 1
'
'    End If

End Sub


Private Sub txt_extra_fl_LostFocus()

    If txt_extra_fl.Text = "N" Then
    
       sdb_discon_prc.Value = 0
       sdb_discon_prc.Enabled = False
'       sdb_discon_prc.BackColor = &HE0E0E0
    Else
       sdb_discon_prc.Enabled = True

    End If

End Sub

Private Sub txt_ord_size_LostFocus()
    
    Dim T As Double
    Dim W As Double
    Dim L As Double
    Dim N1 As Long
    Dim N2 As Long
    Dim N3 As Long
    Dim Num As Long

    If txt_ord_size.Text <> "" Then
       N1 = InStr(1, txt_ord_size.Text, "*")
       N2 = InStr(N1 + 1, txt_ord_size.Text, "*")
       N3 = InStr(N2 + 1, txt_ord_size.Text, "*")
       
       If N1 = 0 Or N2 = 0 Then
          Call Gp_MsgBoxDisplay("尺寸不完整", "I")
          txt_ord_size.SetFocus
          Exit Sub
       End If
       
       If N3 > 0 Then
          Call Gp_MsgBoxDisplay("输入错误", "I")
          txt_ord_size.SetFocus
          Exit Sub
       End If
       
       If InStr(1, txt_ord_size.Text, """") = 0 Then   '判断该字符串中第一个"的位置，为0则没有
          T = Val(Mid(Trim(txt_ord_size.Text), 1, N1 - 1))
          W = Val(Mid(Trim(txt_ord_size.Text), N1 + 1, N2 - 1))
          Num = Len(Trim(txt_ord_size.Text))
            If txt_prod_cd.Text = "HC" Then
               If Mid(Trim(txt_ord_size.Text), N2 + 1, Num - N2) = "C" Then
                  L = 0
               Else
                  Call Gp_MsgBoxDisplay("输入错误", "I")
                  txt_ord_size.SetFocus
                  Exit Sub
               End If
            Else
               L = Val(Mid(Trim(txt_ord_size.Text), N2 + 1, Num - N2))
            End If
       Else
          N1 = InStr(1, txt_ord_size.Text, "*")
          N2 = InStr(N1 + 1, txt_ord_size.Text, "*")
          If InStr(1, txt_ord_size.Text, """") = N1 - 1 And InStr(N1, txt_ord_size.Text, """") = N2 - 1 Then
            T = Val(Mid(Trim(txt_ord_size.Text), 1, N1 - 2)) * 2.54
            W = Val(Mid(Trim(txt_ord_size.Text), N1 + 1, N2 - 2)) * 2.54
            Num = Len(Trim(txt_ord_size.Text))
              If txt_prod_cd.Text = "HC" Then
                 If Mid(Trim(txt_ord_size.Text), N2 + 1, Num - N2) = "C" Then
                    L = 0
                 Else
                    Call Gp_MsgBoxDisplay("输入错误", "I")
                    txt_ord_size.SetFocus
                    Exit Sub
                 End If
              Else
                  If InStr(N2, txt_ord_size.Text, """") = Num Then
                    L = Val(Mid(Trim(txt_ord_size.Text), N2 + 1, Num - N2 - 1)) * 2.54
                  Else
                     Call Gp_MsgBoxDisplay("输入错误", "I")
                     txt_ord_size.SetFocus
                     Exit Sub
                  End If
              End If
          Else
             Call Gp_MsgBoxDisplay("输入错误,", "I")
             txt_ord_size.SetFocus
             Exit Sub
          End If
       End If
    Else
       Exit Sub
    End If
    
    sdb_ord_thk.Value = T
    sdb_ord_wid.Value = W
    sdb_ord_len.Value = L

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

Private Sub txt_pack_way_GotFocus()

    If (sdb_prod_wgt_min = 0) Or (sdb_prod_wgt_min = 0) Then
    
 '      Call Gp_MsgBoxDisplay("请输入产品单重上下限", "I")
       
    End If

End Sub

Private Sub txt_pack_way_LostFocus()

    If txt_prod_cd.Text = "HC" Then
    
        If txt_pack_way.Text = "NO" Then
           sdb_pack_wgt_min.Enabled = False
           sdb_pack_wgt_max.Enabled = False
'           sdb_pack_wgt_min.BackColor = &HE0E0E0
'           sdb_pack_wgt_max.BackColor = &HE0E0E0
           sdb_pack_wgt_min.Value = sdb_prod_wgt_min.Value
           sdb_pack_wgt_max.Value = sdb_prod_wgt_max.Value
        Else
           sdb_pack_wgt_min.Enabled = True
           sdb_pack_wgt_max.Enabled = True
        End If
    End If
    
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

Private Sub txt_stamp_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0030"
        DD.rControl.Add Item:=txt_stamp
        DD.rControl.Add Item:=txt_stamp_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_stamp)) = txt_stamp.MaxLength Then
        txt_stamp_name.Text = Gf_ComnNameFind(M_CN1, "B0030", Trim(txt_stamp.Text), 2)
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
'       cbo_india.BackColor = &HE0E0E0
'       sdb_outdia.BackColor = &HE0E0E0
    
    End If

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


Private Sub txt_transp_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0020"
        DD.rControl.Add Item:=txt_transp_way
        DD.rControl.Add Item:=txt_transp_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_transp_way)) = txt_transp_way.MaxLength Then
        txt_transp_way_name.Text = Gf_ComnNameFind(M_CN1, "B0020", Trim(txt_transp_way.Text), 2)
    Else
        txt_transp_way_name.Text = ""
    End If

End Sub

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

Private Sub txt_pack_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0025"
        DD.rControl.Add Item:=txt_pack_way
        DD.rControl.Add Item:=txt_pack_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_pack_way)) = txt_pack_way.MaxLength Then
        txt_pack_way_name.Text = Gf_ComnNameFind(M_CN1, "B0025", Trim(txt_pack_way.Text), 2)
    Else
        txt_pack_way_name.Text = ""
    End If

End Sub


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

Private Sub txt_del_cond_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0016"
        DD.rControl.Add Item:=txt_del_cond
        DD.rControl.Add Item:=txt_del_cond_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_del_cond)) = txt_del_cond.MaxLength Then
        txt_del_cond_name.Text = Gf_ComnNameFind(M_CN1, "B0016", Trim(txt_del_cond.Text), 2)
    Else
        txt_del_cond_name.Text = ""
    End If

End Sub
Private Sub txt_payment_cond_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0015"
        DD.rControl.Add Item:=txt_payment_cond
        DD.rControl.Add Item:=txt_payment_cond_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_payment_cond)) = txt_payment_cond.MaxLength Then
        txt_payment_cond_name.Text = Gf_ComnNameFind(M_CN1, "B0015", Trim(txt_payment_cond.Text), 2)
    Else
        txt_payment_cond_name.Text = ""
    End If

End Sub
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

Private Sub txt_extra_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0014"
        DD.rControl.Add Item:=txt_extra_fl
        DD.rControl.Add Item:=txt_extra_fl_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_extra_fl)) = txt_extra_fl.MaxLength Then
        txt_extra_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0014", Trim(txt_extra_fl.Text), 2)
    Else
        txt_extra_fl_name.Text = ""
    End If

End Sub
Private Sub txt_currency_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0013"
        DD.rControl.Add Item:=txt_currency
        DD.rControl.Add Item:=txt_currency_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_currency)) = txt_currency.MaxLength Then
        txt_currency_name.Text = Gf_ComnNameFind(M_CN1, "B0013", Trim(txt_currency.Text), 2)
    Else
        txt_currency_name.Text = ""
    End If

End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_name

        DD.nameType = "2"

        Call Gf_Stlgrd_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
       txt_stlgrd.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
    Else
       txt_stlgrd_name.Text = ""
    End If

End Sub


Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec
        DD.rControl.Add Item:=txt_stdspec_yy
        DD.rControl.Add Item:=txt_stdspec_name

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub


Private Sub txt_enduse_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

           '  ABX1050C.Show 1
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


Private Sub txt_prod_dgr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0029"
        DD.rControl.Add Item:=txt_prod_dgr
        DD.rControl.Add Item:=txt_prod_dgr_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_prod_dgr)) = txt_prod_dgr.MaxLength Then
        txt_prod_dgr_name.Text = Gf_ComnNameFind(M_CN1, "B0029", Trim(txt_prod_dgr.Text), 2)
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




