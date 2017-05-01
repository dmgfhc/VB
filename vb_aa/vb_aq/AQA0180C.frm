VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0180C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "质量设计键输入_AQA0180C"
   ClientHeight    =   4365
   ClientLeft      =   1455
   ClientTop       =   3600
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   10770
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10710
      TabIndex        =   26
      Top             =   0
      Width           =   10770
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   600
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   15420
         _ExtentX        =   27199
         _ExtentY        =   1058
         BandCount       =   1
         _CBWidth        =   15420
         _CBHeight       =   600
         _Version        =   "6.7.9782"
         Child1          =   "MenuTool"
         MinHeight1      =   540
         Width1          =   15360
         NewRow1         =   0   'False
         BandStyle1      =   1
         Begin MSComctlLib.Toolbar MenuTool 
            Height          =   540
            Left            =   30
            TabIndex        =   28
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
                  Key             =   "Line3"
                  Style           =   4
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
   Begin VB.TextBox txt_upd_name 
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
      Left            =   6810
      MaxLength       =   20
      TabIndex        =   19
      Top             =   3900
      Width           =   1785
   End
   Begin VB.TextBox txt_ins_name 
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
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   18
      Top             =   3900
      Width           =   1785
   End
   Begin VB.TextBox txt_STDSPEC 
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
      Left            =   1470
      MaxLength       =   18
      TabIndex        =   17
      Top             =   1140
      Width           =   2895
   End
   Begin VB.TextBox txt_MLT_STD_NO 
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
      Left            =   5040
      MaxLength       =   6
      TabIndex        =   16
      Top             =   2505
      Width           =   1995
   End
   Begin VB.TextBox txt_MILL_STD_NO 
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
      Left            =   8610
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2505
      Width           =   1995
   End
   Begin VB.TextBox txt_NISCO_QUALITY_NO 
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
      Left            =   1470
      MaxLength       =   8
      TabIndex        =   14
      Top             =   2505
      Width           =   1995
   End
   Begin Threed.SSCheck Chk_DRT_CNF_TYP 
      Height          =   315
      Left            =   1470
      TabIndex        =   13
      Top             =   2940
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   196609
      BackColor       =   -2147483641
      BackStyle       =   1
   End
   Begin VB.TextBox txt_DRT_CNF_TYP 
      Height          =   270
      Left            =   4290
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   285
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   60
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "品种"
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
   Begin VB.TextBox txt_UPD_EMP 
      Height          =   300
      Left            =   6810
      TabIndex        =   11
      Top             =   3900
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txt_UPD_DATE 
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
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3495
      Width           =   1200
   End
   Begin VB.TextBox txt_INS_EMP 
      Height          =   300
      Left            =   1470
      TabIndex        =   9
      Top             =   3900
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txt_INS_DATE 
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
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3495
      Width           =   1200
   End
   Begin VB.TextBox txt_STEEL_GRD 
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
      Left            =   1470
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2115
      Width           =   1500
   End
   Begin VB.TextBox txt_STEEL_GRD_NAME 
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
      Left            =   3000
      TabIndex        =   7
      Top             =   2115
      Width           =   6255
   End
   Begin VB.TextBox txt_CUST_NAME 
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
      Left            =   7650
      TabIndex        =   5
      Top             =   1140
      Width           =   3045
   End
   Begin VB.TextBox txt_CUST_CD 
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
      Left            =   6810
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1140
      Width           =   825
   End
   Begin VB.TextBox txt_ENDUSE_NAME 
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
      Left            =   7650
      TabIndex        =   3
      Top             =   720
      Width           =   3045
   End
   Begin VB.TextBox txt_ENDUSE_CD 
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
      Left            =   6810
      MaxLength       =   4
      TabIndex        =   2
      Top             =   720
      Width           =   825
   End
   Begin VB.TextBox txt_PROD_KND_NAME 
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
      Left            =   1935
      TabIndex        =   1
      Top             =   720
      Width           =   2445
   End
   Begin VB.TextBox txt_PROD_KND 
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
      Left            =   1470
      MaxLength       =   2
      TabIndex        =   0
      Top             =   720
      Width           =   435
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   5400
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   60
      Top             =   1140
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "标准编号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   5400
      Top             =   1140
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "客户代码"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   4
      Left            =   60
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "厚度组"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   5
      Left            =   3630
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "宽度组"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   6
      Left            =   7200
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "长度组"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   7
      Left            =   60
      Top             =   2505
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "企标材质编号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   8
      Left            =   60
      Top             =   2115
      Width           =   1335
      _ExtentX        =   2355
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   9
      Left            =   3630
      Top             =   2505
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "炼钢规程编号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   10
      Left            =   7200
      Top             =   2505
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "轧钢规程编号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   11
      Left            =   60
      Top             =   2940
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "是否直接投入"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   12
      Left            =   45
      Top             =   3495
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "录入日期"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   13
      Left            =   60
      Top             =   3900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "录入人"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   14
      Left            =   5400
      Top             =   3495
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "修改日期"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   15
      Left            =   5400
      Top             =   3900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "修改人"
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
   Begin CSTextLibCtl.sidbEdit sdb_THK_MIN 
      Height          =   315
      Left            =   1470
      TabIndex        =   20
      Tag             =   "厚度组下限"
      Top             =   1530
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_MAX 
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Tag             =   "厚度组上限"
      Top             =   1530
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_WID_MIN 
      Height          =   315
      Left            =   5040
      TabIndex        =   22
      Tag             =   "宽度组下限"
      Top             =   1560
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_WID_MAX 
      Height          =   315
      Left            =   6090
      TabIndex        =   23
      Tag             =   "宽度组上限"
      Top             =   1560
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_LEN_MIN 
      Height          =   315
      Left            =   8610
      TabIndex        =   24
      Tag             =   "长度组下限"
      Top             =   1560
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_LEN_MAX 
      Height          =   315
      Left            =   9660
      TabIndex        =   25
      Tag             =   "长度组上限"
      Top             =   1560
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9990
      Top             =   3480
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
            Picture         =   "AQA0180C.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":04B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":07D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":09C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":0AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":0D9B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9360
      Top             =   3510
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
            Picture         =   "AQA0180C.frx":124D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":154D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":162D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":1836
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0180C.frx":1BA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   45
      X2              =   10710
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   60
      X2              =   10710
      Y1              =   1980
      Y2              =   1980
   End
End
Attribute VB_Name = "AQA0180C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      质量设计键输入
'-- Program ID        AQA0170C (Master-AQA0180C)
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       质量设计键输入
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

Dim Mc1 As New Collection           'Master Collection
Dim sQuery As String


Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "PopMaster"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary )", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_PROD_KND, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PROD_KND_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ENDUSE_CD, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ENDUSE_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STDSPEC, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_CUST_CD, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_CUST_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_THK_MIN, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_THK_MAX, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_WID_MIN, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_WID_MAX, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_LEN_MIN, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_LEN_MAX, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STEEL_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STEEL_GRD_Name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_NISCO_QUALITY_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_MLT_STD_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_MILL_STD_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_DRT_CNF_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ins_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_INS_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_upd_emp, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_UPD_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
     Mc1.Add Item:="AQA0170C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQA0170C.P_REFER", Key:="P-R"
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
            
        Case "txt_PROD_KND"             '品种
            sCode = "Q0001"
            Set oCodeName = txt_PROD_KND_NAME
            
        Case "txt_ENDUSE_CD"            '订单用途
            sCode = "ENDUSE_CD"
            Set oCodeName = txt_ENDUSE_NAME
            DD.sKey = txt_PROD_KND.Text
            
        Case "txt_STDSPEC"              '标准编号
            sCode = "STDSPEC"
        
        Case "txt_CUST_CD"              '客户代码
            sCode = "CUST_CD"
            Set oCodeName = txt_CUST_NAME
                
        Case "txt_STEEL_GRD"            '钢种
            sCode = "STLGRD"
            Set oCodeName = txt_STEEL_GRD_Name
                
        Case "txt_NISCO_QUALITY_NO"     '企标材质编号
            sCode = "NISCO_QUALITY_NO"
            
        Case "txt_MLT_STD_NO"           '炼钢规程编号
            sCode = "MLT_STD_NO"
            
        Case "txt_MILL_STD_NO"          '轧钢规程编号
            sCode = "MILL_STD_NO"
                
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub



Private Sub Chk_DRT_CNF_TYP_Click(Value As Integer)
    
    Select Case Chk_DRT_CNF_TYP.Value
            Case 0
                txt_DRT_CNF_TYP.Text = "0"
            Case -1
                txt_DRT_CNF_TYP.Text = "1"
            Case Else
                txt_DRT_CNF_TYP.Text = "0"
    End Select
    
End Sub

Private Sub Form_Activate()

    If Mc1("pControl").Item(1).Text = "" Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
    End If
    
    If txt_DRT_CNF_TYP.Text = "1" Then
        Chk_DRT_CNF_TYP.Value = ssCBChecked
    Else
        Chk_DRT_CNF_TYP.Value = ssCBUnchecked
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
    
    sAuthority = Gf_Pgm_Authority("AQA0180C", True)

    
    Call Popup_Menu_Setting
    
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

    Call AQA0170C.Form_Ref

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    MenuTool.Buttons(4).Enabled = False    'Delete
    MenuTool.Buttons(6).Enabled = False    'Copy
    MenuTool.Buttons(7).Enabled = False    'Paste
    
    txt_CUST_NAME.Text = ""
    txt_STEEL_GRD_Name.Text = ""
    txt_PROD_KND_NAME.Text = ""
    txt_ENDUSE_NAME.Text = ""
    txt_DRT_CNF_TYP.Text = "1"
    
    pControl(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then MenuTool.Buttons(4).Enabled = False   'Delete
    
End Sub

Public Sub Form_Pro()
  Dim i_Save_Check As Integer
  Dim s_Msg As String
  
   i_Save_Check = Data_Save_Check(Trim(txt_PROD_KND.Text))
    
    If i_Save_Check = 0 Then
        If Gf_Mc_Authority(sAuthority, Mc1) Then
            txt_ins_emp.Text = sUserID
            If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
                Call Popup_Menu_Setting
            End If
        End If
    Else
        Select Case i_Save_Check
            Case 1
                s_Msg = "请输入钢种!"
            Case 2
                s_Msg = "请输入炼钢/连铸操作规程编号!"
            Case 3
                s_Msg = "请输入轧钢操作规程编号!"
        End Select
        
        Call Gp_MsgBoxDisplay(s_Msg)
    
    End If
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then
        Call Popup_Menu_Setting
    End If
    
End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "粘贴"
            '应做:添加 '粘贴' 按钮代码。
            MsgBox "添加 '粘贴' 按钮代码。"
        Case "复制"
            '应做:添加 '复制' 按钮代码。
            MsgBox "添加 '复制' 按钮代码。"
        Case "删除"
            '应做:添加 '删除' 按钮代码。
            MsgBox "添加 '删除' 按钮代码。"
        Case "保存"
            '应做:添加 '保存' 按钮代码。
            MsgBox "添加 '保存' 按钮代码。"
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






Private Sub sdb_LEN_MAX_KeyPress(KeyAscii As Integer)

  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub sdb_LEN_MAX_Validate(Cancel As Boolean)
        
        If txt_PROD_KND.Text <> "H" Then
        
            If Len(Trim(sdb_LEN_MAX.Text)) <> 0 Then
                If Not (txt_Max_Check(sdb_LEN_MAX.Text, sdb_LEN_MIN.Text)) Then
                    
                   MsgBox ("请检查长度组最小值和最大值，后者不能小与前者")
                   
                   Cancel = True
        
                End If
            
            Else
                   MsgBox ("请输入数值")
                   
                   Cancel = True
            
            End If
        
        Else
            If sdb_LEN_MAX.Value = 0 Then sdb_LEN_MAX.Value = 999999.9
        End If

End Sub

Private Sub sdb_LEN_MIN_KeyPress(KeyAscii As Integer)
  
  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub sdb_LEN_MIN_Validate(Cancel As Boolean)
        
        If txt_PROD_KND.Text <> "H" Then
        
            If Len(Trim(sdb_LEN_MIN.Text)) <> 0 Then
                If Not (txt_Max_Check(sdb_LEN_MAX.Text, sdb_LEN_MIN.Text)) Then
                    
                   MsgBox ("请检查长度组最小值和最大值，后者不能小与前者")
                   
                   Cancel = True
        
                End If
            
            Else
                   MsgBox ("请输入数值")
                   
                   Cancel = True
            
            End If
    
    
    End If

End Sub



Private Sub sdb_THK_MAX_KeyPress(KeyAscii As Integer)
  
  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub sdb_THK_MAX_Validate(Cancel As Boolean)
        
        If Len(Trim(sdb_THK_MAX.Text)) <> 0 Then
            If Not (txt_Max_Check(sdb_THK_MAX.Text, sdb_THK_MIN.Text)) Then
                
               MsgBox ("请检查厚度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub sdb_THK_MIN_KeyPress(KeyAscii As Integer)
  
  KeyAscii = txt_KeyPress(KeyAscii)

End Sub


Private Function txt_KeyPress(KeyAscii As Integer) As Integer

        Select Case KeyAscii
               
               Case Is <= 32
                    txt_KeyPress = KeyAscii
               Case 48 To 57
                    txt_KeyPress = KeyAscii
               Case 46
                    txt_KeyPress = KeyAscii
               Case Else
                    txt_KeyPress = 0
        End Select

    
End Function

Private Function txt_Max_Check(Max_Num, Min_Num As String) As Boolean
          
        If Len(Trim(Max_Num)) <> 0 Then
   
            If Val(Trim(Max_Num)) < Val(Trim(Min_Num)) Then
               
               txt_Max_Check = False
            
            Else
               
               txt_Max_Check = True
               
            End If
        
        Else
        
            txt_Max_Check = True
        
        End If
    
End Function

Private Sub sdb_THK_MIN_Validate(Cancel As Boolean)
        
        If Len(Trim(sdb_THK_MIN.Text)) <> 0 Then
            If Not (txt_Max_Check(sdb_THK_MAX.Text, sdb_THK_MIN.Text)) Then
                
               MsgBox ("请检查厚度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub sdb_WID_MAX_KeyPress(KeyAscii As Integer)

  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub sdb_WID_MAX_Validate(Cancel As Boolean)
        
        If Len(Trim(sdb_WID_MAX.Text)) <> 0 Then
            If Not (txt_Max_Check(sdb_WID_MAX.Text, sdb_WID_MIN.Text)) Then
                
               MsgBox ("请检查宽度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub sdb_WID_MIN_KeyPress(KeyAscii As Integer)

  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub sdb_WID_MIN_Validate(Cancel As Boolean)
        
        If Len(Trim(sdb_WID_MIN.Text)) <> 0 Then
            If Not (txt_Max_Check(sdb_WID_MAX.Text, sdb_WID_MIN.Text)) Then
                
               MsgBox ("请检查宽度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If
End Sub

Private Function Data_Save_Check(ProdKnd As String) As Integer

 Data_Save_Check = 0

    If txt_STEEL_GRD.Text = "" Or Len(Trim(txt_STEEL_GRD.Text)) = 0 Then
        Data_Save_Check = 1
        Exit Function
    End If
    
    Select Case ProdKnd
        Case "S"
            If txt_MLT_STD_NO.Text = "" Or Len(Trim(txt_MLT_STD_NO.Text)) = 0 Then
                Data_Save_Check = 2
                Exit Function
            End If
        Case "H", "P"
            If txt_MLT_STD_NO.Text = "" Or Len(Trim(txt_MLT_STD_NO.Text)) = 0 Then
                Data_Save_Check = 2
                Exit Function
            End If
            
            If txt_MILL_STD_NO.Text = "" Or Len(Trim(txt_MILL_STD_NO.Text)) = 0 Then
                Data_Save_Check = 3
                Exit Function
            End If
        Case Else
            If txt_MLT_STD_NO.Text = "" Or Len(Trim(txt_MLT_STD_NO.Text)) = 0 Then
                Data_Save_Check = 2
                Exit Function
            End If
            
            If txt_MILL_STD_NO.Text = "" Or Len(Trim(txt_MILL_STD_NO.Text)) = 0 Then
                Data_Save_Check = 3
                Exit Function
            End If
    End Select
 
End Function

