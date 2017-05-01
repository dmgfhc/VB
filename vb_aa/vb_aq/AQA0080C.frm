VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0080C 
   Caption         =   "客户特殊要求共用信息输入_AQA0080C"
   ClientHeight    =   8220
   ClientLeft      =   810
   ClientTop       =   1440
   ClientWidth     =   12390
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12390
   Begin VB.TextBox txt_HTM_SHOT_BLAST_NAME 
      Height          =   300
      Left            =   2415
      TabIndex        =   46
      Top             =   3570
      Width           =   2460
   End
   Begin VB.TextBox txt_HTM_COND_CD_1 
      Height          =   300
      Left            =   6855
      MaxLength       =   4
      TabIndex        =   45
      Top             =   3970
      Width           =   615
   End
   Begin VB.TextBox txt_HTM_COND_CD_2 
      Height          =   300
      Left            =   6855
      MaxLength       =   4
      TabIndex        =   44
      Top             =   4379
      Width           =   615
   End
   Begin VB.TextBox txt_HTM_COND_NAME_2 
      Height          =   300
      Left            =   7515
      TabIndex        =   43
      Top             =   4379
      Width           =   4755
   End
   Begin VB.TextBox txt_HTM_METH_NAME_2 
      Height          =   300
      Left            =   2415
      TabIndex        =   42
      Top             =   4379
      Width           =   2460
   End
   Begin VB.TextBox txt_HTM_METH_CD_2 
      Height          =   300
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   41
      Top             =   4379
      Width           =   390
   End
   Begin VB.TextBox txt_HTM_COND_CD_3 
      Height          =   300
      Left            =   6855
      MaxLength       =   4
      TabIndex        =   40
      Top             =   4788
      Width           =   615
   End
   Begin VB.TextBox txt_HTM_COND_NAME_3 
      Height          =   300
      Left            =   7515
      TabIndex        =   39
      Top             =   4788
      Width           =   4755
   End
   Begin VB.TextBox txt_HTM_COND_NAME_1 
      Height          =   300
      Left            =   7515
      TabIndex        =   38
      Top             =   3970
      Width           =   4755
   End
   Begin VB.TextBox txt_HTM_METH_NAME_3 
      Height          =   300
      Left            =   2415
      TabIndex        =   37
      Top             =   4788
      Width           =   2460
   End
   Begin VB.TextBox txt_HTM_METH_NAME_1 
      Height          =   300
      Left            =   2415
      TabIndex        =   36
      Top             =   3970
      Width           =   2460
   End
   Begin VB.TextBox txt_HTM_METH_CD_3 
      Height          =   300
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   35
      Top             =   4788
      Width           =   390
   End
   Begin VB.TextBox txt_HTM_METH_CD_1 
      Height          =   300
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   34
      Top             =   3970
      Width           =   390
   End
   Begin VB.TextBox txt_HTM_SHOT_BLAST 
      Height          =   300
      Left            =   1980
      MaxLength       =   2
      TabIndex        =   33
      Top             =   3570
      Width           =   390
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12330
      TabIndex        =   30
      Top             =   0
      Width           =   12390
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   600
         Left            =   0
         TabIndex        =   31
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
            TabIndex        =   32
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
   Begin VB.TextBox txt_UPD_NAME 
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
      Left            =   6855
      TabIndex        =   28
      Top             =   7500
      Width           =   2775
   End
   Begin VB.TextBox txt_INS_NAME 
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
      Left            =   1995
      TabIndex        =   27
      Top             =   7500
      Width           =   2775
   End
   Begin VB.PictureBox Img_DRT_CNF_TYP 
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
      Height          =   310
      Left            =   3435
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   5197
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt_DRT_CNF_TYP 
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
      Left            =   1995
      MaxLength       =   1
      TabIndex        =   26
      Top             =   5197
      Width           =   315
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
      Left            =   6855
      MaxLength       =   6
      TabIndex        =   8
      Top             =   3028
      Width           =   1725
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
      Left            =   1995
      MaxLength       =   6
      TabIndex        =   7
      Top             =   3028
      Width           =   1725
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
      Left            =   6855
      MaxLength       =   8
      TabIndex        =   6
      Top             =   2619
      Width           =   1725
   End
   Begin VB.TextBox txt_DEV_STD_CD 
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
      Left            =   1995
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2619
      Width           =   1725
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
      Left            =   1995
      MaxLength       =   18
      TabIndex        =   4
      Top             =   2210
      Width           =   2745
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
      Height          =   300
      Left            =   1995
      TabIndex        =   3
      Top             =   1677
      Width           =   1035
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
      Height          =   300
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1677
      Width           =   3795
   End
   Begin VB.TextBox txt_PROD_CD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1995
      TabIndex        =   1
      Top             =   1268
      Width           =   495
   End
   Begin VB.TextBox txt_PROD_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2550
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1268
      Width           =   2265
   End
   Begin VB.TextBox txt_STEEL_GRD_Name 
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
      Left            =   8205
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1268
      Width           =   4095
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
      Left            =   6855
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1268
      Width           =   1335
   End
   Begin VB.TextBox txt_upd_emp 
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
      Left            =   6855
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txt_upd_date 
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
      Left            =   6855
      TabIndex        =   24
      Top             =   7081
      Width           =   2775
   End
   Begin VB.TextBox txt_ins_emp 
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
      Left            =   1995
      TabIndex        =   23
      Top             =   7830
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txt_ins_date 
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
      Left            =   1995
      TabIndex        =   22
      Top             =   7081
      Width           =   2775
   End
   Begin VB.TextBox txt_CUST_SQ 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3645
      TabIndex        =   16
      Top             =   735
      Width           =   1035
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
      Height          =   300
      Left            =   1995
      TabIndex        =   0
      Top             =   735
      Width           =   1635
   End
   Begin VB.TextBox txt_CUST_SPEC_DETAIL 
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
      Left            =   1995
      MaxLength       =   200
      TabIndex        =   15
      Top             =   6548
      Width           =   7635
   End
   Begin VB.TextBox txt_CUST_NAME 
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
      Height          =   300
      Left            =   4695
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   735
      Width           =   4215
   End
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   1
      Left            =   105
      Top             =   735
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "客户特殊要求编号"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   0
      Left            =   4965
      Top             =   1268
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   2
      Left            =   105
      Top             =   1677
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   3
      Left            =   105
      Top             =   2210
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   4
      Left            =   4965
      Top             =   2210
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "发布年度"
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
      Height          =   300
      Index           =   5
      Left            =   105
      Top             =   2619
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "代表性交付条件标准"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   6
      Left            =   105
      Top             =   5730
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   7
      Left            =   105
      Top             =   6139
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   8
      Left            =   4965
      Top             =   6139
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   9
      Left            =   105
      Top             =   6548
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "适用客户"
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
      Height          =   300
      Index           =   10
      Left            =   105
      Top             =   7081
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "编制日期"
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
      Height          =   300
      Index           =   11
      Left            =   105
      Top             =   7500
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "编制人"
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
      Height          =   300
      Index           =   12
      Left            =   4965
      Top             =   7081
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   13
      Left            =   4965
      Top             =   7500
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   14
      Left            =   105
      Top             =   1268
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   15
      Left            =   4965
      Top             =   2619
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "企标编号"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   16
      Left            =   105
      Top             =   3028
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "炼钢／连铸规程编号"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   17
      Left            =   4965
      Top             =   3028
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   18
      Left            =   105
      Top             =   5197
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "直接投入"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9105
      Top             =   2100
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
            Picture         =   "AQA0080C.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":04B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":07D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":09C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":0AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":0D9B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9105
      Top             =   2670
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
            Picture         =   "AQA0080C.frx":124D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":154D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":162D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":1836
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0080C.frx":1BA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_STDSPEC_YY 
      Height          =   310
      Left            =   6855
      TabIndex        =   29
      Tag             =   "发布年度"
      Top             =   2210
      Width           =   870
      _Version        =   262145
      _ExtentX        =   1535
      _ExtentY        =   547
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
      NumDecDigits    =   0
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_WID_MIN 
      Height          =   315
      Left            =   1995
      TabIndex        =   11
      Top             =   6135
      Width           =   1170
      _Version        =   262145
      _ExtentX        =   2064
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
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
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
   Begin CSTextLibCtl.sidbEdit sdb_LEN_MIN 
      Height          =   315
      Left            =   6825
      TabIndex        =   13
      Top             =   6139
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
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
      NumIntDigits    =   7
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_MAX 
      Height          =   315
      Left            =   3225
      TabIndex        =   10
      Top             =   5730
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
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
   Begin CSTextLibCtl.sidbEdit sdb_WID_MAX 
      Height          =   315
      Left            =   3225
      TabIndex        =   12
      Top             =   6139
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
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
   Begin CSTextLibCtl.sidbEdit sdb_LEN_MAX 
      Height          =   315
      Left            =   8055
      TabIndex        =   14
      Top             =   6139
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
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
      NumIntDigits    =   7
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_MIN 
      Height          =   315
      Left            =   1995
      TabIndex        =   9
      Top             =   5730
      Width           =   1200
      _Version        =   262145
      _ExtentX        =   2117
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   19
      Left            =   105
      Top             =   3561
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "抛丸代码"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   20
      Left            =   105
      Top             =   3970
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "热处理方法 1"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   22
      Left            =   4965
      Top             =   3970
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "热处理条件 1"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   24
      Left            =   105
      Top             =   4788
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "热处理方法 3"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   25
      Left            =   4965
      Top             =   4788
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "热处理条件 3"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   21
      Left            =   105
      Top             =   4379
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "热处理方法 2"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   23
      Left            =   4965
      Top             =   4379
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "热处理条件 2"
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
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   105
      X2              =   12390
      Y1              =   3437
      Y2              =   3437
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   105
      X2              =   12390
      Y1              =   5606
      Y2              =   5606
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   105
      X2              =   12405
      Y1              =   2086
      Y2              =   2086
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   105
      X2              =   12390
      Y1              =   6957
      Y2              =   6957
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   105
      X2              =   12390
      Y1              =   1144
      Y2              =   1144
   End
End
Attribute VB_Name = "AQA0080C"
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
'-- Program Name      客户特殊要求共用信息输入
'-- Program ID        AQA0080C (Master-AQA0071C)
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       客户特殊要求共用信息输入
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
Dim select_text As String

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "PopMaster"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary )", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_CUST_CD, "p", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_CUST_SQ, "p", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PROD_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PROD_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STEEL_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STEEL_GRD_Name, " ", " ", " ", " ", "r", "l", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ENDUSE_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ENDUSE_NAME, " ", " ", " ", "", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STDSPEC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_STDSPEC_YY, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_DEV_STD_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_NISCO_QUALITY_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_MLT_STD_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_MILL_STD_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
       Call Gp_Ms_Collection(txt_HTM_SHOT_BLAST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_METH_CD_1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_METH_NAME_1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_COND_CD_1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_COND_NAME_1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_METH_CD_2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_METH_NAME_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_COND_CD_2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_COND_NAME_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_METH_CD_3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_METH_NAME_3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_COND_CD_3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HTM_COND_NAME_3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
       Call Gp_Ms_Collection(txt_DRT_CNF_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_THK_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_THK_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_WID_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_WID_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_LEN_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_LEN_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_CUST_SPEC_DETAIL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ins_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_INS_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_upd_emp, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_UPD_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
     Mc1.Add Item:="AQA0070C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQA0070C.P_REFER", Key:="P-R"
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
        
        Case "txt_CUST_CD"      '客户特殊要求编号
            sCode = "CUST_CD"
            Set oCodeName = txt_CUST_NAME
            
        Case "txt_PROD_CD"      '产品
            sCode = "B0005"
            Set oCodeName = txt_PROD_NAME
            
        Case "txt_ENDUSE_CD"    '订单用途
            sCode = "ENDUSE_CD"
            Set oCodeName = txt_ENDUSE_NAME
            DD.sKey = Left(txt_PROD_CD.Text, 1)
            
        Case "txt_STEEL_GRD"    '钢种
            sCode = "STLGRD"
            Set oCodeName = txt_STEEL_GRD_Name
        
        Case "txt_MLT_STD_NO"           '炼钢规程编号
            sCode = "MLT_STD_NO"
            
        Case "txt_MILL_STD_NO"          '轧钢规程编号
            sCode = "MILL_STD_NO"
            
        Case "txt_NISCO_QUALITY_NO"     '企标编号
            sCode = "NISCO_QUALITY_NO"
        
        Case "txt_HTM_SHOT_BLAST"       '抛丸代码
            sCode = "Q0074"
            Set oCodeName = txt_HTM_SHOT_BLAST_NAME
        Case "txt_HTM_METH_CD_1"
        
            sCode = "Q0073"             '热处理方法
            Set oCodeName = txt_HTM_METH_NAME_1
        
        Case "txt_HTM_METH_CD_2"        '热处理方法
            sCode = "Q0073"
            Set oCodeName = txt_HTM_METH_NAME_2

        Case "txt_HTM_METH_CD_3"        '热处理方法
            sCode = "Q0073"
            Set oCodeName = txt_HTM_METH_NAME_3

        Case "txt_HTM_COND_CD_1"        '热处理条件
            sCode = "HTM_COND_CD"
            Set oCodeName = txt_HTM_COND_NAME_1
            
            
        Case "txt_HTM_COND_CD_2"        '热处理条件
            sCode = "HTM_COND_CD"
            Set oCodeName = txt_HTM_COND_NAME_2
    
        
        Case "txt_HTM_COND_CD_3"        '热处理条件
            sCode = "HTM_COND_CD"
            Set oCodeName = txt_HTM_COND_NAME_3
    
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub Form_Activate()

    If Mc1("pControl").Item(1).Text = "" Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
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

    Call AQA0070C.Form_Ref

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
    txt_STEEL_GRD_Name = ""
'    Cob_stdspec_yy.Clear
    
    
    pControl(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then MenuTool.Buttons(4).Enabled = False   'Delete
    
End Sub

Public Sub Form_Pro()
   
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        txt_ins_emp.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call Popup_Menu_Setting
        End If
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then
        Call Popup_Menu_Setting
    End If
    
End Sub

Private Sub Img_DRT_CNF_TYP_Click()
    If txt_DRT_CNF_TYP.Text = "Y" Or txt_DRT_CNF_TYP.Text = "y" Then
        txt_DRT_CNF_TYP.Text = "N"
    Else
        txt_DRT_CNF_TYP.Text = "Y"
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

Private Sub txt_CUST_CD_Change()
'    txt_CUST_SQ.Text = Right(txt_CUST_CD.Text, 3)
End Sub

Private Sub txt_DEV_STD_CD_KeyUp(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyF4 Then
        
            DD.sWitch = "MS"
            DD.rControl.Add Item:=txt_DEV_STD_CD
            
            Call Gf_STD_DELV_DD(M_CN1, KeyCode)
            
            Exit Sub
        
        End If

End Sub

Private Sub txt_DRT_CNF_TYP_Change()
    If txt_DRT_CNF_TYP.Text = "Y" Or txt_DRT_CNF_TYP.Text = "y" Then
      ' Img_DRT_CNF_TYP.Picture = ImageList1.ListImages(9).Picture
    Else
      ' Img_DRT_CNF_TYP.Picture = Nothing
    End If
End Sub


'Private Sub sdb_LEN_MAX_KeyPress(KeyAscii As Integer)
'
'  KeyAscii = txt_KeyPress(KeyAscii)
'
'End Sub
'
'Private Sub sdb_LEN_MAX_Validate(Cancel As Boolean)
'
'        If Len(Trim(sdb_LEN_MAX.Text)) <> 0 Then
'            If Not (txt_Max_Check(sdb_LEN_MAX.Text, sdb_LEN_MIN.Text)) Then
'
'               MsgBox ("请检查长度组最小值和最大值，后者不能小与前者")
'
'               Cancel = True
'
'            End If
'
'        Else
'               MsgBox ("请输入数值")
'
'               Cancel = True
'
'        End If
'
'End Sub
'
'Private Sub sdb_LEN_MIN_KeyPress(KeyAscii As Integer)
'
'  KeyAscii = txt_KeyPress(KeyAscii)
'
'End Sub
'
'Private Sub sdb_LEN_MIN_Validate(Cancel As Boolean)
'
'        If Len(Trim(sdb_LEN_MIN.Text)) <> 0 Then
'            If Not (txt_Max_Check(sdb_LEN_MAX.Text, sdb_LEN_MIN.Text)) Then
'
'               MsgBox ("请检查长度组最小值和最大值，后者不能小与前者")
'
'               Cancel = True
'
'            End If
'
'        Else
'               MsgBox ("请输入数值")
'
'               Cancel = True
'
'        End If
'
'End Sub
'
'Private Sub txt_MILL_STD_NO_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.rControl.Add Item:=txt_MILL_STD_NO
'
'        Call Gf_Roll_STD_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'End Sub
'
'Private Sub txt_MLT_STD_NO_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.rControl.Add Item:=txt_MLT_STD_NO
'
'        Call Gf_Melt_STD_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'End Sub

'Private Sub txt_NISCO_QUALITY_NO_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.rControl.Add Item:=txt_NISCO_QUALITY_NO
'
'        Call Gf_Nisco_STD_DD(M_CN1, KeyCode)
'
'
'        Exit Sub
'
'    End If
'
'End Sub



Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_STDSPEC
        DD.rControl.Add Item:=sdb_STDSPEC_YY
        
        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        
        Exit Sub
    
    End If
End Sub


'Private Sub sdb_THK_MAX_KeyPress(KeyAscii As Integer)
'
'  KeyAscii = txt_KeyPress(KeyAscii)
'
'End Sub
'
'Private Sub sdb_THK_MAX_Validate(Cancel As Boolean)
'
'        If Len(Trim(sdb_THK_MAX.Text)) <> 0 Then
'            If Not (txt_Max_Check(sdb_THK_MAX.Text, sdb_THK_MIN.Text)) Then
'
'               MsgBox ("请检查厚度组最小值和最大值，后者不能小与前者")
'
'               Cancel = True
'
'            End If
'
'        Else
'               MsgBox ("请输入数值")
'
'               Cancel = True
'
'        End If
'
'End Sub

'Private Sub sdb_THK_MIN_KeyPress(KeyAscii As Integer)
'
'  KeyAscii = txt_KeyPress(KeyAscii)
'
'End Sub


'Private Function txt_KeyPress(KeyAscii As Integer) As Integer
'
'        Select Case KeyAscii
'
'               Case Is <= 32
'                    txt_KeyPress = KeyAscii
'               Case 48 To 57
'                    txt_KeyPress = KeyAscii
'               Case 46
'                    txt_KeyPress = KeyAscii
'               Case Else
'                    txt_KeyPress = 0
'        End Select
'
'
'End Function

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

'Private Sub sdb_THK_MIN_Validate(Cancel As Boolean)
'
'        If Len(Trim(sdb_THK_MIN.Text)) <> 0 Then
'            If Not (txt_Max_Check(sdb_THK_MAX.Text, sdb_THK_MIN.Text)) Then
'
'               MsgBox ("请检查厚度组最小值和最大值，后者不能小与前者")
'
'               Cancel = True
'
'            End If
'
'        Else
'               MsgBox ("请输入数值")
'
'               Cancel = True
'
'        End If
'
'End Sub
'
'Private Sub sdb_WID_MAX_KeyPress(KeyAscii As Integer)
'
'  KeyAscii = txt_KeyPress(KeyAscii)
'
'End Sub
'
'Private Sub sdb_WID_MAX_Validate(Cancel As Boolean)
'
'        If Len(Trim(sdb_WID_MAX.Text)) <> 0 Then
'            If Not (txt_Max_Check(sdb_WID_MAX.Text, sdb_WID_MIN.Text)) Then
'
'               MsgBox ("请检查宽度组最小值和最大值，后者不能小与前者")
'
'               Cancel = True
'
'            End If
'
'        Else
'               MsgBox ("请输入数值")
'
'               Cancel = True
'
'        End If
'
'End Sub
'
'Private Sub sdb_WID_MIN_KeyPress(KeyAscii As Integer)
'
'  KeyAscii = txt_KeyPress(KeyAscii)
'
'End Sub
'
'Private Sub sdb_WID_MIN_Validate(Cancel As Boolean)
'
'        If Len(Trim(sdb_WID_MIN.Text)) <> 0 Then
'            If Not (txt_Max_Check(sdb_WID_MAX.Text, sdb_WID_MIN.Text)) Then
'
'               MsgBox ("请检查宽度组最小值和最大值，后者不能小与前者")
'
'               Cancel = True
'
'            End If
'
'        Else
'               MsgBox ("请输入数值")
'
'               Cancel = True
'
'        End If
'End Sub


