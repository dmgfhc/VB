VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0250C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "外观判定标准输入_AQA0250C"
   ClientHeight    =   3435
   ClientLeft      =   2325
   ClientTop       =   4890
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   10275
   Begin VB.TextBox txt_ENDUSE_NAME 
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
      Left            =   5310
      TabIndex        =   23
      Top             =   735
      Width           =   1665
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
      Left            =   4605
      TabIndex        =   22
      Top             =   735
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10215
      TabIndex        =   19
      Top             =   0
      Width           =   10275
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   600
         Left            =   0
         TabIndex        =   20
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
            TabIndex        =   21
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
      Left            =   8865
      TabIndex        =   12
      Top             =   2970
      Width           =   1215
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
      Left            =   3780
      TabIndex        =   11
      Top             =   2925
      Width           =   1215
   End
   Begin VB.TextBox txt_SHAPE_GRD_NAME 
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
      Left            =   2340
      MaxLength       =   80
      TabIndex        =   10
      Top             =   2265
      Width           =   6435
   End
   Begin VB.TextBox txt_SURF_GRD_NAME 
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
      Left            =   2310
      MaxLength       =   80
      TabIndex        =   9
      Top             =   1815
      Width           =   6465
   End
   Begin VB.TextBox txt_SHAPE_GRD 
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
      Left            =   1290
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2265
      Width           =   1035
   End
   Begin VB.TextBox txt_SURF_GRD 
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
      Left            =   1290
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1815
      Width           =   1035
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
      Left            =   8190
      MaxLength       =   18
      TabIndex        =   6
      Top             =   735
      Width           =   2055
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
      Left            =   1305
      TabIndex        =   2
      Top             =   2895
      Width           =   1215
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
      Left            =   6435
      TabIndex        =   4
      Top             =   2940
      Width           =   1215
   End
   Begin VB.TextBox txt_UPD_EMP 
      Height          =   300
      Left            =   8910
      TabIndex        =   5
      Top             =   2970
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_INS_EMP 
      Height          =   300
      Left            =   3780
      TabIndex        =   3
      Top             =   2925
      Visible         =   0   'False
      Width           =   1215
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   2625
      Top             =   2925
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
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
      Left            =   1275
      TabIndex        =   0
      Top             =   735
      Width           =   705
   End
   Begin VB.TextBox txt_PROD_KND_NAME 
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
      Left            =   1995
      TabIndex        =   1
      Top             =   735
      Width           =   1395
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   0
      Left            =   3465
      Top             =   1215
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   1
      Left            =   105
      Top             =   733
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   2
      Left            =   7050
      Top             =   733
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   3
      Left            =   3465
      Top             =   733
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   5
      Left            =   105
      Top             =   1215
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   6
      Left            =   6855
      Top             =   1215
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   9
      Left            =   150
      Top             =   1815
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "表面等级"
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
      Height          =   315
      Index           =   10
      Left            =   150
      Top             =   2265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "形状等级"
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
      Index           =   1
      Left            =   7710
      Top             =   2970
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   11
      Left            =   150
      Top             =   2895
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
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
      Height          =   315
      Index           =   4
      Left            =   5280
      Top             =   2940
      Width           =   1095
      _ExtentX        =   1931
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
   Begin CSTextLibCtl.sidbEdit sdb_LEN_MIN 
      Height          =   315
      Left            =   7980
      TabIndex        =   13
      Tag             =   "发布年度"
      Top             =   1215
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
      NumDecDigits    =   1
      NumIntDigits    =   7
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_LEN_MAX 
      Height          =   315
      Left            =   9060
      TabIndex        =   14
      Tag             =   "发布年度"
      Top             =   1215
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
      NumDecDigits    =   1
      NumIntDigits    =   7
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_MIN 
      Height          =   315
      Left            =   1260
      TabIndex        =   15
      Tag             =   "发布年度"
      Top             =   1215
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
      Left            =   2340
      TabIndex        =   16
      Tag             =   "发布年度"
      Top             =   1215
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
      Left            =   4575
      TabIndex        =   17
      Tag             =   "发布年度"
      Top             =   1215
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
   Begin CSTextLibCtl.sidbEdit sdb_WID_MAX 
      Height          =   315
      Left            =   5655
      TabIndex        =   18
      Tag             =   "发布年度"
      Top             =   1215
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9060
      Top             =   2055
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
            Picture         =   "AQA0250C.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":04B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":07D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":09C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":0AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":0D9B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9375
      Top             =   1860
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
            Picture         =   "AQA0250C.frx":124D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":154D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":162D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":1836
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0250C.frx":1BA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   45
      X2              =   10215
      Y1              =   2730
      Y2              =   2745
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   45
      X2              =   10170
      Y1              =   1665
      Y2              =   1665
   End
End
Attribute VB_Name = "AQA0250C"
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
'-- Program Name      外观判定标准输入
'-- Program ID        AQA0250C (Master-AQA0180C)
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       外观判定标准输入
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
         Call Gp_Ms_Collection(sdb_THK_MIN, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_THK_MAX, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_WID_MIN, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_WID_MAX, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_LEN_MIN, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_LEN_MAX, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SURF_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_SURF_GRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
       Call Gp_Ms_Collection(txt_SHAPE_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_SHAPE_GRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
        Call Gp_Ms_Collection(txt_ins_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_INS_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_upd_emp, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_UPD_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
     Mc1.Add Item:="AQA0240C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQA0240C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
          
'     sQuery = "select distinct StdSpec from qp_std_head"
'     Cob_STDSPEC.Clear
'
'     Call Gf_ComboAdd(M_CN1, Cob_STDSPEC, sQuery)
     
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
        
        Case "txt_SURF_GRD"             '表面等级
            sCode = "Q0050"
            Set oCodeName = txt_SURF_GRD_NAME
                
        Case "txt_SHAPE_GRD"            '形状等级
            sCode = "Q0051"
            Set oCodeName = txt_SHAPE_GRD_NAME
                            
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
    
'    sAuthority = Gf_Pgm_Authority("AQA0250C", True)

    
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

    Call AQA0240C.Form_Ref

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
    


    txt_PROD_KND_NAME.Text = ""
    txt_ENDUSE_NAME.Text = ""

    
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

Private Sub sdb_LEN_MIN_KeyPress(KeyAscii As Integer)
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


Private Sub sdb_THK_MAX_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub sdb_THK_MIN_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub sdb_WID_MAX_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub sdb_WID_MIN_KeyPress(KeyAscii As Integer)
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

