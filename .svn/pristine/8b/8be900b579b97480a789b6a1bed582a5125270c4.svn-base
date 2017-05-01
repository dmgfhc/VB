VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "发货管理"
   ClientHeight    =   7560
   ClientLeft      =   4335
   ClientTop       =   3090
   ClientWidth     =   9990
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Tag             =   "H"
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet 
      Left            =   4800
      Top             =   2295
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9930
      TabIndex        =   0
      Top             =   0
      Width           =   9990
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   600
         Left            =   0
         TabIndex        =   1
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
            TabIndex        =   2
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
               NumButtons      =   17
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Clear"
                  Object.ToolTipText     =   "空界面"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Refer"
                  Object.ToolTipText     =   "查询"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line1"
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Save"
                  Object.ToolTipText     =   "保存"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Delete"
                  Object.ToolTipText     =   "删除"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line2"
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "RowIns"
                  Object.ToolTipText     =   "追加行"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "RowDel"
                  Object.ToolTipText     =   "删除行"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "RowCan"
                  Object.ToolTipText     =   "取消行"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line3"
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Copy"
                  Object.ToolTipText     =   "复制"
                  ImageIndex      =   8
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   3
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Acopy"
                        Text            =   "Screen Copy"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Mcopy"
                        Text            =   "Master Copy"
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Scopy"
                        Text            =   "Spread Copy"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Paste"
                  Object.ToolTipText     =   "粘贴"
                  ImageIndex      =   9
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   3
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Apaste"
                        Text            =   "Screen Paste"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Mpaste"
                        Text            =   "Master Paste"
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Spaste"
                        Text            =   "Spread Paste"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line4"
                  Style           =   3
               EndProperty
               BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Excel"
                  Object.ToolTipText     =   "导出"
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Print"
                  Object.ToolTipText     =   "打印"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line5"
                  Style           =   3
               EndProperty
               BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Exit"
                  Object.ToolTipText     =   "退出"
                  ImageIndex      =   12
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   1965
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":121F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":12FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":1508
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":16CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":1888
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":1ACD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":1C05
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":1E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":1F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":2196
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":24A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":2960
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":2C63
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":2F83
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":316C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":32BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":3405
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":3592
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":367C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":396B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":3A77
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiMain.frx":3D4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   7095
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12876
            MinWidth        =   12876
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1059
            MinWidth        =   1059
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1059
            MinWidth        =   1059
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1059
            MinWidth        =   1059
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2013-11-27"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "15:04"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3352
            MinWidth        =   3352
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
            Picture         =   "MdiMain.frx":41FE
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Mnu_Control 
      Caption         =   "Control"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Clear 
         Caption         =   "Clear"
      End
      Begin VB.Menu Mnu_Refer 
         Caption         =   "Refer"
      End
      Begin VB.Menu Mnu_Save 
         Caption         =   "Save"
      End
      Begin VB.Menu Mnu_Delete 
         Caption         =   "Del"
      End
      Begin VB.Menu Mnu_RowIns 
         Caption         =   "RowIns"
      End
      Begin VB.Menu Mnu_RowDel 
         Caption         =   "RowDel"
      End
      Begin VB.Menu Mnu_RowCan 
         Caption         =   "RowCan"
      End
      Begin VB.Menu Mnu_Copy 
         Caption         =   "Copy"
         Begin VB.Menu Mnu_Acopy 
            Caption         =   "Acopy"
         End
         Begin VB.Menu Mnu_Mcopy 
            Caption         =   "Mcopy"
         End
         Begin VB.Menu Mnu_Scopy 
            Caption         =   "Scopy"
         End
      End
      Begin VB.Menu Mnu_Paste 
         Caption         =   "Paste"
         Begin VB.Menu Mnu_Apaste 
            Caption         =   "Apaste"
         End
         Begin VB.Menu Mnu_Mpaste 
            Caption         =   "Mpaste"
         End
         Begin VB.Menu Mnu_Spaste 
            Caption         =   "Spaste"
         End
      End
      Begin VB.Menu Mnu_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu Mnu_Print 
         Caption         =   "Print"
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu PopUp_Spread 
      Caption         =   "PopUp-Spread"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Sorting 
         Caption         =   "Columns Sorting"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_FrozenSetting 
         Caption         =   "Columns Frozen Setting"
      End
      Begin VB.Menu Mnu_FrozenCancel 
         Caption         =   "Columns Frozen Cancel"
      End
   End
   Begin VB.Menu Mnu_HStandard 
      Caption         =   "发货标准管理"
      Begin VB.Menu Mnu_AHA0010C 
         Caption         =   "发货能力管理"
      End
      Begin VB.Menu Mnu_AHA0020C 
         Caption         =   "车辆信息录入"
      End
   End
   Begin VB.Menu Mnu_HProcess 
      Caption         =   "发货进程管理"
      Begin VB.Menu Mnu_AHC0130C 
         Caption         =   "退货处理"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHC0110C 
         Caption         =   "产品库存现状"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHC0120C 
         Caption         =   "产品信息修改"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHC0200C 
         Caption         =   "振华发货明细"
      End
   End
   Begin VB.Menu Mnu_SLAB_SALESProcess 
      Caption         =   "板坯发货管理"
      Visible         =   0   'False
      Begin VB.Menu Mnu_AHG0010C 
         Caption         =   "板坯发货实绩"
      End
      Begin VB.Menu Mnu_AHG0020C 
         Caption         =   "板坯发货实绩取消"
      End
      Begin VB.Menu Mnu_AHG0030C 
         Caption         =   "板坯计量重量录入"
      End
   End
   Begin VB.Menu Mnu_HStatistics 
      Caption         =   "发货统计管理"
      Begin VB.Menu Mnu_AHD0010C 
         Caption         =   "日入库实绩查询"
      End
      Begin VB.Menu Mnu_AHD0020C 
         Caption         =   "日出库实绩查询"
      End
      Begin VB.Menu Mnu_AHD0100C 
         Caption         =   "经营部 出入库月平衡报表(综判)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHD0130C 
         Caption         =   "厂库别 出入库月平衡报表(综判)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHD0120C 
         Caption         =   "经营部 库存收发存报表(综判)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHD0140C 
         Caption         =   "厂库别 库存收发存报表(综判)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu__AHD0180C 
         Caption         =   "经营部 出入库月平衡报表(入库)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHD0160C 
         Caption         =   "厂库别 出入库月平衡报表(入库)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu__AHD0170C 
         Caption         =   "经营部 库存收发存报表(入库)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu__AHD0210C 
         Caption         =   "厂库别 库存收发存报表(入库)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu__AHD0280C 
         Caption         =   "厂库别 库存收发存报表（中板入库）"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu__AHD0270C 
         Caption         =   "厂库别 配送库库存（入库）"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHD0110C 
         Caption         =   "中厚板卷厂板材产量(综判)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AHD0220C 
         Caption         =   "仓库库存报表（订/余）"
      End
      Begin VB.Menu Mnu_AHD0240C 
         Caption         =   "仓库库存明细表（订/余）"
      End
      Begin VB.Menu Mnu_AHD0260C 
         Caption         =   "销售总公司正品材收发存日报表"
      End
      Begin VB.Menu Mnu_AHD0230C 
         Caption         =   "发货台帐"
      End
      Begin VB.Menu Mnu_AHD0190C 
         Caption         =   "出库实绩查询"
      End
      Begin VB.Menu Mnu_AHD0250S 
         Caption         =   "年入出库实绩发放"
      End
   End
   Begin VB.Menu Mnu_Report 
      Caption         =   "统计报表"
      Begin VB.Menu Mnu_AHD0400C 
         Caption         =   "销售总公司出入库月平衡报表_AHD0400C"
      End
      Begin VB.Menu Mnu_AHD0430C 
         Caption         =   "厂库别出入库月平衡报表_AHD0430C"
      End
      Begin VB.Menu Mnu_AHD0420C 
         Caption         =   "销售总公司库存收发存报表_AHD0420C"
      End
      Begin VB.Menu Mnu_AHD0440C 
         Caption         =   "厂库别库存收发存报表_AHD0440C"
      End
      Begin VB.Menu Mnu_AHD0450C 
         Caption         =   "厂库别库存收发存报表(新)_AHD0450C"
      End
      Begin VB.Menu Mnu_AHD0451C 
         Caption         =   "厂库别库存收发存质量报表_AHD0451C"
      End
      Begin VB.Menu Mnu_AHD0510C 
         Caption         =   "产成品转库报表_AHD0510C"
      End
      Begin VB.Menu Mnu_AHD0520C 
         Caption         =   "未入库板材库存情况报表_AHD0520C"
      End
      Begin VB.Menu Mnu_AHD0530C 
         Caption         =   "营销部板材对应工序10/30/60天以上未作处理进程情况报表_AHD0530C"
      End
      Begin VB.Menu Mnu_AHD0540C 
         Caption         =   "热处理订单进度跟踪表_AHD0540C"
      End
   End
   Begin VB.Menu Mnu_ZHIBAO 
      Caption         =   "质保书查询"
      Begin VB.Menu Mnu_AHQ0010C 
         Caption         =   "质保书发放信息查询"
      End
      Begin VB.Menu Mnu_AHQ0040C 
         Caption         =   "材质试验实绩确认"
      End
      Begin VB.Menu Mnu_AHQ0050C 
         Caption         =   "钢板(卷)--质保书信息查询"
      End
   End
   Begin VB.Menu Mnu_Windows 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu Mnu_Horiz 
         Caption         =   "Tile Horiz"
      End
      Begin VB.Menu Mnu_Vertical 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu Mnu_Cascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu Mnu_HELP 
         Caption         =   "界面说明书"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Public Sub FormMenuSetting(Fm As Variant, FormType As String, ButtonType As String, sAuthority As String)

On Error Resume Next
    
    With MenuTool
    
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
                    
                Case "Master"
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
                    .Buttons(11).Enabled = True                 'Copy
                    .Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
                    .Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
                    .Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy
                    
                    .Buttons(12).Enabled = True                 'Paste
                    .Buttons(12).ButtonMenus(1).Enabled = False 'All Paste
                    .Buttons(12).ButtonMenus(2).Enabled = False 'Master Paste
                    .Buttons(12).ButtonMenus(3).Enabled = False 'Spread Paste
                    
                    .Buttons(13).Enabled = True                 'Separator
                    .Buttons(14).Enabled = False                'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = True                 'Separator
                    .Buttons(17).Enabled = True                 'Exit
                
                Case "Sheet", "Msheet"
                    .Buttons(1).Enabled = True                  'Screen Clear
                    .Buttons(2).Enabled = True                  'Refer
                    .Buttons(3).Enabled = True                  'Separator
                    .Buttons(4).Enabled = True                  'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = True                  'Separator
                    .Buttons(7).Enabled = True                  'Row Insert
                    .Buttons(8).Enabled = False                 'Row Delete
                    .Buttons(9).Enabled = True                  'Row Cancel
                    .Buttons(10).Enabled = True                 'Separator
                    
                    .Buttons(11).Enabled = True                 'Copy
                    .Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
                    .Buttons(11).ButtonMenus(2).Enabled = False 'Master Copy
                    .Buttons(11).ButtonMenus(3).Enabled = True  'Spread Copy
                    
                    .Buttons(12).Enabled = True                 'Paste
                    .Buttons(12).ButtonMenus(1).Enabled = False 'All Paste
                    .Buttons(12).ButtonMenus(2).Enabled = False 'Master Paste
                    .Buttons(12).ButtonMenus(3).Enabled = False 'Spread Paste
                    
                    .Buttons(13).Enabled = True                 'Separator
                    .Buttons(14).Enabled = True                 'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = True                 'Separator
                    .Buttons(17).Enabled = True                 'Exit
                
                Case "PopSheet"
                    .Buttons(1).Enabled = True                  'Screen Clear
                    .Buttons(2).Enabled = True                  'Refer
                    .Buttons(3).Enabled = True                  'Separator
                    .Buttons(4).Enabled = False                 'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = True                  'Separator
                    .Buttons(7).Enabled = False                 'Row Insert
                    .Buttons(8).Enabled = False                 'Row Delete
                    .Buttons(9).Enabled = False                 'Row Cancel
                    .Buttons(10).Enabled = True                 'Separator
                    .Buttons(11).Enabled = False                'Copy
                    .Buttons(12).Enabled = False                'Paste
                    .Buttons(13).Enabled = True                 'Separator
                    .Buttons(14).Enabled = False                'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = True                 'Separator
                    .Buttons(17).Enabled = True                 'Exit
                
                Case "Hsheet"
                    .Buttons(1).Enabled = True                  'Screen Clear
                    .Buttons(2).Enabled = True                  'Refer
                    .Buttons(3).Enabled = True                  'Separator
                    .Buttons(4).Enabled = True                  'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = True                  'Separator
                    .Buttons(7).Enabled = True                  'Row Insert
                    .Buttons(8).Enabled = False                 'Row Delete
                    .Buttons(9).Enabled = True                  'Row Cancel
                    .Buttons(10).Enabled = True                 'Separator
                    
                    .Buttons(11).Enabled = True                 'Copy
                    .Buttons(11).ButtonMenus(1).Enabled = True  'All Copy
                    .Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
                    .Buttons(11).ButtonMenus(3).Enabled = True  'Spread Copy
                    
                    .Buttons(12).Enabled = True                 'Paste
                    .Buttons(12).ButtonMenus(1).Enabled = False 'All Paste
                    .Buttons(12).ButtonMenus(2).Enabled = False 'Master Paste
                    .Buttons(12).ButtonMenus(3).Enabled = False 'Spread Paste
                    
                    .Buttons(13).Enabled = True                 'Separator
                    .Buttons(14).Enabled = False                'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = True                 'Separator
                    .Buttons(17).Enabled = True                 'Exit
                
                Case "Refer"
                    .Buttons(1).Enabled = True                  'Screen Clear
                    .Buttons(2).Enabled = True                  'Refer
                    .Buttons(3).Enabled = True                  'Separator
                    .Buttons(4).Enabled = False                 'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = True                  'Separator
                    .Buttons(7).Enabled = False                 'Row Insert
                    .Buttons(8).Enabled = False                 'Row Delete
                    .Buttons(9).Enabled = False                 'Row Cancel
                    .Buttons(10).Enabled = True                 'Separator
                    .Buttons(11).Enabled = False                'Copy
                    .Buttons(12).Enabled = False                'Paste
                    .Buttons(13).Enabled = True                 'Separator
                    .Buttons(14).Enabled = False                'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = True                 'Separator
                    .Buttons(17).Enabled = True                 'Exit
                
        End Select
        
        Fm.Toolbar_St = ButtonType
                 
        .Wrappable = True
        
        Call MenuStatus(FormType, ButtonType, sAuthority)
        
    End With
    
End Sub
       
Public Sub MenuStatus(FormType As String, ButtonType As String, sAuthority As String)

    With MenuTool
    
        Select Case ButtonType
                 'Save, Refer
            Case "SE", "RE"
                
                Select Case FormType
                
                    Case "Master"
                        .Buttons(5).Enabled = True              'Delete
                        
                    Case "Sheet", "Msheet"
                        .Buttons(7).Enabled = True              'Row Insert
                        .Buttons(8).Enabled = True              'Row Delete
                        .Buttons(9).Enabled = True              'Row Cancel
                        .Buttons(14).Enabled = True             'Excel
                    
                    Case "PopSheet"
                        .Buttons(14).Enabled = True             'Excel
                        
                    Case "Hsheet"
                        .Buttons(5).Enabled = True              'Delete
                        .Buttons(7).Enabled = True              'Row Insert
                        .Buttons(8).Enabled = True              'Row Delete
                        .Buttons(9).Enabled = True              'Row Cancel
                        .Buttons(14).Enabled = True             'Excel
                    
                    Case "Refer"
                        .Buttons(14).Enabled = True             'Excel
                        .Buttons(15).Enabled = False            'Print
                    
                End Select
                
                 'Form Start, Screen Clear
            Case "FS", "CLS"
                
                Select Case FormType

                    Case "Master"
                        .Buttons(5).Enabled = False             'Delete
                        
                    Case "Sheet", "Msheet"
                        .Buttons(7).Enabled = True              'Row Insert
                        .Buttons(8).Enabled = False             'Row Delete
                        .Buttons(9).Enabled = True              'Row Cancel
                        .Buttons(14).Enabled = False            'Excel
                    
                    Case "PopSheet"
                        .Buttons(14).Enabled = False            'Excel
                        
                    Case "Hsheet"
                        .Buttons(5).Enabled = False             'Delete
                        .Buttons(7).Enabled = True              'Row Insert
                        .Buttons(8).Enabled = False             'Row Delete
                        .Buttons(9).Enabled = True              'Row Cancel
                        .Buttons(14).Enabled = False            'Excel
                    
                    Case "Refer"
                        .Buttons(14).Enabled = False            'Excel
                        .Buttons(15).Enabled = False            'Print
                    
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



Private Sub MDIForm_Activate()

    'Call MDIMain.FormMenuSetting(me,"Start", Toolbar_St,"")

End Sub

Private Sub MDIForm_Load()

   Dim Active_YN As String
    Dim args  As Variant ' 2012.11.09 新增  耿朝雷
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
    Me.BackColor = &HE0E0E0
    
    If GF_DbConnect = False Then
        
        Unload Me
    
    Else
    
    args = Split(Trim(Command), " ") ' 2012.11.09 新增  耿朝雷
    If UBound(args) = 1 Then
         MainFrmType = "New"
         sUserID = args(0) ' 2012.11.09 新增  耿朝雷
         sUserName = args(1) ' 2012.11.09 新增  耿朝雷
         MDIMain.StatusBar1.Panels(1) = "提示信息 ：" ' 2012.11.09 新增  耿朝雷
         MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName ' 2012.11.09 新增  耿朝雷
    Else
        Active_YN = GetSetting("NISCO", "EXE-FILE", "AH.exe")
        If Active_YN = "1" Then
            MainFrmType = "Old"
            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
            MDIMain.StatusBar1.Panels(1) = "提示信息 ：："
            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
        Else
            Call Gp_MsgBoxDisplay("只能从主画面登陆...", "W")
            Unload Me
            Exit Sub
        End If
    End If  ' 2012.11.09 新增  耿朝雷
    
        
'
'        sUserID = "1JS1005"
'        sUserName = "杨猛"
'        MDIMain.StatusBar1.Panels(1) = "提示信息 ："
'        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName


        If Mid(M_CN1, Len(M_CN1), 1) = "9" Then
            MDIMain.StatusBar1.Panels(8) = "正式机"
        Else
            MDIMain.StatusBar1.Panels(8) = "测试机"
        End If

    End If
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim CurrentForm As Form
Dim FormLD As Boolean

    FormLD = False
    
    For Each CurrentForm In Forms
        If CurrentForm.Name <> Me.Name Then
            FormLD = True
            Exit For
        End If
    Next CurrentForm
    
    If FormLD Then
    
        If MsgBox("Low rank program was not ended," + vbCrLf + "end Program ?", MB_YESNO _
                        + MB_ICONQUESTION, Me.Caption) = IDYES Then
            
            For Each CurrentForm In Forms
                If CurrentForm.Name <> Me.Name Then
                    Unload CurrentForm
                End If
            Next CurrentForm
            
        Else
            Cancel = True
        End If
        
    End If
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "AH.exe", ""

End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    
    If Screen.ActiveForm.Name = "MDIMain" Then
        
        If Button.Key = "Exit" Then
            If vbYes = MsgBox(Me.Caption + " Terminate ?", vbQuestion + vbYesNo, Me.Caption) Then
                Unload Me
            End If
        End If
        
        Exit Sub
       
    End If
    
    If TypeOf Screen.ActiveForm.ActiveControl Is vaSpread Then
        Call Gp_Sp_EventMake(Screen.ActiveForm.ActiveControl)
    End If
    
    Select Case Button.Key
        Case "Clear"               'Clear
            Call Mnu_Clear_Click
        Case "Refer"               'Refer
            Call Mnu_Refer_Click
        Case "Save"                'Save
            Call Mnu_Save_Click
        Case "Delete"              'Delete
            Call Mnu_Delete_Click
        Case "RowIns"              'RowIns
            Call Mnu_RowIns_Click
        Case "RowDel"              'RowDel
            Call Mnu_RowDel_Click
        Case "RowCan"              'RowCan
            Call Mnu_RowCan_Click
        Case "Excel"               'Excel
            Call Mnu_Excel_Click
        Case "Print"               'Print
            Call Mnu_Print_Click
        Case "Exit"                'Exit
            Call Mnu_Exit_Click
    End Select
        
End Sub

Private Sub MenuTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    
    Select Case ButtonMenu.Key
    
        Case "Acopy"    'All Copy
            Call Mnu_Acopy_Click
        
        Case "Mcopy"    'Master Copy
            Call Mnu_Mcopy_Click
        
        Case "Scopy"    'Spread Copy
            Call Mnu_Scopy_Click
        
        Case "Apaste"    'All Paste
            Call Mnu_Apaste_Click
        
        Case "Mpaste"    'Master Paste
            Call Mnu_Mpaste_Click
        
        Case "Spaste"    'Spread Paste
            Call Mnu_Spaste_Click
        
    End Select
    
End Sub

Private Sub Mnu__AHD0170C_Click()
    AHD0170C.Show
    AHD0170C.SetFocus
End Sub

Private Sub Mnu__AHD0180C_Click()
    AHD0180C.Show
    AHD0180C.SetFocus
End Sub

Private Sub Mnu__AHD0210C_Click()
    AHD0210C.Show
    AHD0210C.SetFocus
End Sub

Private Sub Mnu__AHD0270C_Click()
    AHD0270C.Show
    AHD0270C.SetFocus
End Sub

Private Sub Mnu__AHD0280C_Click()
    AHD0280C.Show
    AHD0280C.SetFocus
End Sub

Private Sub Mnu_AHA0010C_Click()
    AHA0010C.Show
    AHA0010C.SetFocus
End Sub

Private Sub Mnu_AHA0020C_Click()
    AHA0020C.Show
    AHA0020C.SetFocus
End Sub

Private Sub Mnu_AHC0019C_Click()
    AHC0019C.Show
    AHC0019C.SetFocus
End Sub

Private Sub Mnu_AHB0020C_Click()
    AHB0020C.Show
    AHB0020C.SetFocus
End Sub

Private Sub Mnu_AHB0030C_Click()
    AHB0030C.Show
    AHB0030C.SetFocus
End Sub

Private Sub Mnu_AHC0040C_Click()
    AHC0040C.Show
    AHC0040C.SetFocus
End Sub

Private Sub Mnu_AHC0050C_Click()
    AHC0050C.Show
    AHC0050C.SetFocus
End Sub

'Private Sub Mnu_AHC0070C_Click()
'    AHC0070C.Show
'    AHC0070C.SetFocus
'End Sub


Private Sub Mnu_AHC0110C_Click()
'    AHC0110C.Show
'    AHC0110C.SetFocus
End Sub

Private Sub Mnu_AHC0120C_Click()
'    AHC0120C.Show
'    AHC0120C.SetFocus
End Sub

Private Sub Mnu_AHC0130C_Click()
'    AHC0130C.Show
'    AHC0130C.SetFocus
End Sub

Private Sub Mnu_AHC0200C_Click()
    AHC0200C.Show
    AHC0200C.SetFocus
End Sub

Private Sub Mnu_AHD0010C_Click()
    AHD0010C.Show
    AHD0010C.SetFocus
End Sub

Private Sub Mnu_AHD0020C_Click()
    AHD0020C.Show
    AHD0020C.SetFocus
End Sub

Private Sub Mnu_AHD0100C_Click()
    AHD0100C.Show
    AHD0100C.SetFocus
End Sub

Private Sub Mnu_AHD0130C_Click()
    AHD0130C.Show
    AHD0130C.SetFocus
End Sub

Private Sub Mnu_AHD0110C_Click()
    AHD0110C.Show
    AHD0110C.SetFocus
End Sub

Private Sub Mnu_AHD0120C_Click()
    AHD0120C.Show
    AHD0120C.SetFocus
End Sub

Private Sub Mnu_AHD0140C_Click()
    AHD0140C.Show
    AHD0140C.SetFocus
End Sub

Private Sub Mnu_AHD0150C_Click()

End Sub

'Private Sub Mnu_AHD0150C_Click()
'    AHD0150C.Show
'    AHD0150C.SetFocus
'End Sub

Private Sub Mnu_AHD0160C_Click()
    AHD0160C.Show
    AHD0160C.SetFocus
End Sub

'Private Sub Mnu_AHD0170C_Click()
'    AHD0170C.Show
'    AHD0170C.SetFocus
'End Sub

Private Sub Mnu_AHD0200C_Click()
    AHD0200C.Show
    AHD0200C.SetFocus
End Sub

Private Sub Mnu_AHD0180C_Click()

End Sub

Private Sub Mnu_AHD0190C_Click()
    AHD0190C.Show
    AHD0190C.SetFocus
End Sub

Private Sub Mnu_AHD0220C_Click()
    AHD0220C.Show
    AHD0220C.SetFocus
End Sub

Private Sub Mnu_AHD0230C_Click()
    AHD0230C.Show
    AHD0230C.SetFocus
End Sub

Private Sub Mnu_AHD0240C_Click()
    AHD0240C.Show
    AHD0240C.SetFocus
End Sub

'Private Sub Mnu_AHD0180C_Click()
'    AHD0180C.Show
'    AHD0180C.SetFocus
'End Sub

'Private Sub Mnu_AHD0210C_Click()
'    AHD0210C.Show
'    AHD0210C.SetFocus
'End Sub

Private Sub Mnu_AHD0250S_Click()
    AHD0250S.Show
    AHD0250S.SetFocus
End Sub

'Private Sub Mnu_AHE0010C_Click()
'    AHE0010C.Show
'    AHE0010C.SetFocus
'End Sub
'
'Private Sub Mnu_AHE0020C_Click()
'    AHE0020C.Show
'    AHE0020C.SetFocus
'End Sub
'
'Private Sub Mnu_AHE0030C_Click()
'    AHE0030C.Show
'    AHE0030C.SetFocus
'End Sub
'
'Private Sub Mnu_AHE0040C_Click()
'    AHE0040C.Show
'    AHE0040C.SetFocus
'End Sub
'
'Private Sub Mnu_AHE0050C_Click()
'    AHE0050C.Show
'    AHE0050C.SetFocus
'End Sub

'Private Sub Mnu_AHE0060C_Click()
'    AHE0060C.Show
'    AHE0060C.SetFocus
'End Sub
'
'Private Sub Mnu_AHE0110C_Click()
'    AHE0110C.Show
'    AHE0110C.SetFocus
'End Sub
'Private Sub Mnu_AHE0130C_Click()
'    AHE0130C.Show
'    AHE0130C.SetFocus
'End Sub

'Private Sub Mnu_AHE0150C_Click()
'    AHE0150C.Show
'    AHE0150C.SetFocus
'End Sub

'Private Sub Mnu_AHE0160C_Click()
'    AHE0160C.Show
'    AHE0160C.SetFocus
'End Sub

'Private Sub Mnu_AHE0170C_Click()
'    AHE0170C.Show
'    AHE0170C.SetFocus
'End Sub

'Private Sub Mnu_AHE0180C_Click()
'    AHF0180C.Show
'    AHF0180C.SetFocus
'End Sub

'Private Sub Mnu_AHE0190C_Click()
'    AHE0190C.Show
'    AHE0190C.SetFocus
'End Sub

'Private Sub Mnu_AHE0210C_Click()
'    AHE0210C.Show
'    AHE0210C.SetFocus
'End Sub

'Private Sub Mnu_AHE0220C_Click()
'    AHE0220C.Show
'    AHE0220C.SetFocus
'End Sub

'Private Sub Mnu_AHE0230C_Click()
'    AHE0230C.Show
'    AHE0230C.SetFocus
'End Sub

'Private Sub Mnu_AHE0250C_Click()
'    AHE0250C.Show
'    AHE0250C.SetFocus
'End Sub

'Private Sub Mnu_AHE0260C_Click()
'    AHE0260C.Show
'    AHE0260C.SetFocus
'
'End Sub

'Private Sub Mnu_AHF0010C_Click()
'    AHF0010C.Show
'    AHF0010C.SetFocus
'End Sub
'
'Private Sub Mnu_AHF0011C_Click()
'    AHF0011C.Show
'    AHF0011C.SetFocus
'End Sub

'Private Sub Mnu_AHF0020C_Click()
'    AHF0020C.Show
'    AHF0020C.SetFocus
'End Sub
'Private Sub Mnu_AHF0030C_Click()
'    AHF0030C.Show
'    AHF0030C.SetFocus
'End Sub
'
'Private Sub Mnu_AHF0040C_Click()
'    AHF0040C.Show
'    AHF0040C.SetFocus
'End Sub
'
'Private Sub Mnu_AHF0050C_Click()
'    AHF0050C.Show
'    AHF0050C.SetFocus
'End Sub

'Private Sub Mnu_AHF0110C_Click()
'    AHF0110C.Show
'    AHF0110C.SetFocus
'End Sub

'Private Sub Mnu_AHF0120C_Click()
'    AHF0120C.Show
'    AHF0120C.SetFocus
'End Sub

'Private Sub Mnu_AHF0140C_Click()
'    AHF0140C.Show
'    AHF0140C.SetFocus
'End Sub

'Private Sub Mnu_AHE0240C_Click()
'    AHE0240C.Show
'    AHE0240C.SetFocus
'End Sub

Private Sub Mnu_AHF0050C_Click()

End Sub

Private Sub Mnu_AHD0400C_Click()
    AHD0400C.Show
    AHD0400C.SetFocus
End Sub

Private Sub Mnu_AHD0260C_Click()
    AHD0260C.Show
    AHD0260C.SetFocus
End Sub

Private Sub Mnu_AHD0420C_Click()
    AHD0420C.Show
    AHD0420C.SetFocus

End Sub

Private Sub Mnu_AHD0430C_Click()
    AHD0430C.Show
    AHD0430C.SetFocus

End Sub

Private Sub Mnu_AHD0440C_Click()
    AHD0440C.Show
    AHD0440C.SetFocus
End Sub

Private Sub Mnu_AHD0450C_Click()
    AHD0450C.Show
    AHD0450C.SetFocus
End Sub

Private Sub Mnu_AHD0451C_Click()
    AHD0451C.Show
    AHD0451C.SetFocus
End Sub

Private Sub Mnu_AHD0510C_Click()
    AHD0510C.Show
    AHD0510C.SetFocus
End Sub

Private Sub Mnu_AHD0520C_Click()
    AHD0520C.Show
    AHD0520C.SetFocus
End Sub


Private Sub Mnu_AHD0530C_Click()
 AHD0530C.Show
 AHD0530C.SetFocus
End Sub

Private Sub Mnu_AHD0540C_Click()
 AHD0540C.Show
 AHD0540C.SetFocus
End Sub

Private Sub Mnu_AHG0010C_Click()
    AHG0010C.Show
    AHG0010C.SetFocus
End Sub

Private Sub Mnu_AHG0020C_Click()
    AHG0020C.Show
    AHG0020C.SetFocus
End Sub

Private Sub Mnu_AHG0030C_Click()
    AHG0030C.Show
    AHG0030C.SetFocus
End Sub
Private Sub Mnu_AHG0040C_Click()
    AHG0040C.Show
    AHG0040C.SetFocus
End Sub
Private Sub Mnu_AHQ0010C_Click()
    AHQ0010C.Show
    AHQ0010C.SetFocus
End Sub
'Private Sub Mnu_AHF0150C_Click()
'    AHF0150C.Show
'    AHF0150C.SetFocus
'End Sub

Private Sub Mnu_AHQ0040C_Click()
    AHQ0040C.Show
    AHQ0040C.SetFocus
End Sub


Private Sub Mnu_AHQ0050C_Click()
    AHQ0050C.Show
    AHQ0050C.SetFocus
End Sub

Private Sub Mnu_Cascade_Click()
    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    MDIMain.Arrange 0
End Sub

Private Sub Mnu_Acopy_Click()
    'Screen All Copy
    Call ActiveForm.Form_Cpy
    Call MDIMain.FormMenuSetting(Me, "", "Acopy", "")
End Sub

Private Sub Mnu_Apaste_Click()
    'Screen All Paste
    Call ActiveForm.Form_Pst
End Sub

Private Sub Mnu_Clear_Click()
    'Screen Clera
    Call ActiveForm.Form_Cls
End Sub

Private Sub Mnu_Delete_Click()
    'Delete
    Call ActiveForm.Form_Del
End Sub

Private Sub Mnu_Excel_Click()
    'Excel
    Call ActiveForm.Form_Exc
End Sub

Private Sub Mnu_Exit_Click()
    'Exit
    Call ActiveForm.Form_Exit
End Sub

Private Sub Mnu_FrozenCancel_Click()
    'Spread Col Frozens Cancel
    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    Call ActiveForm.Spread_Forzens_Cancel
End Sub

Private Sub Mnu_FrozenSetting_Click()
    'Spread Col Frozens Setting
    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    Call ActiveForm.Spread_Forzens_Setting
End Sub

Private Sub Mnu_Help_Click()
    Dim FormLD As Boolean
    For Each CurrentForm In Forms
        If CurrentForm.Name <> Me.Name Then
            FormLD = True
            Exit For
        End If
    Next CurrentForm
    If FormLD Then
        HelpDiaplay.Tag = ActiveForm.Name
    End If
    HelpDiaplay.Show (0)
    HelpDiaplay.SetFocus

End Sub

Private Sub Mnu_Horiz_Click()
    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    MDIMain.Arrange 1
End Sub

Private Sub Mnu_Mcopy_Click()
    'Screen Control Copy
    Call ActiveForm.Master_Cpy
    Call MDIMain.FormMenuSetting(Me, "", "Mcopy", "")
End Sub

Private Sub Mnu_Mpaste_Click()
    'Screen Control Paste
    Call ActiveForm.Master_Pst
End Sub

Private Sub Mnu_Print_Click()
    'Print
End Sub

Private Sub Mnu_Refer_Click()
    'Refer
    Call ActiveForm.Form_Ref
End Sub

Private Sub Mnu_RowCan_Click()
    'Spread Row Cancel
    Call ActiveForm.Spread_Can
End Sub

Private Sub Mnu_RowDel_Click()
    'Spread Row Delete
    Call ActiveForm.Spread_Del
End Sub

Private Sub Mnu_RowIns_Click()
    'Spread Row Insert
    Call ActiveForm.Form_Ins
End Sub

Private Sub Mnu_Save_Click()
    'Save
    Call ActiveForm.Form_Pro
End Sub

Private Sub Mnu_Scopy_Click()
    'Spread Row Copy
    Call ActiveForm.Spread_Cpy
    Call MDIMain.FormMenuSetting(Me, "", "Scopy", "")
End Sub

Private Sub Mnu_Sorting_Click()
    'Spread Col Sortting
    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    Call ActiveForm.Spread_ColumnsSort
End Sub

Private Sub Mnu_Spaste_Click()
    'Spread Row Paste
    Call ActiveForm.Spread_Pst
End Sub

Private Sub Mnu_Vertical_Click()
    MDIMain.StatusBar1.Panels(1) = "提示信息: "
    MDIMain.Arrange 2
End Sub
