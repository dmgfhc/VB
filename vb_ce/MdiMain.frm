VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "中板轧钢工序管理"
   ClientHeight    =   6930
   ClientLeft      =   840
   ClientTop       =   3405
   ClientWidth     =   12390
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Tag             =   "CE"
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet 
      Left            =   15
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12330
      TabIndex        =   0
      Top             =   0
      Width           =   12390
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
         _Version        =   "6.7.9782"
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
      Top             =   6465
      Width           =   12390
      _ExtentX        =   21855
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
            TextSave        =   "2016-01-05"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "17:07"
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
   Begin VB.Menu PopUp_Slab 
      Caption         =   "PopUp-Slab"
      Visible         =   0   'False
      Begin VB.Menu mnu_Slab 
         Caption         =   "Slab"
      End
   End
   Begin VB.Menu Mnu_AEA 
      Caption         =   "标准管理"
      Begin VB.Menu Mnu_AEA1 
         Caption         =   "输入板坯设计标准"
         Begin VB.Menu Mnu_AEA1020C 
            Caption         =   "录入板坯宽度决定标准 "
         End
         Begin VB.Menu Mnu_AEA1040C 
            Caption         =   "录入炼钢编制标准"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_AEA1050C 
            Caption         =   "录入板坯设计标准"
         End
         Begin VB.Menu Mnu_AEA1060C 
            Caption         =   "录入板坯长度设计标准"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_AEA1070C 
            Caption         =   "录入母板长度余量标准"
         End
         Begin VB.Menu Mnu_CEA1160C 
            Caption         =   "录入头尾放尺长度标准"
         End
         Begin VB.Menu Mnu_CEA1170C 
            Caption         =   "录入板坯堆冷信息标准"
         End
         Begin VB.Menu Mnu_CEA1140C 
            Caption         =   "录入产品厚度别轧制长度标准"
         End
         Begin VB.Menu Mnu_CEA1150C 
            Caption         =   "录入产品宽度余量标准"
         End
         Begin VB.Menu Mnu_AEA1000C 
            Caption         =   "录入头尾坯切割标准"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_CEA1130C 
            Caption         =   "录入加热炉负荷量标准(T/H) "
         End
         Begin VB.Menu Mnu_AEA4020C 
            Caption         =   "录入加热炉装炉长度标准"
         End
         Begin VB.Menu Mnu_AEA4010C 
            Caption         =   "录入轧钢工序计划标准(T/H)"
         End
      End
      Begin VB.Menu Mnu_AEA1080C 
         Caption         =   "录入炉次编制标准"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AEA3 
         Caption         =   "输入连浇炉数标准"
         Visible         =   0   'False
         Begin VB.Menu Mnu_AEA1090C 
            Caption         =   "录入连浇炉数编制标准"
         End
         Begin VB.Menu Mnu_AEA1100C 
            Caption         =   "录入炉次间钢种混浇标准"
         End
         Begin VB.Menu Mnu_AEA1110C 
            Caption         =   "录入作业准备时间"
         End
         Begin VB.Menu Mnu_AEA1120C 
            Caption         =   "录入炼钢周期"
         End
         Begin VB.Menu Mnu_AEA1121C 
            Caption         =   "钢包移动时间"
         End
         Begin VB.Menu Mnu_AEA1130C 
            Caption         =   "录入浇铸作业时间"
         End
      End
      Begin VB.Menu Mnu_AEA4 
         Caption         =   "输入轧辊编制标准"
         Begin VB.Menu Mnu_AEA2011C 
            Caption         =   "录入轧辊单位编制厚/宽度组标准"
         End
         Begin VB.Menu Mnu_AEA2012C 
            Caption         =   "录入轧辊单位编制量标准"
         End
         Begin VB.Menu Mnu_AEA2013C 
            Caption         =   "录入轧辊单位编制量厚/宽度标准"
         End
         Begin VB.Menu Mnu_AEA2014C 
            Caption         =   "录入轧辊单位调整编制对象标准"
         End
         Begin VB.Menu Mnu_AEA2015C 
            Caption         =   "录入轧辊单位调整编制对象除外标准"
         End
         Begin VB.Menu Mnu_AEA2016C 
            Caption         =   "录入辊期标准公里数"
         End
         Begin VB.Menu Mnu_AEA2017C 
            Caption         =   "录入轧辊钢种系数"
         End
      End
   End
   Begin VB.Menu Mnu_AEB 
      Caption         =   "坯料设计"
      Begin VB.Menu Mnu_AEB0010C 
         Caption         =   "连铸断面查询/修改"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AEB0020C 
         Caption         =   "非运转设备查询/修改"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AEB1010C 
         Caption         =   "设计订单查询/选定"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CEC0000C 
         Caption         =   "标准板坯设计"
      End
      Begin VB.Menu Mnu_AEB1060C 
         Caption         =   "订单分析结果查询"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AEB2060C 
         Caption         =   "板坯设计结果修改"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AEB1070C 
         Caption         =   "HMI 板坯设计"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AEB3050C 
         Caption         =   "炉次编制结果修改"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CEF1510C 
         Caption         =   "板坯替代计划"
      End
   End
   Begin VB.Menu Mnu_CEG0 
      Caption         =   "月单位坯料使用计划"
      Begin VB.Menu Mnu_CEG1010C 
         Caption         =   "月坯料使用计划对象选定/查询"
      End
      Begin VB.Menu Mnu_Line1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEG1040C 
         Caption         =   "加热炉均衡查询/调整"
      End
      Begin VB.Menu Mnu_CEG1050C 
         Caption         =   "炼钢厂均衡查询/调整"
      End
      Begin VB.Menu Mnu_Line14 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEG1060C 
         Caption         =   "月产品生产计划查询/调整/确定"
      End
   End
   Begin VB.Menu Mnu_CEG 
      Caption         =   "周单位坯料使用计划"
      Visible         =   0   'False
      Begin VB.Menu Mnu_CEG2010C 
         Caption         =   "周坯料使用计划对象选定/查询"
      End
      Begin VB.Menu Mnu_CEG2060C 
         Caption         =   "HMI板坯设计"
      End
      Begin VB.Menu Mnu_Line2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CEG2040C 
         Caption         =   "加热炉均衡查询/调整"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CEG2050C 
         Caption         =   "炼钢厂均衡查询/调整"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Line3 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEG2100C 
         Caption         =   "单/多订单板坯设计"
      End
      Begin VB.Menu Mnu_CEG2130C 
         Caption         =   "强制订单板坯设计"
      End
      Begin VB.Menu Mnu_CEG2150C 
         Caption         =   "周坯料使用计划结果信息查询/确定"
      End
      Begin VB.Menu Mnu_Line31 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEG2140C 
         Caption         =   "申请紧急坯"
      End
      Begin VB.Menu Mnu_Line32 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEG3020C 
         Caption         =   "坯料使用计划的分析信息查询"
      End
      Begin VB.Menu Mnu_CEG3010C 
         Caption         =   "使用坯料申请信息查询/取消"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_CEC 
      Caption         =   "周单位坯料使用计划"
      Begin VB.Menu Mnu_CEC1010C 
         Caption         =   "设计订单查询/选定"
      End
      Begin VB.Menu Mnu_CEC2060C 
         Caption         =   "板坯设计结果修改"
      End
      Begin VB.Menu Mnu_CEC1070C 
         Caption         =   "HMI板坯设计"
      End
      Begin VB.Menu Mnu_Line81 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEC2050C 
         Caption         =   "炼钢厂均衡查询/调整"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CEC4010C 
         Caption         =   "长坯料设计"
      End
      Begin VB.Menu Mnu_CEC4030C 
         Caption         =   "强制长坯料设计"
      End
      Begin VB.Menu Mnu_CEC2150C 
         Caption         =   "坯料使用计划结果信息查询/确定"
      End
      Begin VB.Menu Mnu_Line8 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEC2140C 
         Caption         =   "申请紧急坯"
      End
      Begin VB.Menu Mnu_Line4 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEC3020C 
         Caption         =   "坯料使用计划的分析信息查询"
      End
   End
   Begin VB.Menu Mnu_CEH 
      Caption         =   "坯料分段作业计划"
      Begin VB.Menu Mnu_CEH1010C 
         Caption         =   "坯料分段作业指示对象选定/查询"
      End
      Begin VB.Menu Mnu_CEH2010C 
         Caption         =   "坯料紧急分段作业指示"
      End
      Begin VB.Menu Mnu_Line5 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEH1011C 
         Caption         =   "坯料分段作业指示对象选定/查询(临时)"
      End
      Begin VB.Menu Mnu_CEH2011C 
         Caption         =   "坯料紧急分段作业指示(临时)"
      End
      Begin VB.Menu Mnu_CGA2088C 
         Caption         =   "中板厂外板坯切割作业界面(临时)"
      End
   End
   Begin VB.Menu Mnu_CEI 
      Caption         =   "轧钢工序计划"
      Begin VB.Menu Mnu_CED1010C 
         Caption         =   "轧钢工序计划对象坯料选定/查询"
      End
      Begin VB.Menu Mnu_Line6 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_AEC2900C 
         Caption         =   "轧辊单位编制结果修改"
      End
      Begin VB.Menu Mnu_AEC2910C 
         Caption         =   "轧辊单位编制结果查询"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_AEE 
      Caption         =   "作业指示"
      Begin VB.Menu Mnu_CED4010C 
         Caption         =   "确定轧钢作业生产管制指示"
      End
      Begin VB.Menu Mnu_Line15 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEE1040C 
         Caption         =   "坯料分段切割指示查询"
      End
      Begin VB.Menu Mnu_CEE2010C 
         Caption         =   "轧钢作业指示查询"
      End
      Begin VB.Menu Mnu_CEE3010C 
         Caption         =   "精整作业指示查询"
      End
   End
   Begin VB.Menu Mnu_CEF 
      Caption         =   "替代管理"
      Begin VB.Menu Mnu_CEF1010C 
         Caption         =   "可替代订单选定"
      End
      Begin VB.Menu Mnu_CEF1030C 
         Caption         =   "可替代余坯选定"
      End
      Begin VB.Menu Mnu_Line7 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEF1065C 
         Caption         =   "余坯替代"
      End
      Begin VB.Menu Mnu_CEF1150C 
         Caption         =   "HMI余坯替代"
      End
      Begin VB.Menu Mnu_CEF1200C 
         Caption         =   "替代结果查询及修改"
      End
      Begin VB.Menu Mnu_Line10 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_CEF1310C 
         Caption         =   "替代产品长度变更"
      End
      Begin VB.Menu Mnu_CEF1410C 
         Caption         =   "订单材申请坯料替代处理"
      End
      Begin VB.Menu Mnu_Line22 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_ACE1209C 
         Caption         =   "替代履历查询"
      End
      Begin VB.Menu Mnu_ACB1020C 
         Caption         =   "物料库存现状查询"
      End
      Begin VB.Menu Mnu_ACE5010C 
         Caption         =   "指定及解除委托加工"
      End
   End
   Begin VB.Menu Mnu_AEZ 
      Caption         =   "其它管理"
      Begin VB.Menu Mnu_AEZ2010C 
         Caption         =   "查询相关工作管理信息"
      End
      Begin VB.Menu Mnu_Line11 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACB4070C 
         Caption         =   "板坯待判/判定实绩录入"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Line12 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_AEC3000C 
         Caption         =   "中板坯料申请信息查询/选定工序计划炼钢"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AEC3010C 
         Caption         =   "中板坯料申请信息查询/确定生产"
      End
      Begin VB.Menu Mnu_Line13 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_ACB5025C 
         Caption         =   "半产品装车实绩录入"
      End
      Begin VB.Menu Mnu_ACB5030C 
         Caption         =   "半产品/产品卸车实绩录入"
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
      Begin VB.Menu Line4 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Help 
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
'操作人员
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

'Private Sub MDIForm_Load()
'
'    Dim Active_YN As String
'
'    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
'
'    Me.BackColor = &HE0E0E0
'
'    If GF_DbConnect = False Then
'        Unload Me
'    Else
'
'        Active_YN = GetSetting("NISCO", "EXE-FILE", "CE.exe")
'
''        If Active_YN = "1" Then
''            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
''            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
''            MDIMain.StatusBar1.Panels(1) = "提示信息 : "
''            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
''        Else
''            Call Gp_MsgBoxDisplay("只能从主画面登陆...", "W")
''            Unload Me
''            Exit Sub
''        End If
'
'        sUserID = "0860011"
'        sUserName = "金成浩"
'        MDIMain.StatusBar1.Panels(1) = "提示信息 ："
'        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'
'        If Mid(M_CN1, Len(M_CN1), 1) = "9" Then
'            MDIMain.StatusBar1.Panels(8) = "正式机"
'        Else
'            MDIMain.StatusBar1.Panels(8) = "测试机"
'        End If
'
'    End If
'
'End Sub
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
        Active_YN = GetSetting("NISCO", "EXE-FILE", "CE.exe")
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

        

'        sUserID = "1JS1005"
'        sUserName = "杨猛"
'        MDIMain.StatusBar1.Panels(1) = "提示信息 ："
'        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName

'
'        If Mid(M_CN1, Len(M_CN1), 1) = "9" Then
'            MDIMain.StatusBar1.Panels(8) = "正式机"
'        Else
'            MDIMain.StatusBar1.Panels(8) = "测试机"
'        End If

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
    
        'If Gf_MessConfirm("Low rank program was not ended," + vbCrLf + "end Program ?", "Q", Me.Caption) Then
        If MsgBox("还有未关闭的操作界面," + vbCrLf + "是否退出当前系统 ?", MB_YESNO _
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
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "CE.exe", ""

End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
    
    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
    
    If Screen.ActiveForm.Name = "MDIMain" Then
        
        If Button.Key = "Exit" Then
            If vbYes = MsgBox(Me.Caption + " 系统是否退出 ?", vbQuestion + vbYesNo, Me.Caption) Then
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

    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
    
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

Private Sub Mnu_ACB1020C_Click()
    ACB1020C.Show
    ACB1020C.SetFocus
End Sub

Private Sub Mnu_ACB4070C_Click()
    ACB4070C.Show
    ACB4070C.SetFocus
End Sub

Private Sub Mnu_ACB5025C_Click()
    ACB5025C.Show
    ACB5025C.SetFocus
End Sub

Private Sub Mnu_ACB5030C_Click()
    ACB5035C.Show
    ACB5035C.SetFocus
End Sub

Private Sub Mnu_ACE1209C_Click()
    ACE1209C.Show
    ACE1209C.SetFocus
End Sub

Private Sub Mnu_ACE5010C_Click()
    ACE5010C.Show
    ACE5010C.SetFocus
End Sub

Private Sub Mnu_AEA1020C_Click()
    AEA1020C.Show
    AEA1020C.SetFocus
End Sub

Private Sub Mnu_AEA1040C_Click()
    AEA1040C.Show
    AEA1040C.SetFocus
End Sub

Private Sub Mnu_AEA1050C_Click()
    AEA1050C.Show
    AEA1050C.SetFocus
End Sub

Private Sub Mnu_AEA1060C_Click()
    AEA1060C.Show
    AEA1060C.SetFocus
End Sub

Private Sub Mnu_AEA1070C_Click()
    AEA1070C.Show
    AEA1070C.SetFocus
End Sub

Private Sub Mnu_AEA1080C_Click()
    AEA1080C.Show
    AEA1080C.SetFocus
End Sub

Private Sub Mnu_AEA1090C_Click()
    AEA1090C.Show
    AEA1090C.SetFocus
End Sub

Private Sub Mnu_AEA1100C_Click()
    AEA1100C.Show
    AEA1100C.SetFocus
End Sub

Private Sub Mnu_AEA1110C_Click()
    AEA1110C.Show
    AEA1110C.SetFocus
End Sub

Private Sub Mnu_AEA1120C_Click()
    AEA1120C.Show
    AEA1120C.SetFocus
End Sub

Private Sub Mnu_AEA1121C_Click()
    AEA1121C.Show
    AEA1121C.SetFocus
End Sub

Private Sub Mnu_AEA1130C_Click()
    AEA1130C.Show
    AEA1130C.SetFocus
End Sub


Private Sub Mnu_AEA2011C_Click()
    AEA2011C.Show
    AEA2011C.SetFocus
End Sub

Private Sub Mnu_AEA2012C_Click()
    AEA2012C.Show
    AEA2012C.SetFocus
End Sub

Private Sub Mnu_AEA2013C_Click()
    AEA2013C.Show
    AEA2013C.SetFocus
End Sub

Private Sub Mnu_AEA2014C_Click()
    AEA2014C.Show
    AEA2014C.SetFocus
End Sub

Private Sub Mnu_AEA2015C_Click()
    AEA2015C.Show
    AEA2015C.SetFocus
End Sub

Private Sub Mnu_AEA2016C_Click()
    AEA2016C.Show
    AEA2016C.SetFocus
End Sub

Private Sub Mnu_AEA2017C_Click()
    AEA2017C.Show
    AEA2017C.SetFocus
End Sub

Private Sub Mnu_AEA4010C_Click()
    AEA4010C.Show
    AEA4010C.SetFocus
End Sub

Private Sub Mnu_AEA4020C_Click()
    AEA4020C.Show
    AEA4020C.SetFocus
End Sub

Private Sub Mnu_AEB0010C_Click()
    AEB0010C.Show
    AEB0010C.SetFocus
End Sub

Private Sub Mnu_AEB0020C_Click()
    AEB0020C.Show
    AEB0020C.SetFocus
End Sub

Private Sub Mnu_AEB1010C_Click()
    CEB1010C.Show
    CEB1010C.SetFocus
End Sub

Private Sub Mnu_AEB1060C_Click()
    'AEB1060C.Show
    'AEB1060C.SetFocus
End Sub

Private Sub Mnu_AEB1070C_Click()
    AEB1070C.Show
    AEB1070C.SetFocus
End Sub

Private Sub Mnu_AEB2060C_Click()
'    CEB2060C.Show
'    CEB2060C.SetFocus
End Sub

Private Sub Mnu_AEB3050C_Click()
    AEB3050C.Show
    AEB3050C.SetFocus
End Sub

Private Sub Mnu_AEC2900C_Click()
    CEC2900C.Show
    CEC2900C.SetFocus
End Sub

Private Sub Mnu_AEC2910C_Click()
    AEC2910C.Show
    AEC2910C.SetFocus
End Sub

Private Sub Mnu_AEC3000C_Click()
'    AEC3000C.Show
'    AEC3000C.SetFocus
End Sub

Private Sub Mnu_AEC3010C_Click()
    AEC3010C.Show
    AEC3010C.SetFocus
End Sub

Private Sub Mnu_AEZ2010C_Click()
    AEZ2010C.Show
    AEZ2010C.SetFocus
End Sub

Private Sub Mnu_Cascade_Click()
    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
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

'Private Sub Mnu_CEA1110C_Click()
'    CEA1110C.Show
'    CEA1110C.SetFocus
'End Sub

Private Sub Mnu_CEA1130C_Click()
    CEA1130C.Show
    CEA1130C.SetFocus
End Sub

Private Sub Mnu_CEA1140C_Click()
    CEA1140C.Show
    CEA1140C.SetFocus
End Sub

Private Sub Mnu_CEA1150C_Click()
    CEA1150C.Show
    CEA1150C.SetFocus
End Sub

Private Sub Mnu_CEA1160C_Click()
    CEA1160C.Show
    CEA1160C.SetFocus
End Sub

Private Sub Mnu_CEA1170C_Click()
    CEA1170C.Show
    CEA1170C.SetFocus

End Sub

Private Sub Mnu_CEC0000C_Click()
    CEC0000C.Show
    CEC0000C.SetFocus
End Sub

Private Sub Mnu_CEC1010C_Click()
    CEC1010C.Show
    CEC1010C.SetFocus
End Sub

Private Sub Mnu_CEC1070C_Click()
    CEC1070C.Show
    CEC1070C.SetFocus
End Sub

Private Sub Mnu_CEC2050C_Click()
    CEC2050C.Show
    CEC2050C.SetFocus
End Sub

Private Sub Mnu_CEC2060C_Click()
    CEC2060C.Show
    CEC2060C.SetFocus
End Sub

Private Sub Mnu_CEC4010C_Click()
    CEC4010C.Show
    CEC4010C.SetFocus
End Sub

Private Sub Mnu_CEC4030C_Click()
    CEC4030C.Show
    CEC4030C.SetFocus
End Sub

Private Sub Mnu_CED1010C_Click()
    CED1010C.Show
    CED1010C.SetFocus
End Sub

Private Sub Mnu_CED4010C_Click()
    CED4010C.Show
    CED4010C.SetFocus
End Sub

Private Sub Mnu_CEE1040C_Click()
    CEE1040C.Show
    CEE1040C.SetFocus
End Sub

Private Sub Mnu_CEE2010C_Click()
    CEE2010C.Show
    CEE2010C.SetFocus
End Sub

Private Sub Mnu_CEE3010C_Click()
    CEE3010C.Show
    CEE3010C.SetFocus
End Sub

Private Sub Mnu_CEF1010C_Click()
    CEF1010C.Show
    CEF1010C.SetFocus
End Sub

Private Sub Mnu_CEF1030C_Click()
    CEF1030C.Show
    CEF1030C.SetFocus
End Sub

Private Sub Mnu_CEF1065C_Click()
    CEF1065C.Show
    CEF1065C.SetFocus
End Sub

Private Sub Mnu_CEF1150C_Click()
    CEF1150C.Show
    CEF1150C.SetFocus
End Sub

Private Sub Mnu_CEF1200C_Click()
    CEF1200C.Show
    CEF1200C.SetFocus
End Sub

Private Sub Mnu_CEF1310C_Click()
    CEF1310C.Show
    CEF1310C.SetFocus
End Sub

Private Sub Mnu_CEF1410C_Click()
    CEF1410C.Show
    CEF1410C.SetFocus
End Sub

Private Sub Mnu_CEF1510C_Click()
    CEF1510C.Show
    CEF1510C.SetFocus
End Sub

Private Sub Mnu_CEG1010C_Click()
    CEG1010C.Show
    CEG1010C.SetFocus
End Sub

Private Sub Mnu_CEG1040C_Click()
    CEG1040C.Show
    CEG1040C.SetFocus
End Sub

Private Sub Mnu_CEG1050C_Click()
    CEG1050C.Show
    CEG1050C.SetFocus
End Sub

Private Sub Mnu_CEG1060C_Click()
    CEG1060C.Show
    CEG1060C.SetFocus
End Sub

Private Sub Mnu_CEG2010C_Click()
    CEG2010C.Show
    CEG2010C.SetFocus
End Sub

Private Sub Mnu_CEG2040C_Click()
    CEG2040C.Show
    CEG2040C.SetFocus
End Sub

Private Sub Mnu_CEG2050C_Click()
    CEG2050C.Show
    CEG2050C.SetFocus
End Sub

Private Sub Mnu_CEG2060C_Click()
    CEG2060C.Show
    CEG2060C.SetFocus
End Sub

Private Sub Mnu_CEG2100C_Click()
    CEG2100C.Show
    CEG2100C.SetFocus
End Sub

Private Sub Mnu_CEG2130C_Click()
    CEG2130C.Show
    CEG2130C.SetFocus
End Sub

Private Sub Mnu_CEC2140C_Click()
    CEC2140C.Show
    CEC2140C.SetFocus
End Sub

Private Sub Mnu_CEC2150C_Click()
    CEC2150C.Show
    CEC2150C.SetFocus
End Sub

Private Sub Mnu_CEG2140C_Click()
    CEG2140C.Show
    CEG2140C.SetFocus
End Sub

Private Sub Mnu_CEG2150C_Click()
    CEG2150C.Show
    CEG2150C.SetFocus
End Sub

Private Sub Mnu_CEG3010C_Click()
    CEG3010C.Show
    CEG3010C.SetFocus
End Sub

Private Sub Mnu_CEC3020C_Click()
    CEG3020C.Show
    CEG3020C.SetFocus
End Sub

Private Sub Mnu_CEG3020C_Click()
    CEG3020C.Show
    CEG3020C.SetFocus
End Sub

Private Sub Mnu_CEH1010C_Click()
    CEH1010C.Show
    CEH1010C.SetFocus
End Sub

Private Sub Mnu_CEH1011C_Click()
    CEH1011C.Show
    CEH1011C.SetFocus
End Sub

Private Sub Mnu_CEH2010C_Click()
    CEH2010C.Show
    CEH2010C.SetFocus
End Sub

Private Sub Mnu_CEH2011C_Click()
    CEH2011C.Show
    CEH2011C.SetFocus
End Sub

Private Sub Mnu_CGA2088C_Click()
    CGA2088C.Show
    CGA2088C.SetFocus
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
    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
    Call ActiveForm.Spread_Forzens_Cancel
End Sub

Private Sub Mnu_FrozenSetting_Click()
    'Spread Col Frozens Setting
    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
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
    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
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
    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
    Call ActiveForm.Spread_ColumnsSort
End Sub

Private Sub Mnu_Spaste_Click()
    'Spread Row Paste
    Call ActiveForm.Spread_Pst
End Sub

Private Sub Mnu_Vertical_Click()
    MDIMain.StatusBar1.Panels(1) = "提示信息 : "
    MDIMain.Arrange 2
End Sub

