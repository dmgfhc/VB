VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "板卷轧钢作业管理"
   ClientHeight    =   8010
   ClientLeft      =   345
   ClientTop       =   2655
   ClientWidth     =   11400
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Tag             =   "BG"
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11340
      TabIndex        =   0
      Top             =   0
      Width           =   11400
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
         _Version        =   "6.0.8169"
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
      Top             =   7545
      Width           =   11400
      _ExtentX        =   20108
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
            TextSave        =   "2017/10/27"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "10:02"
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
   Begin VB.Menu mnu_aga 
      Caption         =   "轧钢作业实绩管理"
      Begin VB.Menu mnu_aga2010c 
         Caption         =   "加热炉作业实绩查询及修改界面_AGA2010C"
      End
      Begin VB.Menu mnu_agb2010c 
         Caption         =   "轧制作业实绩查询及修改界面_AGB2010C"
      End
      Begin VB.Menu mnu_agb2021c 
         Caption         =   "钢卷卷取实绩查询及修改界面_AGB2021C"
      End
      Begin VB.Menu mnu_agc2010c 
         Caption         =   "热矫直实绩查询及修改界面_AGC2010C"
      End
      Begin VB.Menu mnu_agb2060c 
         Caption         =   "在线钢卷入库界面_AGB2060C"
      End
   End
   Begin VB.Menu mnu_agc 
      Caption         =   "精整作业实绩管理"
      Begin VB.Menu mnu_agc2035c 
         Caption         =   "钢板标印信息发送界面_AGC2035C"
      End
      Begin VB.Menu mnu_agc2031c 
         Caption         =   "钢板剪切实绩查询及修改界面_AGC2031C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_agc2400c 
         Caption         =   "钢板取样查询及修改界面_AGC2400C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_agc2410c 
         Caption         =   "钢板取样实绩查询及修改界面_AGC2400C"
      End
      Begin VB.Menu mnu_agc2430c 
         Caption         =   "理化检验委托单_AGC2430C"
      End
      Begin VB.Menu mnu_agc2432c 
         Caption         =   "理化检验委托单-PWHT_AGC2432C"
      End
      Begin VB.Menu mnu_agc2440c 
         Caption         =   "剪切前当班取样项目查询界面_AGC2440C"
      End
      Begin VB.Menu mnu_agc2020c 
         Caption         =   "表面检查实绩查询及修改界面_AGC2020C"
      End
      Begin VB.Menu mnu_agc2022c 
         Caption         =   "抽检实绩查询及修改界面_AGC2022C"
      End
      Begin VB.Menu Line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_agc2021c 
         Caption         =   "钢板剩磁检查实绩查询及修改界面_AGC2021C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_agc2045c 
         Caption         =   "在线探伤指示下达画面_AGC2045C"
      End
      Begin VB.Menu mnu_agc2046c 
         Caption         =   "在线探伤实绩确认画面_AGC2046C"
      End
      Begin VB.Menu mnu_agc2040c 
         Caption         =   "探伤实绩查询及修改界面_AGC2040C"
      End
      Begin VB.Menu mnu_AGC2050C 
         Caption         =   "火切实绩查询及修改界面_AGC2050C"
      End
      Begin VB.Menu mnu_AGC2051C 
         Caption         =   "分板实绩查询及修改界面_AGC2051C"
      End
      Begin VB.Menu mnu_agc2070c 
         Caption         =   "冷矫实绩查询及修改界面_AGC2070C"
      End
      Begin VB.Menu mnu_age2021c 
         Caption         =   "板卷未入库产品垛位管理界面_AGE2021C"
      End
      Begin VB.Menu mnu_age2031c 
         Caption         =   "板卷未入库产品垛位变更界面_AGE2031C"
      End
   End
   Begin VB.Menu mnu_agc2 
      Caption         =   "2#精整作业实绩管理"
      Begin VB.Menu mnu_agb3010c 
         Caption         =   "母板指示界面_AGB3010C"
      End
      Begin VB.Menu mnu_agb3011c 
         Caption         =   "钢板指示界面_AGB3011C"
      End
      Begin VB.Menu mnu_agb3012c 
         Caption         =   "切割计划下达界面_AGB3012C"
      End
      Begin VB.Menu mnu_agc2060c 
         Caption         =   "上下线实绩查询与修改界面_AGC2060C"
      End
      Begin VB.Menu Mnu_agc2011c 
         Caption         =   "冷床实绩查询及修改界面_AGC2011C"
      End
      Begin VB.Menu mnu_agb3020c 
         Caption         =   "母板分段实绩查询与修改界面_AGB3020C"
      End
      Begin VB.Menu mnu_agc2030c 
         Caption         =   "双边剪实绩查询及修改界面_AGC2030C"
      End
      Begin VB.Menu mnu_agc2036c 
         Caption         =   "钢板标印信息发送界面_AGC2036C"
      End
      Begin VB.Menu mnu_agc2037c 
         Caption         =   "钢板剪切实绩查询及修改界面_AGC2037C"
      End
   End
   Begin VB.Menu mnu_age 
      Caption         =   "成品库库管理"
      Begin VB.Menu mnu_age2020c 
         Caption         =   "在线钢板入库界面_AGE2020C"
      End
      Begin VB.Menu mnu_age2030c 
         Caption         =   "钢板垛位变更及查询界面_AGE2030C"
      End
      Begin VB.Menu mnu_age2040c 
         Caption         =   "钢板库库存现状查询_AGE2040C"
      End
      Begin VB.Menu LIEN10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_age2010c 
         Caption         =   "钢卷垛位变更及查询界面_AGE2010C"
      End
      Begin VB.Menu mnu_age2080c 
         Caption         =   "钢卷库库存现状查询_AGE2080C"
      End
   End
   Begin VB.Menu mnu_agf 
      Caption         =   "轧辊管理"
      Begin VB.Menu mnu_agf2010c 
         Caption         =   "轧辊、轴承(座)入库实绩查询及修改界面_AGF2010C"
      End
      Begin VB.Menu mnu_agf2030c 
         Caption         =   "轧辊磨削实绩查询及修改界面_AGF2030C"
      End
      Begin VB.Menu mnu_agf2050c 
         Caption         =   "轧辊装配实绩查询及修改界面_AGF2050C"
      End
      Begin VB.Menu mnu_agb2040c 
         Caption         =   "轧辊使用实绩查询及修改界面_AGB2040C"
      End
      Begin VB.Menu mnu_agf2032c 
         Caption         =   "轧辊磨削实绩查询(按时间)_AGF2032C"
      End
      Begin VB.Menu mnu_agf2051c 
         Caption         =   "轧辊装配实绩查询及发送界面_AGF2051C"
      End
      Begin VB.Menu mnu_agf2070c 
         Caption         =   "轧辊使用实绩查询(按时间)_AGF2070C"
      End
      Begin VB.Menu mnu_agf2090c 
         Caption         =   "轧辊堆焊实绩查询及修改界面_AGF2090C"
      End
      Begin VB.Menu mnu_agf2020c 
         Caption         =   "轧辊、轴承(座)报废实绩查询及修改界面_AGF2020C"
      End
      Begin VB.Menu mu 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_agf2031c 
         Caption         =   "工作辊轴承(座)保养管理界面_AGF2031C"
      End
      Begin VB.Menu mnu_agf2033c 
         Caption         =   "支撑辊轴承(座)保养管理界面_AGF2033C"
      End
      Begin VB.Menu mnu_agf2040c 
         Caption         =   "轧辊、轴承(座)库存管理界面_AGF2040C"
      End
      Begin VB.Menu mnu_agf2060c 
         Caption         =   "轧辊/轴承座/轴承号管理_AGF2060C"
         Visible         =   0   'False
      End
      Begin VB.Menu mu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_agf3010c 
         Caption         =   "卷筒入库实绩查询及修改界面_AGF3010C"
      End
      Begin VB.Menu mnu_agf3020c 
         Caption         =   "卷筒使用实绩及库存查询界面_AGF3020C"
      End
      Begin VB.Menu mnu_agf3030c 
         Caption         =   "卷筒报废实绩查询及修改界面_AGF3030C"
      End
      Begin VB.Menu mnu_agf3040c 
         Caption         =   "卷筒上下线实绩查询及修改界面_AGF3040C"
      End
      Begin VB.Menu mnu_agf3050c 
         Caption         =   "卷筒修复实绩查询及修改界面_AGF3050C"
      End
   End
   Begin VB.Menu mnu_agg 
      Caption         =   "作业指示管理"
      Begin VB.Menu mnu_agg2010c 
         Caption         =   "指示调整_AGG2040C"
      End
      Begin VB.Menu mnu_agg2060c 
         Caption         =   "轧钢计划查询界面_AGG2060C"
      End
      Begin VB.Menu mnu_agg2020c 
         Caption         =   "轧钢作业指示查询界面_AGG2020C"
      End
      Begin VB.Menu mnu_agg2030c 
         Caption         =   "精整作业指示查询界面_AGG2030C"
      End
      Begin VB.Menu mnu_agg2080c 
         Caption         =   "作业指示状态查询_AGG2080C"
      End
      Begin VB.Menu mnu_agg2050c 
         Caption         =   "订单坯库存查询_AGG2050C"
      End
      Begin VB.Menu LINE7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_CGD2070C 
         Caption         =   "录入精整作业指示_CGD2070C"
      End
      Begin VB.Menu mnu_ACB4110C 
         Caption         =   "精整作业对象查询_ACB4110C"
      End
   End
   Begin VB.Menu Mnu_Other 
      Caption         =   "轧线运行实绩"
      Begin VB.Menu mnu_agb2030c 
         Caption         =   "轧钢生产线进程现状界面_AGB2030C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_agc2090c 
         Caption         =   "轧钢生产线停机实绩查询及修改界面_AGC2090C"
      End
      Begin VB.Menu mnu_agc2100c 
         Caption         =   "公辅材料使用实绩查询及修改界面_AGC2100C"
      End
      Begin VB.Menu mnu_agf2080c 
         Caption         =   "废钢实绩查询及修改界面_AGF2080C"
      End
   End
   Begin VB.Menu Mnu_Others 
      Caption         =   "各种实绩查询"
      Begin VB.Menu Mnu_akx1010c 
         Caption         =   "加热炉实绩查询_AGA2011C"
      End
      Begin VB.Menu Mnu_agb2011c 
         Caption         =   "轧钢实绩查询_AGB2011C"
      End
      Begin VB.Menu Mnu_agb2012c 
         Caption         =   "轧钢测厚实绩查询_AGB2012C"
      End
      Begin VB.Menu Mnu_agb2013c 
         Caption         =   "母板实绩查询界面_AGB2013C"
      End
      Begin VB.Menu mnu_agc2200c 
         Caption         =   "钢板实绩查询_AGC2200C"
      End
      Begin VB.Menu mnu_agc2038c 
         Caption         =   "标印信息查询_AGC2038C"
      End
      Begin VB.Menu mnu_agc2041c 
         Caption         =   "探伤实绩查询_AGC2041C"
      End
      Begin VB.Menu mnu_agc2042c 
         Caption         =   "探伤日报表查询_AGC2042C"
      End
      Begin VB.Menu mnu_agz1010c 
         Caption         =   "综合查询_AGC2901C"
      End
      Begin VB.Menu mnu_agt1090c 
         Caption         =   "物料全息查询_AGT1090C"
      End
      Begin VB.Menu mnu_spa 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_agc2201c 
         Caption         =   "钢板几何尺寸测量数据查询_AGC2201C"
      End
      Begin VB.Menu mnu_agt1100c 
         Caption         =   "在制品查询_AGT1100C"
      End
      Begin VB.Menu mnu_agt1050c 
         Caption         =   "非计划查询_AGT1050C"
      End
      Begin VB.Menu Mnu_agc3020c 
         Caption         =   "产品退判查询_AGC3020C"
      End
      Begin VB.Menu Mnu_agc3030c 
         Caption         =   "改判统计查询_AGC3030C"
      End
   End
   Begin VB.Menu Mnu_Windows 
      Caption         =   "Windows"
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
      Begin VB.Menu Mnu_Help 
         Caption         =   "界面说明书 F1"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub line_Click()

End Sub

Private Sub MDIForm_Activate()

'    Call MDIMain.FormMenuSetting("Start", Toolbar_St)

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
'        If UBound(args) = 1 Then
'             MainFrmType = "New"
'             sUserID = args(0) ' 2012.11.09 新增  耿朝雷
'             sUserName = args(1) ' 2012.11.09 新增  耿朝雷
'             MDIMain.StatusBar1.Panels(1) = "提示信息 ：" ' 2012.11.09 新增  耿朝雷
'             MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName ' 2012.11.09 新增  耿朝雷
'        Else
'            Active_YN = GetSetting("NISCO", "EXE-FILE", "BG.exe")
'            If Active_YN = "1" Then
'                MainFrmType = "Old"
'                sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'                sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'                MDIMain.StatusBar1.Panels(1) = "提示信息 ："
'                MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'            Else
'                Call Gp_MsgBoxDisplay("只能从主画面登陆...", "W")
'                Unload Me
'                Exit Sub
'            End If
'        End If  ' 2012.11.09 新增  耿朝雷
        
        sUserID = "1JS1005"
        sUserName = "杨猛"
        MDIMain.StatusBar1.Panels(1) = "提示信息 ："
        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'

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
    
        If MsgBox("有尚未结束的程序," + vbCrLf + "结束程序么 ?", MB_YESNO _
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
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "BG.exe", ""

End Sub

Private Sub mnu_aga1010c_Click()
    AGA2010C.Show
    AGA2010C.SetFocus
End Sub


Private Sub mnu_ACB4110C_Click()
    ACB4110C.Show
    ACB4110C.SetFocus
End Sub

Private Sub Mnu_agb2011c_Click()
    AGB2011C.Show
    AGB2011C.SetFocus
End Sub

Private Sub mnu_agc2010c_old_Click()

End Sub

'Private Sub mnu_AGC2700C_Click()
'    AGC2700C.Show
'    AGC2700C.SetFocus
'End Sub

Private Sub mnu_aga2010c_Click()
    AGA2010C.Show
    AGA2010C.SetFocus
End Sub

Private Sub mnu_agb2010c_Click()
    AGB2015C.Show
    AGB2015C.SetFocus
End Sub

Private Sub Mnu_agb2012c_Click()
    AGB2012C.Show
    AGB2012C.SetFocus
End Sub

Private Sub Mnu_agb2013c_Click()
   AGB2013C.Show
   AGB2013C.SetFocus
End Sub

Private Sub mnu_agb2021c_Click()
    AGB2021C.Show
    AGB2021C.SetFocus
End Sub

'Private Sub mnu_agb2030c_Click()
'    AGB2030C.Show
'    AGB2030C.SetFocus
'End Sub

Private Sub mnu_agb2040c_Click()
    AGB2040C.Show
    AGB2040C.SetFocus
End Sub

Private Sub mnu_agb2060c_Click()
    AGB2060C.Show
    AGB2060C.SetFocus
End Sub

Private Sub mnu_agb3010c_Click()
    AGB3010C.Show
    AGB3010C.SetFocus
End Sub

Private Sub mnu_agb3011c_Click()
    AGB3011C.Show
    AGB3011C.SetFocus
End Sub

Private Sub mnu_agb3012c_Click()
    AGB3012C.Show
    AGB3012C.SetFocus
End Sub

Private Sub mnu_agb3020c_Click()
    AGB3020C.Show
    AGB3020C.SetFocus
End Sub

Private Sub mnu_agc2010c_Click()
    AGC2010C.Show
    AGC2010C.SetFocus
End Sub

Private Sub mnu_agc2011c_Click()
    AGC2011C.Show
    AGC2011C.SetFocus
End Sub

Private Sub mnu_agc2020c_Click()
    AGC2020C.Show
    AGC2020C.SetFocus
End Sub

Private Sub mnu_agc2022c_Click()
    AGC2022C.Show
    AGC2022C.SetFocus
End Sub

'Private Sub mnu_agc2021c_Click()
'    AGC2021C.Show
'    AGC2021C.SetFocus
'End Sub

Private Sub mnu_agc2030c_Click()
    AGC2030C.Show
    AGC2030C.SetFocus
End Sub
Private Sub mnu_agc2031c_Click()
    AGC2031C.Show
    AGC2031C.SetFocus
End Sub

Private Sub mnu_agc2035c_Click()
    AGC2035C.Show
    AGC2035C.SetFocus
End Sub

Private Sub mnu_agc2036c_Click()
    AGC2036C.Show
    AGC2036C.SetFocus
End Sub

Private Sub mnu_agc2037c_Click()
    AGC2037C.Show
    AGC2037C.SetFocus
End Sub

Private Sub mnu_agc2038c_Click()
    AGC2038C.Show
    AGC2038C.SetFocus
End Sub

Private Sub mnu_AGC2040C_Click()
    AGC2040C.Show
    AGC2040C.SetFocus
End Sub

Private Sub mnu_agc2041c_Click()
    AGC2041C.Show
    AGC2041C.SetFocus
End Sub

Private Sub mnu_agc2042c_Click()
    AGC2042C.Show
    AGC2042C.SetFocus
End Sub

Private Sub mnu_agc2045c_Click()
    AGC2045C.Show
    AGC2045C.SetFocus
End Sub

Private Sub mnu_agc2046c_Click()
    AGC2046C.Show
    AGC2046C.SetFocus
End Sub

Private Sub mnu_agc2050c_Click()
    AGC2050C.Show
    AGC2050C.SetFocus
End Sub

Private Sub mnu_agc2051c_Click()
    AGC2051C.Show
    AGC2051C.SetFocus
End Sub

Private Sub mnu_agc2060c_Click()
    AGC2060C.Show
    AGC2060C.SetFocus
End Sub

Private Sub mnu_agc2070c_Click()
    AGC2070C.Show
    AGC2070C.SetFocus
End Sub

Private Sub mnu_agc2090c_Click()
    AGC2090C.Show
    AGC2090C.SetFocus
End Sub

Private Sub mnu_agc2100c_Click()
    AGC2100C.Show
    AGC2100C.SetFocus
End Sub

Private Sub mnu_agc2200c_Click()
    AGC2200C.Show
    AGC2200C.SetFocus
End Sub

Private Sub mnu_agc2201c_Click()
    AGC2201C.Show
    AGC2201C.SetFocus
End Sub

'Private Sub mnu_agc2300c_Click()
'    AGC2300C.Show
'    AGC2300C.SetFocus
'End Sub

'Private Sub mnu_agc2400c_Click()
'    AGC2400C.Show
'    AGC2400C.SetFocus
'End Sub

'Private Sub mnu_agc2430c_Click()
'    AGC2430C.Show
'    AGC2430C.SetFocus
'End Sub

Private Sub mnu_agc2410c_Click()
    AGC2410C.Show
    AGC2410C.SetFocus
End Sub

'Private Sub mnu_agc2500c_Click()
'    AGC2500C.Show
'    AGC2500C.SetFocus
'End Sub

'Private Sub mnu_agc2530c_Click()
'    AGC2530C.Show
'    AGC2530C.SetFocus
'End Sub

Private Sub mnu_agc2600c_Click()
'    AGC2600C.Show
'    AGC2600C.SetFocus
End Sub

Private Sub mnu_agc2430c_Click()
    AGC2430C.Show
    AGC2430C.SetFocus
End Sub

'Private Sub mnu_agc2432c_Click()
'    AGC2432C.Show
'    AGC2432C.SetFocus
'End Sub

Private Sub mnu_agc2440c_Click()
    AGC2440C.Show
    AGC2440C.SetFocus
End Sub
'Private Sub mnu_AGC2800C_Click()
'    AGC2800C.Show
'    AGC2800C.SetFocus
'End Sub

Private Sub mnu_agc3011c_Click()
'    AGC3011C.Show
'    AGC3011C.SetFocus
End Sub

Private Sub Mnu_agc3020c_Click()
    AGC3020C.Show
    AGC3020C.SetFocus
End Sub

Private Sub Mnu_agc3030c_Click()
    AGC3030C.Show
    AGC3030C.SetFocus
End Sub

Private Sub mnu_age2010c_Click()
    AGE2010C.Show
    AGE2010C.SetFocus
End Sub

Private Sub mnu_age2020c_Click()
    AGE2020C.Show
    AGE2020C.SetFocus
End Sub

Private Sub mnu_age2021c_Click()
    AGE2021C.Show
    AGE2021C.SetFocus
End Sub

Private Sub mnu_age2030c_Click()
    AGE2030C.Show
    AGE2030C.SetFocus
End Sub

Private Sub mnu_age2031c_Click()
    AGE2031C.Show
    AGE2031C.SetFocus
End Sub

Private Sub mnu_age2040c_Click()
    AGE2040C.Show
    AGE2040C.SetFocus
End Sub

Private Sub mnu_age2060c_Click()

End Sub

'Private Sub mnu_age2050c_Click()
'    AGE2050C.Show
'    AGE2050C.SetFocus
'End Sub

'Private Sub mnu_age2060c_Click()
'    AGE2060C.Show
'    AGE2060C.SetFocus
'End Sub

Private Sub mnu_age2080c_Click()
    AGE2080C.Show
    AGE2080C.SetFocus
End Sub

Private Sub mnu_agf2010c_Click()
    AGF2010C.Show
    AGF2010C.SetFocus
End Sub

Private Sub mnu_agf2020c_Click()
    AGF2020C.Show
    AGF2020C.SetFocus
End Sub

Private Sub mnu_agf2030c_Click()
    AGF2030C.Show
    AGF2030C.SetFocus
End Sub

Private Sub mnu_agf2031c_Click()
    AGF2031C.Show
    AGF2031C.SetFocus
End Sub

Private Sub mnu_agf2032c_Click()
    AGF2032C.Show
    AGF2032C.SetFocus
End Sub

Private Sub mnu_agf2033c_Click()
    AGF2033C.Show
    AGF2033C.SetFocus
End Sub

Private Sub mnu_agf2051c_Click()
    AGF2051C.Show
    AGF2051C.SetFocus
End Sub

Private Sub mnu_agf2040c_Click()
    AGF2040C.Show
    AGF2040C.SetFocus
End Sub

Private Sub mnu_agf2050c_Click()
    AGF2050C.Show
    AGF2050C.SetFocus
End Sub

'Private Sub mnu_agf2060c_Click()
'    AGF2060C.Show
'    AGF2060C.SetFocus
'End Sub

Private Sub mnu_agf2070c_Click()
    AGF2070C.Show
    AGF2070C.SetFocus
End Sub

Private Sub mnu_agf2080c_Click()
    AGF2080C.Show
    AGF2080C.SetFocus
End Sub

Private Sub mnu_agf2090c_Click()
    AGF2090C.Show
    AGF2090C.SetFocus
End Sub

Private Sub mnu_agf3010c_Click()
    AGF3010C.Show
    AGF3010C.SetFocus
End Sub

Private Sub mnu_agf3020c_Click()
    AGF3020C.Show
    AGF3020C.SetFocus
End Sub

Private Sub mnu_agf3030c_Click()
    AGF3030C.Show
    AGF3030C.SetFocus
End Sub

Private Sub mnu_agf3040c_Click()
    AGF3040C.Show
    AGF3040C.SetFocus
End Sub

Private Sub mnu_agf3050c_Click()
    AGF3050C.Show
    AGF3050C.SetFocus
End Sub

Private Sub mnu_agg2010c_Click()
   AGG2040C.Show
   AGG2040C.SetFocus
End Sub

Private Sub mnu_agg2020c_Click()
    AGG2020C.Show
    AGG2020C.SetFocus
End Sub

Private Sub mnu_agg2030c_Click()
    AGG2030C.Show
    AGG2030C.SetFocus
End Sub

Private Sub mnu_agg2050c_Click()
    AGG2050C.Show
    AGG2050C.SetFocus
End Sub

'Private Sub mnu_agg2040c_Click()
'    AGG2010C.Show
'    AGG2010C.SetFocus
'End Sub

Private Sub mnu_agg2060c_Click()
    AGG2060C.Show
    AGG2060C.SetFocus
End Sub

Private Sub mnu_agg2080c_Click()
    AGG2080C.Show
    AGG2080C.SetFocus
End Sub

Private Sub mnu_agt1050c_Click()
    AGT1050C.Show
    AGT1050C.SetFocus
End Sub

Private Sub mnu_agt1090c_Click()
    AGT1090C.Show
    AGT1090C.SetFocus
End Sub

Private Sub mnu_agt1100c_Click()
    AGT1100C.Show
    AGT1100C.SetFocus
End Sub

Private Sub mnu_agz1010c_Click()
'    AGZ1010C.Show
'    AGZ1010C.SetFocus
    AGC2901C.Show
    AGC2901C.SetFocus
End Sub
Private Sub Mnu_akx1010c_Click()
    AGA2011C.Show
    AGA2011C.SetFocus
End Sub

'Private Sub Mnu_akx1020c_Click()
'    AKK1010C.Show
'    AKK1010C.SetFocus
'End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
    
    MDIMain.StatusBar1.Panels(1) = "Message : "
    
    If Screen.ActiveForm.Name = "MDIMain" Then
        
        If Button.Key = "Exit" Then
            If vbYes = MsgBox(Me.Caption + " 结束 ?", vbQuestion + vbYesNo, Me.Caption) Then
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

    MDIMain.StatusBar1.Panels(1) = "Message : "
    
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



Private Sub Mnu_akx1020c_Click()

End Sub

Private Sub Mnu_Cascade_Click()
    MDIMain.StatusBar1.Panels(1) = "Message : "
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

Private Sub MNU_CGA2011C_Click()
'    CGA2011C.Show
'    CGA2011C.SetFocus
End Sub

Private Sub mnu_CGD2070C_Click()
    CGD2070C.Show
    CGD2070C.SetFocus
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
    MDIMain.StatusBar1.Panels(1) = "Message : "
    Call ActiveForm.Spread_Forzens_Cancel
End Sub

Private Sub Mnu_FrozenSetting_Click()
    'Spread Col Frozens Setting
    MDIMain.StatusBar1.Panels(1) = "Message : "
    Call ActiveForm.Spread_Forzens_Setting
End Sub

Private Sub Mnu_Help_Click()
    Dim FormLD As Boolean
    Dim CurrentForm As Form
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
    MDIMain.StatusBar1.Panels(1) = "Message : "
    MDIMain.Arrange 1
End Sub

'Private Sub Mnu_jobslog_Click()
'    AGT1030C.Show
'    AGT1030C.SetFocus
'End Sub

'Private Sub Mnu_Log_Click()
'    AGT1010C.Show
'    AGT1010C.SetFocus
'End Sub

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

Private Sub Mnu_Ref_Click()

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
    MDIMain.StatusBar1.Panels(1) = "Message : "
    Call ActiveForm.Spread_ColumnsSort
End Sub

Private Sub Mnu_Spaste_Click()
    'Spread Row Paste
    Call ActiveForm.Spread_Pst
End Sub
Private Sub Mnu_Vertical_Click()
    MDIMain.StatusBar1.Panels(1) = "提示信息："
    MDIMain.Arrange 2
End Sub

Private Sub StatusBar2_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub
