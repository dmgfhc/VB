VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "中板轧钢作业管理"
   ClientHeight    =   7845
   ClientLeft      =   1665
   ClientTop       =   1755
   ClientWidth     =   12390
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Tag             =   "G"
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
      Top             =   7380
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
            TextSave        =   "2017/5/23"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "16:59"
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
   Begin VB.Menu yard_main 
      Caption         =   "板坯库管理"
      Begin VB.Menu mnu_CGA2010C 
         Caption         =   "板坯库库存修改及查询_CGA2010C"
      End
      Begin VB.Menu mnu_CGA2011C 
         Caption         =   "板坯垛位修改及查询_CGA2011C"
      End
      Begin VB.Menu mnu_CGA2020C 
         Caption         =   "板坯库库图界面_CGA2020C"
      End
      Begin VB.Menu mnu_CGA2030C 
         Caption         =   "标准垛位管理_CGA2030C"
      End
      Begin VB.Menu mnu_CGA2060C 
         Caption         =   "库情况查询_CGA2060C"
      End
      Begin VB.Menu MNU_CGA2061C 
         Caption         =   "库详细情况查询_CGA2061C"
      End
      Begin VB.Menu mnu_CGA2080C 
         Caption         =   "板坯切割作业_CGA2080C"
      End
      Begin VB.Menu mnu_CGA2081C 
         Caption         =   "板坯库判废实绩录入_CGA2081C"
      End
      Begin VB.Menu mnu_CGA2090C 
         Caption         =   "外来板坯实绩录入_CGA2090C"
      End
      Begin VB.Menu mnu_CGA2070C 
         Caption         =   "板坯切割实绩查询界面_CGA2070C"
      End
      Begin VB.Menu mnu_CGA2100C 
         Caption         =   "板坯检验及退判实绩_CGA2100C"
      End
      Begin VB.Menu mnu_CGA2110C 
         Caption         =   "板坯入库规格修改界面_CGA2110C"
      End
      Begin VB.Menu MENU_CGA2120C 
         Caption         =   "板坯产出实绩查询_CGA2120C"
      End
      Begin VB.Menu MENU_CGA3000C 
         Caption         =   "板坯检验实绩录入_CGA3000C"
      End
   End
   Begin VB.Menu mnu_aga 
      Caption         =   "轧钢作业实绩管理"
      Begin VB.Menu mnu_cgb2010c 
         Caption         =   "加热炉装炉作业实绩查询及修改界面_CGB2010C"
      End
      Begin VB.Menu mnu_cgb2020c 
         Caption         =   "加热炉出炉作业实绩查询及修改界面_CGB2020C"
      End
      Begin VB.Menu mnu_cgb2030c 
         Caption         =   "再设计申请界面_CGB2030C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_cgc2000c 
         Caption         =   "粗轧制作业实绩查询及修改界面_CGC2000C"
      End
      Begin VB.Menu mnu_cgc2010c 
         Caption         =   "精轧作业实绩查询及修改界面_CGC2010C"
      End
      Begin VB.Menu mnu_CGC2020C 
         Caption         =   "热矫直实绩查询及修改界面_CGC2020C"
      End
      Begin VB.Menu mnu_cgc2021c 
         Caption         =   "热喷信息发送界面_CGC2021C"
      End
      Begin VB.Menu mnu_cgc2060c 
         Caption         =   "母板分段剪实绩处理_CGC2060C"
      End
   End
   Begin VB.Menu mnu_AGC 
      Caption         =   "精整作业实绩管理"
      Begin VB.Menu mnu_CGD2010C 
         Caption         =   "上冷床实绩处理界面_CGD2010C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_CGD2020C 
         Caption         =   "下冷床实绩处理界面_CGD2020C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_CGD2035C 
         Caption         =   "钢板剪切查询界面_CGD2035C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_CGD2030C 
         Caption         =   "左/右纵切实绩处理界面_CGD2030C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_CGD2031C 
         Caption         =   "圆盘切边剪实绩处理界面_CGD2031C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_cgd2040c 
         Caption         =   "取样查询及修改界面_CGD2040C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_cgd2051c 
         Caption         =   "标识指示查询界面_CGD2051C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_CGD2080C 
         Caption         =   "标识（标印、标签）打印信息发送界面_CGD2080C"
      End
      Begin VB.Menu mnu_cgd2050c 
         Caption         =   "表面检查实绩查询及修改界面_CGD2050C"
      End
      Begin VB.Menu mnu_cgd2041c 
         Caption         =   "在线钢板取样信息查询及修改界面_CGD2041C"
      End
      Begin VB.Menu mnu_AGC2420C 
         Caption         =   "理化检验委托单_AGC2420C"
      End
      Begin VB.Menu mnu_AGC2432C 
         Caption         =   "理化检验委托单-PWHT_AGC2432C"
      End
      Begin VB.Menu mnu_AGC2440C 
         Caption         =   "剪切前当班取样项目查询界面_AGC2440C"
      End
      Begin VB.Menu mnu_cge2021c 
         Caption         =   "中板未入库产品垛位管理界面_CGE2021C"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_cgc2070c 
         Caption         =   "母板分产线处理作业_CGC2070C"
      End
      Begin VB.Menu mnu_CGD2037C 
         Caption         =   "上/下线实绩处理界面_CGD2037C"
      End
      Begin VB.Menu Mnu_cgc2071c 
         Caption         =   "精整线在线查询_CGC2071C"
      End
      Begin VB.Menu Mnu_cgc2072c 
         Caption         =   "实物取样实绩录入界面_cgc2072c"
      End
      Begin VB.Menu Mnu_cgd2081c 
         Caption         =   "钢板剪切、表面检查实绩查询及修改_CGD2081C"
      End
      Begin VB.Menu Mnu_cgd2082c 
         Caption         =   "钢板标识信息发送界面_CGD2082C"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_CGZ2031C 
         Caption         =   "钢板剪切实绩集中处理界面_CGZ2031C"
      End
      Begin VB.Menu mnu_cgd2060c 
         Caption         =   "探伤实绩查询及修改界面_CGD2060C"
      End
      Begin VB.Menu mnu_CGC2050C 
         Caption         =   "火切实绩查询及修改界面(中板)_CGC2050C"
      End
      Begin VB.Menu mnu_cgd2042c 
         Caption         =   "火切钢板取样信息查询及修改界面_CGD2042C"
      End
      Begin VB.Menu mnu_AGC2051C 
         Caption         =   "钢板分板实绩修改界面_AGC2051C"
      End
   End
   Begin VB.Menu mnu_age 
      Caption         =   "成品库库管理"
      Begin VB.Menu mnu_CGE2020C 
         Caption         =   "在线钢板入库界面_CGE2020C"
      End
      Begin VB.Menu mnu_cge2030c 
         Caption         =   "钢板垛位变更及查询界面_CGE2030C"
      End
      Begin VB.Menu mnu_cge2040c 
         Caption         =   "钢板库库存现状查询_CGE2040C"
      End
      Begin VB.Menu mnu_AGC3020C 
         Caption         =   "产品退判查询_AGC3020C"
      End
   End
   Begin VB.Menu mnu_agf 
      Caption         =   "轧辊管理"
      Begin VB.Menu mnu_agf2010c 
         Caption         =   "轧辊、轴承座和轴承的入库、查询及修改界面_CGF2010C"
      End
      Begin VB.Menu mnu_agf2020c 
         Caption         =   "轧辊、轴承座和轴承的报废、查询及修改界面_CGF2020C"
      End
      Begin VB.Menu mnu_agf2030c 
         Caption         =   "轧辊磨削实绩查询及修改界面_CGF2030C"
      End
      Begin VB.Menu mnu_cgf2032c 
         Caption         =   "轧辊磨削实绩查询(按时间)_CGF2032C"
      End
      Begin VB.Menu mnu_CGF2031C 
         Caption         =   "轴承座、轴承保养的管理界面_CGF2031C"
      End
      Begin VB.Menu mnu_agf2050c 
         Caption         =   "轧辊装配实绩查询及修改界面_CGF2050C"
      End
      Begin VB.Menu mnu_cgf2051c 
         Caption         =   "轧辊装配实绩查询(按时间)_CGF2051C"
      End
      Begin VB.Menu mnu_cgf2052c 
         Caption         =   "轧辊装配实绩查询及发送_CGF2052C"
      End
      Begin VB.Menu mnu_CGF2060C 
         Caption         =   "轧辊使用实绩查询及修改界面_CGF2060C"
      End
      Begin VB.Menu mnu_cgf2070c 
         Caption         =   "轧辊使用情况查询(按时间)_CGF2070C"
      End
      Begin VB.Menu mnu_agf2040c 
         Caption         =   "轧辊/轴承座和轴承库存管理界面_CGF2040C"
      End
      Begin VB.Menu mnu_agf2060c 
         Caption         =   "轧辊/轴承座/轴承号管理_CGF2060C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_CGF2090C 
         Caption         =   "辊堆焊实绩查询及修改界面_CGF2090C"
      End
   End
   Begin VB.Menu mnu_agg 
      Caption         =   "作业指示管理"
      Begin VB.Menu mnu_ckg2040c 
         Caption         =   "轧钢计划查询界面_CGG2040C"
      End
      Begin VB.Menu mnu_ckg2030c 
         Caption         =   "精整作业指示查询界面_CKG2030C"
      End
      Begin VB.Menu Mnu_space 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_cgd2070c 
         Caption         =   "录入精整作业指示_CGD2070C"
      End
      Begin VB.Menu mnu_cKG2010C 
         Caption         =   "指示查询界面_CKG2010C"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu INQ 
      Caption         =   "各种实绩查询"
      Begin VB.Menu Mnu_cgt2000c 
         Caption         =   "加热炉实绩查询_CGT2000C"
      End
      Begin VB.Menu Mnu_cgt2010c 
         Caption         =   "轧钢实绩查询_CGT2010C"
      End
      Begin VB.Menu Mnu_cgt2020c 
         Caption         =   "母板分段实绩查询_CGT2020C"
      End
      Begin VB.Menu Mnu_cgt2030c 
         Caption         =   "双边剪实绩查询_CGT2030C"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CGT2040C 
         Caption         =   "钢板剪切实绩查询界面_CGT2040C"
      End
      Begin VB.Menu Mnu_CGT2060C 
         Caption         =   "火切实绩查询界面_CGT2060C"
      End
      Begin VB.Menu Mnu_cgt2050c 
         Caption         =   "中板厂产品检验实绩_CGT2050C"
      End
      Begin VB.Menu mnu_cgd2061c 
         Caption         =   "探伤实绩查询_CGD2061C"
      End
      Begin VB.Menu mnu_cgd2062c 
         Caption         =   "探伤日报表查询_CGD2062C"
      End
      Begin VB.Menu Mnu_CGT2100C 
         Caption         =   "综合查询_CGT2100C"
      End
      Begin VB.Menu Mnu_CGT2101C 
         Caption         =   "物料全息查询_CGT2101C"
      End
      Begin VB.Menu Mnu_CGT2102C 
         Caption         =   "订单工序作业时间查询_CGT2102C"
      End
      Begin VB.Menu Mnu_CGT2070C 
         Caption         =   "计划与实绩对比查询_CGT2070C"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CGT2001C 
         Caption         =   "堆冷时间统计报表_CGT2001C"
      End
      Begin VB.Menu Mnu_CGT2110C 
         Caption         =   "厚度公差率达标报表_CGT2110C"
      End
      Begin VB.Menu Mnu_CGT2200C 
         Caption         =   "中板图像识别实绩查询_CGT2200C"
      End
   End
   Begin VB.Menu Mnu_Other 
      Caption         =   "轧线运行实绩"
      Begin VB.Menu mnu_bkh2010c 
         Caption         =   "轧钢生产线进程现状界面_BKH2010C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_cgh2020c 
         Caption         =   "轧钢生产线停机实绩查询及修改界面_CGH2020C"
      End
      Begin VB.Menu mnu_cgh2030c 
         Caption         =   "公辅材料使用实绩查询及修改界面_CGH2030C"
      End
   End
   Begin VB.Menu Mnu_Others 
      Caption         =   "其它"
      Begin VB.Menu Mnu_cgh2040c 
         Caption         =   "板坯磅差录入界面_CGH2040C"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_cgh2050c 
         Caption         =   "质量考核报表_CGH2050C"
      End
      Begin VB.Menu Mnu_cgt2090c 
         Caption         =   "板坯未热装原因查询及录入_CGT2090C"
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

Private Sub MDIForm_Activate()

    'Call MDIMain.FormMenuSetting("Start", Toolbar_St)

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
'        Active_YN = GetSetting("NISCO", "EXE-FILE", "CG.exe")
'        sShiftSet = Gf_ShiftSet3(M_CN1)
'
'        If Active_YN = "1" Then
'
'            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'
'            MDIMain.StatusBar1.Panels(1) = "Message : "
'            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'
'        Else
'
'            Call Gp_MsgBoxDisplay("轧钢作业管理...实行中...", "W")
'            Unload Me
'            Exit Sub
'
'        End If
'
''        sUserID = "1JS1014"
''        sUserName = "杨猛"
''        MDIMain.StatusBar1.Panels(1) = "Message : "
''        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
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
'    If UBound(args) = 1 Then
'         MainFrmType = "New"
'         sUserID = args(0) ' 2012.11.09 新增  耿朝雷
'         sUserName = args(1) ' 2012.11.09 新增  耿朝雷
'         MDIMain.StatusBar1.Panels(1) = "提示信息 ：" ' 2012.11.09 新增  耿朝雷
'         MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName ' 2012.11.09 新增  耿朝雷
'    Else
'        Active_YN = GetSetting("NISCO", "EXE-FILE", "CG.exe")
'        If Active_YN = "1" Then
'            MainFrmType = "Old"
'            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'            MDIMain.StatusBar1.Panels(1) = "提示信息 ：："
'            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'        Else
'            Call Gp_MsgBoxDisplay("只能从主画面登陆...", "W")
'            Unload Me
'            Exit Sub
'        End If
'    End If  ' 2012.11.09 新增  耿朝雷
    
        

        sUserID = "1JS1005"
        sUserName = "杨猛"
        MDIMain.StatusBar1.Panels(1) = "提示信息 ："
        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName


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
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "CG.exe", ""

End Sub



Private Sub MENU_CGA3000C_Click()
    CGA3000C.Show
    CGA3000C.SetFocus
End Sub

Private Sub mnu_AGC2432C_Click()
    AGC2432C.Show
    AGC2432C.SetFocus
End Sub

Private Sub mnu_AGC2440C_Click()
    AGC2440C.Show
    AGC2440C.SetFocus
End Sub

Private Sub mnu_AGC3020C_Click()
    AGC3020C.Show
    AGC3020C.SetFocus
End Sub

Private Sub Mnu_agt1040c_Click()

End Sub

Private Sub mnu_CGA2081C_Click()
    CGA2081C.Show
    CGA2081C.SetFocus
End Sub

Private Sub mnu_cgc2000c_Click()
    CGC2000C.Show
    CGC2000C.SetFocus
End Sub

Private Sub mnu_AGC2051C_Click()
    AGC2051C.Show
    AGC2051C.SetFocus
End Sub

Private Sub mnu_AGC2420C_Click()
    AGC2430C.Show
    AGC2430C.SetFocus
End Sub

Private Sub mnu_agf2010c_Click()
    CGF2010C.Show
    CGF2010C.SetFocus
End Sub

Private Sub mnu_agf2020c_Click()
    CGF2020C.Show
    CGF2020C.SetFocus
End Sub

Private Sub mnu_agf2030c_Click()
    CGF2030C.Show
    CGF2030C.SetFocus
End Sub

Private Sub mnu_agf2040c_Click()
    CGF2040C.Show
    CGF2040C.SetFocus
End Sub

Private Sub mnu_agf2050c_Click()
    CGF2050C.Show
    CGF2050C.SetFocus
End Sub

Private Sub mnu_CGA2110C_Click()
    CGA2110C.Show
    CGA2110C.SetFocus
End Sub

Private Sub MENU_CGA2120C_Click()
    CGA2120C.Show
    CGA2120C.SetFocus
End Sub

Private Sub mnu_cgc2021c_Click()
    CGC2021C.Show
    CGC2021C.SetFocus
End Sub

Private Sub mnu_CGC2050C_Click()
    CGC2050C.Show
    CGC2050C.SetFocus
End Sub

Private Sub Mnu_cgc2070c_Click()
    CGC2070C.Show
    CGC2070C.SetFocus
End Sub

Private Sub Mnu_cgc2071c_Click()
    CGC2071C.Show
    CGC2071C.SetFocus
End Sub

Private Sub Mnu_cgc2072c_Click()
    CGC2072C.Show
    CGC2072C.SetFocus
End Sub

Private Sub mnu_CGD2037C_Click()
    CGD2037C.Show
    CGD2037C.SetFocus
End Sub

Private Sub mnu_cgd2042c_Click()
    CGD2042C.Show
    CGD2042C.SetFocus
End Sub

Private Sub mnu_cgd2060c_Click()
    CGD2060C.Show
    CGD2060C.SetFocus
End Sub

Private Sub mnu_cgd2061c_Click()
    CGD2061C.Show
    CGD2061C.SetFocus
End Sub

Private Sub mnu_cgD2070C_Click()
    CGD2070C.Show
    CGD2070C.SetFocus
End Sub

Private Sub mnu_CGD2080C_Click()
    CGD2080C.Show
    CGD2080C.SetFocus
End Sub

Private Sub Mnu_cgd2081c_Click()
    CGD2081C.Show
    CGD2081C.SetFocus
End Sub

Private Sub Mnu_cgd2082c_Click()
    CGD2082C.Show
    CGD2082C.SetFocus
End Sub

Private Sub mnu_CGE2020C_Click()
    CGE2020C.Show
    CGE2020C.SetFocus
End Sub

Private Sub mnu_cge2021c_Click()
    CGE2021C.Show
    CGE2021C.SetFocus
End Sub

Private Sub mnu_CGF2031C_Click()
    CGF2031C.Show
    CGF2031C.SetFocus
End Sub

Private Sub mnu_cgf2032c_Click()
    CGF2032C.Show
    CGF2032C.SetFocus
End Sub

Private Sub mnu_cgf2051c_Click()
    CGF2051C.Show
    CGF2051C.SetFocus
End Sub

Private Sub mnu_cgf2052c_Click()
    CGF2052C.Show
    CGF2052C.SetFocus
End Sub

Private Sub mnu_CGF2060C_Click()
    CGF2060C.Show
    CGF2060C.SetFocus
End Sub

Private Sub mnu_CGF2070c_Click()
    CGF2070C.Show
    CGF2070C.SetFocus
End Sub


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


Private Sub mnu_bkh2010c_Click()
    BKH2010C.Show
    BKH2010C.SetFocus
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

Private Sub mnu_CGA2010C_Click()
    CGA2010C.Show
    CGA2010C.SetFocus
End Sub

Private Sub mnu_CGA2011C_Click()
    CGA2011C.Show
    CGA2011C.SetFocus
End Sub

Private Sub mnu_CGA2020C_Click()
    CGA2020C.Show
    CGA2020C.SetFocus
End Sub

Private Sub mnu_CGA2030C_Click()
    CGA2030C.Show
    CGA2030C.SetFocus
End Sub

Private Sub mnu_CGA2060C_Click()
    CGA2060C.Show
    CGA2060C.SetFocus
End Sub

Private Sub MNU_CGA2061C_Click()
    CGA2061C.Show
    CGA2061C.SetFocus
End Sub

Private Sub mnu_CGA2070C_Click()
    CGA2070C.Show
    CGA2070C.SetFocus
End Sub

Private Sub mnu_CGA2080C_Click()
    CGA2080C.Show
    CGA2080C.SetFocus
End Sub

Private Sub mnu_CGA2090C_Click()
    CGA2090C.Show
    CGA2090C.SetFocus
End Sub

Private Sub mnu_CGA2100C_Click()
    CGA2100C.Show
    CGA2100C.SetFocus
End Sub

Private Sub mnu_cgb2010c_Click()
    CGB2010C.Show
    CGB2010C.SetFocus
End Sub

Private Sub mnu_cgb2020c_Click()
    CGB2020C.Show
    CGB2020C.SetFocus
End Sub

Private Sub mnu_cgb2030c_Click()
    CGB2030C.Show
    CGB2030C.SetFocus
End Sub

Private Sub mnu_cgc2010c_Click()
    CGC2010C.Show
    CGC2010C.SetFocus
End Sub

Private Sub mnu_CGC2020C_Click()
    CGC2020C.Show
    CGC2020C.SetFocus
End Sub

Private Sub mnu_CGC2060C_Click()
    CGC2060C.Show
    CGC2060C.SetFocus
End Sub

Private Sub mnu_CGD2010C_Click()
'    CGD2010C.Show
'    CGD2010C.SetFocus
End Sub

Private Sub mnu_CGD2020C_Click()
'    CGD2020C.Show
'    CGD2020C.SetFocus
End Sub

Private Sub mnu_CGD2035C_Click()
'    CGD2035C.Show
'    CGD2035C.SetFocus
End Sub

Private Sub mnu_cgd2030c_Click()
'    CGD2030C.Show
'    CGD2030C.SetFocus
End Sub

Private Sub mnu_cgd2031c_Click()
'    CGD2031C.Show
'    CGD2031C.SetFocus
End Sub

Private Sub mnu_cgd2040c_Click()
'    CGD2040C.Show
'    CGD2040C.SetFocus
End Sub

Private Sub mnu_cgd2041c_Click()
    CGD2041C.Show
    CGD2041C.SetFocus
End Sub

Private Sub mnu_cgd2050c_Click()
    CGD2050C.Show
    CGD2050C.SetFocus
End Sub

Private Sub mnu_cgd2051c_Click()
'    CGD2051C.Show
'    CGD2051C.SetFocus
End Sub

'Private Sub mnu_cgd2061c_Click()
'    CGD2061C.Show
'    CGD2061C.SetFocus
'End Sub
Private Sub mnu_cgd2062c_Click()
    CGD2062C.Show
    CGD2062C.SetFocus
End Sub


Private Sub mnu_cge2030c_Click()
    CGE2030C.Show
    CGE2030C.SetFocus
End Sub

Private Sub mnu_cge2040c_Click()
    CGE2040C.Show
    CGE2040C.SetFocus
End Sub
Private Sub mnu_CGF2090C_Click()
    CGF2090C.Show
    CGF2090C.SetFocus
End Sub

Private Sub mnu_cgh2020c_Click()
    CGH2020C.Show
    CGH2020C.SetFocus
End Sub

Private Sub mnu_cgh2030c_Click()
    CGH2030C.Show
    CGH2030C.SetFocus
End Sub

Private Sub Mnu_cgh2040c_Click()
    CGH2040C.Show
    CGH2040C.SetFocus
End Sub

Private Sub Mnu_cgh2050c_Click()
    CGH2050C.Show
    CGH2050C.SetFocus
End Sub

Private Sub Mnu_cgt2000c_Click()
    CGT2000C.Show
    CGT2000C.SetFocus
End Sub

Private Sub Mnu_CGT2001C_Click()
    CGT2001C.Show
    CGT2001C.SetFocus
End Sub

Private Sub Mnu_cgt2010c_Click()
    CGT2010C.Show
    CGT2010C.SetFocus
End Sub

Private Sub Mnu_cgt2020c_Click()
    CGT2020C.Show
    CGT2020C.SetFocus
End Sub

'Private Sub Mnu_cgt2030c_Click()
'    CGT2030C.Show
'    CGT2030C.SetFocus
'End Sub

Private Sub Mnu_CGT2040C_Click()
    CGT2040C.Show
    CGT2040C.SetFocus
End Sub

Private Sub Mnu_CGT2050C_Click()
    CGT2050C.Show
    CGT2050C.SetFocus
End Sub

Private Sub Mnu_CGT2060C_Click()
    CGT2060C.Show
    CGT2060C.SetFocus
End Sub

Private Sub Mnu_CGT2070C_Click()
    CGT2070C.Show
    CGT2070C.SetFocus
End Sub
Private Sub Mnu_cgt2090c_Click()
    CGT2090C.Show
    CGT2090C.SetFocus
End Sub

Private Sub Mnu_CGT2100C_Click()
    CGT2100C.Show
    CGT2100C.SetFocus
End Sub

Private Sub Mnu_CGT2101C_Click()
    CGT2101C.Show
    CGT2101C.SetFocus
End Sub

Private Sub Mnu_CGT2102C_Click()
    CGT2102C.Show
    CGT2102C.SetFocus
End Sub

Private Sub Mnu_CGT2110C_Click()
        CGT2110C.Show
        CGT2110C.SetFocus
End Sub

Private Sub Mnu_CGT2200C_Click()
    CGT2200C.Show
    CGT2200C.SetFocus
End Sub

Private Sub MNU_CGZ2031C_Click()
    CGZ2031C.Show
    CGZ2031C.SetFocus
End Sub

Private Sub mnu_ckg2010c_Click()
'    CKG2010C.Show
'    CKG2010C.SetFocus
End Sub

Private Sub mnu_ckg2030c_Click()
    CKG2030C.Show
    CKG2030C.SetFocus
End Sub

Private Sub mnu_ckg2040c_Click()
    CGG2040C.Show
    CGG2040C.SetFocus
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

