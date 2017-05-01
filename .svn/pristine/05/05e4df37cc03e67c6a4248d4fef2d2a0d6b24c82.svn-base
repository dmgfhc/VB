VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "板卷热处理作业管理"
   ClientHeight    =   7845
   ClientLeft      =   630
   ClientTop       =   2445
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
            TextSave        =   "2014-11-17"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "08:51"
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
   Begin VB.Menu mnu_Control 
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
   Begin VB.Menu HTM_process 
      Caption         =   "热处理管制"
      Begin VB.Menu mnu_DKA1010C 
         Caption         =   "板卷热处理指示调整_DKA1010C"
      End
      Begin VB.Menu mnu_CGD2070C 
         Caption         =   "录入精整作业指示_CGD2070C"
      End
      Begin VB.Menu mnu_ACB4110C 
         Caption         =   "精整作业对象查询_ACB4110C"
      End
   End
   Begin VB.Menu mnu_aga 
      Caption         =   "板卷热处理作业实绩管理"
      Begin VB.Menu mnu_dga1010c 
         Caption         =   "抛丸实绩查询及修改_DGA1010C"
      End
      Begin VB.Menu mnu_dga1020c 
         Caption         =   "热处理装炉作业实绩查询及修改_DGA1020C"
      End
      Begin VB.Menu mnu_dga1030c 
         Caption         =   "热处理出炉作业实绩查询及修改_DGA1030C"
      End
      Begin VB.Menu mnu_dga1040c 
         Caption         =   "冷矫直实绩查询及修改_DGA1040C"
      End
      Begin VB.Menu mnu_dga1050c 
         Caption         =   "热矫直实绩查询及修改_DGA1050C"
      End
      Begin VB.Menu mnu_dag1060c 
         Caption         =   "火剪切实绩查询及修改界面_DGA1060C"
      End
      Begin VB.Menu mnu_DGA1061C 
         Caption         =   "钢板分板实绩查询及修改界面_DGA1061C"
      End
      Begin VB.Menu mnu_AGC2020C 
         Caption         =   "表面检查实绩查询及修改_AGC2020C"
      End
      Begin VB.Menu MNU_DGA1130C 
         Caption         =   "钢板取样信息查询及修改界面_DGA1130C"
      End
      Begin VB.Menu MNU_AGC2431C 
         Caption         =   "理化检验委托单_AGC2431C"
      End
      Begin VB.Menu MNU_AGC2440C 
         Caption         =   "剪切前当班取样项目查询界面_AGC2440C"
      End
      Begin VB.Menu MNU_DGA1200C 
         Caption         =   "正火炉能耗实绩查询及修改_DGA1200C"
      End
      Begin VB.Menu MNU_DGA1210C 
         Caption         =   "钢板剩磁检查实绩查询及修改_DGA1210C"
      End
      Begin VB.Menu MNU_line30 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_DGB1000C 
         Caption         =   "淬火机作业实绩查询及修改_DGB1000C"
      End
      Begin VB.Menu MNU_DGB1010C 
         Caption         =   "淬火机水耗实绩查询及修改_DGB1010C"
      End
      Begin VB.Menu MNU_DGC1020C 
         Caption         =   "回火炉延迟实绩查询及修改_DGC1020C"
      End
      Begin VB.Menu MNU_DGC1030C 
         Caption         =   "回火炉能耗实绩查询及修改_DGC1030C"
      End
      Begin VB.Menu MNU_line31 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_DEA1020C 
         Caption         =   "热处理车间接收/转出_DEA1020C"
      End
      Begin VB.Menu mnu_DGA1180C 
         Caption         =   "热处理车间接收/转出总计查询_DGA1180C"
      End
      Begin VB.Menu mnu_AGE2030C 
         Caption         =   "钢板入库，垛位变更及库存查询界面_AGE2030C"
      End
      Begin VB.Menu mnu_CGA2011C 
         Caption         =   "物料堆放位置查询及修改界面_CGA2011C"
      End
      Begin VB.Menu MNU_ACB5070C 
         Caption         =   "产品交接信息查询及修改_ACB5070C"
      End
   End
   Begin VB.Menu mnu_agb 
      Caption         =   "|"
   End
   Begin VB.Menu mnu_Mangment 
      Caption         =   "半成品管制"
      Begin VB.Menu mnu_DKA1011C 
         Caption         =   "半成品热处理指示调整_DKA1010C"
      End
      Begin VB.Menu mnu_CGD2071C 
         Caption         =   "录入精整作业指示_CGD2070C"
      End
      Begin VB.Menu mnu_ACB4111C 
         Caption         =   "精整作业对象查询_ACB4110C"
      End
   End
   Begin VB.Menu mnu_center 
      Caption         =   "半成品作业实绩管理"
      Begin VB.Menu mnu_dga1011c 
         Caption         =   "抛丸实绩查询及修改_DGA1010C"
      End
      Begin VB.Menu mnu_dga1021c 
         Caption         =   "热处理装炉作业实绩查询及修改_DGA1020C"
      End
      Begin VB.Menu mnu_dga1031c 
         Caption         =   "热处理出炉作业实绩查询及修改_DGA1030C"
      End
      Begin VB.Menu mnu_dga1041c 
         Caption         =   "冷矫直实绩查询及修改_DGA1040C"
      End
      Begin VB.Menu mnu_dga1051c 
         Caption         =   "热矫直实绩查询及修改_DGA1050C"
      End
      Begin VB.Menu mnu_dag1063c 
         Caption         =   "火剪切实绩查询及修改界面_DGA1060C"
      End
      Begin VB.Menu mnu_DGA1064C 
         Caption         =   "钢板分板实绩查询及修改界面_DGA1061C"
      End
      Begin VB.Menu mnu_AGC2021C 
         Caption         =   "表面检查实绩查询及修改_AGC2020C"
      End
      Begin VB.Menu MNU_AGC2432C 
         Caption         =   "委托画面_AGC2432C"
      End
      Begin VB.Menu MNU_DGA1131C 
         Caption         =   "钢板取样信息查询及修改界面_DGA1130C"
      End
   End
   Begin VB.Menu MNU_AGC 
      Caption         =   "|"
   End
   Begin VB.Menu INQ 
      Caption         =   "各种实绩查询"
      Begin VB.Menu Mnu_dga1080c 
         Caption         =   "抛丸实绩查询及修改_DGA1080C"
      End
      Begin VB.Menu Mnu_dga1090c 
         Caption         =   "热处理实绩查询及修改(装炉时间)_DGA1090C"
      End
      Begin VB.Menu Mnu_dga1110c 
         Caption         =   "热矫直实绩查询及修改_DGA1110C"
      End
      Begin VB.Menu Mnu_dga1120c 
         Caption         =   "冷矫直实绩查询及修改_DGA1120C"
      End
      Begin VB.Menu mnu_DGA1140C 
         Caption         =   "火剪切实绩查询及修改界面_DGA1140C"
      End
      Begin VB.Menu mnu_DGA1150C 
         Caption         =   "热处理转入、转出实绩查询_DGA1150C"
      End
      Begin VB.Menu mnu_DGA1160C 
         Caption         =   "热处理实绩查询(出炉时间)_DGA1160C"
      End
      Begin VB.Menu mnu_DGA1170C 
         Caption         =   "热处理报表查询_DGA1170C"
      End
      Begin VB.Menu mnu_DGA1190C 
         Caption         =   "热处理产品日报表(按订单)_DGA1190C"
         Visible         =   0   'False
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
        Active_YN = GetSetting("NISCO", "EXE-FILE", "DG.exe")
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
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "DG.exe", ""

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

Private Sub mnu_ACB4110C_Click()
    ACB4110C.txt_cur_inv = "00"
    ACB4110C.Show
    ACB4110C.SetFocus
End Sub

Private Sub mnu_ACB4111C_Click()
    ACB4110C.txt_cur_inv.Text = "SG"
    ACB4110C.txt_cur_inv_nm.Text = "沙钢半成品库"
    ACB4110C.Show
    ACB4110C.SetFocus
End Sub

Private Sub MNU_ACB5070C_Click()
    ACB5070C.Show
    ACB5070C.SetFocus
End Sub

Private Sub mnu_AGC2021C_Click()
    AGC2020C.Caption = "二号线表面检查实绩查询及修改_AGC2020C"
    AGC2020C.text_cur_inv_code = "WD"
    AGC2020C.ss1.MaxRows = 0
    AGC2020C.Show
    AGC2020C.SetFocus
End Sub

Private Sub mnu_AGC2022C_Click()
    CGD2050C.Show
    CGD2050C.SetFocus
    CGD2050C.txt_PrcLine.Text = "4"
    CGD2050C.opt_LineFlag(0).Value = False
    CGD2050C.opt_LineFlag(1).Value = False
    CGD2050C.opt_LineFlag(2).Value = False
    CGD2050C.opt_LineFlag(0).Visible = False
    CGD2050C.opt_LineFlag(1).Visible = False
    CGD2050C.opt_LineFlag(2).Visible = False
    CGD2050C.SSFrame1.Visible = False
    'CGD2050C.ULabel5.Visible = False
End Sub

Private Sub MNU_AGC2431C_Click()
    AGC2430C.Show
    AGC2430C.SetFocus
End Sub

Private Sub MNU_AGC2432C_Click()
    AGC2432C.Show
    AGC2432C.SetFocus
End Sub

Private Sub MNU_AGC2440C_Click()
    AGC2440C.Show
    AGC2440C.SetFocus
End Sub

Private Sub mnu_AGE2030C_Click()
    AGE2030C.Show
    AGE2030C.SetFocus
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

Private Sub mnu_CGA2011C_Click()
    CGA2011C.Show
    CGA2011C.SetFocus
End Sub

Private Sub mnu_CGD2070C_Click()
    CGD2070C.txt_plt = "C1"
    CGD2070C.txt_cur_inv_code = "00"
    Call CGD2070C.txt_cur_inv_code_KeyUp(0, 0)
    CGD2070C.ss1.MaxRows = 0
    CGD2070C.Show
    CGD2070C.SetFocus
End Sub

Private Sub mnu_CGD2071C_Click()
    CGD2070C.txt_plt = "C1"
    CGD2070C.txt_cur_inv_code = "SG"
    Call CGD2070C.txt_cur_inv_code_KeyUp(0, 0)
    CGD2070C.ss1.MaxRows = 0
    CGD2070C.Show
    CGD2070C.SetFocus
End Sub

Private Sub Mnu_Clear_Click()
    'Screen Clera
    Call ActiveForm.Form_Cls
End Sub

Private Sub mnu_dag1060c_Click()
    DGA1060C.Caption = "一号线火剪切实绩查询及修改界面_DGA1060C"
    DGA1060C.txt_WkPlt = "C1"
    DGA1060C.cbo_PrcLine.ListIndex = 0
    DGA1060C.ss1.MaxRows = 0
    DGA1060C.ss2.MaxRows = 0
    DGA1060C.Show
    DGA1060C.SetFocus
End Sub

Private Sub mnu_dag1063c_Click()
    DGA1060C.Caption = "二号线火剪切实绩查询及修改界面_DGA1060C"
    DGA1060C.txt_WkPlt = "C1"
    DGA1060C.cbo_PrcLine.ListIndex = 1
    DGA1060C.ss1.MaxRows = 0
    DGA1060C.ss2.MaxRows = 0
    DGA1060C.Show
    DGA1060C.SetFocus
End Sub

Private Sub mnu_DEA1020C_Click()
    DEA1020C.Show
    DEA1020C.SetFocus
End Sub

Private Sub Mnu_Delete_Click()
    'Delete
    Call ActiveForm.Form_Del
End Sub

Private Sub mnu_dga1010c_Click()
    DGA1010C.Show
    DGA1010C.SetFocus
End Sub

Private Sub mnu_dga1011c_Click()
    DGA1010C.Caption = "二号线抛丸实绩查询与修改_DGA1010C"
    DGA1010C.txt_plt = "C1"
    DGA1010C.ss1.MaxRows = 0
    DGA1010C.cbo_PrcLine.ListIndex = 1
    DGA1010C.Show
    DGA1010C.SetFocus
End Sub

Private Sub mnu_dga1020c_Click()
    DGA1020C.Caption = "一号线热处理装炉作业实绩查询及修改_DGA1020C"
    DGA1020C.txt_plt = "C1"
    DGA1020C.cbo_PrcLine.ListIndex = 0
    DGA1020C.ss1.MaxRows = 0
    DGA1020C.Show
    DGA1020C.SetFocus
End Sub

Private Sub mnu_dga1021c_Click()
    DGA1020C.Caption = "二号线热处理装炉作业实绩查询及修改_DGA1020C"
    DGA1020C.txt_plt = "C1"
    DGA1020C.ss1.MaxRows = 0
    DGA1020C.cbo_PrcLine.ListIndex = 1
    DGA1020C.Show
    DGA1020C.SetFocus
End Sub

Private Sub mnu_dga1022c_Click()
    DGA1020C.Caption = "中板热处理装炉作业实绩查询及修改_DGA1020C"
    DGA1020C.txt_plt = "C3"
    DGA1020C.ss1.MaxRows = 0
    DGA1020C.cbo_PrcLine.ListIndex = 0
    DGA1020C.Show
    DGA1020C.SetFocus
End Sub

Private Sub mnu_DGA1030C_Click()
    DGA1030C.Caption = "一号线热处理出炉作业实绩查询及修改_DGA1030C"
    DGA1030C.txt_plt = "C1"
    DGA1030C.cbo_PrcLine.ListIndex = 0
    DGA1030C.ss1.MaxRows = 0
    DGA1030C.Show
    DGA1030C.SetFocus
End Sub

Private Sub mnu_dga1031c_Click()
    DGA1030C.Caption = "二号线热处理出炉作业实绩查询及修改_DGA1030C"
    DGA1030C.txt_plt = "C1"
    DGA1030C.cbo_PrcLine.ListIndex = 1
    DGA1030C.ss1.MaxRows = 0
    DGA1030C.Show
    DGA1030C.SetFocus
End Sub

Private Sub mnu_dga1032c_Click()
    DGA1030C.Caption = "中板热处理出炉作业实绩查询及修改_DGA1030C"
    DGA1030C.txt_plt = "C3"
    DGA1030C.cbo_PrcLine.ListIndex = 0
    DGA1030C.ss1.MaxRows = 0
    DGA1030C.Show
    DGA1030C.SetFocus
End Sub

Private Sub mnu_DGA1040C_Click()
    DGA1040C.Caption = "一号线冷矫直实绩查询及修改_DGA1040C"
    DGA1040C.txt_plt = "C1"
    DGA1040C.cbo_PrcLine.ListIndex = 0
    DGA1040C.ss1.MaxRows = 0
    DGA1040C.Show
    DGA1040C.SetFocus
End Sub

Private Sub mnu_dga1041c_Click()
    DGA1040C.Caption = "二号线冷矫直实绩查询及修改_DGA1040C"
    DGA1040C.txt_plt = "C1"
    DGA1040C.cbo_PrcLine.ListIndex = 1
    DGA1040C.ss1.MaxRows = 0
    DGA1040C.Show
    DGA1040C.SetFocus
End Sub

Private Sub mnu_dga1050c_Click()
    DGA1050C.Caption = "一号线热矫直实绩查询及修改_DGA1050C"
    DGA1050C.txt_plt = "C1"
    DGA1050C.cbo_PrcLine.ListIndex = 0
    DGA1050C.ss1.MaxRows = 0
    DGA1050C.Show
    DGA1050C.SetFocus
End Sub

Private Sub mnu_dga1051c_Click()
    DGA1050C.Caption = "二号线热矫直实绩查询及修改_DGA1050C"
    DGA1050C.txt_plt = "C1"
    DGA1050C.cbo_PrcLine.ListIndex = 1
    DGA1050C.ss1.MaxRows = 0
    DGA1050C.Show
    DGA1050C.SetFocus
End Sub

Private Sub mnu_dga1052c_Click()
    DGA1052C.Caption = "中板热处理矫直实绩查询及修改_DGA1052C"
    DGA1052C.txt_plt = "C3"
    DGA1052C.cbo_PrcLine.ListIndex = 0
    DGA1052C.ss1.MaxRows = 0
    DGA1052C.Show
    DGA1052C.SetFocus
End Sub

Private Sub mnu_DGA1061C_Click()
    DGA1061C.Caption = "一号线钢板分板实绩修改界面_DGA1061C"
    DGA1061C.txt_WkPlt = "C1"
    DGA1061C.cbo_PrcLine.ListIndex = 0
    DGA1061C.ss1.MaxRows = 0
    DGA1061C.ss2.MaxRows = 0
    DGA1061C.Show
    DGA1061C.SetFocus
End Sub

Private Sub mnu_AGC2020C_Click()
    AGC2020C.Caption = "一号线表面检查实绩查询及修改_AGC2020C"
    AGC2020C.text_cur_inv_code = "00"
    AGC2020C.ss1.MaxRows = 0
    AGC2020C.Show
    AGC2020C.SetFocus
End Sub

Private Sub mnu_DGA1064C_Click()
    DGA1061C.Caption = "二号线钢板分板实绩修改界面_DGA1061C"
    DGA1061C.txt_WkPlt = "C1"
    DGA1061C.cbo_PrcLine.ListIndex = 1
    DGA1061C.ss1.MaxRows = 0
    DGA1061C.ss2.MaxRows = 0
    DGA1061C.Show
    DGA1061C.SetFocus
End Sub

Private Sub Mnu_dga1080c_Click()
    DGA1080C.Show
    DGA1080C.SetFocus
End Sub

Private Sub Mnu_dga1090c_Click()
    DGA1090C.Show
    DGA1090C.SetFocus
End Sub

Private Sub Mnu_dga1110c_Click()
    DGA1110C.Show
    DGA1110C.SetFocus
End Sub

Private Sub Mnu_DGA1120C_Click()
    DGA1120C.Show
    DGA1120C.SetFocus
End Sub

Private Sub MNU_DGA1130C_Click()
    DGA1130C.Show
    DGA1130C.SetFocus
End Sub

Private Sub MNU_DGA1131C_Click()
    DGA1130C.Show
    DGA1130C.SetFocus
End Sub

Private Sub mnu_DGA1140C_Click()
    DGA1140C.Show
    DGA1140C.SetFocus
End Sub

Private Sub mnu_DGA1150C_Click()
    DGA1150C.Show
    DGA1150C.SetFocus
End Sub

Private Sub mnu_DGA1160C_Click()
    DGA1160C.Show
    DGA1160C.SetFocus
End Sub

Private Sub mnu_DGA1170C_Click()
    DGA1170C.Show
    DGA1170C.SetFocus
End Sub

Private Sub mnu_DGA1180C_Click()
    DGA1180C.Show
    DGA1180C.SetFocus
End Sub

Private Sub mnu_DGA1190C_Click()
    DGA1190C.Show
    DGA1190C.SetFocus
End Sub

Private Sub MNU_DGA1200C_Click()
    DGA1200C.Show
    DGA1200C.SetFocus
End Sub

Private Sub MNU_DGA1210C_Click()
    DGA1210C.Show
    DGA1210C.SetFocus
End Sub

Private Sub MNU_DGB1000C_Click()
    DGB1000C.Show
    DGB1000C.SetFocus
End Sub

Private Sub MNU_DGB1010C_Click()
    DGB1010C.Show
    DGB1010C.SetFocus
End Sub

Private Sub MNU_DGC1020C_Click()
    DGC1020C.Show
    DGC1020C.SetFocus
End Sub

Private Sub MNU_DGC1030C_Click()
    DGC1030C.Show
    DGC1030C.SetFocus
End Sub

Private Sub mnu_DKA1010C_Click()
    DKA1010C.Caption = "一号线指示调整_DKA1010C"
    DKA1010C.txt_plt = "C1"
    DKA1010C.cbo_PrcLine.ListIndex = 0
    DKA1010C.Show
    DKA1010C.SetFocus
    DKA1010C.ss1.MaxRows = 0
End Sub

Private Sub mnu_DKA1011C_Click()
    DKA1010C.Caption = "二号线指示调整_DKA1010C"
    DKA1010C.txt_plt = "C1"
    DKA1010C.cbo_PrcLine.ListIndex = 1
    DKA1010C.Show
    DKA1010C.SetFocus
    DKA1010C.ss1.MaxRows = 0
End Sub

Private Sub mnu_DKA1012C_Click()
    DKA1010C.Caption = "中板热处理指示调整_DKA1010C"
    DKA1010C.txt_plt = "C3"
    DKA1010C.cbo_PrcLine.ListIndex = 0
    DKA1010C.Show
    DKA1010C.SetFocus
    DKA1010C.ss1.MaxRows = 0
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

