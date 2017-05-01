VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQA0071C 
   Caption         =   "客户特殊要求共用信息输入_AQA0071C"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   10320
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Img_DRT_CNF_TYP 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1950
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   3570
      Width           =   300
   End
   Begin VB.TextBox txt_DRT_CNF_TYP 
      Height          =   315
      Left            =   2970
      MaxLength       =   1
      TabIndex        =   28
      Top             =   3540
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt_MILL_STD_NO 
      Height          =   300
      Left            =   6810
      TabIndex        =   14
      Top             =   3135
      Width           =   1725
   End
   Begin VB.TextBox txt_MLT_STD_NO 
      Height          =   300
      Left            =   1950
      TabIndex        =   13
      Top             =   3135
      Width           =   1725
   End
   Begin VB.TextBox txt_NISCO_QUALITY_NO 
      Height          =   300
      Left            =   6810
      TabIndex        =   12
      Top             =   2715
      Width           =   1725
   End
   Begin MSComctlLib.Toolbar MenuTool 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1349
      ButtonWidth     =   1244
      ButtonHeight    =   1349
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Line3"
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_DEV_STD_CD 
      Height          =   300
      Left            =   1950
      TabIndex        =   11
      Top             =   2715
      Width           =   1725
   End
   Begin VB.TextBox txt_STDSPEC_YY 
      Height          =   300
      Left            =   6810
      TabIndex        =   10
      Top             =   2280
      Width           =   1725
   End
   Begin VB.TextBox txt_STDSPEC 
      Height          =   300
      Left            =   1950
      TabIndex        =   9
      Top             =   2280
      Width           =   2745
   End
   Begin VB.TextBox txt_ENDUSE_CD 
      Height          =   300
      Left            =   1950
      TabIndex        =   7
      Top             =   1800
      Width           =   1035
   End
   Begin VB.TextBox txt_ENDUSE_NAME 
      Height          =   300
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   1725
   End
   Begin VB.TextBox txt_PROD_CD 
      Height          =   300
      Left            =   1950
      TabIndex        =   3
      Top             =   1410
      Width           =   1035
   End
   Begin VB.TextBox txt_PROD_NAME 
      Height          =   300
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1410
      Width           =   1725
   End
   Begin VB.TextBox txt_len_max 
      Height          =   300
      Left            =   8160
      TabIndex        =   21
      Top             =   4490
      Width           =   1335
   End
   Begin VB.TextBox txt_len_min 
      Height          =   300
      Left            =   6810
      TabIndex        =   20
      Top             =   4490
      Width           =   1335
   End
   Begin VB.TextBox txt_WID_MAX 
      Height          =   300
      Left            =   3300
      TabIndex        =   19
      Top             =   4490
      Width           =   1335
   End
   Begin VB.TextBox txt_wid_min 
      Height          =   300
      Left            =   1950
      TabIndex        =   18
      Top             =   4490
      Width           =   1335
   End
   Begin VB.TextBox txt_thk_min 
      Height          =   300
      Left            =   1950
      TabIndex        =   16
      Top             =   4060
      Width           =   1335
   End
   Begin VB.TextBox txt_thk_max 
      Height          =   300
      Left            =   3300
      TabIndex        =   17
      Top             =   4060
      Width           =   1335
   End
   Begin VB.TextBox txt_STEEL_GRD_Name 
      Height          =   300
      Left            =   8340
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox txt_STEEL_GRD 
      Height          =   300
      Left            =   6810
      TabIndex        =   5
      Top             =   1440
      Width           =   1515
   End
   Begin VB.TextBox txt_upd_emp 
      Height          =   300
      Left            =   6810
      TabIndex        =   26
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox txt_upd_date 
      Height          =   300
      Left            =   6810
      TabIndex        =   25
      Top             =   5430
      Width           =   2775
   End
   Begin VB.TextBox txt_ins_emp 
      Height          =   300
      Left            =   1950
      TabIndex        =   24
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox txt_ins_date 
      Height          =   300
      Left            =   1950
      TabIndex        =   23
      Top             =   5430
      Width           =   2775
   End
   Begin VB.TextBox txt_cust_sq 
      Height          =   300
      Left            =   3600
      TabIndex        =   1
      Top             =   870
      Width           =   1035
   End
   Begin VB.TextBox txt_cust_no 
      Height          =   300
      Left            =   1950
      TabIndex        =   0
      Top             =   870
      Width           =   1635
   End
   Begin VB.TextBox txt_cust_spec_detail 
      Height          =   300
      Left            =   1950
      TabIndex        =   22
      Top             =   4920
      Width           =   7635
   End
   Begin VB.TextBox txt_cust_name 
      Height          =   300
      Left            =   4650
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   870
      Width           =   3135
   End
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   1
      Left            =   60
      Top             =   870
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "客户特殊要求编号"
      Alignment       =   0
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
      Index           =   0
      Left            =   4920
      Top             =   1440
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "钢种"
      Alignment       =   0
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
      Index           =   2
      Left            =   60
      Top             =   1800
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "订单用途"
      Alignment       =   0
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
      Index           =   3
      Left            =   60
      Top             =   2280
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "标准编号"
      Alignment       =   0
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
      Index           =   4
      Left            =   4920
      Top             =   2280
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "发布年度"
      Alignment       =   0
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
      Left            =   60
      Top             =   2715
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "代表性交付条件标准"
      Alignment       =   0
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
      Index           =   6
      Left            =   60
      Top             =   4060
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "厚度组"
      Alignment       =   0
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
      Left            =   60
      Top             =   4490
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "宽度组"
      Alignment       =   0
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
      Left            =   4920
      Top             =   4490
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "长度组"
      Alignment       =   0
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
      Left            =   60
      Top             =   4920
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "适用客户"
      Alignment       =   0
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
      Left            =   60
      Top             =   5430
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "编制日期"
      Alignment       =   0
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
      Left            =   60
      Top             =   5880
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "编制人"
      Alignment       =   0
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
      Left            =   4920
      Top             =   5430
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "修改日期"
      Alignment       =   0
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
      Left            =   4920
      Top             =   5880
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "修改人"
      Alignment       =   0
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
      Left            =   60
      Top             =   1410
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "品种"
      Alignment       =   0
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
      Index           =   15
      Left            =   4920
      Top             =   2715
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "企标编号"
      Alignment       =   0
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
      Index           =   16
      Left            =   60
      Top             =   3135
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "炼钢／连铸规程编号"
      Alignment       =   0
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
      Index           =   17
      Left            =   4920
      Top             =   3135
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "轧钢规程编号"
      Alignment       =   0
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
      Index           =   18
      Left            =   60
      Top             =   3570
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "直接投入"
      Alignment       =   0
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
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   45
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":018D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":0392
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":0545
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":074B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":0890
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":09DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":0B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":0CF3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1680
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   45
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":1045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":119E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":133F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":14CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":1670
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":17B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":190B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0071C.frx":1A7D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   0
      X2              =   11970
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   30
      X2              =   12000
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   12000
      Y1              =   5310
      Y2              =   5310
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   12000
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "AQA0071C"
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
'-- Program Name
'-- Program ID        AQA0070C (Master-AQA0071C)
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
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
       Call Gp_Ms_Collection(txt_cust_no, "p", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cust_sq, "p", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PROD_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PROD_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STEEL_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STEEL_GRD_Name, " ", " ", " ", " ", "r", "l", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ENDUSE_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ENDUSE_NAME, " ", " ", " ", "", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STDSPEC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STDSPEC_YY, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_DEV_STD_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_NISCO_QUALITY_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_MLT_STD_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_MILL_STD_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_DRT_CNF_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_thk_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_thk_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_wid_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_WID_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_len_min, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_len_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cust_spec_detail, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ins_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_INS_EMP, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_upd_emp, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

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
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
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
    
    txt_cust_name.Text = ""
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
        txt_INS_EMP.Text = sUserID
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


Private Sub txt_cust_no_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        'DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_cust_no
        DD.rControl.Add Item:=txt_cust_name
        
        DD.nameType = "1"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
        
        
        'Call Gf_Common_DD(M_CN1, KeyCode)
        Call Gf_Customer_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_cust_no)) = txt_cust_no.MaxLength Then
        'Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        txt_cust_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_no.Text), 1)
    Else
        txt_cust_name.Text = ""
    End If

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
        Img_DRT_CNF_TYP.Picture = ImageList1.ListImages(9).Picture
    Else
        Img_DRT_CNF_TYP.Picture = Nothing
    End If
End Sub

Private Sub txt_ENDUSE_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = Left(txt_PROD_CD.Text, 1)
        DD.rControl.Add Item:=txt_ENDUSE_CD
        DD.rControl.Add Item:=txt_ENDUSE_NAME

        DD.nameType = "2"

        Call Gf_Usage_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_ENDUSE_CD)) = txt_ENDUSE_CD.MaxLength Then
        txt_ENDUSE_NAME.Text = Gf_UsageNameFind(M_CN1, Left(txt_PROD_CD.Text, 1), Trim(txt_ENDUSE_CD.Text))
    Else
        txt_ENDUSE_NAME.Text = ""
    End If
    
End Sub

Private Sub txt_LEN_MAX_KeyPress(KeyAscii As Integer)

  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub txt_len_max_Validate(Cancel As Boolean)
        
        If Len(Trim(txt_len_max.Text)) <> 0 Then
            If Not (txt_Max_Check(txt_len_max.Text, txt_len_min.Text)) Then
                
               MsgBox ("请检查长度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub txt_LEN_MIN_KeyPress(KeyAscii As Integer)
  
  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub txt_len_min_Validate(Cancel As Boolean)
        
        If Len(Trim(txt_len_min.Text)) <> 0 Then
            If Not (txt_Max_Check(txt_len_max.Text, txt_len_min.Text)) Then
                
               MsgBox ("请检查长度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub txt_MILL_STD_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_MILL_STD_NO
       
        Call Gf_ROLL_STD_DD(M_CN1, KeyCode)
        
        Exit Sub
    
    End If
End Sub

Private Sub txt_MLT_STD_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_MLT_STD_NO
        
        Call Gf_Melt_STD_DD(M_CN1, KeyCode)
        
        Exit Sub
    
    End If

End Sub

Private Sub txt_NISCO_QUALITY_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_NISCO_QUALITY_NO
        
        Call Gf_Nisco_STD_DD(M_CN1, KeyCode)
        
        Exit Sub
    
    End If

End Sub

Private Sub txt_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_PROD_CD
        DD.rControl.Add Item:=txt_PROD_NAME

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_PROD_CD)) = txt_PROD_CD.MaxLength Then
        txt_PROD_NAME.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_PROD_CD.Text), 2)
    Else
        txt_PROD_NAME.Text = ""
    End If

End Sub

Private Sub txt_STDSPEC_Change()
    
    If Len(Trim(txt_STDSPEC.Text)) = 0 Then
        
        txt_STDSPEC_YY.Text = ""
    
    End If

End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_STDSPEC
        DD.rControl.Add Item:=txt_STDSPEC_YY
        
        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        
        Exit Sub
    
    End If
End Sub



Private Sub txt_STEEL_GRD_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        'DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_STEEL_GRD
        DD.rControl.Add Item:=txt_STEEL_GRD_Name
        
        DD.nameType = "2"
        
        
        'Call Gf_Common_DD(M_CN1, KeyCode)
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_STEEL_GRD)) = txt_STEEL_GRD.MaxLength Then
        'Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        txt_STEEL_GRD_Name.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_STEEL_GRD.Text))
    Else
        txt_STEEL_GRD_Name.Text = ""
    End If

End Sub

Private Sub txt_THK_MAX_KeyPress(KeyAscii As Integer)
  
  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub txt_thk_max_Validate(Cancel As Boolean)
        
        If Len(Trim(txt_thk_max.Text)) <> 0 Then
            If Not (txt_Max_Check(txt_thk_max.Text, txt_thk_min.Text)) Then
                
               MsgBox ("请检查厚度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub txt_THK_MIN_KeyPress(KeyAscii As Integer)
  
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

Private Sub txt_thk_min_Validate(Cancel As Boolean)
        
        If Len(Trim(txt_thk_min.Text)) <> 0 Then
            If Not (txt_Max_Check(txt_thk_max.Text, txt_thk_min.Text)) Then
                
               MsgBox ("请检查厚度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub txt_WID_MAX_KeyPress(KeyAscii As Integer)

  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub txt_WID_MAX_Validate(Cancel As Boolean)
        
        If Len(Trim(txt_WID_MAX.Text)) <> 0 Then
            If Not (txt_Max_Check(txt_WID_MAX.Text, txt_wid_min.Text)) Then
                
               MsgBox ("请检查宽度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If

End Sub

Private Sub txt_WID_MIN_KeyPress(KeyAscii As Integer)

  KeyAscii = txt_KeyPress(KeyAscii)

End Sub

Private Sub txt_wid_min_Validate(Cancel As Boolean)
        
        If Len(Trim(txt_wid_min.Text)) <> 0 Then
            If Not (txt_Max_Check(txt_WID_MAX.Text, txt_wid_min.Text)) Then
                
               MsgBox ("请检查宽度组最小值和最大值，后者不能小与前者")
               
               Cancel = True
    
            End If
        
        Else
               MsgBox ("请输入数值")
               
               Cancel = True
        
        End If
End Sub
