VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQC0320C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "综合判定结果详细查询共用 - AQC0320C"
   ClientHeight    =   5160
   ClientLeft      =   795
   ClientTop       =   2595
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   Begin InDate.ULabel lbl_ORD_NO 
      Height          =   315
      Left            =   2100
      Top             =   855
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
      ForeColor       =   0
   End
   Begin VB.TextBox txt_PROD_NO 
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
      Left            =   2100
      MaxLength       =   14
      TabIndex        =   0
      Tag             =   "产品编号"
      Top             =   240
      Width           =   1905
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "产品编号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   13
      Left            =   240
      Top             =   855
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "订单号/序列号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   16
      Left            =   240
      Top             =   1380
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "产品"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   17
      Left            =   240
      Top             =   1935
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "订单厚度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   22
      Left            =   240
      Top             =   4050
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "材质等级"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   25
      Left            =   240
      Top             =   3000
      Width           =   1785
      _ExtentX        =   3149
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   26
      Left            =   240
      Top             =   3525
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "综合判定日期"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   240
      Top             =   2475
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "产品厚度"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   5040
      Top             =   855
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "钢种"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   5040
      Top             =   1380
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "订单用途"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   4
      Left            =   5040
      Top             =   1935
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "订单宽度"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   8
      Left            =   5040
      Top             =   3000
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "外观等级"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   9
      Left            =   5040
      Top             =   3525
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   $"AQC0320C.frx":0000
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   10
      Left            =   5040
      Top             =   2475
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "产品宽度"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   11
      Left            =   9510
      Top             =   855
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "标准号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   12
      Left            =   9510
      Top             =   1380
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "客户"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   14
      Left            =   9510
      Top             =   1935
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "订单长度"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   15
      Left            =   5040
      Top             =   4050
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "UST"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   21
      Left            =   9510
      Top             =   3525
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "入库日期"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   27
      Left            =   9510
      Top             =   2475
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "产品长度"
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
   Begin Threed.SSCommand cmd_AQC0330C 
      Height          =   390
      Left            =   11040
      TabIndex        =   1
      Top             =   165
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   688
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "成分／材质详细"
      BevelWidth      =   1
   End
   Begin InDate.ULabel lbl_PROD_CD 
      Height          =   315
      Left            =   2100
      Top             =   1380
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_THK 
      Height          =   315
      Left            =   2100
      Top             =   1935
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_THK2 
      Height          =   315
      Left            =   2100
      Top             =   2475
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_PROD_GRD 
      Height          =   315
      Left            =   2100
      Top             =   3000
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_DSC_DATE 
      Height          =   315
      Left            =   2100
      Top             =   3525
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_QUALITY_GRD 
      Height          =   315
      Left            =   2100
      Top             =   4050
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_ORD_ITEM 
      Height          =   315
      Left            =   3720
      Top             =   855
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_STLGRD 
      Height          =   315
      Left            =   6900
      Top             =   855
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_ENDUSE_CD 
      Height          =   315
      Left            =   6900
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_WID 
      Height          =   315
      Left            =   6900
      Top             =   1935
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_WID2 
      Height          =   315
      Left            =   6900
      Top             =   2475
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_DSC_DATE2 
      Height          =   315
      Left            =   6900
      Top             =   3525
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_SURF_GRD 
      Height          =   315
      Left            =   6900
      Top             =   3000
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_UST_DSC_RST 
      Height          =   315
      Left            =   6900
      Top             =   4050
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_STDSPEC 
      Height          =   315
      Left            =   11370
      Top             =   855
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_CUST_CD 
      Height          =   315
      Left            =   11370
      Top             =   1380
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_LEN 
      Height          =   315
      Left            =   11370
      Top             =   1935
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_LEN2 
      Height          =   315
      Left            =   11370
      Top             =   2475
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_HOUSING_DATE 
      Height          =   315
      Left            =   11370
      Top             =   3525
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
      Index           =   23
      Left            =   240
      Top             =   4560
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "材质修改等级"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   6
      Left            =   5040
      Top             =   4560
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "等级修改日期"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   18
      Left            =   9510
      Top             =   4560
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "试样编号"
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
   Begin InDate.ULabel lbl_QUALITY_UPD_GRD 
      Height          =   315
      Left            =   2100
      Top             =   4560
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_DRG_UPD_DATE 
      Height          =   315
      Left            =   6900
      Top             =   4560
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
   Begin InDate.ULabel lbl_SMP_NO 
      Height          =   315
      Left            =   11370
      Top             =   4560
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
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
Attribute VB_Name = "AQC0320C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      综合判定结果详细查询共用
'-- Program ID        AQC0320C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.08.22
'-- Description       综合判定结果详细查询共用
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
Public sPRODNO  As String           'Prod_No

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim icontrol As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "PopMaster"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_PROD_NO, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
             Call Gp_Ms_Collection(lbl_ORD_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
           Call Gp_Ms_Collection(lbl_ORD_ITEM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
             Call Gp_Ms_Collection(lbl_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
            Call Gp_Ms_Collection(lbl_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
            Call Gp_Ms_Collection(lbl_PROD_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
          Call Gp_Ms_Collection(lbl_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
            Call Gp_Ms_Collection(lbl_CUST_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
                Call Gp_Ms_Collection(lbl_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
                Call Gp_Ms_Collection(lbl_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
                Call Gp_Ms_Collection(lbl_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
               Call Gp_Ms_Collection(lbl_THK2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
               Call Gp_Ms_Collection(lbl_WID2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
               Call Gp_Ms_Collection(lbl_LEN2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
           Call Gp_Ms_Collection(lbl_PROD_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
           Call Gp_Ms_Collection(lbl_SURF_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
           Call Gp_Ms_Collection(lbl_DSC_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
          Call Gp_Ms_Collection(lbl_DSC_DATE2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
       Call Gp_Ms_Collection(lbl_HOUSING_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
        Call Gp_Ms_Collection(lbl_QUALITY_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
        Call Gp_Ms_Collection(lbl_UST_DSC_RST, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
    Call Gp_Ms_Collection(lbl_QUALITY_UPD_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
       Call Gp_Ms_Collection(lbl_DRG_UPD_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
             Call Gp_Ms_Collection(lbl_SMP_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, icontrol, rControl, aControl, lControl)
         
              
    'MASTER Collection
    
     Mc1.Add Item:="AQC0320C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=icontrol, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
 
End Sub

Private Sub cmd_AQC0330C_Click()
    AQC0330C.txt_PROD_NO.Text = txt_PROD_NO.Text
    AQC0330C.Form_Ref
    Unload Me
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

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

    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set icontrol = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing

'    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    pControl(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()
        
    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(pControl)
    If sMesg = "OK" Then
                       
        If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        End If
                   
    Else
    
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
            
    End If
    
    If Len(lbl_DSC_DATE.Caption) = 8 Then lbl_DSC_DATE.Caption = Format(lbl_DSC_DATE.Caption, "####-##-##")
    If Len(lbl_DSC_DATE2.Caption) = 8 Then lbl_DSC_DATE2.Caption = Format(lbl_DSC_DATE2.Caption, "####-##-##")
    If Len(lbl_HOUSING_DATE.Caption) = 8 Then lbl_HOUSING_DATE.Caption = Format(lbl_HOUSING_DATE.Caption, "####-##-##")
    If Len(lbl_DRG_UPD_DATE.Caption) = 8 Then lbl_DRG_UPD_DATE.Caption = Format(lbl_DRG_UPD_DATE.Caption, "####-##-##")
    
End Sub

Public Sub Form_Pro()
       
    If Gf_Mc_Authority(sAuthority, Mc1) Then
       ' txt_ins_emp.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub


