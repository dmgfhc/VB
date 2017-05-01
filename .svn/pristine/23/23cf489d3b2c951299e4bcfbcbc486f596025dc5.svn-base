VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CGB2021C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "质量设计现状详细查询 "
   ClientHeight    =   9975
   ClientLeft      =   1440
   ClientTop       =   2355
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   14550
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   4485
      Left            =   0
      TabIndex        =   2
      Top             =   615
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   7911
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_ORD_LEN_MAX 
         Height          =   315
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3630
         Width           =   1155
      End
      Begin VB.TextBox txt_MILL_TMP_MIN_2 
         Height          =   315
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1890
         Width           =   1095
      End
      Begin VB.TextBox txt_MILL_TMP_MAX_2 
         Height          =   315
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1890
         Width           =   1155
      End
      Begin VB.TextBox txt_MILL_TMP_TGT_2 
         Alignment       =   1  'Right Justify
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
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1890
         Width           =   1035
      End
      Begin VB.TextBox txt_SLAB_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2055
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   26
         Top             =   180
         Width           =   2250
      End
      Begin VB.TextBox txt_ORD_LEN 
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3630
         Width           =   1035
      End
      Begin VB.TextBox txt_ORD_LEN_MIN 
         Height          =   315
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3630
         Width           =   1095
      End
      Begin VB.TextBox txt_MILL_THK 
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3210
         Width           =   1035
      End
      Begin VB.TextBox txt_ORD_THK_MAX 
         Height          =   345
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3180
         Width           =   1155
      End
      Begin VB.TextBox txt_ORD_THK_MIN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3210
         Width           =   1095
      End
      Begin VB.TextBox txt_COOL_TMP_MAX_2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1155
      End
      Begin VB.TextBox txt_ROLL_TEMP_MAX_2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10440
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   19
         Top             =   2340
         Width           =   1170
      End
      Begin VB.TextBox txt_COOL_TMP_TGT_2 
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2760
         Width           =   1035
      End
      Begin VB.TextBox txt_COOL_TMP_MIN_2 
         Height          =   315
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_ROLL_TEMP_TGT_2 
         Height          =   345
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2310
         Width           =   1035
      End
      Begin VB.TextBox txt_TOTAL_RATE_MAX_2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txt_TOTAL_RATE_MIN_2 
         Alignment       =   1  'Right Justify
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
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1035
         Width           =   885
      End
      Begin VB.TextBox txt_AVE_RATE_MAX_2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txt_AVE_RATE_MIN_2 
         Alignment       =   1  'Right Justify
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
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1035
      End
      Begin VB.TextBox txt_AVE_TEMP_MAX_2 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox txt_HCR_KND 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txt_AVE_TEMP_MIN_2 
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
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   885
      End
      Begin VB.TextBox txt_CR_MILL_RATET1_2 
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
         Left            =   7050
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1455
         Width           =   915
      End
      Begin VB.TextBox txt_HEATTWO_TEMP_MIN_2 
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
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   630
         Width           =   1035
      End
      Begin VB.TextBox txt_CR_MILL_TMPT1_2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1035
      End
      Begin VB.TextBox txt_HEATTWO_TEMP_MAX_2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txt_ROLL_TEMP_MIN_2 
         Height          =   285
         Left            =   7050
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2340
         Width           =   1095
      End
      Begin InDate.ULabel ULabel1 
         Height          =   285
         Left            =   5130
         Top             =   2340
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         Caption         =   "入水温度最小值"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   32
         Left            =   150
         Top             =   630
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   "加热二段温度工艺范围"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   20
         Left            =   5130
         Top             =   630
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "均热段温度工艺范围"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   18
         Left            =   5130
         Top             =   1455
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "中间坯厚度范围"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   3
         Left            =   5130
         Top             =   180
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "冷/热装"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   5
         Left            =   150
         Top             =   1035
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   "均热时间工艺范围"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   36
         Left            =   150
         Top             =   1440
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   "精轧机开轧温度"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   46
         Left            =   8580
         Top             =   3630
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单长度最大值"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   1
         Left            =   3420
         Top             =   630
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   556
         Caption         =   "-"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   4
         Left            =   7950
         Top             =   630
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   556
         Caption         =   "-"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   7
         Left            =   3420
         Top             =   1020
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   556
         Caption         =   "-"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   8
         Left            =   5130
         Top             =   1035
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "总加热时间范围"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   9
         Left            =   7950
         Top             =   1020
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   556
         Caption         =   "-"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   150
         Top             =   2310
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   "入水温度目标值"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   5130
         Top             =   2760
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "返红温度最小值"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   6
         Left            =   8580
         Top             =   2760
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "返红温度最大值"
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   150
         Top             =   2760
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   "返红温度目标值"
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
      Begin InDate.ULabel ULabel5 
         Height          =   345
         Left            =   8580
         Top             =   3180
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         Caption         =   "订单厚度最大值"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   0
         Left            =   5130
         Top             =   3210
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "订单厚度最小值"
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   150
         Top             =   3210
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   "轧件目标厚度"
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   150
         Top             =   3630
         Width           =   2145
         _ExtentX        =   3784
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   5130
         Top             =   3630
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "订单长度最小值"
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
      Begin InDate.ULabel ULabel12 
         Height          =   345
         Index           =   2
         Left            =   150
         Top             =   180
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         Caption         =   "板坯号"
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   5130
         Top             =   1890
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "终轧温度最小值"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   10
         Left            =   150
         Top             =   1890
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   "终轧温度目标值"
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   8580
         Top             =   1890
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "终轧温度最大值"
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   8580
         Top             =   2340
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "入水温度最大值"
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14490
      TabIndex        =   0
      Top             =   0
      Width           =   14550
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
            TabIndex        =   11
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8730
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":04B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":07BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":0CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":0DAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":109E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":1550
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9330
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":1664
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":1964
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":1BB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":1C99
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":1EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":1FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CGB2021C.frx":2215
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CGB2021C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量设计
'-- Program Name      质量设计现状详细查询
'-- Program ID        AQB0020C (Master-AQB0010C)
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       质量设计现状详细查询
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
Dim SLAB_NO As String            'Save The Previous ORD_NO For Next Use

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "PopMaster"

           'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary )", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_slab_no, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HCR_KND, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_HEATTWO_TEMP_MIN_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_HEATTWO_TEMP_MAX_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_AVE_TEMP_MIN_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_AVE_TEMP_MAX_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_AVE_RATE_MIN_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_AVE_RATE_MAX_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_TOTAL_RATE_MIN_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_TOTAL_RATE_MAX_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_CR_MILL_RATET1_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_CR_MILL_TMPT1_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ROLL_TEMP_MAX_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ROLL_TEMP_MIN_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ROLL_TEMP_TGT_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_COOL_TMP_MAX_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_COOL_TMP_MIN_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_COOL_TMP_TGT_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_MILL_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ORD_THK_MIN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ORD_THK_MAX, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ORD_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ORD_LEN_MIN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ORD_LEN_MAX, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_MILL_TMP_TGT_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_MILL_TMP_MIN_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_MILL_TMP_MAX_2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
    
    
    'MASTER Collection
     Mc1.Add Item:="CGB2021C.P_REFER", Key:="P-R"
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
    Else
     Call Form_Ref
     Call Gp_Ms_ControlLock(Mc1("pControl"), True)
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

'    Call AQB0010C.Form_Ref

End Sub

Public Sub Form_Ref()
 On Error GoTo Refer_Err

    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()
    
    SLAB_NO = Trim(txt_slab_no.Text)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    MenuTool.Buttons(4).Enabled = False    'save
    MenuTool.Buttons(5).Enabled = False    'delete
    MenuTool.Buttons(7).Enabled = False    'Copy
    MenuTool.Buttons(8).Enabled = True    'Paste

    pControl(1).SetFocus

End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "Clear"              'Clear
            Call Form_Cls
        Case "Refer"               'refer
            Call Form_Ref
        Case "Copy"               'Copy
            Call Form_Copy
        Case "Paste"              'Paste
            Call Form_Paste
        Case "Exit"               'Exit
            Call Form_Exit
    End Select

End Sub


Public Sub Popup_Menu_Setting()

    Select Case Mid(sAuthority, 2, 3)

        Case "000"      'No Authority
            MenuTool.Buttons(4).Enabled = False                     'Save
            MenuTool.Buttons(5).Enabled = False                     'Delete
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "001"      'Delete Authority
            MenuTool.Buttons(4).Enabled = False                     'Save
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "010"      'Update Authority
            MenuTool.Buttons(5).Enabled = False                     'Delete
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "011"      'Update, Delete Authority
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "100"      'Insert Authority
            MenuTool.Buttons(5).Enabled = False                     'Delete

        Case "101"      'Insert, Delete Authority

        Case "110"      'Insert, Update Authority
            MenuTool.Buttons(5).Enabled = False                     'Delete

        Case "111"      'Insert, Update, Delete Authority

    End Select

End Sub

Private Sub Form_Copy()
    
    SLAB_NO = Trim(txt_slab_no.Text)

End Sub

Private Sub Form_Paste()
    txt_slab_no.Text = ""
    txt_slab_no.Text = SLAB_NO

End Sub

