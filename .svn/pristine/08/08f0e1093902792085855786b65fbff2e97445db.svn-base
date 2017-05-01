VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AED1020C 
   Caption         =   "CCR材查询/选定（pop_up）_AED1020C"
   ClientHeight    =   4365
   ClientLeft      =   1335
   ClientTop       =   2955
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   6270
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2355
      TabIndex        =   5
      Text            =   " "
      Top             =   3015
      Width           =   1905
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2355
      TabIndex        =   4
      Text            =   " "
      Top             =   2310
      Width           =   1155
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4500
      TabIndex        =   3
      Text            =   " "
      Top             =   2310
      Width           =   1155
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   2535
      Top             =   1215
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "起始"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Height          =   315
      Left            =   3585
      Top             =   1635
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "－－"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4500
      TabIndex        =   2
      Text            =   " "
      Top             =   1605
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2355
      TabIndex        =   1
      Top             =   1605
      Width           =   1155
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2355
      TabIndex        =   0
      Top             =   705
      Width           =   1905
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   510
      Top             =   3030
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      Caption         =   "最大编制量"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Left            =   510
      Top             =   2340
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      Caption         =   "厚 度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Left            =   510
      Top             =   1635
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      Caption         =   "宽 度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Left            =   510
      Top             =   705
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      Caption         =   "产品分类"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Left            =   4650
      Top             =   1215
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "终止"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
      Left            =   3585
      Top             =   2310
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "－－"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
Attribute VB_Name = "AED1020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
