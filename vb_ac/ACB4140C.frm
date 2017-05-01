VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB4140C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "录入评审对象处理_ACB4140C"
   ClientHeight    =   9225
   ClientLeft      =   690
   ClientTop       =   1875
   ClientWidth     =   15315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_slab_no2 
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
      Left            =   3270
      MaxLength       =   10
      TabIndex        =   26
      Tag             =   "板坯号"
      Top             =   180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9165
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   16166
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACB4140C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   3705
         Left            =   0
         TabIndex        =   10
         Top             =   5460
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   6535
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BackColor       =   12632319
         PaneTree        =   "ACB4140C.frx":0052
         Begin TabDlg.SSTab SSTab1 
            Height          =   3645
            Left            =   30
            TabIndex        =   19
            Top             =   30
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   6429
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "板坯单位处理"
            TabPicture(0)   =   "ACB4140C.frx":0084
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "SSSplitter5"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "板坯批量处理"
            TabPicture(1)   =   "ACB4140C.frx":00A0
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "SSSplitter4"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin SSSplitter.SSSplitter SSSplitter4 
               Height          =   3285
               Left            =   30
               TabIndex        =   20
               Top             =   330
               Width           =   15075
               _ExtentX        =   26591
               _ExtentY        =   5794
               _Version        =   196609
               SplitterBarWidth=   2
               SplitterBarJoinStyle=   0
               SplitterBarAppearance=   0
               BorderStyle     =   0
               BackColor       =   14737632
               PaneTree        =   "ACB4140C.frx":00BC
               Begin Threed.SSPanel SSPanel2 
                  Height          =   495
                  Left            =   0
                  TabIndex        =   21
                  Top             =   0
                  Width           =   15075
                  _ExtentX        =   26591
                  _ExtentY        =   873
                  _Version        =   196609
                  BackColor       =   14737918
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin VB.TextBox txt_title_est_cd 
                     BackColor       =   &H00C0FFFF&
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
                     Left            =   1470
                     MaxLength       =   4
                     TabIndex        =   24
                     Tag             =   "处理代码"
                     Top             =   90
                     Width           =   555
                  End
                  Begin VB.TextBox txt_title_est_nm 
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
                     Left            =   2025
                     MaxLength       =   60
                     TabIndex        =   23
                     Tag             =   "处理代码"
                     Top             =   90
                     Width           =   3480
                  End
                  Begin VB.TextBox txt_title_est_comm 
                     BackColor       =   &H00C0FFFF&
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
                     Left            =   7035
                     MaxLength       =   200
                     TabIndex        =   22
                     Tag             =   "处理代码"
                     Top             =   90
                     Width           =   7920
                  End
                  Begin InDate.ULabel ULabel1 
                     Height          =   315
                     Left            =   60
                     Top             =   90
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   556
                     Caption         =   "代表处理"
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
                  Begin InDate.ULabel ULabel9 
                     Height          =   315
                     Left            =   5640
                     Top             =   90
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   556
                     Caption         =   "代表处理详细"
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
               End
               Begin FPSpread.vaSpread ss3 
                  Height          =   2760
                  Left            =   0
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   525
                  Width           =   15075
                  _Version        =   393216
                  _ExtentX        =   26591
                  _ExtentY        =   4868
                  _StockProps     =   64
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MaxCols         =   9
                  MaxRows         =   1
                  RetainSelBlock  =   0   'False
                  SpreadDesigner  =   "ACB4140C.frx":010E
               End
            End
            Begin SSSplitter.SSSplitter SSSplitter5 
               Height          =   3285
               Left            =   -74970
               TabIndex        =   27
               Top             =   330
               Width           =   15075
               _ExtentX        =   26591
               _ExtentY        =   5794
               _Version        =   196609
               SplitterBarWidth=   2
               SplitterBarJoinStyle=   0
               SplitterBarAppearance=   0
               BorderStyle     =   0
               BackColor       =   14737632
               PaneTree        =   "ACB4140C.frx":0653
               Begin Threed.SSPanel SSPanel3 
                  Height          =   840
                  Left            =   0
                  TabIndex        =   28
                  Top             =   0
                  Width           =   15075
                  _ExtentX        =   26591
                  _ExtentY        =   1482
                  _Version        =   196609
                  BackColor       =   14737918
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin VB.TextBox txt_title_est_comm1 
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
                     Left            =   7035
                     MaxLength       =   200
                     TabIndex        =   34
                     Tag             =   "处理代码"
                     Top             =   450
                     Width           =   7920
                  End
                  Begin VB.TextBox txt_title_est_nm1 
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
                     Left            =   2025
                     MaxLength       =   60
                     TabIndex        =   33
                     Tag             =   "处理代码"
                     Top             =   450
                     Width           =   3480
                  End
                  Begin VB.TextBox txt_title_est_cd1 
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
                     Left            =   1470
                     MaxLength       =   4
                     TabIndex        =   32
                     Tag             =   "处理代码"
                     Top             =   450
                     Width           =   555
                  End
                  Begin VB.TextBox txt_title_reason_comm1 
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
                     Left            =   7035
                     MaxLength       =   200
                     TabIndex        =   31
                     Tag             =   "处理代码"
                     Top             =   90
                     Width           =   7920
                  End
                  Begin VB.TextBox txt_title_reason_nm1 
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
                     Left            =   2025
                     MaxLength       =   60
                     TabIndex        =   30
                     Tag             =   "处理代码"
                     Top             =   90
                     Width           =   3480
                  End
                  Begin VB.TextBox txt_title_reason_cd1 
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
                     Left            =   1470
                     MaxLength       =   4
                     TabIndex        =   29
                     Tag             =   "处理代码"
                     Top             =   90
                     Width           =   555
                  End
                  Begin InDate.ULabel ULabel5 
                     Height          =   315
                     Left            =   60
                     Top             =   90
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   556
                     Caption         =   "代表原因"
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
                  Begin InDate.ULabel ULabel7 
                     Height          =   315
                     Left            =   5640
                     Top             =   90
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   556
                     Caption         =   "代表原因详细"
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
                  Begin InDate.ULabel ULabel8 
                     Height          =   315
                     Left            =   60
                     Top             =   450
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   556
                     Caption         =   "代表处理"
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
                  Begin InDate.ULabel ULabel10 
                     Height          =   315
                     Left            =   5640
                     Top             =   450
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   556
                     Caption         =   "代表处理详细"
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
               End
               Begin FPSpread.vaSpread ss2 
                  Height          =   2415
                  Left            =   0
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   870
                  Width           =   15075
                  _Version        =   393216
                  _ExtentX        =   26591
                  _ExtentY        =   4260
                  _StockProps     =   64
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MaxCols         =   35
                  MaxRows         =   1
                  RetainSelBlock  =   0   'False
                  SpreadDesigner  =   "ACB4140C.frx":06A5
               End
            End
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   5400
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   9525
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         PaneTree        =   "ACB4140C.frx":150A
         Begin Threed.SSPanel SSPanel1 
            Height          =   900
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   1588
            _Version        =   196609
            BackColor       =   14737632
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_rec_sts 
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
               Left            =   14670
               MaxLength       =   1
               TabIndex        =   17
               Tag             =   "处理代码"
               Top             =   90
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.TextBox txt_est_nm 
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
               Left            =   6795
               MaxLength       =   60
               TabIndex        =   5
               Tag             =   "处理代码"
               Top             =   480
               Width           =   3480
            End
            Begin VB.TextBox txt_est_cd 
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
               Left            =   6240
               MaxLength       =   4
               TabIndex        =   4
               Tag             =   "处理代码"
               Top             =   480
               Width           =   555
            End
            Begin VB.TextBox txt_reason_nm 
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
               Left            =   6795
               MaxLength       =   60
               TabIndex        =   1
               Tag             =   "原因代码"
               Top             =   90
               Width           =   3480
            End
            Begin VB.TextBox txt_reason_cd 
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
               Left            =   6240
               MaxLength       =   4
               TabIndex        =   0
               Tag             =   "原因代码"
               Top             =   90
               Width           =   555
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   645
               Left            =   10440
               TabIndex        =   7
               Top             =   90
               Width           =   4635
               Begin Threed.SSOption opt_all 
                  Height          =   285
                  Left            =   210
                  TabIndex        =   13
                  Top             =   210
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  _Version        =   196609
                  Font3D          =   1
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "全部"
               End
               Begin Threed.SSOption opt_wait 
                  Height          =   285
                  Left            =   2220
                  TabIndex        =   14
                  Top             =   210
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _Version        =   196609
                  Font3D          =   1
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "等待确定"
               End
               Begin Threed.SSOption opt_complete 
                  Height          =   285
                  Left            =   3390
                  TabIndex        =   15
                  Top             =   210
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _Version        =   196609
                  Font3D          =   1
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "确定完成"
               End
               Begin Threed.SSOption opt_in_wait 
                  Height          =   285
                  Left            =   1020
                  TabIndex        =   18
                  Top             =   210
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _Version        =   196609
                  Font3D          =   1
                  ForeColor       =   255
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "处理等待"
                  Value           =   -1
               End
            End
            Begin VB.TextBox txt_slab_no1 
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
               Left            =   1590
               MaxLength       =   10
               TabIndex        =   6
               Tag             =   "板坯号"
               Top             =   95
               Width           =   1395
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   180
               Top             =   480
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "录入日期"
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
               Left            =   180
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
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
            Begin InDate.ULabel ULabel2 
               Height          =   315
               Left            =   4830
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "原因代码"
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
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   4830
               Top             =   480
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "处理代码"
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
               ForeColor       =   16711680
            End
            Begin InDate.UDate dpt_ins_date_fr 
               Height          =   315
               Left            =   1590
               TabIndex        =   2
               Tag             =   "处理日期"
               Top             =   480
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483630
               BackColor       =   16777215
            End
            Begin InDate.UDate dpt_ins_date_to 
               Height          =   315
               Left            =   3180
               TabIndex        =   3
               Tag             =   "处理日期"
               Top             =   480
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483630
               BackColor       =   16777215
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "~"
               Height          =   120
               Left            =   3045
               TabIndex        =   16
               Top             =   540
               Width           =   90
            End
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   4470
            Left            =   0
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   930
            Width           =   15225
            _Version        =   393216
            _ExtentX        =   26855
            _ExtentY        =   7885
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   37
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACB4140C.frx":155C
         End
      End
   End
End
Attribute VB_Name = "ACB4140C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   PROCESS MANAGEMENT
'-- Program Name      SLAB DELIBERATION PROCESS EVENT
'-- Program ID        ACB4140C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2009.8.18
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

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lRowSelect As Long
Dim sOptFl, sTabFl As Boolean

Private Sub Form_Define()
        
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_slab_no1, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_reason_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_reason_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dpt_ins_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dpt_ins_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_est_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_est_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_REC_STS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_slab_no2, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
  
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    For i = 2 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next i
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4140C.P_REFER1", Key:="P-R"
    sc1.Add Item:="ACB4140C.P_ONEROW1", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, "p", "n", "m", "i", "a", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, "p", "n", "m", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACB4140C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:="ACB4140C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:="ACB4140C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="ACB4140C.P_MODIFY2", Key:="P-M"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss2, 8, True)
    Call Gp_Sp_ColHidden(ss2, 11, True)
    Call Gp_Sp_ColHidden(ss2, 17, True)
    Call Gp_Sp_ColHidden(ss2, 21, True)
    Call Gp_Sp_ColHidden(ss2, 34, True)
    Call Gp_Sp_ColHidden(ss2, 35, True)
    
    Call Gp_Sp_ColHidden(ss3, 8, True)
    
    lRowSelect = 0
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet

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
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    Call Gp_Sp_Setting(Sc3.Item("Spread"))
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(sc2.Item("Spread"), 2)
    Call Gp_Sp_HdColColor(sc2.Item("Spread"), 13)
    
    Call Gp_Sp_HdColColor(Sc3.Item("Spread"), 1)
    Call Gp_Sp_HdColColor(Sc3.Item("Spread"), 3)
    
    TXT_REC_STS.Text = "1"
    
    dpt_ins_date_fr.RawData = ""
    dpt_ins_date_to.RawData = ""
    txt_title_est_cd.Text = ""
    txt_title_est_nm.Text = ""
    txt_title_est_comm.Text = ""
    
    SSTab1.Tab = 0
    sTabFl = False
    sOptFl = False
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    
End Sub

Public Sub Form_Cls()
    
    If SSTab1.Tab = 0 Then
        If Not Gf_Sp_Cls(sc2) Then Exit Sub
    Else
        If Not Gf_Sp_Cls(Sc3) Then Exit Sub
    End If
        
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
    End If
    
    dpt_ins_date_fr.RawData = ""
    dpt_ins_date_to.RawData = ""
    TXT_REC_STS.Text = "1"
    opt_in_wait.Value = True
    txt_title_est_cd.Text = ""
    txt_title_est_nm.Text = ""
    txt_title_est_comm.Text = ""

End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
    End If
    
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gp_Ms_Cls(Mc2("rControl"))
    txt_title_est_cd.Text = ""
    txt_title_est_nm.Text = ""
    txt_title_est_comm.Text = ""
    txt_title_reason_cd1.Text = ""
    txt_title_reason_nm1.Text = ""
    txt_title_reason_comm1.Text = ""
    txt_title_est_cd1.Text = ""
    txt_title_est_nm1.Text = ""
    txt_title_est_comm1.Text = ""
            
End Sub

Public Sub Form_Pro()
        
    Dim i As Long
    Dim sQuery As String
    Dim OutParam(2, 4) As Variant
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    If SSTab1.Tab = 0 Then
    
        If Len(txt_title_reason_cd1.Text) = 4 Or txt_title_reason_comm1.Text <> "" Or _
               Len(txt_title_est_cd1.Text) = 4 Or txt_title_est_comm1.Text <> "" Then
            
            sQuery = "{call ACB4140C.P_MODIFY1 ( 'U', '" + txt_slab_no2.Text + "','" + _
                                                           txt_title_reason_cd1.Text + "','" + txt_title_reason_comm1.Text + "','" + _
                                                           txt_title_est_cd1.Text + "','" + txt_title_est_comm1.Text + "',?,?) }"
                                                           
            If Not Gf_Ms_ExecQuery(OutParam, M_CN1, sQuery) Then Exit Sub
            
            sQuery = "{call ACB4140C.P_ONEROW1 ( '" + txt_slab_no2.Text + "') }"
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, ss1, lRowSelect)
            
        End If
            
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc2) Then
            
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
            ss2.OperationMode = OperationModeNormal
            
            If ss2.MaxRows = 0 Then
                Call Gp_Sp_DeleteRow(ss1, lRowSelect)
            Else
                sQuery = "{call ACB4140C.P_ONEROW1 ( '" + txt_slab_no2.Text + "') }"
                Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, ss1, lRowSelect)
            End If
            
        End If
        
        For i = 1 To ss2.MaxRows
        
            ss2.Row = i
            ss2.Col = 2
            
            If ss2.Text = "1010" Then
                Call Gp_Sp_BlockLock(ss2, 19, 19, i, i, False)
            Else
                Call Gp_Sp_BlockLock(ss2, 19, 19, i, i, True)
            End If
            
        Next i
        
    Else
    
        If Len(txt_title_est_cd.Text) = 4 And txt_title_est_comm.Text <> "" Then
            
            For i = 1 To ss1.MaxRows
            
                ss1.Row = i
                ss1.Col = 0
                
                If ss1.Text <> "" Then
                    ss1.Col = 1
                    sQuery = "{call ACB4140C.P_MODIFY1 ( 'T', '" + ss1.Text + "','','','" + _
                                                                   txt_title_est_cd.Text + "','" + txt_title_est_comm.Text + "',?,?) }"
                                                           
                    If Not Gf_Ms_ExecQuery(OutParam, M_CN1, sQuery) Then Exit Sub
                End If
            
            Next i
            
        End If
        
        If Sp_Process(M_CN1, Sc3) Then
    
            If Gf_Sp_Refer(M_CN1, sc1, Mc1) Then
                ss1.OperationMode = OperationModeNormal
            End If
            
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
            
            ss3.MaxRows = 0
            txt_title_est_cd.Text = ""
            txt_title_est_nm.Text = ""
            txt_title_est_comm.Text = ""
            
        End If
        
    End If
    
End Sub

Public Sub Form_Ins()
    
    If SSTab1.Tab = 0 Then Exit Sub
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    'Call INPUT_SLAB_INFO(ss2.ActiveRow)

End Sub

Public Sub Spread_Cpy()

    If SSTab1.Tab = 0 Then Exit Sub
    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    If SSTab1.Tab = 0 Then Exit Sub

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    'Call INPUT_SLAB_INFO(ss2.ActiveRow)
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()
    
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub

Private Sub opt_all_Click(Value As Integer)

    If sOptFl Then
        sOptFl = False
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        If Not Gf_Sp_Cls(sc2) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    Else
        If Not Gf_Sp_Cls(Sc3) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    End If
    
    If Not Gf_Sp_Cls(sc1) Then Exit Sub
    
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    rControl(1).SetFocus
    
    If opt_all.Value Then
        opt_all.ForeColor = &HFF&
        opt_in_wait.ForeColor = &H80000012
        opt_wait.ForeColor = &H80000012
        opt_complete.ForeColor = &H80000012
        TXT_REC_STS.Text = "A"
    End If
    
    txt_title_reason_cd1.Text = ""
    txt_title_reason_nm1.Text = ""
    txt_title_reason_comm1.Text = ""
    txt_title_est_cd1.Text = ""
    txt_title_est_nm1.Text = ""
    txt_title_est_comm1.Text = ""
    txt_title_est_cd.Text = ""
    txt_title_est_nm.Text = ""
    txt_title_est_comm.Text = ""

End Sub

Private Sub opt_complete_Click(Value As Integer)

    If sOptFl Then
        sOptFl = False
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        If Not Gf_Sp_Cls(sc2) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    Else
        If Not Gf_Sp_Cls(Sc3) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    End If
    
    If Not Gf_Sp_Cls(sc1) Then Exit Sub
    
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    rControl(1).SetFocus
    
    If opt_complete.Value Then
        opt_complete.ForeColor = &HFF&
        opt_in_wait.ForeColor = &H80000012
        opt_wait.ForeColor = &H80000012
        opt_all.ForeColor = &H80000012
        TXT_REC_STS.Text = "3"
    End If
    
    txt_title_reason_cd1.Text = ""
    txt_title_reason_nm1.Text = ""
    txt_title_reason_comm1.Text = ""
    txt_title_est_cd1.Text = ""
    txt_title_est_nm1.Text = ""
    txt_title_est_comm1.Text = ""
    txt_title_est_cd.Text = ""
    txt_title_est_nm.Text = ""
    txt_title_est_comm.Text = ""
    
End Sub

Private Sub opt_in_wait_Click(Value As Integer)

    If sOptFl Then
        sOptFl = False
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        If Not Gf_Sp_Cls(sc2) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    Else
        If Not Gf_Sp_Cls(Sc3) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    End If
    
    If Not Gf_Sp_Cls(sc1) Then Exit Sub
    
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    rControl(1).SetFocus

    If opt_in_wait.Value Then
        opt_in_wait.ForeColor = &HFF&
        opt_wait.ForeColor = &H80000012
        opt_complete.ForeColor = &H80000012
        opt_all.ForeColor = &H80000012
        TXT_REC_STS.Text = "1"
    End If
    
    txt_title_reason_cd1.Text = ""
    txt_title_reason_nm1.Text = ""
    txt_title_reason_comm1.Text = ""
    txt_title_est_cd1.Text = ""
    txt_title_est_nm1.Text = ""
    txt_title_est_comm1.Text = ""
    txt_title_est_cd.Text = ""
    txt_title_est_nm.Text = ""
    txt_title_est_comm.Text = ""


End Sub

Private Sub opt_wait_Click(Value As Integer)

    If sOptFl Then
        sOptFl = False
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        If Not Gf_Sp_Cls(sc2) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    Else
        If Not Gf_Sp_Cls(Sc3) Then
            sOptFl = True
            If opt_all.ForeColor = &HFF& Then
                opt_all.Value = True
            ElseIf opt_in_wait.ForeColor = &HFF& Then
                opt_in_wait.Value = True
            ElseIf opt_wait.ForeColor = &HFF& Then
                opt_wait.Value = True
            ElseIf opt_complete.ForeColor = &HFF& Then
                opt_complete.Value = True
            End If
            Exit Sub
        End If
    End If
    
    If Not Gf_Sp_Cls(sc1) Then Exit Sub
    
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    rControl(1).SetFocus
    
    If opt_wait.Value Then
        opt_wait.ForeColor = &HFF&
        opt_in_wait.ForeColor = &H80000012
        opt_complete.ForeColor = &H80000012
        opt_all.ForeColor = &H80000012
        TXT_REC_STS.Text = "2"
    End If
    
    txt_title_reason_cd1.Text = ""
    txt_title_reason_nm1.Text = ""
    txt_title_reason_comm1.Text = ""
    txt_title_est_cd1.Text = ""
    txt_title_est_nm1.Text = ""
    txt_title_est_comm1.Text = ""
    txt_title_est_cd.Text = ""
    txt_title_est_nm.Text = ""
    txt_title_est_comm.Text = ""

    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Public Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i As Integer
    
    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    If SSTab1.Tab = 0 Then
        
        If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
        
        If lRowSelect = 0 Then
            lRowSelect = Row
        Else
            ss1.Row = lRowSelect
            ss1.Col = 0
            ss1.Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRowSelect, lRowSelect)
            lRowSelect = Row
        End If
        
        ss1.Row = Row
        ss1.Col = 0
        ss1.Text = "选择"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
        
        ss1.Col = 1
        txt_slab_no2.Text = ss1.Text
        ss1.Col = 8
        txt_title_reason_cd1.Text = ss1.Text
        ss1.Col = 9
        txt_title_reason_nm1.Text = ss1.Text
        ss1.Col = 10
        txt_title_reason_comm1.Text = ss1.Text
        ss1.Col = 11
        txt_title_est_cd1.Text = ss1.Text
        ss1.Col = 12
        txt_title_est_nm1.Text = ss1.Text
        ss1.Col = 13
        txt_title_est_comm1.Text = ss1.Text
        
        Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
        ss2.OperationMode = OperationModeNormal
        
        For i = 1 To ss2.MaxRows
        
            ss2.Row = i
            ss2.Col = 2
            
            If ss2.Text = "1010" Then
                Call Gp_Sp_BlockLock(ss2, 19, 19, i, i, False)
            Else

'                If ss2.Text = "9090" Then
'                    Call Gp_Sp_BlockLock(ss2, 1, ss2.MaxCols, I, I, True)
'                Else
                    Call Gp_Sp_BlockLock(ss2, 19, 19, i, i, True)
'                End If
                
            End If
            
        Next i
        
    Else
    
        ss1.Row = Row
        ss1.Col = 0
            
        If ss1.Text <> "选择" Then
            ss1.Col = 0
            ss1.Text = "选择"
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
        Else
           ss1.Col = 0
           ss1.Text = ""
           Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
        End If
    
    End If
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    
    ACB4160C.txt_slab_no1.Text = ss1.Text
    ACB4160C.opt_all.Value = True
    ACB4160C.opt_all.ForeColor = &HFF&
    ACB4160C.opt_wait.ForeColor = &H80000012
    ACB4160C.opt_complete.ForeColor = &H80000012
    ACB4160C.TXT_REC_STS.Text = "A"

    Call ACB4160C.Form_Ref
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

'    If ss2.MaxRows < 1 Or Row = 0 Then Exit Sub
'
'    ss2.Row = Row
'    ss2.Col = 2
'
'    If ss2.Text = "9090" Then
'
'        ss2.Col = 20
'
'        If ss2.Text = "" Then
'            ss2.Col = 1
'            ACB4070C.txt_change_no.Text = ss2.Text
'            ACB4070C.opt_end.Value = True
'            ACB4070C.lbl_dir.Caption = "==>"
'            Call ACB4070C.Form_Ref
'        End If
'
'    End If

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    Dim i As Integer
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
    End If
    
    If Mode = 1 And Col = 5 Then

        If ss2.Text = "0" Or ss2.Text = "" Then
        
            For i = 1 To ss2.MaxRows
            
                ss2.Row = i
                ss2.Col = 5
                
                If i <> Row Then
                    ss2.Text = "0"
                End If
            
            Next i
            
            ss2.Row = Row
            
            ss2.Col = 2
            txt_title_reason_cd1.Text = ss2.Text
            ss2.Col = 3
            txt_title_reason_nm1.Text = ss2.Text
            ss2.Col = 4
            txt_title_reason_comm1.Text = ss2.Text
            
            ss2.Col = 5
            ss2.Tag = ss2.Text
            
        Else
        
            txt_title_reason_cd1.Text = ""
            txt_title_reason_nm1.Text = ""
            txt_title_reason_comm1.Text = ""
        
        End If
        
    ElseIf Mode = 1 And Col = 16 Then
        
        If ss2.Text = "0" Or ss2.Text = "" Then
        
            For i = 1 To ss2.MaxRows
            
                ss2.Row = i
                ss2.Col = 16
                
                If i <> Row Then
                    ss2.Text = "0"
                End If
            
            Next i
            
            ss2.Row = Row
            
            ss2.Col = 13
            txt_title_est_cd1.Text = ss2.Text
            ss2.Col = 14
            txt_title_est_nm1.Text = ss2.Text
            ss2.Col = 15
            txt_title_est_comm1.Text = ss2.Text
            
            ss2.Col = 16
            ss2.Tag = ss2.Text
            
        Else
        
            txt_title_est_cd1.Text = ""
            txt_title_est_nm1.Text = ""
            txt_title_est_comm1.Text = ""
        
        End If
        
    End If

End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
        Call INPUT_SLAB_INFO(ss2.MaxRows)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code, sQuery As String

    If ss2.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss2.ActiveCol
    
        Case 2    'REASON_CD
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss2
                
                DD.sWitch = "SP"
                DD.sKey = "C0017"
                DD.rControl.Add Item:=2
                DD.rControl.Add Item:=3
                
                DD.nameType = "2"
'                DD.sWhere = "AND CD  <>  '9090' "
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Else
            
                ss2.Col = ss2.ActiveCol
                
                If Len(Trim(ss2.Text)) = ss2.TypeMaxEditLen Then
                
                    sTemp_Code = ss2.Text
                    ss2.Col = 3
                    ss2.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(sTemp_Code), 2)
                    
                Else
                
                    ss2.Col = 3
                    ss2.Text = ""
                    
                End If
            
            End If
            
        Case 8    'STLGRD
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss2
                
                DD.sWitch = "SP"
                DD.rControl.Add Item:=8
                DD.rControl.Add Item:=9
                
                DD.nameType = "2"
                
                Call Gf_Stlgrd_DD(M_CN1, KeyCode)
                
            Else
            
                ss2.Col = ss2.ActiveCol
                
                If Len(Trim(ss2.Text)) = ss2.TypeMaxEditLen Then
                
                    sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(ss2.Text) + "'"
                    ss2.Col = ss2.Col + 1
                    ss2.Text = Gf_CodeFind(M_CN1, sQuery)
                    ss2.Col = ss2.Col - 1
                    
                Else
                
                    ss2.Col = 9
                    ss2.Text = ""
                    
                End If
            
            End If
            

        Case 13    'EST_CD
        
            If KeyCode = vbKeyF4 Then
            
                sTemp_Code = ss2.Text
            
                Set DD.sPname = Me.ss2
                
                DD.sWitch = "SP"
                DD.sKey = "C0018"
                DD.rControl.Add Item:=13
                DD.rControl.Add Item:=14
                
                DD.nameType = "2"
'                DD.sWhere = "AND CD  <>  '9090' "
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
                ss2.Col = 13
                
                If sTemp_Code <> ss2.Text And ss2.Text <> "" Then
                
                    ss2.Col = 17
                    ss2.Text = Gf_CodeFind(M_CN1, "SELECT STLGRD FROM FP_SLAB WHERE SLAB_NO = '" + Trim(txt_slab_no2.Text) + "'")
                    ss2.Col = 18
                    ss2.Text = Gf_CodeFind(M_CN1, "SELECT A.STEEL_GRD_DETAIL FROM QP_NISCO_CHMC A, FP_SLAB B WHERE B.STLGRD = A.STLGRD AND B.SLAB_NO = '" + Trim(txt_slab_no2.Text) + "'")
                    ss2.Col = 21
                    ss2.Text = sUserID
                    ss2.Col = 22
                    ss2.Text = sUserName
                
                End If
                
            Else
            
                ss2.Col = ss2.ActiveCol
                
                If Len(Trim(ss2.Text)) = ss2.TypeMaxEditLen Then
                
                    sTemp_Code = ss2.Text
                    ss2.Col = 14
                    ss2.Text = Gf_ComnNameFind(M_CN1, "C0018", Trim(sTemp_Code), 2)
                    
                    ss2.Col = 17
                    ss2.Text = Gf_CodeFind(M_CN1, "SELECT STLGRD FROM FP_SLAB WHERE SLAB_NO = '" + Trim(txt_slab_no2.Text) + "'")
                    ss2.Col = 18
                    ss2.Text = Gf_CodeFind(M_CN1, "SELECT A.STEEL_GRD_DETAIL FROM QP_NISCO_CHMC A, FP_SLAB B  WHERE B.STLGRD = A.STLGRD AND B.SLAB_NO = '" + Trim(txt_slab_no2.Text) + "'")
                    ss2.Col = 21
                    ss2.Text = sUserID
                    ss2.Col = 22
                    ss2.Text = sUserName
            
                Else
                
                    ss2.Col = 14
                    ss2.Text = ""
                    ss2.Col = 17
                    ss2.Text = ""
                    ss2.Col = 18
                    ss2.Text = ""
                    ss2.Col = 21
                    ss2.Text = ""
                    ss2.Col = 22
                    ss2.Text = ""
                    
                End If
            
            End If
            
'            ss2.Col = 13
'
'            If ss2.Text = "9090" Then
'                ss2.Col = 0
'                ss2.Text = ""
'                ss2.Col = 13
'                ss2.Text = ""
'                ss2.Col = 14
'                ss2.Text = ""
'                ss2.Col = 17
'                ss2.Text = ""
'                ss2.Col = 18
'                ss2.Text = ""
'                ss2.Col = 21
'                ss2.Text = ""
'                ss2.Col = 22
'                ss2.Text = ""
'            End If
            
'        Case 15   'STLGRD
'
'            If KeyCode = vbKeyF4 Then
'
'                Set DD.sPname = Me.ss2
'
'                DD.sWitch = "SP"
'                DD.rControl.Add Item:=15
'                DD.rControl.Add Item:=16
'
'                DD.nameType = "2"
'
'                Call Gf_Stlgrd_DD(M_CN1, KeyCode)
'
'            Else
'
'                ss2.Col = ss2.ActiveCol
'
'                If Len(Trim(ss2.Text)) = ss2.TypeMaxEditLen Then
'
'                    sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(ss2.Text) + "'"
'                    ss2.Col = ss2.Col + 1
'                    ss2.Text = Gf_CodeFind(M_CN1, sQuery)
'                    ss2.Col = ss2.Col - 1
'
'                Else
'
'                    ss2.Col = 16
'                    ss2.Text = ""
'
'                End If
'
'            End If
            
    End Select

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Sc3.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    Dim i As Integer
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
    End If
    
    If Mode = 1 And Col = 7 Then

        If ss3.Text = "0" Or ss3.Text = "" Then
        
            For i = 1 To ss3.MaxRows
            
                ss3.Row = i
                ss3.Col = 7
                
                If i <> Row Then
                    ss3.Text = "0"
                End If
            
            Next i
            
            ss3.Row = Row
            
            ss3.Col = 3
            txt_title_est_cd.Text = ss3.Text
            ss3.Col = 4
            txt_title_est_nm.Text = ss3.Text
            ss3.Col = 5
            txt_title_est_comm.Text = ss3.Text
            
            ss3.Col = 7
            ss3.Tag = ss3.Text
            
        Else
        
            txt_title_est_cd.Text = ""
            txt_title_est_nm.Text = ""
            txt_title_est_comm.Text = ""
        
        End If
        
    End If
    
End Sub

Private Sub ss3_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss3_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code, sQuery As String

    If ss3.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss3.ActiveCol
    
        Case 1    'REASON_CD
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss3
                
                DD.sWitch = "SP"
                DD.sKey = "C0017"
                DD.rControl.Add Item:=1
                DD.rControl.Add Item:=2
                
                DD.nameType = "2"
'                DD.sWhere = "AND CD  <>  '9090' "
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Else
            
                ss3.Col = ss3.ActiveCol
                
                If Len(Trim(ss3.Text)) = ss3.TypeMaxEditLen Then
                
                    sTemp_Code = ss3.Text
                    ss3.Col = 2
                    ss3.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(sTemp_Code), 2)
                    
                Else
                
                    ss3.Col = 2
                    ss3.Text = ""
                    
                End If
            
            End If
            
            ss3.Col = 1
            
            If ss3.Text = "1010" Then
                Call Gp_Sp_BlockLock(ss3, 6, 6, ss3.ActiveRow, ss3.ActiveRow, False)
            Else
                Call Gp_Sp_BlockLock(ss3, 6, 6, ss3.ActiveRow, ss3.ActiveRow, True)
            End If
            
'            ss3.Col = 1
'
'            If ss3.Text = "9090" Then
'                ss3.Col = 1
'                ss3.Text = ""
'                ss3.Col = 2
'                ss3.Text = ""
'            End If
            
        Case 3    'EST_CD
        
            If KeyCode = vbKeyF4 Then
            
                sTemp_Code = ss3.Text
            
                Set DD.sPname = Me.ss3
                
                DD.sWitch = "SP"
                DD.sKey = "C0018"
                DD.rControl.Add Item:=3
                DD.rControl.Add Item:=4
                
                DD.nameType = "2"
'                DD.sWhere = "AND CD  <>  '9090' "
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Else
            
                ss3.Col = ss3.ActiveCol
                
                If Len(Trim(ss3.Text)) = ss3.TypeMaxEditLen Then
                
                    sTemp_Code = ss3.Text
                    ss3.Col = 4
                    ss3.Text = Gf_ComnNameFind(M_CN1, "C0018", Trim(sTemp_Code), 2)
                    
                Else
                
                    ss3.Col = 4
                    ss3.Text = ""
                    
                End If
            
            End If
            
'            ss3.Col = 3
'
'            If ss3.Text = "9090" Then
'                ss3.Col = 3
'                ss3.Text = ""
'                ss3.Col = 4
'                ss3.Text = ""
'            End If

    End Select

End Sub

Private Sub ss3_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss3
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub SSSplitter1_Resize(ByVal BorderPanes As SSSplitter.Panes)

    SSSplitter5.Height = SSSplitter1.Panes(1).Height - 420
    SSSplitter4.Height = SSSplitter1.Panes(1).Height - 420
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    Dim i As Integer
    
    If sTabFl Then
        sTabFl = False
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
    
        If Gf_Sp_Change(Proc_Sc, sc2) Then
        
            Call Gf_Sp_Cls(Sc3)
            txt_title_est_cd.Text = ""
            txt_title_est_nm.Text = ""
            txt_title_est_comm.Text = ""
            
            For i = 1 To ss1.MaxCols
            
                ss1.Row = i
                ss1.Col = 0
            
                If ss1.Text <> "" Then
                    ss1.Text = ""
                    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, i, i)
                End If
           
            Next i
        
        Else
        
            sTabFl = True
            SSTab1.Tab = PreviousTab
        
        End If
        
        Call MenuTool_ReSet
        
    Else
    
        If Gf_Sp_Change(Proc_Sc, Sc3) Then
        
            Call Gf_Sp_Cls(sc2)
            txt_title_reason_cd1.Text = ""
            txt_title_reason_nm1.Text = ""
            txt_title_reason_comm1.Text = ""
            txt_title_est_cd1.Text = ""
            txt_title_est_nm1.Text = ""
            txt_title_est_comm1.Text = ""
            
            If lRowSelect <> 0 Then
                ss1.Row = lRowSelect
                ss1.Col = 0
                ss1.Text = ""
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRowSelect, lRowSelect)
                lRowSelect = 0
            End If
           
        Else
        
            sTabFl = True
            SSTab1.Tab = PreviousTab
        
        End If
        
        Call MenuTool_ReSet
    
    End If
    
End Sub

Private Sub txt_est_cd_DblClick()

    Call txt_est_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_est_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0018"
        DD.rControl.Add Item:=txt_est_cd
        DD.rControl.Add Item:=txt_est_nm
        
        DD.nameType = "2"
'        DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_est_cd)) = txt_est_cd.MaxLength Then
            txt_est_nm.Text = Gf_ComnNameFind(M_CN1, "C0018", Trim(txt_est_cd.Text), 2)
        Else
            txt_est_nm.Text = ""
        End If
        
    End If
    
'    If txt_est_cd.Text = "9090" Then
'        txt_est_cd.Text = ""
'        txt_est_nm.Text = ""
'    End If

End Sub

Private Sub txt_reason_cd_DblClick()

    Call txt_reason_cd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_reason_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0017"
        DD.rControl.Add Item:=txt_reason_cd
        DD.rControl.Add Item:=txt_reason_nm
        
        DD.nameType = "2"
'        DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_reason_cd)) = txt_reason_cd.MaxLength Then
            txt_reason_nm.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(txt_reason_cd.Text), 2)
        Else
            txt_reason_nm.Text = ""
        End If
        
    End If
    
'    If txt_reason_cd.Text = "9090" Then
'        txt_reason_cd.Text = ""
'        txt_reason_nm.Text = ""
'    End If

End Sub

Private Sub txt_title_est_cd_DblClick()

    Call txt_title_est_cd_KeyUp(vbKeyF4, 0)
        
End Sub

Private Sub txt_title_est_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0018"
        DD.rControl.Add Item:=txt_title_est_cd
        DD.rControl.Add Item:=txt_title_est_nm
        
        DD.nameType = "2"
'        DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_title_est_cd)) = txt_title_est_cd.MaxLength Then
            txt_title_est_nm.Text = Gf_ComnNameFind(M_CN1, "C0018", Trim(txt_title_est_cd.Text), 2)
        Else
            txt_title_est_nm.Text = ""
        End If
        
    End If
    
'    If txt_title_est_cd.Text = "9090" Then
'        txt_title_est_cd.Text = ""
'        txt_title_est_nm.Text = ""
'    End If

End Sub

Private Sub INPUT_SLAB_INFO(ActiveRow As Long)

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If M_CN1.State = 0 Then
        If GF_DbConnect = False Then Exit Sub
    End If
    
    Set AdoRs = New ADODB.Recordset

    sQuery = "SELECT STLGRD, GF_STLGRD_DETAIL(STLGRD), THK, WID, LEN, WGT, PROC_CD, PROD_CD, ORD_FL, ORD_NO, ORD_ITEM  FROM  FP_SLAB  WHERE SLAB_NO = '" + txt_slab_no2.Text + "'"
    
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                
            Else
                 ss2.Row = ActiveRow
                 ss2.Col = 8
                 ss2.Text = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
                 ss2.Col = 9
                 ss2.Text = IIf(VarType(AdoRs.Fields(1)) = vbNull, "", AdoRs.Fields(1))
                 
                 ss2.Col = 23
                 ss2.Text = Trim(AdoRs.Fields(2))
                 ss2.Col = 24
                 ss2.Text = Trim(AdoRs.Fields(3))
                 ss2.Col = 25
                 ss2.Text = Trim(AdoRs.Fields(4))
                 ss2.Col = 26
                 ss2.Text = Trim(AdoRs.Fields(5))
                 ss2.Col = 27
                 ss2.Text = IIf(VarType(AdoRs.Fields(6)) = vbNull, "", AdoRs.Fields(6))
                 ss2.Col = 28
                 ss2.Text = IIf(VarType(AdoRs.Fields(7)) = vbNull, "", AdoRs.Fields(7))
                 ss2.Col = 29
                 ss2.Text = IIf(VarType(AdoRs.Fields(8)) = vbNull, "", AdoRs.Fields(8))
                 ss2.Col = 30
                 ss2.Text = IIf(VarType(AdoRs.Fields(9)) = vbNull, "", AdoRs.Fields(9))
                 ss2.Col = 31
                 ss2.Text = IIf(VarType(AdoRs.Fields(10)) = vbNull, "", AdoRs.Fields(10))
                 
            End If
        End If
        
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing

End Sub

Private Function Sp_Process(Conn As ADODB.Connection, Sc As Collection) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim sSLAB_NO As String
    Dim ret_Result_ErrMsg As String
    Dim lRow, lRow2 As Integer
    Dim sTemp As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
    
    Screen.MousePointer = vbHourglass
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Sp_Process = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To 22
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    For lRow = 1 To ss1.MaxRows
    
        ss1.Row = lRow
        ss1.Col = 0
        
        If ss1.Text <> "" Then
        
            ss1.Col = 1
            sSLAB_NO = ss1.Text
    
            For lRow2 = 1 To ss3.MaxRows
            
                ss3.Row = lRow2
                
                ss3.Col = 1
                sTemp = ss3.Text
                ss3.Col = 3
                
                If sTemp <> "" And ss3.Text <> "" Then
                
                    adoCmd.Parameters(0).Value = "U"
                    adoCmd.Parameters(1).Value = sSLAB_NO         'SLAB_NO
                    
                    ss3.Col = 1
                    adoCmd.Parameters(2).Value = ss3.Text         'REASON_CD
                    adoCmd.Parameters(3).Value = ""               'REASON_COMMENT
                    adoCmd.Parameters(4).Value = ""               'NISCO_CHEM_GRD
                    adoCmd.Parameters(5).Value = ""               'CUST_CHEM_GRD
                    adoCmd.Parameters(6).Value = ""               'STLGRD
                    adoCmd.Parameters(7).Value = ""               'INS_EMP
                    ss3.Col = 3
                    adoCmd.Parameters(8).Value = ss3.Text         'EST_CD
                    ss3.Col = 5
                    adoCmd.Parameters(9).Value = ss3.Text         'EST_COMMENT
                    adoCmd.Parameters(10).Value = ""              'EST_STLGRD
                    ss3.Col = 6
                    adoCmd.Parameters(11).Value = ss3.Text        'EST_QUALITY_GRD
                    ss3.Col = 8
                    adoCmd.Parameters(12).Value = ss3.Text        'EST_EMP
                    adoCmd.Parameters(13).Value = ""              'THK
                    adoCmd.Parameters(14).Value = ""              'WID
                    adoCmd.Parameters(15).Value = ""              'LEN
                    adoCmd.Parameters(16).Value = ""              'WGT
                    adoCmd.Parameters(17).Value = ""              'PROC_CD
                    adoCmd.Parameters(18).Value = ""              'PROD_CD
                    adoCmd.Parameters(19).Value = ""              'ORD_FL
                    adoCmd.Parameters(20).Value = ""              'ORD_NO
                    adoCmd.Parameters(21).Value = ""              'ORD_ITEM
                    adoCmd.Parameters(22).Value = ""              'INS_PGMID
                    
                    adoCmd.Execute
                        
                    'Error Check
                    If adoCmd("Error") <> "0" Then
                    
                        ret_Result_ErrCode = adoCmd("Error")
                        ret_Result_ErrMsg = adoCmd("Messg")
                
                        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                        
                        Call Gp_Sp_RowColor(ss3, lRow2, , vbYellow)
                        Call Gp_MsgBoxDisplay(sErrMessg)
                        
                        Screen.MousePointer = vbDefault
                        Set adoCmd = Nothing
                        
                        Conn.RollbackTrans
                        Sp_Process = False
                        Exit Function
                
                    End If
                
                End If
                
            Next lRow2
            
        End If
        
    Next lRow
    
    Conn.CommitTrans
    
    Sc.Item("Spread").ReDraw = True
    
    If iProcessCount > 0 Then
        
        MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了" & iProcessCount & "条记录"
        
    End If
            
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

Private Sub txt_title_est_cd1_DblClick()

    Call txt_title_est_cd1_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_title_est_cd1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0018"
        DD.rControl.Add Item:=txt_title_est_cd1
        DD.rControl.Add Item:=txt_title_est_nm1
        
        DD.nameType = "2"
'        DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_title_est_cd1)) = txt_title_est_cd1.MaxLength Then
            txt_title_est_nm1.Text = Gf_ComnNameFind(M_CN1, "C0018", Trim(txt_title_est_cd1.Text), 2)
            
        Else
            txt_title_est_nm1.Text = ""
        End If
        
    End If
    
'    If txt_title_est_cd1.Text = "9090" Then
'        txt_title_est_cd1.Text = ""
'        txt_title_est_nm1.Text = ""
'    End If
    
End Sub

Private Sub txt_title_reason_cd1_DblClick()

    Call txt_title_reason_cd1_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_title_reason_cd1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0017"
        DD.rControl.Add Item:=txt_title_reason_cd1
        DD.rControl.Add Item:=txt_title_reason_nm1
        
        DD.nameType = "2"
'        DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_title_reason_cd1)) = txt_title_reason_cd1.MaxLength Then
            txt_title_reason_nm1.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(txt_title_reason_cd1.Text), 2)
        Else
            txt_title_reason_nm1.Text = ""
        End If
        
    End If
    
'    If txt_title_reason_cd1.Text = "9090" Then
'        txt_title_reason_cd1.Text = ""
'        txt_title_reason_nm1.Text = ""
'    End If

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
    
        If SSTab1.Tab = 0 Then
            .Buttons(7).Enabled = False                  'Row Insert
            .Buttons(8).Enabled = False                  'Row Delete
        Else
            .Buttons(7).Enabled = True                   'Row Insert
        End If
    End With

End Sub

