VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CGC2021C 
   Caption         =   "������Ϣ���ͽ���_CGC2021C"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12435
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSP6 
      Height          =   315
      Left            =   7590
      TabIndex        =   51
      Top             =   4050
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "���ڶ���"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP5 
      Height          =   315
      Left            =   6000
      TabIndex        =   50
      Top             =   4050
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   255
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��������"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.TextBox TXT_STLGRD 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15330
      MaxLength       =   80
      TabIndex        =   48
      Tag             =   "���Ϻ�"
      Top             =   1380
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TXT_WGT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4740
      MaxLength       =   14
      TabIndex        =   47
      Tag             =   "���Ϻ�"
      Top             =   1380
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox TXT_CUT_NO 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   16575
      MaxLength       =   10
      TabIndex        =   42
      Tag             =   "��������"
      Top             =   2190
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6840
      Top             =   60
   End
   Begin VB.TextBox TXT_SLAB_NO 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   16590
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "��������"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1980
   End
   Begin SSSplitter.SSSplitter SSSp1 
      Height          =   8655
      Left            =   90
      TabIndex        =   1
      Top             =   510
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   15266
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "CGC2021C.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   1305
         Left            =   0
         TabIndex        =   2
         Tag             =   "172.18.151.145"
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   2302
         _Version        =   196609
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox TXT_PAINT_POS 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7290
            MaxLength       =   14
            TabIndex        =   38
            Tag             =   "���Ϻ�"
            Top             =   870
            Width           =   975
         End
         Begin VB.TextBox TXT_ROLL_TEMP 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   13890
            MaxLength       =   14
            TabIndex        =   37
            Tag             =   "���Ϻ�"
            Text            =   "600"
            Top             =   870
            Width           =   855
         End
         Begin VB.TextBox TXT_PAINT_CNT 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10740
            MaxLength       =   14
            TabIndex        =   36
            Tag             =   "���Ϻ�"
            Top             =   870
            Width           =   615
         End
         Begin VB.TextBox TXT_THK 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1830
            MaxLength       =   14
            TabIndex        =   35
            Tag             =   "���Ϻ�"
            Top             =   870
            Width           =   765
         End
         Begin VB.TextBox TXT_LEN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3510
            MaxLength       =   14
            TabIndex        =   34
            Tag             =   "���Ϻ�"
            Top             =   870
            Width           =   1125
         End
         Begin VB.TextBox TXT_WID 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2610
            MaxLength       =   14
            TabIndex        =   33
            Tag             =   "���Ϻ�"
            Top             =   870
            Width           =   885
         End
         Begin VB.TextBox Winsock 
            Height          =   360
            Left            =   150
            TabIndex        =   3
            Tag             =   "������"
            Top             =   3060
            Width           =   14835
         End
         Begin Threed.SSFrame SSFrame7 
            Height          =   705
            Left            =   12150
            TabIndex        =   14
            Top             =   90
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1244
            _Version        =   196609
            BackColor       =   12632319
            Begin VB.TextBox TXT_BOT 
               Alignment       =   2  'Center
               CausesValidation=   0   'False
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2310
               MaxLength       =   10
               TabIndex        =   32
               Tag             =   "��������"
               Text            =   "0"
               Top             =   360
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.TextBox TXT_TOP 
               Alignment       =   2  'Center
               CausesValidation=   0   'False
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2310
               MaxLength       =   10
               TabIndex        =   31
               Tag             =   "��������"
               Text            =   "1"
               Top             =   30
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Bottom Selected"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   390
               TabIndex        =   16
               Top             =   390
               Width           =   1890
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Top Selected"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   390
               TabIndex        =   15
               Top             =   60
               Value           =   1  'Checked
               Width           =   1890
            End
         End
         Begin Threed.SSFrame SSFrame6 
            Height          =   705
            Left            =   90
            TabIndex        =   17
            Top             =   90
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1244
            _Version        =   196609
            BackColor       =   12632319
            Begin VB.TextBox TXT_RL 
               Alignment       =   2  'Center
               CausesValidation=   0   'False
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   4830
               MaxLength       =   10
               TabIndex        =   29
               Tag             =   "��������"
               Text            =   "1"
               Top             =   240
               Visible         =   0   'False
               Width           =   510
            End
            Begin Threed.SSOption opt_line1 
               Height          =   285
               Left            =   360
               TabIndex        =   18
               Top             =   60
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " 90  (Right-hand) Orientation of marking"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line2 
               Height          =   285
               Left            =   360
               TabIndex        =   19
               Top             =   390
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " 270 (Left-hand) Orientation of marking"
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   705
            Left            =   5550
            TabIndex        =   20
            Top             =   90
            Width           =   6555
            _ExtentX        =   11562
            _ExtentY        =   1244
            _Version        =   196609
            BackColor       =   12632319
            Begin VB.TextBox TXT_COM 
               Alignment       =   2  'Center
               CausesValidation=   0   'False
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   600
               MaxLength       =   10
               TabIndex        =   30
               Tag             =   "��������"
               Text            =   "5"
               Top             =   330
               Visible         =   0   'False
               Width           =   510
            End
            Begin Threed.SSOption opt_line3 
               Height          =   255
               Left            =   2370
               TabIndex        =   21
               Top             =   60
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "75%"
            End
            Begin Threed.SSOption opt_line4 
               Height          =   255
               Left            =   2370
               TabIndex        =   22
               Top             =   390
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "84%"
            End
            Begin Threed.SSOption opt_line5 
               Height          =   255
               Left            =   3345
               TabIndex        =   23
               Top             =   60
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "94%"
            End
            Begin Threed.SSOption opt_line6 
               Height          =   255
               Left            =   3345
               TabIndex        =   24
               Top             =   390
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "116%"
            End
            Begin Threed.SSOption opt_line7 
               Height          =   255
               Left            =   4335
               TabIndex        =   25
               Top             =   60
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "130%"
            End
            Begin Threed.SSOption opt_line8 
               Height          =   255
               Left            =   4335
               TabIndex        =   26
               Top             =   390
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "150%"
            End
            Begin Threed.SSOption opt_line10 
               Height          =   255
               Left            =   300
               TabIndex        =   27
               Top             =   60
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "No compression"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line9 
               Height          =   255
               Left            =   5310
               TabIndex        =   28
               Top             =   390
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "200%"
            End
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   90
            Top             =   870
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "�� / �� / ��"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   12150
            Top             =   870
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "�����¶�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   5550
            Top             =   870
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "��ӡ��ʼλ��"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   9000
            Top             =   870
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "��ӡ����"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   2160
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1335
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   3810
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         MaxRows         =   10
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CGC2021C.frx":0092
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4755
         Left            =   0
         TabIndex        =   9
         Top             =   3900
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   8387
         _StockProps     =   64
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   30
         MaxRows         =   20
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGC2021C.frx":0892
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   0
         TabIndex        =   10
         Top             =   3525
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   609
         _Version        =   196609
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSP1 
            Height          =   315
            Left            =   10890
            TabIndex        =   11
            Top             =   15
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   16761087
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "����"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSP2 
            Height          =   315
            Left            =   12330
            TabIndex        =   12
            Top             =   15
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "ĸ��"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSP3 
            Height          =   315
            Left            =   13770
            TabIndex        =   13
            Top             =   15
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�ְ�"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   405
            Left            =   0
            TabIndex        =   39
            Top             =   -30
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   714
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_slab_no 
               Height          =   285
               Left            =   1920
               TabIndex        =   40
               Top             =   60
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "������"
            End
            Begin Threed.SSOption opt_cut_no 
               Height          =   285
               Left            =   450
               TabIndex        =   41
               Top             =   60
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "�ֶκ�"
               Value           =   -1
            End
         End
         Begin Threed.SSPanel SSP4 
            Height          =   315
            Left            =   9120
            TabIndex        =   49
            Top             =   15
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�ص㶩��"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   15375
      Top             =   1800
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   345
      Left            =   3090
      TabIndex        =   5
      Top             =   90
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   609
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label tcpMsg 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   6
         Top             =   60
         Width           =   2805
      End
      Begin VB.Shape tcpStatus 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   -30
         Shape           =   3  'Circle
         Top             =   60
         Width           =   435
      End
   End
   Begin Threed.SSCommand SSCmd_cnn 
      Height          =   345
      Left            =   780
      TabIndex        =   7
      Top             =   90
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   14804173
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��������״̬"
      ButtonStyle     =   3
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7320
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "172.18.128.101"
      RemotePort      =   2121
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   1
      Left            =   15360
      Top             =   2190
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "�ֶκ�"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   345
      Left            =   11850
      TabIndex        =   44
      Top             =   90
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   609
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Shape tcpStatus2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   -30
         Shape           =   3  'Circle
         Top             =   60
         Width           =   435
      End
      Begin VB.Label tcpMsg2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   45
         Top             =   60
         Width           =   2805
      End
   End
   Begin Threed.SSCommand SSCmd_cnn2 
      Height          =   345
      Left            =   9540
      TabIndex        =   46
      Top             =   90
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   14804173
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����������״̬"
      ButtonStyle     =   3
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7830
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   25298
   End
   Begin VB.CheckBox chk_Cond 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   390
      TabIndex        =   4
      Top             =   90
      Width           =   2400
   End
   Begin VB.CheckBox chk_Cond 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   9150
      TabIndex        =   43
      Top             =   90
      Width           =   2400
   End
End
Attribute VB_Name = "CGC2021C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      LABEL PRINTER SEND DATA
'-- Program ID        CGC2080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2008.3.24
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_LINE1 = 1
Const SS1_SLAB_NO = 2
Const SS1_LOT_NO = 3
Const SS1_ROLL_THK = 4
Const SS1_ROLL_WID = 5
Const SS1_ROLL_LEN = 6
Const SS1_SIZE_KND = 7
Const SS1_CUT_NO = 9
Const SS1_DATE = 10
Const SS1_SHIFT = 11
Const SS1_ROLL_WGT = 12
Const SS1_STLGRD = 13

Const SS2_BLOCK_SEQ = 2
Const SS2_SEQ = 3
Const SS2_PROD_CD = 4
Const SS2_THK = 5
Const SS2_WID = 6
Const SS2_LEN = 7
Const SS2_ORD_NO = 9
Const SS2_UST_FL = 13
Const SS2_HTM = 14
Const SS2_POS = 15
Const SS2_PAINT_NO = 16
Const SS2_PAINT_ADD = 17
Const SS2_DEL_TO_DATE = 19
Const SS2_URGNT_FL = 20
Const SS2_IMP_CONT = 21
Const SS2_JIT_FLAG = 22
Const SS2_ERPORT = 23
Const SS2_ORD_CNT = 24 'һ���ඩ��
Const SS2_OVER_FL = 25 '�쳣��
Const SS2_TRIM_FL = 26 '�Ƿ��б�
Const SS2_FLAB_NO = 1
Const SS2_SURFACE_REQUESTS = 30

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSrc As Any, ByVal cbLength As Long)
Public Property Get HLByte(ByRef Word As Long, HL As Long) As Byte
CopyMemory HLByte, ByVal VarPtr(Word) + HL, 1
End Property
Public Property Get LoByte(ByRef Word As Integer) As Byte
CopyMemory LoByte, ByVal VarPtr(Word), 1
End Property

Public Property Let LoByte(ByRef Word As Integer, ByVal LowByte As Byte)
CopyMemory Word, LowByte, 1
End Property

Public Property Get HiByte(ByRef Word As Integer) As Byte
CopyMemory HiByte, ByVal VarPtr(Word) + 1, 1
End Property

Public Property Let HiByte(ByRef Word As Integer, ByVal HighByte As Byte)
CopyMemory ByVal VarPtr(Word) + 1, HighByte, 1
End Property

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
       
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_CUT_NO, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGC2021C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) ' Add by liqian at 2012-07-30 ��������
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) 'һ���ඩ��
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�쳣��
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�Ƿ��б�
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�Ƿ��б�
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�Ƿ��б�
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�Ƿ��б�
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�Ƿ��б�
  
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGC2021C.P_SREFER2", Key:="P-R"
    sc2.Add Item:="CGC2021C.P_MODIFY", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
'    Call Gp_Sp_ColHidden(ss1, 18, True)

    Call Gp_Sp_ColHidden(ss2, 26, True)
    
End Sub

Private Sub CmdSEND_Click()

End Sub

Private Sub chk_Cond_Click(Index As Integer)

    Dim strState As String
    Dim strState2 As String

    If Index = 1 Then
        If chk_Cond(Index) Then
           TXT_TOP = "1"
        Else
           TXT_TOP = "0"
        End If
    End If
    If Index = 2 Then
        If chk_Cond(Index) Then
           TXT_BOT = "1"
        Else
           TXT_BOT = "0"
        End If
    End If
    If Index = 0 Then
       If chk_Cond(Index) = 1 Then
          Winsock1.Connect
       Else
          Winsock1.Close
          strState = "���Ӷ���"
          tcpStatus.BackColor = &HFF&
          chk_Cond(0).ForeColor = &HFF&
          tcpMsg.Caption = "����״̬ : " & strState
       End If
    End If
    If Index = 3 Then
       If chk_Cond(Index) = 1 Then
          Winsock2.Connect
       Else
          Winsock2.Close
          strState2 = "���Ӷ���"
          tcpStatus2.BackColor = &HFF&
          chk_Cond(Index).ForeColor = &HFF&
          tcpMsg2.Caption = "������״̬ : " & strState2
       End If
    End If
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
'    Call Gp_Sp_ColHidden(ss2, SS2_PAINT_ADD, True)
'    Call Gp_Sp_ColHidden(ss1, SS1_STLGRD, True)

'    Winsock1.RemoteHost = "172.18.56.254" 'Gf_ComnNameFind(M_CN1, "G0034", "03", 1)
'    Winsock1.RemotePort = "5151" 'Gf_ComnNameFind(M_CN1, "G0034", "03", 2)
'    Winsock2.RemoteHost = "172.18.128.167" 'Gf_ComnNameFind(M_CN1, "G0034", "03", 1)
'    Winsock2.RemotePort = "20020" 'Gf_ComnNameFind(M_CN1, "G0034", "03", 2)
    Winsock1.RemoteHost = Gf_ComnNameFind(M_CN1, "G0034", "03", 1)
    Winsock1.RemotePort = Gf_ComnNameFind(M_CN1, "G0034", "03", 2)
    Winsock2.RemoteHost = Gf_ComnNameFind(M_CN1, "G0047", "01", 1)
    Winsock2.RemotePort = Gf_ComnNameFind(M_CN1, "G0047", "01", 2)
'
    TXT_RL.Text = "1"
    TXT_COM.Text = "5"
    TXT_TOP.Text = "1"
    TXT_BOT.Text = "0"    'LiQian 2012-08-23 �ײ�ѡ��Ĭ�ϲ�ѡ,������TXT_BOTĬ��Ҳ��Ϊ0�Ͳ�ѡ��
'
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Winsock1.Close

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
'    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
'        Cancel = 1
'        Exit Sub
'    End If
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
     
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
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) And Gf_Sp_Cls(sc2) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        
    End If
    
End Sub

Public Sub Form_Exc()
    
'    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
    
    Dim iCount       As Integer
    Dim sPlateNo     As String
  
    Dim inum         As Integer
    Dim lRow         As Integer
    Dim iRow         As Integer
    Dim iCol         As Integer
    Dim simpcont     As String
    Dim sFlag        As String
    Dim sexport      As String

            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
        Call Gf_Sp_Cls(sc2)
    End If
    
    With ss2
        For iRow = 1 To .MaxRows
           .ROW = iRow:
           .Col = SS2_IMP_CONT:    simpcont = Trim(.Text)
          
            If simpcont = "Y" Then
               Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, SSP4.BackColor)
            End If
            
            '�Ƿ�������
            .ROW = iRow:
            .Col = SS2_JIT_FLAG: sFlag = Trim(.Text)
            If sFlag = "Y" Then
               Call Gp_Sp_BlockColor(ss2, SS2_FLAB_NO, SS2_FLAB_NO, iRow, iRow, SSP5.BackColor)
            End If
            '�Ƿ���ڶ���
            .ROW = iRow:
            .Col = SS2_ERPORT: sexport = Trim(.Text)
            If sexport = "Y" Then
               Call Gp_Sp_BlockColor(ss2, SS2_FLAB_NO, SS2_FLAB_NO, iRow, iRow, SSP6.BackColor)
            End If
        Next iRow
     End With

End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    
    Dim sMark_no As String
    Dim sPlate_no As String
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sSpec As String
    Dim sStdspec_YY As String
    
    If Trim(TXT_SLAB_NO.Text) = "" Then
       MsgBox (" ������Ϊ�գ���ȷ�� ")
       Exit Sub
    End If
    
    If Val(TXT_THK.Text) < 5 Or Val(TXT_THK.Text) > 150 Or _
       Val(TXT_WID.Text) < 1500 Or Val(TXT_WID.Text) > 2650 Or _
       Val(TXT_LEN.Text) < 1700 Or Val(TXT_LEN.Text) > 50000 Then
       MsgBox (" ������񳬳������豸������Χ����ȷ�� ")
       Exit Sub
    End If
    
    If Val(TXT_ROLL_TEMP.Text) < 250 Or Val(TXT_ROLL_TEMP.Text) > 1000 Then
       MsgBox (" �����¶ȳ��������豸������Χ����ȷ�� ")
       Exit Sub
    End If
    
    If Val(TXT_PAINT_CNT.Text) < 1 Or Val(TXT_PAINT_CNT.Text) > 10 Then
       MsgBox (" ��ӡ�������������豸������Χ����ȷ�� ")
       Exit Sub
    End If
    
    If Val(TXT_PAINT_POS.Text) < 1700 Or Val(TXT_PAINT_POS.Text) > 49300 Then
       MsgBox (" ��ӡ��ʼλ�ó��������豸������Χ����ȷ�� ")
       Exit Sub
    End If

    If chk_Cond(0) = 1 Then
       Call Cmd_SEND
       
       If Gf_Sp_Process(M_CN1, Proc_Sc("SC2"), Mc1, False) Then
          Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
       End If
        
    End If
        
    If chk_Cond(3) = 1 Then
       Call Cmd_SEND_Surf
    End If
    
    Call Form_Ref
    
End Sub
Public Sub Form_Ins()

End Sub

Public Sub Spread_Forzens_Setting()
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
End Sub

Public Sub Spread_Forzens_Cancel()
    Me.ActiveControl.ColsFrozen = 0
End Sub

Public Sub Spread_Del()

End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_cut_no_Click(Value As Integer)
    Call Form_Ref
End Sub

Private Sub opt_line1_Click(Value As Integer)
    If opt_line1 = True Then
       TXT_RL.Text = "1"
    End If
End Sub

Private Sub opt_line10_Click(Value As Integer)

    If opt_line10 = True Then
       TXT_COM.Text = "5"
    End If
    
End Sub

Private Sub opt_line2_Click(Value As Integer)
    If opt_line2 = True Then
       TXT_RL.Text = "2"
    End If
End Sub

Private Sub opt_line3_Click(Value As Integer)
    If opt_line3 = True Then
       TXT_COM.Text = "2"
    End If
End Sub
Private Sub opt_line4_Click(Value As Integer)
    If opt_line4 = True Then
       TXT_COM.Text = "3"
    End If
End Sub
Private Sub opt_line5_Click(Value As Integer)
    If opt_line5 = True Then
       TXT_COM.Text = "4"
    End If
End Sub
Private Sub opt_line6_Click(Value As Integer)
    If opt_line6 = True Then
       TXT_COM.Text = "6"
    End If
End Sub
Private Sub opt_line7_Click(Value As Integer)
    If opt_line7 = True Then
       TXT_COM.Text = "7"
    End If
End Sub
Private Sub opt_line8_Click(Value As Integer)
    If opt_line8 = True Then
       TXT_COM.Text = "8"
    End If
End Sub
Private Sub opt_line9_Click(Value As Integer)
    If opt_line9 = True Then
       TXT_COM.Text = "9"
    End If
End Sub

Private Sub opt_slab_no_Click(Value As Integer)
    Call Form_Ref
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal ROW As Long, ByVal ButtonDown As Integer)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    Dim lRow As Long
    Dim sBlockSeq As String
    Dim sSeq As String
    
    Dim iCol As Long
    Dim iRow As Long
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    Dim iPaint_cnt As Integer
    Dim iPaint_Len As Long
    Dim sDate As String
    Dim sShift As String
    Dim sCut_no As String
    
    Dim sCurDate     As String
    Dim sOrderNo     As String
    Dim sDel_To_Date As String
    
    Dim sAdd_W       As String
    Dim sAdd_S       As String
    Dim sAdd_T       As String
    Dim sAdd_H       As String
    Dim sAdd_B       As String
    Dim sAdd_D       As String  '�������Ͷ�����ĩβ����ĸD  add by lichao 20140709
    Dim sAdd_E       As String  'һ���ඩ����ĩβ����ĸP  add by lichao 20150317
    Dim sAdd_F       As String  '�쳣����ĩβ����ĸX  add by lichao 20150317
    Dim sAdd_G       As String  '�Ƿ��б߼�M
    Dim iPaint_Add   As String
    
    Dim sSthk   As String
    Dim sSord   As String
    
    Dim sXm    As String
    
    sCurDate = Format(Now, "YYYYMM")
    
    iCol = Col
    iRow = ROW
    
    If ROW <= 0 Then Exit Sub
    If Col > 1 Then Exit Sub
    
    iRowto = iRow - 1
    iRowfr = iRow + 1
    
    If iRowto > 0 Then
        For iRowNum = 1 To iRowto
             
             ss1.Col = 0
             ss1.ROW = iRowNum
             If ss1.Text <> "" Then
                ss1.Text = ""
                ss1.Col = SS1_LINE1
                ss1.Value = 0
                Exit For
             End If
        Next iRowNum
    End If
    
    If iRowfr <= ss1.MaxRows Then
        For iRowNum = iRowfr To ss1.MaxRows
             
             ss1.Col = 0
             ss1.ROW = iRowNum
             If ss1.Text <> "" Then
                ss1.Text = ""
                ss1.Col = SS1_LINE1
                ss1.Value = 0
                Exit For
             End If
        Next iRowNum
    End If
    
    ss1.ROW = iRow
    If Col = SS1_LINE1 And ButtonDown = 1 Then
        ss1.Col = 0
        ss1.Text = "Update"
    Else
        ss1.Text = ""
    End If
    
    ss1.ROW = ROW
    ss1.Col = SS1_SLAB_NO:    TXT_SLAB_NO.Text = ss1.Text
    ss1.Col = SS1_CUT_NO:     TXT_CUT_NO.Text = ss1.Text
    sCut_no = Replace(TXT_CUT_NO.Text, "-", " ")
    ss1.Col = SS1_DATE:       sDate = ss1.Text
    ss1.Col = SS1_SHIFT:      sShift = ss1.Text
    ss1.Col = SS1_ROLL_THK:   TXT_THK.Text = ss1.Text
    ss1.Col = SS1_ROLL_WID:   TXT_WID.Text = ss1.Text
    ss1.Col = SS1_ROLL_LEN:   TXT_LEN.Text = ss1.Text
    ss1.Col = SS1_ROLL_WGT:   TXT_WGT.Text = ss1.Text
    ss1.Col = SS1_STLGRD:     TXT_STLGRD.Text = ss1.Text
    
    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows)
    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, , SSP1.BackColor)

    If ButtonDown = 1 And Col = 1 Then
       Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
    End If
    
    ss2.OperationMode = OperationModeNormal
    
    'TXT_SLAB_NO.Text = ""
    iPaint_cnt = 0
    iPaint_Len = 0
    
    For lRow = 1 To ss2.MaxRows
    
        ss2.ROW = lRow
        ss2.Col = SS2_BLOCK_SEQ: sBlockSeq = ss2.Text
        ss2.Col = SS2_SEQ:       sSeq = ss2.Text
        
        If sBlockSeq & sSeq = "0000" Then
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ss2.ROW, ss2.ROW, , SSP1.BackColor)
            ss2.Col = SS2_PROD_CD:       ss2.Text = "����"
        ElseIf sSeq = "00" Then
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ss2.ROW, ss2.ROW, , SSP2.BackColor)
            ss2.Col = SS2_PROD_CD:       ss2.Text = "ĸ��" & sBlockSeq
        Else
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ss2.ROW, ss2.ROW, , SSP3.BackColor)
            ss2.Col = SS2_PROD_CD: ss2.Text = "�ְ�"
            ss2.Col = 0: ss2.Text = "Update"
            
            iPaint_cnt = iPaint_cnt + 1
            If iPaint_cnt = 1 Then
               ss2.Col = SS2_LEN:   TXT_PAINT_POS.Text = Round(Val(ss2.Text) / 3)
               If Val(TXT_PAINT_POS.Text) < 1700 Then
                  TXT_PAINT_POS.Text = 1700
               End If
            End If
            
            ss2.Col = SS2_POS:           ss2.Text = TXT_PAINT_POS.Text + iPaint_Len
            ss2.Col = SS2_LEN:           iPaint_Len = iPaint_Len + Val(ss2.Text)
            
            '������ȡ�24mm������ǰ��������S��̽�˼���T���ȴ�������N �����𣩻�Q����� �����¼���������磩
            '������ȣ�24mmֻ���統����ǰ����S��̽�ˡ��ȴ�������
            '������������λ�ã���9-11λ��˳��ֱ�ΪS T N��Q ���磨21B002-1STN��Q��
            '��û����������ָʾ���9-11λ�ֱ���
            
            '20110514 �޸�
            '������OB5/OM8��ͷ����W,������ǰ��������S��̽�˼���T���ȴ�������N �����𣩻�Q����� �����¼���������磩
            
            '20110519 �޸�
            '������OB5/OM8��ͷ����W,������ǰ��������S��̽�˼���T���ȴ�������N �����𣩻�Q����� �����¼���������磩
            '���磺��21B002-1WSTN��Q��
            ss2.Col = SS2_ORD_NO:   sOrderNo = Mid(ss2.Text, 1, 3)
            If sOrderNo = "OB5" Or sOrderNo = "OM8" Then
                sAdd_W = "C"
            End If
                                            
            ss2.Col = SS2_DEL_TO_DATE:   sDel_To_Date = Mid(ss2.Value, 1, 6)
            If sDel_To_Date < sCurDate Then
                sAdd_S = "S"
            End If
            
            ss2.Col = SS2_UST_FL
            If ss2.Text = "̽��" Then
                sAdd_T = "T"
            End If
            ' Add by liqian at 2012-07-30 ��������
            ss2.Col = SS2_URGNT_FL
            If ss2.Text = "Y" Then
                sAdd_B = "B"
            Else
                sAdd_B = ""
            End If
            ' Add by lichao at 2014-07-09 �������Ͷ���
            ss2.Col = SS2_JIT_FLAG
            If ss2.Text = "Y" Then
                sAdd_D = "D"
            Else
                sAdd_D = ""
            End If
            
            ' Add by lichao at 2015-03-17 ��������
            ss2.Col = SS2_ORD_CNT:   sSord = ss2.Text
            ss2.Col = SS2_THK:       sSthk = Trim(Str(ss2.Text))
            If sSord > "1" And sSthk < "25" Then
                sAdd_E = "P"
            Else
                sAdd_E = ""
            End If
            
            ' Add by lichao at 2015-03-17 �쳣��
            ss2.Col = SS2_OVER_FL
            If ss2.Text <> "" And ss2.Text <> "0" Then
                sAdd_F = "X"
            Else
                sAdd_F = ""
            End If
                        
            'ss2.Col = SS2_JIT_FLAG:      sAdd_D = ss2.Text
            ss2.Col = SS2_HTM
            If ss2.Text <> "" Then
               sAdd_H = "N"
            Else
               sAdd_H = ""
            End If
            
            ss2.Col = SS2_SURFACE_REQUESTS: sXm = ss2.Text
            
            ss2.Col = SS2_TRIM_FL:       sAdd_G = ss2.Text
            
            ss2.Col = SS2_PAINT_ADD:     ss2.Text = sAdd_W & sAdd_S & sAdd_T & sAdd_H & sAdd_B & sAdd_D & sAdd_F & sAdd_G & sAdd_E & sXm
            
            ss2.Col = SS2_PAINT_ADD:     iPaint_Add = ss2.Text
            
            If opt_cut_no.Value = True Then
               ss2.Col = SS2_PAINT_NO:   ss2.Text = sDate & sShift & sCut_no & "-" & iPaint_cnt & iPaint_Add
            Else
               ss2.Col = SS2_PAINT_NO:   ss2.Text = TXT_SLAB_NO.Text & "-" & iPaint_cnt & iPaint_Add
            End If
            ss2.Text = Trim(ss2.Text)
            
        End If
    
    Next lRow
    
    TXT_PAINT_CNT.Text = iPaint_cnt
    
End Sub

Private Sub SSCmd_cnn_Click()
'    Winsock1.Close
'    Winsock1.Connect
End Sub

Private Sub Cmd_SEND()

    Dim sMesg As String
    
    Dim Header As String * 2
    Dim sPlate_no As String * 18
    Dim iThk As Long
    Dim iWid As Long
    Dim iLen As Long
    Dim iTemp As Integer
    Dim iRL As Integer
    Dim iCOM As Integer
    Dim iPaint_cnt As Integer
    Dim iTop As Integer
    Dim iBot As Integer
        
    Dim sSeq As String
    Dim sPaint_Pos(9) As Long
    Dim sPaint_Str(9) As String * 16
    Dim StrSend(1) As String
    Dim sCtrl As String * 2
    
    Dim lRow As Integer
        
    Header = "MD"
    sPlate_no = Trim(TXT_SLAB_NO.Text)
    iThk = Val(TXT_THK.Text)
    iWid = Val(TXT_WID.Text)
    iLen = Val(TXT_LEN.Text)
    iTemp = Val(TXT_ROLL_TEMP.Text)
    iRL = Val(TXT_RL.Text)
    iCOM = Val(TXT_COM.Text)
    iPaint_cnt = Val(TXT_PAINT_CNT.Text)
    iTop = Val(TXT_TOP.Text)
    iBot = Val(TXT_BOT.Text)
    
    StrSend(0) = Chr(16)
    StrSend(1) = Chr(16)
    sCtrl = StrSend(0) & StrSend(1)
    
    iPaint_cnt = 0
    For lRow = 1 To ss2.MaxRows
    
        ss2.ROW = lRow
        ss2.Col = SS2_SEQ:            sSeq = ss2.Text
        If sSeq <> "00" Then
            ss2.Col = SS2_POS:        sPaint_Pos(iPaint_cnt) = Val(ss2.Text)
            ss2.Col = SS2_PAINT_NO:   sPaint_Str(iPaint_cnt) = ss2.Text
            iPaint_cnt = iPaint_cnt + 1
        End If
        
    Next lRow
            
    Winsock1.SendData Header & "  "
    Winsock1.SendData Chr(18) & Chr(10) & sPlate_no
    Winsock1.SendData HLByte(iLen, 3)
    Winsock1.SendData HLByte(iLen, 2)
    Winsock1.SendData HLByte(iLen, 1)
    Winsock1.SendData HLByte(iLen, 0)
    Winsock1.SendData HLByte(iWid, 3)
    Winsock1.SendData HLByte(iWid, 2)
    Winsock1.SendData HLByte(iWid, 1)
    Winsock1.SendData HLByte(iWid, 0)
    Winsock1.SendData HLByte(iThk, 3)
    Winsock1.SendData HLByte(iThk, 2)
    Winsock1.SendData HLByte(iThk, 1)
    Winsock1.SendData HLByte(iThk, 0)
    Winsock1.SendData HiByte(iTemp)
    Winsock1.SendData LoByte(iTemp)
    Winsock1.SendData HiByte(iRL)
    Winsock1.SendData LoByte(iRL)
    Winsock1.SendData HiByte(iCOM)
    Winsock1.SendData LoByte(iCOM)
    Winsock1.SendData HiByte(iPaint_cnt)
    Winsock1.SendData LoByte(iPaint_cnt)
    Winsock1.SendData HiByte(iTop)
    Winsock1.SendData LoByte(iTop)
    Winsock1.SendData HiByte(iBot)
    Winsock1.SendData LoByte(iBot)
    
    For lRow = 1 To 10
    
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 3)
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 2)
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 1)
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 0)
        Winsock1.SendData sCtrl & sPaint_Str(lRow - 1)
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 3)
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 2)
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 1)
        Winsock1.SendData HLByte(sPaint_Pos(lRow - 1), 0)
        Winsock1.SendData sCtrl & sPaint_Str(lRow - 1)
    
    Next lRow

End Sub

Private Sub Cmd_SEND_Surf()

    Dim sMesg As String
    
    Dim Header As String * 12
    Dim sPlate_no As String * 20
    Dim sWgt As String * 5
    Dim sWid As String * 5
    Dim sThk As String * 5
    Dim sLen As String * 10
    Dim sStlgrd As String * 30

    Dim iWid As Long
    Dim iLen As Long
    Dim iThk As Long
    Dim iWGT As Long
        
    iWGT = Val(TXT_WGT.Text) * 1000
    iWid = Val(TXT_WID.Text)
    iThk = Val(TXT_THK.Text) * 1000
    iLen = Val(TXT_LEN.Text)
    
    Header = "GPDIMCD 0077"
    sPlate_no = Trim(TXT_SLAB_NO.Text)
    sWgt = Format(Str(iWGT), "00000")
    sWid = Format(Str(iWid), "00000")
    sThk = Format(Str(iThk), "00000")
    sLen = Format(Str(iLen), "0000000000")
    sStlgrd = TXT_STLGRD.Text
            
    Winsock2.SendData Header & sPlate_no & sWgt & sWid & sThk & sLen & sStlgrd


End Sub



Private Sub Timer1_Timer()

    'sckClosed            0 ȱʡ�ġ�--�ر� û�е�
    'sckOpen              1 �� --�򿪵�
    'sckListening         2 ���� --�쿴��û����������
    'sckConnectionPending 3 ���ӹ���
    'sckResolvingHost     4 ʶ������
    'sckHostResolved      5 ��ʶ������
    'sckConnecting        6 ��������
    'sckConnected         7 ������
    'sckClosing           8 ͬ����Ա���ڹر����� -˵���Է��ر���������
    'sckError             9 ����
    
    Dim strState As String
    Dim strState2 As String
    
    If chk_Cond(0) <> 1 And chk_Cond(3) <> 1 Then
       Exit Sub
    Else
    
        If chk_Cond(0) = 1 Then
        
            Select Case Winsock1.State
                Case 0
                    strState = "���ӹر�"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 1
                    strState = "���Ӵ�"
                Case 2
                    strState = "���ӱ���"
                Case 3
                    strState = "Close"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 4
                    strState = "Find Host...."
                Case 5
                    strState = "Finded Host"
                Case 6
                    strState = "��������"
                Case 7
                    strState = "��������"
                    tcpStatus.BackColor = &HC000&
                    chk_Cond(0).ForeColor = &HC000&
                Case 8
                    strState = "���Ӷ���"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 9
                    strState = "���Ӵ���"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
            Case Else
                strState = "StateNum:" & Winsock1.State
                tcpStatus.BackColor = &HFF&
                chk_Cond(0).ForeColor = &HFF&
            End Select

            tcpMsg.Caption = "����״̬ : " & strState
            
        End If
        
        If chk_Cond(3) = 1 Then

            Select Case Winsock2.State
                Case 0
                    strState2 = "���ӹر�"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
                Case 1
                    strState2 = "���Ӵ�"
                Case 2
                    strState2 = "���ӱ���"
                Case 3
                    strState2 = "Close"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
                Case 4
                    strState2 = "Find Host...."
                Case 5
                    strState2 = "�ҵ�����"
                Case 6
                    strState2 = "��������"
                Case 7
                    strState2 = "��������"
                    tcpStatus2.BackColor = &HC000&
                    chk_Cond(3).ForeColor = &HC000&
                Case 8
                    strState2 = "���Ӷ���"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
                Case 9
                    strState2 = "���Ӵ���"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
            Case Else
                strState2 = "StateNum:" & Winsock2.State
                tcpStatus2.BackColor = &HFF&
                chk_Cond(3).ForeColor = &HFF&
            End Select

            tcpMsg2.Caption = "������״̬ : " & strState2

        End If
        
    End If
    
End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
'        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

