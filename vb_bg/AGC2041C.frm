VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2041C 
   Caption         =   "̽��ʵ����ѯ����_AGC2041C"
   ClientHeight    =   10680
   ClientLeft      =   15
   ClientTop       =   1740
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   16325
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "AGC2041C.frx":0000
      Begin Threed.SSFrame Single 
         Height          =   1650
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15090
         _ExtentX        =   26617
         _ExtentY        =   2910
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
         Begin VB.TextBox TXT_ORD_NO 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1395
            MaxLength       =   11
            TabIndex        =   29
            Tag             =   "CD_MANA_NO"
            Top             =   1230
            Width           =   1530
         End
         Begin VB.ComboBox CBO_ORD_ITEM 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2955
            TabIndex        =   28
            Top             =   1230
            Width           =   750
         End
         Begin VB.TextBox TXT_CO_CD 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   11520
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "��׼����"
            Top             =   1230
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox TXT_STDSPEC_CD 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   12720
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "��׼����"
            Top             =   840
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.ComboBox CBO_SHIFT 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "AGC2041C.frx":0052
            Left            =   5820
            List            =   "AGC2041C.frx":0062
            TabIndex        =   17
            Top             =   120
            Width           =   930
         End
         Begin VB.ComboBox CBO_UST_DEC 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "AGC2041C.frx":0071
            Left            =   5820
            List            =   "AGC2041C.frx":007E
            TabIndex        =   16
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox TXT_UST_STAND_NO 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1395
            TabIndex        =   15
            Tag             =   "����׼"
            Top             =   480
            Width           =   630
         End
         Begin VB.TextBox TXT_UST_STAND_NAME 
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
            Height          =   315
            Left            =   2025
            TabIndex        =   14
            Top             =   480
            Width           =   2565
         End
         Begin VB.TextBox TXT_ADDR 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   8460
            MaxLength       =   10
            TabIndex        =   13
            Top             =   2670
            Width           =   1005
         End
         Begin VB.TextBox TXT_STDSPEC 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8220
            TabIndex        =   12
            Tag             =   "��׼����"
            Top             =   120
            Width           =   2205
         End
         Begin VB.TextBox TXT_ADDR 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   9480
            MaxLength       =   10
            TabIndex        =   11
            Top             =   2670
            Width           =   585
         End
         Begin VB.TextBox TXT_ADDR 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   10350
            MaxLength       =   10
            TabIndex        =   10
            Top             =   2670
            Width           =   585
         End
         Begin VB.TextBox TXT_MAT_NO 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8220
            MaxLength       =   14
            TabIndex        =   9
            Tag             =   "��׼����"
            Top             =   480
            Width           =   2205
         End
         Begin VB.CheckBox chk_Cond_W 
            BackColor       =   &H00E0E0E0&
            Caption         =   "����"
            Height          =   255
            Left            =   14040
            TabIndex        =   8
            Tag             =   "W"
            Top             =   480
            Width           =   720
         End
         Begin VB.CheckBox chk_Cond_B 
            BackColor       =   &H00E0E0E0&
            Caption         =   "����"
            Height          =   255
            Left            =   14040
            TabIndex        =   7
            Tag             =   "B"
            Top             =   120
            Width           =   720
         End
         Begin VB.TextBox txt_f_addr 
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
            Left            =   8220
            TabIndex        =   5
            Tag             =   "��׼����"
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox TXT_EMP 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "AGC2041C.frx":0098
            Left            =   5820
            List            =   "AGC2041C.frx":00BD
            TabIndex        =   4
            Top             =   840
            Width           =   1140
         End
         Begin VB.CheckBox chk_Cond_J 
            BackColor       =   &H00E0E0E0&
            Caption         =   "������"
            Height          =   255
            Left            =   11820
            TabIndex        =   3
            Tag             =   "J"
            Top             =   2700
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.ComboBox CBO_SURFGRD 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "AGC2041C.frx":011D
            Left            =   11400
            List            =   "AGC2041C.frx":0133
            TabIndex        =   2
            Tag             =   "�ȼ�"
            Top             =   840
            Width           =   1080
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   315
            Left            =   9240
            TabIndex        =   6
            Top             =   1230
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            Caption         =   "̽�˱���"
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   4620
            Top             =   120
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "���"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   4620
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "����"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel29 
            Height          =   315
            Left            =   10440
            Top             =   120
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            Caption         =   "���"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CSTextLibCtl.sidbEdit SDB_UST_THK 
            Height          =   315
            Left            =   11400
            TabIndex        =   18
            Top             =   120
            Width           =   1080
            _Version        =   262145
            _ExtentX        =   1905
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   195
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "UST��׼"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
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
            Index           =   3
            Left            =   7350
            Top             =   2670
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Caption         =   "��λ��"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   1
            Left            =   6990
            Top             =   120
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "��׼��"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   12570
            Top             =   1200
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "�ϼ�����"
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
            ForeColor       =   255
         End
         Begin CSTextLibCtl.sidbEdit SDB_WGT 
            Height          =   315
            Left            =   13800
            TabIndex        =   19
            Top             =   1200
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.000"
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   10440
            Top             =   480
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            Caption         =   "����"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CSTextLibCtl.sidbEdit SDB_UST_WID 
            Height          =   315
            Left            =   11400
            TabIndex        =   21
            Top             =   480
            Width           =   1080
            _Version        =   262145
            _ExtentX        =   1905
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   0
            Left            =   6990
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "��ѯ��"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   315
            Left            =   6360
            TabIndex        =   23
            Top             =   1230
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            Caption         =   "��ѱ���"
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   2
            Left            =   6990
            Top             =   840
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "��Ѷ�λ"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   3
            Left            =   4620
            Top             =   840
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "̽����Ա"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   195
            Top             =   840
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "��������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   10440
            Top             =   840
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            Caption         =   "����"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   195
            Top             =   1230
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
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.UDate SDT_PROD_DATE 
            Height          =   315
            Left            =   1395
            TabIndex        =   30
            Tag             =   "��ʼ����"
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.UDate SDT_PROD_DATETO 
            Height          =   315
            Left            =   3150
            TabIndex        =   31
            Tag             =   "��ʼ����"
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   195
            Top             =   120
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "̽������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.UDate SDT_PROD_DATE_FR 
            Height          =   315
            Left            =   1395
            TabIndex        =   32
            Tag             =   "��ʼ����"
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.UDate SDT_PROD_DATE_TO 
            Height          =   315
            Left            =   3150
            TabIndex        =   33
            Tag             =   "��ʼ����"
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_UST_THK_TO 
            Height          =   315
            Left            =   12690
            TabIndex        =   36
            Top             =   120
            Width           =   1080
            _Version        =   262145
            _ExtentX        =   1905
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_UST_WID_TO 
            Height          =   315
            Left            =   12690
            TabIndex        =   37
            Top             =   480
            Width           =   1080
            _Version        =   262145
            _ExtentX        =   1905
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Index           =   3
            Left            =   12540
            TabIndex        =   35
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   12540
            TabIndex        =   34
            Top             =   630
            Width           =   195
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   180
            Index           =   0
            Left            =   2880
            TabIndex        =   26
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   240
            Index           =   1
            Left            =   10050
            TabIndex        =   25
            Top             =   2790
            Width           =   300
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   180
            Index           =   2
            Left            =   2880
            TabIndex        =   24
            Top             =   960
            Width           =   240
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7545
         Left            =   0
         TabIndex        =   27
         Top             =   1710
         Width           =   15090
         _Version        =   393216
         _ExtentX        =   26617
         _ExtentY        =   13309
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   49
         MaxRows         =   50
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "AGC2041C.frx":016D
      End
   End
End
Attribute VB_Name = "AGC2041C"
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
'-- Program Name      ̽��ʵ����ѯ����
'-- Program ID        AGC2041C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2005.7.22
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
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_BEF_APLY_STDSPE = 1
Const SS1_UST_LOC = 2
Const SS1_PLATE_NO = 3
Const SS1_SIZE = 4
Const SS1_CNT = 5
Const SS1_BEF_APLY_STDSPEC = 6
Const SS1_STDSPEC_UPD_FL1 = 7
Const SS1_UST_DEC = 8
Const SS1_PROD_GRD = 9
Const SS1_SURF_GRD = 10
Const SS1_BEF_UST_WGT = 11
Const SS1_UST_WGT = 12
Const SS1_UST_MACHINE_NO = 13
Const SS1_UST_HEAD_KIND = 14
Const SS1_UST_METHOD = 15
Const SS1_UST_STATESCOPE = 16
Const SS1_UST_FL = 17
Const SS1_UST_END_DATE = 18
Const SS1_UST_MAN = 19
Const SS1_PROD_DATE = 20
Const SS1_UST_REMARTS = 21
Const SS1_GAS = 22
Const SS1_CL = 23
Const SS1_HTM_SHOT_BLAST = 24
Const SS1_HTM = 25
Const SS1_SIZE_KND = 26
Const SS1_ACT_SMP_FL = 27
Const SS1_CUR_INV = 28
Const SS1_LOC = 29
Const SS1_BED_PILE_DATE = 30
Const SS1_FLAW = 31
Const SS1_THK = 32
Const SS1_WID = 33
Const SS1_LEN = 34
Const SS1_ORD_NO = 35
Const SS1_PROC_CD = 37
Const SS1_SLAB_THK = 38
Const SS1_PRC_LINE = 39
Const SS1_STLGRD_CD = 40
Const SS1_STLGRD = 41
Const SS1_COOLING_TIME = 42
Const SS1_HEAT_NO = 43
Const SS1_OVER_FL = 43


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Refer"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(SDT_PROD_DATE_FR, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_UST_STAND_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_UST_DEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_UST_THK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_UST_WID, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_UST_THK_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '���������
       Call Gp_Ms_Collection(SDB_UST_WID_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '����������
          Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_CO_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(TXT_ADDR(0), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(TXT_ADDR(1), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(TXT_ADDR(2), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_EMP, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDT_PROD_DATETO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_SURFGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-24 ¯����
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-24 ��¯ʱ��
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-24 ��¯�¶�
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-30 �����и�ʱ��
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' ADD BY LICHAO 20140903
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2041C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub

Private Sub CBO_UST_DEC_Click()
   Select Case CBO_UST_DEC.ListIndex
          Case 1
               CBO_UST_DEC.Text = "Y"
          Case 2
               CBO_UST_DEC.Text = "N"
   End Select
End Sub

Private Sub chk_Cond_B_Click()

    If chk_Cond_B Then
        TXT_CO_CD.Text = chk_Cond_B.Tag
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        chk_Cond_W = False
        chk_Cond_J = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False And chk_Cond_J = False Then
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        TXT_CO_CD.Text = ""
    End If
    
End Sub

Private Sub chk_Cond_W_Click()

    If chk_Cond_W Then
        TXT_CO_CD.Text = chk_Cond_W.Tag
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        chk_Cond_B = False
        chk_Cond_J = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False And chk_Cond_J = False Then
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        TXT_CO_CD.Text = ""
    End If
    
End Sub
Private Sub chk_Cond_J_Click()

    If chk_Cond_J Then
        TXT_CO_CD.Text = chk_Cond_J.Tag
        SSCommand1.Enabled = False
        SSCommand2.Enabled = False
        chk_Cond_B = False
        chk_Cond_W = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False And chk_Cond_J = False Then
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        TXT_CO_CD.Text = ""
    End If
    
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_MAT_NO.Text) >= 8 Then
           Call Form_Ref
        End If
'        KeyAscii = 0
'        SendKeys "{TAB}"
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
'    If Trim(TXT_UST_STAND_NO.Text) = "" Then
'        Call Gp_MsgBoxDisplay(TXT_UST_STAND_NO.Tag & "��������", "", "������ʾ")
'        Exit Sub
'    End If
'
'    Call ExcelPrn

End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
     End If

End Sub

Public Sub Form_Ref()
    
    Dim iCount          As Integer
    Dim dMillCal_Wgt    As Double
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    dMillCal_Wgt = 0
    With ss1
        If .MaxRows = 0 Then
            SDB_WGT.Value = 0
            Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount
            .Col = SS1_UST_WGT
             If .Value = 0 Then
                .Col = SS1_BEF_UST_WGT
                 dMillCal_Wgt = dMillCal_Wgt + .Value
             Else
                 dMillCal_Wgt = dMillCal_Wgt + .Value
             End If
        Next iCount
    End With
    SDB_WGT.Value = dMillCal_Wgt
               
End Sub

Public Sub Form_Pro()

'     If Gf_Mc_Authority(sAuthority, Mc1) Then
'       ' txt_ins_emp.Text = sUserID
'       If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'    End If

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    If ss1.MaxRows < 1 Then Exit Sub
    
    If Row = 0 Then 'And (Col = 1 Or Col = 2 Or Col = 3 Or Col = 4) Then
        Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
    End If

End Sub

Private Sub SSCommand1_Click()

    If Trim(TXT_UST_STAND_NO.Text) = "" Then
        Call Gp_MsgBoxDisplay(TXT_UST_STAND_NO.Tag & "��������", "", "������ʾ")
        Exit Sub
    End If
    
    Call ExcelPrn
    
End Sub

Private Sub SSCommand2_Click()

    If Trim(TXT_UST_STAND_NO.Text) = "" Then
        Call Gp_MsgBoxDisplay(TXT_UST_STAND_NO.Tag & "��������", "", "������ʾ")
        Exit Sub
    End If
    
    Call ExcelPrn_Pile

End Sub

Private Sub TXT_STDSPEC_CD_Change()
    If Len(Trim(TXT_STDSPEC_CD)) = TXT_STDSPEC_CD.MaxLength Then
       txt_stdspec.Text = Gf_ComnNameFind(M_CN1, "G0018", Trim(TXT_STDSPEC_CD.Text), 1)
    End If
End Sub

Private Sub TXT_STDSPEC_Change()
    If Len(txt_stdspec.Text) = 0 Then
       TXT_STDSPEC_CD.Text = ""
    End If
End Sub

Private Sub txt_stdspec_DblClick()
    DD.sWitch = "MS"
    DD.sKey = "G0018"
    DD.rControl.Add Item:=TXT_STDSPEC_CD

    DD.nameType = "2"
    
    Call Gf_Common_DD(M_CN1, vbKeyF4)
End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec
        Call Gf_StdSPEC_DD2(M_CN1, vbKeyF4)
        Exit Sub
    End If
End Sub

Private Sub TXT_UST_STAND_NO_Change()
    If Len(TXT_UST_STAND_NO.Text) = 4 Then
       TXT_UST_STAND_NAME.Text = Gf_ComnNameFind(M_CN1, "Q0046", TXT_UST_STAND_NO.Text, 1)
    End If
End Sub

Private Sub TXT_UST_STAND_NO_dblClick()

    DD.sWitch = "MS"
    DD.sKey = "Q0046"
    DD.rControl.Add Item:=TXT_UST_STAND_NO

    DD.nameType = "2"
    
    Call Gf_Mill_Common_DD(M_CN1, vbKeyF4)
    
End Sub

Private Sub ExcelPrn()
    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGC2041C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = SDT_PROD_DATE_FR.Text
    
    If SDT_PROD_DATE_FR.Text <> SDT_PROD_DATE_TO.Text Then
        xlApp.Range("D2").Value = "����: " & Left(sDate, 4) + "��" + Mid(sDate, 6, 2) + "��" + Mid(sDate, 9, 2) + "�� - " + Mid(SDT_PROD_DATE_TO.Text, 9, 2) + "��"
    Else
        xlApp.Range("D2").Value = "����: " & Left(sDate, 4) + "��" + Mid(sDate, 6, 2) + "��" + Mid(sDate, 9, 2) + "��"
    End If
 
    ss1.Row = 1
    ss1.Col = SS1_UST_MACHINE_NO:      xlApp.Range("A4").Value = ss1.Text
    ss1.Col = SS1_UST_HEAD_KIND:       xlApp.Range("B4").Value = ss1.Text
    ss1.Col = SS1_UST_METHOD:          xlApp.Range("C4").Value = ss1.Text
    ss1.Col = SS1_UST_STATESCOPE:      xlApp.Range("D4").Value = ss1.Text
    ss1.Col = SS1_UST_FL:              xlApp.Range("G4").Value = ss1.Text
    
    Clipboard.Clear
    ss1.SetSelection 1, 1, 8, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("A7").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
'
'    xlApp.Range("I2").Select
'    xlApp.ActiveSheet.Paste
    
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
'     xlApp.Application.Visible = False
'     xlSheet.Close False
'     xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub ExcelPrn_Pile()

    Dim i               As Integer
    Dim j               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    Dim sShift          As String
    
    Dim sPage_Num       As Integer
    Dim sPage_X         As Integer
    Dim sPage           As Double
    Dim sLastPage       As Double
    Dim sRow1           As Integer
    Dim sRow2           As Integer
    
    Dim xl_A            As String
    Dim xl_B            As String
    Dim xl_C            As String
    Dim xl_E            As String
    Dim xl_F            As String
    Dim xl_G            As String
    Dim xl_H            As String
    Dim xl_I            As String
    Dim xl_J            As String
    Dim xl_K            As String
    
    Dim xl_clr_body     As String
    Dim xl_clr_sum      As String
    Dim xl_clr_spc      As String
    
    Dim Xl_Cnt          As String
    Dim Xl_Wgt          As String
    Dim Xl_Wgt_Val      As String
    Dim Xl_Ust          As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGC2043C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = SDT_PROD_DATE_FR.Text
    
    If SDT_PROD_DATE_FR.Text <> SDT_PROD_DATE_TO.Text Then
        xlApp.Range("A2").Value = "����: " & Left(sDate, 4) + "��" + Mid(sDate, 6, 2) + "��" + Mid(sDate, 9, 2) + "�� - " + Mid(SDT_PROD_DATE_TO.Text, 9, 2) + "��"
    Else
        xlApp.Range("A2").Value = "����: " & Left(sDate, 4) + "��" + Mid(sDate, 6, 2) + "��" + Mid(sDate, 9, 2) + "��"
    End If
    
    If CBO_SHIFT.Text = "1" Then
       sShift = "��ҹ��"
    ElseIf CBO_SHIFT.Text = "2" Then
       sShift = "�װ�"
    ElseIf CBO_SHIFT.Text = "3" Then
       sShift = "Сҹ��"
    Else
       sShift = ""
    End If
    
    xlApp.Range("C2").Value = Mid(xlApp.Range("C2").Value, 1, 3) & sShift
        
    sPage_Num = 30
    sPage_X = 32
    
    sPage = Int(ss1.MaxRows / sPage_Num) + 1
    sLastPage = ss1.MaxRows - Int(ss1.MaxRows / sPage_Num) * sPage_Num
    
    For i = 0 To 11
        xl_clr_body = "A" + CStr(4 + i * sPage_X) + ":" + "I" + CStr(33 + i * sPage_X)
        xl_clr_sum = "C" + CStr(34 + i * sPage_X) + ":" + "C" + CStr(35 + i * sPage_X)
        xl_clr_spc = "D" + CStr(34 + i * sPage_X)
        xlApp.Range(xl_clr_body).Value = Null
        xlApp.Range(xl_clr_sum).Value = Null
        xlApp.Range(xl_clr_spc).Value = Mid(xlApp.Range(xl_clr_spc).Value, 1, 5)
    Next i
    
    For i = 0 To sPage - 1
       
        sRow1 = 1 + sPage_Num * i
        sRow2 = sPage_Num * (i + 1)

        If i = sPage - 1 Then
           sRow2 = sPage_Num * i + sLastPage
        End If

        xl_A = "A" + CStr(4 + i * sPage_X)
        xl_B = "B" + CStr(4 + i * sPage_X)
        xl_C = "C" + CStr(4 + i * sPage_X)
        xl_E = "E" + CStr(4 + i * sPage_X)
        xl_F = "F" + CStr(4 + i * sPage_X)
        xl_G = "G" + CStr(4 + i * sPage_X)
        xl_H = "H" + CStr(4 + i * sPage_X)
        xl_I = "I" + CStr(4 + i * sPage_X)
        xl_J = "J" + CStr(4 + i * sPage_X)
        xl_K = "K" + CStr(4 + i * sPage_X)
        
        Xl_Cnt = "C" + CStr(3 + (i + 1) * sPage_X - 1)
        Xl_Wgt = "C" + CStr(3 + (i + 1) * sPage_X)
        Xl_Ust = "D" + CStr(3 + (i + 1) * sPage_X - 1)
        
        Clipboard.Clear
        ss1.SetSelection 2, sRow1, 2, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_A).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 1, sRow1, 1, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_B).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 3, sRow1, 4, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_C).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 12, sRow1, 12, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_E).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 8, sRow1, 8, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_F).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 19, sRow1, 19, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_G).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 7, sRow1, 7, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_H).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 36, sRow1, 36, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_I).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 35, sRow1, 35, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_J).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection 21, sRow1, 21, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_K).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        If i = sPage - 1 Then
           xlApp.Range(Xl_Cnt).Value = sLastPage
        Else
           xlApp.Range(Xl_Cnt).Value = sPage_Num
        End If
        
        For j = 1 To sPage_Num
            Xl_Wgt_Val = "E" & CStr((Val(Mid(xl_E, 2)) + j - 1))
            xlApp.Range(Xl_Wgt).Value = xlApp.Range(Xl_Wgt).Value + xlApp.Range(Xl_Wgt_Val).Value
        Next j
        
        ss1.Row = 1
        ss1.Col = 17
        xlApp.Range(Xl_Ust).Value = xlApp.Range(Xl_Ust).Value + ss1.Text
              
    Next i
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub



Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub
Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1

End Sub
Private Sub txt_f_addr_DblClick()
     Call txt_f_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "F0009"
        txt_f_addr.Text = "P"
        DD.rControl.Add Item:=txt_f_addr
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If
    

End Sub
Public Sub Spread_Forzens_Setting()
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
End Sub

Public Sub Spread_Forzens_Cancel()
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
End Sub