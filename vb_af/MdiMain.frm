VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "���������ҵ����"
   ClientHeight    =   8700
   ClientLeft      =   1800
   ClientTop       =   2775
   ClientWidth     =   12390
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Tag             =   "F"
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet 
      Left            =   15
      Top             =   2640
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
                  Object.ToolTipText     =   "�ս���"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Refer"
                  Object.ToolTipText     =   "��ѯ"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line1"
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Save"
                  Object.ToolTipText     =   "����"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Delete"
                  Object.ToolTipText     =   "ɾ��"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line2"
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "RowIns"
                  Object.ToolTipText     =   "׷����"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "RowDel"
                  Object.ToolTipText     =   "ɾ����"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "RowCan"
                  Object.ToolTipText     =   "ȡ����"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line3"
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Copy"
                  Object.ToolTipText     =   "����"
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
                  Object.ToolTipText     =   "ճ��"
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
                  Object.ToolTipText     =   "����"
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Print"
                  Object.ToolTipText     =   "��ӡ"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line5"
                  Style           =   3
               EndProperty
               BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Exit"
                  Object.ToolTipText     =   "�˳�"
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
      Top             =   8235
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
            TextSave        =   "2016-06-17"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "10:59"
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
         Name            =   "����"
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
         Caption         =   "������"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_FrozenSetting 
         Caption         =   "�ж�������"
      End
      Begin VB.Menu Mnu_FrozenCancel 
         Caption         =   "��ȡ������"
      End
   End
   Begin VB.Menu Mnu_AFA 
      Caption         =   "����վ/Ԥ����"
      Begin VB.Menu Mnu_AFA3000C 
         Caption         =   "��¯��ˮʵ������"
      End
      Begin VB.Menu Mnu_AFA2000C 
         Caption         =   "����վʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFB2010C 
         Caption         =   "��ˮԤ����ʵ���޸ļ���ѯ����"
      End
   End
   Begin VB.Menu Mnu_AFC 
      Caption         =   "ת¯��ҵ"
      Begin VB.Menu Mnu_AFC2010C 
         Caption         =   "ת¯ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFC2020C 
         Caption         =   "CASʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFG2010C 
         Caption         =   "������/��ԭ��ʹ����ʵ���޸ļ���ѯ����"
      End
   End
   Begin VB.Menu Mnu_AFD 
      Caption         =   "¯�⾫����ҵ"
      Begin VB.Menu Mnu_AFE2010C 
         Caption         =   "LF ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFF2010C 
         Caption         =   "VD ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFF2020C 
         Caption         =   "RH ʵ���޸ļ���ѯ����"
      End
   End
   Begin VB.Menu Mnu_AFF 
      Caption         =   "������ҵ"
      Begin VB.Menu Mnu_AFH2010C 
         Caption         =   "����ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFL2050C 
         Caption         =   "�����и�/װ¯��״����"
      End
      Begin VB.Menu Mnu_AFH2020C 
         Caption         =   "�����и�ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFH2040C 
         Caption         =   "����ȱ��ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFH6010C 
         Caption         =   "�м��ʵ��"
      End
   End
   Begin VB.Menu Mnu_AFG 
      Caption         =   "�ɷַ���"
      Begin VB.Menu Mnu_AFK2010C 
         Caption         =   "��ѧ�ɷ�ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFK2030C 
         Caption         =   "�ɷ��ж�ʵ����ѯ����"
      End
   End
   Begin VB.Menu Mnu_AFH 
      Caption         =   "������"
      Begin VB.Menu Mnu_AFL2010C 
         Caption         =   "���������޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFL2040C 
         Caption         =   "�������ͼ����"
      End
      Begin VB.Menu Mnu_AFL2060C 
         Caption         =   "��׼��λ��������"
      End
      Begin VB.Menu Mnu_AFL2070C 
         Caption         =   "���״̬��ѯ����"
      End
      Begin VB.Menu Mnu_AFL2080C 
         Caption         =   "�����������ѯ����"
      End
      Begin VB.Menu Mnu_AFL2090C 
         Caption         =   "�����������ѯ����"
      End
      Begin VB.Menu Mnu_AFL2030C 
         Caption         =   "���Ͱ��������ʵ��¼�����"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AFL2100C 
         Caption         =   "����������ʷ����ѯ"
      End
      Begin VB.Menu Mnu_AFL2101C 
         Caption         =   "����ͳ������ۺϲ�ѯ"
      End
      Begin VB.Menu Mnu_AFL2102C 
         Caption         =   "������������"
      End
      Begin VB.Menu Mnu_AFL2103C 
         Caption         =   "����CADͳ������ۺϲ�ѯ"
      End
      Begin VB.Menu Mnu_AFL2110C 
         Caption         =   "����������¼��"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_AFI 
      Caption         =   "��/����"
      Begin VB.Menu Mnu_AFM2010C 
         Caption         =   "��/����ά��ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFM2020C 
         Caption         =   "��/����������״��ѯ����"
      End
   End
   Begin VB.Menu Mnu_AFJ 
      Caption         =   "�쳣��"
      Begin VB.Menu Mnu_AFM2030C 
         Caption         =   "���͸�ˮʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFT4060C 
         Caption         =   "�����ж�ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFM2040C 
         Caption         =   "������ĥ���ϸ�ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFM2050C 
         Caption         =   "������ĥ���ϸ�ʵ����ѯ����"
      End
      Begin VB.Menu Mnu_AFM2060C 
         Caption         =   "�����и���ҵ"
      End
      Begin VB.Menu Mnu_AFM2080C 
         Caption         =   "����������ҵ"
      End
      Begin VB.Menu Mnu_AFM2090C 
         Caption         =   "������ĥ���ʵ���޸ļ���ѯ����"
      End
   End
   Begin VB.Menu Mnu_AFK 
      Caption         =   "��ҵָʾ"
      Visible         =   0   'False
      Begin VB.Menu Mnu_AFN2010C 
         Caption         =   "������ҵָʾ��ѯ����"
      End
      Begin VB.Menu Mnu_AFN2020C 
         Caption         =   "�����и���ҵָʾ��ѯ����"
      End
   End
   Begin VB.Menu Mnu_AFL 
      Caption         =   "��������ʵ��"
      Begin VB.Menu Mnu_AFO2020C 
         Caption         =   "���ֹ�������ʹ��ʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AFO2030C 
         Caption         =   "����������ͣ��ʵ������"
      End
   End
   Begin VB.Menu Mnu_AFN 
      Caption         =   "������¼"
      Begin VB.Menu Mnu_AFC5100C 
         Caption         =   "ת¯ԭʼ������¼"
      End
      Begin VB.Menu Mnu_AFE5010C 
         Caption         =   "����ԭʼ������¼"
      End
      Begin VB.Menu Mnu_AFH5010C 
         Caption         =   "����ԭʼ������¼"
      End
      Begin VB.Menu Mnu_AFN1 
         Caption         =   "�յ���Ϣ��ѯ"
         Begin VB.Menu Mnu_AFT3014C 
            Caption         =   "��ˮԤ������ת¯�յ���Ϣ��ѯ"
         End
         Begin VB.Menu Mnu_AFT3013C 
            Caption         =   "�����������յ���Ϣ��ѯ"
         End
      End
   End
   Begin VB.Menu Mnu_AFP 
      Caption         =   "����"
      Begin VB.Menu Mnu_AFP1000C 
         Caption         =   "��¯��ˮʵ��������ɾ��"
      End
      Begin VB.Menu Mnu_AFP1010C 
         Caption         =   "����վʵ������"
      End
      Begin VB.Menu Mnu_AFP2010C 
         Caption         =   "��ˮԤ����ʵ������"
      End
      Begin VB.Menu Mnu_AFP3010C 
         Caption         =   "ת¯ʵ������"
      End
      Begin VB.Menu Mnu_AFP3012C 
         Caption         =   "CASʵ������"
      End
      Begin VB.Menu Mnu_AFP4010C 
         Caption         =   "LFʵ������"
      End
      Begin VB.Menu Mnu_AFP5010C 
         Caption         =   "VDʵ������"
      End
      Begin VB.Menu Mnu_AFP5020C 
         Caption         =   "RHʵ������"
      End
      Begin VB.Menu Mnu_AFP6010C 
         Caption         =   "����ʵ������"
      End
      Begin VB.Menu Mnu_AFP7010C 
         Caption         =   "���͸�ˮʵ������"
      End
      Begin VB.Menu Mnu_AFP8010C 
         Caption         =   "��ѧ�ɷ�ʵ������"
      End
      Begin VB.Menu Mnu_AFP9010C 
         Caption         =   "���ֳ���ְ�׼��������������"
      End
      Begin VB.Menu Mnu_AFP3011C 
         Caption         =   "����������ʹ��ʵ������"
      End
   End
   Begin VB.Menu Mnu_AFM 
      Caption         =   "����"
      Begin VB.Menu Mnu_AFO3020C 
         Caption         =   "���ֹ���ʱ�̱�"
      End
      Begin VB.Menu Mnu_AFO2010C 
         Caption         =   "���������ڽ��̸��ٽ���"
      End
      Begin VB.Menu Mnu_AFH2030C 
         Caption         =   "�м��ά��ʵ���޸ļ���ѯ����"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AFO3010C 
         Caption         =   "�ϸִ����´����"
      End
      Begin VB.Menu Mnu_AFO2060C 
         Caption         =   "����ԭ�ϳ�/���(��)��ѯ����"
      End
      Begin VB.Menu Mnu_AFO2070C 
         Caption         =   "����ԭ�ϳ�/���(��)��ѯ����"
      End
      Begin VB.Menu Mnu_AFO2080C 
         Caption         =   "ԭ������������"
      End
      Begin VB.Menu Mnu_AFO2090C 
         Caption         =   "ԭ�ϳ����������"
      End
      Begin VB.Menu Mnu_AFT2010C 
         Caption         =   "�����������������"
      End
      Begin VB.Menu Mnu_AFT3010C 
         Caption         =   "������־��ѯ"
      End
      Begin VB.Menu Mnu_AFT3011C 
         Caption         =   "�����ƻ���ʵ�ʱȽ�"
      End
      Begin VB.Menu Mnu_AFT3015C 
         Caption         =   "�����Ǽƻ���ѯ"
      End
      Begin VB.Menu Mnu_AFT3016C 
         Caption         =   "�����Ǽƻ�ͳ�Ʊ���"
      End
      Begin VB.Menu Mnu_AFT3030C 
         Caption         =   "�����������ٲ�ѯ_AFT3030C"
      End
      Begin VB.Menu Mnu_line 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_AFT3999C 
         Caption         =   "���а峧��������¼�����(����)"
      End
      Begin VB.Menu Mnu_AFT3998C 
         Caption         =   "���а峧����������¼����(����)"
      End
      Begin VB.Menu Mnu_AFT4000C 
         Caption         =   "�⹺��ʵ���޸ļ���ѯ����_AFT4000C"
      End
      Begin VB.Menu Mnu_AFT4010C 
         Caption         =   "ұ�����ڱ���"
      End
      Begin VB.Menu Mnu_AFT4020C 
         Caption         =   "���������ϸ����ı���"
      End
      Begin VB.Menu Mnu_AFT4030C 
         Caption         =   "�ְ�ʢ��ʱ����ٱ���"
      End
      Begin VB.Menu Mnu_AFT4040C 
         Caption         =   "������汨��"
      End
      Begin VB.Menu Mnu_AFT4050C 
         Caption         =   "�����������ݲɼ�����"
      End
      Begin VB.Menu Mnu_AFT4070C 
         Caption         =   "���������ƻ���ѯ���޸Ľ���"
      End
      Begin VB.Menu Mnu_AFT4080C 
         Caption         =   "���������Ϣ���ݽ���_AFT4080C"
      End
      Begin VB.Menu Mnu_AFT4090C 
         Caption         =   "������������ϵ���ձ�_AFT4090C"
      End
   End
   Begin VB.Menu Mnu_Windows 
      Caption         =   "Window"
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
      Begin VB.Menu Line4 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Help 
         Caption         =   "����˵����"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

    'Call MDIMain.FormMenuSetting(me,"Start", Toolbar_St,"")

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
'         Active_YN = GetSetting("NISCO", "EXE-FILE", "AF.exe")
'
'         If Active_YN = "1" Then
'             sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'             sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'             MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
'             MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'         Else
'             Call Gp_MsgBoxDisplay("Process Management...Exectue", "W")
'            Unload Me
'         End If
'
''        sUserID = "1JS6005"
''        sUserName = "NISCO"
''        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
''        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
''''
'        If Mid(M_CN1, Len(M_CN1), 1) = "9" Then
'            MDIMain.StatusBar1.Panels(8) = "��ʽ��"
'        Else
'            MDIMain.StatusBar1.Panels(8) = "���Ի�"
'        End If
'
'    End If
'
'End Sub
Private Sub MDIForm_Load()

    Dim Active_YN As String
    Dim args  As Variant ' 2012.11.09 ����  ������
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
    Me.BackColor = &HE0E0E0
    
    If GF_DbConnect = False Then
        
        Unload Me
    
    Else

    args = Split(Trim(Command), " ") ' 2012.11.09 ����  ������
    If UBound(args) = 1 Then
         MainFrmType = "New"
         sUserID = args(0) ' 2012.11.09 ����  ������
         sUserName = args(1) ' 2012.11.09 ����  ������
         MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��" ' 2012.11.09 ����  ������
         MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName ' 2012.11.09 ����  ������
    Else
        Active_YN = GetSetting("NISCO", "EXE-FILE", "AF.exe")
        If Active_YN = "1" Then
            MainFrmType = "Old"
            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
            MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ�� ��"
            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
        Else
            Call Gp_MsgBoxDisplay("ֻ�ܴ��������½...", "W")
            Unload Me
            Exit Sub
        End If
    End If  ' 2012.11.09 ����  ������



'
'        sUserID = "1JS1014"
'        sUserName = "�¾���"
'        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
'        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'
'''
   If Mid(M_CN1, Len(M_CN1), 1) = "9" Then
            MDIMain.StatusBar1.Panels(8) = "��ʽ��"
        Else
            MDIMain.StatusBar1.Panels(8) = "���Ի�"
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
    
        If Gf_MessConfirm("����δ�رյĲ������棬" + vbCrLf + "�Ƿ��˳���ǰϵͳ��", "Q", Me.Caption) Then
            
            For Each CurrentForm In Forms
                If CurrentForm.Name <> Me.Name Then
                    Unload CurrentForm
                End If
            Next CurrentForm
            
        Else
            Cancel = True
        End If
        
    End If
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "AF.exe", ""

End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
    
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
    
    If Screen.ActiveForm.Name = "MDIMain" Then
        
        If Button.Key = "Exit" Then
            If vbYes = MsgBox(Me.Caption + " ϵͳ�Ƿ��˳� ?", vbQuestion + vbYesNo, Me.Caption) Then
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

    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ��"
    
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

Private Sub Mnu_ACB4070C_Click()
     ACB4070C.Show
     ACB4070C.SetFocus
End Sub

Private Sub Mnu_AFA2000C_Click()
     AFA2010C.Show
     AFA2010C.SetFocus
End Sub

Private Sub Mnu_AFA3000C_Click()
     AFA2000C.Show
     AFA2000C.SetFocus
End Sub

Private Sub Mnu_AFB2010C_Click()
    AFB2010C.Show
    AFB2010C.SetFocus
End Sub

Private Sub Mnu_AFC2010C_Click()
   AFC2010C.Show
   AFC2010C.SetFocus
End Sub

Private Sub Mnu_AFC2020C_Click()
   AFC2020C.Show
   AFC2020C.SetFocus
End Sub

Private Sub Mnu_AFC5100C_Click()
    AFC5100C.Show
    AFC5100C.SetFocus
End Sub

Private Sub Mnu_AFE2010C_Click()
    AFE2010C.Show
    AFE2010C.SetFocus
End Sub

Private Sub Mnu_AFE5010C_Click()
    AFE5010C.Show
    AFE5010C.SetFocus
End Sub

Private Sub Mnu_AFF2010C_Click()
    AFF2010C.Show
    AFF2010C.SetFocus
End Sub
Private Sub Mnu_AFF2020C_Click()
    AFF2020C.Show
    AFF2020C.SetFocus
End Sub

Private Sub Mnu_AFG2010C_Click()
    AFG2010C.Show
    AFG2010C.SetFocus
End Sub

Private Sub Mnu_AFH2010C_Click()
    AFH2010C.Show
    AFH2010C.SetFocus
End Sub

Private Sub Mnu_AFH2020C_Click()
    AFH2020C.Show
    AFH2020C.SetFocus
End Sub

Private Sub Mnu_AFH2030C_Click()
    AFH2030C.Show
    AFH2030C.SetFocus
End Sub

Private Sub Mnu_AFH2040C_Click()
    AFH2040C.Show
    AFH2040C.SetFocus
End Sub

Private Sub Mnu_AFH5010C_Click()
    AFH5010C.Show
    AFH5010C.SetFocus
End Sub

Private Sub Mnu_AFH6010C_Click()
    AFH6010C.Show
    AFH6010C.SetFocus
End Sub

Private Sub Mnu_AFK2010C_Click()
    AFK2010C.Show
    AFK2010C.SetFocus
End Sub

Private Sub Mnu_AFK2030C_Click()
    AFK2030C.Show
    AFK2030C.SetFocus
End Sub

Private Sub Mnu_AFL2010C_Click()
    AFL2010C.Show
    AFL2010C.SetFocus
End Sub

Private Sub Mnu_AFL2030C_Click()
    AFL2030C.Show
    AFL2030C.SetFocus
End Sub

Private Sub Mnu_AFL2040C_Click()
    AFL2040C.Show
    AFL2040C.SetFocus
End Sub

Private Sub Mnu_AFL2050C_Click()
    AFL2050C.Show
    AFL2050C.SetFocus
End Sub

Private Sub Mnu_AFL2060C_Click()
    AFL2060C.Show
    AFL2060C.SetFocus
End Sub

Private Sub Mnu_AFL2070C_Click()
    AFL2070C.Show
    AFL2070C.SetFocus
End Sub

Private Sub Mnu_AFL2080C_Click()
    AFL2080C.Show
    AFL2080C.SetFocus
End Sub

Private Sub Mnu_AFL2090C_Click()
    AFL2090C.Show
    AFL2090C.SetFocus
End Sub

Private Sub Mnu_AFL2100C_Click()
    AFL2100C.Show
    AFL2100C.SetFocus
End Sub

Private Sub Mnu_AFL2101C_Click()
    AFL2101C.Show
    AFL2101C.SetFocus
End Sub

Private Sub Mnu_AFL2102C_Click()
    AFL2102C.Show
    AFL2102C.SetFocus
End Sub

Private Sub Mnu_AFL2103C_Click()
    AFL2103C.Show
    AFL2103C.SetFocus
End Sub

Private Sub Mnu_AFL2110C_Click()
    AFL2110C.Show
    AFL2110C.SetFocus
End Sub

Private Sub Mnu_AFM2010C_Click()
    AFM2010C.Show
    AFM2010C.SetFocus
End Sub

Private Sub Mnu_AFM2020C_Click()
    AFM2020C.Show
    AFM2020C.SetFocus
End Sub

Private Sub Mnu_AFM2030C_Click()
    AFM2030C.Show
    AFM2030C.SetFocus
End Sub

Private Sub Mnu_AFM2040C_Click()
    AFM2040C.Show
    AFM2040C.SetFocus
End Sub

Private Sub Mnu_AFM2050C_Click()
    AFM2050C.Show
    AFM2050C.SetFocus
End Sub

Private Sub Mnu_AFM2060C_Click()
    AFM2060C.Show
    AFM2060C.SetFocus
End Sub

Private Sub Mnu_AFM2080C_Click()
    AFM2080C.Show
    AFM2080C.SetFocus
End Sub

Private Sub Mnu_AFM2090C_Click()
    AFM2090C.Show
    AFM2090C.SetFocus
End Sub

Private Sub Mnu_AFN2010C_Click()
    AFN2010C.Show
    AFN2010C.SetFocus
End Sub

Private Sub Mnu_AFN2020C_Click()
    AFN2020C.Show
    AFN2020C.SetFocus
End Sub

Private Sub Mnu_AFO2010C_Click()
    AFO2010C.Show
    AFO2010C.SetFocus
End Sub

Private Sub Mnu_AFO2020C_Click()
    AFO2020C.Show
    AFO2020C.SetFocus
End Sub

Private Sub Mnu_AFO2030C_Click()
    AFO2030C.Show
    AFO2030C.SetFocus
End Sub

Private Sub Mnu_AFO2060C_Click()
    AFO2060C.Show
    AFO2060C.SetFocus
End Sub

Private Sub Mnu_AFO2070C_Click()
    AFO2070C.Show
    AFO2070C.SetFocus
End Sub

Private Sub Mnu_AFO2080C_Click()
    AFO2090C.Show
    AFO2090C.SetFocus
End Sub

Private Sub Mnu_AFO2090C_Click()
    AFO2080C.Show
    AFO2080C.SetFocus
End Sub

Private Sub Mnu_AFO3010C_Click()
    AFO3010C.Show
    AFO3010C.SetFocus
End Sub

Private Sub Mnu_AFO3020C_Click()
    AFO3020C.Show
    AFO3020C.SetFocus
End Sub

Private Sub Mnu_AFP1000C_Click()
    AFP1000C.Show
    AFP1000C.SetFocus
End Sub

Private Sub Mnu_AFP1010C_Click()
    AFP1010C.Show
    AFP1010C.SetFocus
End Sub

Private Sub Mnu_AFP2010C_Click()
    AFP2010C.Show
    AFP2010C.SetFocus
End Sub

Private Sub Mnu_AFP3010C_Click()
    AFP3010C.Show
    AFP3010C.SetFocus
End Sub

Private Sub Mnu_AFP3011C_Click()
    AFP3011C.Show
    AFP3011C.SetFocus
End Sub

Private Sub Mnu_AFP3012C_Click()
    AFP3012C.Show
    AFP3012C.SetFocus
End Sub

Private Sub Mnu_AFP4010C_Click()
    AFP4010C.Show
    AFP4010C.SetFocus
End Sub

Private Sub Mnu_AFP5010C_Click()
    AFP5010C.Show
    AFP5010C.SetFocus
End Sub

Private Sub Mnu_AFP5020C_Click()
    AFP5020C.Show
    AFP5020C.SetFocus
End Sub

Private Sub Mnu_AFP6010C_Click()
    AFP6010C.Show
    AFP6010C.SetFocus
End Sub

Private Sub Mnu_AFP7010C_Click()
    AFP7010C.Show
    AFP7010C.SetFocus
End Sub

Private Sub Mnu_AFP8010C_Click()
    AFP8010C.Show
    AFP8010C.SetFocus
End Sub
Private Sub Mnu_AFP9010C_Click()
    AFP9010C.Show
    AFP9010C.SetFocus
End Sub

Private Sub Mnu_AFT2010C_Click()
    AFT2010C.Show
    AFT2010C.SetFocus
End Sub

Private Sub Mnu_AFT3010C_Click()
    AFT3010C.Show
    AFT3010C.SetFocus
End Sub

Private Sub Mnu_AFT3011C_Click()
    AFT3011C.Show
    AFT3011C.SetFocus
End Sub

Private Sub Mnu_AFT3013C_Click()
    AFT3013C.Show
    AFT3013C.SetFocus
End Sub

Private Sub Mnu_AFT3014C_Click()
    AFT3014C.Show
    AFT3014C.SetFocus
End Sub

Private Sub Mnu_AFT3015C_Click()
    AFT3015C.Show
    AFT3015C.SetFocus
End Sub

Private Sub Mnu_AFT3016C_Click()
    AFT3016C.Show
    AFT3016C.SetFocus
End Sub

Private Sub Mnu_AFT3030C_Click()
    AFT3030C.Show
    AFT3030C.SetFocus
End Sub

Private Sub Mnu_AFT3998C_Click()
    AFT3998C.Show
    AFT3998C.SetFocus
End Sub

Private Sub Mnu_AFT3999C_Click()
    AFT3999C.Show
    AFT3999C.SetFocus
End Sub
Private Sub Mnu_AFT4000C_Click()
    AFT4000C.Show
    AFT4000C.SetFocus
End Sub

Private Sub Mnu_AFT4010C_Click()
    AFT4010C.Show
    AFT4010C.SetFocus
End Sub

Private Sub Mnu_AFT4020C_Click()
    AFT4020C.Show
    AFT4020C.SetFocus
End Sub

Private Sub Mnu_AFT4030C_Click()
    AFT4030C.Show
    AFT4030C.SetFocus
End Sub

Private Sub Mnu_AFT4040C_Click()
    AFT4040C.Show
    AFT4040C.SetFocus
End Sub

Private Sub Mnu_AFT4050C_Click()
    AFT4050C.Show
    AFT4050C.SetFocus
End Sub

Private Sub Mnu_AFT4060C_Click()
    AFT4060C.Show
    AFT4060C.SetFocus
End Sub

Private Sub Mnu_AFT4070C_Click()
    AFT4070C.Show
    AFT4070C.SetFocus
End Sub

Private Sub Mnu_AFT4080C_Click()
    AFT4080C.Show
    AFT4080C.SetFocus
End Sub

Private Sub Mnu_AFT4090C_Click()
    AFT4090C.Show
    AFT4090C.SetFocus
End Sub

Private Sub Mnu_Cascade_Click()
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ��"
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
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ��"
    Call ActiveForm.Spread_Forzens_Cancel
End Sub

Private Sub Mnu_FrozenSetting_Click()
    'Spread Col Frozens Setting
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ��"
    Call ActiveForm.Spread_Forzens_Setting
End Sub

Private Sub Mnu_Help_Click()
    Dim FormLD As Boolean
    
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
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ��"
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
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ��"
    Call ActiveForm.Spread_ColumnsSort
End Sub

Private Sub Mnu_Spaste_Click()
    'Spread Row Paste
    Call ActiveForm.Spread_Pst
End Sub

Private Sub Mnu_Vertical_Click()
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ��"
    MDIMain.Arrange 2
End Sub