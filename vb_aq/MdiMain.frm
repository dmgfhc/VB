VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "��������"
   ClientHeight    =   6795
   ClientLeft      =   870
   ClientTop       =   2835
   ClientWidth     =   11280
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11220
      TabIndex        =   0
      Top             =   0
      Width           =   11280
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
      Top             =   6330
      Width           =   11280
      _ExtentX        =   19897
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
            Enabled         =   0   'False
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
            TextSave        =   "2017-04-21"
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
   Begin VB.Menu Mnu_AQA 
      Caption         =   "������׼����"
      Begin VB.Menu Mnu_AQA0010C 
         Caption         =   "��׼������Ϣ��ѯ(AQA0010C)"
      End
      Begin VB.Menu Mnu_AQA0130C 
         Caption         =   "���ɷֲ�ѯ(AQA0130C)"
      End
      Begin VB.Menu Mnu_AQA0140C 
         Caption         =   "��������Ϣ����(AQA0140C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0170C 
         Caption         =   "������Ƽ���ѯ(AQA0170C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0190C 
         Caption         =   "������;��ѯ(AQA0190C)"
      End
      Begin VB.Menu Mnu_AQA0200C 
         Caption         =   "��׼������;��ѯ(AQA0200C)"
      End
      Begin VB.Menu Mnu_AQA0210C 
         Caption         =   "����/�����������(AQA0210C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0220C 
         Caption         =   "���ֹ������(AQA0220C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0221C 
         Caption         =   "���ȹ�̸��ַ������(AQA0221C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0222C 
         Caption         =   "���ȹ������(AQA0222C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0260C 
         Caption         =   "ȡ����׼��¼(AQA0260C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0270C 
         Caption         =   "��������֤��������֤������(AQA0270C)"
      End
      Begin VB.Menu Mnu_AQA0271C 
         Caption         =   "���彻��״̬ά������AQA0271C)"
      End
      Begin VB.Menu Mnu_AQA0272C 
         Caption         =   "ERP��������������Ϣ(AQA0271C)"
      End
      Begin VB.Menu Mnu_AQA0280C 
         Caption         =   "��Ʒ����������������(AQA0280C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0400C 
         Caption         =   "��������׼����(AQA0400C)"
      End
      Begin VB.Menu Mnu_AQA0410C 
         Caption         =   "��Ʒ�ɷ���Լ��׼����(AQA0410C)"
      End
      Begin VB.Menu Mnu_AQA0420C 
         Caption         =   "¼����ͬ¯�γɷ�ƫ�����(AQA0420C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0430C 
         Caption         =   "�ȴ���������������ѯ(AQA0430C)"
      End
      Begin VB.Menu Mnu_AQA0440C 
         Caption         =   "�ȴ�����ҵ��׼��ѯ(AQA0440C)"
      End
      Begin VB.Menu Mnu_AQA0450C 
         Caption         =   "�ȴ�������/�������ʱ�׼����(AQA0450C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQA0470C 
         Caption         =   "�ȴ���������ѯ���£�(AQA0470C)"
      End
      Begin VB.Menu Mnu_AQA0460C 
         Caption         =   "��Ʒϸ�����ѯ����(AQA0460C)"
      End
      Begin VB.Menu Mnu_AQA0500C 
         Caption         =   "��ȸ���ֵ��ѯ(AQA0500C)"
      End
      Begin VB.Menu Mnu_AQA0600C 
         Caption         =   "�������������淶ά������(AQA0600C)"
      End
      Begin VB.Menu Mnu_AQA0700C 
         Caption         =   "�Զ����й���ά��(AQA0700C)"
      End
      Begin VB.Menu Mnu_AQA0800C 
         Caption         =   "�ϲ�ȡ������ά��(AQA0800C)"
      End
      Begin VB.Menu Mnu_AQA0900C 
         Caption         =   "�ᾧ����ȴ�ƶȹ�̽��棨AQA0900C��"
      End
      Begin VB.Menu Mnu_AQA0910C 
         Caption         =   "�����ƶȹ�̽��棨AQA0910C��"
      End
      Begin VB.Menu Mnu_AQA0920C 
         Caption         =   "�񶯲�����̽��棨AQA0920C��"
      End
   End
   Begin VB.Menu Mnu_AQB 
      Caption         =   "�������"
      Begin VB.Menu Mnu_AQB0010C 
         Caption         =   "������ƽ�����״��ѯ(AQB0010C)"
      End
      Begin VB.Menu mnu_AQB0110C 
         Caption         =   "�ɷ���ƽ����ѯ(AQB0110C)"
      End
      Begin VB.Menu mnu_AQB0120C 
         Caption         =   "������ƽ����ѯ(AQB0120C)"
      End
      Begin VB.Menu mnu_AQB0121C 
         Caption         =   "PWHT������ƽ����ѯ(AQB0121C)"
      End
      Begin VB.Menu mnu_AQB0150C 
         Caption         =   "����������ƽ����ѯ(AQB0150C)"
      End
      Begin VB.Menu mnu_AQB0160C 
         Caption         =   "�����淶��ƽ����ѯ(AQB0160C)"
      End
      Begin VB.Menu mnu_AQB0200C 
         Caption         =   "������� ERROR ��ѯ(AQB0200C)"
      End
      Begin VB.Menu Mnu_AQF1 
         Caption         =   "����������Ϣ����"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQF0010C 
         Caption         =   "����ȡ����׼����(AQF0010C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQF0100C 
         Caption         =   "����ȡ��ʵ������(AQF0100C)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_AQC 
      Caption         =   "�ж�����"
      Begin VB.Menu mnu_AQC0060C 
         Caption         =   "�ְ�ȡ����Ϣ��ѯ���޸�(AQC0060C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQC0010C 
         Caption         =   "���������״��ѯ(AQC0010C)"
      End
      Begin VB.Menu Mnu_AQC0034C 
         Caption         =   "��Ʒ����ʵ��¼�루��ѧ��_AQC0034C"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQC0032C 
         Caption         =   "��Ʒ����ʵ��¼�루���ࣩ_AQC0032C"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQC0040C 
         Caption         =   "����ʵ��ȷ��(AQC0040C)"
      End
      Begin VB.Menu Mnu_AQC0050C 
         Caption         =   "����ָʾ���´����(AQC0050C)"
      End
      Begin VB.Menu Mnu_AQC0051C 
         Caption         =   "PWHT����ָʾ���´����(AQC0051C)"
      End
      Begin VB.Menu mnu_AQC0080C 
         Caption         =   "�ɷ�/������Ʊ�׼�޸ļ���ѯ(AQC0080C)"
      End
      Begin VB.Menu Mnu_AQC0090C 
         Caption         =   "���ճɷ�ʵ���޸ļ���ѯ��AQC0090C��"
      End
      Begin VB.Menu Mnu_AQC0092C 
         Caption         =   "���ճɷ�ʵ����ѯ����(AQC0092C)"
      End
      Begin VB.Menu mnu_AQC0310C 
         Caption         =   "��Ʒ�ж������ѯ(AQC0310C)"
      End
      Begin VB.Menu mnu_AQC0360C 
         Caption         =   "�ۺ��ж����ϸ��Ʒ����(AQC0360C)"
      End
      Begin VB.Menu mnu_AQC0110C 
         Caption         =   "��컯��ϵͳ״̬��ѯ(AQC0110C)"
      End
      Begin VB.Menu mnu_AQC0120C 
         Caption         =   "ί�е���Ϣ��ѯ����(AQC0120C)"
      End
      Begin VB.Menu mnu_AQC0130C 
         Caption         =   "δ���ʵ����������״̬��ѯ(AQC0130C)"
      End
      Begin VB.Menu mnu_AQC0140C 
         Caption         =   "����δ��������Ų�ѯ���ս����(AQC0140C)"
      End
      Begin VB.Menu mnu_AQC0150C 
         Caption         =   "�쳣������ѯ(AQC0150C)"
      End
   End
   Begin VB.Menu Mnu_AQD 
      Caption         =   "����֤�������"
      Begin VB.Menu mnu_AQD0050C 
         Caption         =   "��������֤�������(AQD0050C)"
      End
      Begin VB.Menu Mnu_AQD0010C 
         Caption         =   "��ѯ����֤����(AQD0010C)"
      End
      Begin VB.Menu mnu_AQD0030C 
         Caption         =   "����֤������η���(AQD0030C)"
      End
      Begin VB.Menu Mnu_AQD0012C 
         Caption         =   "�ظ�����֤����(AQD0012C)"
      End
      Begin VB.Menu Mnu_AQD0090C 
         Caption         =   "���������͸�ƽ�⿨��ӡ(AQD0090C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQD0091C 
         Caption         =   "���������͸�ƽ�⿨��ӡ(AQD0091C)"
      End
      Begin VB.Menu Mnu_AQD0100C 
         Caption         =   "����/��¯�������߸�ͳ�ƽ���_AQD0100C"
      End
      Begin VB.Menu Mnu_AQD0101C 
         Caption         =   "�����ʱ���-�Զ���������Ų�ѯ_AQD0101C"
      End
      Begin VB.Menu Mnu_AQD0102C 
         Caption         =   "����-�ְ�-�ʱ���-����Ų�ѯ_AQD0102C"
      End
      Begin VB.Menu Mnu_AQD1010C 
         Caption         =   "����֤�������ȷ�Ϲ�������(AQD1010C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQD1 
         Caption         =   "ԭ�ϳɷ�ʵ��"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQD1000C 
         Caption         =   "ԭ�ϳɷ�ʵ������(AQD1000C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQD1030C 
         Caption         =   "�����ʱ�����ƺ����������ѯ_AQD1030C"
      End
      Begin VB.Menu mnu_AQD1040C 
         Caption         =   "��ϵ�ְָ�ͳ�ƽ���_AQD1040C"
      End
   End
   Begin VB.Menu Mnu_AQE 
      Caption         =   "����������Ϣ��ѯ"
      Begin VB.Menu Mnu_AQE0010C 
         Caption         =   "��Ʒ�ɷ������ۺϲ�ѯ(AQE0010C)"
      End
      Begin VB.Menu Mnu_AQE0110C 
         Caption         =   "��ó�ɷ������ۺϲ�ѯ(AQE0110C)"
      End
      Begin VB.Menu Mnu_AQE0011C 
         Caption         =   "��Ʒ���ܸ��в�ѯ(AQE0011C)"
      End
      Begin VB.Menu Mnu_AQE0020C 
         Caption         =   "¼��ɲ���Ŀ��ֵ(AQE0020C)"
      End
      Begin VB.Menu Mnu_AQE0030C 
         Caption         =   "¼��һ�κϸ���Ŀ��ֵ(AQE0030C)"
      End
      Begin VB.Menu mnu_AQE0040C 
         Caption         =   "��Ʒ�������̲�����ѯ(AQE0040C)"
      End
      Begin VB.Menu Mnu_AQE0050C 
         Caption         =   "���ղ���Υ�¼�¼(AQE0050C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQE0060C 
         Caption         =   "���ַ������_AQE0060C"
      End
      Begin VB.Menu mnu_AQE0070C 
         Caption         =   "���ϼ۸�����������_AQE0070C"
      End
      Begin VB.Menu mnu_AQE1000C 
         Caption         =   "��������״�� (AQE1010C)"
      End
      Begin VB.Menu mnu_AQE1020C 
         Caption         =   "��������״�� (AQE1020C)"
      End
      Begin VB.Menu Mnu_AQE1030C 
         Caption         =   "��Ʒ���������ѯ(AQE1030C)"
      End
      Begin VB.Menu Mnu_AQE1050C 
         Caption         =   "�ϸ��к����������������(AQE1050C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQE1051C 
         Caption         =   "�ϸ��к����������������_TEST(AQE1051C)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQE1060C 
         Caption         =   "�ϸ��к������ְ�/�־��������(AQE1060C)"
      End
      Begin VB.Menu Mnu_AQE1062C 
         Caption         =   "�ϸ��к������ְ�/�־��������(AQE1062C)"
      End
      Begin VB.Menu mnu_AQE1070C 
         Caption         =   "�к��/�����ܲ��ϸ�ͳ�Ʊ�_AQE1070C"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQE1080C 
         Caption         =   "�ְ��ȳ����ѯ(AQE1080C)"
      End
      Begin VB.Menu mnu_AQE1090C 
         Caption         =   "��ѯ��Ʒ�ɲ���(AQE1090C)"
      End
      Begin VB.Menu mnu_AQE1100C 
         Caption         =   "��ѯһ�κϸ���(AQE1100C)"
      End
      Begin VB.Menu mnu_AQE1110C 
         Caption         =   "����Ʒ�������ڵ���ҵ��״��ѯ����(AQE1110C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQE1120C 
         Caption         =   "��Ĳ�Ʒ���ʲ��ܼ�(AQE1120C)"
      End
      Begin VB.Menu mnu_AQE1130C 
         Caption         =   "��ĸ��ֺϸ����ܱ�(AQE1130C)"
      End
      Begin VB.Menu mnu_AQE1140C 
         Caption         =   "����������Ϣ�������������ձ�(AQE1140C)"
      End
      Begin VB.Menu mnu_AQE1200C 
         Caption         =   "�ȴ���һ�κϸ���(AQE1200C)"
      End
      Begin VB.Menu Mnu_AQE2000C 
         Caption         =   "��������/�¼���������(AQE2000C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQE2010C 
         Caption         =   "��������˷ϱ�_AQE2010C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQE2020C 
         Caption         =   "����������б�_AQE2020C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQE2030C 
         Caption         =   "��ǰ�����˷ϱ�_AQE2030C"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AQE1121C 
         Caption         =   "������ȡ�����ܲ�ѯ_AQE1121C"
      End
      Begin VB.Menu Mnu_AQE2050C 
         Caption         =   "�г���������̨��_AQE2050C"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AQE2080C 
         Caption         =   "ȡ������ͳ�Ʋ�ѯ_AQE2080C"
      End
      Begin VB.Menu Mnu_AQE2090C 
         Caption         =   "��������ָ���ѯ_AQE2090C"
      End
   End
   Begin VB.Menu Mnu_AQF 
      Caption         =   "����"
      Index           =   1
      Begin VB.Menu Mnu_AQF0020C 
         Caption         =   "����ͱ�ʵ��¼��_AQF0020C"
      End
      Begin VB.Menu Mnu_AQF0030C 
         Caption         =   "��Ʒ�ɷ�ʵ��¼��_AQF0030C"
      End
      Begin VB.Menu Mnu_AQC0095C 
         Caption         =   "�����ͱ���������_AQC0095C"
      End
      Begin VB.Menu Mnu_AQC0093C 
         Caption         =   "�����ͱ�ʵ������_AQC0093C"
      End
      Begin VB.Menu Mnu_AQC0094C 
         Caption         =   "�ͱ���������ѯ_AQC0094C"
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
'        Active_YN = GetSetting("NISCO", "EXE-FILE", "AQ.exe")
'
'        If Active_YN = "1" Then
'
'            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'
'            MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
'            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'
'        Else
'
'            Call Gp_MsgBoxDisplay("ֻ�ܴ��������½��������ϵͳ...", "W")
'            Unload Me
'            Exit Sub
'        End If
'
''        sUserID = "1ZL1005" '"1ZL7209"  '
''        sUserName = "NISCO"
'
'        MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
'        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'
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
''
    args = Split(Trim(Command), " ") ' 2012.11.09 ����  ������
    If UBound(args) = 1 Then
         MainFrmType = "New"
         sUserID = args(0) ' 2012.11.09 ����  ������
         sUserName = args(1) ' 2012.11.09 ����  ������
         MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��" ' 2012.11.09 ����  ������
         MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName ' 2012.11.09 ����  ������
    Else
        Active_YN = GetSetting("NISCO", "EXE-FILE", "AQ.exe")
        If Active_YN = "1" Then
            MainFrmType = "Old"
            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
            MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ����"
            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
        Else
            Call Gp_MsgBoxDisplay("ֻ�ܴ��������½...", "W")
            Unload Me
            Exit Sub
        End If
    End If  ' 2012.11.09 ����  ������

'
'        sUserID = "1ZL1005"
'        sUserName = "����"

'       MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
'       MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName

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
    
        If MsgBox("����δ�رյĲ�������," + vbCrLf + "�Ƿ��˳���ǰϵͳ ?", MB_YESNO _
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
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "AQ.exe", ""

End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
    
    MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
    
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

'Private Sub Mnu_AFH2080C_Click()
'AFH2080C.Show
'    AFH2080C.SetFocus
'End Sub
'
'Private Sub Mnu_AFH2081C_Click()
'AFH2081C.Show
'    AFH2081C.SetFocus
'End Sub

Private Sub Mnu_AQA0010C_Click()
    AQA0010C.Show
    AQA0010C.SetFocus
End Sub


Private Sub Mnu_AQA0130C_Click()
    AQA0130C.Show
    AQA0130C.SetFocus
End Sub
 
Private Sub Mnu_AQA0140C_Click()
    AQA0140C.Show
    AQA0140C.SetFocus
End Sub

Private Sub Mnu_AQA0170C_Click()
    AQA0170C.Show
    AQA0170C.SetFocus
End Sub

Private Sub Mnu_AQA0190C_Click()
    AQA0190C.Show
    AQA0190C.SetFocus
End Sub

Private Sub Mnu_AQA0200C_Click()
    AQA0200C.Show
    AQA0200C.SetFocus
End Sub
Private Sub Mnu_AQA0210C_Click()
    AQA0210C.Show
    AQA0210C.SetFocus
End Sub

Private Sub Mnu_AQA0220C_Click()
    AQA0220C.Show
    AQA0220C.SetFocus
End Sub

Private Sub Mnu_AQA0221C_Click()
    AQA0221C.Show
    AQA0221C.SetFocus
End Sub

Private Sub Mnu_AQA0222C_Click()
    AQA0222C.Show
    AQA0222C.SetFocus
End Sub

Private Sub Mnu_AQA0260C_Click()
    AQA0260C.Show
    AQA0260C.SetFocus
End Sub

Private Sub Mnu_AQA0270C_Click()
    AQA0270C.Show
    AQA0270C.SetFocus
End Sub

Private Sub Mnu_AQA0271C_Click()
    AQA0271C.Show
    AQA0271C.SetFocus
End Sub

Private Sub Mnu_AQA0272C_Click()
    AQA0272C.Show
    AQA0272C.SetFocus
End Sub

Private Sub Mnu_AQA0280C_Click()
    AQA0280C.Show
    AQA0280C.SetFocus
End Sub

Private Sub Mnu_AQA0400C_Click()
    AQA0400C.Show
    AQA0400C.SetFocus
End Sub
Private Sub Mnu_AQA0410C_Click()
    AQA0410C.Show
    AQA0410C.SetFocus
End Sub
Private Sub Mnu_AQA0420C_Click()
    AQA0420C.Show
    AQA0420C.SetFocus
End Sub

Private Sub Mnu_AQA0430C_Click()
    AQA0430C.Show
    AQA0430C.SetFocus
End Sub

Private Sub Mnu_AQA0440C_Click()
    AQA0440C.Show
    AQA0440C.SetFocus
End Sub

Private Sub Mnu_AQA0450C_Click()
    AQA0450C.Show
    AQA0450C.SetFocus
End Sub

Private Sub Mnu_AQA0460C_Click()
    AQA0460C.Show
    AQA0460C.SetFocus
End Sub

Private Sub Mnu_AQA0470C_Click()
    AQA0470C.Show
    AQA0470C.SetFocus
End Sub

Private Sub Mnu_AQA0500C_Click()
    AQA0500C.Show
    AQA0500C.SetFocus
End Sub

Private Sub Mnu_AQA0600C_Click()
    AQA0600C.Show
    AQA0600C.SetFocus
End Sub

Private Sub Mnu_AQA0700C_Click()
    AQA0700C.Show
    AQA0700C.SetFocus
End Sub

Private Sub Mnu_AQA0800C_Click()
    AQA0800C.Show
    AQA0800C.SetFocus
End Sub
Private Sub Mnu_AQA0900C_Click()
    AQA0900C.Show
    AQA0900C.SetFocus
End Sub

Private Sub Mnu_AQA0910C_Click()
    AQA0910C.Show
    AQA0910C.SetFocus
End Sub

Private Sub Mnu_AQA0920C_Click()
    AQA0920C.Show
    AQA0920C.SetFocus
End Sub


Private Sub Mnu_AQB0010C_Click()
    AQB0010C.Show
    AQB0010C.SetFocus
End Sub

Private Sub mnu_AQB0110C_Click()
    AQB0110C.Show
    AQB0110C.SetFocus
End Sub

Private Sub mnu_AQB0120C_Click()
    AQB0120C.Show
    AQB0120C.SetFocus
End Sub
Private Sub mnu_AQB0121C_Click()
    AQB0121C.Show
    AQB0121C.SetFocus
End Sub

Private Sub mnu_AQB0150C_Click()
    AQB0150C.Show
    AQB0150C.SetFocus
End Sub

Private Sub mnu_AQB0160C_Click()
    AQB0160C.Show
    AQB0160C.SetFocus
End Sub


Private Sub mnu_AQB0200C_Click()
    AQB0200C.Show
    AQB0200C.SetFocus
End Sub

Private Sub mnu_AQC0010C_Click()
    AQC0010C.Show
    AQC0010C.SetFocus
End Sub

Private Sub Mnu_AQC0034C_Click()
    AQC0034C.Show
    AQC0034C.SetFocus
End Sub

Private Sub Mnu_AQC0032C_Click()
    AQC0032C.Show
    AQC0032C.SetFocus
End Sub

Private Sub mnu_AQC0040C_Click()
    AQC0040C.Show
    AQC0040C.SetFocus
End Sub

Private Sub mnu_AQC0050C_Click()
    AQC0050C.Show
    AQC0050C.SetFocus
End Sub

Private Sub Mnu_AQC0051C_Click()
    AQC0051C.Show
    AQC0051C.SetFocus
End Sub

Private Sub mnu_AQC0060C_Click()
    AQC0060C.Show
    AQC0060C.SetFocus
End Sub

Private Sub mnu_AQC0070C_Click()
    AQC0070C.Show
    AQC0070C.SetFocus
End Sub

Private Sub mnu_AQC0080C_Click()
    AQC0080C.Show
    AQC0080C.SetFocus
End Sub

Private Sub Mnu_AQC0090C_Click()
    AQC0090C.Show
    AQC0090C.SetFocus
End Sub

Private Sub Mnu_AQC0092C_Click()
    AQC0092C.Show
    AQC0092C.SetFocus
End Sub

Private Sub Mnu_AQC0093C_Click()
 AQC0093C.Show
    AQC0093C.SetFocus
End Sub

Private Sub Mnu_AQC0094C_Click()
 AQC0094C.Show
    AQC0094C.SetFocus
End Sub

Private Sub Mnu_AQC0095C_Click()
     AQC0095C.Show
    AQC0095C.SetFocus
End Sub

Private Sub mnu_AQC0110C_Click()
    AQC0110C.Show
    AQC0110C.SetFocus
End Sub

Private Sub mnu_AQC0120C_Click()
    AQC0120C.Show
    AQC0120C.SetFocus
End Sub

Private Sub mnu_AQC0130C_Click()
    AQC0130C.Show
    AQC0130C.SetFocus
End Sub

Private Sub mnu_AQC0140C_Click()
    AQC0140C.Show
    AQC0140C.SetFocus
End Sub

Private Sub mnu_AQC0150C_Click()
    AQC0150C.Show
    AQC0150C.SetFocus
End Sub

Private Sub mnu_AQC0310C_Click()
    AQC0310C.Show
    AQC0310C.SetFocus
End Sub

Private Sub mnu_AQC0360C_Click()
    AQC0360C.Show
    AQC0360C.SetFocus
End Sub

Private Sub mnu_AQD0010C_Click()
    AQD0010C.Show
    AQD0010C.SetFocus
End Sub

Private Sub Mnu_AQD0012C_Click()
    AQD0012C.Show
    AQD0012C.SetFocus
End Sub

Private Sub mnu_AQD0030C_Click()
    AQD0030C.Show
    AQD0030C.SetFocus
End Sub

Private Sub mnu_AQD0050C_Click()
    AQD0050C.Show
    AQD0050C.SetFocus
End Sub

Private Sub Mnu_AQD0090C_Click()
    AQD0090C.Show
    AQD0090C.SetFocus
End Sub

Private Sub Mnu_AQD0091C_Click()
    AQD0091C.Show
    AQD0091C.SetFocus
End Sub

Private Sub Mnu_AQD0100C_Click()
    AQD0100C.Show
    AQD0100C.SetFocus
End Sub

Private Sub mnu_AQD1000C_Click()
    AQD1000C.Show
    AQD1000C.SetFocus
End Sub

Private Sub Mnu_AQD1010C_Click()
    AQD1010C.Show
    AQD1010C.SetFocus
End Sub

Private Sub Mnu_AQD0101C_Click()
    AQD0101C.Show
    AQD0101C.SetFocus
End Sub

Private Sub Mnu_AQD0102C_Click()
    AQD0102C.Show
    AQD0102C.SetFocus
End Sub

Private Sub mnu_AQD1030C_Click()
    AQD1030C.Show
    AQD1030C.SetFocus
End Sub

Private Sub mnu_AQD1040C_Click()
    AQD1040C.Show
    AQD1040C.SetFocus
End Sub

Private Sub Mnu_AQE0010C_Click()
    AQE0010C.Show
    AQE0010C.SetFocus
End Sub

Private Sub Mnu_AQE0011C_Click()
    AQE0011C.Show
    AQE0011C.SetFocus
End Sub

Private Sub Mnu_AQE0020C_Click()
    AQE0020C.Show
    AQE0020C.SetFocus
End Sub
Private Sub Mnu_AQE0030C_Click()
    AQE0030C.Show
    AQE0030C.SetFocus
End Sub
Private Sub mnu_AQE0040C_Click()
    AQE0040C.Show
    AQE0040C.SetFocus
End Sub

Private Sub Mnu_AQE0050C_Click()
    AQE0050C.Show
    AQE0050C.SetFocus
End Sub

Private Sub mnu_AQE0060C_Click()
    AQE0060C.Show
    AQE0060C.SetFocus
End Sub

Private Sub mnu_AQE0070C_Click()
    AQE0070C.Show
    AQE0070C.SetFocus
End Sub

Private Sub Mnu_AQE0110C_Click()
AQE0110C.Show
    AQE0110C.SetFocus
End Sub

Private Sub mnu_AQE1000C_Click()
    AQE1010C.Show
    AQE1010C.SetFocus
End Sub

Private Sub mnu_AQE1020C_Click()
    AQE1020C.Show
    AQE1020C.SetFocus
End Sub

Private Sub Mnu_AQE1030C_Click()
    AQE1030C.Show
    AQE1030C.SetFocus
End Sub

Private Sub Mnu_AQE1050C_Click()
    AQE1050C.Show
    AQE1050C.SetFocus
End Sub
Private Sub Mnu_AQE1051C_Click()
    AQE1051C.Show
    AQE1051C.SetFocus
End Sub
Private Sub Mnu_AQE1060C_Click()
    AQE1060C.Show
    AQE1060C.SetFocus
End Sub

Private Sub Mnu_AQE1062C_Click()
    AQE1062C.Show
    AQE1062C.SetFocus
End Sub

Private Sub Mnu_AQE1070C_Click()
    AQE1070C.Show
    AQE1070C.SetFocus
End Sub
Private Sub Mnu_AQE1080C_Click()
    AQE1080C.Show
    AQE1080C.SetFocus
End Sub

Private Sub mnu_AQE1090C_Click()
    AQE1090C.Show
    AQE1090C.SetFocus
    End Sub
Private Sub mnu_AQE1100C_Click()
    AQE1100C.Show
    AQE1100C.SetFocus
End Sub

Private Sub mnu_AQE1110C_Click()
    AQE1110C.Show
    AQE1110C.SetFocus
End Sub

Private Sub mnu_AQE1120C_Click()
    AQE1120C.Show
    AQE1120C.SetFocus
End Sub

Private Sub mnu_AQE1121C_Click()
    AQE1121C.Show
    AQE1121C.SetFocus
End Sub


Private Sub mnu_AQE1130C_Click()
    AQE1130C.Show
    AQE1130C.SetFocus
End Sub

Private Sub mnu_AQE1140C_Click()

    AQE1140C.Show
    AQE1140C.SetFocus

End Sub

Private Sub mnu_AQE1200C_Click()
    AQE1200C.Show
    AQE1200C.SetFocus
End Sub

Private Sub Mnu_AQE2000C_Click()
    AQE2000C.Show
    AQE2000C.SetFocus
End Sub

Private Sub mnu_AQE2010C_Click()
    AQE2010C.Show
    AQE2010C.SetFocus
End Sub
Private Sub mnu_AQE2020C_Click()
    AQE2020C.Show
    AQE2020C.SetFocus
End Sub

Private Sub mnu_AQE2030C_Click()
    AQE2030C.Show
    AQE2030C.SetFocus
End Sub


Private Sub Mnu_AQE2050C_Click()
    AQE2050C.Show
    AQE2050C.SetFocus
End Sub

Private Sub Mnu_AQE2080C_Click()
    AQE2080C.Show
    AQE2080C.SetFocus
End Sub

Private Sub Mnu_AQE2090C_Click()
    AQE2090C.Show
    AQE2090C.SetFocus
End Sub

Private Sub Mnu_AQF0010C_Click()
    AQF0010C.Show
    AQF0010C.SetFocus
End Sub

Private Sub Mnu_AQF0020C_Click()
    AQF0020C.Show
    AQF0020C.SetFocus
End Sub

Private Sub Mnu_AQF0030C_Click()
    AQF0030C.Show
    AQF0030C.SetFocus
End Sub

Private Sub Mnu_AQF0100C_Click()
    AQF0100C.Show
    AQF0100C.SetFocus
End Sub


Private Sub Mnu_Cascade_Click()
    MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
    MDIMain.Arrange 0
End Sub

Private Sub Mnu_Acopy_Click()
    'Screen All Copy
    Call ActiveForm.form_Cpy
    Call MDIMain.FormMenuSetting(Me, "", "Acopy", "")
End Sub

Private Sub Mnu_Apaste_Click()
    'Screen All Paste
    Call ActiveForm.form_Pst
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
    MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
    Call ActiveForm.Spread_Forzens_Cancel
End Sub

Private Sub Mnu_FrozenSetting_Click()
    'Spread Col Frozens Setting
    MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
    Call ActiveForm.Spread_Forzens_Setting
End Sub

Private Sub Mnu_Horiz_Click()
    MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
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
    MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
    Call ActiveForm.Spread_ColumnsSort
End Sub

Private Sub Mnu_Spaste_Click()
    'Spread Row Paste
    Call ActiveForm.Spread_Pst
End Sub

Private Sub Mnu_Vertical_Click()
    MDIMain.StatusBar1.Panels(1) = "ϵͳ��Ϣ : "
    MDIMain.Arrange 2
End Sub