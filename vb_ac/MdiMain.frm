VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "���̹���"
   ClientHeight    =   7305
   ClientLeft      =   780
   ClientTop       =   2805
   ClientWidth     =   11250
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Tag             =   "C"
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet 
      Left            =   90
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
      ScaleWidth      =   11190
      TabIndex        =   0
      Top             =   0
      Width           =   11250
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   600
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   20250
         _ExtentX        =   35719
         _ExtentY        =   1058
         BandCount       =   1
         _CBWidth        =   20250
         _CBHeight       =   600
         _Version        =   "6.7.9782"
         Child1          =   "MenuTool"
         MinHeight1      =   540
         Width1          =   20190
         NewRow1         =   0   'False
         BandStyle1      =   1
         Begin MSComctlLib.Toolbar MenuTool 
            Height          =   540
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   20190
            _ExtentX        =   35613
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
      Left            =   90
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
      Left            =   90
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
      Top             =   6840
      Width           =   11250
      _ExtentX        =   19844
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
            TextSave        =   "2017-01-17"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "10:30"
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
   Begin VB.Menu Mnu_ACA 
      Caption         =   "��������"
      Begin VB.Menu Mnu_ACA1020C 
         Caption         =   "����������״��ѯ"
      End
      Begin VB.Menu Mnu_ACA1030C 
         Caption         =   "����������ϸ��ѯ"
      End
      Begin VB.Menu Mnu_ACA1031C 
         Caption         =   "����������ϸ��ѯ"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACA1040C 
         Caption         =   "���Ͻ�����״��ѯ"
      End
      Begin VB.Menu Mnu_ACA1021C 
         Caption         =   "��ͬ���ֲ�ѯ"
      End
      Begin VB.Menu Mnu_ACA1120C 
         Caption         =   "��ʷ�������̲�ѯ_ACA1120C"
      End
      Begin VB.Menu Mnu_ACA1130C 
         Caption         =   "��ͬ���ַ����ۺϱ���_ACA1130C"
      End
      Begin VB.Menu Mnu_ACA1140C 
         Caption         =   "���Ϲ���ʱ���ѯ_ACA1140C"
      End
      Begin VB.Menu Mnu_ACA2010C 
         Caption         =   "�������Ͷ��������"
      End
      Begin VB.Menu Mnu_ACA2030C 
         Caption         =   "�������Ͷ����Ϣ��ѯ��ACA2030C"
      End
      Begin VB.Menu Mnu_ACA2033C 
         Caption         =   "������汨����ACA2033C"
      End
      Begin VB.Menu Mnu_ACA2034C 
         Caption         =   "�¶�������ͳ�Ʊ��������֣�"
      End
      Begin VB.Menu Mnu_ACA1045C 
         Caption         =   "������Ϣ�����ϵͳ��ϵ��ѯ"
      End
   End
   Begin VB.Menu Mnu_ACB 
      Caption         =   "���Ϲ���"
      Begin VB.Menu Mnu_ACB1010C 
         Caption         =   "���Ͽ���ܼƲ�ѯ"
      End
      Begin VB.Menu Mnu_ACB1020C 
         Caption         =   "���Ͽ����״��ѯ"
      End
      Begin VB.Menu Mnu_ACB1030C 
         Caption         =   "����״����ϸ��ѯ"
      End
      Begin VB.Menu Mnu_ACB1022C 
         Caption         =   "����״̬��Ϣ��ѯ"
      End
      Begin VB.Menu Mnu_ACB1024C 
         Caption         =   "������Ϣ���������ѯ"
      End
      Begin VB.Menu Mnu_ACB4090C 
         Caption         =   "��Ʒ���д���"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACB4120C 
         Caption         =   "�ϸ����д���"
      End
      Begin VB.Menu Mnu_ACB4091C 
         Caption         =   "����ERP��Ʒ�����������"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACB4100C 
         Caption         =   "���϶�����Ϣ�������"
      End
      Begin VB.Menu Mnu_ACB4110C 
         Caption         =   "������ҵ�����ѯ"
      End
      Begin VB.Menu Mnu_CGD2070C 
         Caption         =   "¼�뾫����ҵָʾ"
      End
      Begin VB.Menu Mnu_ACB4080C 
         Caption         =   "������Ϣ��ѯ�޸�"
      End
      Begin VB.Menu Mnu_ACB4098C 
         Caption         =   "�����ߴ���Ϣת������_ACB4098C"
      End
      Begin VB.Menu Mnu_ACB4099C 
         Caption         =   "�����ߴ���Ϣ��ѯ�޸�"
      End
      Begin VB.Menu Mnu_ACB1023C 
         Caption         =   "�������ָ��в�ѯ���޸�"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACB4070C 
         Caption         =   "��������/�ж�ʵ��¼��"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACB4060C 
         Caption         =   "��������������ת������"
      End
      Begin VB.Menu Mnu_ACB4200C 
         Caption         =   "���������ж�����"
      End
      Begin VB.Menu Mnu_ACE5010C 
         Caption         =   "ָ�������ί�мӹ�"
      End
      Begin VB.Menu Mnu_ACB1040C 
         Caption         =   "δ�����ж��Ĳ�Ʒ����ѯ"
      End
      Begin VB.Menu Mnu_ACB3010C 
         Caption         =   "�ֻ���Ϣ��ѯ"
      End
      Begin VB.Menu Mnu_ACB4400C 
         Caption         =   "�����������"
         Begin VB.Menu Mnu_ACB4150C 
            Caption         =   "¼���������"
         End
         Begin VB.Menu Mnu_ACB4140C 
            Caption         =   "¼�����������"
         End
         Begin VB.Menu Mnu_ACB4160C 
            Caption         =   "ȷ��������"
         End
         Begin VB.Menu Mnu_ACB4170C 
            Caption         =   "����������ѯ"
         End
      End
      Begin VB.Menu Mnu_ACB4000C 
         Caption         =   "��Ʒת����ҵ����"
         Begin VB.Menu Mnu_ACB4010C 
            Caption         =   "ת����ҵָʾ¼��"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_ACB4020C 
            Caption         =   "��Ʒװ��ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB4030C 
            Caption         =   "��Ʒж��ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB5070C 
            Caption         =   "��Ʒ������Ϣ��ѯ���޸�"
         End
         Begin VB.Menu Mnu_ACB4031C 
            Caption         =   "ת����ҵʵ����ѯ"
         End
         Begin VB.Menu Mnu_ACB5080C 
            Caption         =   "��λ�ű��������ѯ-ACB5080C"
         End
      End
      Begin VB.Menu Line3 
         Caption         =   "����ת����ҵ����"
         Begin VB.Menu Mnu_ACB6020C 
            Caption         =   "����ת��ƻ�ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB6030C 
            Caption         =   "����ת��ƻ���ѯ"
         End
         Begin VB.Menu Mnu_ACB1025C 
            Caption         =   "����װ��ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB1026C 
            Caption         =   "����ж��ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB6060C 
            Caption         =   "���������޸ļ���ѯ"
         End
         Begin VB.Menu Mnu_ACB6070C 
            Caption         =   "�������Ͽ��ͼ����"
         End
         Begin VB.Menu Mnu_ACB4040C 
            Caption         =   "����ʹ��ʵ��¼��"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_ACB4050C 
            Caption         =   "����ʹ��ʵ����ѯ"
         End
      End
      Begin VB.Menu Mnu_Slab 
         Caption         =   "�⹺��������"
         Begin VB.Menu Mnu_ACB2010C 
            Caption         =   "�⹺���ܼƲ�ѯ"
         End
         Begin VB.Menu Mnu_ACB2020C 
            Caption         =   "�⹺����״��ѯ"
         End
      End
      Begin VB.Menu Mnu_PlateForArea 
         Caption         =   "���ڲ�ͬ������������"
         Begin VB.Menu Mnu_ACB5026C 
            Caption         =   "���ڲ�ͬ����װ��ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB5036C 
            Caption         =   "���ڲ�ͬ����ж��ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB5031C 
            Caption         =   "���ڲ�ͬ����ת��ʵ����ѯ"
         End
      End
      Begin VB.Menu Mnu_PlateForInv 
         Caption         =   "���Ʒת����ҵ����"
         Begin VB.Menu Mnu_ACB5025C 
            Caption         =   "���Ʒװ��ʵ��¼��"
         End
         Begin VB.Menu Mnu_ACB5035C 
            Caption         =   "���Ʒж��ʵ��¼��"
         End
      End
   End
   Begin VB.Menu Mnu_ACE 
      Caption         =   "�������"
      Begin VB.Menu Mnu_ACE1280C 
         Caption         =   "������¼��"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACE2000C 
         Caption         =   "��Ľ���¼��"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACE3010C 
         Caption         =   "��;���"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACE1 
         Caption         =   "�Զ��������"
         Begin VB.Menu Mnu_ACE1010C 
            Caption         =   "���������ѡ��"
         End
         Begin VB.Menu Mnu_ACE1030C 
            Caption         =   "��������ѡ��"
         End
         Begin VB.Menu Mnu_ACE1065C 
            Caption         =   "�������"
         End
         Begin VB.Menu Mnu_ACE1260C 
            Caption         =   "HMI�����������"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_ACE1270C 
            Caption         =   "HMI�����������"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_ACE1150C 
            Caption         =   "HMI��İ������ PLATE"
         End
         Begin VB.Menu Mnu_ACE1200C 
            Caption         =   "��������ѯ���޸�"
         End
      End
      Begin VB.Menu Mnu_ACE7 
         Caption         =   "HMI�������"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ACE8 
         Caption         =   "����Ʒ�������"
         Begin VB.Menu Mnu_ACE6010C 
            Caption         =   "����Ʒ�������(���ϱ�׼)"
         End
         Begin VB.Menu Mnu_ACE6020C 
            Caption         =   "����Ʒ�������(������׼)"
         End
         Begin VB.Menu Mnu_ACE6030C 
            Caption         =   "����Ʒ��Ľ���(������׼)"
         End
         Begin VB.Menu Mnu_ACE6040C 
            Caption         =   "����Ʒ����ְ崦��_ACE6040C"
         End
      End
      Begin VB.Menu Mnu_ACE9 
         Caption         =   "����Ʒ�Զ����"
         Begin VB.Menu Mnu_ACE7000C 
            Caption         =   "¼������Ʒ�Զ������׼_ACE7000C"
         End
         Begin VB.Menu Mnu_ACE7010C 
            Caption         =   "����Ʒ�Զ����ȷ��(���ϱ�׼)_ACE7010C"
         End
         Begin VB.Menu Mnu_ACE7020C 
            Caption         =   "����Ʒ�Զ����ȷ��(������׼)_ACE7020C"
         End
      End
      Begin VB.Menu Mnu_ACE1209C 
         Caption         =   "�������������ѯ"
      End
      Begin VB.Menu Mnu_ACE4010C 
         Caption         =   "�����׼¼��"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_ACF 
      Caption         =   "ͳ�ƹ���"
      Visible         =   0   'False
      Begin VB.Menu Mnu_AKP1011C 
         Caption         =   "�����򱨲�ѯ"
      End
      Begin VB.Menu Mnu_ACF2 
         Caption         =   "�����±���ѯ"
         Begin VB.Menu Mnu_ACF2080C 
            Caption         =   "���������±���ѯ"
         End
         Begin VB.Menu Mnu_ACF2090C 
            Caption         =   "���������±���ѯ"
         End
      End
   End
   Begin VB.Menu Mnu_ACF1 
      Caption         =   "�������"
      Begin VB.Menu Mnu_ACF0010C 
         Caption         =   "���嶨�����ͽ��ȸ���_ACF0010C"
      End
      Begin VB.Menu Mnu_ACF0020C 
         Caption         =   "���嶨�����Ͷ������ٲ�ѯ_ACF0020C"
      End
      Begin VB.Menu Mnu_ACF0030C 
         Caption         =   "���嶨�����͹���ʱ�����_ACF0030C"
      End
      Begin VB.Menu Mnu_ACF0040C 
         Caption         =   "������׼����ʱ�����_ACF0040C"
      End
      Begin VB.Menu Mnu_ACF0050C 
         Caption         =   "������_ACF0050C"
      End
      Begin VB.Menu Mnu_ACF0060C 
         Caption         =   "��Ķ������ܱ�_ACF0060C"
      End
      Begin VB.Menu Mnu_ACF0070C 
         Caption         =   "�����ҵ��������Ӫ�ձ�_ACF0070C"
      End
      Begin VB.Menu Mnu_ACF0080C 
         Caption         =   "�����ɱ����ݸ���_ACF0080C"
      End
      Begin VB.Menu Mnu_ACF0090C 
         Caption         =   "�����ɱ����ݸ�����ϸ_ACF0090C"
      End
      Begin VB.Menu Mnu_ACF0091C 
         Caption         =   "���ֳɱ����ݸ�����ϸ_ACF0091C"
      End
   End
   Begin VB.Menu Mnu_ACZ 
      Caption         =   "��������"
      Begin VB.Menu Mnu_ACB4121C 
         Caption         =   "�ְ�ת���ϴ�������"
      End
      Begin VB.Menu Mnu_ACA2035C 
         Caption         =   "����������Ϣ��ϸ��ѯ����_ACA2035C"
      End
      Begin VB.Menu Mnu_ACA2032C 
         Caption         =   "QAB״̬δ��ʱ�����ְ�ͳ�Ʊ�"
      End
      Begin VB.Menu Mnu_AEZ2010C 
         Caption         =   "��ѯ��ع���������Ϣ"
      End
      Begin VB.Menu Mnu_ACZ1010C 
         Caption         =   "�ۺϲ�ѯ����"
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
'        Active_YN = GetSetting("NISCO", "EXE-FILE", "AC.exe")
'
'        If Active_YN = "1" Then
'            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'            MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
'            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'        Else
'            Call Gp_MsgBoxDisplay("ֻ�ܴ��������½...", "W")
'            Unload Me
'            Exit Sub
'        End If
'
''        sUserID = "1JS6005"
''        sUserName = "��ɺ�"
''        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
''        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
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

        args = Split(Trim(Command), " ") ' 2012.11.09 ����  ������
'        If UBound(args) = 1 Then
'             MainFrmType = "New"
'             sUserID = args(0) ' 2012.11.09 ����  ������
'             sUserName = args(1) ' 2012.11.09 ����  ������
'             MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��" ' 2012.11.09 ����  ������
'             MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName ' 2012.11.09 ����  ������
'        Else
'            Active_YN = GetSetting("NISCO", "EXE-FILE", "AC.exe")
'            If Active_YN = "1" Then
'                MainFrmType = "Old"
'                sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'                sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'                MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
'                MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'            Else
'                Call Gp_MsgBoxDisplay("ֻ�ܴ��������½...", "W")
'                Unload Me
'                Exit Sub
'            End If
'        End If  ' 2012.11.09 ����  ������
''
        sUserID = "1JS1005"  '1JS1005
        sUserName = "����"
        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName


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
    
        'If Gf_MessConfirm("Low rank program was not ended," + vbCrLf + "end Program ?", "Q", Me.Caption) Then
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
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "AC.exe", ""
    
End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
    
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
    
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

    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
    
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

Private Sub Mnu_ACA1020C_Click()
    ACA1020C.Show
    ACA1020C.SetFocus
End Sub

Private Sub Mnu_ACA1021C_Click()
    ACA1021C.Show
    ACA1021C.SetFocus
End Sub

Private Sub Mnu_ACA1030C_Click()
    ACA1030C.Show
    ACA1030C.SetFocus
End Sub

Private Sub Mnu_ACA1031C_Click()
ACA1031C.Show
ACA1031C.SetFocus
End Sub

Private Sub Mnu_ACA1040C_Click()
    ACA1040C.Show
    ACA1040C.SetFocus
End Sub

Private Sub Mnu_ACA1045C_Click()
    ACA1045C.Show
    ACA1045C.SetFocus
End Sub

Private Sub Mnu_ACA1120C_Click()
    ACA1120C.Show
    ACA1120C.SetFocus
End Sub

Private Sub Mnu_ACA1130C_Click()
    ACA1130C.Show
    ACA1130C.SetFocus
End Sub

Private Sub Mnu_ACA1140C_Click()
    ACA1140C.Show
    ACA1140C.SetFocus
End Sub





Private Sub Mnu_ACA2010C_Click()
    ACA2010C.Show
    ACA2010C.SetFocus
End Sub

Private Sub Mnu_ACA2030C_Click()
    ACA2030C.Show
    ACA2030C.SetFocus
End Sub

Private Sub Mnu_ACA2032C_Click()
    ACA2032C.Show
    ACA2032C.SetFocus
End Sub
Private Sub Mnu_ACA2033C_Click()
    ACA2033C.Show
    ACA2033C.SetFocus
End Sub

Private Sub Mnu_ACA2034C_Click()
    ACA2034C.Show
    ACA2034C.SetFocus
End Sub

Private Sub Mnu_ACA2035C_Click()
    ACA2035C.Show
    ACA2035C.SetFocus
End Sub

Private Sub Mnu_ACB1010C_Click()
    ACB1010C.Show
    ACB1010C.SetFocus
End Sub

Private Sub Mnu_ACB1020C_Click()
    ACB1020C.Show
    ACB1020C.SetFocus
End Sub

Private Sub Mnu_ACB1022C_Click()
    ACB1022C.Show
    ACB1022C.SetFocus
End Sub

Private Sub Mnu_ACB1023C_Click()
    ACB1023C.Show
    ACB1023C.SetFocus
End Sub

Private Sub Mnu_ACB1024C_Click()
    ACB1024C.Show
    ACB1024C.SetFocus
End Sub

Private Sub Mnu_ACB1025C_Click()
    ACB6040C.Show
    ACB6040C.SetFocus
End Sub

Private Sub Mnu_ACB1026C_Click()
    ACB6050C.Show
    ACB6050C.SetFocus
End Sub

Private Sub Mnu_ACB1030C_Click()
    ACB1030C.Show
    ACB1030C.SetFocus
End Sub

Private Sub Mnu_ACB1040C_Click()
    ACB1040C.Show
    ACB1040C.SetFocus
End Sub

'Private Sub Mnu_ACB1050C_Click()
'    ACB1050C.Show
'    ACB1050C.SetFocus
'End Sub

'Private Sub Mnu_ACB1060C_Click()
'    ACB1060C.Show
'    ACB1060C.SetFocus
'End Sub

Private Sub Mnu_ACB2010C_Click()
    ACB2010C.Show
    ACB2010C.SetFocus
End Sub

Private Sub Mnu_ACB2020C_Click()
    ACB2020C.Show
    ACB2020C.SetFocus
End Sub

Private Sub Mnu_ACB3010C_Click()
    ACB3010C.Show
    ACB3010C.SetFocus
End Sub

Private Sub Mnu_ACB4010C_Click()
    ACB4020C.Show
    ACB4020C.SetFocus
End Sub

Private Sub Mnu_ACB4020C_Click()
    ACB5020C.Show
    ACB5020C.SetFocus
End Sub

Private Sub Mnu_ACB4030C_Click()
    ACB5030C.Show
    ACB5030C.SetFocus
End Sub

Private Sub Mnu_ACB4098C_Click()
    ACB4098C.Show
    ACB4098C.SetFocus
End Sub

Private Sub Mnu_ACB4099C_Click()
    ACB4099C.Show
    ACB4099C.SetFocus
End Sub

Private Sub Mnu_ACB4121C_Click()
    ACB4121C.Show
    ACB4121C.SetFocus
End Sub

Private Sub Mnu_ACB5026C_Click()
    ACB5026C.Show
    ACB5026C.SetFocus
End Sub

Private Sub Mnu_ACB5036C_Click()
    ACB5036C.Show
    ACB5036C.SetFocus
End Sub

Private Sub Mnu_ACB5025C_Click()
    ACB5025C.Show
    ACB5025C.SetFocus
End Sub

Private Sub Mnu_ACB5035C_Click()
    ACB5035C.Show
    ACB5035C.SetFocus
End Sub


Private Sub Mnu_ACB5031C_Click()
    ACB5031C.Show
    ACB5031C.SetFocus
End Sub


Private Sub Mnu_ACB4031C_Click()
    ACB4031C.Show
    ACB4031C.SetFocus
End Sub

Private Sub Mnu_ACB4040C_Click()
'    ACB4040C.Show
'    ACB4040C.SetFocus
End Sub

Private Sub Mnu_ACB4050C_Click()
    ACB4050C.Show
    ACB4050C.SetFocus
End Sub

Private Sub Mnu_ACB4060C_Click()
    ACB4060C.Show
    ACB4060C.SetFocus
End Sub

Private Sub Mnu_ACB4070C_Click()
    ACB4070C.Show
    ACB4070C.SetFocus
End Sub

Private Sub Mnu_ACB4080C_Click()
    ACB4080C.Show
    ACB4080C.SetFocus
End Sub

Private Sub Mnu_ACB4090C_Click()
'    ACB4090C.Show
'    ACB4090C.SetFocus
End Sub

Private Sub Mnu_ACB4091C_Click()
'    ACB4091C.Show
'    ACB4091C.SetFocus
End Sub

Private Sub Mnu_ACB4100C_Click()
    ACB4100C.Show
    ACB4100C.SetFocus
End Sub

Private Sub Mnu_ACB4110C_Click()
    ACB4110C.Show
    ACB4110C.SetFocus
End Sub

Private Sub Mnu_ACB4120C_Click()
    ACB4120C.Show
    ACB4120C.SetFocus
End Sub

Private Sub Mnu_ACB4140C_Click()
    ACB4140C.Show
    ACB4140C.SetFocus
End Sub

Private Sub Mnu_ACB4150C_Click()
    ACB4150C.Show
    ACB4150C.SetFocus
End Sub

Private Sub Mnu_ACB4160C_Click()
    ACB4160C.Show
    ACB4160C.SetFocus
End Sub

Private Sub Mnu_ACB4170C_Click()
    ACB4170C.Show
    ACB4170C.SetFocus
End Sub

Private Sub Mnu_ACB4200C_Click()
    ACB4200C.Show
    ACB4200C.SetFocus
End Sub

Private Sub Mnu_ACB5070C_Click()
    ACB5070C.Show
    ACB5070C.SetFocus
End Sub
Private Sub Mnu_ACB5080C_Click()
    ACB5080C.Show
    ACB5080C.SetFocus
End Sub

Private Sub Mnu_ACB6020C_Click()
    ACB6020C.Show
    ACB6020C.SetFocus
End Sub

Private Sub Mnu_ACB6030C_Click()
    ACB6030C.Show
    ACB6030C.SetFocus
End Sub

Private Sub Mnu_ACB6060C_Click()
    ACB6060C.Show
    ACB6060C.SetFocus
End Sub

Private Sub Mnu_ACB6070C_Click()
    ACB6070C.Show
    ACB6070C.SetFocus
End Sub

Private Sub Mnu_ACE1010C_Click()
    ACE1010C.Show
    ACE1010C.SetFocus
End Sub
Private Sub Mnu_ACE1030C_Click()
    ACE1030C.Show
    ACE1030C.SetFocus
End Sub
'Private Sub Mnu_ACE1040C_Click()
'    ACE1040C.Show
'    ACE1040C.SetFocus
'End Sub

Private Sub Mnu_ACE1065C_Click()
    ACE1065C.Show
    ACE1065C.SetFocus
End Sub

Private Sub Mnu_ACE1150C_Click()
    ACE1150C.Show
    ACE1150C.SetFocus
End Sub

Private Sub Mnu_ACE1200C_Click()
    ACE1200C.Show
    ACE1200C.SetFocus
End Sub

Private Sub Mnu_ACE1209C_Click()
    ACE1209C.Show
    ACE1209C.SetFocus
End Sub

Private Sub Mnu_ACE1260C_Click()
'    ACE1260C.Show
'    ACE1260C.SetFocus
End Sub

Private Sub Mnu_ACE1270C_Click()
'    ACE1270C.Show
'    ACE1270C.SetFocus
End Sub

Private Sub Mnu_ACE1280C_Click()
'    ACE1280C.Show
'    ACE1280C.SetFocus
End Sub

Private Sub Mnu_ACE2000C_Click()
'    ACE2000C.Show
'    ACE2000C.SetFocus
End Sub

Private Sub Mnu_ACE3010C_Click()
'    ACE3010C.Show
'    ACE3010C.SetFocus
End Sub

Private Sub Mnu_ACE4010C_Click()
'    ACE4010C.Show
'    ACE4010C.SetFocus
End Sub

Private Sub Mnu_ACE5010C_Click()
    ACE5010C.Show
    ACE5010C.SetFocus
End Sub

Private Sub Mnu_ACE6010C_Click()
    ACE6010C.Show
    ACE6010C.SetFocus
End Sub

Private Sub Mnu_ACE6020C_Click()
    ACE6020C.Show
    ACE6020C.SetFocus
End Sub

Private Sub Mnu_ACE6030C_Click()
    ACE6030C.Show
    ACE6030C.SetFocus
End Sub

Private Sub Mnu_ACE6040C_Click()
    ACE6040C.Show
    ACE6040C.SetFocus
End Sub

Private Sub Mnu_ACE7000C_Click()
    ACE7000C.Show
    ACE7000C.SetFocus
End Sub

Private Sub Mnu_ACE7010C_Click()
    ACE7010C.Show
    ACE7010C.SetFocus
End Sub

Private Sub Mnu_ACE7020C_Click()
    ACE7020C.Show
    ACE7020C.SetFocus
End Sub

Private Sub Mnu_ACF0010C_Click()
    ACF0010C.Show
    ACF0010C.SetFocus
End Sub

Private Sub Mnu_ACF0020C_Click()
    ACF0020C.Show
    ACF0020C.SetFocus
End Sub

Private Sub Mnu_ACF0030C_Click()
    ACF0030C.Show
    ACF0030C.SetFocus
End Sub

Private Sub Mnu_ACF0040C_Click()
    ACF0040C.Show
    ACF0040C.SetFocus
End Sub

Private Sub Mnu_ACF0050C_Click()
    ACF0051C.Show
    ACF0051C.SetFocus
End Sub

Private Sub Mnu_ACF0060C_Click()
    ACF0060C.Show
    ACF0060C.SetFocus
End Sub

Private Sub Mnu_ACF0070C_Click()
    ACF0070C.Show
    ACF0070C.SetFocus
End Sub

Private Sub Mnu_ACF0080C_Click()
    ACF0080C.Show
    ACF0080C.SetFocus
End Sub

Private Sub Mnu_ACF0090C_Click()
    ACF0090C.Show
    ACF0090C.SetFocus
End Sub

Private Sub Mnu_ACF0091C_Click()
    ACF0091C.Show
    ACF0091C.SetFocus
End Sub

Private Sub Mnu_ACF2080C_Click()
    ACF2080C.Show
    ACF2080C.SetFocus
End Sub

Private Sub Mnu_ACF2090C_Click()
    ACF2090C.Show
    ACF2090C.SetFocus
End Sub

Private Sub Mnu_ACZ1010C_Click()
    ACZ1011C.Show
    ACZ1011C.SetFocus
End Sub

Private Sub Mnu_AEZ2010C_Click()
    AEZ2010C.Show
    AEZ2010C.SetFocus
End Sub
'
'Private Sub Mnu_AKP1011C_Click()
'    AKP1011C.Show
'    AKP1011C.SetFocus
'End Sub

Private Sub Mnu_Cascade_Click()
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
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

Private Sub Mnu_CGD2070C_Click()
    CGD2070C.Show
    CGD2070C.SetFocus
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
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
    Call ActiveForm.Spread_Forzens_Cancel
End Sub

Private Sub Mnu_FrozenSetting_Click()
    'Spread Col Frozens Setting
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
    Call ActiveForm.Spread_Forzens_Setting
End Sub

' Display Help Screen
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
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
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
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
    Call ActiveForm.Spread_ColumnsSort
End Sub

Private Sub Mnu_Spaste_Click()
    'Spread Row Paste
    Call ActiveForm.Spread_Pst
End Sub

Private Sub Mnu_Vertical_Click()
    MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
    MDIMain.Arrange 2
End Sub
