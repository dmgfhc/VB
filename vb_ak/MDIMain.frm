VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "������������"
   ClientHeight    =   7545
   ClientLeft      =   345
   ClientTop       =   4260
   ClientWidth     =   12150
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Tag             =   "K"
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet 
      Left            =   30
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12090
      TabIndex        =   0
      Top             =   0
      Width           =   12150
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
            Picture         =   "MDIMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":121F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":12FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1508
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":16CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1888
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1ACD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1C05
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2196
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
            Picture         =   "MDIMain.frx":24A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2960
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2C63
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2F83
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":316C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":32BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":3405
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":3592
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":367C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":396B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":3A77
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":3D4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   7080
      Width           =   12150
      _ExtentX        =   21431
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
            TextSave        =   "2016-07-13"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "15:53"
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
            Picture         =   "MDIMain.frx":41FE
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
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_FrozenSetting 
         Caption         =   "�ж�������"
      End
      Begin VB.Menu Mnu_FrozenCancel 
         Caption         =   "��ȡ������"
      End
   End
   Begin VB.Menu Mnu_AKA 
      Caption         =   "������ҵ"
      Begin VB.Menu Mnu_AKN2030C 
         Caption         =   "������ҵָʾ�������´����"
      End
      Begin VB.Menu Mnu_AKN2040C 
         Caption         =   "����ָʾ�������´����"
      End
      Begin VB.Menu Mnu_AKN2050C 
         Caption         =   "ָ�������������ҵָʾ��������"
      End
      Begin VB.Menu Mnu_AKN2010C 
         Caption         =   "������ҵָʾ��ѯ����"
      End
      Begin VB.Menu Mnu_AKN2020C 
         Caption         =   "�����и���ҵָʾ��ѯ����"
      End
      Begin VB.Menu Mnu_AKO2010C 
         Caption         =   "���������ڽ��̸��ٽ���"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKL2050C 
         Caption         =   "�����и�/װ¯��״����"
      End
      Begin VB.Menu Mnu_AKM2030C 
         Caption         =   "���͸�ˮʵ���޸ļ���ѯ����"
      End
      Begin VB.Menu Mnu_AKA2000C 
         Caption         =   "��¯��ˮʵ������"
      End
      Begin VB.Menu Mnu_AKN3000C 
         Caption         =   "�����ƻ�ָʾ�ָ�����"
      End
      Begin VB.Menu Mnu_AKE1010C 
         Caption         =   "ת¯���ղ�����ѯ����"
      End
      Begin VB.Menu Mnu_AKE1020C 
         Caption         =   "�������ղ�����ѯ����"
      End
      Begin VB.Menu Mnu_AKE1030C 
         Caption         =   "�������ղ�����ѯ����"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKN2031C 
         Caption         =   "������ҵָʾ��������"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_AKP 
      Caption         =   "����ͳ�Ʋ�ѯ"
      Begin VB.Menu Mnu_AKP1014C 
         Caption         =   "����������ѯ"
      End
      Begin VB.Menu Mnu_AKP1010C 
         Caption         =   "����ҵ�ƻ���ѯ���޸�"
      End
      Begin VB.Menu Mnu_AKP1011C 
         Caption         =   "�к�����������(����)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKP1111C 
         Caption         =   "�к�����������(��������)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKP1211C 
         Caption         =   "�к�����������"
      End
      Begin VB.Menu Mnu_AKP1213C 
         Caption         =   "�к�����������_V1010"
      End
      Begin VB.Menu Mnu_AKP1022C 
         Caption         =   "���ֺϽ𼰸���������ϸ��Ϣ"
      End
      Begin VB.Menu Mnu_AKP1013C 
         Caption         =   "ͣ��ʵ���ۺϲ�ѯ"
      End
      Begin VB.Menu Mnu_AKT 
         Caption         =   "�ɱ����"
         Begin VB.Menu Mnu_AKP1016C 
            Caption         =   "�����ճɱ�����(��¯��)"
         End
         Begin VB.Menu Mnu_AKP1019C 
            Caption         =   "�����ճɱ�����(������)"
         End
         Begin VB.Menu Mnu_AKP1018C 
            Caption         =   "�����ճɱ�����(��¯��)"
         End
         Begin VB.Menu Mnu_AKP1021C 
            Caption         =   "�����ճɱ�����(������)"
         End
         Begin VB.Menu Mnu_AKP1023C 
            Caption         =   "����Ч������"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_AKP1025C 
            Caption         =   "�ɱ���Ϣά��"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_AKP1017C 
            Caption         =   "���Ĳ��ϼ۸�ϵ�������ʷ���ά��"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_AKP1020C 
            Caption         =   "��Ʒ���۵���ά��"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu Mnu_AKH 
      Caption         =   "ԭ�Ϲ���"
      Begin VB.Menu Mnu_AKH1010C 
         Caption         =   "��ԭ�ϼ���ʵ��¼�뼰��ѯ����_AKH1010C"
      End
      Begin VB.Menu Mnu_AKH1020C 
         Caption         =   "���Ͻ����ʵ��¼�뼰��ѯ����_AKH1020C"
      End
      Begin VB.Menu Mnu_AKH1030C 
         Caption         =   "�����Ͳļ���ʵ��¼�뼰��ѯ����_AKH1030C"
      End
      Begin VB.Menu Mnu_AKH1040C 
         Caption         =   "�Ƕ����Ͳļ���ʵ��¼�뼰��ѯ����_AKH1040C"
      End
      Begin VB.Menu Mnu_AKH1050C 
         Caption         =   "��ԭ�ϳɷּ���ʵ����ѯ����_AKH1050C"
      End
      Begin VB.Menu Mnu_AKH1060C 
         Caption         =   "��ԭ�ϴ�����չ�ϵ��_AKH1060C"
      End
      Begin VB.Menu Mnu_AKH1070C 
         Caption         =   "�����Ͳı�׼¼�뼰��ѯ����_AKH1070C"
      End
      Begin VB.Menu Mnu_AKH1080C 
         Caption         =   "��ԭ��������¯��ƥ��ʵ����ѯ����_AKH1080C"
      End
      Begin VB.Menu Mnu_AKH1090C 
         Caption         =   "���Ͻ��յ��ʲ�ѯ����_AKH1090C"
      End
      Begin VB.Menu Mnu_AKH1011C 
         Caption         =   "���Ͻ�۸��׼¼�뼰��ѯ����_AKH1011C"
      End
      Begin VB.Menu Mnu_AKH1012C 
         Caption         =   "���Ͻ�۸񱨱�_AKH1012C"
      End
   End
   Begin VB.Menu Mnu_AKC 
      Caption         =   "����"
      Begin VB.Menu Mnu_AKW2010C 
         Caption         =   "����ϵͳ״̬��ѯ"
      End
      Begin VB.Menu Mnu_AKT1030C 
         Caption         =   "JOBS ����״̬��ѯ"
      End
      Begin VB.Menu Mnu_AKW2030C 
         Caption         =   "��ҵ����ʱ��״����ѯ"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKW2040C 
         Caption         =   "��ҵʵ����ѯ"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKW2050C 
         Caption         =   "ת¯��LF��VD��̬����"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKW2060C 
         Caption         =   "���������߶�̬����"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_AKW2080C 
         Caption         =   "������"
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
'        Active_YN = GetSetting("NISCO", "EXE-FILE", "AK.exe")
'
'        If Active_YN = "1" Then
'            sUserID = GetSetting("NISCO", "AUTHORITY", "sUserID")
'            sUserName = GetSetting("NISCO", "AUTHORITY", "sUsername")
'            MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ : "
'            MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
'        Else
'            Call Gp_MsgBoxDisplay("ֻ�ܴ��������½...", "W")
'            Unload Me
'        End If
'
''        sUserID = "1JS1014"
''        sUserName = "�¾���"
''        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
''        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
''
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
        Active_YN = GetSetting("NISCO", "EXE-FILE", "AK.exe")
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
''
'        sUserID = "1JS1014"
'        sUserName = "�¾���"
'        MDIMain.StatusBar1.Panels(1) = "��ʾ��Ϣ ��"
'        MDIMain.StatusBar1.Panels(7) = sUserID + " " + sUserName
''''''

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
    
    If Cancel = False Then SaveSetting "NISCO", "EXE-FILE", "AK.exe", ""

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

Private Sub Mnu_AKA2000C_Click()
     AKA2000C.Show
     AKA2000C.SetFocus
End Sub

Private Sub Mnu_AKB2030C_Click()
    AKB2030C.Show
    AKB2030C.SetFocus
End Sub

Private Sub Mnu_AKC2900C_Click()
    AKC2900C.Show
    AKC2900C.SetFocus
End Sub

Private Sub Mnu_AKE2070C_Click()
    AKE2070C.Show
    AKE2070C.SetFocus
End Sub

Private Sub Mnu_AKG2020C_Click()
    AKG2020C.Show
    AKG2020C.SetFocus
End Sub

Private Sub Mnu_AKG2030C_Click()
    AKG2030C.Show
    AKG2030C.SetFocus
End Sub

Private Sub Mnu_AKG2040C_Click()
    AKG2040C.Show
    AKG2040C.SetFocus
End Sub

Private Sub Mnu_AKE1010C_Click()
    AKE1010C.Show
    AKE1010C.SetFocus
End Sub

Private Sub Mnu_AKE1020C_Click()
    AKE1020C.Show
    AKE1020C.SetFocus
End Sub

Private Sub Mnu_AKE1030C_Click()
    AKE1030C.Show
    AKE1030C.SetFocus
End Sub

Private Sub Mnu_AKH1010C_Click()
    AKH1010C.Show
    AKH1010C.SetFocus
End Sub

Private Sub Mnu_AKH1011C_Click()
    AKH1011C.Show
    AKH1011C.SetFocus
End Sub

Private Sub Mnu_AKH1012C_Click()
   AKH1012C.Show
    AKH1012C.SetFocus
End Sub

Private Sub Mnu_AKH1020C_Click()
    AKH1020C.Show
    AKH1020C.SetFocus
End Sub

Private Sub Mnu_AKH1030C_Click()
    AKH1030C.Show
    AKH1030C.SetFocus
End Sub

Private Sub Mnu_AKH1040C_Click()
    AKH1040C.Show
    AKH1040C.SetFocus
End Sub

Private Sub Mnu_AKH1050C_Click()
    AKH1050C.Show
    AKH1050C.SetFocus
End Sub

Private Sub Mnu_AKH1060C_Click()
  AKH1060C.Show
  AKH1060C.SetFocus
End Sub

Private Sub Mnu_AKH1070C_Click()
  AKH1070C.Show
  AKH1070C.SetFocus
End Sub

Private Sub Mnu_AKH1080C_Click()
  AKH1080C.Show
  AKH1080C.SetFocus
End Sub

Private Sub Mnu_AKH1090C_Click()
  AKH1090C.Show
  AKH1090C.SetFocus
End Sub

Private Sub Mnu_AKL2050C_Click()
    AKL2050C.Show
    AKL2050C.SetFocus
End Sub

Private Sub Mnu_AKM2030C_Click()
    AKM2030C.Show
    AKM2030C.SetFocus
End Sub

Private Sub Mnu_AKN2010C_Click()
    AKN2010C.Show
    AKN2010C.SetFocus
End Sub

Private Sub Mnu_AKN2020C_Click()
    AKN2020C.Show
    AKN2020C.SetFocus
End Sub

Private Sub Mnu_AKN2030C_Click()
    AKN2030C.Show
    AKN2030C.SetFocus
End Sub

Private Sub Mnu_AKN2031C_Click()
    AKN2031C.Show
    AKN2031C.SetFocus
End Sub

Private Sub Mnu_AKN2040C_Click()
    AKN2040C.Show
    AKN2040C.SetFocus
End Sub

Private Sub Mnu_AKN2050C_Click()
    AKN2050C.Show
    AKN2050C.SetFocus
End Sub

Private Sub Mnu_AKN3000C_Click()
    AKN3000C.Show
    AKN3000C.SetFocus
End Sub

Private Sub Mnu_AKO2010C_Click()
    AFO2010C.Show
    AFO2010C.SetFocus
End Sub

Private Sub Mnu_AKP1010C_Click()
    AKP1010C.Show
    AKP1010C.SetFocus
End Sub

Private Sub Mnu_AKP1011C_Click()
    AKP1011C.Show
    AKP1011C.SetFocus
End Sub
Private Sub Mnu_AKP1013C_Click()
    AKP1013C.Show
    AKP1013C.SetFocus
End Sub

Private Sub Mnu_AKP1014C_Click()
    AKP1014C.Show
    AKP1014C.SetFocus
End Sub

Private Sub Mnu_AKP1016C_Click()
    AKP1016C.Show
    AKP1016C.SetFocus
End Sub

Private Sub Mnu_AKP1017C_Click()
    AKP1017C.Show
    AKP1017C.SetFocus
End Sub

Private Sub Mnu_AKP1018C_Click()
    AKP1018C.Show
    AKP1018C.SetFocus
End Sub

Private Sub Mnu_AKP1019C_Click()
    AKP1019C.Show
    AKP1019C.SetFocus
End Sub

Private Sub Mnu_AKP1020C_Click()
    AKP1020C.Show
    AKP1020C.SetFocus
End Sub

Private Sub Mnu_AKP1021C_Click()
    AKP1021C.Show
    AKP1021C.SetFocus
End Sub

Private Sub Mnu_AKP1022C_Click()
    AKP1022C.Show
    AKP1022C.SetFocus
End Sub

Private Sub Mnu_AKP1023C_Click()
    AKP1023C.Show
    AKP1023C.SetFocus
End Sub
Private Sub Mnu_AKP1025C_Click()
    AKP1025C.Show
    AKP1025C.SetFocus
End Sub

Private Sub Mnu_AKP1111C_Click()
    AKP1111C.Show
    AKP1111C.SetFocus
End Sub

Private Sub Mnu_AKP1211C_Click()
    AKP1211C.Show
    AKP1211C.SetFocus
End Sub

Private Sub Mnu_AKP1213C_Click()
    AKP1213C.Show
    AKP1213C.SetFocus
End Sub

'Private Sub Mnu_AKP3020C_Click()
'    AKP3020C.Show
'    AKP3020C.SetFocus
'End Sub

Private Sub Mnu_AKP3030C_Click()
    AKP3030C.Show
    AKP3030C.SetFocus
End Sub

Private Sub Mnu_AKP3050C_Click()
    AKP3050C.Show
    AKP3050C.SetFocus
End Sub

Private Sub Mnu_AKP3051C_Click()
    AKP3051C.Show
    AKP3051C.SetFocus
End Sub

Private Sub Mnu_AKP3052C_Click()
    AKP3052C.Show
    AKP3052C.SetFocus
End Sub

Private Sub Mnu_AKP3060C_Click()
    AKP3060C.Show
    AKP3060C.SetFocus
End Sub

Private Sub Mnu_AKP3061C_Click()
    AKP3061C.Show
    AKP3061C.SetFocus
End Sub

Private Sub Mnu_AKW2010C_Click()
    AKW2010C.Show
    AKW2010C.SetFocus
End Sub

Private Sub Mnu_AKT1030C_Click()
    AKT1030C.Show
    AKT1030C.SetFocus
End Sub

Private Sub Mnu_AKW2030C_Click()
    AKW2030C.Show
    AKW2030C.SetFocus
End Sub

Private Sub Mnu_AKW2040C_Click()
    AKW2040C.Show
    AKW2040C.SetFocus
End Sub

Private Sub Mnu_AKW2050C_Click()
    AKW2050C.Show
    AKW2050C.SetFocus
End Sub

Private Sub Mnu_AKW2060C_Click()
    AKW2060C.Show
    AKW2060C.SetFocus
End Sub

Private Sub Mnu_AKW2080C_Click()
    AKW2080C.Show
    AKW2080C.SetFocus
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
    Dim CurrentForm As Form
    Dim FormLD      As Boolean
    
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