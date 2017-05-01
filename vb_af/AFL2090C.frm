VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFL2090C 
   Caption         =   "入库出库情况查询界面_AFL2090C"
   ClientHeight    =   9225
   ClientLeft      =   150
   ClientTop       =   1515
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15150
   WindowState     =   2  'Maximized
   Begin InDate.UDate sdt_in_plt_date 
      Height          =   315
      Left            =   3570
      TabIndex        =   10
      Top             =   165
      Width           =   1425
      _ExtentX        =   2514
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
   Begin VB.ComboBox cbo_yard_type 
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
      ItemData        =   "AFL2090C.frx":0000
      Left            =   1455
      List            =   "AFL2090C.frx":000D
      TabIndex        =   0
      Tag             =   "Yard"
      Top             =   165
      Width           =   675
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   6705
      Top             =   165
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "板坯数"
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
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8160
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   14393
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AFL2090C.frx":001A
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   9180
      Top             =   165
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "板坯重量"
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
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   12105
      Top             =   165
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "剩余存储能力"
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
   End
   Begin CSTextLibCtl.sidbEdit sdb_capa 
      Height          =   315
      Left            =   13410
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   165
      Width           =   1545
      _Version        =   262145
      _ExtentX        =   2725
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0"
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_wgt 
      Height          =   315
      Left            =   10500
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   165
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0"
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_count 
      Height          =   315
      Left            =   8010
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   165
      Width           =   1050
      _Version        =   262145
      _ExtentX        =   1852
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
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
      NumDecDigits    =   0
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   2250
      Top             =   165
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "日期"
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   150
      Top             =   165
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "库种类"
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
   End
   Begin Threed.SSCommand cmd_out 
      Height          =   375
      Left            =   13455
      TabIndex        =   7
      Top             =   8775
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   661
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
      Caption         =   "打印出库报表"
   End
   Begin Threed.SSCommand cmd_in 
      Height          =   375
      Left            =   11700
      TabIndex        =   8
      Top             =   8775
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   661
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
      Caption         =   "打印入库报表"
   End
   Begin Threed.SSCommand cmd_store 
      Height          =   375
      Left            =   9945
      TabIndex        =   9
      Top             =   8775
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   661
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
      Caption         =   "打印当前库存清单"
   End
   Begin InDate.UDate sdt_out_plt_date 
      Height          =   315
      Left            =   5130
      TabIndex        =   11
      Top             =   165
      Width           =   1425
      _ExtentX        =   2514
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
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   14985
      TabIndex        =   6
      Top             =   210
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   180
      Left            =   5010
      TabIndex        =   5
      Top             =   300
      Width           =   90
   End
End
Attribute VB_Name = "AFL2090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name
'-- Program Name
'-- Program ID        AFL2090C
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2003.11.10
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim s As String

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(cbo_yard_type, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdt_in_plt_date, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdt_out_plt_date, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_count, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_wgt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_capa, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cbo_yard_type_Change()
   If Trim(cbo_yard_type.Text) = "S" Then
      ULabel6.Caption = "板坯数"
      ULabel7.Caption = "板坯重量"
   ElseIf Trim(cbo_yard_type.Text) = "P" Then
      ULabel6.Caption = "钢板数"
      ULabel7.Caption = "钢板重量"
   ElseIf Trim(cbo_yard_type.Text) = "C" Then
      ULabel6.Caption = "钢卷数"
      ULabel7.Caption = "钢卷重量"
   End If
End Sub


Private Sub cbo_yard_type_Click()
   If Trim(cbo_yard_type.Text) = "S" Then
      ULabel6.Caption = "板坯数"
      ULabel7.Caption = "板坯重量"
   ElseIf Trim(cbo_yard_type.Text) = "P" Then
      ULabel6.Caption = "钢板数"
      ULabel7.Caption = "钢板重量"
   ElseIf Trim(cbo_yard_type.Text) = "C" Then
      ULabel6.Caption = "钢卷数"
      ULabel7.Caption = "钢卷重量"
   End If
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
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sQuery As String
    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
        
            If cbo_yard_type.ListIndex = 0 Then  'SLAB
            
                sQuery = "SELECT TO_DATE(PLT_DATE, 'YYYYMMDD'), SUM(IN_CNT),SUM(IN_WGT), SUM(OUT_CNT)  ,SUM(OUT_WGT) "
                sQuery = sQuery + " FROM (SELECT PROD_DATE PLT_DATE, COUNT(SLAB_NO) IN_CNT, SUM(WGT)  IN_WGT ,0 OUT_CNT ,0 OUT_WGT "
                sQuery = sQuery + "         FROM FP_SLAB "
                sQuery = sQuery + "        WHERE PROD_DATE BETWEEN '" & sdt_in_plt_date.RawData & "' AND '" & sdt_out_plt_date.RawData & "' "
                sQuery = sQuery + "        GROUP BY PROD_DATE "
                sQuery = sQuery + "       UNION All "
                sQuery = sQuery + "       SELECT OUT_PLT_DATE PLT_DATE, 0 IN_CNT, 0 IN_WGT, COUNT(SLAB_NO) OUT_CNT ,SUM(WGT) OUT_WGT  "
                sQuery = sQuery + "         FROM FP_SLAB "
                sQuery = sQuery + "        WHERE OUT_PLT_DATE BETWEEN '" & sdt_in_plt_date.RawData & "' AND '" & sdt_out_plt_date.RawData & "' "
                sQuery = sQuery + "        GROUP BY OUT_PLT_DATE) "
                sQuery = sQuery + " GROUP BY PLT_DATE "
                
            ElseIf cbo_yard_type.ListIndex = 1 Then  'PLATE
            
                sQuery = "SELECT TO_DATE(PLT_DATE, 'YYYYMMDD'),SUM(IN_CNT),SUM(IN_WGT), SUM(OUT_CNT)  ,SUM(OUT_WGT) "
                sQuery = sQuery + " FROM (SELECT PROD_DATE PLT_DATE, COUNT(PLATE_NO) IN_CNT,SUM(WGT) IN_WGT, 0 OUT_CNT,0 OUT_WGT "
                sQuery = sQuery + "         FROM GP_PLATE "
                sQuery = sQuery + "        WHERE PROD_DATE BETWEEN '" & sdt_in_plt_date.RawData & "' AND '" & sdt_out_plt_date.RawData & "' "
                sQuery = sQuery + "        GROUP BY PROD_DATE "
                sQuery = sQuery + "       UNION All "
                sQuery = sQuery + "       SELECT OUT_PLT_DATE PLT_DATE, 0 IN_CNT, 0 IN_WGT,COUNT(PLATE_NO) OUT_CNT,SUM(WGT) OUT_WGT"
                sQuery = sQuery + "         FROM GP_PLATE "
                sQuery = sQuery + "        WHERE OUT_PLT_DATE BETWEEN '" & sdt_in_plt_date.RawData & "' AND '" & sdt_out_plt_date.RawData & "' "
                sQuery = sQuery + "        GROUP BY OUT_PLT_DATE) "
                sQuery = sQuery + " GROUP BY PLT_DATE "
            
            Else                                     'COIL
            
                sQuery = "SELECT TO_DATE(PLT_DATE, 'YYYYMMDD'), SUM(IN_CNT),SUM(IN_WGT), SUM(OUT_CNT)  ,SUM(OUT_WGT) "
                sQuery = sQuery + " FROM (SELECT PROD_DATE PLT_DATE, COUNT(COIL_NO) IN_CNT,SUM(WGT) IN_WGT, 0 OUT_CNT, 0 OUT_WGT "
                sQuery = sQuery + "         FROM GP_COIL "
                sQuery = sQuery + "        WHERE PROD_DATE BETWEEN '" & sdt_in_plt_date.RawData & "' AND '" & sdt_out_plt_date.RawData & "' "
                sQuery = sQuery + "        GROUP BY PROD_DATE "
                sQuery = sQuery + "       UNION All "
                sQuery = sQuery + "       SELECT OUT_PLT_DATE PLT_DATE, 0 IN_CNT,0 IN_WGT, COUNT(COIL_NO) OUT_CNT,SUM(WGT) OUT_WGT "
                sQuery = sQuery + "         FROM GP_COIL "
                sQuery = sQuery + "        WHERE OUT_PLT_DATE BETWEEN '" & sdt_in_plt_date.RawData & "' AND '" & sdt_out_plt_date.RawData & "' "
                sQuery = sQuery + "        GROUP BY OUT_PLT_DATE) "
                sQuery = sQuery + " GROUP BY PLT_DATE "
            
            End If
            
            If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
            
                If cbo_yard_type.ListIndex = 0 Then      'SLAB
                    
                    'COUNT
                    sQuery = "SELECT COUNT(*) FROM FP_SLAB WHERE LOC IS NOT NULL AND REC_STS <> '3' "
                    sdb_count.VALUE = Gf_FloatFind(M_CN1, sQuery)
                    'WGT
                    sQuery = "SELECT NVL(SUM(WGT),0) FROM FP_SLAB WHERE LOC IS NOT NULL AND NVL(OUT_PLT_CD, ' ') = ' ' "
                    sdb_wgt.VALUE = Gf_FloatFind(M_CN1, sQuery)
                    
                    
                ElseIf cbo_yard_type.ListIndex = 1 Then  'PLATE
                
                    'COUNT
                    sQuery = "SELECT COUNT(*) FROM GP_PLATE WHERE LOC IS NOT NULL AND NVL(OUT_PLT_CD, ' ') = ' ' "
                    sdb_count.VALUE = Gf_FloatFind(M_CN1, sQuery)
                    'WGT
                    sQuery = "SELECT NVL(SUM(WGT),0) FROM GP_PLATE WHERE LOC IS NOT NULL AND NVL(OUT_PLT_CD, ' ') = ' ' "
                    sdb_wgt.VALUE = Gf_FloatFind(M_CN1, sQuery)
                                
                Else                                     'COIL
                
                    'COUNT
                    sQuery = "SELECT COUNT(*) FROM GP_COIL WHERE LOC IS NOT NULL AND NVL(OUT_PLT_CD, ' ') = ' ' "
                    sdb_count.VALUE = Gf_FloatFind(M_CN1, sQuery)
                    'WGT
                    sQuery = "SELECT NVL(SUM(WGT),0) FROM GP_COIL WHERE LOC IS NOT NULL AND NVL(OUT_PLT_CD, ' ') = ' ' "
                    sdb_wgt.VALUE = Gf_FloatFind(M_CN1, sQuery)
                    
                
                End If
                
                'CAPACITY
                sdb_capa.VALUE = CLng(5000) * 67 - sdb_wgt.VALUE
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                
            End If
            
        Else
            sMesg = sMesg + "长度不正确"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + "必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        
    End If
    
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
        
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

