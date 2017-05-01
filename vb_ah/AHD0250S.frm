VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AHD0250S 
   Caption         =   "年出入库实绩发放_AHD0250S"
   ClientHeight    =   1335
   ClientLeft      =   2565
   ClientTop       =   3915
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   8970
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_prod_cd 
      Height          =   315
      Left            =   6705
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "产品"
      Top             =   315
      Width           =   615
   End
   Begin InDate.UDate dtp_yy_mm2 
      Height          =   315
      Left            =   3375
      TabIndex        =   1
      Tag             =   "年份月报"
      Top             =   315
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.UDate dtp_yy_mm1 
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Tag             =   "年份月报"
      Top             =   315
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   5310
      Top             =   315
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   2970
      Top             =   315
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "至"
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
      Left            =   270
      Top             =   315
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "年份月报"
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
End
Attribute VB_Name = "AHD0250S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Template System
'-- Sub_System Name   Common
'-- Program Name      Refer Template
'-- Program ID        Refer
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.19
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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(dtp_yy_mm1, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_yy_mm2, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    Me.KeyPreview = True

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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
    Call Form_Define
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set rControl = Nothing
    
    Set Mc1 = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

   Call Gp_Ms_Cls(Mc1("rControl"))
   Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Exc()
    

End Sub

Public Sub Form_Ref()

    If dtp_yy_mm1.RawData = "" Or dtp_yy_mm2.RawData = "" Or dtp_yy_mm1.RawData > dtp_yy_mm2.RawData Then
        Call Gp_MsgBoxDisplay("必须输入正确的日期.....")
       Exit Sub
    End If

    Dim sQuery As String
    Dim sMesg As String
    Dim AdoRs As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
     
 '   sQuery = "{call AHD0250S.P_REFER ('" + dtp_yy_mm1.RawData + "','" + dtp_yy_mm2.RawData + "','" + txt_prod_cd.Text + "')}"

    
    If Trim(txt_prod_cd) = "" Or Trim(txt_prod_cd) = "HC" Then
        sQuery = " SELECT  PROD_CD,GF_STLGRD_DETAIL(STLGRD), APLY_STDSPEC,THK,WID,PROD_GRD,"
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)< '" + dtp_yy_mm1.RawData + "' AND (SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)>= '" + dtp_yy_mm1.RawData + "')) THEN WGT ELSE 0 END),"
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)>= '" + dtp_yy_mm1.RawData + "' AND (SUBSTR(HOUSING_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "')) THEN WGT ELSE 0 END), "
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(SHP_DATE,1,6)>='" + dtp_yy_mm1.RawData + "' AND (SUBSTR(SHP_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "')) then WGT else 0 end) , "
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "' AND (SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)> '" + dtp_yy_mm2.RawData + "')) THEN WGT ELSE 0 END)"
        sQuery = sQuery + "  From gp_coil "
        sQuery = sQuery + "  where NVL(REC_STS,'1')  <>  '1'   and  PROD_GRD  IN  ('1','2','3')  And PROC_CD  >=  'X' and PROD_DATE  >=  '20051101'"
        sQuery = sQuery + "    and BED_PILE_DATE  IS NOT  NULL   and   HOUSING_DATE  IS  NOT  NULL "
'        sQuery = sQuery + "  where substr(HOUSING_DATE,1,6)<='" + dtp_yy_mm1.RawData + "' and ( SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)>= '" + dtp_yy_mm2.RawData + "' )"
 '       sQuery = sQuery + "  AND STLGRD LIKE '" + Trim(txt_stlgrd.Text) + "%'"
        sQuery = sQuery + "  group by PROD_CD,STLGRD, APLY_STDSPEC,THK,WID,PROD_GRD "
    End If
    
    If Trim(txt_prod_cd) = "" Or Trim(txt_prod_cd) = "PP" Then
        If Trim(txt_prod_cd) = "" Then
            sQuery = sQuery + " UNION "
            sQuery = sQuery + " SELECT  PROD_CD,GF_STLGRD_DETAIL(STLGRD), APLY_STDSPEC,THK,WID,PROD_GRD,"
        Else
            sQuery = " SELECT  PROD_CD,GF_STLGRD_DETAIL(STLGRD), APLY_STDSPEC,THK,WID,PROD_GRD,"
        End If
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)< '" + dtp_yy_mm1.RawData + "' AND (SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)>= '" + dtp_yy_mm1.RawData + "')) THEN WGT ELSE 0 END),"
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)>= '" + dtp_yy_mm1.RawData + "' AND (SUBSTR(HOUSING_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "')) THEN WGT ELSE 0 END), "
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(SHP_DATE,1,6)>='" + dtp_yy_mm1.RawData + "' AND (SUBSTR(SHP_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "')) then WGT else 0 end) , "
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "' AND (SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)> '" + dtp_yy_mm2.RawData + "')) THEN WGT ELSE 0 END)"
        sQuery = sQuery + "  From gp_PLATE "
        sQuery = sQuery + "  where NVL(REC_STS,'1')  <>  '1'   and  PROD_GRD  IN  ('1','2','3')  And PROC_CD  >=  'X' and PROD_DATE  >=  '20051101'"
        sQuery = sQuery + "    and BED_PILE_DATE  IS NOT  NULL   and   HOUSING_DATE  IS  NOT  NULL "
'        sQuery = sQuery + "  where substr(HOUSING_DATE,1,6)<='" + dtp_yy_mm1.RawData + "' and ( SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)>= '" + dtp_yy_mm2.RawData + "' )"
   '     sQuery = sQuery + "  AND STLGRD LIKE '" + Trim(txt_stlgrd.Text) + "%'"
        sQuery = sQuery + "  group by PROD_CD,STLGRD, APLY_STDSPEC,THK,WID,PROD_GRD "
        ' sQuery = sQuery + " UNION "
    End If
    
    If Trim(txt_prod_cd) = "" Or Trim(txt_prod_cd) = "SL" Then
        If Trim(txt_prod_cd) = "" Then
            sQuery = sQuery + " UNION "
            sQuery = sQuery + " SELECT  PROD_CD,GF_STLGRD_DETAIL(STLGRD), APLY_STDSPEC,THK,WID,PROD_GRD,"
        Else
            sQuery = " SELECT  PROD_CD,GF_STLGRD_DETAIL(STLGRD), APLY_STDSPEC,THK,WID,PROD_GRD,"
        End If
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)< '" + dtp_yy_mm1.RawData + "' AND (SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)>= '" + dtp_yy_mm1.RawData + "')) THEN WGT ELSE 0 END),"
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)>= '" + dtp_yy_mm1.RawData + "' AND (SUBSTR(HOUSING_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "')) THEN WGT ELSE 0 END), "
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(SHP_DATE,1,6)>='" + dtp_yy_mm1.RawData + "' AND (SUBSTR(SHP_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "')) then WGT else 0 end) , "
        sQuery = sQuery + "  sum(CASE WHEN (SUBSTR(HOUSING_DATE,1,6)<= '" + dtp_yy_mm2.RawData + "' AND (SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)> '" + dtp_yy_mm2.RawData + "')) THEN WGT ELSE 0 END)"
        sQuery = sQuery + "  From FP_SLAB "
        sQuery = sQuery + "  where NVL(REC_STS,'1')  <>  '1'   and  PROD_GRD  IN  ('1','2','3')  And PROC_CD  >=  'X' and PROD_DATE  >=  '20051101'"
        sQuery = sQuery + "    and BED_PILE_DATE  IS NOT  NULL   and   HOUSING_DATE  IS  NOT  NULL "
'        sQuery = sQuery + "  where substr(HOUSING_DATE,1,6)<='" + dtp_yy_mm1.RawData + "' and ( SHP_DATE is null OR SUBSTR(SHP_DATE,1,6)>= '" + dtp_yy_mm2.RawData + "' )"
 '       sQuery = sQuery + "  AND STLGRD LIKE '" + Trim(txt_stlgrd.Text) + "%'"
        sQuery = sQuery + "  group by PROD_CD,STLGRD, APLY_STDSPEC,THK,WID,PROD_GRD "
    End If
    
    Screen.MousePointer = vbHourglass
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If AdoRs.BOF Or AdoRs.EOF Then
    
        Call Gp_MsgBoxDisplay("There is No Relevant Data", "I")
        
        AdoRs.Close
        Set AdoRs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End If
        
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    oExcel.Visible = True
    oSheet.Range("A1") = "产品代码"
    oSheet.Range("B1") = "钢种"
    oSheet.Range("C1") = "标准代码"
    oSheet.Range("D1") = "厚度"
    oSheet.Range("E1") = "宽度"
    oSheet.Range("F1") = "等级"
    oSheet.Range("G1") = "期初库存"
    oSheet.Range("H1") = "本期入库"
    oSheet.Range("I1") = "本期出库"
    oSheet.Range("J1") = "期末库存"
    
    oSheet.Range("A2").CopyFromRecordset AdoRs
    
    Set oSheet = Nothing
    Set oBook = Nothing
    Set oExcel = Nothing
    
    AdoRs.Close
    
    Set AdoRs = Nothing
   
    Screen.MousePointer = vbDefault

End Sub

Public Sub Spread_Forzens_Setting()
    
End Sub

Public Sub Spread_Forzens_Cancel()

    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub




