VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AEA3090C 
   Caption         =   "录入炼钢生产技术参数(POP_UP)_AEA3090C"
   ClientHeight    =   5070
   ClientLeft      =   4035
   ClientTop       =   3345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6495
   Begin VB.TextBox TXT_WID_GRP 
      Height          =   315
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   12
      Tag             =   "宽度组"
      Top             =   4005
      Width           =   1275
   End
   Begin VB.ComboBox Cbo_TIME 
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
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Tag             =   "type"
      Top             =   4545
      Width           =   1275
   End
   Begin VB.TextBox TXT_THK_GRP 
      Height          =   315
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   11
      Tag             =   "厚度组"
      Top             =   3420
      Width           =   1275
   End
   Begin VB.TextBox TXT_PRC_NAME 
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
      Left            =   1845
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "PLT"
      Top             =   1125
      Width           =   1905
   End
   Begin VB.TextBox TXT_PRC 
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
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "工序"
      Top             =   1125
      Width           =   465
   End
   Begin VB.TextBox txt_PRC_line 
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
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "机号"
      Top             =   585
      Width           =   420
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   1815
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "PLT"
      Top             =   90
      Width           =   4515
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   90
      Width           =   465
   End
   Begin VB.TextBox txt_prod_cd 
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
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "产品分类"
      Top             =   1665
      Width           =   465
   End
   Begin VB.TextBox txt_prod_cd_name 
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
      Left            =   1845
      MaxLength       =   40
      TabIndex        =   6
      Tag             =   "PLT"
      Top             =   1665
      Width           =   1185
   End
   Begin VB.TextBox TXT_APLY_ITEM_NAME 
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
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   8
      Tag             =   "PLT"
      Top             =   2205
      Width           =   3075
   End
   Begin VB.TextBox TXT_APLY_ITEM 
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
      Left            =   1350
      MaxLength       =   15
      TabIndex        =   7
      Tag             =   "适用项目"
      Top             =   2205
      Width           =   1860
   End
   Begin VB.TextBox TxT_stdgrd 
      Height          =   315
      Left            =   1350
      MaxLength       =   11
      TabIndex        =   9
      Tag             =   "钢种"
      Top             =   2790
      Width           =   2040
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   45
      Top             =   2790
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "钢种"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   45
      Top             =   2205
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "适用项目"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   45
      Top             =   1665
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "产品分类"
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
      Left            =   45
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "工厂"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   45
      Top             =   585
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "机号"
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
      Left            =   45
      Top             =   1125
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "工序"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   45
      Top             =   3420
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "厚度组"
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
   Begin Threed.SSCommand SSCommand2 
      Height          =   420
      Left            =   5085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "新增记录"
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   45
      Top             =   4005
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "宽度组"
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   45
      Top             =   4545
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "编辑时间"
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   45
      X2              =   6345
      Y1              =   4995
      Y2              =   4995
   End
End
Attribute VB_Name = "AEA3090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Me.Left = 3000
    Me.Top = 3000
    
    Cbo_TIME.AddItem "全年"
     Cbo_TIME.AddItem "3"
    Cbo_TIME.AddItem "6"
    Cbo_TIME.AddItem "9"
    Cbo_TIME.AddItem "12"
    
    Cbo_TIME.ListIndex = 0
End Sub

Private Sub Form_Resize()
Me.Height = 5475
Me.Width = 6615

End Sub

Private Sub SSCommand2_Click()
Dim plt As String

plt = Left(txt_PLT.Text, 1)
If plt = "b" Then   '------------sms
'call aea3100p

End If




If plt = "c" Then   '------------mill

'call aea3600p


End If






End Sub

Private Sub TXT_APLY_ITEM_KeyUp(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If
    
Dim sTemp_Code As String

                 If KeyCode = vbKeyF4 Then
                
                DD.sWitch = "MS"
                DD.sKey = "EP_CAPA_STD"
                DD.rControl.Add Item:=TXT_APLY_ITEM
                
                DD.rControl.Add Item:=TXT_APLY_ITEM_NAME
                
                Call Gf_Apply_DD(M_CN1, KeyCode)
                Exit Sub
                End If
                
               
              If Len(Trim(TXT_APLY_ITEM.Text)) = TXT_APLY_ITEM.MaxLength Then
                    
               sTemp_Code = TXT_APLY_ITEM.Text
            TXT_APLY_ITEM_NAME.Text = Gf_ApplyNameFind(M_CN1, "EP_CAPA_STD", Trim(sTemp_Code))
            Else
            
            TXT_APLY_ITEM_NAME.Text = ""
   
   
             End If
             
               
     
   
End Sub

Private Sub txt_PLT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_PLT
        DD.rControl.Add Item:=txt_PLT_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_PLT.Text)) = txt_PLT.MaxLength Then
        txt_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_PLT.Text), 2)
    Else
        txt_PLT_NAME.Text = ""
        
    End If



End Sub




Private Sub TXT_PRC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        DD.rControl.Add Item:=TXT_PRC
        DD.rControl.Add Item:=TXT_PRC_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(TXT_PRC.Text)) = TXT_PRC.MaxLength Then
        TXT_PRC_NAME.Text = Gf_ComnNameFind(M_CN1, "C0002", Trim(TXT_PRC.Text), 2)
    Else
        TXT_PRC_NAME.Text = ""
        
    End If
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_prod_cd.Text)) = txt_prod_cd.MaxLength Then
        txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
    Else
        txt_prod_cd_name.Text = ""
    End If
End Sub

Private Sub TxT_stdgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
            
               
                DD.nameType = "1"
                DD.sWitch = "MS"
                DD.rControl.Add Item:=TxT_stdgrd
                
                Call Gf_Stlgrd_DD(M_CN1, KeyCode)
                
            End If
End Sub
