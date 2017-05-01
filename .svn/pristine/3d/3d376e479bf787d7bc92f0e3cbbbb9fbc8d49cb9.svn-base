VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReport 
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   780
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   90
      Visible         =   0   'False
      Width           =   1635
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   9975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12015
      lastProp        =   500
      _cx             =   21193
      _cy             =   17595
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    
    CRViewer91.Left = 0
    CRViewer91.Top = 0
    CRViewer91.Height = ScaleHeight
    CRViewer91.Width = ScaleWidth
    
End Sub


Public Sub form_init(ByVal oForm As Form)
    
    CRViewer91.ReportSource = oForm.Report
    CRViewer91.ViewReport
    Call CRViewer91.Zoom(75)
    
End Sub
