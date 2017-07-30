VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Whenever an Add-in is being activated or deactivated
'the Ribbon needs to be refreshed

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookAddinInstall(ByVal Wb As Workbook)
    Call modCodeCount.RibbonRefresh
End Sub

Private Sub App_WorkbookAddinUninstall(ByVal Wb As Workbook)
    Call modCodeCount.RibbonRefresh
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub
