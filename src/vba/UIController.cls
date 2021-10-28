VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIController"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Public Ribbon As IRibbonUI

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    Ribbon.InvalidateControl ("PlantUMLBtn")
    Ribbon.InvalidateControl ("PlantUMLEdit")
End Sub

Private Sub Class_Initialize()
    Set App = Application
End Sub
