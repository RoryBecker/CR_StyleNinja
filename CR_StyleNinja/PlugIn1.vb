Imports System.ComponentModel
Imports DevExpress.CodeRush.Core
Public Class PlugIn1

    'DXCore-generated code...
#Region " InitializePlugIn "
    Public Overrides Sub InitializePlugIn()
        MyBase.InitializePlugIn()
        RegisterRules()
        'TODO: Add your initialization code here.
    End Sub
#End Region
#Region " FinalizePlugIn "
    Public Overrides Sub FinalizePlugIn()
        'TODO: Add your finalization code here.

        MyBase.FinalizePlugIn()
    End Sub
#End Region

#Region "Rule Registration"
    Private Sub RegisterRules()
        SA11XX.RegisterRulesAndFixes(Components)
        SA12XX.RegisterRulesAndFixes(Components)
        SA13XX.RegisterRulesAndFixes(Components)
        SA14XX.RegisterRulesAndFixes(Components)
        SA15XX.RegisterRulesAndFixes(Components)
    End Sub


#End Region

    ' Code: Declare Delegate

    ' CreateIssue("Interfaces start with I", AddressOf InterfacesStartWithI_CheckCodeIssues)
    ' Code: Declare as Extension Method
End Class
