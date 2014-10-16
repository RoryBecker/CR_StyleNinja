Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.CodeRush.Core
Imports DevExpress.CodeRush.PlugInCore
Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.Refactor.Core


Public Delegate Sub CheckCodeIssuesDelegate(ByVal sender As Object, ByVal ea As CheckCodeIssuesEventArgs)
Public Delegate Sub ApplyDelegate(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
Public Delegate Sub CheckContentAvailabilityDelegate(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
Public Class CheckIssuesDelegateCaller
    Private mCheckIssues As CheckCodeIssuesDelegate
    Public Sub New(ByVal CheckIssues As CheckCodeIssuesDelegate)
        mCheckIssues = CheckIssues

    End Sub
    Public Sub Invoke(ByVal sender As Object, ByVal ea As CheckCodeIssuesEventArgs)
        mCheckIssues.Invoke(sender, ea)
    End Sub
End Class
Public Class CheckContentAvailabilityDelegateCaller
    Private mCheckAvailability As CheckContentAvailabilityDelegate
    Public Sub New(ByVal CheckAvailability As CheckContentAvailabilityDelegate)
        mCheckAvailability = CheckAvailability
    End Sub
    Public Sub Invoke(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
        mCheckAvailability.Invoke(sender, ea)
    End Sub
End Class
Public Class ApplyDelegateCaller
    Private mApply As ApplyDelegate
    Public Sub New(ByVal Apply As ApplyDelegate)
        mApply = Apply
    End Sub
    Public Sub Invoke(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
        mApply.Invoke(sender, ea)
    End Sub
End Class
