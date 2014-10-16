Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core

Module HiddenDX
    Public Function GetVisibilityName(ByVal newVisibility As MemberVisibility) As String
        Dim str As String = ""
        If Not CodeRush.Language.GetVisibility(newVisibility, str) Then
            Throw New ArgumentException(String.Format("Cannot change visibility to {0} because it is not supported by this language.", newVisibility), "newVisibility")
        End If
        Return str
    End Function
End Module
