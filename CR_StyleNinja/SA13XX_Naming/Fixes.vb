Imports DevExpress.CodeRush.Core
Imports System.ComponentModel
Namespace SA13XX
    Public Module NamingFixes
#Region "Utility"
        Public Function ChangeFirstCharToUpper(ByVal Name As String) As String
            Return Char.ToUpper(Name.First) & Name.Substring(1)
        End Function
        Public Function ChangeFirstCharToLower(ByVal Name As String) As String
            Return Char.ToLower(Name.First) & Name.Substring(1)
        End Function
        Public Function RemoveHungarianPrefix(ByVal Name As String) As String
            Dim Pos = GetFirstUCaseChar(Name)
            If Pos = -1 Then
                Return Name
            End If
            Return Name.Substring(Pos)
        End Function
        Public Function GetFirstUCaseChar(ByVal Name As String) As Integer
            For i As Integer = 0 To Name.Length - 1
                If Char.IsUpper(Name(i)) Then
                    Return i
                End If
            Next i
            Return -1
        End Function
#End Region
#Region "RegisterRefactorings"
        Public Sub RegisterRefactorings(ByVal c As IContainer)
            Dim UppercaseInitial = c.CreateRefactoring("UppercaseFirstChar", "Change first Character to Uppercase", AddressOf ShouldBeUppercased_Check, AddressOf UppercaseFirstChar_Apply)
            UppercaseInitial.SolvedIssues.Add(Message_SA1300) ' MainElement and InitialLower
            UppercaseInitial.SolvedIssues.Add(Message_SA1304) ' NonPrivateField and InitialLower
            UppercaseInitial.SolvedIssues.Add(Message_SA1307) ' PublicOrInternalField and InitialLower

            '-----------------------------------------
            Dim LowercaseInitial = c.CreateRefactoring("LowercaseFirstChar", "Change first Character to Lowercase", AddressOf ShouldBeLowercased_Check, AddressOf LowercaseFirstChar_Apply)
            LowercaseInitial.SolvedIssues.Add(Message_SA1306) ' Not (NonPrivateField) and InitialUpper

            '-----------------------------------------
            Dim PrefixNameWithI = c.CreateRefactoring("PrefixNameWithI", "Prefix name with 'I'", AddressOf InterfaceNotPrefixedI_Check, AddressOf PrefixNameWithI_Apply)
            PrefixNameWithI.SolvedIssues.Add(Message_SA1302)

            '-----------------------------------------
            Dim RemoveHungarianPrefix = c.CreateRefactoring("RemoveHungarianPrefix", "Remove Hungarian Prefix", AddressOf PrefixedWithHungarian_Check, AddressOf RemoveHungarianPrefix_Apply)
            RemoveHungarianPrefix.SolvedIssues.Add(Message_SA1305)

            '-----------------------------------------
            Dim RemoveUnderscores = c.CreateRefactoring("RemoveUnderscores", "Remove Underscores", AddressOf ContainsUnderscores_Check, AddressOf RemoveUnderscores_Apply)
            RemoveUnderscores.SolvedIssues.Add(Message_SA1309)
            RemoveUnderscores.SolvedIssues.Add(Message_SA1310)

        End Sub
#End Region
#Region "CheckAvailability"
        Friend Sub PrefixedWithHungarian_Check(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1305(ea.CodeActive)
        End Sub
        Friend Sub InterfaceNotPrefixedI_Check(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1302(ea.CodeActive)
        End Sub

        Friend Sub ShouldBeUppercased_Check(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1300(ea.CodeActive) _
                    OrElse Qualifies_SA1304(ea.CodeActive) _
                    OrElse Qualifies_SA1307(ea.CodeActive)
        End Sub
        Friend Sub ShouldBeLowercased_Check(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1306(ea.CodeActive)
        End Sub
        Friend Sub ContainsUnderscores_Check(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1309(ea.CodeActive) OrElse Qualifies_SA1310(ea.CodeActive)
        End Sub
#End Region
#Region "Apply"
        Public Sub UppercaseFirstChar_Apply(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            RenameElement(ea.CodeActive, ChangeFirstCharToUpper(ea.CodeActive.Name))
        End Sub
        Public Sub LowercaseFirstChar_Apply(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            RenameElement(ea.CodeActive, ChangeFirstCharToLower(ea.CodeActive.Name))
        End Sub
        Public Sub PrefixNameWithI_Apply(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            RenameElement(ea.CodeActive, "I" & ea.CodeActive.Name)
        End Sub
        Public Sub RemoveHungarianPrefix_Apply(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            RenameElement(ea.CodeActive, RemoveHungarianPrefix(ea.CodeActive.Name))
        End Sub
        Public Sub RemoveUnderscores_Apply(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            RenameElement(ea.CodeActive, ea.CodeActive.Name.Replace("_", ""))
        End Sub
#End Region
    End Module
End Namespace
