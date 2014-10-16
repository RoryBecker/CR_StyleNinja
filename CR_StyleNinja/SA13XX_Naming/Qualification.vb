Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Namespace SA13XX
    Module Qualification
#Region "Utility"
        Private Function isMainElement(ByVal e As IElement) As Boolean
            Return NamespacesClassesEnumsStructsDelegatesEventsMethodsPropertiesTypes.Contains(e.ElementType)
        End Function

        Private Function isPrivateReadonly(ByVal e As IMemberElement) As Boolean
            Return e.Visibility = MemberVisibility.Private AndAlso e.IsReadOnly
        End Function
        Private Function isPublicOrInternal(ByVal e As IMemberElement) As Boolean
            Return e.Visibility = MemberVisibility.Public _
                                             OrElse e.Visibility = MemberVisibility.Friend
        End Function
        Private Function isNonPrivate(ByVal e As IMemberElement) As Boolean
            Return e.Visibility <> MemberVisibility.Private
        End Function
        Private Function StartsAlpha(ByVal V As IElement) As Boolean
            Return Char.IsLetter(V.Name.First)
        End Function
        Private Function StartsLower(ByVal V As IElement) As Boolean
            Return Char.IsLower(V.Name.First)
        End Function
        Private Function StartsUpper(ByVal V As IElement) As Boolean
            Return Char.IsUpper(V.Name.First)
        End Function
#End Region
#Region "CodeRush Style Rule Qualification"
#Region "Qualifies_LocalWithPoorPrefix"
        Friend Message_LocalsShouldStart As String = String.Format("Locals should start with '{0}'", CodeRush.CodeStyle.PrefixLocal)
        Friend Function Qualifies_LocalWithPoorPrefix(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If

            Return Variable.IsLocal _
             AndAlso CodeRush.CodeStyle.PrefixLocal <> String.Empty _
            AndAlso Not Element.Name.StartsWith(CodeRush.CodeStyle.PrefixLocal)
        End Function
#End Region
#Region "Qualifies_FieldWithPoorPrefix"
        Friend Message_FieldsShouldStart As String = String.Format("Fields should start with '{0}'", CodeRush.CodeStyle.PrefixField)
        Friend Function Qualifies_FieldWithPoorPrefix(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return Variable.IsField _
            AndAlso CodeRush.CodeStyle.PrefixField <> String.Empty _
            AndAlso Not Element.Name.StartsWith(CodeRush.CodeStyle.PrefixParam)
        End Function
#End Region
#Region "Qualifies_ParamWithPoorPrefix"
        Friend Message_ParamsShouldStart As String = String.Format("Params should start with '{0}'", CodeRush.CodeStyle.PrefixParam)
        Friend Function Qualifies_ParamWithPoorPrefix(ByVal Element As IElement) As Boolean
            Return Element.ElementType = LanguageElementType.Parameter _
            AndAlso CodeRush.CodeStyle.PrefixParam <> String.Empty _
            AndAlso Not Element.Name.StartsWith(CodeRush.CodeStyle.PrefixParam)
        End Function
#End Region
#End Region
#Region "SA1300 Series Qualification"
#Region "Qualifies_SA1300"
        Friend Const Message_SA1300 As String = "SA1300: Element does not start with an uppercase char."
        Friend Function Qualifies_SA1300(ByVal Element As IElement) As Boolean
            Return isMainElement(Element) AndAlso StartsAlpha(Element) AndAlso StartsLower(Element)
        End Function
#End Region
#Region "Qualifies_SA1302"
        Friend Const Message_SA1302 As String = "SA1302: Interface should start with 'I'"
        Friend Function Qualifies_SA1302(ByVal Element As IElement) As Boolean
            Return Element.ElementType = LanguageElementType.Interface AndAlso Not Element.Name.StartsWith("I")
        End Function
#End Region
#Region "Qualifies_SA1304"
        Friend Const Message_SA1304 As String = "SA1304: Non Private Field must start with an uppercase char."
        Friend Function Qualifies_SA1304(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return isNonPrivate(Variable) AndAlso Variable.IsReadOnly AndAlso StartsLower(Variable)
        End Function
#End Region
#Region "Qualifies_SA1305"
        Friend Const Message_SA1305 As String = "SA1305: Variable must not use Hungarian Notation"
        Friend Function Qualifies_SA1305(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return (Variable.IsField OrElse Variable.IsLocal) AndAlso Variable.Name.IsHungarian
        End Function
#End Region
#Region "Qualifies_SA1306"
        Friend Const Message_SA1306 As String = "SA1306: Field must start with a lowercase char."
        Friend Function Qualifies_SA1306(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            If Not Variable.IsField Then
                Return False
            End If
            Return Not ((isPublicOrInternal(Variable) _
                         OrElse Variable.IsConst _
                         OrElse (isPrivateReadonly(Variable)))) _
                 AndAlso StartsUpper(Variable)
        End Function
#End Region
#Region "Qualifies_SA1307"
        Friend Const Message_SA1307 As String = "SA1307: Field must start with uppercase char."
        Friend Function Qualifies_SA1307(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return Variable.IsField AndAlso isPublicOrInternal(Variable) AndAlso StartsLower(Variable)
        End Function
#End Region
#Region "Qualifies_SA1308"
        Friend Const Message_SA1308 As String = "SA1308: Field must not be prefixed with m_ or s_"
        Friend Function Qualifies_SA1308(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return Variable.IsField AndAlso Variable.Name.StartsWith("s_") OrElse Variable.Name.StartsWith("m_")
        End Function
#End Region
#Region "Qualifies_SA1309"
        Friend Const Message_SA1309 As String = "SA1309: Field must not be prefixed with underscore"
        Friend Function Qualifies_SA1309(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return Variable.IsField AndAlso Variable.Name.StartsWith("_")
        End Function
#End Region
#Region "Qualifies_SA1310"
        Friend Const Message_SA1310 As String = "SA1310: Field must not contain underscores"
        Friend Function Qualifies_SA1310(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return Variable.IsField AndAlso Variable.Name.IndexOf("_") > 0
        End Function
#End Region
#End Region
    End Module
End Namespace
