
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Public NotInheritable Class BrokenRules

    Public m_BrokenRules As New Hashtable
    Public Event ValidOrNot(ByVal isValid As Boolean)

    Public Sub BrokenRule(ByVal Rule As String, ByVal isBroken As Boolean)

        If isBroken Then
            If Not Me.m_BrokenRules.Contains(Rule) Then
                Me.m_BrokenRules.Add(Rule, True)
                RaiseEvent ValidOrNot(False)
            End If
        Else
            Me.m_BrokenRules.Remove(Rule)
            If Me.m_BrokenRules.Count = 0 Then
                RaiseEvent ValidOrNot(True)
            End If
        End If

    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me.m_BrokenRules.Count
        End Get
    End Property

    Public ReadOnly Property AmIValid() As Boolean
        Get
            Return Me.m_BrokenRules.Count = 0
        End Get
    End Property

End Class
