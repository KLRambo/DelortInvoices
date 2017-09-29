
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Public NotInheritable Class LineItems

    Inherits CollectionBase

#Region " Public Properties "

    Default ReadOnly Property Item(ByVal index As Integer) As LineItem
        Get
            Return DirectCast(Me.InnerList.Item(index), LineItem)

        End Get
    End Property

#End Region

#Region " Public Routines "

    Public Sub Add(ByVal LineItem As LineItem)
        Me.InnerList.Add(LineItem)
    End Sub

#End Region

End Class
