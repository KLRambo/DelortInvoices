
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Public NotInheritable Class LineItem

#Region " Public Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal InvoiceNo As Integer)
        Me._InvoiceNo = InvoiceNo

    End Sub

#End Region

#Region " Enums "
    Public Enum Crud
        Insert = 0
        Update = 1
        Delete = 2
    End Enum
#End Region

#Region " Private Variables "

    Private _ItemNo As Integer
    Private _Date As Date = Now
    Private _LaborCost As String = ""
    Private _InvoiceNo As Integer = 0
    Private _LaborDesc As String = ""
    Private _IsValid As Boolean = False
    Private _Mode As Crud = Crud.Insert
    Private _MaterialDesc As String = ""
    Private _MaterialCost As String = ""
    Private _MarkedForDeletion As Boolean
    Private WithEvents _BrokenRules As New BrokenRules

#End Region

#Region " Events "

    Public Event MaterialCostChanged()
    Public Event LaborCostChanged()
    Public Event Valid(ByVal IsValid As Boolean)

#End Region

#Region " Public Properties "

    Public Property Mode() As Crud

        Get
            Return Me._Mode
        End Get

        Set(ByVal Value As Crud)
            Me._Mode = Value
        End Set

    End Property
    Public Property ItemDate() As Date
        Get
            Return Me._Date

        End Get
        Set(ByVal Value As Date)
            Me._Date = Value
            Me.CheckValidity()

        End Set
    End Property
    Public ReadOnly Property InvoiceNo() As Integer
        Get
            Return Me._InvoiceNo
        End Get
    End Property
    Public Property ItemNo() As Integer
        Get
            Return Me._ItemNo
        End Get
        Set(ByVal Value As Integer)
            Me._ItemNo = Value

        End Set
    End Property
    Public Property LaborDesc() As String

        Get
            Return Me._LaborDesc
        End Get

        Set(ByVal Value As String)

            If Value.Length > 255 Then
                Throw New Exception("Field can contain only 255 characters; last character will be truncated")
            End If

            Me._LaborDesc = Value

            CheckValidity()
        End Set

    End Property
    Public Property MaterialDesc() As String
        Get
            Return Me._MaterialDesc

        End Get

        Set(ByVal Value As String)

            If Value.Length > 255 Then
                Throw New Exception("Field can contain only 255 characters; last character will be truncated")
            End If

            Me._MaterialDesc = Value

            CheckValidity()

        End Set

    End Property

    Public Property LaborCost() As String
        Get
            Return Me._LaborCost
        End Get

        Set(ByVal Value As String)

            Me._LaborCost = Value
            RaiseEvent LaborCostChanged()
            CheckValidity()
        End Set

    End Property
    Public Property MaterialCost() As String
        Get

            Return Me._MaterialCost

        End Get

        Set(ByVal Value As String)
            Me._MaterialCost = Value
            RaiseEvent MaterialCostChanged()

            CheckValidity()

        End Set
    End Property
    Public ReadOnly Property Total() As Double

        Get

            Return Me.BlankToZero(Me._LaborCost) + Me.BlankToZero(Me._MaterialCost)

        End Get

    End Property
    Public Property IsValid() As Boolean
        Get
            Return Me._IsValid

        End Get
        Set(ByVal Value As Boolean)
            Me._IsValid = Value
            RaiseEvent Valid(Value)


        End Set
    End Property
    Public Property MarkedForDeletion() As Boolean

        Get
            Return Me._MarkedForDeletion

        End Get

        Set(ByVal Value As Boolean)
            Me._MarkedForDeletion = Value


        End Set
    End Property
#End Region

#Region " Public Routine(s) "

    Public Sub CheckValidity()

        'Rules for Valid LineItem
        If Me.LaborCost.Length > 0 And Me.LaborDesc.Length = 0 Then
            IsValid = False
            RaiseEvent Valid(False)
        ElseIf Me.LaborDesc.Length > 0 And Me.LaborCost.Length = 0 Then
            IsValid = False
            RaiseEvent Valid(False)
        ElseIf Me.MaterialDesc.Length > 0 And Me.MaterialCost.Length = 0 Then
            IsValid = False
            RaiseEvent Valid(False)
        ElseIf Me.MaterialCost.Length > 0 And Me.MaterialDesc.Length = 0 Then
            IsValid = False
            RaiseEvent Valid(False)
        ElseIf Me.LaborDesc.Length = 0 And Me.LaborCost.Length = 0 And Me.MaterialDesc.Length = 0 And Me.MaterialCost.Length = 0 Then
            IsValid = False
            RaiseEvent Valid(False)
        Else
            IsValid = True
            RaiseEvent Valid(True)
        End If

    End Sub
#End Region

#Region " Private Routines "
    Private Function BlankToZero(ByVal data As Object) As Double

        If IsNumeric(data.ToString) Then
            Return Convert.ToDouble(data.ToString)
        Else
            Return 0
        End If

    End Function
#End Region

End Class
