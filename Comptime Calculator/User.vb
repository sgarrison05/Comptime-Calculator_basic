Imports Microsoft.VisualBasic.ApplicationServices

Public Class Employee
    'Attempt to simulate a standard user
    Public Const dname As String = "Orange County Juvenile Probation Dept."
    Public user As String
    Public mposition As String
    Private _rate As Decimal

    Public Property Rate() As Decimal

        Get
            Return _rate
        End Get
        Set(value As Decimal)
            _rate = value
        End Set

    End Property

    Public Sub New()
        'Sets the Defaults
        user = Nothing
        mposition = Nothing
        Rate = 1.5D
    End Sub

End Class

Public Class Staff
    ' Used to create a standard staff member - i.e. Office Manager or Secretary

    Inherits Employee

    Public Sub New()
        'Set the defaults
        mposition = "Staff"
        Rate = 1.5D
    End Sub

End Class

Public Class JPO
    ' Used to create a JPO - Juvenile Probation Officer or ordinary line officer.

    Inherits Employee

    Public Sub New()
        'Set the defaults
        mposition = "JPO"
        Rate = 1.5D
    End Sub

End Class

Public Class Chief
    ' Used to create a Chief Juvenile Officer.

    Inherits Employee

    Public Sub New()
        'Set the defaults
        mposition = "Chief"
        Rate = 1D
    End Sub

End Class
