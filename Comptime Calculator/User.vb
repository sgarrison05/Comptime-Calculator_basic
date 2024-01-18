Imports Microsoft.VisualBasic.ApplicationServices

Public Class Employee
    'Attempt to simulate a standard user
    Public dname As String = "Orange County Juvenile Probation Dept."
    Public user As String
    Public mposition As String
    Public rate As Decimal

    Public Sub New()
        user = Nothing
        mposition = Nothing
        rate = Nothing
    End Sub

End Class

Public Class Staff
    ' Used to create a standard staff member - i.e. Office Manager or Secretary

    Inherits Employee

    Public Sub New()
        'Set the defaults
        mposition = "Office Staff"
        rate = 1.5D
    End Sub

End Class

Public Class JPO
    ' Used to create a JPO - Juvenile Probation Officer or ordinary line officer.

    Inherits Employee

    Public Sub New()
        'Set the defaults
        mposition = "JPO"
        rate = 1.5D
    End Sub

End Class

Public Class Chief
    ' Used to create a Chief Juvenile Officer.

    Inherits Employee

    Public Sub New()
        'Set the defaults
        mposition = "Chief"
        rate = 1D
    End Sub

End Class
