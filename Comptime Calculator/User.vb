Imports Microsoft.VisualBasic.ApplicationServices

Public Class Users
    'Attempt to simulate a standard user

    Public user As String

End Class

Public Class Staff
    ' Used to create a standard staff member - i.e. Office Manager or Secretary

    Inherits Users
    Public rate As Integer = 1.5D


End Class

Public Class JPO
    ' Used to create a JPO - Juvenile Probation Officer or ordinary line officer.

    Inherits Users
    Public rate As Integer = 1.5D


End Class

Public Class Chief
    ' Used to create a Chief Juvenile Officer.

    Inherits Users
    Public rate As Integer = 1D


End Class
