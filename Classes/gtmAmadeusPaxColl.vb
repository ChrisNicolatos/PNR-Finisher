Option Strict Off
Option Explicit On
Friend Class gtmAmadeusPaxColl
    Inherits Collections.Generic.Dictionary(Of String, gtmAmadeusPax)
	
    Friend Sub AddItem(ByVal pElementNo As Short, ByVal pInitial As String, ByVal pLastName As String, ByVal pID As String)

        Dim pobjClass As gtmAmadeusPax

        pobjClass = New gtmAmadeusPax

        pobjClass.SetValues(pElementNo, pInitial, pLastName, pID)
        MyBase.Add(Format(pElementNo), pobjClass)


    End Sub
End Class