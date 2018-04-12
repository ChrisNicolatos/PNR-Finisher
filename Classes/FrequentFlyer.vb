Option Strict Off
Option Explicit On
Namespace FrequentFlyer
    Friend Class FrequentFlyerItem
        Private Structure ClassProps
            Dim PaxName As String
            Dim Airline As String
            Dim FrequentTravelerNo As String
            Dim CrossAccrual As String
        End Structure
        Dim mudtProps As ClassProps

        Public ReadOnly Property PaxName As String
            Get
                PaxName = mudtProps.PaxName
            End Get
        End Property
        Public ReadOnly Property Airline As String
            Get
                Airline = mudtProps.Airline
            End Get
        End Property
        Public ReadOnly Property FrequentTravelerNo As String
            Get
                FrequentTravelerNo = mudtProps.FrequentTravelerNo
            End Get
        End Property
        Public ReadOnly Property CrossAccrual As String
            Get
                CrossAccrual = mudtProps.CrossAccrual
            End Get
        End Property
        Public Sub SetValues(ByVal pPaxName As String, ByVal pAirline As String, ByVal pFrequentTravelerNo As String, ByVal pCrossAccrual As String)
            With mudtProps
                .PaxName = pPaxName
                .Airline = pAirline
                .FrequentTravelerNo = pFrequentTravelerNo
                .CrossAccrual = pCrossAccrual
            End With
        End Sub
    End Class
    Friend Class FrequentFlyerColl
        Inherits Collections.Generic.Dictionary(Of String, FrequentFlyerItem)
        Friend Sub AddItem(ByVal pPaxName As String, ByVal pAirline As String, ByVal pFrequentTravelerNo As String, ByVal pCrossAccrual As String)
            If Not MyBase.ContainsKey(pPaxName) Then
                Dim pobjClass As FrequentFlyerItem

                pobjClass = New FrequentFlyerItem

                pobjClass.SetValues(pPaxName, pAirline, pFrequentTravelerNo, pCrossAccrual)
                MyBase.Add(pPaxName, pobjClass)
            End If

        End Sub
    End Class
End Namespace
