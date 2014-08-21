VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Dim NetworkRAC As Double
Dim Entries As Double
Public CobblestoneMultiplier As Integer
Public URL As String
Dim UserTotalRAC As Double
Public id As Integer
Public Average As Double
Dim NetworkRacWithDay(999) As Double
Dim NetworkRacWithDayEntries(999) As Double
Dim NetworkEntries As Double
Dim UserEntries As Double
Dim UserRacWithDay(999) As Double
Dim UserRacWithDayEntries(999) As Double
Dim UserAverage As Double
Public bNonPopular As Boolean
Public Sub AddNetworkRacToMother(amount As Double, iDay As Long)
NetworkRAC = NetworkRAC + amount
NetworkRacWithDay(iDay) = NetworkRacWithDay(iDay) + amount
NetworkRacWithDayEntries(iDay) = NetworkRacWithDayEntries(iDay) + 1
NetworkEntries = NetworkEntries + 1
Average = NetworkRAC / NetworkEntries
End Sub
Public Sub AddNetworkRac(amount As Double, iDay As Long)
Dim MotherProject As Project
Set MotherProject = BoincProjects(Me.id)
Call MotherProject.AddNetworkRacToMother(amount, iDay)
End Sub
Public Sub Purge30DayNetworkDMAFromMother(iDay As Long)
Dim PriorDay As Long
PriorDay = iDay - 30
If PriorDay < 0 Then PriorDay = 0
NetworkRAC = NetworkRAC - NetworkRacWithDay(PriorDay)
NetworkEntries = NetworkEntries - NetworkRacWithDayEntries(PriorDay)
NetworkRacWithDay(PriorDay) = 0
NetworkRacWithDayEntries(PriorDay) = 0
End Sub
Public Sub Purge30DayNetworkDMA(iDay As Long)
BoincProjects(Me.id).Purge30DayNetworkDMAFromMother (iDay)
End Sub
Public Function GetNetworkRacFromMother() As Double
GetNetworkRacFromMother = NetworkRAC
End Function
Public Function GetNetworkRac() As Double
GetNetworkRac = BoincProjects(Me.id).GetNetworkRacFromMother
End Function
Public Function GetUserAverage() As Double
GetUserAverage = UserAverage
End Function
Public Sub AddUserRac(amount As Double, iDay As Long)
UserTotalRAC = UserTotalRAC + amount
UserEntries = UserEntries + 1
UserRacWithDay(iDay) = UserRacWithDay(iDay) + amount
UserRacWithDayEntries(iDay) = UserRacWithDayEntries(iDay) + 1
UserAverage = UserTotalRAC / UserEntries
End Sub
Public Sub Purge30DayUserDMA(iDay As Long)
Dim PriorDay As Long
PriorDay = iDay - 30
If PriorDay < 0 Then PriorDay = 0
UserTotalRAC = UserTotalRAC - UserRacWithDay(PriorDay)
UserEntries = UserEntries - UserRacWithDayEntries(PriorDay)
UserRacWithDay(PriorDay) = 0
UserRacWithDayEntries(PriorDay) = 0
End Sub
Public Function GetUserEntries() As Double
GetUserEntries = UserEntries
End Function
Public Function GetUserTotalRac() As Double
GetUserTotalRac = UserTotalRAC
End Function

