Attribute VB_Name = "modStationIdentity"
Option Explicit

Public Function CurrentComputerStationId() As String
    Dim computerName As String

    computerName = Trim$(Environ$("COMPUTERNAME"))
    If computerName = "" Then computerName = "LOCAL-COMPUTER"
    CurrentComputerStationId = computerName
End Function
