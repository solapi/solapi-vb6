Attribute VB_Name = "UTCDateTime"
Option Explicit

Private Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(63) As Byte  'unicode (0-based)
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(63) As Byte  'unicode (0-based)
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Sub Form_Load()
    txtDate.Text = GetUDTDateTime()
End Sub

Public Function GetUDTDateTime() As String
    Const TIME_ZONE_ID_DAYLIGHT As Long = 2
    Dim tzi As TIME_ZONE_INFORMATION
    Dim dwBias As Long
    'Dim sZone As String
    Dim tmp As String
    Select Case GetTimeZoneInformation(tzi)
        Case TIME_ZONE_ID_DAYLIGHT
            dwBias = tzi.Bias + tzi.DaylightBias
            'sZone = " (" & Left$(tzi.DaylightName, 1) & "DT)"
        Case Else
            dwBias = tzi.Bias + tzi.StandardBias
            'sZone = " (" & Left$(tzi.StandardName, 1) & "ST)"
    End Select
    tmp = " " & Right$("00" & CStr(dwBias \ 60), 2) & ":" & Right$("00" & CStr(dwBias Mod 60), 2)
    If dwBias > 0 Then
        Mid$(tmp, 1, 1) = "-"
    Else
        Mid$(tmp, 1, 2) = "+0"
    End If
    GetUDTDateTime = Format$(Now, "yyyy-mm-ddTHh:Mm:Ss") & tmp
End Function
