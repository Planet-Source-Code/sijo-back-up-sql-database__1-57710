Attribute VB_Name = "MdlMain"
Public Sub RestoreDB()
Dim DBPath As String
    Dim DIYA As New DIYAINI
        With DIYA
            .path = App.path & "\Settings.ini"
            .Section = "Database"
            .Key = "BPath"
            DBPath = .Value & "\"
            .Section = "RPoint" & FrmRestore.CAL.SDate
            .Key = FrmRestore.LstRestore.Text
            MsgBox DBPath & .Value & ".CAB"
        End With
End Sub

Public Sub AutoRestore(RestoreEvent As String)
KillReport
Dim FNME As String
Dim DIYAINI As New DIYAINI
    With DIYAINI
        .path = App.path & "\Settings.ini"
        .Section = "AvailDate"
        .Key = Date
        .Value = "Yes"
        .Section = "RPoint" & Date
        .Key = Time & " " & RestoreEvent
        .Value = Month(Date) & Day(Date) & Year(Date) & RestoreEvent & Hour(Time) & Minute(Time) & Second(Time) & Right(Time, 2)
    End With
'===============
FNME = Month(Date) & Day(Date) & Year(Date) & RestoreEvent & Hour(Time) & Minute(Time) & Second(Time) & Right(Time, 2) & ".CAB"
    With DIYAINI
        .path = App.path & "\Settings.ini"
        .Section = "DataBase"
        .Key = "BPath"
        BUPPath = .Value
    End With
        Dim SZIP As String
        SZIP = ".Option EXPLICIT" _
        & vbCrLf & ".Set MaxDiskSize = CDRom" _
        & vbCrLf & ".Set ReservePerCabinetSize = 44" _
        & vbCrLf & ".Set RptFileName=" & App.path & "\MKCB.rpt" _
        & vbCrLf & ".Set DiskDirectoryTemplate =" & Chr(34) & BUPPath & Chr(34) _
        & vbCrLf & ".Set CompressionType = MSZIP" _
        & vbCrLf & ".Set CompressionLevel = 7" _
        & vbCrLf & ".Set CompressionMemory = 21" _
        & vbCrLf & ".Set CabinetNameTemplate =" & Chr(34) & FNME & Chr(34) _
        & vbCrLf & ".Set Cabinet=on" _
        & vbCrLf & ".Set Compress=on" _
        '& vbCrLf & "Settings.ini"
        For a = 0 To FrmCrRPF.FL.ListCount - 1
            FrmCrRPF.FL.ListIndex = a
            SZIP = SZIP & vbCrLf & Chr(34) & FrmCrRPF.FL.FileName & Chr(34)
        Next
    Open App.path & "\DBZIP.SED" For Output As #1
        Print #1, SZIP
    Close #1
    Shell App.path & "\SSCB.exe /f DBZIP.SED", vbHide
    End
End Sub

Public Sub KillReport()
On Error Resume Next
    Kill App.path & "\MKCB.rpt"
End Sub
