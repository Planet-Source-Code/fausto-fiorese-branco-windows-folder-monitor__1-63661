VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitoring Folder"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   540
      TabIndex        =   2
      Text            =   "c:\Scripts"
      Top             =   0
      Width           =   7350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Folder:"
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Private Sub Pause(iSecs As Integer)
    
    Dim i As Integer

    For i = 1 To iSecs * 10
        Sleep 100
        DoEvents
    Next
End Sub

Private Sub Command1_Click()

Dim arrArquivo() As String

Me.Top = 100000
DoEvents
strComputer = "."  'Name of the computer or . to local machine
'Open WMI Services Connection
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

'Get the name of folder
strfolder = txtFolder.Text
'For this script each "\" need be replaced to "\\\\"
strfolder = Replace(Trim(strfolder), "\", "\\\\")

'Run query to monitor the folder.
'WITHIN 10 = is with 10 seconds
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceOperationEvent WITHIN 10 WHERE Targetinstance ISA 'CIM_DirectoryContainsFile' and TargetInstance.GroupComponent= 'Win32_Directory.Name=""" & strfolder & """'")

 
Do While True
    'Wait for any change in the folder.
    Set objEventObject = colMonitoredEvents.NextEvent()
        
    Select Case objEventObject.Path_.Class
        Case "__InstanceCreationEvent"
            
            ''To know all "Properties" in objEventObject.TargetInstance.PartComponent. Uncomment this block (this is interesting)
            'For Each objProperty In objWMIService.get(objEventObject.TargetInstance.PartComponent).Properties_
            '   Debug.Print objProperty.Name & "  -  " & objWMIService.get(objEventObject.TargetInstance.PartComponent).Properties_(objProperty.Name)
            'Next

            'Check if copy of the file is complete. This is necessary for big files.
            ' this monitor method, show when the process file is started, not when is finished. So I need know only on the process is finished.
            Do While Not FileExists(objWMIService.get(objEventObject.TargetInstance.PartComponent).Properties_("Name"))
               Call Pause(3)
               DoEvents
            Loop

            Call MsgBox("A new file was just created: " & _
                objWMIService.get(objEventObject.TargetInstance.PartComponent).Properties_("Name") & Chr(13) & " Do Something!")
            
            'Close all objects
            Set objEventObject = Nothing
            Set objWMIService = Nothing
            Set colMonitoredEvents = Nothing
            
            Me.Top = 0

            Exit Sub
        Case "__InstanceDeletionEvent"
            Call MsgBox("A file was just deleted: " & _
                objEventObject.TargetInstance.PartComponent & Chr(13) & " Do Something!")
            
            Me.Top = 0
            Exit Sub
    End Select
    DoEvents
Loop

End Sub


Private Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        'This is only way (simple) that I find to know if the copy of file is completly finished
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function


Private Sub Form_Load()

End Sub


