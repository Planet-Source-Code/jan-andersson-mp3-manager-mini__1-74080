VERSION 5.00
Begin VB.Form frmDemoMX3 
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtInfo 
      Height          =   1395
      Left            =   1740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   900
      Width           =   1695
   End
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   3120
      Top             =   3240
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4500
      TabIndex        =   4
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   435
      Visible         =   0   'False
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000002&
      Height          =   195
      Index           =   4
      Left            =   4740
      TabIndex        =   5
      Top             =   1800
      Width           =   525
      Visible         =   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4920
      TabIndex        =   3
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   435
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000002&
      Height          =   195
      Index           =   1
      Left            =   3540
      TabIndex        =   2
      Top             =   60
      Width           =   525
      Visible         =   0   'False
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000003&
      Height          =   195
      Index           =   2
      Left            =   3180
      TabIndex        =   1
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Select file(s) you want to encode, or just press the F2 key."
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   15
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000003&
      Height          =   195
      Index           =   5
      Left            =   3900
      TabIndex        =   6
      Top             =   1860
      Width           =   1185
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuF 
         Caption         =   "Select file(s) to encode..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuF 
         Caption         =   "Select destination folder encoded files..."
         Index           =   1
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuF 
         Caption         =   "Rip audio from CD..."
         Index           =   2
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuF 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuF 
         Caption         =   "Start encoding"
         Index           =   4
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuF 
         Caption         =   "Cancel ongoing job"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuF 
         Caption         =   "Clear the log window"
         Index           =   6
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuF 
         Caption         =   "Set encoding output folder to default"
         Index           =   7
      End
      Begin VB.Menu mnuF 
         Caption         =   "Exit"
         Index           =   8
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu mnuDemo 
      Caption         =   "Tools"
      Begin VB.Menu mnuD 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuDS 
         Caption         =   "Format MP3 file names as"
         Index           =   0
         Begin VB.Menu mnuFF 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuDS 
         Caption         =   "Change the value max diff in milliseconds"
         Index           =   1
      End
      Begin VB.Menu mnuDS 
         Caption         =   "Select CD drive"
         Index           =   2
      End
      Begin VB.Menu mnuDS 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuS 
         Caption         =   "Bitrate"
         Index           =   0
         Begin VB.Menu mnuB 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuS 
         Caption         =   "Priority"
         Index           =   1
         Begin VB.Menu mnuP 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuS 
         Caption         =   "Channels"
         Index           =   2
         Begin VB.Menu mnuM 
            Caption         =   "Both channels"
            Index           =   0
         End
         Begin VB.Menu mnuM 
            Caption         =   "Left channel"
            Index           =   1
         End
         Begin VB.Menu mnuM 
            Caption         =   "Right channel"
            Index           =   2
         End
         Begin VB.Menu mnuM 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuM 
            Caption         =   "Mono"
            Index           =   4
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuM 
            Caption         =   "Swap channels"
            Index           =   5
         End
      End
      Begin VB.Menu mnuS 
         Caption         =   "If destination file(s) exist"
         Index           =   3
         Begin VB.Menu mnuX 
            Caption         =   "Create a backup"
            Index           =   0
         End
         Begin VB.Menu mnuX 
            Caption         =   "Delete it"
            Index           =   1
         End
         Begin VB.Menu mnuX 
            Caption         =   "Stop encoding"
            Index           =   2
         End
      End
      Begin VB.Menu mnuS 
         Caption         =   "Interface language M3X "
         Index           =   4
         Begin VB.Menu mnuL 
            Caption         =   "Only avalible in my compiled MX3.dll"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuH 
         Caption         =   ""
         Index           =   0
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmDemoMX3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*************************************************************************************************
'*  Copyright© Pappsegull Sweden, http://freetranslator.webs.com <pappsegull@yahoo.se>
'*
'*
'* FEATURES
'* --------
'* - MP3 encoding
'* - Create playlists using filter
'* - Batch edit MP3 tags using filter.
'* - CD audio ripping with auto tagging.
'* - Auto add MP3 tags, integrated with CDDB.
'* - + some more useful stuff;-)
'* - For more features Download my FREE MX3.dll at http://www.mediafire.com/?2rqfqid592a7c
'*
'* This software is provided "as-is," without any express or implied warranty.
'* In no event shall the author be held liable for any damages arising from the use of this software.
'* If you do not agree with these terms, do not use it!
'* Use of the program implicitly means you have agreed to these terms.
'*
'* Permission is granted to anyone to use this software for any purpose,
'* including commercial use, and to alter and redistribute it, provided that
'* the following conditions are met:
'*
'* CONDITIONS
'* ----------
'*   1. All redistribution of source code files must retain all copyright
'*      notices that are currently in place, and this list of conditions without
'*      any modification.
'*   2. All redistribution in binary form must retain all occurrences of the
'*      above copyright notice and web site addresses that are currently in
'*      place (for example, in the About boxes).
'*   3. Modified versions in source or binary form must be plainly marked as
'*      such, and must not be misrepresented as being the original software.
'*
'*************************************************************************************************

Option Explicit
Dim WithEvents MX3 As clsMX3
Attribute MX3.VB_VarHelpID = -1

'// Click on the file sub menus
Private Sub mnuF_Click(Index As Integer)
Dim s$
    With MX3
        Select Case Index
            Case 0 'Select file(s) to encode.
                .SelectFilesToEncode
                If .EncFileInCount > 0 Then lbl(0) = .GetText([Info Enc F2 to start])
            Case 1 'Select output folder
                .SelectEncodeOutputFolder
                If .EncFileInCount > 0 And LenB(.EncFolderOut) Then _
                  lbl(0) = .GetText([Info Enc done folder])
            Case 2 'Rip Audio CD
                Call SettingsApply: ArrangeMenu True 'Disable file menus
                'Format Tracks$ with space between the tracks, i.e: .CDRIP , , "1 5 7 17"
                .CDRIP: ArrangeMenu
            Case 4 'Apply settings and start encoding
                ArrangeMenu True: SettingsApply
                '*****************************************************************************
                '*** Valid file formats is WAV and AIFF and need to be 32, 44.1 or 48 kHz! ***
                '*** ALL propertys need to be set BEFORE calling Encode()                  ***
                .Encode '*********************************************************************
                '*****************************************************************************
                ArrangeMenu
            Case 5 'Cancel the Encoding/Ripping/Searching process
                .StopWork
            Case 6 'Clear the info text box
                txtInfo = ""
            Case 7 'Clear property destination folder
                .SelectEncodeOutputFolder True
            Case 8 'Exit
                Unload Me
        End Select
        ArrangeMenu
    End With
End Sub

'// Click on the more sub menu
Private Sub mnuD_Click(Index As Integer)
Dim s$, t$, m$, e$, x%, y%, v$(), b As Boolean
Const c_MP3 = ".mp3", c_Def = "http://www.planet-source-code.com/vb/scripts/ShowZip.asp?lngWId=1&lngCodeId=72092&strZipAccessCode=tp%2FU720927461", c = vbLf & "Do you want to try again?" ', c_Done = "Done! This is the result." & vbLf & vbLf
    With MX3
        Select Case Index
            Case 0 'Tag and rename - Audio CD as source
                .TagRename [Audio CD as source]
            Case 1 'Tag and rename - MP3 album folder as source
                .TagRename [MP3 album folder as source]
            Case 2 'Tag and rename - By using CDID and category
                .TagRename [By CDID and category]
            Case 3 'Tag and rename - By using free search in CDDB
                .ShowFreeSearch
            Case 5 'Get genre name from number
                .ShowGenreFromNumber Me
            Case 6 'Get genre number from name
                .ShowGenreNumberFromName Me
            Case 7 'Show the File Size Calculator
               .ShowFileSizeCalc Me, .MP3File, .EncBitrate
            Case 8 'Show MP3 Tag editor
                .ShowTaggerMP3 Me
            Case 9 'Show MP3 Tag Batch Editor
                .ShowTagBatchEdit Me
            Case 10 'Create playlist from folder
                .CreatePlaylist
            Case 11 'Show Play List Creator
                .ShowPlaylistCreator Me
            Case 13 'Download a file
                m$ = .ShowInput(.GetText([Select file to download]) & " The default file is my Decimal Clock project from PSC;-)", "Download Demo", c_Def, Me, False, 0, txtLeft)
                'm$ = InputBox(.GetText([Select file to download]) & vbLf & vbLf & "The default file is my Decimal Clock project from PSC;-)", "Download Demo", c_Def)
                If m$ <> vbNullString Then b = (MsgBox("Do you want to delete the cache?", 292) = vbYes) Else Exit Sub
                If m$ = c_Def Then t$ = "Decimal Clock.zip": e$ = ".zip": s$ = "Zip Files (*.zip)|*.zip" Else s$ = "All Files (*.*)|*.*"
                t$ = .FileFolder(encShowSave, s$, .GetText([Select save download]), , , , t$, e$)
                If t$ = vbNullString Then Exit Sub
                lbl(0) = "Wait...Downloading your request. If you like to have a assync. downloader, there is one built in in my translator GAT ActiveX for FREE;-)"
                WindowState = 2: Refresh: DoEvents
                If Not .DownloadFile(m$, t$, b) Then
                    t$ = "Sorry, could not download the file": lbl(0) = t$
                    MsgBox t$ & ":" & vbLf & m$, 48
                Else
                    lbl(0) = "Done! The file is saved to: " & t$
                    If MsgBox("Done! Do you want to open it?", 36) = vbYes Then .ExecuteShell t$
                End If
            Case 14 'Backup a file
                .MsgBoxW "This function renames a file or copy it if it's locked, in this demo it force to copy."
                s$ = .ShowOpen("All Files (*.*)|*.*", .GetText([Select file to backup]))
                If s$ = vbNullString Then Exit Sub
                t$ = .BackupFile(s$, True): If LenB(t$) Then _
                  .MsgBoxW "Done! the file:" & vbLf & s$ & vbLf & "Is copied to:" & vbLf & t$
            Case 15 'Demo - Merge MP3-Files Dialtone
                s$ = .ShowInput("Input a number to create a dialer tone.", _
                  "Demo of merging MP3-Files", "0123456789", Me, True, 20, txtRight)
                m$ = .PathApp & "Sounds\": e$ = m$ & "Dial tone " & s$ & c_MP3
                For x% = 1 To Len(s$)
                    t$ = Val(Mid$(s$, x%, 1)) & c_MP3
                    If x% = 1 Then
                        If .FileExists(e$) Then Kill e$
                        If Not .FileExists(m$ & t$) Then _
                          .MsgBoxW "Can't find file:" & vbLf & m$ & t$, 48: Exit Sub
                        FileCopy m$ & t$, e$
                    Else: .MP3Merge e$, m$ & t$: End If
                Next
                .MP3AutoAddTag e$: .ExecuteShell e$
            Case 16 'Demo - Merge MP3-Files you select files
                If .MP3Merge(s$) Then .ExecuteShell s$
            Case 17 'Show built in strings in MX3
                v$() = .InterfaceLanguage(True)
                For x% = 0 To UBound(v$())
                    If v$(x%) <> vbNullString Then _
                      s$ = s$ & "Index=" & Format(x%, "000") & "  " & v$(x%) & vbNewLine
                Next
                .ShowTextForm Me, s$, "Built in text in MX3": Erase v$()
        End Select
    End With
End Sub

'// MP3 file name format menu
Private Sub mnuFF_Click(Index As Integer)
    MX3.MP3FileFormat = Index: ArrangeMenu
End Sub

'// Demo settings menu
Private Sub mnuDS_Click(Index As Integer)
Dim s$, t$, m$, l&, d$()
    With MX3
        Select Case Index
            Case 1 'Edit the maximum allowable diff time in milliseconds
                .ShowAdjustDiffMs
            Case 2 'Select CD drive
                .ShowSelectCD
        End Select
    End With
End Sub

'// Click on the help menu
Private Sub mnuH_Click(Index As Integer)
    With MX3
        Select Case Index
            Case 0 'Read the language file again
            Case 1 'Translate the user interface again
            Case 2 'Show the language file in notepad
            Case 3 'Download source code
                .ExecuteShell "http://www.mediafire.com/?2rqfqid592a7c"
            Case 4 'Vote on my code
                .ExecuteShell "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=Jan+andersson&lngWId=1"
            Case 5 'My homepage
                .ExecuteShell "http://freetranslator.webs.com"
            Case 6 'Contact me
                .ExecuteShell "mailto:Pappsegull Sweden <pappsegull@yahoo.se>?subject=" & App.Title
            Case 7 'About...
                .MsgBoxW "About..."
        End Select
    End With
End Sub

'// Click on the sub menu "Bitrate", menu mnuB()
Private Sub mnuB_Click(Index As Integer)
    Dim x%: For x% = 0 To mnuB.UBound: mnuB(x%).Checked = False: Next
    mnuB(Index).Checked = True: MX3.EncBitrate = mnuB(Index).Tag ': .frmCalc.Bitrate = mnuB(Index).Tag
End Sub

'// Click on the sub menu "Priority", menu mnuB()
Private Sub mnuP_Click(Index As Integer)
    Dim x%: For x% = 0 To mnuP.UBound: mnuP(x%).Checked = False: Next: mnuP(Index).Checked = True
End Sub

'// Click on the sub menu "Mono channels", menu mnuX()
Private Sub mnuM_Click(Index As Integer)
Dim x%
    If Index = 4 Then
        mnuM(4).Checked = Not mnuM(4).Checked: MX3.EncMono = mnuM(4).Checked: ArrangeMenu
    ElseIf Index = 5 Then
        mnuM(5).Checked = Not mnuM(5).Checked: MX3.EncSwapChannels = mnuM(5).Checked: ArrangeMenu
    Else: For x% = 0 To 2: mnuM(x%).Checked = False: Next: mnuM(Index).Checked = True: End If
End Sub

'// Click on the sub menu "If destionation file exist", menu mnuX()
Private Sub mnuX_Click(Index As Integer)
    Dim x%: For x% = 0 To mnuX.UBound: mnuX(x%).Checked = False: Next: mnuX(Index).Checked = True
End Sub

'// Click on the sub menu "Settings", menu mnuS()
Private Sub mnuS_Click(Index As Integer)
    Select Case Index
        Case 24 'Save settings to disk
            Call SettingsApply: MX3.SettingsSave: Exit Sub
        Case 25 'Reset to default settings
            MX3.SettingsResetToDefault: ArrangeMenu: Exit Sub
        Case 26 'Show the log file in notepad if any data to show
            MX3.LogFileShow: Exit Sub
        Case Is < 6 Or Index > 22: Exit Sub
    End Select
    mnuS(Index).Checked = Not mnuS(Index).Checked: SettingsApply: ArrangeMenu
End Sub

'// Event MX3, show status while working...
Private Sub MX3_IsWorking(ByVal PecentDone As Single, ByVal Info As String, EventType As encEventTypes, ByVal CurretFileName As String, ByVal CurretFilePecentDone As Single)
Dim b As Boolean, s$, l&
    lbl(0) = Info: b = IIf(PecentDone = 100 Or PecentDone = 0, False, True)
    If b Then
        lbl(1).Width = (lbl(2).Width / 100) * PecentDone
        lbl(3) = Format(PecentDone, "0.00") & "%"             'Total process
    End If
    lbl(1).Visible = b: If Not b Then lbl(3) = MX3.GetText([Info Please vote]): lbl(1).Width = 0
    b = IIf(CurretFilePecentDone = 100 Or CurretFilePecentDone = 0, False, True)
    If b Then
        lbl(4).Width = (lbl(5).Width / 100) * CurretFilePecentDone
        lbl(6) = CurretFileName & " (" & Format(CurretFilePecentDone, "0.00") & "%)"  'Current file process
    End If
    lbl(4).Visible = b: lbl(5).Visible = b: lbl(6).Visible = b
    If Not b Then lbl(6) = "": lbl(4).Width = 0
    If EventType = encEventFileDone Or EventType = encEventJobDone Then
        s$ = Now & " --> " & Info & vbNewLine & _
          IIf(EventType = encEventJobDone, vbNewLine, "")      'Current file/Job complete
        l& = Len(txtInfo): txtInfo = txtInfo & s$
        txtInfo.SelStart = l&: txtInfo.SelLength = Len(s$)
    End If
    DoEvents
End Sub

'// Adjust progressbars and other controls.
Private Sub Form_Resize()
Dim l&: Const H = 14: On Local Error Resume Next
    lbl(0).Move 0, ScaleHeight - (H * 3), ScaleWidth, H: txtInfo.Move 0, 0, ScaleWidth, lbl(0).Top
    For l& = 4 To 6: lbl(l&).Move 0, lbl(0).Top + H, ScaleWidth, H: Next 'Progress bar last file
    For l& = 1 To 3: lbl(l&).Move 0, lbl(4).Top + H, ScaleWidth, H: Next 'Progress bar total
End Sub

'//Enable menus i.e Stop job if MX3 is working
Private Sub tmr_Timer()
    mnuF(5).Enabled = MX3.IsWorking
    mnuF(6).Enabled = LenB(txtInfo)
    mnuF(7).Enabled = LenB(MX3.EncFolderOut)
End Sub

'// Apply form settings to the class
Private Sub SettingsApply()
Dim l&, b As Boolean
    With MX3
        For l& = 0 To mnuB.UBound 'Bitrate (Same enum value as the tag property;)
            If mnuB(l&).Checked Then .EncBitrate = mnuB(l&).Tag: Exit For
        Next
        For l& = 0 To mnuX.UBound 'If destionation file exist (Same enum value as index;)
            If mnuX(l&).Checked Then .IfDestExist = l&: Exit For
        Next
        For l& = 0 To mnuP.UBound 'Priority "Lowest" is default in Blade (Same enum value as index;)
            If mnuP(l&).Checked Then .EncPriority = l&: Exit For
        Next
        .EncMono = mnuM(4).Checked         'Mono selected and mono channel
        .EncSwapChannels = mnuM(5).Checked 'Swap audio channels
        For l& = 0 To 2 'Both channels is default in Blade (Same enum value as index;)
            If mnuM(l&).Checked Then .EncMonoChannels = l&: Exit For
        Next
        'If display Blade window or not
        b = mnuS(6).Checked: .EncDisplayBlade = IIf(b, vbHide, vbNormalFocus)
        .EncCloseWhenDone = b Or mnuS(7).Checked           'Auto close Blade window
        .EncNoScreenOutput = mnuS(8).Checked               'No screen output in Blade window
        .EncDeleteSourceFiles = mnuS(9).Checked            'Blade delete source files when done
        .EncAddChecksum = mnuS(10).Checked                 'Add checksum to the MP3 file
        .EncIsPrivate = mnuS(11).Checked                   'Add "Private" flag to the MP3 file
        .EncIsCopyrighted = mnuS(12).Checked               'Add "Copyright" flag to the MP3 file
        .EncFilesJoin = mnuS(13).Checked                   'Join files into one MP3-File
        .EncFilesSorted = mnuS(14).Checked                 'Sort files if join
        .EncFilesNoGapIfJoin = mnuS(15).Checked            'Skip silece between files if join
        .AutoAddTag = mnuS(16).Checked                     'Add ID3 version 1 tag to the MP3 file
        '--------------------------------------------------------------------------------------
        .EncWaitUntilDone = mnuS(18).Checked 'If false the class just send the command to-
        'Blade, check that 1st file start creating, return True if do, and exit the function.
        'The StopWork() sub have then no function and event IsWorking won't fire.
        'Display a message box with result when done, if WaitUntilDone = True
        .EncMsgboxResult = mnuS(19).Checked
        .SettingsSaveOnExit = mnuS(21).Checked          'Save settings when class terminates
        .LogSave = mnuS(22).Checked                     'Log file will be created if true.
    End With
End Sub

'// Get default settings form the class and apply them to the menus
Private Sub ArrangeMenu(Optional IsWorking As Boolean)
Dim l&, b As Boolean
    With MX3
        For l& = 0 To mnuFF.UBound 'MP3 file format (Same enum value as the index;)
            mnuFF(l&).Checked = (.MP3FileFormat = l&)
        Next
        For l& = 0 To mnuB.UBound 'Bitrate (Same enum value as the tag property;)
            mnuB(l&).Checked = (.EncBitrate = mnuB(l&).Tag)
        Next
        For l& = 0 To mnuX.UBound 'If destionation file exist (Same enum value as index;)
            mnuX(l&).Checked = (.IfDestExist = l&)
        Next
        For l& = 0 To mnuP.UBound 'Priority "Lowest" is default in Blade (Same enum value as index;)
            mnuP(l&).Checked = (.EncPriority = l&)
        Next
        mnuM(4).Checked = .EncMono          'Mono selected
        mnuM(5).Checked = .EncSwapChannels  'Swap audio channels
        For l& = 0 To 2 'Mono channel, both channels is default in Blade (Same enum value as index;)
            mnuM(l&).Checked = (.EncMonoChannels = l&)
        Next
        b = .EncDisplayBlade = vbHide: mnuS(6).Checked = b   'If display Blade window or not
        mnuS(7).Checked = .EncCloseWhenDone Or b             'Auto close Blade window
        mnuS(8).Checked = .EncNoScreenOutput                 'No screen output in Blade window
        mnuS(9).Checked = .EncDeleteSourceFiles              'Blade delete source files when done
        mnuS(10).Checked = .EncAddChecksum                   'Add checksum to the MP3 file
        mnuS(11).Checked = .EncIsPrivate                     'Add "Private" flag to the MP3 file
        mnuS(12).Checked = .EncIsCopyrighted                 'Add "Copyright" flag to the MP3 file
        mnuS(13).Checked = .EncFilesJoin                     'Join files into one MP3-File
        mnuS(14).Checked = .EncFilesSorted                   'Sort files if join
        mnuS(15).Checked = .EncFilesNoGapIfJoin              'Skip silece between files if join
        mnuS(16).Checked = .AutoAddTag                       'Add ID3 version 1 tag to the MP3 file
        '--------------------------------------------------------------------------------------
        mnuS(18).Checked = .EncWaitUntilDone   'If false the class just send the command to-
        'Blade, check that 1st file start creating, return True if do, and exit the function.
        'The EncodeStop() sub have then no function and event IsWorking won't fire.
        'Display a message box with result when done, if WaitUntilDone = True
        mnuS(19).Checked = .EncMsgboxResult
        mnuS(21).Checked = .SettingsSaveOnExit            'Save settings when class terminates
        mnuS(22).Checked = .LogSave                       'Log file will be created if true.
        'Hide & auto close Blade window
        b = mnuS(6).Checked: mnuS(7).Enabled = Not b
        mnuS(8).Enabled = Not b: If b Then mnuS(7).Checked = True
        'Auto tag & display MsgBox result.
        b = mnuS(18).Checked: mnuS(16).Enabled = b: mnuS(19).Enabled = b
        For l& = 0 To 2: mnuM(l&).Enabled = mnuM(4).Checked: Next 'Enable mono channels if mono.
        'The class automaticly change to default bitrate when change the mono property.
        'Default bitrate for stereo is 196kbps, mono 64kbps so edit the sub menu "Bitrate".
        For l& = 0 To mnuB.UBound: mnuB(l&).Checked = .EncBitrate = Val(mnuB(l&).Tag): Next
        mnuF(1).Enabled = Not mnuS(10).Checked 'Save file as and auto file name out
        'On Local Error Resume Next
        For l& = 0 To mnuF.UBound 'Enabled/Disable file menus
            If mnuF(l&).Caption <> "-" Then mnuF(l&).Enabled = Not IsWorking
        Next
        mnuD(5).Enabled = Not IsWorking                  'CD RIP
        mnuF(5).Enabled = (IsWorking And .EncWaitUntilDone) 'Cancel encoding process menu.
        b = .EncFileInCount > 1: mnuS(13).Enabled = b 'Check if more than one file so can join files
        b = b And mnuS(13).Checked: mnuS(14).Enabled = b: mnuS(15).Enabled = b
    End With
End Sub

'//Load the form and add some sub menus...
Private Sub Form_Load()
Dim l&, s$(): Const c = "...", c_CT = "Tag and rename - ": On Error GoTo Form_LoadErr

Set MX3 = New clsMX3: Caption = App.Title: Icon = MX3.Icon: Show
'Load sub menus to the demo menu.
    s$() = Split(c_CT & "Audio CD as source;" & c_CT & "MP3 album folder as source;" & c_CT & "By using CDID and category;" & c_CT & "With full text search to FreeDB;-;Find genre name from number;Find genre number from name;Show File Size Calculator;Show MP3 Tag editor;Show MP3 Tag Batch Editor;Create playlist from folder;Show Play List Creator;-;Download a file;Backup a file;Merge MP3-Files (Dialtone);Merge MP3-Files (Select files);Show built in strings", ";"): mnuD(0).Caption = s$(0) & c: Debug.Print 0; s$(0)
    For l& = 1 To UBound(s$())
        Load mnuD(l&): mnuD(l&).Caption = s$(l&) & IIf(s$(l&) <> "-", c, "")
         Debug.Print l; s$(l)
    Next
    s$() = MX3.MP3FileNameFormats: mnuFF(0).Caption = s$(0)
    For l& = 1 To UBound(s$()): Load mnuFF(l&): mnuFF(l&).Caption = s$(l&): Next
'Load sub menus from index 5 to the settings menu
    s$() = Split("-;Hide the Blade window;Close Blade window when done;No screen output in Blade window;Delete source files;Add checksum;Add Private flag;Add Copyright flag;Join files to encode into one MP3-File;Sort files when join;Remove gap between files when join;Auto add ID tag (Only avalible if waiting);-;Wait until Blade is done;Show result in a message box when done, if waiting;-;Save settings when exit;Save logfile;-;Save settings;Reset to default settings;Show logfile", ";")
    For l& = 0 To UBound(s$()): Load mnuS(l& + 5): mnuS(l& + 5).Caption = s$(l&): Next
'Load sub menu with the valid bitrates, and save values to tag, coz is the same as enum values.
    s$() = MX3.EncBitrates: mnuB(0).Caption = s$(1) & s$(0)
    mnuB(0).Tag = s$(1) ': lbl(4) = lbl(4).Tag & "(s):"
    For l& = 1 To UBound(s$()) - 1
        Load mnuB(l&): mnuB(l&).Tag = s$(l& + 1)
        mnuB(l&).Caption = s$(l& + 1) & s$(0)
    Next
'Load sub menu with valid priority levels, default level in Blade is Lowest
    s$() = Split("Highest;Higher;Normal;Lower;Lowest;Idle", ";")
    mnuP(0).Caption = s$(1) & s$(0): mnuP(0).Tag = s$(1)
    For l& = 0 To UBound(s$())
        If l& > 0 Then Load mnuP(l&)
        mnuP(l&).Caption = s$(l&)
    Next
'Load sub menus to the help menu.

    s$() = Split("Read the language file again;Translate the user interface again;Open the language file;Download source code;Vote on my code;My homepage;Contact me;About", ";"): mnuH(0).Caption = s$(0)
    For l& = 1 To UBound(s$()): Load mnuH(l&): mnuH(l&).Caption = s$(l&) & c: Next
    For l& = 0 To 2: mnuH(l&).Enabled = False: Next
    Call ArrangeMenu: lbl(3) = MX3.GetText([Info Please vote]): Erase s$()
    Exit Sub
Form_LoadErr:
    MsgBox "Error number: " & Err & ", in Form_Load(). " & Err.Description, 16
End Sub

'// Exit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide: DoEvents: tmr.Enabled = False: MX3.StopWork: SettingsApply
    MX3.Terminate: Set MX3 = Nothing: Set frmDemoMX3 = Nothing
End Sub
Private Sub Form_Terminate(): End: End Sub
