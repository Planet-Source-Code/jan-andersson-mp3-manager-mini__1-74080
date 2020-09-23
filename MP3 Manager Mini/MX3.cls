VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMX3"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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


'More info aboute Blade encoder here: http://home.swipnet.se/~w-82625/default.htm
'More ino about encoding here: http://teamcom.biz/MP3Handbook/16.htm
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Private Declare Function GetModuleHandleW Lib "Kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CopyFile Lib "Kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal sDrive As String) As Long

'// CD Ripping stuff START
Private Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "Kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Type WAVCHUNKHEADER
    ChunkID As String * 4
    ChunkSize As Long
End Type
Private Type WAVCHUNKFORMAT
    wFormatTag As Integer
    wChannels As Integer
    dwSamplesPerSec As Long
    dwAvgBytesPerSec As Long
    wBlockAlign As Integer
    wBitsPerSample   As Integer
End Type
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Enum TRACK_MODE_TYPE
    YellowMode2
    XAForm2
    CDDA
End Enum
Private Type RAW_READ_INFO
    DiskOffset  As LARGE_INTEGER
    SectorCount As Long
    TrackMode   As TRACK_MODE_TYPE
End Type
Private Type HeaderInfoWAV
    channels As Integer
    Frequency As Long
    bits As Integer
    Kbps As Long
    Playtime As String * 5
    FileSize As Double
End Type

Public Enum txtAlignment
    txtLeft
    txtRight
    txtCenter
End Enum

Public Enum WAVHeaderInfo
    wavChannels
    wavFrequency
    wavBits
    wavKbps
    wavFilesize
    wavPlaytime
End Enum

Public Enum encInfoCD
    encTracksCount
    encLenght
    encLenghtMs
    encTrackLengthMs
    encTrackStart
    encTrackStop
    encTrackLength
    encTrackStartLBA
    encTrackStopLBA
    encTrackLenghtLBA
    encTrackStartMin
    encTrackStartSec
    encTrackStartFra
    encTrackStopMin
    encTrackStopSec
    encTrackStopFra
    encReturnHandle
    encReturnQueryS
End Enum
Public Enum encMP3FileFormats
    [Do not rename the file(s)]
    [Track# - Artist - Year - Album - Title]
    [Track# - Artist - Album - Title]
    [Track# - Artist - Title]
    [Track# - Title]
    [Artist - Year - Album - Title]
    [Artist - Album - Title]
    [Artist - Title]
    [Title]
End Enum
Public Enum encCDDCQuerys
    [Audio CD as source]
    [MP3 album folder as source]
    [By CDID and category]
End Enum
'// CD Ripping stuff END

Public Enum encBuiltInText 'Some built in text string you can get, translated or not
    [Info Connect to Internet] = 4
    [Select Max Diff Ms] = 11
    [Select CD drive]
    [Select MP3 to add data]
    [Select track number]
    [Select file to download]
    [Select save download]
    [Select file to backup]
    [Translating] = 160
    [Translating done]
    [Translating failed]
    [Info Edit Language file]
    [Info Please vote]
    [Info Enc destination folder]
    [Info Enc done folder] = 167
    [Info Enc F2 to start]
End Enum

'// Event
Public Event IsWorking(ByVal PecentDone As Single, ByVal Info As String, EventType As encEventTypes, ByVal CurretFileName As String, ByVal CurretFilePecentDone As Single)
'Implements clsMP3.ShowFileCalc

'// Enum
Public Enum encEncodeBitRates   'Encoding bitrate
    [320 Kbps] = 320
    [256 Kbps] = 256
    [224 Kbps] = 224
    [192 Kbps] = 192
    [160 Kbps] = 160
    [128 Kbps] = 128
    [112 Kbps] = 112
    [ 96 Kbps] = 96
    [ 80 Kbps] = 80
    [ 64 Kbps] = 64
    [ 56 Kbps] = 56
    [ 48 Kbps] = 48
    [ 40 Kbps] = 40
    [ 32 Kbps] = 32
End Enum
Public Enum encIfDestExist      'What to do if destination file exists
    encCreateBackup             'Backup the file as default
    encDelete
    encStop
End Enum
Public Enum encStdInOut         'If use
    encOff                      'The default is Off, filname will be used.
    encIn                       'Read the input from a redirection or pipe
    encOut                      'Send the generated MP3 file through the-
End Enum                        'standard output stream
Public Enum EncEncMonoChannels
    encBothChannels
    encLeftChannel
    encRightChannel
End Enum
Public Enum encLogTypes         'To the class log to BladeMP3.log
    encLogEvent
    encLogError
    encLogErrorNoMsgBox
    encLogBackup
End Enum
Public Enum encEventTypes       'Type off message in event
    encEventWorking
    encEventJobDone
    encEventFileDone
    encEventError
    encEventInfo
End Enum
Public Enum encEncPrioritys        'Changes the task EncPriority of Blade
    encHighest
    encHigher
    encNormal
    encLower
    encLowest
    encIdle
End Enum
Public Enum encPathInfo        'Get info from path
    encExt
    encNameExt
    encName
    encPath
    encPathBS
    encNewExt
    encDrive
End Enum
Public Enum mciDlgType          'Dialog type to the FileFolder$ function
    encShowOpen
    encShowSave
    encShowFolder
End Enum
Public Enum encComDlgFlags      'Common dialog flags
    encAllowMultiSelect = &H200
    encCreatePrompt = &H2000
    encFileMustExist = &H1000
    encHideReadOnly = &H4
    encNoNetworkButton = &H20000
    encNoReadOnlyReturn = &H8000
    encNoTestfileCreate = &H10000
    encNoValidate = &H100
    encOverwritePrompt = &H2
    encPathMustExist = &H800
    encReadOnly = &H1
    encShareAware = &H4000
    encExplorer = &H80000
    encNoChangeDir = &H8
End Enum

'// This type hold file info when encoding
Private Type encEncodeEncFileInfo
    InName As String
    InPath As String
    InPathS As String
    OutName As String
    OutPath As String
    OutPathS As String
    SizeOrg As Long
    SizeCalc As Long
    SizeDone As Long
    HzOrg As Double
    LenMs As Long
    NotOk As Boolean
    TicStart As Long
    TicDone As Long
End Type
'// This type hold info when encoding
Private Type encEncodeInfo
    FI() As encEncodeEncFileInfo
    Path As String
    PathOut As String
    PathFolder As String
    nFiles As Long
    nFileNo As Long
    TicStart As Long
    LenMsTot As Double
    SizeOrgTot As Double
    SizeCalcTot As Double
    SizeDoneTot As Double
    SizeAdjustTot As Double
End Type
'ID3v1 tag format (128 bytes)
Private Type encID3v1Tag
    Tag As String * 3               'Always "TAG"
    Title As String * 30            'Title, 30 characters
    Artist As String * 30           'Artist, 30 characters
    Album As String * 30            'Album, 30 characters
    Year As String * 4              'Year, 4 characters
    Comment As String * 30          'Comment, 30 characters (or 28 if track# included)
    Genre As Byte                   'Genre, 255 for none defined
End Type
Public Enum encID3v1TagInfo
    encTagTitle
    encTagArtist
    encTagAlbum
    encTagYear
    encTagTrack
    encTagComment
    encTagGenre
    encTagGenreNo
End Enum
'// Private Constants
Private Const c_Tag = "TAG", c_BOM = "¥¿", c_D = """", c_Q = "?", c_MP3MaxDiffMs = 400, c_F = "nn:ss", c_3D = "...", GWL_WNDPROC = -4, c_Prio = "-prio=", c_eMailCDDB$ = "your@mail.pls", c_InetTest = "http://www.google.com", c_Home = "freetranslator.webs.com", c_ShortName = "MX3", c_BladeFile = "BladeEnc.mx3", c_LogFile = c_ShortName & ".log", c_SettingsFile = c_ShortName & " Settings.dat", c_Title = "MP3 Manager Mini"
'// Private Variables
Private TOC As CDROM_TOC, m_PathLanguageFile$, sLog$, m_eMailCDDB$, WAVHeader As HeaderInfoWAV, KbpsChk$(), _
   m_EncFileInCount&, m_TagExist As Boolean, m_Tag As encID3v1Tag, FSO As Object

Friend Sub EventRaise(sInfo$, Optional Percent%)
    RaiseEvent IsWorking(Percent%, sInfo$, encEventWorking, "", 0)
End Sub

'// Set default property settings
Sub SettingsResetToDefault()
Dim s$, P$, l&, b As Boolean, sT As encSettings: Static bInitDone As Boolean
    With m_Settings
        If bInitDone Then m_Settings = sT
        s$ = App.Path: If Right$(s$, 1) <> "\" Then s$ = s$ & "\"
        .PathApp = s$ 'Just if fail to load settings from file when initialize the class.
        If Not bInitDone Then b = SettingsLoad() 'Load from disk file
        If Not b Or bInitDone Then 'If bInitDone = True, user reset to default settings.
            .EncPriority = encLowest: .LogSave = True: .AutoAddTag = True: EncWaitIdleMsCheck = 500
            .EncBitrate = [128 Kbps]: .EncCloseWhenDone = True ': .FileOutAuto = True
            .EncMsgboxResult = True: .EncDisplayBlade = vbHide: .EncWaitUntilDone = True
            .MP3MaxDiffMs = c_MP3MaxDiffMs: CDGetDriveLetters$ .CDDriveLetter
            .DefComment = c_Home: .MP3FileFormat = 3: .CreatePlaylist = True: .LanguageCode = "en"
        End If
        P$ = Space(255): l& = GetSystemDirectory(P$, 255)
        P$ = Left(P$, l&): If Right(P$, 1) <> "\" Then P$ = P$ & "\"
        .PathApp = s$: .PathBlade = P$ & c_BladeFile: bInitDone = True
     End With
End Sub

'// Class Initialize & Terminate
Private Sub Class_Initialize() 'Load text in arrays and set default property settings
Dim l&, s$, m$: s$ = " - ": m$ = "#"
'Load strings
    For l& = 0 To c_ResStrUbound: IL(l&) = LoadResString(2000 + l&): Next
    IL(61) = IL(54) & m$ & s$ & IL(51) & s$ & IL(53) & s$ & IL(52) & s$ & IL(50)
    IL(62) = IL(54) & m$ & s$ & IL(51) & s$ & IL(52) & s$ & IL(50)
    IL(63) = IL(54) & m$ & s$ & IL(51) & s$ & IL(50)
    IL(64) = IL(54) & m$ & s$ & IL(50)
    IL(65) = IL(51) & s$ & IL(53) & s$ & IL(52) & s$ & IL(50)
    IL(66) = IL(51) & s$ & IL(52) & s$ & IL(50)
    IL(67) = IL(51) & s$ & IL(50)
    IL(68) = IL(50): ILOrg() = IL() 'Store the original english text to ILOrg()
    For l& = 0 To 147: m_GenreID3$(l&) = LoadResString(500 + l&): Next 'Load genre
    m_GenreID3$(255) = LoadResString(755): m_GenreID3Org$ = m_GenreID3$
    Set FSO = CreateObject("Scripting.FileSystemObject")
    KbpsChk$() = Split(" Kbps;320;256;224;192;160;128;112;96;80;64;56;48;40;32", ";")
    m_eMailCDDB$ = GetTickCount & "@hotmail.com": SettingsResetToDefault: Set M3 = Me
End Sub
Sub Terminate()
    Class_Terminate
End Sub
Private Sub Class_Terminate()
Dim x%
    m_Settings.StopWork = True: Set FSO = Nothing
    If m_Settings.SettingsSaveOnExit Then SettingsSave
    If m_Settings.LogSave Then LogfileSave
    Erase KbpsChk$(): Erase IL$(): Erase ILOrg$(): Erase m_GenreID3$()
    For x% = 0 To UBound(frm)
        frm(x%).TimeToUnload = True: Unload frm(x%): Set frm(x%) = Nothing
    Next
End Sub



'// Start the encoding, if you want you can change the propertys before you call this function
'// If not, the default propertys will be used, and you need to select source/destination files.
Public Function Encode(Optional MyOwnBladeCommand$) As Boolean: On Local Error GoTo EncodeErr
Dim EI As encEncodeInfo, bArr() As Byte, l&, lP&, n&, t&, P!, P2!, bJoin As Boolean, PF&, _
  bSort As Boolean, sIN$, sOUT$, sI$, sO$, s$, m$, sCmd$, sCap$, v$(), v2$(), _
  bD As Boolean, b As Boolean
EncodeStart:
    With m_Settings
    
        If (Dir(m_Settings.PathBlade) = vbNullString) Then 'Check if blade file is there
            l& = FreeFile: bArr = LoadResData(158, 0)
            Open .PathBlade For Binary As #l&: Put #l&, , bArr: Close #l&: Erase bArr
        End If
        If LenB(MyOwnBladeCommand$) Then sCmd$ = MyOwnBladeCommand$: GoTo RunBladeCommand
        n& = UBound(KbpsChk$): s$ = "" 'Check if valid bitrate.
        For l& = 1 To n&
            s$ = s$ & KbpsChk$(l&) & KbpsChk$(0) & ", "
            If .EncBitrate = Val(KbpsChk$(l&)) Then b = True
        Next
        If Not b Then 'Not valid bitrate
            MsgBoxW IL(2) & "." & vbLf & vbLf & IL(3) & ": " & _
              Left$(s$, Len(s$) - 2) & ".", 48: Exit Function
        End If
        sIN$ = .EncFileIn: b = False
        'Add command flags to Blade
        s$ = IIf(.EncBitrate = [128 Kbps], " ", " -br " & .EncBitrate & " ")
        If .EncAddChecksum Then s$ = s$ & "-crc "
        If .EncDeleteSourceFiles Then s$ = s$ & "-del "
        If .EncMono Then s$ = s$ & "-dm "
        If .EncMonoChannels = encLeftChannel Then s$ = s$ & "-lm "
        If .EncMonoChannels = encRightChannel Then s$ = s$ & "-lr "
        If .EncSwapChannels Then s$ = s$ & "-swap "
        If .EncIsPrivate Then s$ = s$ & "-p "
        If .EncIsCopyrighted Then s$ = s$ & "-c "
        If .EncClearOriginalFlag Then s$ = s$ & "-copy "
        If .EncNoScreenOutput Then s$ = s$ & "-quiet "
        If .EncCloseWhenDone Then s$ = s$ & "-q "
        If LenB(.EncFolderOut) Then
            If Right$(.EncFolderOut, 1) = "\" Then .EncFolderOut = Left$(.EncFolderOut, Len(.EncFolderOut) - 1)
            s$ = s$ & "-outdir=" & c_D & .EncFolderOut & c_D & " "
        End If
        If m_EncFileInCount > 1 Then 'Check if join files
            If .EncFilesJoin Then bJoin = True: s$ = s$ & "-concat "
            If bJoin Then
                If .EncFilesNoGapIfJoin Then s$ = s$ & "-nogap "
                If .EncFilesSorted Then bSort = True: s$ = s$ & "-sort "
            End If
        End If
        Select Case .EncPriority
            Case encHighest: sCmd$ = s$ & c_Prio & "HIGHEST"
            Case encHigher:  sCmd$ = s$ & c_Prio & "HIGHER"
            Case encNormal:  sCmd$ = s$ & c_Prio & "NORMAL"
            Case encLower:   sCmd$ = s$ & c_Prio & "LOWER"
            'Case encLowest:  sCmd$ = s$ & c_Prio & "LOWEST" 'Default Blade
            Case encIdle:    sCmd$ = s$ & c_Prio & "IDLE"
            Case Else:       sCmd$ = s$ 'If LOWEST or other integer used than valid
        End Select
        If LenB(sIN$) = 0 Then 'User have to select input file(s)
            sIN$ = FileFolder(encShowOpen, , , , encAllowMultiSelect + encExplorer): If LenB(sIN$) = 0 Then Exit Function
        End If
        If InStr(1, sIN$, vbNullChar) Then  'More than one file in selected
            n& = InStr(1, sIN$, vbNullChar & vbNullChar)
            If n& Then
                v$() = Split(Left$(sIN$, n& - 1), vbNullChar)
            Else: v$() = Split(sIN$, vbNullChar): End If
            EI.nFiles = UBound(v$()) - LBound(v$())
        Else                                'One file selected
            ReDim v$(1): v$(0) = PathInfo(sIN$, encPathBS)
            v$(1) = PathInfo(sIN$, encNameExt): EI.nFiles = 1
        End If
        EI.Path = v$(0)
        If LenB(.EncFolderOut) Then
            EI.PathOut = .EncFolderOut: If Right$(EI.PathOut, 1) <> "\" Then EI.PathOut = EI.PathOut & "\"
        End If
        If Right$(EI.Path, 1) <> "\" Then EI.Path = EI.Path & "\"
        If LenB(EI.PathOut) = 0 Then EI.PathOut = EI.Path
        sIN$ = c_D & EI.Path & v$(1) & c_D & " ": s$ = "": n& = 0: ReDim EI.FI(EI.nFiles)
    
        For l& = 1 To EI.nFiles 'Create file info for each file
            'If Trim$(v$(l&)) = vbNullString Then GoTo WasEmpty
            EI.FI(l&).InName = v$(l&): EI.FI(l&).InPath = EI.Path & v$(l&)  'Name & Path in
            EI.FI(l&).InPathS = GetShortPath(EI.Path)                       'Short path in
            v2$() = Split(v$(l&), "."): EI.FI(l&).OutName = v2$(0) & ".mp3" 'MP3 name out
            EI.FI(l&).OutPath = EI.PathOut & EI.FI(l&).OutName              'Path out
            EI.FI(l&).OutPathS = c_D & EI.PathOut & v$(l&) & c_D            'Short path out
            If FileExists(EI.FI(l&).OutPath) Then     'Check if destination file exists.
                Select Case .IfDestExist
                    Case encCreateBackup: BackupFile EI.FI(l&).OutPath
                    Case encDelete: 'Skip if locked, let Blade handle it :P
                    Case encStop: MsgBoxW IL(7) & ";-)" & vbLf & vbLf & _
                      EI.FI(l&).OutPath, 48: Exit Function
                End Select
            End If
            m$ = "": EI.FI(l&).SizeCalc = MP3CalcSize&(EI.Path & v$(l&), _
              .EncBitrate, False, EI.FI(l&).LenMs, EI.FI(l&).HzOrg, m$)        'Calculate
            EI.FI(l&).SizeOrg = FileLen(EI.Path & v$(l&))                   'Org. size byte
            If l& > 1 Then sIN$ = sIN$ & c_D & v$(l&) & c_D & " "           'Add name to sIn$
            EI.SizeCalcTot = EI.SizeCalcTot + EI.FI(l&).SizeCalc            'Total calc size
            EI.SizeOrgTot = EI.SizeOrgTot + EI.FI(l&).SizeOrg               'Total org. size
            EI.LenMsTot = EI.LenMsTot + EI.FI(l&).LenMs                     'Total lenght Ms
        Next
        sIN$ = Left$(sIN$, Len(sIN$) - 1): EI.SizeAdjustTot = EI.SizeCalcTot
        m$ = IL(112) & c_3D & IL(129) & " " & EI.nFiles & " " & _
          LCase(IL(IIf(EI.nFiles = 1, 39, 38)))
        sCap$ = GetShortPath(.PathBlade): sCmd$ = sCap$ & sCmd$ & " " & sIN$
RunBladeCommand:
    Debug.Print sCmd$
        '********************************************************************************
        Shell sCmd$, .EncDisplayBlade '*********   Done, run the stuff with shell   ********
        If LenB(MyOwnBladeCommand$) Then Encode = True: GoTo EncodeExit          '*******
        '********************************************************************************
        s$ = IIf(.AscyncRIP, IL(57) & " - " & EI.FI(1).InName & _
          " --> " & EI.FI(1).OutName, m$): P! = IIf(.AscyncRIP, .AscyncPercent, 1)
        RaiseEvent IsWorking(P!, s$, IIf(.AscyncRIP, encEventFileDone, encEventWorking), "", 0)
        EI.nFileNo = 1: .IsWorking = True
        sOUT$ = EI.FI(EI.nFileNo).OutPath: .MP3File = sOUT$
        Do
            l& = FindWindow(vbNullString, sCap$) 'Find the blade window, by checking caption
            DoEvents: If .StopWork Then GoTo EncodeExit
            If l& Then  'Blade Window found, change caption
                sCap$ = c_Title: CaptionW l&, sCap$: t& = GetTickCount + .EncWaitIdleMsCheck
                If Not .EncWaitUntilDone Then Encode = True: GoTo EncodeExit  'Not wait
                .hWndBladeWindow = l&: b = True: lP& = -1: DoEvents
                Do
                    DoEvents: If .StopWork Then GoTo EncodeExit
                    If GetTickCount - t& > .EncWaitIdleMsCheck Then 'Check if stil there every X sec...
                        If b Then
                            If Not FileExists(sOUT$) Then
                                If bJoin And bSort Then 'Find file to join into if sorted
                                    For n& = 1 To EI.nFiles
                                        sOUT$ = EI.FI(n&).OutPath
                                        If FileExists(sOUT$) Then
                                            EI.FI(1).OutName = EI.FI(n&).OutName
                                            EI.FI(1).OutPath = sOUT$: b = False: Exit For
                                        End If
                                    Next
                                    If b Then Err.Raise 53 'Blade can't encode the file
                                Else: Err.Raise 53: End If
                            End If
                            b = False: EI.TicStart = GetTickCount
                        End If
                        'Check filesize so far, calculate and raise event
                        'n& = FileLen(EI.FI(EI.nFileNo).OutPath)
                        n& = FileLen(sOUT$)
                        If n& = lP& Then  'Same file size as last check, so might ended
                            If bJoin Then 'Joining files
                                If Not IsFileOpen(sOUT$) Then b = True: bD = True
                            ElseIf EI.nFileNo < EI.nFiles Then  'Check if next file started
                                If Dir(EI.FI(EI.nFileNo + 1).OutPath) <> vbNullString Then
                                    EI.FI(EI.nFileNo + 1).TicStart = GetTickCount: bD = True
                                End If
                            Else
                                If (EI.nFiles = 1 Or EI.nFiles = PF&) Then
                                    'Check if Blade is done, if is last file...
                                    If Not IsFileOpen(sOUT$) Then b = True: bD = True
                                End If
                            End If
                        End If
                        P! = ((EI.SizeDoneTot + n&) / EI.SizeAdjustTot) * 100
                        If P! > 99 Then P! = 99
                        If P! < 0 Then P! = 0
                        If EI.nFiles > 1 Then _
                          P2! = (n& / EI.FI(EI.nFileNo).SizeCalc) * 100 Else P2! = P!
                        If n& = -1 Or bD Then
                            'Add tag on completed file
                            If .AutoAddTag Then MP3AutoAddTag sOUT$
                            n& = FileLen(sOUT$): EI.FI(EI.nFileNo).SizeDone = n&
                            t& = GetTickCount: EI.SizeDoneTot = EI.SizeDoneTot + n&
                            EI.FI(EI.nFileNo).TicDone = t&
                            'Adjust the calculated total file size to the real
                            EI.SizeAdjustTot = EI.SizeAdjustTot - _
                              (EI.FI(EI.nFileNo).SizeCalc - n&)
                             
                            bD = True: n& = 0: PF& = EI.nFileNo + 1: m$ = IL(108) & "! "
                        Else: m$ = IL(5) & " ": End If
                        If bJoin Then
                            s$ = EI.FI(1).OutName: P2! = 0
                             m$ = m$ & EI.nFiles & " files --> " & s$
                        Else
                            s$ = EI.FI(EI.nFileNo).OutName
                            m$ = m$ & EI.nFileNo & "/" & EI.nFiles & _
                              " --> " & EI.FI(EI.nFileNo).OutName
                        End If
                        If EI.nFiles = 1 Then P2! = 0
                        m$ = m$ & IIf(n& > 0, ", " & FormatKMG(n&, True, 2) & _
                          " " & .EncBitrate & KbpsChk$(0), "")
                        RaiseEvent IsWorking(P!, m$, IIf(bD, encEventFileDone, encEventWorking), s$, P2!)
                        t& = GetTickCount: m$ = "" ': lP& = n& ': P2! = 0
                        If bD Then
                            If EI.nFileNo < EI.nFiles Then EI.nFileNo = EI.nFileNo + 1
                            sOUT$ = EI.FI(EI.nFileNo).OutPath: Log m$: .MP3File = sOUT$
                        End If
                        lP& = n&: bD = False
                        If b Then  'Encoding done :-)
                            'Get filesize in bytes and seconds elapsed, and format message.
                            m$ = IL(108) & "! " & EI.nFiles & " " & _
                              IIf(EI.nFiles = 1, IL(39), LCase(IL(38))) & " " & IL(107) & " " & _
                              FormatMs(GetTickCount - EI.TicStart) & ", " & _
                              FormatKMG$(EI.SizeOrgTot, True, 2) & " --> " & _
                              FormatKMG$(EI.SizeAdjustTot, True, 2) & ", (" & _
                              FormatKMG$(EI.SizeOrgTot / ((GetTickCount - EI.TicStart) / 1000), True, 2) & "/sec) " & _
                              .EncBitrate & KbpsChk$(0) & " " & _
                              IIf(.EncMono, "EncMono", "Stereo")
                            Encode = True: Exit Do
                        End If
                    End If
                Loop
                If .AutoAddTag Then MP3AutoAddTag sOUT$ 'Add tag
                .hWndBladeWindow = 0: Log m$ & vbNewLine
                RaiseEvent IsWorking(100, m$, encEventJobDone, "", 100)
                If .EncMsgboxResult Then 'Show message box with result
                    If EI.nFiles > 1 And Not bJoin Then  'Did encode multiple files.
                        s$ = IL(36) & " " & IL(43)
                        sOUT$ = EI.PathOut ' IIf(LenB(.EncFolderOut), .EncFolderOut, EI.FI(0).OutPath)
                    Else
                        If bJoin Then sOUT$ = EI.FI(1).OutPath
                        m$ = m$ & vbLf & sOUT$: s$ = IL(6): .MP3File = sOUT$
                    End If
                    If MsgBoxW(m$ & vbLf & vbLf & s$, 36) = vbYes Then ExecuteShell sOUT$
                End If
            End If
        Loop Until Encode = True Or .StopWork = True
EncodeExit:
        If .StopWork Then 'Canceled by user
            If .hWndBladeWindow And .EncDisplayBlade = vbHide Then _
              SendMessage .hWndBladeWindow, &H10, 0&, 0&  'Close Blade
            RaiseEvent IsWorking(100, IL(5) & " " & LCase(IL(136)) & "!", encEventFileDone, "", 100)
        End If
        Erase EI.FI(): .IsWorking = False: .StopWork = False: Erase v$(): Erase v2$()
        Exit Function
EncodeErr:
        If Err = 5 Then 'The file BladeEnc.m3x is corrupt.
            Kill .PathBlade: Resume EncodeStart
        End If
        If .hWndBladeWindow And .EncDisplayBlade = vbHide Then _
          SendMessage .hWndBladeWindow, &H10, 0&, 0&  'Close Blade
        If Err = 53 Then 'File not found, Blade can't encode the selected file
            m$ = IL(58): s$ = m$: If EI.nFiles = 1 Then _
              m$ = EI.FI(EI.nFileNo).InName: s$ = EI.FI(EI.nFileNo).InPath
            RaiseEvent IsWorking(100, m$ & " " & IL(127), encEventFileDone, "", 100)
            MsgBoxW s$ & " " & IL(127) & vbLf & _
              IL(128) & " 32, 44.1 or 48 kHz.", 48: Resume EncodeExit
        End If
        Log "Encode(): " & Err.Description & vbNewLine & EI.FI(EI.nFileNo).InPath & vbNewLine & _
          EI.FI(EI.nFileNo).OutPath, encLogError: Resume EncodeExit
    End With
    Resume
End Function

'// Calculate the final size on the MP3 file by using the lenght of the source file and bitrate.
'// The function return the calculated file size in bytes, and -1 if error occurs.
'// RetValLenMs& return the lenght of the file in ms, sErr$ error string if error occurs.
'// RetSamplesPerSec# return the sampling frequency in Hz -1 if fail.
Public Function MP3CalcSize&(ByVal File$, ByVal Bitrate%, Optional ShowErr As Boolean, Optional RetValLenMs&, Optional RetSamplesPerSec#, Optional sErr$, Optional RetBytesPerSec#)
Const Sz = 32: Dim s As String * Sz, a$, l&, b As Boolean: Static lMs&, Hz#, FilePrev$, RetBps#
    If LenB(File$) And File$ <> FilePrev$ Then 'Skip MCI if same file as last request...
        'Exit if can't find file
        If Dir(File$) = vbNullString Then MP3CalcSize& = -1: Exit Function
        a$ = Trim$(CStr(GetTickCount)) 'Create a alias by using GetTickCount
        'Open MCI, exit and return error string to sErr$ if error.
        l& = mciSendString("open " & GetShortPath(File$, True) & " alias " & a$, s, Sz, 0&)
        If l& <> 0 Then sErr$ = ErrMCI$(l&, ShowErr): MP3CalcSize& = -1: Exit Function Else b = True
        'Set time format to milliseconds, return error string to sErr$ if error and close MCI
        l& = mciSendString("set " & a$ & " time format milliseconds", s, Sz, 0&)
        If l& <> 0 Then sErr$ = ErrMCI$(l&, ShowErr): MP3CalcSize& = -1: GoTo CloseMCI
        'Request length of the media in Ms, return error string to sErr$ if error and close MCI
        l& = mciSendString("status " & a$ & " length", s, Sz, 0&)
        If l& <> 0 Then sErr$ = ErrMCI$(l&, ShowErr): MP3CalcSize& = -1: GoTo CloseMCI _
          Else: lMs& = Val(s): RetValLenMs& = lMs&: FilePrev$ = File$
        'Request sampling frequency in Hz of the media, return -1 to RetSamplesPerSec# if fail.
        l& = mciSendString("status " & a$ & " samplespersec", s, Sz, 0&)
        If l& = 0 Then Hz# = Val(s): RetSamplesPerSec# = Hz# Else _
          Hz# = -1: RetSamplesPerSec# = -1
        'Request bitrate of the media, return -1 to RetBytesPerSec# if fail.
        l& = mciSendString("status " & a$ & " bytespersec", s, Sz, 0&)
        If l& = 0 Then RetBps# = Val(s): RetBytesPerSec# = RetBps# Else _
          RetBps# = -1: RetBytesPerSec# = -1
    Else: RetValLenMs& = lMs&: RetSamplesPerSec# = Hz#: RetBytesPerSec# = RetBps#: End If
    MP3CalcSize& = ((lMs& / 1000) * Bitrate% * 1024) / 8 'Calculate file size in bytes
CloseMCI:
    If b Then mciSendString "close " & a$, s, Sz, 0&
End Function

'// Play a media file
Public Function MP3Play(sFile$, sAlias$) As Boolean
Dim s$: If Not FileExists(sFile$) Then Exit Function
    s$ = SendString("open " & GetShortPath$(sFile$) & " alias " & sAlias$)
    If Not IsNumeric(s$) Then Exit Function Else SendString "play " & sAlias$: MP3Play = True
End Function

'// Stop play media file
Public Function MP3Stop(sAlias$) As Boolean
Dim s$
    s$ = SendString("stop " & sAlias$)
    If Val(s$) <> "0" Then Exit Function Else SendString "close " & sAlias$: MP3Stop = True
End Function

'// Get MCI error description
Private Function ErrMCI$(ErrNo&, Optional ShowMsgBoxw As Boolean)
Dim s As String * 256: mciGetErrorString ErrNo&, s, Len(s)
    s = Left$(s, InStr(s, vbNullChar) - 1)
    ErrMCI$ = Trim$(s): If ShowMsgBoxw Then MsgBoxW ErrMCI$, 16
End Function

'// Return short file path name.
Public Function GetShortPath$(File$, Optional IfFailChr34 As Boolean)
Dim s As String * 255: GetShortPathName File$, s$, Len(s$)
    GetShortPath$ = Left$(s, InStr(s, vbNullChar) - 1)
    If LenB(GetShortPath$) = 0 Then 'Can't create if file not exists
       ' Debug.Print GetShortPath$
        If IfFailChr34 Then 'Add " in start and end instead.
            GetShortPath$ = Chr(34) & File$ & Chr(34)
        Else: GetShortPath$ = File$: End If
    End If
End Function

'// Check if have a Internet connection
Public Function IsInternetConnected() As Boolean
Dim i&, u&
    i& = InternetOpen(vbNullString, 0&, vbNullString, vbNullString, 0&)
    If i& Then
      u& = InternetOpenUrl(i&, c_InetTest, vbNullString, 0&, -2076180480, 0&)
      If u& Then IsInternetConnected = True: InternetCloseHandle u&
    End If
    InternetCloseHandle i&
End Function

'// Download a file or page from internet, return true if success
Public Function DownloadFile(SourceFile$, DestinationFile$, Optional DeleteCache As Boolean, Optional NoMsgboxIfNotConnected As Boolean) As Boolean
Dim x&: If DeleteCache Then x& = DeleteUrlCacheEntry(SourceFile$): DoEvents 'Delete cached data
Retry:
'Debug.Print SourceFile$
    x& = URLDownloadToFile(0&, SourceFile$, DestinationFile$, &H10, 0&)
    If x& = 0 Then 'Succes to download the file if X& = 0
        DownloadFile = FileExists(DestinationFile$)
       ' If Not DownloadFile Then DownloadFile = Dir(DestinationFile$) <> vbNullString
    Else
        If Not IsInternetConnected Then 'Check if connected to internet
            If Not NoMsgboxIfNotConnected Then _
              If MsgBoxW(IL(4), vbRetryCancel + vbExclamation) = vbRetry Then GoTo Retry
        End If
    End If
End Function

'// Execute a shell command
Public Sub ExecuteShell(CmdString$, Optional hWnd&)
    ShellExecute hWnd&, "Open", CmdString$, vbNullString, "C:\", 1
End Sub

'// Rename a file, or copy it if locked, instead of delete it. Return the new file name if success.
Public Function BackupFile$(sFile$, Optional ForceCopy As Boolean)
Dim l&, v$(): On Local Error GoTo BackupFileErr
    v$() = Split(sFile$, "\"): v$() = Split(v$(UBound(v$())), ".")
    v$(0) = v$(UBound(v$())):  v$(1) = Left$(sFile$, Len(sFile$) - Len(v$(0)) - 1)
    Do
        l& = l& + 1: BackupFile$ = v$(1) & " [Backup " & l& & "]." & v$(0)
        If Not FileExists(BackupFile$) Then
            If ForceCopy Then l& = -1 Else Name sFile$ As BackupFile$
            Exit Do
        End If
    Loop
BackupFileExit:
    Log sFile$ & " " & IIf(l& = -1, "COPIED TO --> ", "RENAMED AS --> ") & BackupFile$, encLogBackup
    If l& = -1 Then _
      l& = CopyFile(sFile$, BackupFile$, 0&): If l& = 0 Then Err.Raise 70 'Permission denied
    Erase v$(): Exit Function
     
BackupFileErr:
    If Err = 75 Then '75 Path/File access error. The file is locked so copy it instead.
        l& = -1: Resume BackupFileExit
    Else 'Ask user if like to continue.
        If MsgBoxW(IL(10) & vbLf & vbLf & sFile$ & vbLf & BackupFile$ & vbLf & _
          vbLf & IL(11), 276) = vbNo Then m_Settings.StopWork = True
        BackupFile$ = "": l& = 0: Resume BackupFileExit
    End If
End Function

'// 'Set unicode caption in window.
Private Sub CaptionW(hWnd, ByRef NewValue$)
Dim m_Caption$, WndProc&, VBWndProc&: m_Caption = NewValue
    WndProc& = GetProcAddress(GetModuleHandleW(StrPtr("user32")), "DefWindowProcW")
    VBWndProc& = GetWindowLongA(hWnd, GWL_WNDPROC)
    If WndProc <> 0 Then
        SetWindowLongW hWnd, GWL_WNDPROC, WndProc&
        SetWindowTextW hWnd, StrPtr(m_Caption)
        SetWindowLongA hWnd, GWL_WNDPROC, VBWndProc&
    Else: NewValue = m_Caption: End If 'No Unicode
End Sub

'// Check if a file or folder exist
Function FolderExists(Folder$) As Boolean
    FolderExists = FileFolderExists(Folder$)
End Function
Function FileExists(File$) As Boolean
    FileExists = FileFolderExists(File$, True)
End Function
Private Function FileFolderExists(FileOrFolder$, Optional File As Boolean) As Boolean
    On Local Error Resume Next
    If File Then
        FileFolderExists = FSO.FileExists(FileOrFolder$)
    Else
        FileFolderExists = FSO.FolderExists(FileOrFolder$)
    End If
End Function


'//Create a playlist with the MP3 files in selected folder and sub directorys if selected
Function CreatePlaylist$(Optional ByVal sFolder$, Optional FileName$, Optional Pfx$ = "mp3", Optional bNestedDirs As Boolean, Optional RetFullPath As Boolean = True, Optional RetFileArray, Optional ShowMsgboxResult As Boolean)
Dim l&, n&, m&, s$, v$(), sF$, sFB$, b As Boolean
SelectFolder:
    If LenB(sFolder$) = 0 Then
        sFolder$ = FileFolder(encShowFolder, , IL(70))
        If LenB(sFolder$) Then
            b = True: If MsgBoxW(sFolder$ & vbLf & vbLf & _
              IL(71), 36) = vbYes Then bNestedDirs = True
        Else: Exit Function: End If
    End If
    If Not FolderExists(sFolder$) Then _
      MsgBox IL(82) & " " & sFolder$, 48: sFolder$ = "": GoTo SelectFolder
    If Right$(sFolder$, 1) <> "\" Then sFolder$ = sFolder$ & "\"
    v$() = DirW(sFolder$, Pfx$, bNestedDirs, RetFullPath)
    If v$(0) = "" Then MsgBox IL(74) & " " & UCase(Pfx$) & "-" & LCase(IL(38)) & _
      " " & IL(69) & "." & vbLf & sFolder$, 64: sFolder$ = "": GoTo SelectFolder
    If LenB(FileName$) Then s$ = FileReplBad$(FileName$) Else s$ = IL(75)
    CreatePlaylist$ = sFolder$ & s$ & ".m3u": sF$ = CreatePlaylist$
    If FileExists(sF$) Then 'Create backup if the playlist file exists
        sFB$ = BackupFile(sF$): sFB$ = vbLf & vbLf & IL(79) & ":" & vbLf & sFB$
    End If
    l& = FreeFile: Open CreatePlaylist$ For Output As #l&
    n& = UBound(v$()): Print #l&, "#EXTM3U"
    For m& = 0 To n&
        MP3ReadTagV1 v$(m&): s$ = MP3TagInfoGet(encTagTitle)
        Print #l&, "#EXTINF:" & MP3LenMs(v$(m&)) / 1000 & "," & s$
        Print #l&, Right$(v$(m&), Len(v$(m&)) - Len(sFolder$))
    Next
    Close #l&: RetFileArray = v$(): Erase v$()
    RaiseEvent IsWorking(0, nFiles& & " " & UCase(Pfx$) & "-" & LCase(IL(38)) & _
      ", " & IL(70) & " " & sFolder$, encEventJobDone, "", 0)
    If (b Or ShowMsgboxResult Or LenB(sFB$)) And nFiles& Then _
      If MsgBoxW(IL(72) & ", " & nFiles& & " " & LCase(IL(38)) & " " & IL(69) & "." & _
        vbLf & CreatePlaylist$ & sFB$ & vbLf & vbLf & IL(73), 36) = vbYes Then ExecuteShell CreatePlaylist$
End Function

'// Return all the files of the Filetype$ in a folder, and return result as an array
Friend Function DirW(ByVal Path$, Optional Pfx$ = "mp3", Optional ByVal bNestedDirs As Boolean, Optional RetFullPath As Boolean) As String()
Dim s$, FSO As Object, Folder As Object: nFiles& = 0: ReDim sTmpList$(0)
'Returns a 1-based string array containing the listed files
    If Not FolderExists(Path$) Then DirW = sTmpList$(): MsgBoxW IL(82) & Path$, 48: Exit Function
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(Path$)                'Get the starting folder
    m_Settings.IsWorking = True
    ListFiles Folder, bNestedDirs, Pfx$, RetFullPath 'Let the private subroutine do all the work
    DirW = sTmpList$(): Set FSO = Nothing: Set Folder = Nothing: Erase sTmpList$()
    s$ = nFiles& & " " & LCase(IL(IIf(nFiles& = 1, 39, 38))) & " " & IL(69) & " (" & Path$ & "*." & Pfx$ & ")"
    If m_Settings.StopWork Then s$ = IL(136) & "! " & s$
    RaiseEvent IsWorking(0, s$, encEventJobDone, "", 0)
    m_Settings.IsWorking = False: m_Settings.StopWork = False
End Function

'// List all the files in a directory
Private Sub ListFiles(ByVal Folder As Object, ByVal bNestedDirs As Boolean, Optional Pfx$, Optional RetFullPath As Boolean)
Dim File As Object, SubFolder As Object, b As Boolean, x%, s$, f$
On Error Resume Next
    DoEvents: If m_Settings.StopWork Then Exit Sub Else s$ = Folder.Path
    For Each File In Folder.Files 'List all the files in this directory
        If RetFullPath Then RaiseEvent IsWorking(0, nFiles& & " " & UCase(Pfx$) & "-" & _
          LCase(IL(38)) & ", " & IL(69) & " --> " & s$, encEventWorking, "", 0)
        b = True: If LenB(Pfx$) Then If LCase(Right$(File.Name, Len(Pfx$))) <> Pfx$ Then b = False
        With SearchTag
            If b Then
                If .CheckTag And LCase(Pfx$) = "mp3" Then 'Check search criteria in MP3 tag
                    MP3ReadTagV1 File.Path
                    For x% = 0 To 6
                        If x% = 6 Then 'Genre
                            If Val(.Check(x%)) <> -1 Then
                                s$ = MP3TagInfoGet(encTagGenreNo)
                                If .Check(x%) <> s$ Then b = False
                            End If
                        Else 'Title, Artist, Album, Year, Track, Comment
                            If .Check(x%) <> c_Q Then
                                s$ = MP3TagInfoGet(x%)
                                If .ExcactText Then
                                    If .Check(x%) <> s$ Then b = False: Exit For
                                Else
                                    If InStr(1, s$, .Check(x%)) = 0 Then b = False: Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                If b Then
                    DoEvents: If m_Settings.StopWork Then Exit Sub
                    f$ = IIf(RetFullPath, File.Path, PathInfo(File.Path, encNameExt))
                    ReDim Preserve sTmpList$(nFiles&): sTmpList$(nFiles&) = f$: nFiles& = nFiles& + 1
                End If
            End If
        End With
    Next
    If bNestedDirs Then 'If requested, search also subdirectories
        For Each SubFolder In Folder.Subfolders
            ListFiles$ SubFolder, bNestedDirs, Pfx$, RetFullPath
        Next
    End If
End Sub

'// Try to rename file to see if have full access to it Return True if file is open
Private Function IsFileOpen(sFile$) As Boolean
Dim s$, l&: On Local Error Resume Next
    s$ = sFile$ & ".Test": Name sFile$ As s$
    If Err Then
        Err.Clear: IsFileOpen = True: l& = GetTickCount 'Wait 50 Ms
        Do
            DoEvents: If GetTickCount - l& > 50 Then Exit Do
        Loop
    Else: Name s$ As sFile$: End If
End Function

''// Get TickCount Milliseconds
Friend Function GetTicCount&(): GetTicCount& = GetTickCount: End Function

'// Opens Common file dialogs folder select, save or file selection
Public Function FileFolder$(Optional DlgType As mciDlgType = encShowOpen, Optional ByVal Filter$, Optional DialogTitle$, Optional DirInit$, Optional flags As encComDlgFlags = encOverwritePrompt + encAllowMultiSelect + encExplorer, Optional FilterIndex% = -1, Optional FileName$, Optional ForceExtention = ".mp3")
Dim s$, e$, obj As Object, obj2 As Object: Static InitDir$, IdxF%, Count%: On Error GoTo FileFolderErr
    If DlgType = encShowFolder Then 'Show select folder
        Set obj = CreateObject("Shell.Application")
        s$ = IIf(LenB(DialogTitle$), DialogTitle$, "Select folder")
        If flags = &H11& Then
            Set obj = obj.BrowseForFolder(0&, s$, 0&, &H11&)     'This computer
        Else
            If LenB(DirInit$) = 0 Then DirInit$ = InitDir$
            Set obj = obj.BrowseForFolder(0&, s$, 0&, InitDir$)   'Specify directory
        End If
        If Not obj Is Nothing Then
            Set obj = obj.Self: FileFolder$ = obj.Path: InitDir$ = FileFolder$
        End If
        GoTo FileFolderExit
    End If
    Set obj = CreateObject("MSCOMDLG.COMMONDIALOG")
    If FilterIndex% <> -1 Then IdxF% = FilterIndex%
    If LenB(InitDir$) = 0 Then InitDir$ = App.Path: IdxF% = 1: Count% = 1
    If LenB(DirInit$) Then InitDir$ = DirInit$
       With obj
        .MaxFileSize = 1024: .InitDir = InitDir$: .FileName = "": .flags = flags
        If DlgType = encShowSave Then 'Show save MP3 sound as...
            .DialogTitle = IIf(LenB(DialogTitle$), DialogTitle$, "Save the MP3 as...")
            If LenB(Filter$) = 0 Then Filter$ = "MP3-" & IL(38) & " (*.mp3)|*.mp3"
            .Filter = Filter$: e$ = ForceExtention
            
            If LenB(FileName$) = 0 Then
                .FileName = "My MP3 " & Count% & e$: Count% = Count% + 1
            Else: .FileName = FileName$: End If
            s$ = .FileName
            .ShowSave: If .FileName <> s$ Then FileFolder$ = .FileName
            If LenB(FileFolder$) And LenB(ForceExtention) Then 'Check that filname i.e *.mp3
                If LCase(Right$(FileFolder$, 4)) <> e$ Then
                    FileFolder$ = FileFolder$ & e$
                End If
            End If
        ElseIf DlgType = encShowOpen Then  'Show open file
             If LenB(Filter$) = 0 Then Filter$ = "Supported formats (*.wav;*.aif;*.aiff)|*.wav;*.aif;*.aiff|All Files (*.*)|*.*"
            .DialogTitle = IIf(LenB(DialogTitle$), DialogTitle$, _
              IL(130))
            .Filter = Filter$: .FilterIndex = IdxF%
            .ShowOpen: FileFolder$ = .FileName
            IdxF% = .FilterIndex: FilterIndex% = IdxF%
        End If
        If LenB(FileFolder$) Then InitDir$ = FileFolder$
    End With
FileFolderExit:
    Set obj = Nothing: Exit Function
FileFolderErr:
   ' Debug.Print Err.Description
    Log "FileFolder$(): " & Err.Description & vbLf & FileFolder$, encLogError: FileFolder$ = "": Resume FileFolderExit
    Resume
 End Function
''// Show Open dialog, return path if selected and not canceled
Public Function ShowOpen$(Optional Filter$, Optional DlgTitle$): ShowOpen$ = FileFolder$(encShowOpen, Filter$, DlgTitle$): End Function
''// Show Save dialog, return path if selected and not canceled
Public Function ShowSave$(Optional Filter$, Optional DlgTitle$): ShowSave$ = FileFolder$(encShowSave, Filter$, DlgTitle$): End Function
''// Show folder dialog, return path if selected and not canceled
Public Function ShowFolder$(Optional DlgTitle$): ShowFolder$ = FileFolder$(encShowFolder, , DlgTitle$): End Function

'// Format milliseconds to selected format
Public Function FormatMs$(ByVal Ms&, Optional sFormat$ = "hh:nn:ss", Optional ShowDecimal As Boolean)
Const f = "hh:nn:ss"
    FormatMs$ = Format$(DateAdd("s", Ms& / 1000, #12:00:00 AM#), f)
    If sFormat$ <> f Then FormatMs$ = Format$(CDate(FormatMs$), sFormat$)
    If ShowDecimal Then FormatMs$ = FormatMs$ & "." & Format$(Ms& Mod 1000, "000")
End Function

'// Format a number like Kb, Mb, G...
Public Function FormatKMG$(Number, Optional IsByte As Boolean, Optional nDecimals%)
Dim x%, y%, s$, d$, v: Const b = "B": On Local Error Resume Next
    y% = IIf(IsByte, 1024, 1000): v = Val(Number): d$ = "# ###"
    If nDecimals% > 0 Then d$ = d$ & "." & String(nDecimals%, "0")
    Do
        If v < y% Then Exit Do
        x% = x% + 1: v = v / y%
    Loop
    s$ = Choose(x% + 1, "", "k", "M", "G", "T", "P", "E", "Z", "Y")
    If Err <> 0 Then s$ = ":-)": Err.Clear Else If Not IsByte Then s$ = UCase(s$)
    FormatKMG$ = Trim$(Format(v, d$) & " " & s$ & IIf(IsByte, b, ""))
    If FormatKMG$ = "" Then FormatKMG$ = "0"
    If FormatKMG$ = "b" Then FormatKMG$ = "0 " & b
End Function

'// MP3 ID3 Tag Stuff ******************************************************************

'// Return some info from MP3 ID3 version 1 tag stored in m_Tag
Public Function MP3TagInfoGet$(Optional ByVal Info As encID3v1TagInfo = encTagTitle, Optional RetValTagName$)
Dim s$, l&: l& = Info: If l& > 6 Then l& = 6
    If l& < 0 Or l& > 7 Then Exit Function Else RetValTagName$ = IL(50 + l&)
    If Not m_TagExist Then Exit Function       'Return Tag title only
    On Local Error Resume Next
    With m_Tag
        Select Case Info
            Case encTagTitle:   s$ = .Title
            Case encTagArtist:  s$ = .Artist
            Case encTagAlbum:   s$ = .Album
            Case encTagYear:    s$ = .Year
            Case encTagTrack:   s$ = IIf(InStr(1, .Comment, Chr(0)), _
                                  Asc(Right$(Trim$(.Comment), 1)), "0")        'Get Track# from comment
            Case encTagComment: s$ = IIf(InStr(1, .Comment, Chr(0)), _
                                  Left$(Trim$(.Comment), 28), .Comment)        'Remove Track# from comment
            Case encTagGenre:   s$ = m_GenreID3$(Val(.Genre))
            Case encTagGenreNo: s$ = .Genre
        End Select
        If Info = encTagTrack Then If s$ = "0" Then s$ = ""
    End With
    If Err Then Err.Clear: Exit Function
    l& = InStr(s$, vbNullChar): If l& Then s$ = Left$(s$, l& - 1)
    MP3TagInfoGet = Trim$(s$)
End Function
'// Reset previous read MP3-Tag
Sub MP3TagReset()
Dim Tag As encID3v1Tag: m_Tag = Tag: m_TagExist = False
End Sub
'// Return genre as string from input number (MP3 ID3 version 1 tag)
Public Function MP3GetTagGenreName$(ByVal Number As Byte)
    MP3GetTagGenreName$ = m_GenreID3$(Number)
End Function

'// Auto adding tag to created MP3
Public Sub MP3AutoAddTag(File$)
Dim s$, v$(): v$() = Split(File$, "\"): s$ = v$(UBound(v$())): v$() = Split(s$, "."): s$ = v$(0)
    MP3TagInfoLet File$, s$, App.Title, c_Title, Year(Now), 24, , c_Home: Erase v$()
End Sub

'// Read MP3 ID3 version 1 tag and store it in m_Tag
Public Function MP3ReadTagV1(File$) As Boolean
Dim x%: MP3TagReset: x% = FreeFile: On Error GoTo MP3TagV1GetErr
'Debug.Print File$
    Open File$ For Binary Access Read As #x% 'Open file for read
    'Check if file is big enough for it to contain a tag
    If LOF(x%) >= Len(m_Tag) Then
        Get #x%, LOF(x%) - Len(m_Tag) + 1, m_Tag 'Read the tag
        Close #x%: If m_Tag.Tag = c_Tag Then m_TagExist = True: MP3ReadTagV1 = True
    Else: Close #x%: End If
    Exit Function
MP3TagV1GetErr:
    Err.Clear: On Local Error Resume Next: Close #x%: Err.Clear
End Function

'// Write Tag to MP3-File
Public Function MP3TagInfoLet(File$, Optional ByVal Title$ = c_Q, Optional ByVal Artist$ = c_Q, Optional ByVal Album$ = c_Q, Optional ByVal Year$ = c_Q, Optional ByVal Genre$ = c_Q, Optional ByVal Track$ = c_Q, Optional ByVal Comment$ = c_Q, Optional AliasMCI$) As Boolean
Dim x%, l&, b As Boolean, bP As Boolean, bO As Boolean, s$, WriteTag As encID3v1Tag, ReadTag As encID3v1Tag
On Error GoTo MP3TagInfoLetErr
    'Error correction
    If Dir(File$) = vbNullString Then Exit Function
    If Len(Title$) > 30 Then Title$ = Left$(Title$, 30)
    If Len(Artist$) > 30 Then Artist$ = Left$(Artist$, 30)
    If Len(Album$) > 30 Then Album$ = Left$(Album$, 30)
    If Genre$ <> c_Q Then
        If IsNumeric(Genre$) Then
            If Val(Genre$) > 255 Or Val(Genre$) < 0 Then Genre$ = 24 '24 = Soundtrack
        Else
            For x% = 0 To 255 'If Genre$ is text check if match in m_GenreID3$()
                If Trim$(LCase(Genre$)) = LCase(m_GenreID3$(x%)) Then
                    Genre$ = x%: Exit For
                End If
            Next
            If Not IsNumeric(Genre$) Then Genre$ = 255
        End If
    End If
    x% = Val(Format(Date, "yyyy")): If Year$ <> c_Q Then _
      If Val(Year$) > x% Or Val(Year$) <= 0 Then Year$ = x%
    Comment$ = IIf(Comment$ <> c_Q, Comment$, m_Tag.Comment)
    If Len(Comment$) > 30 Then Comment$ = Left$(Comment$, 30)
    If Track$ <> c_Q Then
        If IsNumeric(Track$) Then
            If Val(Track$) > 255 Or Val(Track$) < 0 Then x% = -1
        Else: x% = -1: End If
        If x% = -1 Then Track$ = "0"
    Else: Track$ = "0": End If
    If Track$ <> "0" Then 'If there is track information to be written
        Dim sT As String * 30: sT$ = Comment$
        Comment$ = Left$(sT$, 28) & Chr(0) & Chr$(Val(Track$))
    End If
    With WriteTag 'Set up tag to be written
        .Tag = c_Tag
        .Title = IIf(Title$ <> c_Q, Title$, m_Tag.Title)
        .Artist = IIf(Artist$ <> c_Q, Artist$, m_Tag.Artist)
        .Album = IIf(Album$ <> c_Q, Album$, m_Tag.Album)
        .Year = IIf(Year$ <> c_Q, Format$(Year$, "0000"), m_Tag.Year)
        .Comment$ = Comment$
        .Genre = Val(IIf(Genre$ <> c_Q, Genre$, m_Tag.Genre))
    End With
    x% = FreeFile 'Open file for processing
'    If File$ = m_File$ Then 'Edit the open file
'        If IsOpen([Media Player]) Then 'Store where we are playing
'            B = True: l& = Val(TimePosition([Media Player])): bP = IsPlaying([Media Player])
'            CloseOne [Media Player]
'        End If
'    End If
    Open File$ For Binary As #x%
    If LOF(x%) < Len(ReadTag) Then Close #x%: Exit Function
    Get #x%, LOF(x%) - Len(ReadTag) + 1, ReadTag 'Read a tag
    If ReadTag.Tag = c_Tag Then 'Check if there is a tag
        Put #x%, LOF(x%) - Len(WriteTag) + 1, WriteTag 'Overwrite old tag
    Else 'Go to end of file and add the tag if not found
        Seek #x%, LOF(x%) + 1: Put #x%, , WriteTag
    End If
    MP3File = File$
    Close #x%
    MP3TagInfoLet = True: m_Tag = WriteTag 'Success!
MP3TagInfoLetResume:
'    If B Then 'Open and continue play
'        OpenMedia File$
'        If l& Then If bP Then Play [Media Player], l& Else SeekPosMs l&, [Media Player]
'    End If
    If bO Then MsgBoxW s$ & vbLf & File$, l&: Close #x%
    Exit Function
MP3TagInfoLetErr:
    s$ = Err.Description: l& = 16
    If Err = 75 Then s$ = IL(131): l& = 48
    bO = True: Resume MP3TagInfoLetResume
    Resume
End Function

'// Call this to add a text to the logfile
Friend Sub Log(LogText$, Optional LogType As encLogTypes = encLogEvent)
Dim s As String * 10
    Select Case LogType
        Case encLogEvent: s = "- EVENT"
        Case encLogError, encLogErrorNoMsgBox
          s = "- ERROR": If LogType = encLogError Then MsgBoxW LogText$, 16
        Case encLogBackup: s = "- BACKUP"
    End Select
    sLog$ = sLog$ & s & Now & " --> " & LogText$ & vbNewLine
End Sub

'// Save log text to log file
Private Sub LogfileSave()
Dim x%: If LenB(sLog$) = 0 Then Exit Sub
    x% = FreeFile: On Local Error Resume Next
    Open PathApp & c_LogFile For Append As #x%
    Print #x%, sLog$: Close #x%: sLog$ = ""
End Sub

'// Show the logfile in Notepad
Sub LogFileShow()
    On Local Error Resume Next: If LenB(sLog$) Then LogfileSave
    If Not FileExists(PathApp & c_LogFile) Then Exit Sub
    Shell "notepad " & PathApp & c_LogFile, vbMaximizedFocus
End Sub

'// Load the class settings from disk file and use as default if FileName$ = ""
Friend Function SettingsLoad(Optional FileName$, Optional ToLoad As Variant) As Boolean
Dim x%, s$: If LenB(FileName$) = 0 Then s$ = PathApp & c_SettingsFile Else s$ = FileName$
    On Local Error Resume Next: If Not FileExists(s$) Then Exit Function
    x% = FreeFile: Open s$ For Binary As #x%
    If LenB(FileName$) Then Get #x%, 1, ToLoad Else Get #x%, 1, m_Settings
    Close #x%: If Err Then Err.Clear: BackupFile s$: Kill s$ Else SettingsLoad = True
    If LenB(FileName$) = 0 Then _
      If m_Settings.BOM <> c_BOM Then Kill s$: SettingsLoad = False 'Error in settings file
End Function

'// Save the class settings as default to disk file, or any binary file if you like...
Public Sub SettingsSave(Optional FileName$, Optional ToSave As Variant)
Dim x%, s$: If LenB(FileName$) Then s$ = FileName$ Else _
      s$ = PathApp & c_SettingsFile: m_Settings.BOM = c_BOM
    On Local Error Resume Next: If FileExists(s$) Then Kill s$
    x% = FreeFile: Open s$ For Binary As #x%
    If LenB(FileName$) Then Put #x%, 1, ToSave Else Put #x%, 1, m_Settings
    Close #x%
End Sub


'// START CDDB & CD RIP STUFF

Public Function CDRIP(Optional ToPath$, Optional ByVal Tracks$, Optional EncodeToMp3 As Boolean = True) As Boolean
Dim TrackNo&, v$(), sE$(), s$(2), n&, nT&, StartT&, PosT&, EndT&, SizeTot&, _
  H&, l&, t&, Sz&, SizePrev&, DL$, sMsg$, sName$, nC&, b As Boolean, b2 As Boolean, _
  bMatchFound As Boolean, sT$: Const Hz = ", 44.1 kHz PCM"
    On Error GoTo CDRIPErr: DL$ = m_Settings.CDDriveLetter & ":"
    'User will select tracks to rip, and destination folder.
    If (LenB(ToPath$) = 0) Then
        If LenB(Tracks$) = 0 Then
            Tracks$ = FileFolder$(encShowOpen, "CD Tracks (*.cda)|*.cda", IL(132), DL$)
            If Tracks$ = vbNullString Then Exit Function
        End If
        ToPath$ = ShowFolder(IL(133) & ":"): If ToPath$ = vbNullString Then Exit Function
    End If
    If Not FolderExists(ToPath$) Then
        Log IL(134) & ": " & ToPath$, encLogError: Exit Function
    End If
    H& = CDInfo(encReturnHandle): If H& = -1 Then Exit Function 'Init Drive
    nT& = TOC.LastTrack - TOC.FirstTrack + 1
    If InStr(1, Tracks$, vbNullChar) Then 'Selected from Common dialog.
        v$() = Split(Tracks$, vbNullChar): n& = UBound(v$())
        For l& = 1 To n& 'Extract track number i.e Track02.cda
            v$(l& - 1) = Right$(v$(l&), 6): v$(l& - 1) = Left$(v$(l& - 1), 2)
        Next
        n& = n& - 1
    Else 'Track numbers separated with space, single track or rip all.
        If InStr(1, Tracks$, "\") Then 'Single track selected
            v$() = Split(Tracks$, "\"): n& = 0
            v$(1) = Right$(v$(1), 6): v$(0) = Left$(v$(1), 2)
        Else
            If LenB(Tracks$) = 0 Then 'RIP all tracks on the CD
                ReDim v$(nT& - 1): For l& = 0 To nT& - 1: v$(l&) = l& + 1: Next
                n& = nT& - 1
            Else
                If InStr(1, Tracks$, " ") Then
                    v$() = Split(Tracks$, " "): n& = UBound(v$())
                Else: ReDim v$(0): v$(0) = Tracks$: End If
            End If
        End If
    End If
    ReDim sE$(n&): If Right$(ToPath$, 1) <> "\" Then ToPath$ = ToPath$ & "\"
    For l = 0 To n& 'Check and calculate total lenght of tracks
        If Val(v$(l&)) <= nT& And Val(v$(l&)) > 0 Then
            CalcTrack TOC, Val(v$(l&)), StartT&, EndT&
            SizeTot& = SizeTot& + (EndT& - StartT& + 1)
        Else
            Log IL(135) & " '" & v$(l&) & "'", encLogError: v$(l&) = -1
        End If
    Next
    t& = GetTickCount: m_Settings.StopWork = False: m_Settings.IsWorking = True
    If CDDBQuery <> encFoundExactMatch Then
        ReDim InfoCDDB.tr(nT&)
    Else: bMatchFound = True: End If
    For l& = 0 To n& 'Start ripping tracks
        If v$(l&) <> -1 Then
            TrackNo& = Val(v$(l&)): CalcTrack TOC, TrackNo&, StartT&, EndT&
            If bMatchFound Then
                sT$ = " [" & InfoCDDB.tr(TrackNo&).Title & "]"
            Else: sT$ = "": End If
            sName$ = "Track" & Format$(TrackNo&, "00") & ".wav": sMsg$ = ToPath$ & sName$
            If FileExists(sMsg$) Then
                Select Case m_Settings.IfDestExist
                    Case encCreateBackup: BackupFile sMsg$
                    Case encDelete: 'Skip if locked, let system handle it :P
                    Case encStop: MsgBoxW IL(7) & ";-)" & vbLf & vbLf & _
                      sMsg$, 48: GoTo CDRIPExit
                End Select
            End If
            sMsg = IIf(EncodeToMp3, ", " & m_Settings.EncBitrate & KbpsChk$(0), "")
            sMsg$ = "ripping" & " " & UCase(DL$) & ", " & LCase(IL(54)) & " " & Format$(TrackNo&, "00") & _
              " (" & Format$(l& + 1, "00") & "/" & Format$(n& + 1, "00") & ") --> " & sName$ & sT$ & Hz & sMsg$
            InfoCDDB.tr(TrackNo&).PathWAV = ToPath$ & sName$
            If Not CDTrack2WAV(ToPath$ & sName$, _
              StartT&, EndT&, SizeTot&, SizePrev&, sMsg$, sName$, H&, EncodeToMp3, n&) Then
                If m_Settings.StopWork Then
                    sMsg$ = "CDRIP() " & IL(136) & "!": Log sMsg$
                    RaiseEvent IsWorking(100, sMsg$, encEventJobDone, sName$, 100)
                GoTo CDRIPExit
                Else: Log IL(135) & " " & TrackNo&, encLogError: End If
            Else
                If EncodeToMp3 Then
                    With m_Settings 'Start encoding each track to MP3 when one track is done.
                        b = .EncWaitUntilDone: .EncWaitUntilDone = False
                        'b2 = .FileOutAuto: .FileOutAuto = True
                        s$(0) = .EncFileIn: .EncFileIn = ToPath$ & sName$: .AscyncRIP = True
                        's$(1) = .FileOut: .FileOut = ""
                        s$(2) = .EncFolderOut: .EncFolderOut = "": Encode
                        .EncWaitUntilDone = b: .AscyncRIP = False ': .FileOutAuto = b2
                        .EncFileIn = s$(0): .EncFolderOut = s$(2) ': .FileOut = s$(1)
                    End With
                End If
            End If
            nC& = nC& + 1: Sz& = Sz& + FileLen(ToPath$ & sName$)
        End If
    Next
    CDRIP = True: sMsg$ = IL(108) & "! " & nC& & "/" & n& + 1 & " " & IL(137) & _
      Hz & " (" & FormatMs(GetTickCount - t&, c_F) & "  " & FormatKMG(Sz&, True, 2) & _
      "  " & FormatKMG(Sz& / (GetTickCount - t&), True, 2) & "/sec)"
    Log sMsg$ & vbNewLine: RaiseEvent IsWorking(100, sMsg$, encEventJobDone, "", 100)
    If EncodeToMp3 And bMatchFound Then
        sMsg$ = IL(138) & c_3D 'Wait until Blade have encode all files and then tag & rename them
        For l& = 1 To InfoCDDB.nTracks: InfoCDDB.tr(l&).PathMP3 = "": Next
        For l& = 0 To n&
            sName$ = "Track" & Format$(Val(v$(l&)), "00") & ".mp3"
            RaiseEvent IsWorking((l& / (n& + 1)) * 100, sMsg$ & " " & l& & "/" & n& + 1 & _
              " --> " & sName$, encEventWorking, sName$, 0)
            sName$ = ToPath$ & sName$
            If n& = 0 Then 'Wait until Blade start to encode if ony one track
                Do
                    If FileExists(sName$) Then Exit Do Else DoEvents
                    If m_Settings.StopWork Then GoTo Canceled
                Loop
            End If
            If FileExists(sName$) Then
                InfoCDDB.tr(l& + 1).TrackNoTmp = Val(v$(l&)): InfoCDDB.tr(l& + 1).PathMP3 = sName$
                Do
                    If m_Settings.StopWork Then 'Canceled by you
Canceled:
                        sMsg$ = "CDRIP() " & IL(136) & "!": Log sMsg$
                        RaiseEvent IsWorking(100, sMsg$, encEventJobDone, sName$, 100)
                        GoTo CDRIPExit
                    End If
                    If Not IsFileOpen(sName$) Then Exit Do Else DoEvents
                Loop
            End If
        Next
        TagRenameApply True 'Tag and rename all tracks
    Else
        ExecuteShell ToPath$ 'Open folder
    End If
CDRIPExit:
    Erase s$(): Erase v$(): Erase sE$(): CloseHandle H&: m_Settings.IsWorking = False
    Exit Function
CDRIPErr:
    Log "CDRIP() " & Err.Description, encLogError: Resume CDRIPExit
    Resume
End Function
'// RIP one track
Private Function CDTrack2WAV(ByVal TrackWavFile$, StartT&, EndT&, SizeTot&, SizePrev&, sMsg$, sName$, H&, EncodeToMp3 As Boolean, nTracks&) As Boolean
Dim Hdr As WAVCHUNKHEADER, FMT As WAVCHUNKFORMAT, INF As RAW_READ_INFO, _
  b() As Byte, OFS&, SizeT&, PosT&, Sz&, l&, P!, x%: Const CL = 27&, RSZ = 2352&
    
On Local Error GoTo CDTrack2WAVErr
    'Open the file to write the WAV-File.
    x% = FreeFile: Open TrackWavFile For Binary Access Write Lock Read As #x%
    FMT.wFormatTag = 1: FMT.wChannels = 2: FMT.dwSamplesPerSec = 44100              'Format the Tag of wave
    FMT.dwAvgBytesPerSec = 176400: FMT.wBlockAlign = 4: FMT.wBitsPerSample = 16
    Hdr.ChunkID = "RIFF": Hdr.ChunkSize = 0: Put #x%, , Hdr                         'Wave header
    Put #x%, , "WAVE": ReDim b(RSZ * CL - 1): PosT& = StartT&                       'Wave signature
    Hdr.ChunkID = "fmt ": Hdr.ChunkSize = Len(FMT): Put #x%, , Hdr: Put #x%, , FMT  'Format
    Hdr.ChunkID = "data": Hdr.ChunkSize = 0: OFS& = Seek(1): Put #x%, , Hdr         'Data
    Do While PosT& + CL < EndT&
        DoEvents: If m_Settings.StopWork = True Then Close #x%: Exit Function
        INF.DiskOffset.lowpart = PosT& * 2048&: INF.SectorCount = CL: INF.TrackMode = CDDA
        DeviceIoControl H&, &H2403E, INF, Len(INF), b(0), RSZ * CL, l&, ByVal 0&
        Sz& = Sz& + UBound(b) + 1: Put #x%, , b: PosT& = PosT& + CL
        If nTracks& > 0 Then P! = ((PosT& - StartT&) / (EndT& - StartT&)) * 100
        m_Settings.AscyncPercent = ((PosT& - StartT& + SizePrev&) / SizeTot&) * 100
        RaiseEvent IsWorking(m_Settings.AscyncPercent, IL(112) & c_3D & sMsg$, encEventWorking, _
          sName$, P!)
    Loop
    Hdr.ChunkID = "data": Hdr.ChunkSize = Sz&: Put #x%, OFS&, Hdr                       'Data header
    Hdr.ChunkID = "RIFF": Hdr.ChunkSize = LOF(1) - 4: Put #x%, 1, Hdr                   'File Header
    Close #x%: Erase b(): CDTrack2WAV = True: SizePrev& = SizePrev& + (EndT& - StartT&) 'Close file
    Log IL(108) & "! " & sMsg$: RaiseEvent IsWorking(m_Settings.AscyncPercent, _
      IL(108) & "! " & sMsg$, encEventFileDone, sName$, 100)
Exit Function
CDTrack2WAVErr:
    Log "CDTrack2WAV(): " & Err.Description, encLogError: Err.Clear
End Function

'// Try to calculate a TOC ID from a album folder with MP3 files
Private Function CDDBTocIdFromFolder$(sFolder$, Optional StartOffsetSec% = 2)
Dim s$, l&, n&, m&, v$(), v2$(), t&(2)
    v$() = DirW(sFolder$): n& = UBound(v$()) ': ReDim v2$(n&): bTrackSorting = False
    SortFilesByTrack sFolder$, v$()
    t&(1) = StartOffsetSec% * 75:  t&(1) = (t&(1) * 1000) / 75
    For l& = 0 To n&
        DoEvents: t&(2) = lTrackTimeMs&(l&)
        s$ = s$ & CLng(t&(1) / 1000) * 75 & "+": t&(1) = t&(1) + t&(2)
    Next
    v2$() = Split(s$, "+"): t&(1) = CLng(t&(1) / 1000) * 75
    For l& = 1 To n& - 1
        t&(2) = Val(v2$(l&)) \ 75
        Do While t&(2) > 0
            t&(0) = t&(0) + (t&(2) Mod 10): t&(2) = t&(2) \ 10
        Loop
    Next 'Save the CD ID to InfoCDDB.ID, and build the whole TOC query string
    InfoCDDB.id = LCase$(ZeroL(Hex$(t&(0) Mod &HFF), 2) & _
       ZeroL(Hex$((t&(1) \ 75) - (Val(v2$(0)) \ 75)), 4) & _
       ZeroL(Hex$(n& + 1), 2)): CDDBTocIdFromFolder$ = InfoCDDB.id & "+" & _
       n& + 1 & "+" & s$ & (t&(1) \ 75)
    sTmpArr$() = v$(): Erase v$(): Erase v2$(): Erase t&()
End Function


'// Access FreeCDDB.org to get CDDB info
Friend Function CDDBQuery(Optional eMail$ = c_eMailCDDB$, Optional sRetInfo$, Optional sRetMatches$, Optional TocID$, Optional CDID$, Optional Category$, Optional ArrFileNames, Optional bInfo As Boolean, Optional RetStrResult$, Optional bDetails As Boolean) As encResponseCDDB
Dim s$, q$, P$, d#, sC$, e$(), v$(), x%, y%, l&, n&, CI As encInfoCDDB, b As Boolean
    s$ = InfoCDDB.id: InfoCDDB = CI: InfoCDDB.id = s$: On Error GoTo CDDBQueryErr
    If eMail$ = c_eMailCDDB$ Then eMail$ = m_eMailCDDB$
    s$ = eMail$ & "@" & "clsBladeEnc" & "@1.0.0@" & "freedb.freedb.org" 'Create TOC query string
    e$() = Split(s$, "@"): If LenB(TocID$) Then s$ = TocID$ _
      Else: s$ = CDInfo(encReturnQueryS): InfoCDDB.TocID = s$ 'Get TocID from CD instead of input
    'Debug.Print s
    e$(4) = "http://" & e$(4) & "/~cddb/cddb.cgi/?cmd=": P$ = PathApp & "~freedb.tmp"
    e$(0) = "&hello=" & e$(0) & "+" & e$(1) & "+" & e$(2) & "+" & e$(3) & "&proto=5"
    q$ = e$(4) & "cddb+query+" & s$ & e$(0)
    'If LenB(CDID$) = 0 Then
        RaiseEvent IsWorking(0, "Wait" & c_3D & "Sending request to CDDB database" & c_3D, encEventWorking, "", 0)
    'End If
    If LenB(CDID$) Then sC$ = Category$: e$(3) = CDID$: GoTo ShowDetails
    If DownloadFile(q$, P$) Then
        x% = FreeFile: Open P$ For Input As #x% 'Read result of query string
        s$ = Input(LOF(x%), #x%): Close #x%: CDDBQuery = Val(Left$(s$, 3))
        sRetInfo$ = q$ & vbNewLine & vbNewLine & s$ & vbNewLine
        If CDDBQuery = encFoundInExactMatch Then 'Inexact match
            v$() = Split(s$, vbNewLine)
            For x% = 1 To UBound(v$())
                If LenB(v$(x%)) > 4 Then
                    sRetMatches = sRetMatches & _
                      v$(x%) & vbNewLine: y% = y% + 1
                End If
            Next
            If y% = 1 Then 'Only one match in the inexact list so continue with that one
                v$() = Split(v$(1), " "): sC$ = Trim$(v$(0)): e$(3) = v$(1)
                CDDBQuery = encFoundExactMatch
            Else 'Let the user select one...
             RaiseEvent IsWorking(0, IL(139), encEventWorking, "", 0)
                v$(0) = ShowCDDBMatches$(IL(140) & _
                  vbNewLine & IL(141) & "e: " & InfoCDDB.id, sRetMatches)
                If v$(0) = "" Then GoTo CDDBQueryExit Else CDDBQuery = encFoundExactMatch
                v$() = Split(v$(Val(v$(0))), " "): sC$ = Trim$(v$(0)): e$(3) = v$(1)
            End If
        ElseIf CDDBQuery = encFoundExactMatch Then 'Excact match
            v$() = Split(s$, " "): sC$ = Trim$(v$(1)): e$(3) = v$(2)
            sRetMatches = Right$(s$, Len(s$) - 4)
        Else: GoTo CDDBQueryExit: End If
        
ShowDetails:
        
        q$ = e$(4) & "cddb+read+" & sC$ & "+" & e$(3) & e$(0) ': Debug.Print q$
        If DownloadFile(q$, P$) Then
            x% = FreeFile: Open P$ For Input As #x%: sTmpStr$ = Input(LOF(x%), #x%)
            s$ = sTmpStr$: CDDBQuery = Val(Left$(s$, 3)): Close #x%: b = False
            sRetInfo$ = sRetInfo$ & vbNewLine & q$ & vbNewLine & vbNewLine & s$
            If CDDBQuery >= encErrorQuery Then GoTo CDDBQueryExit Else n& = 0: Open P$ For Input As #x%
            e$() = Split(" ;Disc ID: ;Artist: ;Album: ;Year: ;Categ.: ;Genre: ;Track;Extra: ;Length;Start;End;MilliSec.;Title", ";")
            Do While Not EOF(x%) 'Fill RetInfo$ and the private type InfoCDDB with info
                With InfoCDDB: Line Input #x%, s$
                    If InStr(1, s$, "Track frame offsets") Or b Then
                        If b Then 'Read offsets and calculate track lenght in milliseconds
                            If LenB(s$) > 2 Then
                                ReDim Preserve .tr(n&): .tr(n&).OffSet = Val(Trim$(Right$(s$, Len(s$) - 1)))
                                If n& > 0 Then
                                    d# = (.tr(n& - 1).OffSet / 75) * 1000: .tr(n&).StartT = FormatMs(d#, "nn:ss")
                                    .tr(n&).LengthMs = ((.tr(n&).OffSet / 75) * 1000) - d#
                                    .tr(n&).LengthT = FormatMs(.tr(n&).LengthMs, "nn:ss")
                                    .tr(n&).StopT = FormatMs(d# + .tr(n&).LengthMs, "nn:ss")
                                End If
                                n& = n& + 1
                            Else: b = False: End If
                        Else: b = True: End If
                    End If
                    If InStr(1, s$, "Disc length") Then 'Read disc length in seconds to get last track length
                        v$() = Split(s$, ":"): v$() = Split(v$(1), " "): ReDim Preserve .tr(n&)
                        d# = (.tr(n& - 1).OffSet / 75) * 1000: .tr(n&).StartT = FormatMs(d#, "nn:ss")
                        .tr(n&).LengthMs = (Val(v$(1)) - (.tr(n& - 1).OffSet / 75)) * 1000
                        .tr(n&).StopT = FormatMs(d# + .tr(n&).LengthMs, "nn:ss")
                        .tr(n&).LengthT = FormatMs(.tr(n&).LengthMs, "nn:ss"): n& = 1
                    End If
                    If InStr(1, s$, "=") Then
                        v$() = Split(s$, "="): b = False
                        Select Case v$(0)
                            Case "DISCID": .id = v$(1): e$(0) = e$(1) & vbTab & v$(1) & vbNewLine
                            Case "DTITLE"
                                v$() = Split(v$(1), "/"): CDDBQuery = encFoundExactMatch
                                .Artist = Trim$(v$(0)): e$(0) = e$(0) & e$(2) & vbTab & Trim$(v$(0)) & vbNewLine
                                If UBound(v$()) > 0 Then 'If no / char separating Artist from / Album
                                    .Album = Trim$(v$(1)): e$(0) = e$(0) & e$(3) & vbTab & Trim$(v$(1)) & vbNewLine
                                End If
                            Case "DYEAR": .Year = v$(1): e$(0) = e$(0) & e$(4) & vbTab & v$(1) & vbNewLine
                            Case "DGENRE"
                                .Category = sC$: e$(0) = e$(0) & e$(5) & vbTab & sC$ & vbNewLine
                                .Genre = v$(1): e$(0) = e$(0) & e$(6) & vbTab & v$(1) & vbNewLine
                            Case "EXTD":  .TitleX = .TitleX & v$(1)
                            Case Else 'Track stuff...
                                If CBool(LenB(v$(1))) Then 'It is some text to the right of =
                                    If UBound(.tr) < n& Then ReDim Preserve .tr(n&)
                                    If CBool(InStr(1, s$, "TTITLE")) Then    'Track title
                                        .tr(n&).Title = v$(1): n& = n& + 1 'Counter tracks
                                    ElseIf CBool(InStr(1, s$, "EXTT")) Then  'Extra info track
                                        y% = Mid$(v$(0), 5) + 1: .tr(y%).TitleX = .tr(y%).TitleX & v$(1)
                                    End If
                                End If
                        End Select
                    End If
                End With
            Loop
            With InfoCDDB: Close #x%
              If IsArray(ArrFileNames) Then v$() = ArrFileNames: b = True Else b = False
              e$(0) = e$(0) & IIf(LenB(.TitleX), e$(8) & vbTab & .TitleX & vbNewLine, "") & vbNewLine
              SendString$ "open cdaudio alias x wait shareable"
              SendString$ "set x time format 0" 'Let MCI check track length in Ms
                If LenB(TocID$) = 0 And LenB(CDID$) = 0 And Not bInfo Then e$(0) = e$(0) & e$(7) & vbTab & e$(9) & vbTab & e$(12) & _
                  vbTab & e$(10) & vbTab & e$(11) & vbTab & e$(13) & vbNewLine & vbNewLine
              For x% = 1 To n& - 1 'Add the track info
                If LenB(TocID$) = 0 And LenB(CDID$) = 0 And Not bInfo Then 'Calculation on CD if not using TocID$ from folder.
                  .tr(x%).LengthT = CDInfo(encTrackLength, x%)
                  .tr(x%).StartT = CDInfo(encTrackStart, x%)
                  .tr(x%).StopT = CDInfo(encTrackStop, x%)
                  l& = Val(SendString("status x length track " & x% & " wait"))
                  .tr(x%).LengthMs = l&: .LengthMs = .LengthMs + l&
                  e$(0) = e$(0) & Format$(x%, "00") & vbTab & .tr(x%).LengthT & _
                    vbTab & l& & vbTab & .tr(x%).StartT & vbTab & .tr(x%).StopT & vbTab & _
                    Trim$(.tr(x%).Title) & IIf(LenB(.tr(x%).TitleX), vbNewLine & e$(8) & vbTab & _
                    Replace(.tr(x%).TitleX, "\n", ", ") & vbNewLine, "") & vbNewLine
                Else 'No calculation on CD if using TocID$ from folder, just track & title..
                    If b Then
                        Dim sT As String * 20 's$ = String(50, " "):
                        sT = Left$(Trim$(.tr(x%).Title), 18)
                        If x% - 1 > UBound(v$()) Then
                            s$ = IL(46) & " " & n& - UBound(v$()) - 2 & " " & IL(142) & "!"
                        Else: s$ = v$(x% - 1): End If
                        q$ = "  " & sT & vbTab & " --> " & vbTab & s$
                    Else: q$ = Trim$(.tr(x%).Title): End If
                    e$(0) = e$(0) & "#" & Format$(x%, "00") & "  " & FormatMs(.tr(x%).LengthMs, "nn:ss") & q$ & vbNewLine
                End If
              Next
              If LenB(TocID$) Then e$(0) = Left$(e$(0), Len(e$(0)) - 2) & "."
              If UBound(v$()) > n& - 2 Then 'More files in selected folder than tracks
                    For x% = n& - 1 To UBound(v$())
                        If x% < 0 Then Exit For
                        e$(0) = e$(0) & IL(46) & vbTab & v$(x%) & vbNewLine
                    Next
              End If
              If n& > 1 Then
                sRetInfo$ = e$(0): sTmpStr$ = sTmpStr$ & String(3, vbNewLine) & e$(0)
                SendString$ "close x": .FoundCDDB = CDDBQuery: .nTracks = n& - 1
              End If
            End With
        Else: CDDBQuery = encInternetNotConected: End If
    Else: CDDBQuery = encInternetNotConected: End If
CDDBQueryExit:
    Erase e$(): Erase v$(): InfoCDDB.RetInfo = sRetInfo$: If FileExists(P$) Then Kill P$
    b = False: s$ = IL(150) & ":" & " " & s$
    Select Case CDDBQuery
        Case encCanceledByUser:         s$ = IL(136)
        Case encFoundExactMatch:        s$ = IL(143): b = True
        Case encFoundInExactMatch:      s$ = IL(139): b = True
        Case encNoMatchFound:           s$ = IL(144)
        Case Is >= encErrorQuery:       s$ = s$ & IL(145)
        Case encNoHandshake:            s$ = s$ & IL(146)
        Case encDatabaseEntryCorrupt:   s$ = s$ & IL(147)
        Case encNoEntryInDatabase:      s$ = IL(148)
        Case encInternetNotConected:    s$ = s$ & IL(149)
        Case Else:                      s$ = ""
    End Select
    RetStrResult$ = s$ & ". :-" & IIf(b, ")", "("): InfoCDDB.MsgReturn = RetStrResult$
    'If frmMX3.Visible Then frmMX3.Hide 'This event won't fire when frminput is modal:S
    RaiseEvent IsWorking(0, RetStrResult$, encEventInfo, "", 0)
    Exit Function
CDDBQueryErr:
    If Err = 55 Then Resume Next 'Temp file is open
    sRetInfo$ = "CDDBQuery(): " & Err.Description: InfoCDDB.RetInfo = sRetInfo$
    RaiseEvent IsWorking(0, sRetInfo$, encEventError, "", 0)
    Log sRetInfo$, encLogError: CDDBQuery = encNoMatchFound: Err.Clear: Exit Function
    Resume
End Function

'// Will return some mixed info about the CD drive, returns -1 if fail
Public Function CDInfo(Optional ByVal Info As encInfoCD = encTracksCount, Optional ByVal TrackNo& = 1)
Dim DL$, s$, l&, v$(), n&, t&(8), H&, StartT&, EndT&: CDInfo = -1
    'Init Drive
    DL$ = m_Settings.CDDriveLetter & ":"
    H& = CreateFile("\\.\" & DL$, &H80000000, &H1, ByVal 0&, 3, &H80, 0)
    If (H& = -1) Then Log "Can't access " & DL$, encLogError: Exit Function
    'Table of Content
    If (DeviceIoControl(H&, &H24000, ByVal 0&, 0, TOC, Len(TOC), l&, ByVal 0&) = 0) Then
        Log "Can't read " & DL$, encLogError: CloseHandle H&: ShowSelectCD: Exit Function
    End If
    n& = TOC.LastTrack - TOC.FirstTrack + 1
    If TrackNo& <= n& And TrackNo& > 0 Then
        CalcTrack TOC, TrackNo&, StartT&, EndT&
    Else: CloseHandle H&: Exit Function: End If
    For l& = 1 To 3: t&(l&) = TOC.TrackData(TrackNo&).Address(l&): Next
    TrackNo& = TrackNo& - 1
    For l& = 1 To 3: t&(l& + 3) = TOC.TrackData(TrackNo&).Address(l&): Next
    Select Case Info
        Case encTracksCount:    CDInfo = n&
        Case encLenght:         CDInfo = TOC.Length
        Case encLenghtMs:       CDInfo = InfoCDDB.LengthMs
        Case encTrackLengthMs:  CDInfo = InfoCDDB.tr(TrackNo& + 1).LengthMs
        Case encTrackStartLBA:  CDInfo = StartT&
        Case encTrackStopLBA:   CDInfo = EndT&
        Case encTrackLenghtLBA: CDInfo = (EndT& - StartT& + 1)
        Case encTrackStopMin:   CDInfo = t&(1)
        Case encTrackStopSec:   CDInfo = t&(2)
        Case encTrackStopFra:   CDInfo = t&(3)
        Case encTrackStartMin:  CDInfo = t&(4)
        Case encTrackStartSec:  CDInfo = t&(5)
        Case encTrackStartFra:  CDInfo = t&(6)
        Case encTrackStart:     CDInfo = Format$(t&(4), "00") & ":" & Format$(t&(5), "00")
        Case encTrackStop:      CDInfo = Format$(t&(1), "00") & ":" & Format$(t&(2), "00")
        Case encTrackLength:    CDInfo = FormatMs((((t&(1) - t&(4)) * 60) + _
                                  (t&(2) - t&(5))) * 1000, c_F)
        Case encReturnHandle:   CDInfo = H&
        Case encReturnQueryS 'Create TOC query string to FreeDB
            For l& = 0 To n&
                t&(6) = TOC.TrackData(l&).Address(1): t&(7) = TOC.TrackData(l&).Address(2)
                s$ = s$ & (t&(6) * 60 * 75) + (t&(7) * 75) + TOC.TrackData(l&).Address(3) & ";"
            Next
            v$() = Split(s$, ";"): s$ = "": InfoCDDB.TOC = TOC
            For l& = 0 To n& - 1
                t&(7) = Val(v$(l&)) \ 75: s$ = s$ & v$(l&) & "+"
                Do While t&(7) > 0
                    t&(8) = t&(8) + (t&(7) Mod 10): t&(7) = t&(7) \ 10
                Loop
            Next 'Save the CD ID to InfoCDDB.ID, and return the whole query string
            InfoCDDB.id = LCase$(ZeroL(Hex$(t&(8) Mod &HFF), 2) & _
               ZeroL(Hex$((Val(v$(n&)) \ 75) - (Val(v$(0)) \ 75)), 4) & _
               ZeroL(Hex$(n&), 2)): CDInfo = InfoCDDB.id & "+" & _
               n& & "+" & s$ & (Val(v$(n&)) \ 75)
    End Select
    Erase t&(): Erase v$(): If Info <> encReturnHandle Then CloseHandle H&
End Function

'// Try to find out genre to use in MP3 tag, with data  from CDDB
Friend Function GenreFromCDDB%()
Dim x%, y%, G%
    With InfoCDDB
        For x% = 0 To 255 'Try to find the genere
            If Trim$(LCase(m_GenreID3$(x%)) = Trim$(LCase(.Genre))) Then G% = x%: Exit For
            If InStr(1, LCase(m_GenreID3$(x%)), LCase(.Genre)) Then If y% = 0 Then y% = x%
        Next
        If G% + y% = 0 Then 'Nothing found in genre so try category instead
            For x% = 0 To 255
                If Trim$(LCase(m_GenreID3$(x%)) = Trim$(LCase(.Category))) Then G% = x%: Exit For
                If InStr(1, LCase(m_GenreID3$(x%)), LCase(.Category)) Then If y% = 0 Then y% = x%
            Next
        End If
        GenreFromCDDB% = IIf(G% = 0, IIf(y% = 0, 255, y%), G%)
        If GenreFromCDDB% > 147 Then GenreFromCDDB% = 255
    End With
End Function
'// Format Minutes, Seconds and Frames to LBA.
Private Function FormatLBA(ByVal Minutes As Long, ByVal Seconds As Long, ByVal Frames As Long) As Long
    FormatLBA = ((60 * 75 * (Minutes)) + (75 * (Seconds)) + ((Frames) - 150))
End Function

'// Calculate start and end position on selected track
Private Sub CalcTrack(TOC As CDROM_TOC, ByVal TrackNo&, RetStartT&, RetEndT&)
    With TOC
        TrackNo& = TrackNo& - 1: RetStartT& = FormatLBA(.TrackData(TrackNo&).Address(1), _
          .TrackData(TrackNo&).Address(2), .TrackData(TrackNo&).Address(3))
        RetEndT& = FormatLBA(.TrackData(TrackNo& + 1).Address(1), _
          .TrackData(TrackNo& + 1).Address(2), .TrackData(TrackNo& + 1).Address(3))
    End With
End Sub

'// END CDDB & RIP STUFF

'// Tag and rename a folder with a MP3 abum.
Public Sub TagRename(Optional QueryType As encCDDCQuerys = [Audio CD as source], Optional sFolder$)
Dim s$, f$, e$, m$, sCap$, l&, n&, v$(), t&(2), RetQ As encResponseCDDB, bRetry As Boolean
On Local Error GoTo TagRenameExit
    If LenB(sFolder$) = 0 Then
        If LenB(sTmpFolder$) = 0 Then
            sFolder$ = FileFolder(encShowFolder, , IL(151))
        Else: sFolder$ = sTmpFolder$: End If
    End If
    If LenB(sFolder$) Then
        If Right$(sFolder$, 1) <> "\" Then sFolder$ = sFolder$ & "\"
        sCap$ = sFolder$ & "*.mp3"
    Else: Exit Sub: End If
    sTmpFolder$ = sFolder$: ReDim lTrackTimeMs&(0)
    'Calculate TOC ID from files in a folder
    If QueryType = [MP3 album folder as source] Then
        f$ = CDDBTocIdFromFolder$(sFolder$)
        RetQ = CDDBQuery(, s$, , f$, , , sTmpArr$(), True, e$)
    ElseIf QueryType = [Audio CD as source] Then 'Calculate TOC ID from audio CD
        sTmpArr$() = DirW(sFolder$): RetQ = CDDBQuery(, s$, , , , , sTmpArr$(), True, e$)
        If Right$(s$, 2) = vbNewLine Then s$ = Left$(s$, Len(s$) - 2)
        sCap$ = CDDriveLetter & ":\ --> " & sCap$
    Else                                         'By using input CDID and category
RetryCDID:
        If LenB(sTmpStr$) = 0 Then
            's$ = Me.sh
            s$ = frm(frmSearchCDDB_CDID).ShowSearchCDDB: If LenB(s$) = 0 Then Exit Sub
        Else: s$ = sTmpStr$: End If
        v$() = Split(s$, ";"): sTmpArr$() = DirW(sFolder$)
        RetQ = CDDBQuery(, s$, , , v$(0), v$(1), sTmpArr$(), True, e$)
        If Right$(s$, 2) = vbNewLine Then s$ = Left$(s$, Len(s$) - 2)
        If LenB(sTmpStr$) = 0 And RetQ = encNoMatchFound Or _
          RetQ = encNoEntryInDatabase Or RetQ >= encErrorQuery Then
            m$ = vbLf & e$ & vbLf & vbLf & IL(153)
            If MsgBoxW(IL(152) & " (" & v$(1) & " '" & v$(0) & "')." & _
              m$, vbYesNo + vbExclamation) = vbYes Then
                sTmpStr$ = "": GoTo RetryCDID
            End If
        End If
    End If
    RaiseEvent IsWorking(0, e$, encEventWorking, "", 0)
    If RetQ = encFoundExactMatch Then 'Match found in CDDB
        frm(frmTagRename).ShowTagRename s$, sCap$
    End If
TagRenameExit:
    If Err Then e$ = Err.Description: Err = 0
    RaiseEvent IsWorking(0, e$, encEventWorking, "", 0)
    Erase sTmpArr$(): Erase v$(): Erase t&(): sTmpFolder$ = "": Exit Sub
    Resume
End Sub

'// Tag and rename check tracks
Function TagRenameCheck(sFo$)
Dim x%, y%, t%, sMsg$, bSL As Boolean, n%, s$, m$, sT$, sSL$, l&, lDiff&, nT%, sF$, v$(), v2$(), sMP3$()
    With InfoCDDB
        'Get all MP3 files in selected folder
        If UBound(lTrackTimeMs&()) = 0 Then
            SortFilesByTrack sFo$, sTmpArr$(): sMP3$() = sTmpArr$()
        Else: sMP3$() = sTmpArr$(): End If
        RaiseEvent IsWorking(0, .MsgReturn, encEventJobDone, "", 0)
        y% = UBound(sMP3$()): nT% = 1: ReDim aDiffMs&(y% + 1)
        If LenB(sMP3$(0)) = 0 Then MsgBoxW IL(41) & vbNewLine & sFo$, 48: Exit Function
        If y% + 1 <> .nTracks Then 'More/Less MP3-Files than tracks in selected folder
            sMsg$ = IL(18) & " " & .nTracks & " " & IL(19) & " " & _
              y% + 1 & " " & IL(20) & vbNewLine & sFo$ & vbNewLine
        End If
        'Compair times on Audio CD or times from CDDB info found with files in selected folder.
        For x% = 0 To y% 'Loop all files found in the selected folder.
            l& = lTrackTimeMs&(x%)
            For t% = 1 To .nTracks               'Loop tracks to find matching length from CDDB
                lDiff& = .tr(t%).LengthMs
                lDiff& = lDiff& - l&: If lDiff& < 0 Then lDiff& = lDiff& * -1
                If t% = x% + 1 Then aDiffMs&(t%) = lDiff&
                If x% = y% And t% = .nTracks Then  'Last track so compensate for start offset
                    If lDiff& < ((.tr(0).OffSet / 75) * 1000) Then lDiff& = m_Settings.MP3MaxDiffMs
                End If
                If lDiff& <= m_Settings.MP3MaxDiffMs Then  'Found matching length
                    v$() = Split(sT$, ";"): bSL = False
                    For n% = 0 To UBound(v$()) - 1 'Check if already a captured track
                        If Val(v$(n%)) = t% Then
                            bSL = True: sSL$ = sSL$ & Format$(t%, "00") & _
                              "|" & sMP3$(x%) & "|" & l& & ";"
                        End If
                    Next
                    If bSL Then Exit For
                    .tr(nT%).TrackNoTmp = t%: .tr(nT%).PathMP3 = sFo$ & sMP3$(x%)
                    sT$ = sT$ & Format(t%, "00") & ";": nT% = nT% + 1
                End If
            Next
        Next
      '  Debug.Print sT$: Debug.Print sSL$
        If nT% - 1 <> .nTracks Or LenB(sSL$) Then     'Match found but not all, or simular length
            On Local Error Resume Next
            sT$ = Left$(sT$, Len(sT$) - 1): v$() = Split(sT$, ";"): v2$() = sMP3$()
            n% = UBound(v$())
            For x% = 0 To n%
                sMP3$(x%) = v2$(Val(v$(x%)) - 1)
            Next
            For x% = 0 To UBound(v2$())
                bSL = False
                For y% = 0 To UBound(v$())
                    If v2$(x%) = sMP3$(y%) Then bSL = True: Exit For
                Next
                If Not bSL Then
                    n% = n% + 1: sMP3$(n%) = v2$(x%)
                End If
            Next
            sTmpArr$() = sMP3$()
            If LenB(sT$) And nT% - 1 <> .nTracks Then 'Match found but not all
                s$ = "": SortArray v$(): l& = 292
                For x% = 0 To UBound(v$())
                    y% = Val(v$(x%))
                    s$ = s$ & Format(y%, "00") & "." & vbTab & .tr(y%).Title & vbNewLine
                Next
                y% = nT% - 1: m$ = IL(22) & " " & nT% - 1 & "/" & _
                  .nTracks & " " & IL(23) & "."
                sMsg$ = sMsg$ & m$: s$ = m$ & vbNewLine & vbNewLine & s$ & vbNewLine
            Else: s$ = "": End If
            If LenB(sSL$) Then 'Have found tracks that have simular length.
                sSL$ = Left$(sSL$, Len(sSL$) - 1): v$() = Split(sSL$, ";")
                sF$ = "": SortArray v$(): l& = 308
                For x% = 0 To UBound(v$())
                    v2$() = Split(v$(x%), "|"): y% = Val(v2$(0))
                    For t% = 1 To nT% - 1
                        If .tr(t%).TrackNoTmp = y% Then Exit For
                    Next
                    n% = Len(.tr(y%).Title): m$ = Left$(.tr(y%).Title, 15)
                    If 15 - n% > 0 Then m$ = m$ & String(15 - n%, " ") Else m$ = m$ & String(3, ".")
                    sF$ = sF$ & m$ & vbTab & "Track " & v2$(0) & " <-- " & _
                      FormatMs(.tr(y%).LengthMs, c_F, True) & " <--> " & _
                      FormatMs(Val(v2$(2)), c_F, True) & " --> " & v2$(1) & " {" & _
                      .tr(y%).LengthMs - Val(v2$(2)) & "}" & vbNewLine
                Next
                sF$ = vbNewLine & IL(24) & vbNewLine & IL(25) & _
                  "' " & vbNewLine & vbNewLine & IL(26) & "." & _
                  vbNewLine & vbNewLine & sF$ & vbNewLine
            Else: sF$ = "": End If 'Display the result to user...
            If LenB(s$) + LenB(sF$) = 0 Then l& = 48: s$ = IL(27) & _
              ", " & IL(28) & "." & vbNewLine & vbNewLine: sMsg$ = s$ & sMsg$
            
            sSL$ = "": m$ = IL(29) & _
              " '" & m_Settings.MP3MaxDiffMs & "', " & IL(30) & " " & _
              IIf(l& = 308, IL(31), IL(32)) & " " & IL(33) & "."
            s$ = s$ & sF$ & m$: If LenB(sMsg$) Then _
              sMsg$ = sMsg$ & vbLf & vbLf & m$ & vbLf & IL(154)
            TagRenameCheck = s$ & vbNewLine & vbNewLine
            If LenB(sMsg$) Or l& = 48 Then MsgBoxW sMsg$
        Else: sTmpArr$() = sMP3$(): End If
    End With
TagRenameApplyExit:
End Function

'// Tag and rename sort files by track if possible and get file length in milliseconds
Private Sub SortFilesByTrack(sFolder$, Files() As String)
Dim s$, m&, n&, l&, v$(), v2$()
    v$() = Files(): n& = UBound(v$()): ReDim v2$(n&): ReDim lTrackTimeMs&(n&) ': bTrackSorting = False
    For l& = 0 To n& 'Read MP3 track# in tag
        MP3ReadTagV1 sFolder$ & v$(l&): v2$(l&) = MP3TagInfoGet(encTagTrack)
        If IsNumeric(v2$(l&)) Then m& = m& + 1
    Next
    If m& - 1 = n& Then 'Sort files by track#
        For l& = 0 To n&
            For m& = 0 To n&
                If Val(v2$(m&)) = (l& + 1) Then s$ = s$ & v$(m&) & ";": Exit For
            Next
        Next
        If LenB(s$) = 0 Then Exit Sub
        s$ = Left$(s$, Len(s$) - 1)
        Files() = Split(s$, ";")
    End If
    For l& = 0 To n& 'Read MP3 track# in tag
        RaiseEvent IsWorking(((l& / (n& + 1))) * 100, IL(112) & c_3D & IL(155), encEventWorking, "", 0)
        If UBound(Files()) < l& Then Exit For
        lTrackTimeMs&(l&) = MP3LenMs(sFolder$ & Files(l&))
    Next
    Erase v$(): Erase v2$()
End Sub

'// Apply the track info collected from FreeCDDB, to a whole folder or a single MP3 file if TrackNumber% > 0
Friend Function TagRenameApply(Optional IsRipped As Boolean) As Boolean
Dim x%, y%, t%, G%, n&, m$, s$, sF$, sPrevF$, sPL$, sSL$, sT$, sFo$, b As Boolean: Const c = " - "
    If Not InfoCDDB.FoundCDDB Then Exit Function Else On Error GoTo TagRenameApplyErr
    With InfoCDDB
        G% = GenreFromCDDB%
        For x% = 1 To .nTracks 'Start adding tag and rename
            t% = .tr(x%).TrackNoTmp: sF$ = .tr(x%).PathMP3
            If LenB(sF$) Then
                b = MP3TagInfoLet(sF$, .tr(t%).Title, .Artist, .Album, _
                  .Year, G%, t%, c_Home) 'Add MP3 tag
                y% = y% + 1: sPrevF$ = sF$: If b Then n& = n& + 1
                If b And m_Settings.MP3FileFormat <> 0 Then 'Rename the file
                    Select Case m_Settings.MP3FileFormat
                        Case 1: s$ = ZeroL(t%, 2) & c & .Artist & c & .Year & c & .Album & c & .tr(t%).Title
                        Case 2: s$ = ZeroL(t%, 2) & c & .Artist & c & .Album & c & .tr(t%).Title
                        Case 3: s$ = ZeroL(t%, 2) & c & .Artist & c & .tr(t%).Title
                        Case 4: s$ = ZeroL(t%, 2) & c & .tr(t%).Title
                        Case 5: s$ = .Artist & c & .Year & c & .Album & c & .tr(t%).Title
                        Case 6: s$ = .Artist & c & .Album & c & .tr(t%).Title
                        Case 7: s$ = .Artist & c & .tr(t%).Title
                        Case 8: s$ = .tr(t%).Title
                    End Select
                    sFo$ = PathInfo(sF$, encPathBS)
                    s$ = FileReplBad$(s$): s$ = sFo$ & s$ & ".mp3"  'Check and create the new file name.
                    If s$ <> sF$ Then 'No need to backup if same name as before
                        If FileExists(s$) Then                          'Check if destination file exist.
                            Select Case IfDestExist                     'Force copy if sF$ = s$
                                Case encCreateBackup: BackupFile s$, IIf(sF$ = s$, True, False)
                                Case encDelete: Kill s$
                                Case encStop: MsgBoxW IL(7) & ";-)" & vbLf & vbLf & _
                                  s$, 48: GoTo TagRenameApplyExit
                            End Select
                        End If
                        Name sF$ As s$: sF$ = s$: sPrevF$ = s$: MP3File = sPrevF$
                        .tr(t%).ListName = PathInfo(s$, encNameExt)
                    End If
                Else
                    If m_Settings.MP3FileFormat <> 0 Then sSL$ = sSL$ & IL(35) & " '" & _
                      sF$ & "' (" & FormatKMG(FileLen(sF$), True, 2) & ")." & vbLf
                End If
            End If
        Next
        sT$ = IL(108) & "! ": s$ = IL(36)        'Create playlist
        If m_Settings.CreatePlaylist And Not IsRipped Then sPL$ = CreatePlaylist$(sFo$, .Artist & " - " & .Album)
        m$ = " " & IIf(m_Settings.MP3FileFormat, " " & IL(37), IL(9)) & "."
        If y% <> 1 Then 'Done! Ask if user want to open forder/play the song.
            If n& <> y% Then x% = 52 Else x% = 36
            m$ = n& & "/" & y% & " " & LCase(IL(38)) & IIf(IsRipped, " " & IL(49), "") & m$
            s$ = sT$ & m$ & vbLf & vbLf & s$ & " " & IL(43) & vbLf & sFo$: sT$ = sFo$
        Else
            m$ = IL(40) & " " & IL(39) & IIf(IsRipped, " " & IL(49), "") & m$: x% = 36
            
            s$ = sT$ & m$ & vbLf & vbLf & _
            s$ & " " & IL(44) & " '" & PathInfo(sPrevF$, encNameExt) & "'?": sT$ = sPrevF$
        End If
        RaiseEvent IsWorking(100, m$, encEventJobDone, "", 100)
        If LenB(sSL$) Then s$ = sSL$: x% = 48
        b = True: If MsgBoxW(s$, x%) = vbYes Then ExecuteShell sT$
TagRenameApplyExit:
        TagRenameApply = b
    End With
    Exit Function
TagRenameApplyErr:
    Log "TagRenameApply(): " & Err.Description & " --> " & _
      sF$, encLogError: Err.Clear: Resume TagRenameApplyExit
    Resume
End Function



'// SHOW FORMS START

'// Show the MP3 Tagger window form
Public Sub ShowTaggerMP3(Optional OwnerForm As Object, Optional MP3File$)
Dim FT As mx3FormTypes: FT = frmTaggerForm
    With frm(FT)
        .FormType = FT: .WindowState = 0
        .ShowTagger OwnerForm, MP3File$
    End With
End Sub
'// Show the File size calculator form
Public Sub ShowFileSizeCalc(Optional OwnerForm As Object, Optional FilePath$, Optional Bitrate As encEncodeBitRates = [128 Kbps])
Dim FT As mx3FormTypes: FT = frmCalcForm
    With frm(FT)
        .FormType = FT: .WindowState = 0: .ShowFileCalc OwnerForm
        .CalcFilePath = FilePath$: .CalcBitrate = Bitrate
    End With
End Sub
'// Show the File size calculator form
Public Function ShowFreeSearch$(Optional OwnerForm As Object, Optional SearchText$)
Dim FT As mx3FormTypes: FT = frmSearchCDDB_Free
    With frm(FT)
        .FormType = FT: .WindowState = 0
        ShowFreeSearch$ = .ShowFreeSearch(OwnerForm, SearchText$)
    End With
End Function
'// Show input form (OwnerForm As Object, Promt$, Title$, Default$)
Public Function ShowInput$(Promt$, Title$, Optional Default$, Optional OwnerForm As Object, Optional OnlyNumbers As Boolean, Optional MaxLength& = 0, Optional Alignment As txtAlignment = txtLeft)
Dim FT As mx3FormTypes: FT = frmInput
    With frm(FT)
        .FormType = FT: ShowInput$ = .ShowInputForm(OwnerForm, Promt$, Title$, Default$, OnlyNumbers, MaxLength&, Alignment)
    End With
End Function
'// Show genre name from number
Public Sub ShowGenreFromNumber(Optional OwnerForm As Object)
Dim FT As mx3FormTypes: FT = frmGenreFromNumber
    With frm(FT)
        .FormType = FT: .ShowInputForm OwnerForm, IL(175) & " (0-255):", _
          IL(176), "37", True, 3, txtRight
    End With
End Sub
'// Show genre number from name
Public Sub ShowGenreNumberFromName(Optional OwnerForm As Object)
Dim FT As mx3FormTypes: FT = frmGenreNumberFromName
    With frm(FT)
        .FormType = FT: .ShowInputForm OwnerForm, IL(173), IL(174), _
          LCase(IL(99)), False, 50, txtLeft
    End With
End Sub
'// Show the pure text form
Public Sub ShowTextForm(OwnerForm As Object, sText$, sCaption$)
Dim FT As mx3FormTypes: FT = frmTextForm
    With frm(FT)
        .FormType = FT: .WindowState = 0: .ShowTextForm OwnerForm, sText$, sCaption$
    End With
End Sub
'// Show MP3 tag batch edit form
Public Sub ShowTagBatchEdit(Optional OwnerForm As Object)
Dim FT As mx3FormTypes: FT = frmTagBatchEdit
    With frm(FT)
        .FormType = FT: .WindowState = 0: .ShowTagBatchEdit OwnerForm
    End With
End Sub
'// Show playlist creator form
Public Sub ShowPlaylistCreator(Optional OwnerForm As Object)
Dim FT As mx3FormTypes: FT = frmPlayListCreator
    With frm(FT)
        .FormType = FT: .WindowState = 0: .ShowPlaylistCreator OwnerForm
    End With
End Sub
'// Show the user change CD drive form
Public Sub ShowSelectCD()
Dim FT As mx3FormTypes: FT = frmSelectCD
    With frm(FT)
        .FormType = FT: .WindowState = 0
        CDDriveLetter = .ShowDrive(CDDriveLetter, CDGetDriveLetters$)
    End With
End Sub
'// Will return a string with all CD drive letters i.e "D:|F:", first found to RetFirstDrive$
Private Function CDGetDriveLetters$(Optional RetFirstDrive$)
Dim x%, s$, d$
    For x% = 65 To 90
        d$ = Chr(x%) & ":":
        If GetDriveType(d$) = 5 Then
            If LenB(s$) = 0 Then RetFirstDrive$ = d$
            s$ = s$ & d$ & "|"
        End If
    Next
    If LenB(s$) Then CDGetDriveLetters$ = Left$(s$, Len(s$) - 1)
End Function
'// Show form, let user adjust maximum diff in Ms between MP3 tracks when tag & rename
Public Function ShowAdjustDiffMs$(Optional OwnerForm As Object)
Dim s$, FT As mx3FormTypes: FT = frmAdjustMilliSec
    With frm(FT)
        .FormType = FT: .WindowState = 0
        s$ = Val(.ShowAdjustDiffMs(OwnerForm, m_Settings.MP3MaxDiffMs))
        If s$ <> "0" Then m_Settings.MP3MaxDiffMs = Val(s$)
    End With
End Function

'// Show form, let user select if several matches from CDDB
Private Function ShowCDDBMatches$(sInfotext$, ByVal sListItems$)
Dim s$, s2$, Ct As String * 10, r$(), c$(), n&, l&, FT As mx3FormTypes: FT = frmShowCDDBMatches
    r$() = Split(sListItems$, vbNewLine): n& = UBound(r$())
    s$ = "#" & vbTab & "CDID" & vbTab & vbTab & "Category" & vbTab & _
      vbTab & "Artist/Album" & vbNewLine & vbNewLine
    For l& = 0 To n& - 1
        c$() = Split(r$(l&), " "): s2$ = Right$(r$(l&), Len(r$(l&)) - _
          (Len(c$(0)) + Len(c$(1)) + 2)): Ct = c$(1)
        s$ = s$ & ZeroL(l& + 1, 2) & vbTab & Ct & _
          vbTab & UCase(c$(0)) & String(IIf(Len(c$(0)) > 8, 1, 2), vbTab) & s2$ & vbNewLine
    Next
    With frm(FT)
        .FormType = FT: .WindowState = 0
        ShowCDDBMatches$ = .ShowCDDBMatches(sInfotext$, s$): Erase r$(): Erase c$()
    End With
End Function

'// SHOW FORMS END


Public Function MsgBoxW(ByVal Prompt$, Optional ByVal Buttons As VbMsgBoxStyle = vbInformation, Optional ByVal Title$, Optional ByVal hWndOwner&) As VbMsgBoxResult
On Local Error Resume Next 'Show a unicode message box...
    If hWndOwner = 0 Then _
        If Not Screen.ActiveForm Is Nothing Then _
          hWndOwner& = Screen.ActiveForm.hWnd
    If LenB(Title$) = 0 Then Title$ = c_Title
    If Buttons < vbSystemModal Then Buttons = Buttons + vbSystemModal
    MsgBoxW = MessageBoxW(hWndOwner&, StrPtr(Prompt$), StrPtr(Title$), Buttons)
End Function

'// Get WAV header info, need to call WAVHeaderRead first to call this function
Public Function WAVHeaderInfo$(Info As WAVHeaderInfo)
    With WAVHeader
        Select Case Info
            Case wavChannels:   WAVHeaderInfo$ = .channels
            Case wavFrequency:  WAVHeaderInfo$ = .Frequency
            Case wavBits:       WAVHeaderInfo$ = .bits
            Case wavKbps:       WAVHeaderInfo$ = .Kbps
            Case wavFilesize:   WAVHeaderInfo$ = .FileSize
            Case wavPlaytime:   WAVHeaderInfo$ = .Playtime
        End Select
    End With
End Function

'// Read WAV header info
Public Function WAVHeaderRead(Optional FileWAV$, Optional RetInfo$, Optional BaseIn% = 10, Optional BaseOut% = 16) As Boolean
Dim riff As String * 4, l&, b As Boolean, v$(), sChr$, iAsc%, pv%, dTot#, d#, OutB#, _
  s$, sFreq$, bFreq As Byte: On Error GoTo WAVHeaderReadErr
    
    If LenB(FileWAV$) = 0 Then
        b = True: FileWAV$ = FileFolder(encShowOpen, "WAVE " & IL(38) & " (*.wav)|*.wav", IL(156), , encExplorer + encHideReadOnly)
    End If
    If Not FileExists(FileWAV$) Then Exit Function
    With WAVHeader
        l& = FreeFile: Open FileWAV$ For Binary As #l&
        Get #l&, 1, riff: Get #l&, 23, .channels: Get #l&, 35, .bits
        Get #l&, 25, bFreq: Close #l&
        If riff <> "RIFF" Then Exit Function Else sFreq$ = Val(Str(bFreq))
        If sFreq$ = "" Or BaseIn% < 2 Or BaseIn% > 36 Or BaseOut% < 1 Or BaseOut% > 36 Then _
          Exit Function Else s$ = UCase$(sFreq$): pv% = Len(s$): sFreq$ = ""
        For l& = 1 To Len(s$)
            pv% = pv% - 1: sChr$ = Mid$(s$, l&, 1): iAsc% = 0
            If Asc(sChr$) > 64 And Asc(sChr$) < 91 Then iAsc% = Asc(sChr$) - 55
            If iAsc% = 0 Then
                If Asc(sChr$) < 48 Or Asc(sChr$) > 57 Then Exit Function Else iAsc% = Val(sChr$)
            End If
            If iAsc% < 0 Or iAsc% > BaseIn% - 1 Then Exit Function
            dTot# = dTot# + iAsc% * (BaseIn% ^ pv%)
        Next
        Do
            OutB# = CDbl(BaseOut%): d# = (dTot# - (Int(dTot# / OutB#) * OutB#))
            dTot# = (dTot# - d#) / BaseOut%
            If d# >= 10 Then
                sChr$ = Chr$(d# + 55)
            Else
                sChr$ = Right$(Str$(d#), Len(Str$(d#)) - 1)
            End If
            sFreq$ = sChr$ & sFreq$
        Loop While dTot# > 0
        Select Case sFreq$
            Case 40: l& = 8000
            Case 11: l& = 11025
            Case 22: l& = 22050
            Case 0:  l& = 32000
            Case 44: l& = 44100
            Case 80: l& = 48000
        End Select
        .Frequency = l&: d# = FileLen(FileWAV$): .FileSize = d#
        .Kbps = (.bits * .channels * .Frequency)
        d# = (((d# * 8) - 8000) / .Kbps): l& = d# \ 60: d# = d# - (l& * 60)
        .Playtime = Format(l&, "00") & ":" & Format(d#, "00"): s$ = vbTab
         .Kbps = .Kbps / 1000
        v$() = Split("Channels  ;Frequency;Bits            ;Bitrate       ;Play time   ;Filesize      ", ";")
          RetInfo$ = "Wave-File header:" & vbLf & vbLf & v$(0) & s$ & .channels & vbLf & v$(1) & s$ & .Frequency / 1000 & " kHz" & vbLf & _
          v$(2) & s$ & .bits & "" & vbLf & v$(3) & s$ & .Kbps & " Kbps" & vbLf & _
          v$(4) & s$ & .Playtime & vbLf & v$(5) & s$ & FormatKMG(.FileSize, True, 2): Erase v$()
        If b Then MsgBoxW RetInfo$ & vbLf & vbLf & FileWAV$, 64
    End With
    WAVHeaderRead = True: Exit Function
WAVHeaderReadErr:
    Log "WAVHeaderRead(): " & Err.Description, encLogErrorNoMsgBox: Err.Clear
End Function

'// Get the length of a MP3 file in milliseconds from MCI
Function MP3LenMs&(sFile$)
  SendString$ "open " & GetShortPath(sFile$) & " alias x": SendString$ "set x time format 0"
  MP3LenMs& = Val(SendString("status x length")): SendString$ "close x": DoEvents
End Function

'// MCI string handeling
Private Function SendString(StrMCI$)
Const SL& = 128: Dim l&, s As String * SL: l& = mciSendString(StrMCI$, s, SL, 0&)
    If l& = 0 Then
        SendString = Left$(s, InStr(s, vbNullChar) - 1)
    Else: SendString = ErrMCI$(l&): Log StrMCI$ & ": " & SendString: End If
End Function

'// Add n zero to the left of the input string
Private Function ZeroL$(ByVal s$, ByVal n%)
    If Len(s$) < n% Then ZeroL = String$(n% - Len(s$), "0") & s$ Else ZeroL = s$
End Function


'// Replace any bad character in a file name
Function FileReplBad$(ByVal Name$, Optional ReplaceChar$ = " ")
Dim x%: Const BAD = "<>*?|/\:$"""""
    For x% = 1 To Len(BAD)
        If InStr(Name$, Mid$(BAD, x%, 1)) > 0 Then
            Name$ = Replace(Name$, Mid$(BAD, x%, 1), ReplaceChar$)
        End If
    Next
    FileReplBad$ = Name$
End Function

'// Extract info from a file path
Public Function PathInfo$(ByVal Path$, Optional ByVal Info As encPathInfo = encNameExt, Optional NewExtention$)
Dim s$(): s$() = Split(Path$, "\"): On Local Error Resume Next
    Select Case Info
        Case encExt, encNewExt 'Extract the extension of the file or replace it.
            s$() = Split(s$(UBound(s$())), "."): s$(0) = UCase(s$(1))
            If Info = encNewExt Then s$(0) = Left$(Path$, Len(Path$) - Len(s$(0))) & NewExtention$
        Case encNameExt 'Extract the file name and Extension
            s$(0) = s$(UBound(s$()))
        Case encName 'Extract the filname
            s$() = Split(s$(UBound(s$())), ".")
        Case encPath, encPathBS 'Extract only Path, with or without ending backslash \
            s$(0) = Mid$(Path$, 1, InStrRev(Path$, "\") - IIf(Info = encPathBS, 0, 1))
        Case encDrive 'Extract only Path
            s$(0) = UCase(Left$(Path$, 2))
    End Select
    PathInfo$ = s$(0): Erase s$()
End Function

Public Sub SortArray(ByRef vArr As Variant, Optional ByVal lLeft&, Optional ByVal lRight&)
Dim x&, y&, vM, vS 'Sort a array
   
    If lRight = 0 Then: lLeft = LBound(vArr): lRight = UBound(vArr)
    x& = lLeft: y& = lRight
    vM = vArr((lLeft + lRight) \ 2)
    Do
        Do While vArr(x&) < vM And x& < lRight
            x& = x& + 1
        Loop
        Do While vM < vArr(y&) And y& > lLeft
            y& = y& - 1
        Loop
        If x& <= y& Then
            vS = vArr(x&): vArr(x&) = vArr(y&): vArr(y&) = vS
            x& = x& + 1: y& = y& - 1
        End If
    Loop Until x& > y&
    If lLeft < y& Then SortArray vArr, lLeft, y&
    If x& < lRight Then SortArray vArr, x&, lRight
End Sub
Public Sub SelectFilesToEncode()
Dim s$: s$ = ShowOpen: If LenB(s$) Then EncFileIn = s$
End Sub
Public Sub SelectEncodeOutputFolder(Optional ResetToDefault As Boolean)
Dim s$: If ResetToDefault Then EncFolderOut = "": Exit Sub
    s$ = ShowFolder(IL(166)): If LenB(s$) Then EncFolderOut = s$
End Sub



'// Return the built in text strings from this class (Use enum or index).
Public Function GetText$(ByVal Text As encBuiltInText, Optional EnglishLanguage As Boolean)
    If Val(Text) < 0 Or Val(Text) > UBound(IL()) Then Exit Function
    GetText$ = IIf(EnglishLanguage, ILOrg(Text), IL(Text))
End Function

''// Open a file in Notepad
'Sub FileOpenInNotepad(sFile$)
'Dim Shell As Object 'Open a file in notepad
'On Local Error Resume Next
'    Set Shell = CreateObject("Wscript.Shell")
'    Shell.Run "%windir%\notepad " & sFile$: Set Shell = Nothing
'    If Err <> 0 Then Err.Clear
'End Sub

'// Merge MP3 files
Function MP3Merge(Optional MainFile$, Optional FileToAdd$) As Boolean
On Error GoTo ErrMP3Merge
Dim Sz!, s$(), l&, f$, n&: f$ = "MP3-" & IL(38) & " (*.mp3)|*.mp3": Const Chunk = 2 ^ 20 '1 MB
    If LenB(MainFile$) = 0 Then
        MainFile$ = FileFolder(encShowOpen, f$, IL(59), , encExplorer)
        If LenB(MainFile$) = 0 Then Exit Function
    End If
    If LenB(FileToAdd$) = 0 Then
        FileToAdd$ = FileFolder(encShowOpen, f$, IL(159) & _
          " '" & PathInfo(MainFile$, encNameExt) & "'", , encExplorer)
        If LenB(FileToAdd$) = 0 Then Exit Function
    End If
    Sz! = FileLen(FileToAdd$)
    If Sz! < Chunk Then  'Less than 1MB
        ReDim s(1): s(1) = GetFromFile(FileToAdd$, 1, Sz!)
    Else
        n& = Sz! \ Chunk: ReDim s(1 To n&)
        For l& = 1 To n&
            s(l&) = GetFromFile(FileToAdd$, (l& - 1) * Chunk + 1, Chunk): DoEvents
        Next
        If (n& * Chunk) < Sz! Then
            Sz! = Sz! - (n& * Chunk): ReDim Preserve s$(1 To n& + 1)
            s(n& + 1) = GetFromFile(FileToAdd$, n& * Chunk, Sz!)
        End If
    End If
    Sz! = FileLen(MainFile$) + 1: n& = UBound(s$)
    For l& = 1 To n&
        PutToFile MainFile$, Sz! + (l& - 1) * Chunk, s$(l&): DoEvents
    Next
    MP3Merge = True
    Exit Function
ErrMP3Merge:
    Log Err.Description & vbLf & MainFile$ & vbLf & _
      FileToAdd$, encLogError: Exit Function
    Resume
End Function
Private Function GetFromFile$(Path$, nRecord!, nByte!)
Dim f%, s$: f% = FreeFile: s$ = Space$(nByte!)
    Open Path$ For Binary Access Read As #f%
    Get #f%, nRecord!, s$: Close #f%: GetFromFile = s$
End Function

Private Sub PutToFile(Path$, nRecord!, Data$)
Dim f%: f% = FreeFile
    Open Path$ For Binary Access Write As #f%
    Put #f%, nRecord!, Data: Close #f%
End Sub



'// Propertys ************************************************************************

Public Property Get Icon() As Object
    Set Icon = frmMX3.Icon
End Property
''// Can be used when translate the interface language using GAT
Public Property Get InterfaceLanguage(Optional TheTranslated As Boolean) As String()
Dim s$
    If TheTranslated Then
        s$ = Join(m_GenreID3$(), vbNewLine) & vbNewLine & Join(IL(), vbNewLine)
    Else
        s$ = Join(m_GenreID3Org$(), vbNewLine) & vbNewLine & Join(ILOrg(), vbNewLine)
    End If
    InterfaceLanguage = Split(s$, vbNewLine)
End Property
Public Property Let InterfaceLanguage(TheTranslated As Boolean, Strings() As String)
Dim l&: If UBound(Strings()) <> 512 Then _
  MsgBoxW "InterfaceLanguage(): " & IL(157), 16: Exit Property
    For l& = 0 To 255: m_GenreID3$(l&) = Strings(l&): Next
    For l& = 0 To c_ResStrUbound: IL(l&) = Strings(256 + l&): Next
End Property

'// To store default language code if using my compiled MX3.
Public Property Get LanguageCode$()
    LanguageCode = m_Settings.LanguageCode
End Property
Public Property Let LanguageCode(Code$)
    m_Settings.LanguageCode = Code$
    m_PathLanguageFile = PathApp & "MX3-" & Code$ & ".lng"
End Property

'// Get path to the language file
Public Property Get PathLanguageFile() As String
    PathLanguageFile = m_PathLanguageFile
End Property

'// Get applications path
Public Property Get PathApp() As String
    PathApp = m_Settings.PathApp
End Property

'// Return if waiting when encoding, the event IsWorking is raised if true
Public Property Get IsWorking() As Boolean
    IsWorking = m_Settings.IsWorking
End Property

'// Call this sub if you like to cancel waiting during encoding
Public Sub StopWork()
    m_Settings.StopWork = True
End Sub

'// File to encode, if this is a null string, a dialog will be displayed to select file
Public Property Get EncFileIn() As String
    EncFileIn = m_Settings.EncFileIn
End Property
Public Property Let EncFileIn(ByVal SourceFile As String)
Dim t$(), l&: On Error GoTo EncFileInErr
    If LenB(SourceFile) Then
        If InStr(1, SourceFile, vbNullChar) Then
            t$() = Split(SourceFile, vbNullChar)
            If Not FolderExists(t$(0)) Then GoTo EncFileInErr
            For l& = 1 To UBound(t$())
                If Not FileExists(t$(0) & "\" & t$(l&)) Then _
                  GoTo EncFileInErr Else WAVFile = t$(0) & "\" & t$(l&)
            Next
            m_EncFileInCount& = l& - 1: Erase t$()
        Else: m_EncFileInCount& = 1: WAVFile = SourceFile: End If
    Else: m_EncFileInCount& = 0: End If
    m_Settings.EncFileIn = SourceFile
    Exit Property
EncFileInErr:
    MsgBoxW IL(157) & " EncFileIn():" & vbLf & SourceFile & vbLf & vbLf & _
      IL(158) & ": 'Folder & vbNullChar & FileName1 & vbNullChar & FileName2'...", 16
End Property
Public Property Get EncFileInCount() As Long
    EncFileInCount = m_EncFileInCount&
End Property

'// How to show the Blade proccecing window
Public Property Get EncDisplayBlade() As VbAppWinStyle
    EncDisplayBlade = m_Settings.EncDisplayBlade
End Property
Public Property Let EncDisplayBlade(ShowAs As VbAppWinStyle)
    m_Settings.EncDisplayBlade = ShowAs
End Property

'// If true no messagebox before downloading Blade
Public Property Get QuietDownload() As Boolean
    QuietDownload = m_Settings.QuietDownload
End Property
Public Property Let QuietDownload(Value As Boolean)
    m_Settings.QuietDownload = Value
End Property

'// If true this class will wait until the job is done and raise the event WhileEncoding
Public Property Get EncWaitUntilDone() As Boolean
    EncWaitUntilDone = m_Settings.EncWaitUntilDone
End Property
Public Property Let EncWaitUntilDone(Value As Boolean)
    m_Settings.EncWaitUntilDone = Value
End Property

'// If true and EncWaitUntilDone = True, display a message box of the result.
Public Property Get EncMsgboxResult() As Boolean
    EncMsgboxResult = m_Settings.EncMsgboxResult
End Property
Public Property Let EncMsgboxResult(Value As Boolean)
    m_Settings.EncMsgboxResult = Value
End Property

'// If destination MP3-File exist actions Create Backup, Delete or Stop
Public Property Get IfDestExist() As encIfDestExist
    IfDestExist = m_Settings.IfDestExist
End Property
Public Property Let IfDestExist(Action As encIfDestExist)
    m_Settings.IfDestExist = Action
End Property
'// Will sort all selected files nice if join
Public Property Get EncFilesSorted() As Boolean
    EncFilesSorted = m_Settings.EncFilesSorted
End Property
Public Property Let EncFilesSorted(Value As Boolean)
    m_Settings.EncFilesSorted = Value
End Property
'// Will join all selected files into one MP3
Public Property Get EncFilesJoin() As Boolean
    EncFilesJoin = m_Settings.EncFilesJoin
End Property
Public Property Let EncFilesJoin(Value As Boolean)
    m_Settings.EncFilesJoin = Value
End Property
'// Make sure there's no gap between files when join into one MP3
Public Property Get EncFilesNoGapIfJoin() As Boolean
    EncFilesNoGapIfJoin = m_Settings.EncFilesNoGapIfJoin
End Property
Public Property Let EncFilesNoGapIfJoin(Value As Boolean)
    m_Settings.EncFilesNoGapIfJoin = Value
End Property

Public Property Get AutoAddTag() As Boolean
    AutoAddTag = m_Settings.AutoAddTag
End Property
Public Property Let AutoAddTag(Value As Boolean)
    m_Settings.AutoAddTag = Value
End Property

Public Property Get hWndBladeWindow() As Long
    hWndBladeWindow = m_Settings.hWndBladeWindow
End Property

Public Property Get MP3TagExist() As Boolean
    MP3TagExist = m_TagExist
End Property
'

'// Save settings to disk file  when this class terminates
Public Property Get SettingsSaveOnExit() As Boolean
    SettingsSaveOnExit = m_Settings.SettingsSaveOnExit
End Property
Public Property Let SettingsSaveOnExit(Value As Boolean)
    m_Settings.SettingsSaveOnExit = Value
End Property
'// Save logfile when this class terminates
Public Property Get LogSave() As Boolean
    LogSave = m_Settings.LogSave
End Property
Public Property Let LogSave(Value As Boolean)
    m_Settings.LogSave = Value
End Property
'// Time in milliseconds to wait before checking if file size on the file Blade creates have change.
Public Property Get EncWaitIdleMsCheck() As Long
    EncWaitIdleMsCheck = m_Settings.EncWaitIdleMsCheck
End Property
Public Property Let EncWaitIdleMsCheck(Ms As Long)
    m_Settings.EncWaitIdleMsCheck = Ms
End Property

'// The format the MP3 filenames will get when using data from FreeCDDB.org
Public Property Get MP3FileFormat() As encMP3FileFormats
    MP3FileFormat = m_Settings.MP3FileFormat
End Property
Public Property Let MP3FileFormat(Format As encMP3FileFormats)
    m_Settings.MP3FileFormat = Format
End Property

'// The maximum diff in milliseconds when comparing MP3-File with audio data from FreeCDDB.org
Public Property Get MP3MaxDiffMs%()
    MP3MaxDiffMs = m_Settings.MP3MaxDiffMs
End Property
Public Property Let MP3MaxDiffMs(Ms%)
    m_Settings.MP3MaxDiffMs = Ms%
End Property

'// The selected CD drive letter.
Public Property Get CDDriveLetter$()
    CDDriveLetter = m_Settings.CDDriveLetter
End Property
Public Property Let CDDriveLetter(Drive$)
    If LenB(Drive$) Then Drive$ = UCase(Left$(Drive$, 1)) Else Exit Property
    m_Settings.CDDriveLetter = Drive$
End Property

'// Return valid MP3 file name formats as a string array
Public Property Get MP3FileNameFormats() As String()
Dim x%, v$(8): For x% = 0 To 8: v$(x%) = IL(x% + 60): Next: MP3FileNameFormats = v$: Erase v$()
End Property


'// To store last MP3-File used, this property will automaticly change to last file used by M3X
Public Property Get MP3File$()
    MP3File = m_Settings.MP3File
End Property
Public Property Let MP3File(MP3File$)
    m_Settings.MP3File = MP3File$
End Property

'// To store last WAV-File used, this property will automaticly change to last file used by M3X
Public Property Get WAVFile$()
    WAVFile = m_Settings.WAVFile
End Property
Public Property Let WAVFile(WAVFile$)
    m_Settings.WAVFile = WAVFile$: frm(frmCalcForm).CalcFilePath = WAVFile
End Property


'// ***********************************************************************************************
'// BLADE FLAGS from here, the info text copied from: http://home.swipnet.se/~w-82625/default.htm
'// ***********************************************************************************************


'// Defines the bitrate for the MP3-file. Higher bitrates gives better quality,-
'// but also bigger files. Most people prefer to generate 128 kBit MP3s.
'// Please note that it's the total bitrate that is specified, so if you're-
'// generating a stereo MP3-file at 128 kBit you get 64 kBit for left channel-
'// and 64 kBit for the right channel. The default setting is 128 kBit for stereo-
'// files and 64 kBit for EncMono files.
'// Allowed bitrates are: 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256 and 320.

Public Property Get EncBitrate() As encEncodeBitRates
    EncBitrate = m_Settings.EncBitrate
End Property
Public Property Let EncBitrate(Kbps As encEncodeBitRates)
    m_Settings.EncBitrate = Kbps
    frm(frmCalcForm).CalcBitrate = Kbps
End Property
'// This is my and return a string array with valid bitrates, Index 0 = " Kbps", 1 = "320"...
Public Property Get EncBitrates() As String()
    EncBitrates = KbpsChk$()
End Property

'// Adds checksum data to each frame in the MP3 file. The checksum data is useful-
'// for error-correction when streaming the MP3 in realtime over internet-
'// (as done by internet radio stations). It lowers the quality of the sound slightly-
'// since the checksum data also needs to fit in the specified bitrate and is not needed-
'// for normal use.
Public Property Get EncAddChecksum() As Boolean
    EncAddChecksum = m_Settings.EncAddChecksum
End Property
Public Property Let EncAddChecksum(Value As Boolean)
    m_Settings.EncAddChecksum = Value
End Property

'// The WAV-file is automatically deleted after having been encoded. Be careful with-
'// this switch because the WAV-file is deleted even if the encoding process failed-
'//(for example due to insufficient drive space).
Public Property Get EncDeleteSourceFiles() As Boolean
    EncDeleteSourceFiles = m_Settings.EncDeleteSourceFiles
End Property
Public Property Let EncDeleteSourceFiles(Value As Boolean)
    m_Settings.EncDeleteSourceFiles = Value
End Property

'// Downmixes stereo input files to EncMono before they are compressed,-
'// thus generating EncMono MP3 files. This switch does nothing if the input file already is EncMono.
Public Property Get EncMono() As Boolean
    EncMono = m_Settings.EncMono
End Property
Public Property Let EncMono(Value As Boolean)
Static Prev As Boolean
    m_Settings.EncMono = Value 'Set default bitrate if change between EncMono/Stereo
    If Prev <> Value Then m_Settings.EncBitrate = IIf(Value, [ 64 Kbps], [128 Kbps])
    Prev = Value
End Property

'// If left, takes only the left channel of the input file and encodes it into a EncMono MP3.
'// If right, takes only the right channel of the input file and encodes it into a EncMono MP3.
'// If the input file is EncMono it will be compressed as it is.
Public Property Get EncMonoChannels() As EncEncMonoChannels
    EncMonoChannels = m_Settings.EncMonoChannels
End Property
Public Property Let EncMonoChannels(channels As EncEncMonoChannels)
    m_Settings.EncMonoChannels = channels
End Property

'// Swaps the left and right channels of stereo input files before encoding them.
Public Property Get EncSwapChannels() As Boolean
    EncSwapChannels = m_Settings.EncSwapChannels
End Property
Public Property Let EncSwapChannels(Value As Boolean)
    m_Settings.EncSwapChannels = Value
End Property

'// Sets the private-flag in the MP3-file, specifying that this is a private MP3.
'// As far as I know there is no program that treats private MP3's differently-
'// from normal ones so it is practically useless. Setting this flag doesn't-
'// affect the encoding time, file size or quality in any way, so set it if it makes you happy.
Public Property Get EncIsPrivate() As Boolean
    EncIsPrivate = m_Settings.EncIsPrivate
End Property
Public Property Let EncIsPrivate(Value As Boolean)
    m_Settings.EncIsPrivate = Value
End Property

'// This switch clears the MP3 file's original-flag that is set by default-
'// from version 0.50 of BladeEnc. Clearing the original-flag doesn't affect the-
'// encoding time, filesize or quality in any way, so do it if it makes you happy.
Public Property Get EncIsCopyrighted() As Boolean
    EncIsCopyrighted = m_Settings.EncIsCopyrighted
End Property
Public Property Let EncIsCopyrighted(Value As Boolean)
    m_Settings.EncIsCopyrighted = Value
End Property

'// This switch clears the MP3 file's original-flag that is set by default from-
'// version 0.50 of BladeEnc. Clearing the original-flag doesn't affect the encoding-
'// time, filesize or quality in any way, so do it if it makes you happy.
Public Property Get EncClearOriginalFlag() As Boolean
    EncClearOriginalFlag = m_Settings.EncClearOriginalFlag
End Property
Public Property Let EncClearOriginalFlag(Value As Boolean)
    m_Settings.EncClearOriginalFlag = Value
End Property

'// Prevents BladeEnc from doing any screen output, except for error messages.
'// This is mostly useful if you want to run bladeenc as a background process-
'// on UNIX systems. The screen output takes nearly no time at all, so you won't-
'// even save 1% by disabling screen output. This switch is global only.
Public Property Get EncNoScreenOutput() As Boolean
    EncNoScreenOutput = m_Settings.EncNoScreenOutput
End Property
Public Property Let EncNoScreenOutput(Value As Boolean)
    m_Settings.EncNoScreenOutput = Value
End Property

'// Makes BladeEnc quit automatically when all files have been encoded.
'// Normally BladeEnc waits for someone to press RETURN before quitting.
'// This switch is global only.
Public Property Get EncCloseWhenDone() As Boolean
    EncCloseWhenDone = m_Settings.EncCloseWhenDone
End Property
Public Property Let EncCloseWhenDone(Value As Boolean)
    m_Settings.EncCloseWhenDone = Value
End Property

'// Specifies an output path for the encoded files. Normally the MP3 files-
'// ends up in the same directory as their corresponding WAV-files resides in,-
'// but using this switch you can get them to end up wherever you like.-
'// This switch is global only.
Public Property Get EncFolderOut() As String
    EncFolderOut = m_Settings.EncFolderOut
End Property
Public Property Let EncFolderOut(FolderPath As String)
    m_Settings.EncFolderOut = FolderPath
End Property

'// Changes the task EncPriority of BladeEnc. Valid settings are-
'// HIGHEST, HIGHER, NORMAL, LOWER, LOWEST and IDLE.
'// BladeEnc is by default set to LOWEST EncPriority which basically means that-
'// BladeEnc functions fine in the background without disturbing or slowing down-
'// any other program. This switch is global only. This setting is not available in all ports.
Public Property Get EncPriority() As encEncPrioritys
    EncPriority = m_Settings.EncPriority
End Property
Public Property Let EncPriority(Prio As encEncPrioritys)
    m_Settings.EncPriority = Prio
End Property

'// Stdin can be specified as input by simply substituting the name of the input-
'// file with stdin. This will make bladeenc read the input from a redirection or pipe instead.
'// Stdout can be specified as output by simply substituting the name of the output-
'// file with stdout. This will make bladeenc send the generated MP3 file through-
'// the standard output stream.

'Public Property Get StdInOut() As encStdInOut
'    StdInOut = m_Settings.StdInOut
'End Property
'Public Property Let StdInOut(InOut As encStdInOut)
'    m_Settings.StdInOut = InOut
'End Property
