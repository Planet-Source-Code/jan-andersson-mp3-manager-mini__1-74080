Attribute VB_Name = "modMX3"

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
Global Const c_ResStrUbound = 176, c_FrmUbound = 13 'Number of frmMX3
Enum mx3FormTypes
    frmAdjustMilliSec
    frmTaggerForm
    frmCalcForm
    frmSelectCD
    frmSearchCDDB_Free
    frmSearchCDDB_CDID
    frmTagRename
    frmShowCDDBMatches
    frmTextForm
    frmPlayListCreator
    frmTagBatchEdit
    frmInput
    frmGenreFromNumber
    frmGenreNumberFromName
End Enum
Private Type mx3TagSearch
    CheckTag As Boolean
    ExcactText As Boolean
    Check(6) As String
End Type
Type TRACK_DATA
    Reserved As Byte
    Adr As Byte
    TrackNumber As Byte
    Reserved1 As Byte
    Address(3) As Byte
End Type
Type CDROM_TOC
    Length(1) As Byte
    FirstTrack As Byte
    LastTrack As Byte
    TrackData(99) As TRACK_DATA
End Type
Enum encResponseCDDB
    encInternetNotConected = -2
    encCanceledByUser = -1
    encFoundExactMatch = 200
    encFoundInExactMatch = 211
    encNoMatchFound = 202
    encNoEntryInDatabase = 401
    encDatabaseEntryCorrupt = 403
    encNoHandshake = 409
    encErrorQuery = 500
End Enum
'// Track info
Type encInfoTrackCDDB
    Title As String
    TitleX As String
    LengthMs As Long
    OffSet As Long
    TrackNoTmp As Integer
    LengthT As String * 5
    StartT As String * 5
    StopT As String * 5
    ListTime As String
    ListName As String
    PathWAV As String
    PathMP3 As String
    TagOK As Boolean
End Type
Type encInfoCDDB
    FoundCDDB As Boolean
    ResponseCDDB As encResponseCDDB
    TocID As String
    id As String * 8
    nTracks As Integer
    LengthT As String * 5
    LengthMs As Double
    RetInfo As String
    Artist As String
    Album As String
    TitleX As String
    Year As String * 4
    Category As String
    MsgReturn As String
    Genre As String
    TOC As CDROM_TOC
    tr() As encInfoTrackCDDB
End Type
'// This type hold my property and other settings
Type encSettings
    BOM As String * 2
    EncFileIn As String
    EncFolderOut As String
    EncBitrate As Integer ' encEncodeBitRates
    EncIsPrivate As Boolean
    EncClearOriginalFlag As Boolean
    EncMono As Boolean
    EncMonoChannels As Integer ' encEncMonoChannels
    EncAddChecksum As Boolean
    EncSwapChannels As Boolean
    EncDeleteSourceFiles As Boolean
    EncPriority As Integer ' encEncPrioritys
    EncIsCopyrighted As Boolean
    EncDisplayBlade As Integer ' VbAppWinStyle
    EncWaitUntilDone As Boolean
    EncNoScreenOutput As Boolean
    EncCloseWhenDone As Boolean
    EncFilesSorted As Boolean
    EncFilesNoGapIfJoin As Boolean
    EncFilesJoin As Boolean
    EncWaitIdleMsCheck As Long
    EncMsgboxResult As Boolean
    IfDestExist As Integer ' encIfDestExist
    IsWorking As Boolean
    StopWork As Boolean
    PathBlade As String
    LanguageCode As String
    MP3File As String
    WAVFile As String
    QuietDownload As Boolean
    PathApp As String
    AutoAddTag As Boolean
    AscyncRIP As Boolean
    AscyncPercent As Single
    hWndBladeWindow As Long
    SettingsSaveOnExit As Boolean
    LogSave As Boolean
    MP3FileFormat As Integer ' encMP3FileFormats
    MP3MaxDiffMs As Integer
    CDDriveLetter As String * 1
    DefComment As String * 30
    CreatePlaylist As Boolean
    FrmTop(c_FrmUbound) As Long
    FrmLeft(c_FrmUbound) As Long
End Type

Global frm(c_FrmUbound) As New frmMX3, nFiles&, M3 As clsMX3, IL$(255), ILOrg$(), sTmpArr$(), sTmpStr$, _
   m_Settings As encSettings, sTmpList$(), sTmpStrTag$, sTmpFolder$, InfoCDDB As encInfoCDDB, _
   SearchTag As mx3TagSearch, lTrackTimeMs&(), m_GenreID3Org$(), m_GenreID3$(255) ', _
   MnuIL$(255), MnuILOrg$()



