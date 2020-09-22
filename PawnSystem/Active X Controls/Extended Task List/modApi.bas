Attribute VB_Name = "modApi"
Option Explicit

'*************************************************************

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const TH32CS_SNAPMODULE = &H8&

Public Const hNull = 0
Public Const WIN95_System_Found = 1
Public Const WINNT_System_Found = 2
Public Const Default_Log_Size = 10000000
Public Const Default_Log_Days = 0
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
Public Const STANDARD_RIGHTS_ALL = &H1F0000
' ===== From Win32 Ver.h =================
' ----- VS_VERSION.dwFileFlags -----
Public Const VS_FFI_SIGNATURE = &HFEEF04BD
Public Const VS_FFI_STRUCVERSION = &H10000
Public Const VS_FFI_FILEFLAGSMASK = &H3F&

' ----- VS_VERSION.dwFileFlags -----
Public Const VS_FF_DEBUG = &H1
Public Const VS_FF_PRERELEASE = &H2
Public Const VS_FF_PATCHED = &H4
Public Const VS_FF_PRIVATEBUILD = &H8
Public Const VS_FF_INFOINFERRED = &H10
Public Const VS_FF_SPECIALBUILD = &H20

' ----- VS_VERSION.dwFileOS -----
Public Const VOS_UNKNOWN = &H0
Public Const VOS_DOS = &H10000
Public Const VOS_OS216 = &H20000
Public Const VOS_OS232 = &H30000
Public Const VOS_NT = &H40000

Public Const VOS__BASE = &H0
Public Const VOS__WINDOWS16 = &H1
Public Const VOS__PM16 = &H2
Public Const VOS__PM32 = &H3
Public Const VOS__WINDOWS32 = &H4

Public Const VOS_DOS_WINDOWS16 = &H10001
Public Const VOS_DOS_WINDOWS32 = &H10004
Public Const VOS_OS216_PM16 = &H20002
Public Const VOS_OS232_PM32 = &H30003
Public Const VOS_NT_WINDOWS32 = &H40004

' ----- VS_VERSION.dwFileType -----
Public Const VFT_UNKNOWN = &H0
Public Const VFT_APP = &H1
Public Const VFT_DLL = &H2
Public Const VFT_DRV = &H3
Public Const VFT_FONT = &H4
Public Const VFT_VXD = &H5
Public Const VFT_STATIC_LIB = &H7

' ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Public Const VFT2_UNKNOWN = &H0
Public Const VFT2_DRV_PRINTER = &H1
Public Const VFT2_DRV_KEYBOARD = &H2
Public Const VFT2_DRV_LANGUAGE = &H3
Public Const VFT2_DRV_DISPLAY = &H4
Public Const VFT2_DRV_MOUSE = &H5
Public Const VFT2_DRV_NETWORK = &H6
Public Const VFT2_DRV_SYSTEM = &H7
Public Const VFT2_DRV_INSTALLABLE = &H8
Public Const VFT2_DRV_SOUND = &H9
Public Const VFT2_DRV_COMM = &HA



'*************************************************************
Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type


Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type


Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long ' This process's parent process
    pcPriClassBase As Long ' Base priority of process threads
    dwFlags As Long
    szExeFile As String * 260 ' MAX_PATH
End Type

Public Type MODULEENTRY32
  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
  szModule As String * 256
  szExePath As String * 260
End Type


Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long '1 = Windows 95.
    '2 = Windows NT
    szCSDVersion As String * 128
End Type



Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

'*************************************************************
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
            "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle _
                As Long, ByVal dwlen As Long, lpData As Any) As Long

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" _
    Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
                                        lpdwHandle As Long) As Long

Public Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As _
      Long, lppe As Any) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As _
      Long, lppe As Any) As Long

Public Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" _
                                                                    (pBlock As Any, _
                                                                    ByVal lpSubBlock As String, _
                                                                    lplpBuffer As Any, _
                                                                    puLen As Long) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
                                                                    (ByVal Path As String, _
                                                                    ByVal cbBytes As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                                                (dest As Any, _
                                                                ByVal Source As Long, _
                                                                ByVal length As Long)

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
                                                                (ByVal lpString1 As String, _
                                                                ByVal lpString2 As Long) As Long



Public moCollection As Collection

Public Sub PopulateTaskInfo(lvItems As ListView, rsExeName As String)
Dim lclsTask        As clsTask
Dim loitem          As ListItem
Dim loColHeader     As ColumnHeader
Dim lsItem          As String
Dim llCount         As Long

    lvItems.ListItems.Clear
    
    ' set the default columns headers
    If lvItems.ColumnHeaders.Count = 0 Then
        lvItems.ColumnHeaders.Add , , "Filename"
        lvItems.ColumnHeaders.Add , , "Path"
        lvItems.ColumnHeaders.Add , , "FileDate"
        lvItems.ColumnHeaders.Add , , "lFileSize"
        lvItems.ColumnHeaders.Add , , "CompanyName"
        lvItems.ColumnHeaders.Add , , "FileDescription"
        lvItems.ColumnHeaders.Add , , "FileVersion"
        lvItems.ColumnHeaders.Add , , "InternalName"
        lvItems.ColumnHeaders.Add , , "LegalCopyright"
        lvItems.ColumnHeaders.Add , , "OriginalFileName"
        lvItems.ColumnHeaders.Add , , "ProductName"
        lvItems.ColumnHeaders.Add , , "ProductVersion"
        lvItems.ColumnHeaders.Add , , "Comments"
        lvItems.ColumnHeaders.Add , , "LegalTrademarks"
        lvItems.ColumnHeaders.Add , , "PrivateBuild"
        lvItems.ColumnHeaders.Add , , "SpecialBuild"
        lvItems.ColumnHeaders.Add , , "FileType"
        lvItems.ColumnHeaders.Add , , "FileFlags"
        lvItems.ColumnHeaders.Add , , "FileOS"
    End If
    
    Set moCollection = GetProcesses(rsExeName)
        
    For Each lclsTask In moCollection
        llCount = 0
        For Each loColHeader In lvItems.ColumnHeaders
        
            ' get the infromation to place in the lv by the column text
            Select Case loColHeader.Text
                Case "Filename"
                    lsItem = lclsTask.Filename
                    
                Case "Path"
                    lsItem = lclsTask.Path
                    
                Case "FileDate"
                    lsItem = lclsTask.FileDate
                    
                Case "lFileSize"
                    lsItem = FormatFileSize(lclsTask.lFileSize)
                    
                Case "CompanyName"
                    lsItem = lclsTask.CompanyName
                    
                Case "FileDescription"
                    lsItem = lclsTask.FileDescription
                    
                Case "FileVersion"
                    lsItem = lclsTask.FileVersion
                    
                Case "InternalName"
                    lsItem = lclsTask.InternalName
                    
                Case "LegalCopyright"
                    lsItem = lclsTask.LegalCopyright
                    
                Case "OriginalFileName"
                    lsItem = lclsTask.OriginalFileName
                    
                Case "ProductName"
                    lsItem = lclsTask.ProductName
                    
                Case "ProductVersion"
                    lsItem = lclsTask.ProductVersion
                    
                Case "Comments"
                    lsItem = lclsTask.Comments
                    
                Case "LegalTrademarks"
                    lsItem = lclsTask.LegalTrademarks
                    
                Case "PrivateBuild"
                    lsItem = lclsTask.PrivateBuild
                    
                Case "SpecialBuild"
                    lsItem = lclsTask.SpecialBuild
                    
                Case "FileType"
                    lsItem = lclsTask.FileType
                    
                Case "FileFlags"
                    lsItem = lclsTask.FileFlags
                    
                Case "FileOS"
                    lsItem = lclsTask.FileOS
            End Select
            
            ' put information in the listview.
            llCount = llCount + 1
            If llCount = 1 Then
                Set loitem = lvItems.ListItems.Add(, , lsItem)
            Else
                loitem.SubItems(llCount - 1) = lsItem
            End If
        Next
    Next
    
End Sub


' Function that formats the file size in bytes into KB/MB/GB
Private Function FormatFileSize(ByVal TheSize_BYTEs As Long) As String
On Error Resume Next
  
  Const KB As Long = 1024
  Const MB As Long = KB * KB
  Dim FormatSoFar As String
  
  ' Return size of file in kilobytes.
  If TheSize_BYTEs = -1 Then
    FormatFileSize = "0.0KB (0 bytes)"
  ElseIf TheSize_BYTEs < KB Then
    FormatSoFar = Format(TheSize_BYTEs, "#,##0") & " bytes"
  Else
    Select Case TheSize_BYTEs \ KB
      Case Is < 10
        FormatSoFar = Format(TheSize_BYTEs / KB, "0.00") & "KB"
      Case Is < 100
        FormatSoFar = Format(TheSize_BYTEs / KB, "0.0") & "KB"
      Case Is < 1000
        FormatSoFar = Format(TheSize_BYTEs / KB, "0.0") & "KB"
      Case Is < 10000
        FormatSoFar = Format(TheSize_BYTEs / MB, "0.00") & "MB"
      Case Is < 100000
        FormatSoFar = Format(TheSize_BYTEs / MB, "0.0") & "MB"
      Case Is < 1000000
        FormatSoFar = Format(TheSize_BYTEs / MB, "0.0") & "MB"
      Case Is < 10000000
        FormatSoFar = Format(TheSize_BYTEs / MB / KB, "0.00") & "GB"
    End Select
    FormatSoFar = FormatSoFar & " (" & Format(TheSize_BYTEs, "#,##0") & " bytes)"
  End If
  
  FormatFileSize = FormatSoFar
  
End Function

Function StrZToStr(s As String) As String
   StrZToStr = Left$(s, Len(s) - 1)
End Function

Public Function GetProcesses(ByVal EXEName As String) As Collection
Dim booResult               As Boolean
Dim lngLength               As Long
Dim lngProcessID            As Long
Dim strProcessName          As String
Dim lngSnapHwnd             As Long
Dim udtProcEntry            As PROCESSENTRY32
Dim lngCBSize               As Long 'Specifies the size, In bytes, of the lpidProcess array
Dim lngCBSizeReturned       As Long 'Receives the number of bytes returned
Dim lngNumElements          As Long
Dim lngProcessIDs()         As Long
Dim lngCBSize2              As Long
Dim lngModules(1 To 200)    As Long
Dim lngReturn               As Long
Dim strModuleName           As String
Dim lngSize                 As Long
Dim lngHwndProcess          As Long
Dim lngLoop                 As Long
Dim b                       As Long
Dim c                       As Long
Dim e                       As Long
Dim d                       As Long
Dim pmc                     As PROCESS_MEMORY_COUNTERS
Dim lRet                    As Long
Dim strProcName2            As String
Dim strProcName             As String
Dim lclsFile                As clsTask
Dim loColRet                As Collection
Dim llLoop                  As Long
Dim llEnd                   As Long
Dim sname                   As String
Dim hSnap                   As Long
Dim proc                    As PROCESSENTRY32
Dim f2                      As Long
Dim hSnapMod                As Long
Dim proc2                   As MODULEENTRY32

'Turn on Error handler
On Error GoTo Error_handler

    Set loColRet = New Collection
    Set GetProcesses = loColRet
    
    EXEName = UCase$(Trim$(EXEName))
    lngLength = Len(EXEName)

    'ProcessInfo.bolRunning = False

    Select Case getOsVersion()
        Case WIN95_System_Found 'Windows 95/98
            hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
            
            If hSnap = hNull Then
                Exit Function
            End If
            
            proc.dwSize = Len(proc)
            
            ' Iterate through the processes
            lRet = Process32First(hSnap, proc)
            
            Do While lRet
                sname = StrZToStr(proc.szExeFile)
                
                ' if the name is the one being sought,
                If InStr(1, UCase$(sname), UCase$(EXEName)) > 0 Then
                
                    ' enum the loaded modulates for the task
                    hSnapMod = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, proc.th32ProcessID)
                    
                    If hSnapMod <> 0 Then
                    
                        proc2.dwSize = Len(proc2)
                        f2 = Module32First(hSnapMod, proc2)
                            
                        Do While f2
                            sname = StrZToStr(proc2.szModule)
                            
                            Set lclsFile = New clsTask
                            lclsFile.Filename = sname
                            lclsFile.Path = StrZToStr(proc2.szExePath) & "\" & sname
                            loColRet.Add lclsFile
                            
                            ' get the extended info
                            GetExtendedInfo lclsFile
                            
                            ' goto next
                            f2 = Module32Next(hSnapMod, proc2)
                            
                        Loop
                    
                        CloseHandle hSnapMod
                    End If
                End If
                
                lRet = Process32Next(hSnap, proc)
            Loop

        Case WINNT_System_Found 'Windows NT

            lngCBSize = 8 ' Really needs To be 16, but Loop will increment prior to calling API
            lngCBSizeReturned = 96

            Do While lngCBSize <= lngCBSizeReturned
                DoEvents
                'Increment Size
                lngCBSize = lngCBSize * 2
                'Allocate Memory for Array
                ReDim lngProcessIDs(lngCBSize / 4) As Long
                'Get Process ID's
                lngReturn = EnumProcesses(lngProcessIDs(1), lngCBSize, lngCBSizeReturned)
            Loop

            'Count number of processes returned
            lngNumElements = lngCBSizeReturned / 4
            'Loop thru each process

            For lngLoop = 1 To lngNumElements
                'Get a handle to the Process and Open
                lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessIDs(lngLoop))
                
                If lngHwndProcess <> 0 Then
                    'Get an array of the module handles for the specified process
                    lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
    
                    'If the Module Array is retrieved, Get the ModuleFileName
                    If lngReturn <> 0 Then
                        llEnd = lngCBSize2 / 4
                        'Buffer with spaces first to allocate memory for byte array
                        strModuleName = Space(MAX_PATH)
                        
                        'Must be set prior to calling API
                        lngSize = 500
                        'Get Process Name
                        lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, lngSize)
                        
                        'Remove trailing spaces
                        strProcessName = Left(strModuleName, lngReturn)
    
                        'Check for Matching Upper case result
                        strProcessName = UCase$(Trim$(strProcessName))
                        strProcName2 = GetElement(Trim(Replace(strProcessName, Chr$(0), "")), "\", 0, 0, _
                                    GetNumElements(Trim(Replace(strProcessName, Chr$(0), "")), "\") - 1)
                        
                        ' all the items for the process
                        If strProcName2 = EXEName Then
                            For llLoop = 1 To llEnd
                            
                                lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(llLoop), strModuleName, lngSize)
                                
                                'Remove trailing spaces
                                strProcessName = Left(strModuleName, lngReturn)
            
                                'Check for Matching Upper case result
                                strProcessName = UCase$(Trim$(strProcessName))
                                strProcName2 = GetElement(Trim(Replace(strProcessName, Chr$(0), "")), "\", 0, 0, _
                                            GetNumElements(Trim(Replace(strProcessName, Chr$(0), "")), "\") - 1)
                            
                                Set lclsFile = New clsTask
                                lclsFile.Filename = strProcName2
                                lclsFile.Path = strProcessName
                                GetExtendedInfo lclsFile
                                loColRet.Add lclsFile
                                
                            Next
                            
                            'Get the Site of the Memory Structure
                            pmc.cb = LenB(pmc)
                            lRet = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
                            
                        End If
                    End If
                End If
                
                'Close the handle to this process
                lngReturn = CloseHandle(lngHwndProcess)
                'DoEvents
            Next
    
    End Select

IsProcessRunning_Exit:

'Exit early to avoid error handler
Exit Function
Error_handler:
    Err.Raise Err, Err.Source, "ProcessInfo", Error
    Resume Next
End Function


Private Function getOsVersion() As Long

    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    getOsVersion = osinfo.dwPlatformId

End Function


Public Function GetElement(ByVal strList As String, ByVal strDelimiter As String, ByVal lngNumColumns As Long, ByVal lngRow As Long, ByVal lngColumn As Long) As String

    Dim lngCounter As Long

    ' Append delimiter text to the end of the list as a terminator.
    strList = strList & strDelimiter

    ' Calculate the offset for the item required based on the number of columns the list
    ' 'strList' has i.e. 'lngNumColumns' and from which row the element is to be
    ' selected i.e. 'lngRow'.
    lngColumn = IIf(lngRow = 0, lngColumn, (lngRow * lngNumColumns) + lngColumn)

    ' Search for the 'lngColumn' item from the list 'strList'.
    For lngCounter = 0 To lngColumn - 1

        ' Remove each item from the list.
        strList = Mid$(strList, InStr(strList, strDelimiter) + Len(strDelimiter), Len(strList))

        ' If list becomes empty before 'lngColumn' is found then just
        ' return an empty string.
        If Len(strList) = 0 Then
            GetElement = ""
            Exit Function
        End If

    Next lngCounter

    ' Return the sought list element.
    GetElement = Left$(strList, InStr(strList, strDelimiter) - 1)

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Function GetNumElements (ByVal strList As String,
' ByVal strDelimiter As String)
' As Integer
'
' strList = The element list.
' strDelimiter = The delimiter by which the elements in
' 'strList' are seperated.
'
' The function returns an integer which is the count of the
' number of elements in 'strList'.
'
' Author: Roger Taylor
'
' Date:26/12/1998
'
' Additional Information:
'
' Revision History:
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetNumElements(ByVal strList As String, ByVal strDelimiter As String) As Integer

    Dim intElementCount As Integer

    ' If no elements in the list 'strList' then just return 0.
    If Len(strList) = 0 Then
        GetNumElements = 0
        Exit Function
    End If

    ' Append delimiter text to the end of the list as a terminator.
    strList = strList & strDelimiter

    ' Count the number of elements in 'strlist'
    While InStr(strList, strDelimiter) > 0
        intElementCount = intElementCount + 1
        strList = Mid$(strList, InStr(strList, strDelimiter) + 1, Len(strList))
    Wend

    ' Return the number of elements in 'strList'.
    GetNumElements = intElementCount

End Function

Private Sub GetExtendedInfo(roclsTask As clsTask)
Dim Buffer                  As String
Dim rc                      As Long
Dim FullFileName            As String
Dim iLoop                   As Integer
Dim lBufferLen              As Long
Dim lDummy                  As Long
Dim strVersionInfo          As Variant
Dim strTemp                 As String
Dim bytebuffer(255)         As Byte
Dim Lang_Charset_String     As String
Dim HexNumber               As Long
Dim lVerPointer             As Long
Dim sBuffer()               As Byte
Dim lVerbufferLen           As Long
Dim udtVerBuffer            As VS_FIXEDFILEINFO

    FullFileName = roclsTask.Path
    
    '*** Get size ****
    lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
    If lBufferLen < 1 Then
       Exit Sub
    End If
    
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, _
                            0&, _
                            lBufferLen, _
                            sBuffer(0))
    If rc = 0 Then
       Exit Sub
    End If
    
    ' get the size and date of the file.
    roclsTask.FileDate = FileDateTime(FullFileName)
    roclsTask.lFileSize = FileLen(FullFileName)
    
    ' get the file properties
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    
    With roclsTask
        .FileFlags = ""
        If udtVerBuffer.dwFileFlags And VS_FF_DEBUG Then
            .FileFlags = "Debug "
        End If
    
        If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE _
           Then .FileFlags = .FileFlags & "PreRel "
        If udtVerBuffer.dwFileFlags And VS_FF_PATCHED _
           Then .FileFlags = .FileFlags & "Patched "
        If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD _
           Then .FileFlags = .FileFlags & "Private "
        If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRED _
           Then .FileFlags = .FileFlags & "Info "
        If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD _
           Then .FileFlags = .FileFlags & "Special "
        If udtVerBuffer.dwFileFlags And VFT2_UNKNOWN _
           Then .FileFlags = .FileFlags + "Unknown "
    
        '**** Determine OS for which file was designed ****
        Select Case udtVerBuffer.dwFileOS
           Case VOS_DOS_WINDOWS16
             .FileOS = "DOS-Win16"
           Case VOS_DOS_WINDOWS32
             .FileOS = "DOS-Win32"
           Case VOS_OS216_PM16
             .FileOS = "OS/2-16 PM-16"
           Case VOS_OS232_PM32
             .FileOS = "OS/2-16 PM-32"
           Case VOS_NT_WINDOWS32
             .FileOS = "NT-Win32"
           Case Else
             .FileOS = "Unknown"
        End Select
        Select Case udtVerBuffer.dwFileType
           Case VFT_APP
              .FileType = "App"
           Case VFT_DLL
              .FileType = "DLL"
           Case VFT_DRV
              .FileType = "Driver"
              Select Case udtVerBuffer.dwFileSubtype
                 Case VFT2_DRV_PRINTER
                    .FileType = .FileType & " (Printer drv)"
                 Case VFT2_DRV_KEYBOARD
                    .FileType = .FileType & " (Keyboard drv)"
                 Case VFT2_DRV_LANGUAGE
                    .FileType = .FileType & " (Language drv)"
                 Case VFT2_DRV_DISPLAY
                    .FileType = .FileType & " (Display drv)"
                 Case VFT2_DRV_MOUSE
                    .FileType = .FileType & " (Mouse drv)"
                 Case VFT2_DRV_NETWORK
                    .FileType = .FileType & " (Network drv)"
                 Case VFT2_DRV_SYSTEM
                    .FileType = .FileType & " (System drv)"
                 Case VFT2_DRV_INSTALLABLE
                    .FileType = .FileType & " (Installable)"
                 Case VFT2_DRV_SOUND
                    .FileType = .FileType & " (Sound drv)"
                 Case VFT2_DRV_COMM
                    .FileType = .FileType & " (Comm drv)"
                 Case VFT2_UNKNOWN
                    .FileType = .FileType & " (Unknown)"
              End Select
           Case VFT_FONT
              .FileType = "Font"
              #If 0 Then
              Select Case udtVerBuffer.dwFileSubtype
                 Case VFT_FONT_RASTER
                    .FileType = .FileType & " (Raster Font)"
                 Case VFT_FONT_VECTOR
                    .FileType = .FileType & " (Vector Font)"
                 Case VFT_FONT_TRUETYPE
                    .FileType = .FileType & " (TrueType Font)"
              End Select
              #End If
              
           Case VFT_VXD
              .FileType = "VxD"
           Case VFT_STATIC_LIB
              .FileType = "Lib"
           Case Else
              .FileType = "Unknown"
        End Select
    End With
    
    rc = VerQueryValue(sBuffer(0), _
                      "\VarFileInfo\Translation", _
                      lVerPointer, _
                      lBufferLen)
    
    If rc = 0 Then
       Exit Sub
    End If
    
    'lVerPointer is a pointer to four 4 bytes of Hex number,
    'first two bytes are language id, and last two bytes are code
    'page. However, Lang_Charset_String needs a  string of
    '4 hex digits, the first two characters correspond to the
    'language id and last two the last two character correspond
    'to the code page id.
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + _
    bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    Lang_Charset_String = Hex(HexNumber)
    
    'now we change the order of the language id and code page
    'and convert it into a string representation.
    'For example, it may look like 040904E4
    'Or to pull it all apart:
    '04------        = SUBLANG_ENGLISH_USA
    '--09----        = LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows:Multilingual
    Do While Len(Lang_Charset_String) < 8
       Lang_Charset_String = "0" & Lang_Charset_String
    Loop
    
    strVersionInfo = Array("CompanyName", _
                   "FileDescription", _
                   "FileVersion", _
                   "InternalName", _
                   "LegalCopyright", _
                   "OriginalFileName", _
                   "ProductName", _
                   "ProductVersion", _
                   "Comments", _
                   "LegalTrademarks", _
                   "PrivateBuild", _
                   "SpecialBuild")
    
    
    ' get the infromation for each of the items
    For iLoop = 0 To UBound(strVersionInfo)
    
       ' get the info from the header
       Buffer = String(255, 0)
       strTemp = "\StringFileInfo\" & Lang_Charset_String & "\" & strVersionInfo(iLoop)
       rc = VerQueryValue(sBuffer(0), strTemp, lVerPointer, lBufferLen)
       
       ' place the info into the class
       If rc <> 0 Then
           lstrcpy Buffer, lVerPointer
           Buffer = Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1)
           Select Case strVersionInfo(iLoop)
               Case "CompanyName"
                   roclsTask.CompanyName = Buffer
               Case "FileDescription"
                   roclsTask.FileDescription = Buffer
               Case "FileVersion"
                   roclsTask.FileVersion = Buffer
               Case "InternalName"
                   roclsTask.InternalName = Buffer
               Case "LegalCopyright"
                   roclsTask.LegalCopyright = Buffer
               Case "OriginalFileName"
                   roclsTask.OriginalFileName = Buffer
               Case "ProductName"
                   roclsTask.ProductName = Buffer
               Case "ProductVersion"
                   roclsTask.ProductVersion = Buffer
               Case "Comments"
                   roclsTask.Comments = Buffer
               Case "LegalTrademarks"
                   roclsTask.LegalTrademarks = Buffer
               Case "PrivateBuild"
                   roclsTask.PrivateBuild = Buffer
               Case "SpecialBuild"
                   roclsTask.SpecialBuild = Buffer
           End Select
       End If
    Next iLoop
    
End Sub


