; ExportHostDrv.pb.  To Compile as Executable Console with Request Administrator
; 
; Export All Host Third-Party Driver Packages.
;
; ExportHostDrv.exe <Path_to_the_Destination_Folder>
;
;    Example: ExportHostDrv.exe D:\ExportHostDriver
;  
; Based on Driver Info (Structures SP_DRVINFO_DATA and DETAIL_DATA) by GoodNPlenty and JHPJHP 
; https://www.purebasic.fr/english/viewtopic.php?p=423755#p423755

EnableExplicit

CompilerIf #PB_Compiler_OS <> #PB_OS_Windows
  CompilerError "Error: Windows Only"
  End
CompilerEndIf

CompilerIf #PB_Compiler_Unicode = #False
  CompilerError "Error: Must Compile with Unicode"
  End
CompilerEndIf

If OSVersion() < #PB_OS_Windows_2000
  MessageBox_(#Null, "Invalid Operating System", "Requires Windows 2000 or Later", #MB_ICONERROR)
  End
EndIf

#PB_FileSystem_Link         = #FILE_ATTRIBUTE_REPARSE_POINT
#SPDIT_COMPATDRIVER         = $00000002
#CR_SUCCESS                 = $00000000
#DI_QUIETINSTALL            = $00800000
#DI_NOFILECOPY              = $01000000
#DI_FLAGSEX_INSTALLEDDRIVER	= $04000000
#MAX_CLASS_NAME_LEN         = 32
#LINE_LEN                   = 256        ;// Windows 9x max for displayable strings from a device INF.

Structure SP_DEVINFO_DATA Align #PB_Structure_AlignC
  cbSize.l
  ClassGuid.GUID
  DevInst.l
  *Reserved
EndStructure

Structure SP_DEVINSTALL_PARAMS Align #PB_Structure_AlignC
  cbSize.l
  Flags.l
  FlagsEx.l
  hwndParent.i 
  *InstallMsgHandler
  *InstallMsgHandlerContext
  *FileQueue
  *ClassInstallReserved
  Reserved.l
  DriverPath.c[#MAX_PATH]
EndStructure

Structure SP_DRVINFO_DATA Align #PB_Structure_AlignC
  cbSize.l
  DriverType.l
  *Reserved
  Description.c[#LINE_LEN]
  MfgName.c[#LINE_LEN]
  ProviderName.c[#LINE_LEN]
  DriverDate.FILETIME
  DriverVersion.u[4]
EndStructure

Structure SP_DRVINFO_DETAIL_DATA Align #PB_Structure_AlignC
  cbSize.l
  InfDate.FILETIME
  CompatIDsOffset.l
  CompatIDsLength.l
  *Reserved
  SectionName.c[#LINE_LEN]
  InfFileName.c[#MAX_PATH]
  DrvDescription.c[#LINE_LEN]
  HardwareID.c[#ANYSIZE_ARRAY]
EndStructure

Structure ExportDrvInfo
  DeviceIndex.i
  DriverIndex.i
  InstanceID.s
  HardwareID.s
  InfFileName.s
  SectionName.s
  ClassName.s
  ClassDescription.s
  DrvDescription.s
  InfDriverStorePath.s
  InfDriverStoreSize.i
  DriverStoreFolder.s
  OrgInfFileName.S
  ProviderName.s
  DriverDate.s
  DriverVersion.s
EndStructure
Global NewList ExportDrv.ExportDrvInfo()

CompilerIf #PB_Compiler_Processor = #PB_Processor_x64
CompilerElseIf #PB_Compiler_Processor = #PB_Processor_x86
CompilerElse
  CompilerError "Error: Unsupported Processor"
  End
CompilerEndIf

Prototype CM_Get_DevNode_Status(*Status, *ProblemNumber, DevInst, Flags)
Prototype SetupGetInfDriverStoreLocation(InfFileName, AlternatePlatformInfo.l, LocaleName.l, InfDriverStorePath, InfDriverStorePathSize.l, RequiredSize.l)

Global DestPath$

Procedure Usage()
  PrintN(GetFilePart(ProgramFilename()) + " <Path_to_the_Destination_Folder>")
  PrintN("")
  PrintN("    Example: " + GetFilePart(ProgramFilename()) + " D:\ExportHostDriver")
  PrintN("")
  Delay(1000)
EndProcedure

Procedure PrintNError(text.s)
  ConsoleColor(12, 0)
  PrintN(text)
  PrintN("")
  ConsoleColor(15, 0)
EndProcedure

Procedure.s GetLastError(LastError.i)
  ; Based on ABBKlaus procedure
  Protected *Buffer, BufferLength.i, Flags.i, ErrorMessage$
  Flags = #FORMAT_MESSAGE_ALLOCATE_BUFFER|#FORMAT_MESSAGE_FROM_SYSTEM
  
  BufferLength = FormatMessage_(Flags, 0, LastError, GetUserDefaultLangID_(), @*Buffer, 0, 0)
  If BufferLength
    ErrorMessage$ = PeekS(*Buffer)
    LocalFree_(*Buffer)
    ErrorMessage$ = RemoveString(ErrorMessage$, Chr(13)+Chr(10))
    ProcedureReturn "Error Code: " + Str(LastError) + " Error Message: " + ErrorMessage$
  Else
    ProcedureReturn Str(LastError)
  EndIf
EndProcedure

Procedure.s GetSystemProductName()
  Define SystemProductName$, hKey.l, Type.l, cbData.l = 256, lpbData.l = AllocateMemory(256)
  
  If RegOpenKeyEx_(#HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SystemInformation", 0, #KEY_QUERY_VALUE , @hKey) = #ERROR_SUCCESS
    If hKey
      If RegQueryValueEx_(hKey, "SystemProductName", 0, @Type, lpbData, @cbData) = #ERROR_SUCCESS
        SystemProductName$ = PeekS(lpbData)
      EndIf
      RegCloseKey_(hKey)
    EndIf
  EndIf
  FreeMemory(lpbData)
  
  ProcedureReturn SystemProductName$
EndProcedure

Procedure.s SizeIt(Value.q)
  ; SizeIt: Format Size in bytes, KB, MB, GB, TB, PB
  Protected unit.b=0, byte.q, pos.l, nSize.s
  byte = Value
  While byte >= 1024
    byte / 1024 : unit + 1
  Wend
  If unit : nSize = StrD(Value/Pow(1024, unit), 15) : pos.l = FindString(nSize, ".") : Else : nSize = Str(Value) : EndIf
  If unit : If pos <  4 : nSize=Mid(nSize,1,pos+2) : Else : nSize = Mid(nSize, 1, pos-1) : EndIf : EndIf
  ;ProcedureReturn nSize+Space(1)+StringField("b ,KB,MB,GB,TB,PB", unit+1, ",")
  ProcedureReturn nSize+StringField("b ,Kb,Mb,Gb,Tb,Pb", unit+1, ",")
EndProcedure

Procedure.q GetFolderSize(Path.s, State.b = 1)
  ;GetFolderSize by Thunder93 - Test: Path.s = #PB_Compiler_Home :  Debug #CRLF$+"GetFolderSize()" :  Debug "Path: " + Path :  Debug "  "+Str(GetFolderSize(Path)) + " Bytes"
  ;https://www.purebasic.fr/english/viewtopic.php?f=12&t=65963
  Protected Path$ = Path, hF
  Static.q nSize
  
  If State : nSize = 0 : EndIf
  
  hF = ExamineDirectory(#PB_Any, Path, "*.*")
  If hF <> #False   
    While NextDirectoryEntry(hF)
      If DirectoryEntryType(hF) = #PB_DirectoryEntry_Directory
        If DirectoryEntryName(hF) <> "." And DirectoryEntryName(hF) <> ".."
          
          If DirectoryEntryAttributes(hF) & #PB_FileSystem_Link
            Continue
          EndIf
          
          Path$ = Path + "\" + DirectoryEntryName(hF) 
          GetFolderSize(Path$, 0) ; Traverse Directory
        EndIf
      Else
        nSize + FileSize(Path + "\"  + DirectoryEntryName(hF))
      EndIf
    Wend
    FinishDirectory(hF)
  EndIf
  
  ProcedureReturn nSize
EndProcedure

Procedure.s InitialCaps(text.s)
  ;Fastest way for InitialCaps by Demivec and wilbert
  ;https://www.purebasic.fr/english/viewtopic.php?p=406482#p406482
  Protected *b.Character = @text, cur, wordStarted = #False
  
  cur = *b\c
  While cur
    Select cur
      Case 97 To 122
        If Not wordStarted
          *b\c = cur - 32
          wordStarted = #True
        EndIf
      Case 65 To 90
        If wordStarted
          *b\c = cur + 32
        Else
          wordStarted = #True
        EndIf
      Case 39, '-', '_' ;apostrophe, hyphen and underscore treated as part of word
      Default
        wordStarted = #False
    EndSelect
    
    *b + SizeOf(Character)
    cur = *b\c
  Wend 
  
  ProcedureReturn text
EndProcedure

Procedure.s ValidPath(Path.s, ReplaceChar.s = " ")
  ;Forbidden path & Filename in Windows : /\:*?"<>|  ...
  If CheckFilename(Path) = 0
    Path = ReplaceString(Path, "/", ReplaceChar)
    Path = ReplaceString(Path, "\", ReplaceChar)
    Path = ReplaceString(Path, ":", ReplaceChar)
    Path = ReplaceString(Path, "*", ReplaceChar)
    Path = ReplaceString(Path, "?", ReplaceChar)
    Path = ReplaceString(Path, "<", ReplaceChar)
    Path = ReplaceString(Path, ">", ReplaceChar)
    Path = ReplaceString(Path, "|", ReplaceChar)
  EndIf
  ProcedureReturn Path
EndProcedure

Procedure GetDriverAndInfClassInfo()
  Protected DeviceInfoData.SP_DEVINFO_DATA
  Protected DeviceInstallParms.SP_DEVINSTALL_PARAMS
  Protected DriverInfoData.SP_DRVINFO_DATA
  Protected *DriverInfoDetailData.SP_DRVINFO_DETAIL_DATA
  Protected SetupGetInfDriverStoreLocation_.SetupGetInfDriverStoreLocation
  
  Protected.i libSetupAPI, hDevice, DeviceIndex, ConfigReturn, DriverIndex, InfPathLen, DriverInfoDetailSize, ActualDriverInfoDetailSize
  Protected.i Status, Problem, Result, InstanceIdStatus, GetDriverInfoDetailStatus, DeviceInstanceIdSize
  Protected InfFileName$, InfClassFileName$, InfPath$, ClassName$, ClassDescription$, InfDriverStorePath$, ClassGUID$, CopyFiles$, InstanceID$
  Protected CM_Get_DevNode_Status_.CM_Get_DevNode_Status
  Protected ClassGUID.GUID
  Protected *DeviceInstanceId
  Protected sysTime.SYSTEMTIME 
  
  DeviceInfoData\cbSize     = SizeOf(SP_DEVINFO_DATA)
  DeviceInstallParms\cbSize = SizeOf(SP_DEVINSTALL_PARAMS)
  DriverInfoData\cbSize     = SizeOf(SP_DRVINFO_DATA)
  
  ClassName$  = Space(#MAX_CLASS_NAME_LEN)
  ClassDescription$ = Space(#MAX_PATH)
  InfPath$    = GetEnvironmentVariable("SystemRoot")+"\inf"
  InfPathLen  = (Len(InfPath$)*SizeOf(Character))+SizeOf(Character)
  InfDriverStorePath$ = Space(#MAX_PATH)
  
  libSetupAPI = OpenLibrary(#PB_Any, "setupapi.dll")
  If libSetupAPI
    CM_Get_DevNode_Status_ = GetFunction(libSetupAPI, "CM_Get_DevNode_Status")
    If Not CM_Get_DevNode_Status_
      PrintNError("ERROR: GetFunction Error: Unable to GetFunction CM_Get_DevNode_Status")
      CloseLibrary(libSetupAPI)
      End
    EndIf
    SetupGetInfDriverStoreLocation_ = GetFunction(libSetupAPI, "SetupGetInfDriverStoreLocationW")
    If Not SetupGetInfDriverStoreLocation_
      PrintNError("ERROR: GetFunction Error: Unable to GetFunction SetupGetInfDriverStoreLocation")
      CloseLibrary(libSetupAPI)
      End
    EndIf
  Else   
    PrintNError("ERROR: OpenLibrary Error: Unable to OpenLibrary(setupapi.dll)")
    End
  EndIf
  
  hDevice = SetupDiGetClassDevs_(#Null, 0, 0, #DIGCF_ALLCLASSES|#DIGCF_PRESENT)
  If hDevice = #INVALID_HANDLE_VALUE
    PrintNError("ERROR: SetupDiGetClassDevs: "+GetLastError(GetLastError_()))
    ProcedureReturn GetLastError_()
  EndIf
  
  If Not SetupDiEnumDeviceInfo_(hDevice, DeviceIndex, @DeviceInfoData)
    PrintNError("ERROR: SetupDiEnumDeviceInfo: "+GetLastError(GetLastError_()))
    ProcedureReturn GetLastError_()
  EndIf   
  
  While GetLastError_() <> #ERROR_NO_MORE_ITEMS
    InstanceIdStatus = SetupDiGetDeviceInstanceId_(hDevice, @DeviceInfoData, #Null, 0, @DeviceInstanceIdSize)
    If InstanceIdStatus = #False And GetLastError_() = #ERROR_INSUFFICIENT_BUFFER
      DeviceInstanceIdSize = (DeviceInstanceIdSize * SizeOf(Character))
      *DeviceInstanceId = AllocateMemory(DeviceInstanceIdSize + SizeOf(Character))
      If *DeviceInstanceId
        If SetupDiGetDeviceInstanceId_(hDevice, @DeviceInfoData, *DeviceInstanceId, DeviceInstanceIdSize, #Null)
          InstanceID$ = PeekS(*DeviceInstanceId)
          FreeMemory(*DeviceInstanceId)
          ConfigReturn = CM_Get_DevNode_Status_(@Status, @Problem, DeviceInfoData\DevInst, 0)
          If ConfigReturn = #CR_SUCCESS
            If SetupDiGetDeviceInstallParams_(hDevice, @DeviceInfoData, @DeviceInstallParms)
              DeviceInstallParms\Flags = #DI_QUIETINSTALL|#DI_NOFILECOPY
              DeviceInstallParms\FlagsEx = #DI_FLAGSEX_INSTALLEDDRIVER
              CopyMemory(@InfPath$, @DeviceInstallParms\DriverPath, InfPathLen) 
              If SetupDiSetDeviceInstallParams_(hDevice, @DeviceInfoData, @DeviceInstallParms)
                If SetupDiBuildDriverInfoList_(hDevice, @DeviceInfoData, #SPDIT_COMPATDRIVER)
                  DriverIndex = 0
                  Repeat
                    If SetupDiEnumDriverInfo_(hDevice, @DeviceInfoData, #SPDIT_COMPATDRIVER, DriverIndex, @DriverInfoData)
                      GetDriverInfoDetailStatus = SetupDiGetDriverInfoDetail_(hDevice, @DeviceInfoData, @DriverInfoData, 0, 0, @DriverInfoDetailSize)
                      If GetDriverInfoDetailStatus = #False And GetLastError_() = #ERROR_INSUFFICIENT_BUFFER
                        *DriverInfoDetailData = AllocateMemory(DriverInfoDetailSize + SizeOf(Character))
                        If *DriverInfoDetailData
                          *DriverInfoDetailData\cbSize = SizeOf(SP_DRVINFO_DETAIL_DATA) - (SizeOf(SP_DRVINFO_DETAIL_DATA)%8/2)
                          If SetupDiGetDriverInfoDetail_(hDevice, @DeviceInfoData, @DriverInfoData, *DriverInfoDetailData,  DriverInfoDetailSize, @ActualDriverInfoDetailSize)
                            ;Only OEM drivers
                            InfFileName$    = PeekS(@*DriverInfoDetailData\InfFileName)
                            If Not PathMatchSpec_(@InfFileName$, "*\Inf\oem*.inf")
                              DriverIndex + 1
                              FreeMemory(*DriverInfoDetailData)
                              Continue
                            EndIf
                            If SetupGetInfDriverStoreLocation_(@InfFileName$, 0, 0, @InfDriverStorePath$, #MAX_PATH, 0)
                              If SetupDiGetINFClass_(@InfFileName$ , @ClassGUID, @ClassName$, #MAX_CLASS_NAME_LEN, #Null)
                                If SetupDiGetClassDescription_(@ClassGUID, @ClassDescription$, #MAX_PATH, 0)
                                  
                                  AddElement(ExportDrv())
                                  With ExportDrv()
                                    \DeviceIndex        = DeviceIndex
                                    \DriverIndex        = DriverIndex
                                    \InstanceID         = InstanceID$
                                    \HardwareID         = PeekS(@*DriverInfoDetailData\HardwareID)
                                    \InfFileName        = PeekS(@*DriverInfoDetailData\InfFileName)
                                    \SectionName        = PeekS(@*DriverInfoDetailData\SectionName)
                                    \ClassName          = ClassName$
                                    \ClassDescription   = ClassDescription$
                                    \DrvDescription     = PeekS(@*DriverInfoDetailData\DrvDescription)
                                    \DrvDescription     = ValidPath(\DrvDescription, "-")
                                    \OrgInfFileName     = GetFilePart(InfDriverStorePath$)
                                    \InfDriverStorePath = GetPathPart(InfDriverStorePath$)
                                    \InfDriverStoreSize = GetFolderSize(\InfDriverStorePath)
                                    \DriverStoreFolder  = GetFilePart(RTrim(\InfDriverStorePath, "\"))
                                    \ProviderName       = PeekS(@DriverInfoData\ProviderName, #PB_Ascii)
                                    \ProviderName       = InitialCaps(\ProviderName)
                                    FileTimeToSystemTime_(@DriverInfoData\DriverDate, sysTime)
                                    \DriverDate         = Str(sysTime\wYear)+"."+RSet(Str(sysTime\wMonth),2,"0")+"."+RSet(Str(sysTime\wDay),2,"0")
                                    \DriverVersion      = Str(DriverInfoData\DriverVersion[3]) + "." +
                                                          Str(DriverInfoData\DriverVersion[2]) + "." +
                                                          Str(DriverInfoData\DriverVersion[1]) + "." +
                                                          Str(DriverInfoData\DriverVersion[0])
                                  EndWith
                                  
                                Else      ;// If SetupDiGetClassDescription_(@ClassGUID, @ClassDescription$, #MAX_PATH, 0)
                                  PrintNError("ERROR: SetupDiGetClassDescription : "+GetLastError(GetLastError_()))
                                EndIf      ;// If SetupDiGetClassDescription_(@ClassGUID, @ClassDescription$, #MAX_PATH, 0)
                              Else         ;// If SetupDiGetINFClass_(@InfFileName$ , @ClassGUID, @ClassName$, #MAX_CLASS_NAME_LEN, #Null)
                                PrintNError("ERROR: SetupDiGetINFClass : "+GetLastError(GetLastError_()))
                              EndIf     ;// If SetupDiGetINFClass_(@InfFileName$ , @ClassGUID, @ClassName$, #MAX_CLASS_NAME_LEN, #Null)
                            Else        ;// If SetupGetInfDriverStoreLocation_(@InfFileName$, 0, 0, @InfDriverStorePath$, #MAX_PATH, 0)
                              PrintNError("ERROR: SetupGetInfDriverStoreLocation : "+GetLastError(GetLastError_())) 
                            EndIf       ;// If SetupGetInfDriverStoreLocation_(@InfFileName$, 0, 0, @InfDriverStorePath$, #MAX_PATH, 0)
                          Else          ;// If SetupDiGetDriverInfoDetail_(hDevice, @DeviceInfoData, @DriverInfoData, *DriverInfoDetailData,  DriverInfoDetailSize, @ActualDriverInfoDetailSize)
                            PrintNError("ERROR: SetupDiGetDriverInfoDetail: "+GetLastError(GetLastError_()))
                          EndIf       ;// If SetupDiGetDriverInfoDetail_(hDevice, @DeviceInfoData, @DriverInfoData, *DriverInfoDetailData,  DriverInfoDetailSize, @ActualDriverInfoDetailSize)
                          FreeMemory(*DriverInfoDetailData)
                        Else          ;// If *DriverInfoDetailData
                          PrintNError("ERROR: Memory Allocation SetupDiGetDriverInfoDetail")
                        EndIf         ;// If *DriverInfoDetailData                         
                      Else            ;// If GetDriverInfoDetailStatus = #False And GetLastError_() = #ERROR_INSUFFICIENT_BUFFER
                        PrintNError("ERROR: SetupDiGetDriverInfoDetail: "+GetLastError(GetLastError_()))
                      EndIf           ;// If GetDriverInfoDetailStatus = #False And GetLastError_() = #ERROR_INSUFFICIENT_BUFFER
                    ElseIf GetLastError_() = #ERROR_NO_MORE_ITEMS             ;// If SetupDiEnumDriverInfo_(hDevice, @DeviceInfoData, #SPDIT_COMPATDRIVER, DriverIndex, @DriverInfoData)
                      Break
                    Else             ;// If SetupDiEnumDriverInfo_(hDevice, @DeviceInfoData, #SPDIT_COMPATDRIVER, DriverIndex, @DriverInfoData)
                      PrintNError("ERROR: SetupDiEnumDriverInfo: "+GetLastError(GetLastError_()))
                    EndIf             ;// If SetupDiEnumDriverInfo_(hDevice, @DeviceInfoData, #SPDIT_COMPATDRIVER, DriverIndex, @DriverInfoData)
                    DriverIndex + 1
                  ForEver 
                  SetupDiDestroyDriverInfoList_(hDevice, @DeviceInfoData, #SPDIT_COMPATDRIVER)
                EndIf                 ;// If SetupDiBuildDriverInfoList_(hDevice, @DeviceInfoData, #SPDIT_COMPATDRIVER)
              EndIf                   ;// If SetupDiSetDeviceInstallParams_(hDevice, @DeviceInfoData, @DeviceInstallParms)
            EndIf                     ;// If SetupDiGetDeviceInstallParams_(hDevice, @DeviceInfoData, @DeviceInstallParms)
          EndIf                       ;// If ConfigReturn = #CR_SUCCESS
        EndIf                         ;// If SetupDiGetDeviceInstanceId_(hDevice, @DeviceInfoData, *DeviceInstanceId, DeviceInstanceIdSize, #Null)
      Else                            ;// If *DeviceInstanceId
        PrintNError("ERROR: Memory Allocation SetupDiGetDeviceInstanceId")
      EndIf                           ;// If *DeviceInstanceId
    Else                              ;// If GetDriverInfoDetailStatus = #False And GetLastError_() = #ERROR_INSUFFICIENT_BUFFER 'SetupDiGetDeviceInstanceId)
    EndIf                             ;// If GetDriverInfoDetailStatus = #False And GetLastError_() = #ERROR_INSUFFICIENT_BUFFER
    DeviceIndex + 1
    SetupDiEnumDeviceInfo_(hDevice, DeviceIndex, @DeviceInfoData)   
  Wend                                ;// While GetLastError_() <> #ERROR_NO_MORE_ITEMS
  
  SetupDiDestroyDeviceInfoList_(hDevice)
  If Result = #ERROR_NO_MORE_ITEMS : Result = #True : EndIf
  
  CloseLibrary(libSetupAPI)
  
  ProcedureReturn Result
EndProcedure

Procedure ListDriver(hFile = 0) 
  Protected SavClassDescription$, NbDriver, Result
  
  NbDriver = ListSize(ExportDrv())
  If NbDriver
    Result = #True
    
    If hFile = 0
      PrintN("List of " + NbDriver + " Third-Party Driver Packages :")
      PrintN("")
    EndIf
    
    With ExportDrv()
      ForEach ExportDrv()
        If \ClassDescription <> SavClassDescription$
          If hFile
            WriteStringN(hFile, \ClassDescription)
          Else
            PrintN(\ClassDescription)
          EndIf
          SavClassDescription$ = \ClassDescription
        EndIf
        If hFile
          WriteStringN(hFile, "  |- "+ \DrvDescription +"--"+ \ProviderName +"--"+ \OrgInfFileName +"-v"+ \DriverVersion + " (" + SizeIt(\InfDriverStoreSize) + ")")
        Else
          PrintN("  |- "+ \DrvDescription +"--"+ \ProviderName +"--"+ \OrgInfFileName +"-v"+ \DriverVersion + " (" + SizeIt(\InfDriverStoreSize) + ")")
        EndIf
      Next
    EndWith
  Else
    PrintNError("No Third-Party Driver Package Found")
  EndIf
  
  ProcedureReturn Result
EndProcedure

Procedure CopyDriver() 
  Protected hFile, FileName$, SavClassDescription$, NbDriver, DriverIndex = 1, Result
  
  NbDriver = ListSize(ExportDrv())
  If NbDriver
    Result = #True
    
    FileName$ = ValidPath(GetSystemProductName()) + " Driver " + FormatDate("%yyyy.%mm.%dd %hh-%ii-%ss", Date())
    hFile = CreateFile(#PB_Any, DestPath$ + "\" + FileName$ + ".txt")
    If hFile
      WriteStringN(hFile, FileName$)
      WriteStringN(hFile, "")
      WriteStringN(hFile, "Export of " + NbDriver + " Third-Party Driver Packages :")
      WriteStringN(hFile, "")
      ListDriver(hFile)
      CloseFile(hFile)
    EndIf 
    
    PrintN("Export of " + NbDriver + " Third-Party Driver Packages :")
    PrintN("")
    
    With ExportDrv()
      ForEach ExportDrv()
        
        If \ClassDescription <> SavClassDescription$
          If FileSize(DestPath$ + "\" + \ClassDescription +"\") <> -2
            If CreateDirectory(DestPath$ + "\" + \ClassDescription +"\") = 0
              PrintNError("Invalid Directory Name.")
            EndIf
          EndIf
          PrintN(\ClassDescription)
          SavClassDescription$ = \ClassDescription
        EndIf
        
        PrintN("  |- "+ \DrvDescription +"--"+ \ProviderName +"--"+ \OrgInfFileName +"-v"+ \DriverVersion + " (" + SizeIt(\InfDriverStoreSize) + ")")
        Print("  |  > Exporting " + Str(DriverIndex) + " of " + Str(NbDriver) + " - " + GetFilePart(\InfFileName) + " (" + \DriverStoreFolder + ").")
        If CopyDirectory(\InfDriverStorePath, DestPath$ +"\"+ \ClassDescription +"\"+ \DrvDescription +"--"+ \ProviderName +"--"+ \OrgInfFileName +"-v"+ \DriverVersion +"\", "", #PB_FileSystem_Recursive|#PB_FileSystem_Force)
          ConsoleColor(1, 0)
          PrintN(" Successfully Exported.")
          ConsoleColor(15, 0)
        Else
          PrintNError(" Failed to Export")
        EndIf
        
        DriverIndex + 1
      Next
    EndWith
  EndIf
  
  ProcedureReturn Result
EndProcedure

;- -- Main --
OpenConsole()
PrintN("")
PrintN("Export Host Third-Party Driver Packages.")
PrintN("")

If CountProgramParameters() = 1
  DestPath$ = ProgramParameter(0)
  If FileSize(DestPath$) <> -2
    PrintNError("The Destination Path " + DestPath$ +" Does Not Exist.")
    Usage()
    End 1
  EndIf
  RTrim(DestPath$, "\")
  ; Possible check: The Destination Path DestPath$ Is Not Empty.
Else
  Usage()
  ; + ListDriver()
EndIf

GetDriverAndInfClassInfo()

If ListSize(ExportDrv())     ;Sort ExportDrv() List by ClassDescription and DrvDescription
  SortStructuredList(ExportDrv(), #PB_Sort_Ascending, OffsetOf(ExportDrvInfo\DrvDescription),TypeOf(ExportDrvInfo\DrvDescription))
  SortStructuredList(ExportDrv(), #PB_Sort_Ascending, OffsetOf(ExportDrvInfo\ClassDescription),TypeOf(ExportDrvInfo\ClassDescription))
EndIf

If FileSize(DestPath$) = -2
  CopyDriver()
Else
  ListDriver()
EndIf
Delay(1000)
; IDE Options = PureBasic 5.73 LTS (Windows - x64)
; ExecutableFormat = Console
; Folding = ---
; EnableXP
; EnableAdmin
; UseIcon = ExportHostDrv.ico
; Executable = ExportHostDrv_x64.exe
; CommandLine = E:\PureBasic\Tools\ExportHostDrv\ExpDriver\
; Compiler = PureBasic 5.73 LTS (Windows - x64)
; IncludeVersionInfo
; VersionField0 = 1.0.0.0
; VersionField1 = 1.0.0.0
; VersionField2 = ChrisR
; VersionField3 = ExpHostDrv
; VersionField4 = 1.0.0.0
; VersionField5 = 1.0.0.0
; VersionField6 = Export Host Third-Party (OEM) Driver Packages
; VersionField7 = ExpHostDrv.exe
; VersionField8 = ExpHostDrv.exe
; VersionField9 = ChrisR