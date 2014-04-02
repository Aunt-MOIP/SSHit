Attribute VB_Name = "ModCLI"
Dim CmdStr, originCmdStr As String


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'私有的数据结构申明


Private Type STARTUPINFO '(createprocess)
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Type PROCESS_INFORMATION '(creteprocess)
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type
Private Type SECURITY_ATTRIBUTES '(createprocess)
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
'常数声明
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_QUERY_INFORMATION = &H400
'函数声明
Private Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" _
                          (ByVal hNamedPipe As Long, _
                           ByVal lpBuffer As Long, _
                           ByVal nBufferSize As Long, _
                           ByRef lpBytesRead As Long, _
                           ByRef lpTotalBytesAvail As Long, _
                           ByRef lpBytesLeftThisMessage As Long _
                           ) As Long

Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hHandle As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
                                                   ByVal lpBuffer As Long, _
                                                   ByVal nNumberOfBytesToWrite As Long, _
                                                   ByRef lpNumberOfBytesWritten As Long, _
                                                   lpOverlapped As Any) As Long

Private Declare Function DuplicateHandle Lib "kernel32" _
               (ByVal hSourceProcessHandle As Long, _
                ByVal hSourceHandle As Long, _
                ByVal hTargetProcessHandle As Long, _
                lpTargetHandle As Long, _
                ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwOptions As Long) As Long
 


Private Const DUPLICATE_SAME_ACCESS = &H2
Private PipeR4InputChannel As Long, PipeW4InputChannel As Long, hInputHandle As Long
Private PipeR4OutputChannel As Long, PipeW4OutputChannel As Long, hOutputHandle As Long
Private Proc As PROCESS_INFORMATION
Public Enum InitResult
  ERROR_OK = 0
  ERROR_INIT_INPUT_HANDLE = 1
  ERROR_INIT_OUTPUT_HANDLE = 2
  ERROR_DUP_READ_HANDLE = 3
  ERROR_DUP_WRITE_HANDLE = 4
  ERROR_CREATE_CHILD_PROCESS = 5
End Enum

Public Enum TermResult
 ERROR_OK = 0
End Enum
Public Enum InputResult
 ERROR_OK = 0
 ERROR_QUERY_WRITE_INFO_SIZE = 1
 ERROR_DATA_TO_LARGE = 2
 ERROR_WRITE_INFO = 3
 ERROR_WRITE_UNEXPECTED = 5
End Enum
Public Enum OutputResult
 ERROR_OK = 0
 ERROR_QUERY_READ_INFO_SIZE = 1
 ERROR_ZERO_INFO_SIZE = 2
 ERROR_READ_INFO = 3
 ERROR_UNEQUAL_INFO_SIZE = 4
 ERROR_READ_UNEXPECTED = 5
End Enum
Public Function InitDosIO() As InitResult
   Dim Sa As SECURITY_ATTRIBUTES, Ret As Long
   
   With Sa
      .nLength = Len(Sa)
      .bInheritHandle = 1&
      .lpSecurityDescriptor = 0&
   End With
    
    Ret = CreatePipe(PipeR4InputChannel, PipeW4InputChannel, Sa, 1024&)
    If Ret = 0 Then '建立进程输入管道
        InitDosIO = ERROR_INIT_INPUT_HANDLE
        Exit Function
    End If
    
    Ret = CreatePipe(PipeR4OutputChannel, PipeW4OutputChannel, Sa, 4096&) '建立输出通道,若建立失败，则关闭管道，退出
    If Ret = 0 Then '建立进程的输出管道
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        InitDosIO = ERROR_INIT_OUTPUT_HANDLE
        Exit Function
    End If
  

   Ret = DuplicateHandle(GetCurrentProcess(), PipeW4InputChannel, GetCurrentProcess(), hInputHandle, 0, True, DUPLICATE_SAME_ACCESS)
   If Ret = 0 Then '转换写句柄
     CloseHandle PipeR4InputChannel
     CloseHandle PipeW4InputChannel
     CloseHandle PipeR4OutputChannel
     CloseHandle PipeW4OutputChannel
     InitDosIO = ERROR_DUP_WRITE_HANDLE
     Exit Function
   End If
   Ret = CloseHandle(PipeW4InputChannel)
   If Ret = 0 Then
    'MsgBox "close handle eerr"
   End If
  Ret = DuplicateHandle(GetCurrentProcess(), PipeR4OutputChannel, GetCurrentProcess(), hOutputHandle, 0, True, DUPLICATE_SAME_ACCESS)
   If Ret = 0 Then '转换读句柄
     CloseHandle PipeR4InputChannel
     CloseHandle PipeW4InputChannel
     CloseHandle PipeR4OutputChannel
     CloseHandle PipeW4OutputChannel
     InitDosIO = ERROR_DUP_READ_HANDLE
     Exit Function
   End If
 Ret = CloseHandle(PipeR4OutputChannel)
 If Ret = 0 Then
  'MsgBox "close handle 2 er"
 End If
   
    
    Dim Start As STARTUPINFO
    Start.cb = Len(Start)
    Start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    Start.hStdOutput = PipeW4OutputChannel
    Start.hStdError = PipeW4OutputChannel
    Start.hStdInput = PipeR4InputChannel
    '##########################################################################################
    '##########################################################################################
    '##########################################################################################
    '##########################################################################################
    '##########################################################################################
    'For Compile Public
    ' CmdStr = "plink.exe " & Form1.IP & " -N -ssh -2 -P " & Form1.Port & " -l " & Form1.User & " -C -D 7070 -v"    '###########
     'originCmdStr = "plink.exe 172.17.18.5 -N -ssh -2 -P 22 -l root -C -D 7070 -v -pw alpine" '###########
     'CmdStr = originCmdStr
    '/For Compile
    
    'CmdStr = "C:\plink.exe 172.17.18.11 -N -ssh -2 -P 1165 -l root -C -D 7070 -v" '###########
   
    '##########################################################################################
    '##########################################################################################
    '##########################################################################################
    '##########################################################################################
    '##########################################################################################
    
    Ret& = CreateProcessA(0&, CmdStr, Sa, Sa, True, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, Proc)
    
    If Ret <> 1 Then '建立控制进程
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        CloseHandle PipeR4OutputChannel
        CloseHandle PipeW4OutputChannel
        InitDosIO = ERROR_CREATE_CHILD_PROCESS
        Exit Function
    End If
    
End Function
Public Function DosInput(ByVal Str As String) As InputResult
 Dim Btarray  As String, Buflen As Long, BtWritten As Long, rtn As Long
 Dim BtTest() As Byte
 Btarray = StrConv(Str + vbCrLf, vbFromUnicode)
 BtTest = StrConv(Str + vbCrLf, vbFromUnicode)
 Buflen = LenB(Btarray)
 rtn = WriteFile(hInputHandle, StrPtr(BtTest), Buflen, BtWritten, ByVal 0&)
 
 If BtWritten = 0 Then
   DosInput = ERROR_WRITE_INFO
   Exit Function
 End If
 DosInput = 0
End Function

Public Function DosOutput(ByRef StrOutput As String) As OutputResult
  Dim Ret As Long, TmpBuf As String * 128, BtRead As Long, BtTotal As Long, BtLeft As Long
  rtn = PeekNamedPipe(hOutputHandle, StrPtr(TmpBuf), 128, BtRead, BtTotal, BtLeft)
  If rtn = 0 Then '查询信息量
    DosOutput = ERROR_QUERY_INFO_SIZE
    Exit Function
  End If
  
  If BtTotal = 0 Then '若信息为空，则退出
    DosOutput = ERROR_ZERO_INFO_SIZE
    Exit Function
  End If
  
  Dim Btbuf() As Byte, BtReaded As Long
  ReDim Btbuf(BtTotal)
  Ret = ReadFile(hOutputHandle, VarPtr(Btbuf(0)), BtTotal, lngbytesread, 0&)
  
  If Ret = 0 Then
    DosOutput = ERROR_READ_INFO
    Exit Function
  End If
  If BtTotal <> lngbytesread Then
   DosOutput = ERROR_UNEQUAL_INFO_SIZE
  End If
     
 Dim strBuf As String
 strBuf = StrConv(Btbuf, vbUnicode)
 Debug.Print strBuf
 StrOutput = strBuf
End Function
Public Function EndDosIo() As Long
 Dim Ret As Long
 CloseHandle PipeR4InputChannel
 CloseHandle PipeW4InputChannel
 CloseHandle PipeR4OutputChannel
 CloseHandle PipeW4OutputChannel
 CloseHandle Proc.hThread
 CloseHandle Proc.hProcess
If EndProcess(Proc.dwProcessId) = False Then
  ' MsgBox "主服务程序[CMD.EXE]没有关闭，请您手动关闭 ", vbInformation, "不好意思"
End If
End Function

Public Function EndProcess(ByVal ProcessID As Long) As Boolean
  Dim hProcess As Long, ExitCode As Long, Rst As Long
  hProcess = OpenProcess(PROCESS_TERMINATE Or PROCESS_QUERY_INFORMATION, True, ProcessID)
   If hProcess <> 0 Then
     GetExitCodeProcess hProcess, ExitCode
     If ExitCode <> 0 Then
       Rst = TerminateProcess(hProcess, ExitCode)
       CloseHandle hProcess
        If Rst = 0 Then

         EndProcess = False
        Else
         EndProcess = True
        End If
     Else
       EndProcess = False
     End If
   Else
     EndProcess = False
   End If
  
End Function


Public Function changeCmdStr(newCmdStr As String)
If newCmdStr <> "" Then
CmdStr = newCmdStr
Else
CmdStr = ""
End If
End Function
