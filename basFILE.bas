Attribute VB_Name = "basFILE"
'------------------------------------------------------------
'*******************************************
' READ READ READ READ READ READ READ
'*******************************************
'
' I am writting this here to let you know that
' when you first open this project, you will get
' an error saying it cannot find the VBDICE control.
'  The whole point of this program is to show you
' how a program can put some file onto a users
' machine and register it if needed.  So, if you
' trust me and my code, when you run it, a copy
' of VBDICE.OCX will be put on your machine and
' registered.  After that, because you are running
' this from the source, you will have to CLOSE
' VB WITHOUT saving.  Then reopen the project and
' all will work then because you have the control.
'  If you compile and give this to somebody [who
' has VB runtimes] when they run it, it will put
' the control on their computer and run perfectly.
'  Only because you are running this from the source
' are there problems.
'------------------------------------------------------------



Option Explicit
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long '
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  To run any file from its registered application,
'                to open the default browser to a web site, to
'                start up the default e-mail program and start
'                a mail message to a particular address, to run
'                a file and pass it arguments.
' Parameters:  LINK=Command;  Args=Optional Arguments
' Example:
'           ExecuteLink "http://lafever.iscool.net"
'           ExecuteLink "mailto:lafeverc@home.com"
'           ExecuteLink app.path & "\readme.doc"
'           ExecuteLink "REGSVR32.EXE","C:\WINNT\SYSTEM32\FILE.DLL"
' Date: August,23 2001 @ 11:37:35
'------------------------------------------------------------
Public Sub ExecuteLink(LINK As String, Optional Args As String = "")
    On Error Resume Next
    Dim lRet As Long
    If LINK <> "" Then
        lRet = ShellExecute(0, "open", LINK, Args, App.Path, SW_SHOWNORMAL)
        If lRet >= 0 And lRet <= 32 Then
            MsgBox "Error jumping to:" & LINK, 48, "Warning"
        End If
    End If
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Gets and returns the System Directory
' Returns:  String containing the System Directory.  Empty string on error.
' Date: June,22 1999 @ 06:13:35
'------------------------------------------------------------
Public Function GetSysDir() As String
    On Error GoTo ErrorGetSysDir
    Dim rSTR As String
    Dim rLEN As Long
    rSTR = String(255, 0)
    rLEN = GetSystemDirectory(rSTR, Len(rSTR))
    If rLEN < Len(rSTR) Then
        rSTR = Left(rSTR, rLEN)
        If Right(rSTR, 1) = "\" Then
            GetSysDir = Left(rSTR, Len(rSTR) - 1)
        Else
            GetSysDir = rSTR
        End If
    Else
        GetSysDir = ""
    End If
    Exit Function
ErrorGetSysDir:
    GetSysDir = ""
    MsgBox Err & ":Error in GetSysDir.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
'------------------------------------------------------------
' Author:  Clint LaFever - [lafeverc@saic.com]
' Purpose:  Extracts a file from the custom resource file
'                to the local hard drive.
' Parameters:  resID=ID of resource  :  resSECTION=Section of custom resource ie. CUSTOM
'                     fEXT=Extension for new file  :  fPATH=Destination path, default is App.Path
'                     fNAME=Name for new file, default is TEMP
' Returns:  Full path and file name of file created
' Example:  retSTR=GenFileFromRes(101,"CUSTOM","JPG",,"IMAGE")
' Date: December,17 1999 @ 10:50:58
'------------------------------------------------------------
Public Function GenFileFromRes(resID As Long, resSECTION As String, fEXT As String, Optional fPath As String = "", Optional fNAME As String = "temp", Optional FullName As String = "") As String
    On Error GoTo ErrorGenFileFromRes
    Dim resBYTE() As Byte
    If fPath = "" Then fPath = App.Path
    If fNAME = "" Then fNAME = "temp"
    '------------------------------------------------------------
    ' Get the file out of the resource file
    '------------------------------------------------------------
    resBYTE = LoadResData(resID, resSECTION)
    '------------------------------------------------------------
    ' Open destination
    '------------------------------------------------------------
    If FullName = "" Then
        Open fPath & "\" & fNAME & "." & fEXT For Binary Access Write As #1
    Else
        Open FullName For Binary Access Write As #1
    End If
    '------------------------------------------------------------
    ' Write it out
    '------------------------------------------------------------
    Put #1, , resBYTE
    '------------------------------------------------------------
    ' Close it
    '------------------------------------------------------------
    Close #1
    If FullName = "" Then
        GenFileFromRes = fPath & "\" & fNAME & "." & fEXT
    Else
        GenFileFromRes = FullName
    End If
    Exit Function
ErrorGenFileFromRes:
    GenFileFromRes = ""
    MsgBox Err & ":Error in GenFileFromRes.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
'------------------------------------------------------------
' Determines wheather or not a file already exists
' or not for the path/file name passed.
'------------------------------------------------------------
Function FileExists(filename As String) As Boolean
    On Error Resume Next
    Dim x As Long
    x = Len(Dir$(filename))
    If Err Or x = 0 Then FileExists = False Else FileExists = True
End Function
