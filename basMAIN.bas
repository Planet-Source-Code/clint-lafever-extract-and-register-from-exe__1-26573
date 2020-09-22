Attribute VB_Name = "basMAIN"
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
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  First checks to see if VBDICE.OCX is in the system
'                folder.  If not, extract it from the resource
'                file to the system folder and register it.  Once
'                VBDICE.OCX is in place [or already there], open
'                frmMAIN
' Date: August,23 2001 @ 11:30:25
'------------------------------------------------------------
Public Sub Main()
    On Error GoTo ErrorMain
    Dim sysDIR As String
    '------------------------------------------------------------
    ' Get SystemFolder Path
    '------------------------------------------------------------
    sysDIR = GetSysDir
    '------------------------------------------------------------
    ' Is VBDICE.OCX in the SystemFolder?
    '------------------------------------------------------------
    If FileExists(sysDIR & "\vbdice.ocx") = True Then 'Yes it is
        frmMAIN.Show
    Else 'No it is not
        '------------------------------------------------------------
        ' Extract a copy of VBDICE.OCX out of the Resource
        ' File to the System Folder
        '------------------------------------------------------------
        GenFileFromRes 101, "OCX", "OCX", , , sysDIR & "\vbdice.ocx"
        '------------------------------------------------------------
        ' Register it
        '------------------------------------------------------------
        ExecuteLink "REGSVR32.EXE", sysDIR & "\vbdice.ocx /s"
        frmMAIN.Show
    End If
    Exit Sub
ErrorMain:
    MsgBox Err & ":Error in Main.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
'------------------------------------------------------------
' Now I know there is great danger in code like
' this.  I now have just informed all you wanna
' be virus writters and malicious jerks how you
' can embed files into your programs and have the
' code extract and execute them.  I just know that
' I have found this code useful for legitimate
' reasons and thought others who are professionals
' would like it too.  I do suggest that anybody
' out there who likes downloading sample projects
' off the web like PSC, please get yourself a source
' code Project Scanner that checks for the calls
' that extract files from a resource file.  That
' way you can catch this activity before running
' source code you download.  I do have a project
' scanner at my web site http://lafever.iscool.net
' and it does check for this.  I know PSC has another
' one listed but I do not know if it checks for
' this type of call.  Anyhow, code safely.
'------------------------------------------------------------

