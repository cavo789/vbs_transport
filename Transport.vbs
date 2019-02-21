' ========================================================
'
' Author: Christophe Avonture
' Date  : February 2019
'
' > Transport "DEV.xlsm" to "PROD.xlsm" and update addins references
'
' Long description:
' -----------------
'
' This script will copy a file from a <source> to a <target> folder.
'
' This utility was mainly developed to copy an Excel file from a (development)
' folder to a (production) folder after modifying it to replace the reference to
' an addin with another.
'
' The development Excel file ("dev.xlsm") can refer to an addin on the hard disk ("dev.xlam")
' while, once in production, the addin must be the production one ("prod.xlam").
'
' Usage: `Transport.vbs -s -t [-o=xxx] [-n=xxx] [/?]`
'
' -s=xxx    The <s>ource file i.e. the file that should be copied
' -t=xxx    The <t>arget file i.e. the location where the file should be copied
' -o=xxx    The <o>ld addin filename i.e. the addin that should be removed [optional parameter]
' -n=xxx    The <n>ew addin filename to add in references [optional parameter]
' -o2=xxx   The <o2>ld addin filename (if there is a second addin to remove) [optional parameter]
' -n2=xxx   The <n2>ew addin filename (if there is a second addin to add) [optional parameter]
'
' /?        Show this help screen (or /help)
' /force    Force the script to run, don't ask for confirmations
' /silent   Don't display messages, silent mode.
' /readonly Once copied, set the Read-Only file's attributes
' /hidden   Once copied, set the Hidden file's attributes
' /open     Once copied, open the folder with Windows Explorer
'
' Examples:
' --------
'
' - Simply copy from <s>ource to <t>arget, no changes
'
' ```
' cscript transport.vbs -s="C:\Christophe\App.xlsm" -t="L:\Prod\App.xlsm"
' ```
'
' - Copy from <s>ource (=dev) to <t>arget (=prod) but also remove <o>ld addin and add <n>ew one
'   So "L:\Prod\App.xlsm" will then be a copy of "C:\Christophe\App.xlsm" but using the
'   "L:\Prod\addin.xlam" production file and not the addin from development
'
' ```
' cscript transport.vbs -s="C:\Christophe\App.xlsm" -t="L:\Prod\App.xlsm" -o="C:\Christophe\addin.xlam" -n="L:\Prod\addin.xlam"
' ```
' ========================================================

Option Explicit

' Debug mode enabled or not.
Const DebugMode = False

Class clsParameters

    ' Fullname of the source file (file to copy)
    Dim sSourceFileName

    ' Fullname of the target file (copied filename)
    Dim sTargetFileName

    ' Fullname of the old addin (if there is one)
    Dim sOldAddIn
    Dim sOldAddIn2

    ' Fullname of the new addin (if there is one)
    Dim sNewAddIn
    Dim sNewAddIn2

    ' When bForceRun is equal to 1, the user won't be prompted for,
    ' for instance, confirmations.
    Dim bForceRun

    ' When bSilentMode is equal to 1, the user doesn't want to see
    ' messages (except errors)
    Dim bSilentMode

    ' When bReadOnly is equal to 1, once copied to his target folder,
    ' file's attributes will be set to "+R" i.e. current attributes + read-only
    Dim bReadOnly

    ' When bHidden is equal to 1, once copied to his target folder,
    ' file's attributes will be set to "+R" i.e. current attributes + hidden
    Dim bHidden

    ' Once copied, open the folder with Windows Explorer
    Dim bOpenFolder

    ' Working variables/objects
    Dim objFSO
    Dim bVerbose

    ' --------------------------------------------------------
    ' Class initialization
    ' --------------------------------------------------------
    Private Sub Class_Initialize()

        Set objFSO = CreateObject("Scripting.FileSystemObject")

        bVerbose = False
        bForceRun = False
        bSilentMode = False
        bReadOnly = False
        bHidden = False
        bOpenFolder = False

        sSourceFileName = ""
        sTargetFileName = ""
        sOldAddIn = ""
        sOldAddIn2 = ""
        sNewAddIn = ""
        sNewAddIn2 = ""

    End Sub

    ' --------------------------------------------------------
    ' Before leaving the class, release objects
    ' --------------------------------------------------------
    Private Sub Class_Terminate()
        Set objFSO = Nothing
    End Sub

    ' --------------------------------------------------------
    ' Allow (true) or Disallow (false) this class to echoed
    ' information's messages. Errors will always be echoed.
    ' --------------------------------------------------------
    Public Property Let verbose(bYesNo)
        bVerbose = bYesNo
    End Property

    ' --------------------------------------------------------
    ' Force i.e. don't ask for, f.i., confirmations
    ' --------------------------------------------------------
    Public Property Get Force()
        Force = bForceRun
    End Property

    ' --------------------------------------------------------
    ' SilentMode will (dis)allow messages (errors will always
    ' be echoed)
    ' --------------------------------------------------------
    Public Property Get SilentMode()
        SilentMode = bSilentMode
    End Property

    ' --------------------------------------------------------
    ' ReadOnly, if set, means that the file, once copied to his
    ' target destination, will be set as a Read-Only file
    ' --------------------------------------------------------
    Public Property Get ReadOnly()
        ReadOnly = bReadOnly
    End Property

    ' --------------------------------------------------------
    ' Hidden, if set, means that the file, once copied to his
    ' target destination, will be set as an hidden file
    ' --------------------------------------------------------
    Public Property Get Hidden()
        Hidden = bHidden
    End Property

    ' --------------------------------------------------------
    ' Once copied, if True, open the folder with Windows Explorer
    ' --------------------------------------------------------
    Public Property Get OpenFolder()
        OpenFolder = bOpenFolder
    End Property

    ' ------------------------------------------------------
    ' Get the current folder
    ' ------------------------------------------------------
    Public Property Get CurrentFolder

        Dim objFile

        Set objFile = objFSO.GetFile(wScript.ScriptFullName)

        CurrentFolder = objFSO.GetParentFolderName(objFile)

        ' Be sure to have the final slash
        If Not (Right(CurrentFolder, 1) = "\") THen
            CurrentFolder = CurrentFolder & "\"
        End If

        Set objFile = Nothing

    End Property

    ' ------------------------------------------------------
    ' Initialize the Source variable with the full name of
    ' the file to copy
    '
    ' For instance "C:\Christophe\Repository\App\Interface.xlsm"
    ' ------------------------------------------------------
    Public Property Let SourceFile(sFileName)
        sSourceFileName = sFileName
    End Property

    ' ------------------------------------------------------
    ' Initialize the Target variable with the full name of
    ' the file to copy
    '
    ' For instance "I:\Production\App\Interface.xlsm"
    ' ------------------------------------------------------
    Public Property Let TargetFile(sFileName)
        sTargetFileName = sFileName
    End Property

    ' ------------------------------------------------------
    ' Initialize the old addin variable i.e. the addin that
    ' should be removed from the file during the transport
    '
    ' For instance "C:\Christophe\Repository\App\DEV_AddIn.xlam"
    ' ------------------------------------------------------
    Public Property Let OldAddIn(sFileName)
        sOldAddIn = sFileName
    End Property

    Public Property Let OldAddIn2(sFileName)
        sOldAddIn2 = sFileName
    End Property

    ' ------------------------------------------------------
    ' Initialize the new addin variable i.e. the addin that
    ' should be added to the file during the transport
    '
    ' For instance "I:\Production\App\PROD_AddIn.xlam"
    ' ------------------------------------------------------
    Public Property Let NewAddIn(sFileName)
        sNewAddIn = sFileName
    End Property

    Public Property Let NewAddIn2(sFileName)
        sNewAddIn2 = sFileName
    End Property

    ' ------------------------------------------------------
    ' Folder where the source file (probably from DEV) is
    ' located (aka "cDevFolder")
    '
    ' Return, for instance, "C:\Christophe\Repository\App\"
    ' ------------------------------------------------------
    Public Property Get SourceFolder()

        SourceFolder = objFSO.GetParentFolderName(sSourceFileName)

        ' Be sure to have the final slash
        If Not (Right(SourceFolder, 1) = "\") THen
            SourceFolder = SourceFolder & "\"
        End If

    End Property

    ' ------------------------------------------------------
    ' Fullname of the file to copy
    '
    ' Return, for instance, "C:\Christophe\Repository\App\Interface.xlsm"
    ' ------------------------------------------------------
    Public Property Get SourceFileFullName
        SourceFileFullName = sSourceFileName
    End Property

    ' ------------------------------------------------------
    ' Filename (no path)  of the file to transport (aka "cFile")
    '
    ' Return, for instance, "Interface.xlsm"
    ' ------------------------------------------------------
    Public Property Get SourceFileName
        SourceFileName = objFSO.GetFileName(sSourceFileName)
    End Property

    ' ------------------------------------------------------
    ' Folder where the file should be copied (aka "cTargetFolder")
    '
    ' Return, for instance, "I:\Production\App\"
    ' ------------------------------------------------------
    Public Property Get TargetFolder()

        TargetFolder = objFSO.GetParentFolderName(sTargetFileName)

        ' Be sure to have the final slash
        If Not (Right(TargetFolder, 1) = "\") THen
            TargetFolder = TargetFolder & "\"
        End If

    End Property

    ' ------------------------------------------------------
    ' Fullname of the file once copied (on target folder)
    '
    ' Return, for instance, "I:\Production\App\Interface.xlsm"
    ' ------------------------------------------------------
    Public Property Get TargetFileFullName
        TargetFileFullName = sTargetFileName
    End Property

    ' ------------------------------------------------------
    ' Filename (no path) of the file once copied (on target folder)
    '
    ' Return, for instance, "Interface.xlsm"
    ' ------------------------------------------------------
    Public Property Get TargetFileName
        TargetFileName = objFSO.GetFileName(sTargetFileName)
    End Property

    ' ------------------------------------------------------
    ' Fullname of the old addin that should be removed from
    ' the file during the transport
    '
    ' Return, for instance, "C:\Christophe\Repository\App\DEV_AddIn.xlam"
    ' ------------------------------------------------------
    Public Property Get OldAddInFullName
        OldAddInFullName = sOldAddIn
    End Property

    Public Property Get OldAddIn2FullName
        OldAddIn2FullName = sOldAddIn2
    End Property

    ' ------------------------------------------------------
    ' Basename (no extension) of the old addin
    '
    ' Return, for instance, "DEV_AddIn"
    ' ------------------------------------------------------
    Public Property Get OldAddInBaseName
        OldAddInBaseName = objFSO.GetBaseName(sOldAddIn)
    End Property

    Public Property Get OldAddIn2BaseName
        OldAddIn2BaseName = objFSO.GetBaseName(sOldAddIn2)
    End Property

    ' ------------------------------------------------------
    ' Fullname of the new addin that should be added to the file
    ' during the transport
    '
    ' Return, for instance, "I:\Production\App\PROD_AddIn.xlam"
    ' ------------------------------------------------------
    Public Property Get NewAddInFullName
        NewAddInFullName = sNewAddIn
    End Property

    Public Property Get NewAddIn2FullName
        NewAddIn2FullName = sNewAddIn2
    End Property

    ' ------------------------------------------------------
    ' Basename (no extension) of the new addin
    '
    ' Return, for instance, "PROD_AddIn"
    ' ------------------------------------------------------
    Public Property Get NewAddInBaseName
        NewAddInBaseName = objFSO.GetBaseName(sNewAddIn)
    End Property

    Public Property Get NewAddIn2BaseName
        NewAddIn2BaseName = objFSO.GetBaseName(sNewAddIn2)
    End Property

    ' --------------------------------------------------------
    ' Small wScript.echo shortcut
    ' --------------------------------------------------------
    Private Sub Echo(sLine)
        wScript.echo " " & sLine
    End Sub

    ' --------------------------------------------------------
    '
    ' Process command line parameters / options
    '
    '   Parameters starts with a "-" (like in "-s=")
    '   Options starts with a "/" (like in "/?" or "/help")
    '
    ' --------------------------------------------------------
    Public Sub Read()

        Dim wCount, I
        Dim sArgument

        ' No arguments, ouch, Houston, something is wrong
        If (wScript.Arguments.Count = 0) Then
            Call ShowHelp()
            Exit Sub
        End If

        wCount = wScript.Arguments.Count - 1

        ' Process arguments one by one
        For I = 0 To wCount

            ' Get the argument
            sArgument = Trim(wScript.Arguments(I))

            If (Left(sArgument, 3) = "-s=") Then
                ' -s for the source file
                cParameters.SourceFile = Right(sArgument, Len(sArgument) - 3)
            ElseIf (Left(sArgument, 3) = "-t=") Then
                ' -t for the target file
                cParameters.TargetFile = Right(sArgument, Len(sArgument) - 3)
            ElseIf (Left(sArgument, 3) = "-o=") Then
                ' -o for the old Addin filename
                cParameters.OldAddIn = Right(sArgument, Len(sArgument) - 3)
            ElseIf (Left(sArgument, 4) = "-o2=") Then
                ' -o2 for the second old Addin filename
                cParameters.OldAddIn2 = Right(sArgument, Len(sArgument) - 4)
            ElseIf (Left(sArgument, 3) = "-n=") Then
                ' -n for the new Addin filename
                cParameters.NewAddin = Right(sArgument, Len(sArgument) - 3)
            ElseIf (Left(sArgument, 4) = "-n2=") Then
                ' -n2 for the second new Addin filename
                cParameters.NewAddin2 = Right(sArgument, Len(sArgument) - 4)
            Else
                If (LCase(sArgument) = "/forcerun") or (LCase(sArgument) = "/force") Then
                    bForceRun = true
                ElseIf (LCase(sArgument) = "/silentmode") or (LCase(sArgument) = "/silent") Then
                    bSilentMode = True
                ElseIf (LCase(sArgument) = "/readonly") Then
                    bReadOnly = True
                ElseIf (LCase(sArgument) = "/hidden") Then
                    bHidden = True
                ElseIf (LCase(sArgument) = "/open") Then
                    bOpenFolder = True
                ElseIf (sArgument = "/?") Then
                    ShowHelp()
                ElseIf (LCase(sArgument) = "/help") Then
                    ShowHelp()
                Else
                    If bVerbose Then
                        wScript.echo "Debug - UNSUPPORTED PARAMETER " & I & " is " & wScript.Arguments(I)
                    End If
                End If
            End If

        Next

        Call Debug()

    End Sub

    ' --------------------------------------------------------
    '
    ' Helper, Iif function
    '
    ' --------------------------------------------------------
    Private Function IIf(expr, truepart, falsepart)
        IIf = falsepart
        If expr Then IIf = truepart
    End Function

    ' ------------------------------------------------------
    ' Small debug code for displaying values
    ' ------------------------------------------------------
    Private Sub Debug()

        ' Only when the DebugMode constant is set
        ' Don't check the SilentMode since this is only during
        ' development.
        If DebugMode Then
            wScript.echo " --- DebugMode SET - List of Parameters ---"
            wScript.echo " Source file"
            wScript.echo "      fullname " & SourceFileFullName
            'wScript.echo "      filename " & SourceFileName
            'wScript.echo "      folder   " & SourceFolder
            'wScript.echo ""
            wScript.echo " Target file"
            wScript.echo "      fullname " & TargetFileFullName
            'wScript.echo "      filename " & TargetFileName
            'wScript.echo "      folder   " & TargetFolder
            'wScript.echo ""
            wScript.echo " Old AddIn (to remove from the references)"
            wScript.echo "      fullname " & OldAddInFullName
            wScript.echo "      fullname2 " & OldAddIn2FullName
            'wScript.echo "      basename " & OldAddInBaseName
            'wScript.echo ""
            wScript.echo " New AddIn (to add to the references)"
            wScript.echo "      fullname " & NewAddInFullName
            wScript.echo "      fullname2 " & NewAddIn2FullName
            'wScript.echo "      basename " & NewAddInBaseName
            wScript.echo ""
            wScript.echo " Options"
            wScript.echo "      force " & Iif(bForceRun, "True", "False")
            wScript.echo "      hidden " & Iif(bHidden, "True", "False")
            wScript.echo "      open " & Iif(bOpenFolder, "True", "False")
            wScript.echo "      readonly " & Iif(bReadOnly, "True", "False")
            wScript.echo "      silentmode " & Iif(bSilentMode, "True", "False")
            wScript.echo " ----------------"
        End If

    End Sub

    ' --------------------------------------------------------
    '
    ' Display how to use the script and the list of parameters
    ' Quit the script
    '
    ' --------------------------------------------------------
    Private Sub showHelp

        Echo "This script will copy a file from a <source> to a <target> folder."
        Echo ""
        Echo "This utility was mainly developed to copy an Excel file from a (development) "
        Echo "folder to a (production) folder after modifying it to replace the reference to "
        Echo "an addin with another."
        Echo ""
        Echo "The development Excel file (""dev.xlsm"") can refer to an addin on the hard disk (""dev.xlam"") "
        Echo "while, once in production, the addin must be the production one (""prod.xlam"")."
        Echo ""
        Echo "Usage: Transport.vbs -s -t [-o=xxx] [-n=xxx] [/?]"
        Echo ""
        Echo "-s=xxx    The <s>ource file i.e. the file that should be copied"
        Echo "-t=xxx    The <t>arget file i.e. the location where the file should be copied"
        Echo "-o=xxx    The <o>ld addin filename i.e. the addin that should be removed [optional parameter]"
        Echo "-n=xxx    The <n>ew addin filename to add in references [optional parameter]"
        Echo "-o2=xxx   The <o2>ld addin filename (if there is a second addin to remove) [optional parameter]"
        Echo "-n2=xxx   The <n2>ew addin filename (if there is a second addin to add) [optional parameter]"
        Echo ""
        Echo "/?        Show this help screen (or /help)"
        Echo "/force    Force the script to run, don't ask for confirmations"
        Echo "/silent   Don't display messages, silent mode."
        Echo "/readonly Once copied, set the Read-Only file's attributes"
        Echo "/hidden   Once copied, set the Hidden file's attributes"
        Echo "/open     Once copied, open the folder with Windows Explorer"
        Echo ""
        Echo "Examples: "
        Echo ""
        Echo "- Simply copy from <s>ource to <t>arget, no changes"
        Echo ""
        Echo "    cscript " & Wscript.ScriptName & " -s=""C:\Christophe\App.xlsm"" -t=""L:\Prod\App.xlsm"""
        Echo ""
        Echo "- Copy from <s>ource (=dev) to <t>arget (=prod) but also remove <o>ld addin and add <n>ew one"
        Echo "  So ""L:\Prod\App.xlsm"" will then be a copy of ""C:\Christophe\App.xlsm"" but " & _
                    "using the ""L:\Prod\addin.xlam"" production file and not the addin from development"
        Echo ""
        Echo "    cscript " & Wscript.ScriptName & " -s=""C:\Christophe\App.xlsm"" -t=""L:\Prod\App.xlsm"" -o=""C:\Christophe\addin.xlam"" -n=""L:\Prod\addin.xlam"""

        ' And quit
        wScript.Quit 0

    End Sub

End Class

' ========================================================
'
' Author : Christophe Avonture
' Date	: November / December 2017
'
' MS Excel helper
'
' This class provide functionnalities to facilitate automation of
' MS Excel
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSExcel.md
'
' Changes
' =======
'
' March 2018 - Improve OpenCSV method
'
' ========================================================

Class clsMSExcel

    Private oApplication
    Private sFileName
    Private bVerbose, bEnableEvents, bDisplayAlerts

    Private bAppHasBeenStarted

    ' --------------------------------------------------------
    ' Allow (true) or Disallow (false) this class to echoed
    ' information's messages. Errors will always be echoed.
    ' --------------------------------------------------------
    Public Property Let verbose(bYesNo)
        bVerbose = bYesNo
    End Property

    Public Property Let EnableEvents(bYesNo)
        bEnableEvents = bYesNo

        If Not (oApplication Is Nothing) Then
            oApplication.EnableEvents = bYesNo
        End If
    End Property

    Public Property Let DisplayAlerts(bYesNo)
        bDisplayAlerts = bYesNo

        If Not (oApplication Is Nothing) Then
            oApplication.DisplayAlerts = bYesNo
        End If

    End Property

    Public Property Let FileName(ByVal sName)
        sFileName = sName
    End Property

    Public Property Get FileName
        FileName = sFileName
    End Property

    ' Make oApplication accessible
    Public Property Get app
        Set app = oApplication
    End Property

    Private Sub Class_Initialize()
        bVerbose = False
        bAppHasBeenStarted = False
        bEnableEvents = False
        bDisplayAlerts = False
        Set oApplication = Nothing
    End Sub

    Private Sub Class_Terminate()
        Set oApplication = Nothing
    End Sub

    ' --------------------------------------------------------
    ' Initialize the oApplication object variable : get a pointer
    ' to the current Excel.exe app if already in memory or start
    ' a new instance.
    '
    ' If a new instance has been started, initialize the variable
    ' bAppHasBeenStarted to True so the rest of the script knows
    ' that Excel should then be closed by the script.
    ' --------------------------------------------------------
    Public Function Instantiate()

        If (oApplication Is Nothing) Then

            On error Resume Next

            Set oApplication = GetObject(,"Excel.Application")

            If (Err.number <> 0) or (oApplication Is Nothing) Then
                Set oApplication = CreateObject("Excel.Application")
                ' Remember that Excel has been started by
                ' this script ==> should be released
                bAppHasBeenStarted = True
            End If

            oApplication.EnableEvents = bEnableEvents
            oApplication.DisplayAlerts = bDisplayAlerts

            Err.clear

            On error Goto 0

        End If

        ' Return True if the application was created right
        ' now
        Instantiate = bAppHasBeenStarted

    End Function

    Public Sub Quit()
        If not (oApplication Is Nothing) Then
            oApplication.Quit
        End If
    End Sub

    ' --------------------------------------------------------
    ' Open a standard Excel file and allow to specify if the
    ' file should be opened in a read-only mode or not
    ' --------------------------------------------------------
    Public Sub Open(sFileName, bReadOnly)

        If not (oApplication Is nothing) Then

            If bVerbose Then
                wScript.echo "Open " & sFileName & _
                    " (clsMSExcel::Open)"
            End If

            ' False = UpdateLinks
            oApplication.Workbooks.Open sFileName, False, _
                bReadOnly

        End If

    End sub

    ' --------------------------------------------------------
    ' Close the active workbook
    ' --------------------------------------------------------
    Public Sub CloseFile(sFileName)

        Dim wb
        Dim I
        Dim objFSO
        Dim sBaseName

        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            If (sFileName = "") Then
                If Not (oApplication.ActiveWorkbook Is Nothing) Then
                    sFileName = oApplication.ActiveWorkbook.FullName
                End If
            End If

            If (sFileName <> "") Then

                If bVerbose Then
                    wScript.echo "Close " & sFileName & _
                        " (clsMSExcel::CloseFile)"
                End If

                ' Only the basename and not the full path
                sBaseName = objFSO.GetFileName(sFileName)

                On Error Resume Next
                Set wb = oApplication.Workbooks(sBaseName)
                If Not (err.number = 0) Then
                    ' Not found, workbook not loaded
                    Set wb = Nothing
                Else
                    If bVerbose Then
                        wScript.echo "	Closing " & sBaseName & _
                            " (clsMSExcel::CloseFile)"
                    End If
                    ' Close without saving
                    wb.Close False
                End If

                On Error Goto 0

            End If

            Set objFSO = Nothing

        End If

    End Sub

    ' --------------------------------------------------------
    ' Save the active workbook on disk
    ' --------------------------------------------------------
    Public Sub SaveFile(sFileName)

        Dim wb, objFSO

        ' If Excel isn't loaded or has no active workbook, there
        ' is thus nothing to save.
        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            On Error Resume Next

            Set wb = oApplication.Workbooks(objFSO.GetFileName(sFileName))

            If (Err.Number <> 0) Then
                Err.clear
                ' Perhaps the file isn't a .xlsx (.xlsm) file but an Addin$
                ' Try with the AddIns2 collection
                Set wb = oApplication.AddIns2(objFSO.GetFileName(sFileName))
            End If

            On Error Goto 0

            If Not (wb is Nothing) Then

                If (bVerbose) Then
                    wScript.echo "Save file " & sFileName & _
                        " (clsMSExcel::SaveFile)"
                End If

                If (wb.FullName = sFileName) Then
                    wb.Save
                Else
                    ' Don't specify extension because if we've opened
                    ' a .xlsm file and save the file elsewhere, we need
                    ' to keep the same extension
                    wb.SaveAs sFileName
                End If
            End If

            Set wb = Nothing
            Set objFSO = Nothing

        End If

    End Sub

    ' --------------------------------------------------------
    ' Check if a specific file is already opened in Excel
    ' This function will return True if the file is already loaded.
    ' --------------------------------------------------------
    Public Function IsLoaded(sFileName)

        Dim bLoaded, bShouldClose
        Dim bCheckAddins2
        Dim I, J

        bLoaded = false

        If (oApplication Is Nothing) Then
            bShouldClose = Instantiate()
        End If

        On Error Resume Next

        If bVerbose Then
            wScript.echo "Check if " & sFileName & _
                " is already loaded (clsMSExcel::IsLoaded)"
        End If

        If (Right(sFileName, 5) = ".xlam") Then

            ' The AddIns2 collection only exists since MSOffice
            ' 2014 (version 14)
            On Error Resume Next
            J = oApplication.AddIns2.Count
            bCheckAddins2 = (Err.Number = 0)
            On Error Goto 0

            If (bCheckAddins2) then

                J = oApplication.AddIns2.Count

                If J > 0 Then
                    For I = 1 To J
                        If (StrComp(oApplication.AddIns2(I).FullName,sFileName,1)=0) Then
                            bLoaded = True
                            Exit For
                        End If
                    Next ' For I = 1 To J
                End If

            End If ' If (oApplication.version >=14) then

        Else ' If (Right(sFileName, 5) = ".xlam") Then

            ' It's a .xls, .xlsm, ... file, not an AddIn
            J = oApplication.Workbooks.Count

            If J > 0 Then
                For I = 1 To J
                    If (StrComp(oApplication.Workbooks(I).FullName,sFileName,1)=0) Then
                        bLoaded = True
                        Exit For
                    End If
                Next ' For I = 1 To J
            End If ' If J > 0 Then

        End If ' If (Right(sFileName, 5) = ".xlam") Then

        On Error Goto 0

        ' Quit Excel if it was started here, in this script
        If bShouldClose then
            oApplication.Quit
            Set oApplication = Nothing
        End If

        IsLoaded = bLoaded

    End Function

    ' --------------------------------------------------------
    ' Remove a specific reference. For instance, remove a linked
    ' .xlam library from the list of references used by the file
    '
    ' sAddin   The name of the reference; without the extension
    '          For instance "MyLibrary" (and not MyLibrary.xlam)
    '
    ' Note: Make sure Events are disabled before calling this subroutine
    '       ==> .EnableEvents = False
    ' --------------------------------------------------------
    Public Sub References_Remove(sFileName, sAddin)

        Dim wb, ref
        Dim objFSO
        Dim sRefFullName, sBaseName
        Dim bEvents

        If Not (oApplication Is Nothing) Then

            bEvents = oApplication.EnableEvents
            oApplication.EnableEvents = false

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            ' sAddin should be a relative filename -
            ' Without the extension !
            If Instr(sAddin, "\")>0 Then
                sAddin = objFSO.GetBaseName(sAddin)
            End If

            If (StrComp(Right(sAddIn, 5), ".xlam", 1) = 0) Then
                sAddIn = Left(sAddIn, Len(sAddIn) - 5)
            End If

            'If bVerbose Then
            '    wScript.echo "Try to remove " & sAddin & " from the references"
            'End If

            ' IF STILL, EXCEL RUN A MACRO, IT'S MORE PROBABLY DUE TO A RIBBON
            ' PRESENT IN THE FILE AND AN "ONLOAD" SUBROUTINE.
            ' In this case, update the OnLoad code and check if
            ' Application.EnableEvents is equal to False and in this case,
            ' don't run your onLoad code; exit your subroutine.
            ' Add something like here below in the top of your subroutine:
            '
            '       If Not (Application.EnableEvents) Then
            '           Exit Sub
            '       End If
            '
            ' ALSO MAKE SURE TO NOT START EXCEL VISIBLE: THE RIBON IS LOADED
            ' IN THAT CASE

            Set wb = oApplication.Workbooks(objFSO.GetFileName(sFileName))

            If Not (wb Is Nothing) Then

                With wb

                    For Each ref In .VBProject.References

                        'If bVerbose Then
                        '    wScript.echo "   Found " & ref.Name & _
                        '        " (clsMSExcel::References_Remove)"
                        'End If

                        If (ref.Name = sAddIn) Then

                            If bVerbose Then
                                wScript.echo "      Remove " & ref.Name & _
                                    " addin (clsMSExcel::References_Remove)"
                            End If

                            ' Get the fullpath of the reference
                            sRefFullName = ref.FullPath

                            .VBProject.References.Remove ref

                            ' IF A PROBLEM OCCURS ON THE LINE BELOW, EVEN WHEN THE
                            ' .EnableEvents PROPERTY IS SET TO False, TRY TO PUT IN COMMENT
                            ' THE LINE "Option Explicit" THAT IS PRESENT IN THE MODULE WHERE
                            ' THE ERROR IS COMING. THE ERROR CAN COMES FROM AN UNDECLARED ADDIN
                            ' DUE TO THE "".VBProject.References.Remove ref" COMMAND
                            '
                            ' SO EDIT THE ORIGINAL EXCEL FILE, LOCATE THE MODULE / CLASS AND
                            ' COMMENT THE "Option Explicit" LINE
                            .Save

                            ' --------------------------------------
                            ' Once unloaded, close the .xlam file
                            ' This should be made by closing the
                            ' filename (addin.xlam) and not just
                            ' the name (addin) or the fullname
                            ' So get the filename
                            sBaseName = objFSO.GetFileName(sRefFullName)

                            If bVerbose Then
                                wScript.echo "Unload " & sBaseName & _
                                " (clsMSExcel::References_Remove)"
                            End If

                            Call oApplication.Workbooks(sBaseName).Close

                            Exit For

                        End If

                    Next
                End With

            End If

            Set wb = Nothing
            Set objFSO = Nothing

            oApplication.EnableEvents = bEnableEvents

        End If

    End Sub

    ' --------------------------------------------------------
    ' Add an addin to the list of references.
    ' For instance, add MyAddin.xlam to the MyInterface.xlsm
    '
    ' sAddin  The full filename to the addin to add as reference
    ' --------------------------------------------------------
    Public Sub References_AddFromFile(sFileName, sAddinFile)

        Dim bReturn
        Dim wb, ref

        bReturn = true

        Dim objFSO

        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            Set wb = oApplication.Workbooks(objFSO.GetFileName(sFileName))

            If Not (wb Is Nothing) Then

                If bVerbose Then
                    wScript.echo "Add a reference to " & sAddInFile
                End If

                wb.VBProject.References.AddFromFile sAddInFile

            End If

        End If

    End Sub

End Class

Class clsTransport

    ' Working variables/objects
    Dim objFSO
    Dim bVerbose
    Dim cParameters
    Dim cMSExcel

    ' List of constants (the const directive isn't supported in a class)
    ' Values are then defined in the Class_Initialize subroutine
    Private cErrorExcelIsRunning        ' -1
    Private cErrorNoParams              ' -2
    Private cErrorNoSourceFileName      ' -3
    Private cErrorNoTargetFileName      ' -4
    Private cErrorSourceFileNotFound    ' -5
    Private cErrorOldAddInFileNotFound  ' -6
    Private cErrorNewAddInFileNotFound  ' -7
    Private cErrorOldAddIn2FileNotFound ' -8
    Private cErrorNewAddIn2FileNotFound ' -9
    Private cErrorCopyFileInUse         ' -25

    ' --------------------------------------------------------
    ' Class initialization
    ' --------------------------------------------------------
    Private Sub Class_Initialize()

        Set objFSO = CreateObject("Scripting.FileSystemObject")
        bVerbose = False

        ' Define errors constants
        cErrorExcelIsRunning = -1
        cErrorNoParams = -2
        cErrorNoSourceFileName = -3
        cErrorNoTargetFileName = -4
        cErrorSourceFileNotFound = -5
        cErrorOldAddInFileNotFound = -6
        cErrorNewAddInFileNotFound = -7
        cErrorOldAddIn2FileNotFound = -8
        cErrorNewAddIn2FileNotFound = -9
        cErrorCopyFileInUse = -25

        Set cMSExcel = Nothing

    End Sub

    ' --------------------------------------------------------
    ' Before leaving the class, release objects
    ' --------------------------------------------------------
    Private Sub Class_Terminate()

        Set objFSO = Nothing
        Set cParameters = Nothing

        ' Release Excel
        If Not (cMSExcel Is Nothing) Then

            ' Restore important settings
            cMSExcel.App.EnableEvents = True
            cMSExcel.App.DisplayAlerts = True

            cMSExcel.Quit
            Set cMSExcel = Nothing

        End If

    End Sub

    ' --------------------------------------------------------
    ' Allow (true) or Disallow (false) this class to echoed
    ' information's messages. Errors will always be echoed.
    ' --------------------------------------------------------
    Public Property Let verbose(bYesNo)
        bVerbose = bYesNo
    End Property

    ' --------------------------------------------------------
    ' Initialize the Parameters object with every paths and
    ' filenames
    ' --------------------------------------------------------
    Public Property Let Parameters(cParams)
        Set cParameters = cParams
    End Property

    ' ------------------------------------------------------
    ' Return the Windows Temporary Folder with "\Transport"
    ' has subfolder i.e. the staging folder where files will
    ' be copied temporarily so changes can be done there
    ' before putting the file in his final destination (PROD)
    '
    ' Return, f.i., "C:\Users\Christophe\AppData\Local\Temp\Transport\"
    ' ------------------------------------------------------
    Public Property Get StagingFolder()

        ' 2 = Windows Temp folder
        StagingFolder = objFSO.GetSpecialFolder(2)

        ' Be sure to have the final slash
        If Not (Right(StagingFolder, 1) = "\") THen
            StagingFolder = StagingFolder & "\"
        End If

        ' Add a subfolder and create the folder if not present
        StagingFolder = StagingFolder & "Transport\"

        If Not (objFSO.FolderExists(StagingFolder)) Then
            Call objFSO.CreateFolder(StagingFolder)
        End If

    End Property

    ' ------------------------------------------------------
    ' Helper, check the existence of a file
    ' ------------------------------------------------------
    Private Function FileExists(sFile)

        If (Trim(sFile) = "") Then
            FileExists = False
        Else
            FileExists = objFSO.FileExists(sFile)
        End If

    End Function

    ' ------------------------------------------------------
    ' Make a few controls and if one of fails, stop
    ' ------------------------------------------------------
    Private Sub Check()

        If (getCountInstancesOfExcel() > 0) Then
            wScript.echo "ERROR - Excel is running"
            wScript.echo "In order to avoid problems with auto-running code " & _
                "(like in ribbons); all instances of Excel should be first terminated"
            wScript.echo "Excel will be started in an automated way where events " & _
                "will be disabled"
            wScript.Quit cErrorExcelIsRunning
        End If

        If Not(IsObject(cParameters)) Then
            wScript.echo "ERROR - Parameters not initialized in clsTransport"
            wScript.Quit cErrorNoParams
        End If

        If (cParameters.SourceFileFullName = "") Then
            wScript.echo "ERROR - You must provide the source filename"
            wScript.echo "Please use -s option in the command line parameters"
            wScript.Quit cErrorNoSourceFileName
        End If

        If (cParameters.TargetFileFullName = "") Then
            wScript.echo "ERROR - You must provide the target filename"
            wScript.echo "Please use -t options in the command line parameters"
            wScript.Quit cErrorNoTargetFileName
        End If

        If Not (FileExists(cParameters.SourceFileFullName)) Then
            wScript.echo "ERROR - The specified source file doesn't exists"
            wScript.echo "File " & cParameters.SourceFileFullName & " not found"
            wScript.echo "Please check file mentionned for the -s option"
            wScript.Quit cErrorSourceFileNotFound
        End If

        If ((cParameters.OldAddInFullName <> "") And _
            Not (FileExists(cParameters.OldAddInFullName))) Then
            wScript.echo "ERROR - The old addin file doesn't exists"
            wScript.echo "File " & cParameters.OldAddInFullName & " not found"
            wScript.echo "Please check file mentionned for the -o option"
            wScript.Quit cErrorOldAddInFileNotFound
        End If

        If ((cParameters.NewAddInFullName <> "") And _
            Not (FileExists(cParameters.NewAddInFullName))) Then
            wScript.echo "ERROR - The new addin file doesn't exists"
            wScript.echo "File " & cParameters.NewAddInFullName & " not found"
            wScript.echo "Please check file mentionned for the -n option"
            wScript.Quit cErrorNewAddInFileNotFound
        End If

        If ((cParameters.OldAddIn2FullName <> "") And _
            Not (FileExists(cParameters.OldAddIn2FullName))) Then
            wScript.echo "ERROR - The second old addin file doesn't exists"
            wScript.echo "File " & cParameters.OldAddIn2FullName & " not found"
            wScript.echo "Please check file mentionned for the -o2 option"
            wScript.Quit cErrorOldAddIn2FileNotFound
        End If

        If ((cParameters.NewAddIn2FullName <> "") And _
            Not (FileExists(cParameters.NewAddIn2FullName))) Then
            wScript.echo "ERROR - The second new addin file doesn't exists"
            wScript.echo "File " & cParameters.NewAddIn2FullName & " not found"
            wScript.echo "Please check file mentionned for the -n2 option"
            wScript.Quit cErrorNewAddIn2FileNotFound
        End If

    End Sub

    ' ------------------------------------------------------
    ' Get the number of instances of already running Excel
    ' ------------------------------------------------------
    Private Function getCountInstancesOfExcel()

        Dim objWMIService, objProcess, colProcess
        Dim wCount

        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = " & "'EXCEL.EXE'")

        wCount = colProcess.Count

        'If we need to kill them...
        'If (wCount > 0) Then
        '    For Each objProcess in colProcess
        '        objProcess.Terminate()
        '    Next
        '    Set objProcess = Nothing
        'End If

        Set objWMIService = Nothing
        Set colProcess = Nothing

        getCountInstancesOfExcel = wCount

    End Function

    ' ------------------------------------------------------
    ' Helper, check if a file is read-only
    ' ------------------------------------------------------
    Private Function IsReadOnly(sFileName)

        IsReadOnly = False

        If FileExists(sFileName) Then
            If objFSO.GetFile(sFileName).Attributes And 1 Then
                IsReadOnly = True
            End If
        End If

    End Function

    ' ------------------------------------------------------
    ' Helper, set the ReadOnly attribute for a file.
    ' Assuming that the file exists.
    ' ------------------------------------------------------
    Private Sub SetReadOnly(sFileName)

        Dim objFile

        On Error Resume Next

        If bVerbose Then
            wScript.echo " Make " & sFileName & " read-only (clsTransport::SetReadOnly)"
        End If

        Set objFile = objFSO.GetFile(sFileName)

        ' Not sure that the connected user can change
        ' file's attributes. If he can't don't raise an error
        objFile.Attributes = objFile.Attributes OR 1 ' 1 = readonly / OR = set

        If Err.Number <> 0 Then

            If bVerbose Then
                wScript.echo " ERROR - " & Err.Description
            End If

            Err.Clear

        End If

        On Error Goto 0

        Set objFile = Nothing

    End Sub

    ' ------------------------------------------------------
    ' Helper, remove the ReadOnly attribute for a file.
    ' Assuming that the file exists.
    ' ------------------------------------------------------
    Private Sub SetReadWrite(sFileName)

        Dim objFile

        On Error Resume Next

        Set objFile = objFSO.GetFile(sFileName)

        If bVerbose Then
            wScript.echo " Make " & sFileName & " writable (clsTransport::SetReadWrite)"
        End If

        ' Not sure that the connected user can change
        ' file's attributes. If he can't don't raise an error
        If objFile.Attributes AND 1 Then
            objFile.Attributes = objFile.Attributes XOR 1 ' 1 = readonly / XOR = remove
        End If

        If Err.Number <> 0 Then

            If bVerbose Then
                wScript.echo " ERROR - " & Err.Description
            End If

            Err.Clear

        End If

        On Error Goto 0

        Set objFile = Nothing

    End Sub

    ' ------------------------------------------------------
    ' Helper, set the Hidden attribute for a file.
    ' Assuming that the file exists.
    ' ------------------------------------------------------
    Private Sub SetHidden(sFileName)

        Dim objFile

        On Error Resume Next

        If bVerbose Then
            wScript.echo " Make " & sFileName & " hidden (clsTransport::SetHidden)"
        End If

        Set objFile = objFSO.GetFile(sFileName)

        ' Not sure that the connected user can change
        ' file's attributes. If he can't don't raise an error
        objFile.Attributes = objFile.Attributes OR 2 ' 2 = Hidden

        If Err.Number <> 0 Then

            If bVerbose Then
                wScript.echo " ERROR - " & Err.Description
            End If

            Err.Clear

        End If

        On Error Goto 0

        Set objFile = Nothing

    End Sub

    ' ------------------------------------------------------
    ' Helper, copy a file to another location
    ' ------------------------------------------------------
    Private Sub Copy(sSource, sTarget)

        Dim bReadOnly

        bReadOnly = IsReadOnly(sTarget)

        If bReadOnly Then

            ' Remove the read-only attribute
            SetReadWrite(sTarget)

        End If

        If bVerbose Then
            wScript.echo " Copy " & sSource & " to " & sTarget & " (clsTransport::Copy)"
        End If

        On Error Resume Next
        objFSO.CopyFile sSource, sTarget

        If Err.number <> 0 Then
            wScript.echo " ERROR - Copying " & sSource & " to " & sTarget & " has failed (clsTransport::Copy)"
            wScript.echo " " & Err.Description
            wScript.Quit cErrorCopyFileInUse
            Err.Clear
        End If

        On Error Goto 0

        If Not (FileExists(sTarget)) Then

            wScript.echo " ERROR - The file " & sTarget & " wasn't copied (clsTransport::Copy)"

        Else

            If bReadOnly Then

                ' And set it again
                SetReadOnly(sTarget)

            End If

        End If

    End Sub

    ' ------------------------------------------------------
    ' Helper, remove a file
    ' ------------------------------------------------------
    Private Sub Delete(sFile)

        Dim bReadOnly

        If (FileExists(sFile)) Then

            bReadOnly = IsReadOnly(sFile)

            If bReadOnly Then
                ' Remove the read-only attribute
                SetReadWrite(sFile)
            End If

            If bVerbose Then
                wScript.echo " Delete " & sFile & "  (clsTransport::Delete)"
            End If

            On Error Resume Next

            objFSO.DeleteFile sFile

            If Err.Number <> 0 Then

                If bVerbose Then
                    wScript.echo " ERROR - " & Err.Description
                End If

                Err.Clear

            End If

            On Error Goto 0

        End If

    End Sub

    ' ------------------------------------------------------
    ' Copy the development file to the staging folder
    ' ------------------------------------------------------
    Public Sub CopyDevToStaging()

        Dim sSource, sTarget

        'Call Check()

        sSource = cParameters.SourceFileFullName
        sTarget = StagingFolder & cParameters.TargetFileName

        If bVerbose Then
            wScript.echo " Copy from DEV to Staging (clsTransport::CopyDevToStaging)"
            wScript.echo ""
        End If

        ' Process the file; make sure the file can be overwritten if
        ' already present in the staging folder
        If (FileExists(sTarget)) Then
            Delete(sTarget)
        End If

        Call Copy(sSource, sTarget)

    End Sub

    ' ------------------------------------------------------
    ' Copy from the staging folder to the final destination
    ' ------------------------------------------------------
    Public Sub CopyStagingToFinal()

        Dim sSource, sTarget

        'Call Check()

        sSource = StagingFolder & cParameters.TargetFileName
        sTarget = cParameters.TargetFileFullName

        ' Create the target folder if not yet there
        If Not (objFSO.FolderExists(cParameters.TargetFolder)) Then
            Call objFSO.CreateFolder(cParameters.TargetFolder)
        End If

        If bVerbose Then
            wScript.echo " Copy from Staging To Final (clsTransport::CopyStagingToFinal)"
            wScript.echo ""
        End If

        ' Process the file; make sure the file can be overwritten if
        ' already present in the staging folder
        If (FileExists(sTarget)) Then
            Delete(sTarget)
        End If

        Call Copy(sSource, sTarget)

        If cParameters.ReadOnly Then
            SetReadOnly(sTarget)
        End If

        If cParameters.Hidden Then
            SetHidden(sTarget)
        End If

    End Sub

    ' ------------------------------------------------------
    ' Instantiate Excel and set properties
    ' ------------------------------------------------------
    Private Sub InstantiateExcel

        If (cMSExcel Is Nothing) Then

            If bVerbose Then
                wScript.echo " Instantiate Excel (clsTransport::InstantiateExcel)"
                wScript.echo ""
            End If

            Set cMSExcel = New clsMSExcel
            cMSExcel.Verbose = bVerbose

            Call cMSExcel.Instantiate()

            cMSExcel.DisplayAlerts = False
            cMSExcel.EnableEvents = False

        End If

        ' In no cases, Excel can be visible
        If (cMSExcel.app.Visible) Then

            wScript.echo " ERROR - Excel is actually visible and this will give problems."
            wScript.echo " When Excel is visible and if the file to process contains a"
            wScript.echo " ribbon, the onLoad code will be executed even if EnableEvents"
            wScript.echo " is unset and therefore replacing addins will fails."
            wScript.echo ""
            wScript.echo " Before running this script; be sure Excel isn't opened."

            cMSExcel.Quit

            wScript.Quit -99

        End If

    End Sub

    ' ------------------------------------------------------
    ' Remove a reference from the file present in the staging
    ' folder
    ' ------------------------------------------------------
    Public Sub RemoveAddIn()

        Dim bReadOnly
        Dim sFile

        If ((cParameters.OldAddInFullName = "") Or _
            (Not FileExists(cParameters.OldAddInFullName))) Then
            Exit Sub
        End If

        sFile = StagingFolder & cParameters.TargetFileName

        If bVerbose Then
            wScript.echo " Remove " & cParameters.OldAddInBaseName & " addIn " & _
                "from " & sFile & " (clsTransport::RemoveAddIn)"
            wScript.echo ""
        End If

        ' Open and initialize Excel
        Call InstantiateExcel

        ' We'll update the target file so should be writable
        bReadOnly = IsReadOnly(sFile)

        ' Remove the read-only attribute
        If bReadOnly Then SetReadWrite(sFile)

        ' Open the file
        ' False = not read-only since we'll modify it
        Call cMSExcel.Open(sFile, False)

        ' Remove and replace references
        ' The addin parameter should only be the name; no path, no extension
        Call cMSExcel.References_Remove(sFile, cParameters.OldAddInBaseName)

        ' Process the second addin to remove if there is one
        If Not (cParameters.OldAddIn2BaseName = "") Then
            Call cMSExcel.References_Remove(sFile, cParameters.OldAddIn2BaseName)
        End If

        Call cMSExcel.SaveFile(sFile)

        ' And set the RO attribute again if needed
        If bReadOnly Then SetReadOnly(sFile)

    End Sub

    ' ------------------------------------------------------
    ' Add a reference to the file present in the staging
    ' folder
    ' ------------------------------------------------------
    Public Sub AddAddIn()

        Dim bReadOnly
        Dim sFile

        If ((cParameters.NewAddInFullName = "") Or _
            (Not FileExists(cParameters.NewAddInFullName))) Then
            Exit Sub
        End If

        sFile = StagingFolder & cParameters.TargetFileName

        If bVerbose Then
            wScript.echo ""
            wScript.echo " Add " & cParameters.NewAddInFullName & " " & _
                "in the list of addins of " & sFile & " (clsTransport::AddAddIn)"
            wScript.echo ""
        End If

        ' We'll update the target file so should be writable
        bReadOnly = IsReadOnly(sFile)

        ' Remove the read-only attribute
        If bReadOnly Then SetReadWrite(sFile)

        Call InstantiateExcel

        ' Open the file
        ' False = not read-only since we'll modify it
        Call cMSExcel.Open(sFile, False)

        ' Add a new reference
        Call cMSExcel.References_AddFromFile(sFile, cParameters.NewAddInFullName)

        ' Process the second addin to add if there is one
        If Not (cParameters.NewAddIn2FullName = "") Then
            Call cMSExcel.References_AddFromFile(sFile, cParameters.NewAddIn2FullName)
        End If

        Call cMSExcel.SaveFile(sFile)

        ' And set the RO attribute again if needed
        If bReadOnly Then SetReadOnly(sFile)

    End Sub

    ' ------------------------------------------------------
    ' Close the file
    ' ------------------------------------------------------
    Public Sub CloseStaging()

        Dim sFile

        sFile = StagingFolder & cParameters.TargetFileName

        If bVerbose Then
            wScript.echo ""
            wScript.echo " Close " & sFile & " (clsTransport::CloseStaging)"
            wScript.echo ""
        End If

        Call cMSExcel.CloseFile(sFile)

    End Sub

    ' ------------------------------------------------------
    ' Start Windows Explorer and open the target folder
    ' ------------------------------------------------------
    Public Sub OpenFinal()

        Dim objShell

        If cParameters.OpenFolder Then

            Set objShell = CreateObject("Wscript.Shell")

            wScript.echo "Opening folder " & cParameters.TargetFolder
            wScript.echo "Your file " & cParameters.TargetFileName & " should be there"
            objShell.Run "explorer.exe /e," & cParameters.TargetFolder

            Set objShell = Nothing

        End If

    End Sub

End Class

' ----------------------------------------------------------
' When the user double-clic on a .vbs file (from Windows explorer f.i.)
' the running process will be WScript.exe while it's CScript.exe when
' the .vbs is started from the command prompt.
'
' This subroutine will check if the script has been started with cscript
' and if not, will run the script again with cscript and terminate the
' "wscript" version. This is usefull when the script generate a lot of
' wScript.echo statements, easier to read in a command prompt.'
' ----------------------------------------------------------
Sub ForceCScriptExecution()

    Dim sArguments, Arg, sCommand

    If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then

        ' Get command lines paramters'
        sArguments = ""
        For Each Arg In WScript.Arguments
            sArguments=sArguments & Chr(34) & Arg & Chr(34) & Space(1)
        Next

        sCommand = "cmd.exe cscript.exe //nologo " & Chr(34) & _
        WScript.ScriptFullName & Chr(34) & Space(1) & Chr(34) & sArguments & Chr(34)

        ' 1 to activate the window
        ' true to let the window opened
        Call CreateObject("Wscript.Shell").Run(sCommand, 1, true)

        ' This version of the script (started with WScript) can be terminated
        wScript.quit

    End If

End Sub

' *******************
' *** Entry point ***
' *******************

Dim cParameters, cTransport
Dim sMsg, wResult

    ' Stop if the script has been executed with wScript.exe i.e.
    ' with Windows and restart the script with cScript.exe (under DOS)
    Call ForceCScriptExecution()

    ' Get command line parameters
    Set cParameters = New clsParameters

    ' Read command line parameters
    cParameters.Read()

    ' Enable or not the display of messages (errors will always be displayed)
    cParameters.Verbose = cParameters.SilentMode

    If Not cParameters.SilentMode Then
        wScript.echo " === TRANSPORT - Start ==="
        wScript.echo ""
    End If

    ' Ask confirmation
    If Not cParameters.Force Then
        sMsg = "Ready to copy " & vbCrLf & cParameters.SourceFileFullName & _
            vbCrLf & " to " & vbCrLf & cParameters.TargetFileFullName & _
            vbCrLf & vbCrLf & "Please confirm..."

        wResult = Msgbox(sMsg, vbYesNo+vbInformation, "Transport")

        If (wResult = vbNo) Then
            wScript.echo " You've choose ""No"" so stop; do nothing"
            wScript.echo ""
            wScript.Quit 0
        End If
    End If

    ' Start the transport
    Set cTransport = New clsTransport

    ' Enable or not the display of messages (errors will always be displayed)
    cTransport.Verbose = Not cParameters.SilentMode

    ' Initialize the parameters object
    cTransport.Parameters = cParameters

    ' Step 1 - Copy the file from DEV to Staging
    cTransport.CopyDevToStaging()

    ' Update addins only if needed
    If ((cParameters.OldAddInFullName <> "") Or _
        (cParameters.NewAddInFullName <> "")) Then

        ' Step 2 - Remove addin from Staging
        If (cParameters.OldAddInFullName <> "") Then
            cTransport.RemoveAddIn()
        End If

        ' Step 3 - Add new addin to Staging
        If (cParameters.NewAddInFullName <> "") Then
            cTransport.AddAddIn()
        End If

        ' Step 4 - Close the file
        cTransport.CloseStaging()

    End If

    ' Step 5 - Copy the file from Staging to his final destination
    cTransport.CopyStagingToFinal()

    ' Finally, open the folder where we've copied the files
    cTransport.OpenFinal()

    If Not cParameters.SilentMode Then
        wScript.echo ""
        wScript.echo " === TRANSPORT - Success - End ==="
    End If

    Set cTransport = Nothing
    Set cParameters = Nothing
