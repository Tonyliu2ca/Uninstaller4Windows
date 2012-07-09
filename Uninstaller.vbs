' ---------------------------------------------------------------------------
' Name: Uninstaller for Windows
'
' Version 0.1
' Support Windows Xp, Vista, 7 (x86 or x64)
'
' History:
'    Created: May 14, 2012
'    Updates: June 20, 2012: Initial
'
' Description: 
'    This script list/test/uninstall a program. User can provide commands/options
'    to specify program to be operated.
'    For help, run: cscript Uninstaller.vbs /h
'
' Copyright (c) 2012, Tony Liu
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
'
' Contact: Tonyliu2ca@gmail.com
'
'
On Error Resume Next


'----------------------------------------------------
' Registry Const
Const HKLM = &H80000002        'HKEY_LOCAL_MACHINE
Const strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Const strKey64 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
' Commands
Const OPT_LISTONLY="/L"
Const OPT_UNINSTALL="/U"
Const OPT_HELP="/H"
Const OPT_TEST="/T"
' Options
Const OPT_COMPUTER="/S"
Const OPT_COMPARE="/C"
Const OPT_NAME="/D"
Const OPT_QUIET="/Q"
Const OPT_MAJOR="/Ma"
Const OPT_MINOR="/Mi"
Const OPT_VERSION="/V"
Const OPT_OPTION="/O"

' Index of each item
Public  arrEntryType, arrEntrys
Const idx_DisplayName=0
Const idx_InstallDate=1
Const idx_VersionMajor=2
Const idx_VersionMinor=3
Const idx_EstimatedSize=4
Const idx_UninstallString=5
Const idx_Publisher=6
Const idx_DisplayVersion=7
Const idx_InstallLocation=8

'-----------------------------------------------------
Dim objDictionary, objFiltered ' Dictionary
Public pstrComputerName ' Computer Name
Public objReg, objReg64, strEntry, strKeyContent ' Registry Related
Public pobjWMIService, objFSObj, pstrCommand ' Others
Public pstrComputer, pstrCompare, piQuietMode, pstrDisplayName, pstrVerMajor, pstrVerMinor, pstrVersion, pstrUnOption ' Options
Public pobjShell, pwshShell

'-----------------------------------------------------
'  Registry Entry Items
arrEntryType = array("REG_SZ", "REG_SZ", "REG_DWORD", "REG_DWORD", "REG_DWORD", "REG_SZ", "REG_SZ", "REG_SZ", "REG_SZ")
arrEntrys = array("DisplayName", "InstallDate", "VersionMajor", "VersionMinor", "EstimatedSize", "UninstallString", "Publisher", "DisplayVersion", "InstallLocation")
Redim arrValues(UBound(arrEntrys))

' -----------------------------------------------
' Start
ListAppProcess
GetArguments
If Err <> 0 Then WScript.Quit
SysInit
If Err <> 0 Then WScript.Quit
ReadApps32
If err <> 0 Then DisInfo "ReadApps32: Error: " & err.number & ":" & err.description
ReadApps64
If err <> 0 Then DisInfo "ReadApps64: Error: " & err.number & ":" & err.description
Filter
ListApps
UnTest
Uninstall
WScript.Quit

'-------------------------------------------
' Show help info lines
'-------------------------------------------
Function ShowHelps ()
   On Error Resume Next
'   if piQuietMode > 0 then
'      Exit Function
'   End if
   WScript.Echo "Usage: command " & OPT_LISTONLY & " [options]"
   WScript.Echo "               " & OPT_UNINSTALL & " [options]"
   WScript.Echo "               /h : This help"
   WScript.Echo "   " & OPT_LISTONLY & ":" & vbTab & vbTab & " List installed application(s) info only."
   WScript.Echo "   " & OPT_UNINSTALL & ":" & vbTab & vbTab & " Uninstall application(s)."
   WScript.Echo VbCrLf & "   Options:" 
   WScript.Echo "      " & OPT_COMPUTER & ": Remote computer numner, example: " & OPT_COMPUTER & " w999-999"
   WScript.Echo "      " & OPT_NAME & ": Filer by display name, example: " & OPT_NAME & " ""Adobe"""
   WScript.Echo "      " & OPT_MAJOR & ": Filer by Major Version, example: " & OPT_MAJOR & " 12"
   WScript.Echo "      " & OPT_MINOR & ": Filer by Minor Version, example: " & OPT_MINOR & " 1"
   WScript.Echo "      " & OPT_VERSION & ": Filer by Display Version, example: " & OPT_VERSION & " 12.1.0"
   WScript.Echo "      " & OPT_OPTION & ": Uninstall options: " & OPT_OPTION & " /quiet /norestart"
   WScript.Echo "      " & OPT_COMPARE & ": Compare to match whole words Only."
   WScript.Echo "      " & OPT_QUIET & ": Quiet Mode."
   WScript.Quit
End Function


'-------------------------------------------
' Get all scripts' arguments
'-------------------------------------------
Function GetArguments ()
   On Error Resume Next
   Err.Clear

   ' DisInfo "Get Command Arguments"
   Set objArgs = WScript.Arguments
   If objArgs.Count <= 0 Then
      ShowHelps
      Exit Function
   End If

   ' -----------------
   ' Commands
   pstrCommand = UCase(objArgs(0))
   ' DisInfo "0:" & pstrCommand
   If pstrCommand <> OPT_LISTONLY AND pstrCommand <> OPT_UNINSTALL AND pstrCommand <> OPT_TEST Then
      ShowHelps
      Exit Function
   End If
   
   ' -----------------
   'Options
   pstrComputer = "."
   pstrCompare = Null
   piQuietMode = 0
   pstrDisplayName = Null
   pstrVersion = Null
   pstrVerMajor = Null
   pstrVerMinor = Null
   for idx = 1 to objArgs.Count -1
      strOption = UCase(objArgs(idx))
      ' DisInfo idx & ": " & strOption
      select case strOption
      case OPT_COMPUTER: ' Remote Computer
         idx = idx + 1
         pstrComputer = Trim(objArgs(idx))
         ' Wscript.Echo idx, ":" & Trim(objArgs(idx))
      case OPT_COMPARE: ' Strict compare mode
         pstrCompare = OPT_COMPARE
         ' Wscript.Echo idx, ":" & Trim(objArgs(idx))
      case OPT_NAME:  ' Filer display name
         idx = idx + 1
         pstrDisplayName = Trim(objArgs(idx))
         ' Wscript.Echo idx, ":" & Trim(objArgs(idx))
      case OPT_QUIET: ' Quiet mode
         piQuietMode = 1
         ' Wscript.Echo idx, ":" & Trim(objArgs(idx))
      case OPT_VERSION:   ' Display Version
         idx = idx + 1
         pstrVersion = Trim(objArgs(idx))
         ' Wscript.Echo idx, ":" & Trim(objArgs(idx))
      case OPT_MAJOR: ' If delete this script self?
         idx = idx + 1
         pstrVerMajor = Trim(objArgs(idx))
         ' Wscript.Echo idx, ":" & Trim(objArgs(idx))
      case OPT_OPTION: ' Uninstall options
         idx= idx + 1
         pstrUnOption = " " & Trim(objArgs(idx))
      case OPT_MINOR:
         idx = idx + 1
         pstrVerMinor = Trim(objArgs(idx))
         ' Wscript.Echo idx, ":" & Trim(objArgs(idx))
      End Select
   Next
End Function

'-------------------------------------------
' System initialize
'-------------------------------------------
Function SysInit ()
   On Error Resume Next
   Const SYSINIT_ID = "<SysInit>:"
   strKeyContent = ""

   DisInfo "System Initializing..."
   Err.Clear 
   ' Get local machine name
   If pstrComputer = "." then
      Set objComputer = CreateObject("WScript.Network")
      pstrComputerName = UCase(Trim(objComputer.ComputerName))
   Else
      pstrComputerName = split(pstrComputer, ".", 1)(0)
   End If
   DisInfo "Computer Number=" & pstrComputerName

   Set objDictionary = CreateObject("Scripting.Dictionary")
   If err <> 0 Then
      DisInfo  "Program Initial: 1.Error: " & err.number & ":" & err.description
      Exit Function
   End If
   
   Set objFiltered = CreateObject("Scripting.Dictionary")
   If err <> 0 Then
      DisInfo  "Program Initial: 2.Error: " & err.number & ":" & err.description
      Exit Function
   End If
   
   Set objReg = GetObject("winmgmts://" & pstrComputer & "/root/default:StdRegProv")
   If err <> 0 Then
      DisInfo  "Program Initial: 3.Error: " & err.number & ":" & err.description
      Exit Function
   End If
   
   Set pobjShell = CreateObject("Shell.Application")
   If err <> 0 Then
      DisInfo  "Program Initial: 4.Error:" & err.number & ":" & err.description
      Exit Function
   End If
   
   Set pwshShell = CreateObject("WScript.Shell")
   If err <> 0 Then
      DisInfo  "Program Initial: 5.Error:" & err.number & ":" & err.description
      Exit Function
   End If
   
   Set pobjWMIService = GetObject("winmgmts:\\" & pstrComputer & "\root\cimv2")
   If err <> 0 Then
      DisInfo  "Program Initial: 6.Error:" & err.number & ":" & err.description
      Exit Function
   End If
   Err.Clear
End Function

'--------------------------------------------------------------
' Read all installed Applications for x32 system
'--------------------------------------------------------------
Function ReadApps32()
   On Error Resume Next

   DisInfo  "ReadApps32(): Start"
   objReg.EnumKey HKLM, strKey, arrSubkeys
   ' DisInfo "arrEntrys: " & UBound(arrEntrys)
   For Each strSubkey In arrSubkeys
      ' DisInfo "A: " & strSubkey
      For idx = 0 to UBound(arrEntrys)
      ' Get the next subkey
         ' DisInfo "Entry: " & arrEntrys(idx)
         ' DisInfo strSubkey & " : " & arrEntryType(idx)
         Select Case arrEntryType(idx)
         Case "REG_SZ"
             intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, arrEntrys(idx), varValue)
         Case "REG_DWORD"
             intRet1 = objReg.GetDWORDValue(HKLM, strKey & strSubkey, arrEntrys(idx), varValue)
         End Select
         if IsNull(varValue) = False Then
             arrValues(idx) = cstr(varValue)
        Else
            arrValues(idx) = ""
         End If
         arrValues(idx) = Trim(arrValues(idx))
         if err <> 0 Then DisInfo " ReadApps 32.Error: ", err.number & ":" & err.description
         ' DisInfo strSubkey & " : " & arrEntrys(idx) & "=" & arrValues(idx)
      Next
      objDictionary.add Trim(strSubkey), arrValues
   Next
   ' DisInfo  "ReadApps32(): End"
End Function

'--------------------------------------------------------------
' Read all installed Applications for x64 system
'--------------------------------------------------------------
Function ReadApps64()
   On Error Resume Next
   ' DisInfo "ReadApps64(): Start"
   objReg.EnumKey HKLM, strKey64, arrSubkeys
   ' DisInfo "arrEntrys: " & UBound(arrEntrys)
   If IsNull(arrSubkeys) Then
      Exit Function
   End If
   For Each strSubkey In arrSubkeys
      ' DisInfo "A: " & strSubkey
      For idx = 0 to UBound(arrEntrys)
      ' Get the next subkey
         Select Case arrEntryType(idx)
         Case "REG_SZ"
             intRet1 = objReg.GetStringValue(HKLM, strKey64 & strSubkey, arrEntrys(idx), varValue)
         Case "REG_DWORD"
             intRet1 = objReg.GetDWORDValue(HKLM, strKey64 & strSubkey, arrEntrys(idx), varValue)
         End Select
         if IsNull(varValue) = False Then
             arrValues(idx) = Trim(cstr(varValue))
        Else
            arrValues(idx) = ""
         End If
      Next
      objDictionary.add Trim(strSubkey), arrValues
   Next   
   ' DisInfo "ReadApps64(): End"
End Function

'--------------------------------------------------------------
' List all Apps
'--------------------------------------------------------------
Function ListApps ()
   On Error Resume Next
   ' DisInfo "Application number=" & objFiltered.Count
   'For each key in objFiltered.Keys
   '    DisInfo key & "," & Join(objFiltered.Item(key), ", ")
   'Next
   If pstrCommand = OPT_LISTONLY Then
      For each item in objFiltered.Items
         DisInfo Join(item, ", ")
      Next
   End If
End Function

'--------------------------------------------------------------
' List uninstall commands
'--------------------------------------------------------------
Function UnTest()
   On Error Resume Next
   If pstrCommand = OPT_TEST Then
      For each item in objFiltered.Items
         If Len(item(idx_UninstallString)) > 0 Then
            DisInfo item(idx_DisplayName) & ", " & item(idx_UninstallString)
         End If
      Next
   End If
End Function

'--------------------------------------------------------------
' Uninstall command
'--------------------------------------------------------------
Function Uninstall()
   On Error Resume Next
   ' DisInfo "Application number=" & objFiltered.Count
   'For each key in objFiltered.Keys
   '    DisInfo key & "," & Join(objFiltered.Item(key), ", ")
   'Next
   If pstrCommand = OPT_UNINSTALL Then
      For each key in objFiltered.Keys
         'each item in objFiltered.Items
         item = objFiltered.Item(key)
         If Len(item(idx_UninstallString)) > 0 Then
            If isMsi(item(idx_UninstallString)) Then
               DisInfo "Is MSI"
               WaitNoPro "msiexec.exe"
               ProcessMsi Key, pstrUnOption
            Else
               DisInfo "Not MSI"
               ProcessCmd item(idx_UninstallString), pstrUnOption
            End If
         Else
            DisInfo item(idx_DisplayName) & "'s uninstall command=(" & item(idx_UninstallString) &")"
         End If
      Next
   End If
   Set WshShell = Nothing
End Function

'--------------------------------------------------------------
' if it's msi command.
'--------------------------------------------------------------
Function isMsi(strCMD)
   isMsi = False
   If InStr(strCMD, "MsiExec.exe") > 0 Then
      isMsi = True
   End If   
End Function

'--------------------------------------------------------------
' Msi uninstall
'--------------------------------------------------------------
Function ProcessMsi (key, strOption)
'    strCMD= "Msiexec.exe /X " & Key & strOption
'    DisInfo strCMD'
'    pwshShell.Run strCMD, 0, true
    strCMD= "/X " & Key & strOption
    pobjShell.ShellExecute "MsiExec.exe", strCMD, "", "runas", 1
End Function

'--------------------------------------------------------------
' Program specified unstall command.
'--------------------------------------------------------------
Function ProcessCmd (strCmd, strOption)
   ' DisInfo "..." & strCMD & "..."
   pwshShell.Run strCMD & strOption, 0, true
'   pobjShell.ShellExecute strCMD, "", "", "runas", 1
'   Do While objShell.Status = 0
'      DisInfo "objShell.Status = " & objShell.Status            
'      WScript.Sleep 100 
'   Loop 
End Function


'--------------------------------------------------------------
' Filter function
'--------------------------------------------------------------
Function Filter ()
   On Error Resume Next
   ' Set objTempDic = CreateObject("Scripting.Dictionary")
   ' DisInfo "<Compare> = " & pstrCompare
   For each key in objDictionary.Keys
       ' DisInfo key & "," & Join(objDictionary.Item(key), ", ")
       arrCurrent = objDictionary.Item(key)
       If IsNull(pstrCompare) Then
       ' Soft filter
           bAdd = insCompare(arrCurrent(idx_DisplayName), pstrDisplayName)
           bAdd = bAdd AND insCompare(arrCurrent(idx_DisplayVersion), pstrVersion)
           bAdd = bAdd AND insCompare(arrCurrent(idx_VersionMajor), pstrVerMajor)
           bAdd = bAdd AND insCompare(arrCurrent(idx_VersionMinor), pstrVerMinor)
        Else
        ' Hard filter
           bAdd = eqCompare(arrCurrent(idx_DisplayName), pstrDisplayName)
           bAdd = bAdd AND eqCompare(arrCurrent(idx_DisplayVersion), pstrVersion)
           bAdd = bAdd AND eqCompare(arrCurrent(idx_VersionMajor) = pstrVerMajor)
           bAdd = bAdd AND eqCompare(arrCurrent(idx_VersionMinor), pstrVerMinor)
        End If
        ' If err <> 0 Then DisInfo "Wrong: " & err.number & ":" & err.description
        'If insCompare(arrCurrent(idx_DisplayName), "Adobe") Then
        '   DisInfo ".." & arrCurrent(idx_DisplayName) & " | " & bAdd
        'End If
        If bAdd = True Then
           ' DisInfo arrCurrent(idx_DisplayName)
           objFiltered.add Key, objDictionary.Item(key)
        End If

        Next
    DisInfo "Application number=" & objFiltered.Count & "/" & objDictionary.Count
End Function

'-------------------------------------------
' Instring compare
'-------------------------------------------
Function insCompare (First, Second)
   If IsNull(Second) Then
      insCompare = True
   Else
      If Instr (First, Second) > 0 Then
         insCompare = True
      Else
      insCompare = False
      End If
   End If
End Function

'-------------------------------------------
' Strict compare
'-------------------------------------------
Function eqCompare (First, Second)
   If IsNull(Second) Then
      eqCompare = True
   Else
       If First = Second Then
         eqCompare = True
       Else
         eqCompare = False
       End If
   End If
   ' If err <> 0 Then DisInfo "__Compare " & First & ":" & Second
End Function

'--------------------------------------------------------------
' Wait for a process to quit.
'--------------------------------------------------------------
Function WaitNoPro (strProgram)
On Error Resume Next
   bContinue = True
   Do
      Set colProcesses = pobjWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = " & "'" & strProgram & "' AND SessionID > 0")
      If colProcesses.Count = 0 Then
         bContinue = False
	  Else
         DisInfo "...Wait for all processes of <" & strProgram & "> to quit. " & colProcesses.Count
      End If 
   Loop While bContinue = True
End Function


'--------------------------------------------------------------
' Internal display
'--------------------------------------------------------------
Function DisInfo (strInfo)
   If piQuietMode < 1 Then
      Wscript.Echo strInfo
   End If
End Function

'--------------------------------------------------------------
' Test function for list all processes
'--------------------------------------------------------------
Function ListAppProcess ()
On Error Resume Next
   Set pobjWMIService = GetObject("winmgmts:\\.\root\cimv2")
   ' Wscript.Echo "...ListAppProcess.."
   Set colProcesses = pobjWMIService.ExecQuery("SELECT * FROM Win32_Process")
   For Each objProcess in colProcesses
      DisInfo "..." & objProcess.Name & ":" & objProcess.SessionID
   Next
End Function
