VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LIB to A Converter v2.0"
   ClientHeight    =   3375
   ClientLeft      =   6375
   ClientTop       =   2370
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   Visible         =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Convert"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Library Files"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6900
      Begin VB.CommandButton Command2 
         Caption         =   "Browse"
         Height          =   285
         Left            =   5280
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DLL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1000
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "LIB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   400
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "LIB to A Converter"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   3390
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim InitDir As String

Dim LibFile As String
Dim LibName As String
Dim DllFile As String

Dim DefFile As String
Dim AFile As String

' Maximum buffer size constant
Private Const MAX_PATH = 260

'----------------------------------------------------------------------------------
'                      Common Dialog Type Declarations
'----------------------------------------------------------------------------------

' Common Dialog OPEN/SAVE information type
Private Type OPENFILENAME
  lStructSize       As Long    ' Size of this type / structure
  hwndOwner         As Long    ' Handle to the owner of the dialog
  hInstance         As Long    ' Instance handle of .EXE that contains custom dialog template
  lpstrFilter       As String  ' File type filter
  lpstrCustomFilter As String  ' Sets / returns the custom file type
  nMaxCustFilter    As Long    ' Length of the lpstrCustomFilter buffer
  nFilterIndex      As Long    ' Sets / returns the index of the filter in the dialog
  lpstrFile         As String  ' Sets / returns the full path to the file
  nMaxFile          As Long    ' Length of the lpstrFile buffer
  lpstrFileTitle    As String  ' Returns the name of the file selected (w/o the path)
  nMaxFileTitle     As Long    ' Length of the lpstrFileTitle buffer
  lpstrInitialDir   As String  ' Sets the initial browsing directory for the dialog
  lpstrTitle        As String  ' Sets the title of the dialog to be displayed
  Flags             As Long    ' Sets / returns the flags used with the dialog
  nFileOffset       As Integer ' Number of characters from the beginning of the full path to the first letter of the file name
  nFileExtension    As Integer ' Number of characters from the beginning of the full path to the file extension
  lpstrDefExt       As String  ' Sets the default extention of the file
  lCustData         As Long    ' Data passed to hook function
  lpfnHook          As Long    ' Pointer to hook function
  lpTemplateName    As String  ' Custom template name
End Type

' General Common Dialog Return Codes
Private Const CDERR_NOERROR = 0
Private Const CDERR_DIALOGFAILURE = &HFFFF    ' The dialog box could not be created. The common dialog box function’s call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle.
Private Const CDERR_FINDRESFAILURE = &H6      ' The common dialog box function failed to find a specified resource.
Private Const CDERR_INITIALIZATION = &H2      ' The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available.
Private Const CDERR_LOADRESFAILURE = &H7      ' The common dialog box function failed to load a specified resource.
Private Const CDERR_LOADSTRFAILURE = &H5      ' The common dialog box function failed to load a specified string.
Private Const CDERR_LOCKRESFAILURE = &H8      ' The common dialog box function failed to lock a specified resource.
Private Const CDERR_MEMALLOCFAILURE = &H9     ' The common dialog box function was unable to allocate memory for internal structures.
Private Const CDERR_MEMLOCKFAILURE = &HA      ' The common dialog box function was unable to lock the memory associated with a handle.
Private Const CDERR_NOHINSTANCE = &H4         ' The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle.
Private Const CDERR_NOHOOK = &HB              ' The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure.
Private Const CDERR_NOTEMPLATE = &H3          ' The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template.
Private Const CDERR_REGISTERMSGFAIL = &HC     ' The RegisterWindowMessage function returned an error code when it was called by the common dialog box function.
Private Const CDERR_STRUCTSIZE = &H1          ' The lStructSize member of the initialization structure for the corresponding common dialog box is invalid.

' DLG_PrintDialog Function Return Codes
Private Const PDERR_CREATEICFAILURE = &H100A  ' The DLG_PrintDialog function failed when it attempted to create an information context.
Private Const PDERR_DEFAULTDIFFERENT = &H100C ' You called the DLG_PrintDialog function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer. (This error occurs when you store the DEVNAMES structure and the user changes the default printer by using the Control Panel.)
                                             ' To use the printer described by the DEVNAMES structure, clear the DN_DEFAULTPRN flag and call DLG_PrintDialog again. To use the default printer, replace the DEVNAMES structure (and the DEVMODE structure, if one exists) with NULL; and call DLG_PrintDialog again.
Private Const PDERR_DNDMMISMATCH = &H1009     ' The data in the DEVMODE and DEVNAMES structures describes two different printers.
Private Const PDERR_GETDEVMODEFAIL = &H1005   ' The printer driver failed to initialize a DEVMODE structure. (This error code applies only to printer drivers written for Windows versions 3.0 and later.)
Private Const PDERR_INITFAILURE = &H1006      ' The DLG_PrintDialog function failed during initialization, and there is no more specific extended error code to describe the failure. This is the generic default error code for the function.
Private Const PDERR_LOADDRVFAILURE = &H1004   ' The DLG_PrintDialog function failed to load the device driver for the specified printer.
Private Const PDERR_NODEFAULTPRN = &H1008     ' A default printer does not exist.
Private Const PDERR_NODEVICES = &H1007        ' No printer drivers were found.
Private Const PDERR_PARSEFAILURE = &H1002     ' The DLG_PrintDialog function failed to parse the strings in the [devices] section of the WIN.INI file.
Private Const PDERR_PRINTERNOTFOUND = &H100B  ' The [devices] section of the WIN.INI file did not contain an entry for the requested printer.
Private Const PDERR_RETDEFFAILURE = &H1003    ' The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was not NULL.
Private Const PDERR_SETUPFAILURE = &H1001     ' The DLG_PrintDialog function failed to load the required resources.

' DLG_ChooseFont Function Return Codes
Private Const CFERR_MAXLESSTHANMIN = &H2002   ' The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member.
Private Const CFERR_NOFONTS = &H2001          ' No fonts exist.

' DLG_GetOpenFileName / DLG_GetSaveFileName Function Return Codes
Private Const FNERR_BUFFERTOOSMALL = &H3003   ' The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the filename specified by the user. The first two bytes of the lpstrFile buffer contain an integer value specifying the size, in bytes (ANSI version) or characters (Unicode version), required to receive the full name.
Private Const FNERR_INVALIDFILENAME = &H3002  ' A filename is invalid.
Private Const FNERR_SUBCLASSFAILURE = &H3001  ' An attempt to subclass a list box failed because sufficient memory was not available.

'----------------------------------------------------------------------------------
'                       Common Dialog Flag Codes
'----------------------------------------------------------------------------------

' Common Dialog OPEN/SAVE Flag Constants (See flag documentation below)
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function DLG_GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function DLG_GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function DLG_GetLastError Lib "comdlg32.dll" Alias "CommDlgExtendedError" () As Long

' Checks to see if an error occured in one of the following functions, and if one did,
' displays it in a standard error message dialog:
'
'  - DLG_GetOpenFileName
'  - DLG_GetSaveFileName
'  - DLG_ChooseColor Lib
'  - DLG_ChooseFont Lib
'  - DLG_PrintDialog Lib
'  - DLG_PageSetupDialog
'
Private Function GetLastError_CDLG(Optional ByVal ShowError As Boolean = True, Optional ByVal LastFunctionCalled As String = "last") As Boolean
On Error Resume Next
  
  Dim ReturnCode As Long
  Dim TheMsg As String
  
  ReturnCode = DLG_GetLastError
  
  Select Case ReturnCode
    ' General Common Dialog Return Codes
    Case CDERR_NOERROR          ' 0
      GetLastError_CDLG = False
      Exit Function
    Case CDERR_DIALOGFAILURE, 65535 ' &HFFFF
      TheMsg = "The dialog box could not be created." & Chr(13) & "The common dialog box function’s call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle."
    Case CDERR_FINDRESFAILURE   ' &H6
      TheMsg = "The common dialog box function failed to find a specified resource."
    Case CDERR_INITIALIZATION   ' &H2
      TheMsg = "The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available."
    Case CDERR_LOADRESFAILURE   ' &H7
      TheMsg = "The common dialog box function failed to load a specified resource."
    Case CDERR_LOADSTRFAILURE   ' &H5
      TheMsg = "The common dialog box function failed to load a specified string."
    Case CDERR_LOCKRESFAILURE   ' &H8
      TheMsg = "The common dialog box function failed to lock a specified resource."
    Case CDERR_MEMALLOCFAILURE  ' &H9
      TheMsg = "The common dialog box function was unable to allocate memory for internal structures."
    Case CDERR_MEMLOCKFAILURE   ' &HA
      TheMsg = "The common dialog box function was unable to lock the memory associated with a handle."
    Case CDERR_NOHINSTANCE      ' &H4
      TheMsg = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but a corresponding instance handle was not provide."
    Case CDERR_NOHOOK           ' &HB
      TheMsg = "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but a pointer to a corresponding hook procedure was not provided."
    Case CDERR_NOTEMPLATE       ' &H3
      TheMsg = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but a corresponding template was not provided."
    Case CDERR_REGISTERMSGFAIL  ' &HC
      TheMsg = "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
    Case CDERR_STRUCTSIZE       ' &H1
      TheMsg = "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
    
    ' DLG_PrintDialog Function Return Codes
    Case PDERR_CREATEICFAILURE  ' &H100A
      TheMsg = "The DLG_PrintDialog function failed when it attempted to create an information context."
    Case PDERR_DEFAULTDIFFERENT ' &H100C
      TheMsg = "You called the DLG_PrintDialog function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer." & Chr(13) & Chr(13) & "This error occurs when the DEVNAMES structure is stored and the user changes the default printer by using the Control Panel."
    Case PDERR_DNDMMISMATCH     ' &H1009
      TheMsg = "The data in the DEVMODE and DEVNAMES structures describes two different printers."
    Case PDERR_GETDEVMODEFAIL   ' &H1005
      TheMsg = "The printer driver failed to initialize a DEVMODE structure." & Chr(13) & "This error code applies only to printer drivers written for Windows versions 3.0 and later."
    Case PDERR_INITFAILURE      ' &H1006
      TheMsg = "The DLG_PrintDialog function failed during initialization, and there is no more specific extended error code to describe the failure."
    Case PDERR_LOADDRVFAILURE   ' &H1004
      TheMsg = "The DLG_PrintDialog function failed to load the device driver for the specified printer."
    Case PDERR_NODEFAULTPRN     ' &H1008
      TheMsg = "A default printer does not exist."
    Case PDERR_NODEVICES        ' &H1007
      TheMsg = "No printer drivers were found."
    Case PDERR_PARSEFAILURE     ' &H1002
      TheMsg = "The DLG_PrintDialog function failed to parse the strings in the [devices] section of the WIN.INI file."
    Case PDERR_PRINTERNOTFOUND  ' &H100B
      TheMsg = "The [devices] section of the WIN.INI file did not contain an entry for the requested printer."
    Case PDERR_RETDEFFAILURE    ' &H1003
      TheMsg = "The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was not NULL."
    Case PDERR_SETUPFAILURE     ' &H1001
      TheMsg = "The DLG_PrintDialog function failed to load the required resources."
    
    ' DLG_ChooseFont Function Return Codes
    Case CFERR_MAXLESSTHANMIN   ' &H2002
      TheMsg = "The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member."
    Case CFERR_NOFONTS          ' &H2001
      TheMsg = "No fonts exist."
    
    ' DLG_GetOpenFileName / DLG_GetSaveFileName Function Return Codes
    Case FNERR_BUFFERTOOSMALL   ' &H3003
      TheMsg = "The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the filename specified by the user."
    Case FNERR_INVALIDFILENAME  ' &H3002
      TheMsg = "A filename is invalid."
    Case FNERR_SUBCLASSFAILURE  ' &H3001
      TheMsg = "An attempt to subclass a list box failed because sufficient memory was not available."
    
    ' Other error
    Case Else
      TheMsg = "Unknown Error"
  End Select
  
  GetLastError_CDLG = True
  
  If ShowError = True Then
    MsgBox "COMDLG32.DLL encountered the following error while calling the '" & LastFunctionCalled & "' function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(ReturnCode) & Chr(13) & "Error Description = " & TheMsg, vbOKOnly + vbExclamation, "  Common Dialog Error"
  End If
  
End Function

' Function that takes the standard "Pipe Seperated" filter statements for
' the ShowOpen / ShowSave functions and changes them to a format that the
' COMDLG32.DLL API functions expect ("NULL Seperated" + Double NULL End)
Private Function StripFilter(ByVal TheFilter As String) As String
On Error Resume Next
  
  Dim MyCounter As Integer
  Dim CharLeft As String
  Dim CharRight As String
  Dim StringSoFar As String
  
  ' If there is no filter specified, return NULL
  If TheFilter = "" Then
    StripFilter = Chr(0)
    Exit Function
  End If
  
  ' Parse the string looking for the VB standard PIPE "|" separator and
  ' replace it with a NULL if found
  For MyCounter = 1 To Len(TheFilter)
    CharLeft = Left(TheFilter, MyCounter)
    CharRight = Right(CharLeft, 1)
    If CharRight = "|" Then
      CharRight = Chr(0)
    End If
    StringSoFar = StringSoFar & CharRight
  Next
  
  ' The filter always has to be terminated with 2 NULLs
  StringSoFar = StringSoFar & Chr(0) & Chr(0)
  StripFilter = StringSoFar
  
End Function

'=============================================================================================================
' CD_ShowOpen_Save
'
' Purpose :
' Displays the standard Windows "Open File" or "Save File" dialog depending
' on which the user wants to display.
'
' Compare to COMDLG32.OCX ShowOpen & ShowSave methods.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( OFN_... )
' FileName             Optional. Sets / recieves the full file path.
'                      If the OFN_ALLOWMULTISELECT flag is used, this returns the
'                      path where the files where selected from, followed by a
'                      NULL (vbNullChar), followed by a NULL-seperated list of
'                      the files that were selected, followed by a DOUBLE-NULL
'                      terminator at the end.
' FileTitle            Optional. Sets / recieves the file name (w/o path)
'                      If the OFN_ALLOWMULTISELECT flag is used, this returns a
'                      blank string.
' DefaultExt           Optional. Sets the default file extension
' DialogTitle          Optional. Sets the titlebar caption of the dialog
' Filter               Optional. Sets / recieves the browse file filter
' InitDir              Optional. Sets the initial browse directory
' Open_Not_Save        Optional. TRUE  = ShowOpen
'                                FALSE = ShowSave
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags     = Dialog return flags
'                    FileName  = Full file path for selected file
'                    FileTitle = File name selected (w/o path)
'                    Fileter   = Filter used to select the file
'
'=============================================================================================================
Private Function CD_ShowOpen_Save(ByVal OwnerHandle As Long, _
                                 Optional ByRef Flags As Long, _
                                 Optional ByRef FileName As String, _
                                 Optional ByRef FileTitle As String, _
                                 Optional ByVal DefaultExt As String, _
                                 Optional ByVal DialogTitle As String, _
                                 Optional ByRef Filter As String = "All Files (*.*)|*.*", _
                                 Optional ByVal InitDir As String = "C:\", _
                                 Optional ByVal Open_Not_Save As Boolean = True)
On Error GoTo ErrorTrap
  
  Dim FileInfo As OPENFILENAME
  Dim ReturnCode As Long
  Dim ReturnFilePath As String
  Dim ReturnFileName As String
  Dim SelectedFilter As String
  Dim FilterIndex As Long
  Dim FileNameOffset As Long
  Dim ExtensionOffset As Long
  
  ' Initialize the buffers to recieve the paths
  ReturnFilePath = FileName & String(MAX_PATH - Len(FileName), Chr(0))
  ReturnFileName = String(MAX_PATH, Chr(0))
  
  ' Initialize the variables that will be used to return the selected filter
  SelectedFilter = String(MAX_PATH, Chr(0))
  FilterIndex = 1
  
  ' Setup the initial directory
  If InitDir = "" Then InitDir = CurDir
  
  ' Make sure DefaultExt is correct
  If DefaultExt = "" Then DefaultExt = Chr(0)
  
  ' Make sure strings are NULL terminated
  If Right(DefaultExt, 1) <> Chr(0) Then DefaultExt = DefaultExt & Chr(0)
  If Right(InitDir, 1) <> Chr(0) Then InitDir = InitDir & Chr(0)
  If Right(DialogTitle, 1) <> Chr(0) Then DialogTitle = DialogTitle & Chr(0)
  
  ' Setup the information to use
  With FileInfo
    .lStructSize = Len(FileInfo)          ' Specifies the length of this type/structure
    .hwndOwner = OwnerHandle              ' Handle of the owner window for the dialog box. If this = NULL then the dialog has no owner
    .lpstrFilter = StripFilter(Filter)    ' Sets the specified filter correctly via the StripFilter function
    .lpstrCustomFilter = SelectedFilter   ' > Sets / recieves the selected filter. If this = NULL then no filter is returned
    .nFilterIndex = FilterIndex           ' Sets the default filter index to use.  1,2,3, etc. for standard filters... 0 if .lpstrCustomFilter is used to return the filter selected
    .lpstrFile = ReturnFilePath           ' > Sets / recieves the full path of the selected file(s)
    .lpstrFileTitle = ReturnFileName      ' > Sets / recieves the file name & extention of the selected file (this can be NULL)
    .lpstrInitialDir = InitDir            ' Sets the initial directory to browse from (if this is NULL, the current directory is used)
    .lpstrTitle = DialogTitle             ' Sets the caption of the dialog's titlebar (if this is NULL, the default "Save As" / "Open" is used)
    .Flags = Flags                        ' > Sets / receives the flags for the dialog
    .lpstrDefExt = DefaultExt             ' Sets the extension to attach to a file name if none is specified by the user.  If this is NULL, no extension is attached to the filename if the user doesn't specify one
    .nFileOffset = FileNameOffset         ' > Sets / recieves the number of characters to the left the first letter of the file name is located at (if lpstrFile = "c:\dir1\dir2\file.ext", nFileOffset = 13)
    .nFileExtension = ExtensionOffset     ' > Sets / recieves the number of characters to the left the first letter of the file extention is located at (if lpstrFile = "c:\dir1\dir2\file.ext", nFileExtension = 18)
    .nMaxFile = Len(ReturnFileName)       ' Specifies the length of .lpstrFile
    .nMaxFileTitle = Len(ReturnFileName)  ' Specifies the length of .lpstrFileTitle
    .nMaxCustFilter = Len(SelectedFilter) ' Specifies the length of the returned Filter from .lpstrCustomFilter (Ignored if .lpstrCustomFilter is NULL or vbNullString)
'------
    .hInstance = 0                        ' This is only needed if Flags contains OFN_ENABLETEMPLATEHANDLE, OFN_ENABLETEMPLATE, or OFN_EXPLORER
    .lCustData = 0                        ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message’s lParam parameter is a pointer to the OPENFILENAME structure specified when the dialog box was created. The hook procedure can use this pointer to get the lCustData value.
    .lpfnHook = 0                         ' Pointer to a hook procedure. This member is ignored unless the Flags member includes the OFN_ENABLEHOOK flag.
                                          ' If the OFN_EXPLORER flag is not set in the Flags member, lpfnHook is a pointer to an OFNHookProcOldStyle hook procedure that receives messages intended for the dialog box. The hook procedure returns FALSE to pass a message to the default dialog box procedure or TRUE to discard the message.
                                          ' If OFN_EXPLORER is set, lpfnHook is a pointer to an OFNHookProc hook procedure. The hook procedure receives notification messages sent from the dialog box. The hook procedure also receives messages for any additional controls that you defined by specifying a child dialog template. The hook procedure does not receive messages intended for the standard controls of the default dialog box.
    .lpTemplateName = vbNullString        ' Pointer to a null-terminated string that names a dialog template resource in the module identified by the hInstance member. This member is ignored unless the OFN_ENABLETEMPLATE flag is set in the Flags member.
                                          ' If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
  End With
  
  ' Show open or save dialog depending on the use specification
  If Open_Not_Save = True Then
    ReturnCode = DLG_GetOpenFileName(FileInfo)
  Else
    ReturnCode = DLG_GetSaveFileName(FileInfo)
  End If
    
  ' Check if dialog canceled or error occured
  If GetLastError_CDLG(True, "DLG_GetOpenFileName") = True Or ReturnCode = 0 Then
    Flags = 0
    FileName = ""
    FileTitle = ""
    Filter = ""
    CD_ShowOpen_Save = False
    
  ' Return the information
  Else
    
    With FileInfo
      SelectedFilter = .lpstrCustomFilter
      Do While Left(SelectedFilter, 1) = Chr(0)
        SelectedFilter = Right(SelectedFilter, Len(SelectedFilter) - 1)
      Loop
      If SelectedFilter <> "" Then
         SelectedFilter = Left(SelectedFilter, InStr(SelectedFilter, Chr(0)) - 1)
      End If
      ReturnFilePath = .lpstrFile
      ReturnFileName = .lpstrFileTitle
    End With
    If (Flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
      Do While Right(ReturnFilePath, 1) = Chr(0)
        ReturnFilePath = Left(ReturnFilePath, Len(ReturnFilePath) - 1)
      Loop
      FileName = ReturnFilePath & Chr(0) & Chr(0)
    Else
      FileName = Left(ReturnFilePath, InStr(ReturnFilePath, Chr(0)) - 1)
    End If
    FileTitle = Left(ReturnFileName, InStr(ReturnFileName, Chr(0)) - 1)
    Filter = SelectedFilter
    Flags = FileInfo.Flags
    CD_ShowOpen_Save = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowOpen_Save function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowOpen_Save = False
  End If
  
End Function

Private Sub Command1_Click()

  Dim Result As String

  With Me
    Call CD_ShowOpen_Save(.hwnd, OFN_PATHMUSTEXIST & OFN_FILEMUSTEXIST, LibFile, Result, "", "", _
      "Linker Library (*.lib)" & Chr$(0) & "*.lib" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*", _
      InitDir)

    If (LibFile <> "") Then
      .Text1.Text = LibFile
      ' LibName = Strings.Left(LibName, Strings.InStr(LibName, "."))
      '.Text1.Text = LibName
      LibName = Result
    End If
    .Text1.SelStart = .Text1.SelLength

    If (.Text2.Text <> "") Then
      .Command4.Enabled = True
    Else
      .Command4.Enabled = False
    End If
  End With

End Sub

Private Sub Command2_Click()

  With Me
    Call CD_ShowOpen_Save(.hwnd, OFN_PATHMUSTEXIST & OFN_FILEMUSTEXIST, DllFile, "", "", "", _
      "Dynamic Library (*.dll)" & Chr$(0) & "*.dll" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*", _
      InitDir)

    If (DllFile <> "") Then .Text2.Text = DllFile
    .Text2.SelStart = .Text2.SelLength

    If (.Text1.Text <> "") Then
      .Command4.Enabled = True
    Else
      .Command4.Enabled = False
    End If
  End With

End Sub

Private Sub Command4_Click()


  Call Shell(App.Path & "\bin\reimp.exe -d " & LibFile, vbHide)

  ' LibName = Strings.Left(LibName, 4)
  LibName = Strings.Left(LibName, Strings.InStr(LibName, "."))
  DefFile = LibName & "def"
  AFile = "lib" & LibName & "a"

  Call Shell(App.Path & "\bin\dlltool.exe -d " & DefFile & " -D " & DllFile & " -k -l " & AFile, vbHide)

End Sub

Private Sub Command5_Click()

  Call ShellAbout(Me.hwnd, "LIB to A Converter", "FX Programmer", Me.Icon)

End Sub

Private Sub Form_Load()

  InitDir = App.Path & "\convert"
  LibFile = ""
  LibName = ""
  DllFile = ""

  With Me
    ' Center Window
    .Top = (Screen.Height / 2) - (.Height / 2)
    .Left = (Screen.Width / 2) - (.Width / 2)

    .Visible = True
  End With

End Sub

Private Sub Text1_Change()

  With Me
    If (.Text1.Text = "") Then
      .Command4.Enabled = False
    Else
      If (.Text2.Text <> "") Then
        .Command4.Enabled = True
      End If
    End If
  End With

End Sub

Private Sub Text2_Change()

  With Me
    If (.Text2.Text = "") Then
      .Command4.Enabled = False
    Else
      If (.Text1.Text <> "") Then
        .Command4.Enabled = True
      End If
    End If
  End With

End Sub
