VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDirDrill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************
'  Copyright ©1992-2007, Karl E. Peterson
'  http://vb.mvps.org/
' *************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************
Option Explicit

' Win32 API declarations
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Const INVALID_FILE_ATTRIBUTES As Long = -1&

' Default values
Private Const defPattern As String = "*.*"
Private Const defRecurse As Boolean = True
Private Const defAttrMask As Long = 0

' Member variables
Private m_Cancel As Boolean
Private m_Pattern As String
Private m_Recurse As Boolean
Private m_Folder As String
Private m_TotalFiles As Long
Private m_TotalDirs As Long
Private m_AttrMask As Long

' Events
Public Event NewFile(ByVal filespec As String, ByRef Cancel As Boolean)
Public Event NewFolder(ByVal FolderSpec As String, ByRef Cancel As Boolean)
Public Event Done(ByVal TotalFiles As Long, ByVal TotalFolders As Long)

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Set defaults
   m_Folder = CurDir$()
   m_Pattern = defPattern
   m_Recurse = defRecurse
   m_AttrMask = defAttrMask
End Sub

Private Sub Class_Terminate()
   '
End Sub

' *********************************************
'  Properties
' *********************************************
Public Property Let AttributeMask(ByVal NewMask As Long)
   ' Needed in order to find hidden/system files, or
   ' simply to narrow the search to specific attribs.
   m_AttrMask = NewMask
End Property

Public Property Get AttributeMask() As Long
   AttributeMask = m_AttrMask
End Property

Public Property Get Folder() As String
   Folder = m_Folder
End Property

Public Property Let Folder(ByVal NewValue As String)
   If IsDirectory(NewValue) Then
      m_Folder = NewValue
   End If
End Property

Public Property Get Pattern() As String
   Pattern = m_Pattern
End Property

Public Property Let Pattern(ByVal NewValue As String)
   ' Accept delimited list of filespecs, replacing all
   ' commas with semi-colons, and eliminating spaces.
   m_Pattern = Replace$(Replace$(NewValue, " ", ""), ",", ";")
End Property

Public Property Let Recursive(ByVal NewRecursive As Boolean)
   ' Flag indicates whether to recurse entire dirtree.
   m_Recurse = NewRecursive
End Property

Public Property Get Recursive() As Boolean
   Recursive = m_Recurse
End Property

' *********************************************
'  Public Methods
' *********************************************
Public Sub BeginSearch(Optional ByVal StartFolder As String)
   ' Assign new path, if asked to.
   If Len(StartFolder) Then
      Me.Folder = StartFolder
   End If
   ' Reset counters, and start recursion.
   m_Cancel = False
   m_TotalFiles = 0
   m_TotalDirs = 1  'always have a starting folder!
   Call RecurseFolders(m_Folder)
   RaiseEvent Done(m_TotalFiles, m_TotalDirs)
   ' Clear any outstanding locks held by Dir$()
   Call Dir$("nul")
End Sub

Public Function FileText(ByVal filespec As String) As String
   Dim hFile As Long
   On Error Resume Next
   ' Return entire file contents as string, ignoring errors.
   hFile = FreeFile
   Open filespec For Input As #hFile
   FileText = Input(LOF(hFile), #hFile)
   Close #hFile
End Function

Public Function ExtractName(ByVal SpecIn As String, Optional BaseOnly As Boolean = False) As String
   Dim nPos As Long
   Dim SpecOut As String
   
   ' Make sure we don't have a directory.
   If IsDirectory(SpecIn) = False Then
      ' Find last backslash, and keep what follows.
      nPos = InStrRev(SpecIn, "\")
      If nPos < Len(SpecIn) Then
         SpecOut = Mid$(SpecIn, nPos + 1)
      Else
         SpecOut = SpecIn
      End If
      
      ' If user requested, remove the extension.
      ' Useful for appending a different extension.
      If BaseOnly Then
         nPos = InStrRev(SpecOut, ".")
         If nPos < Len(SpecOut) Then
            SpecOut = Left$(SpecOut, nPos - 1)
         End If
      End If
   End If
   
   ' Return SpecOut
   ExtractName = SpecOut
End Function

Public Function ExtractPath(ByVal SpecIn As String) As String
   Dim nPos As Long
   Dim SpecOut As String
   
   ' Make sure we don't already have a directory.
   If IsDirectory(SpecIn) = True Then
      SpecOut = Backslashed(SpecIn)
   Else
      ' Find last backslash, and keep what follows.
      nPos = InStrRev(SpecIn, "\")
      If nPos Then
         SpecOut = Left$(SpecIn, nPos)
      End If
   End If
   
   ' Return SpecOut
   ExtractPath = SpecOut
End Function

' *********************************************
'  Private Methods
' *********************************************
Private Function Backslashed(ByVal SpecIn As String) As String
   ' Make sure path ends with backslash.
   If IsDirectory(SpecIn) Then
      If Right(SpecIn, 1) <> "\" Then
         SpecIn = SpecIn & "\"
      End If
   End If
   Backslashed = SpecIn
End Function

Private Function GetAttr(ByVal PathName As String) As Long
   'Const INVALID_FILE_ATTRIBUTES As Long = -1&
   ' Use API to retrieve file attributes, because VBA.GetAttr
   ' chokes on files that have attributes not used by the time
   ' VB5/6 were released.  :-(
   GetAttr = GetFileAttributes(PathName)
End Function

Private Function IsDirectory(ByVal PathName As String) As Boolean
   Dim Attr As Long
   ' Make sure we have a valid attribute to check.
   Attr = GetAttr(PathName)
   If Attr <> INVALID_FILE_ATTRIBUTES Then
      ' Check for presence of Directory attribute.
      If (Attr And vbDirectory) = vbDirectory Then
         IsDirectory = True
      End If
   End If
End Function

Private Sub RecurseFolders(ByVal StartPath As String)
   Dim Dirs() As String, Files() As String
   'Dim Patterns() As String  '<= Use this with VB6 Split
   Dim Patterns As Variant
   Dim nDirs As Long, nFiles As Long
   Dim FileName As String
   Dim This As String
   Dim i As Long
   Const nIncrement As Long = 100
   
   ' Make sure there's a trailing backslash.
   StartPath = Backslashed(StartPath)
   
   ' Alert user that a new folder is being processed.
   RaiseEvent NewFolder(StartPath, m_Cancel)
   
   If m_Cancel = False Then
      ' Build array of all directories under "here."
      ReDim Dirs(0 To nIncrement - 1) As String
      FileName = Dir(StartPath & "*.*", vbDirectory)
      Do While Len(FileName)
         This = StartPath & FileName
         If IsDirectory(This) Then
            ' Ignore dotted directory names
            If Left$(FileName, 1) <> "." Then
               If nDirs > UBound(Dirs) Then
                  ReDim Preserve Dirs(0 To UBound(Dirs) + nIncrement) As String
               End If
               Dirs(nDirs) = This
               nDirs = nDirs + 1
            End If
         End If
         ' Continue...
         FileName = Dir()
      Loop
      
      ' Build array of wildcard filespec patterns.
      Patterns = Split(m_Pattern, ";")
      
      ' Begin scanning for all files that match each spec.
      ReDim Files(0 To nIncrement - 1) As String
      For i = LBound(Patterns) To UBound(Patterns)
         FileName = Dir(StartPath & Patterns(i), m_AttrMask)
         Do While Len(FileName)
            This = StartPath & FileName
            If IsDirectory(This) = False Then
               'If (GetAttr(This) And m_AttrMask) Then
                  If nFiles > UBound(Files) Then
                     ReDim Preserve Files(0 To UBound(Files) + nIncrement) As String
                  End If
                  Files(nFiles) = This
                  nFiles = nFiles + 1
               'End If
            End If
            ' Continue...
            FileName = Dir()
         Loop
      Next i
   End If
      
   ' Increment counters
   m_TotalFiles = m_TotalFiles + nFiles
   m_TotalDirs = m_TotalDirs + nDirs
   
   ' Process each file in this directory.
   For i = 0 To nFiles - 1
      RaiseEvent NewFile(Files(i), m_Cancel)
      If m_Cancel Then Exit For
   Next i
   
   ' Process each directory found.
   If (m_Recurse = True) Then
      For i = 0 To nDirs - 1
         RecurseFolders Dirs(i)
         ' Bail if user cancelled in last procedure.
         If m_Cancel Then Exit For
      Next i
   End If
End Sub

