VERSION 5.00
Begin VB.Form FSearch 
   Caption         =   "Search Declares"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   3555
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "FSEARCH.frx":0000
      Top             =   1320
      Width           =   6195
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin..."
      Default         =   -1  'True
      Height          =   435
      Left            =   4080
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "FSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
'  Copyright ©1992-2005, Karl E. Peterson
'  http://vb.mvps.org/
' *************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************
Option Explicit

Private WithEvents dd As CDirDrill
Attribute dd.VB_VarHelpID = -1
Private m_Cancel As Boolean

Private Sub Command1_Click()
   ' Read search parameters.
   dd.Folder = Text1.Text
   dd.Pattern = Text2.Text
   dd.AttributeMask = vbHidden Or vbSystem Or vbArchive Or vbReadOnly
   ' Clear results text.
   Text3.Text = ""
   ' Clear cancel flag.
   m_Cancel = False
   ' Let it rip!
   dd.BeginSearch
End Sub

Private Sub dd_Done(ByVal TotalFiles As Long, ByVal TotalFolders As Long)
   DebugOutput "Found " & TotalFiles & " files in " & TotalFolders & " folders."
   Clipboard.Clear
   Clipboard.SetText Text3.Text
End Sub

Private Sub dd_NewFile(ByVal filespec As String, Cancel As Boolean)
   ' Show that we found one...
   DebugOutput Space$(5) & dd.ExtractPath & dd.ExtractName(filespec)
   ' Do any other processing here...
End Sub

Private Sub dd_NewFolder(ByVal FolderSpec As String, Cancel As Boolean)
   ' Output new folder found.
   DebugOutput FolderSpec
   ' Take a breath, bail if Escape was pressed.
   DoEvents
   Cancel = m_Cancel
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   ' Set cancel flag if Escape key was pressed.
   m_Cancel = (KeyAscii = vbKeyEscape)
End Sub

Private Sub Form_Load()
   ' Enable watching for an escape key.
   Me.KeyPreview = True
   ' Put some default text into controls.
   Text1.Text = App.Path
   Text2.Text = "*.bas;*.cls;*.frm;*.ctl"
   Text3.Text = ""
   ' Instantiate recursive search class.
   Set dd = New CDirDrill
End Sub

Private Sub DebugOutput(ByVal Data As String, Optional ByVal CrLf As Boolean = True)
   On Error Resume Next
   Text3.SelStart = Len(Text3.Text)
   If CrLf Then
      Debug.Print Data
      Text3.SelStart = Len(Text3.Text)
      Text3.SelText = Data & vbCrLf
   Else
      Debug.Print Data;
      Text3.SelText = Data
   End If
End Sub

Private Sub Text1_GotFocus()
   Call SelectAll(Text1)
End Sub

Private Sub Text2_GotFocus()
   Call SelectAll(Text2)
End Sub

' *********************************************
'  Private Methods
' *********************************************
Private Sub SelectAll(txt As TextBox)
   txt.SelStart = 0
   txt.SelLength = Len(txt.Text)
End Sub

