Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit

Private WithEvents dd As CDirDrill
Private m_Cancel As Boolean
Private Sub Form_Load()
   ' Enable watching for an escape key.
   Me.KeyPreview = True
   ' Put some default text into controls.
   Text1.Text = App.Path
   ' Instantiate recursive search class.
   Set dd = New CDirDrill
End Sub


Private Sub dd_NewFile(ByVal filespec As String, Cancel As Boolean)
   ' Show that we found one...
   Debug.Print dd.ExtractPath(filespec) & dd.ExtractName(filespec)
       Dim lRet&
    Debug.Print "C:\GALLERY\PROGRAMS\CORELGAL.EXE " & dd.ExtractPath(filespec) & dd.ExtractName(filespec)
' Start the Corel application, and store the process id.
lRet = Shell("C:\GALLERY\PROGRAMS\CORELGAL.EXE " & dd.ExtractPath(filespec) & dd.ExtractName(filespec), vbNormalFocus)

' Send the keystrokes to the Corel application.
SendKeys ("{ENTER}"), True
SendKeys ("{ENTER}"), True
SendKeys ("{ENTER}"), True
SendKeys ("{ENTER}"), True
Sleep 300
SendKeys ("%f"), True
Sleep 600
SendKeys ("e"), True
'Sleep 1000
SendKeys ("{ENTER}"), True
Sleep 1000
SendKeys ("{ENTER}"), True
Sleep 1000
SendKeys ("%{F4}"), True
Dim FileName As String
Kill dd.ExtractPath(filespec) & dd.ExtractName(filespec)
Debug.Print "Delete " & dd.ExtractPath(filespec) & dd.ExtractName(filespec)
End Sub

Private Sub Button1_Click()
    Text1.Text = "SoF is the greatest!"
    Dim ProcID As Double
End Sub

Private Sub Command1_Click()
   dd.Folder = "C:\GALLERY\PROGRAMS\CLIPART"
   dd.Pattern = "*.BMF"
   dd.AttributeMask = vbHidden Or vbSystem Or vbArchive Or vbReadOnly
   ' Clear cancel flag.
   m_Cancel = False
   ' Let it rip!
   dd.BeginSearch
End Sub

Private Sub Go_Click()

End Sub
