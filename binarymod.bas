Attribute VB_Name = "Module1"
Option Explicit

Public Type FileStructure
    Warning As String ''FOR WARNING MESSAGE THAT IS TO BE PUT IN SAVED FILE
    RTBtext As String ''FOR RTB
    LBLtext As Long ''FOR LABEL; CAN ALSO BE A STRING IF NECESSARY
    LSTitems As String ''FOR LISTBOX
    LSTitems1 As String ''FOR LISTBOX
End Type

Public loopopen As Long ''FOR THE LISTBOXLOOP-OPEN
Public generalS As String ''FOR GENERAL -STRING- USAGE
Public generalL As Long ''FOR GENERAL -LONG- USAGE
Public FileRec As FileStructure ''FOR RECORD PLACEMENT


Public Sub SaveFile()

Open (Form1.CommonDialog1.FileName) For Binary As #1 ''OPENS FILE; IF THE FILE DOES NOT EXIST THEN IT IS CREATED
Close #1 ''CLOSES FILE
Kill (Form1.CommonDialog1.FileName)
''^^WE MUST FIRST OPEN/CREATE THE FILE TO BE OPERATED ON _
THEN IMMEDIATELY KILL(DELETE) IT. IF THE FILE ALREADY _
EXISTS AND WE DO NOT KILL IT FIRST THEN THE NEW RECORDS _
WILL BE ADDED TO THE FILE LEAVING THE EXISTING RECORDS IN _
THE FILE(eventhough they cannot be accessed) WHICH MAY _
CAUSE DISK SPACE TO BE WASTED. DOING IT THIS WAY WILL _
ENSURE THAT THE ONLY RECORD IN THE FILE IS THE RECORD _
WE PUT INTO IT IN THIS SUB TO KEEP THE SAVED FILE SIZE _
MINIMAL.^^


Open (Form1.CommonDialog1.FileName) For Binary As #1 ''OPENS FILE; IF THE FILE DOES NOT EXIST IT IS CREATED

FileRec.Warning = ("Binary Parsing Engine By: Michael Schmidt (mds@vci.net)|2-26-2000|2:32AM" & Chr(10) & "WARNING!!! EDIT THIS FILE AND YOU WILL NO LONGER BE ABLE TO OPEN IT WITH THE ASSOCIATED APPLICATION!!" & Chr(10)) ''''ASSIGN VALUE TO THE FIELD THAT IS GOING TO BE PUT IN THE RECORD; ONLY SHOWS UP IN BINARY FILE INCASE PEOPLE TRY TO EDIT THE FILE
FileRec.RTBtext = Form1.RichTextBox1 ''ASSIGN VALUE TO THE FIELD THAT IS GOING TO BE PUT IN THE RECORD
FileRec.LBLtext = Form1.Label1 ''''ASSIGN VALUE TO THE FIELD THAT IS GOING TO BE PUT IN THE RECORD

generalS = "" ''CLEARS VARIABLE MUST DO!
For generalL = 0 To Form1.List1.ListCount - 1 ''SETS STARTING/FINISHING VALUES OF LOOP
 generalS = generalS & Form1.List1.List(generalL) & Chr(10) ''ADDS LIST1 CONTENTS TO STRING WITH (chr(10)) BETWEEN EACH
  Next generalL ''RESTARTS LOOP
FileRec.LSTitems = generalS ''ASSIGN VALUE TO THE FIELD THAT IS GOING TO BE PUT IN THE RECORD

generalS = "" ''CLEARS VARIABLE MUST DO!
For generalL = 0 To Form1.List2.ListCount - 1 ''SETS STARTING/FINISHING VALUES OF LOOP
 generalS = generalS & Form1.List2.List(generalL) & Chr(10) ''ADDS LIST2 CONTENTS TO STRING WITH (chr(10) BETWEEN EACH
  Next generalL ''RESTARTS LOOP
FileRec.LSTitems1 = generalS ''ASSIGN VALUE TO THE FIELD THAT IS GOING TO BE PUT IN THE RECORD

Put #1, 1, FileRec ''WRITES RECORD TO FILE
Close #1 ''CLOSES FILE

End Sub

Public Sub OpenFile()

Open (Form1.CommonDialog1.FileName) For Binary As #1 'OPENS FILE
Get #1, 1, FileRec ''READS INFO FROM FILE
Close #1 ''CLOSES FILE

Form1.RichTextBox2.Text = "" ''CLEARS RTB BEFORE FILE IS OPENED IN IT
Form1.Label4.Caption = "0" ''CLEARS LABEL BEFORE FILE IS OPENED IN IT
Form1.RichTextBox2 = FileRec.RTBtext ''FILLS RTB WITH INFO FROM FILE
Form1.Label4 = FileRec.LBLtext ''FILLS LABEL WITH INFO FROM FILE

generalS = FileRec.LSTitems ''ASSIGNS RECORD TO A STRING(FROM FILE)
Form1.List3.Clear ''CLEARS LIST BEFORE FILE IS OPENED IN IT
loopopen = 1 ''ASSIGNS VALUE TO LOOPOPEN FOR LOOP
 For generalL = 1 To Len(generalS) ''SETS STARTING/FINISHING VALUES OF LOOP
  If Mid(generalS, generalL, 1) = Chr(10) Then ''SEARCHES STRING FOR (chr(10); IF FOUND THEN ADDS THE TEXT DIRECTLY BEFORE (chr(10) TO LIST3
    Form1.List3.AddItem Mid(generalS, loopopen, (generalL - loopopen))
     loopopen = generalL + 1 ''MOVES TO NEXT ITEM IN LOOP
  End If
    Next generalL ''RESTARTS LOOP

generalS = FileRec.LSTitems1 ''ASSIGNS RECORD TO A STRING(FROM FILE)
Form1.List4.Clear ''CLEARS LIST BEFORE FILE IS OPENED IN IT
loopopen = 1 ''ASSIGNS VALUE TO LOOPOPEN FOR LOOP
 For generalL = 1 To Len(generalS) ''SETS STARTING/FINISHING VALUES OF LOOP
  If Mid(generalS, generalL, 1) = Chr(10) Then ''SEARCHES STRING FOR (chr(10); IF FOUND THEN ADDS THE TEXT DIRECTLY BEFORE (chr(10) TO LIST4
    Form1.List4.AddItem Mid(generalS, loopopen, (generalL - loopopen))
      loopopen = generalL + 1 ''MOVES TO NEXT ITEM IN LOOP
  End If
    Next generalL ''RESTARTS LOOP

End Sub


Public Function ShowSave(DefaultDir As String, DefaultName As String) As String

With Form1.CommonDialog1
    On Error GoTo errhandler
    .CancelError = True
    .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
    .Filter = "Any File Extension (*.any)|*.any" ''CAN BE ANY EXTENSION
    .FilterIndex = 0
    .InitDir = DefaultDir
    .FileName = DefaultName
    .DialogTitle = "Save As"
    .ShowSave
End With
    Call SaveFile ''CALLS SaveFile SUB
errhandler:
    End Function
''^^COMMONDIALOG SHOWSAVE^^

Public Function ShowOpen(DefaultDir As String, DefaultName As String) As String

With Form1.CommonDialog1
    On Error GoTo errhandler
    .CancelError = True
    .Flags = cdlOFNHideReadOnly
    .Filter = "Any File Extension (*.any)|*.any" '' = SAME EXTENSION AS SAVED FILE
    .FilterIndex = 0
    .InitDir = DefaultDir
    .FileName = DefaultName
    .DialogTitle = "Open File"
    .ShowOpen
End With
    Call OpenFile ''CALLS OpenFile SUB
errhandler:
End Function
''COMMONDIALOG SHOWOPEN^^
