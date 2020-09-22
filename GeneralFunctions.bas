Attribute VB_Name = "GeneralFunctions"
Option Explicit
  Public FSO As FileSystemObject
  Public T As TextStream
  Public J As Integer
  Public LastUpdate As String
  
Public Sub waitFor(ByVal mSecond As Single)
  Dim Start, Finish As Single
  Start = Timer ' Set start time.
  Do While Timer < Start + mSecond
    DoEvents  ' Yield to other processes.
  Loop
  Finish = Timer  ' Set end time.
End Sub

Public Sub WriteHTM()
'USED THIS ROUTINE TO GET ALL THE AVAILABLE PICTURES FORM WEATHERCHANNEL.COM
  Set FSO = New FileSystemObject
  Set T = FSO.OpenTextFile(App.Path & "\gifs.htm", ForWriting, True)
  For J = 1 To 200
    T.Write "<IMG SRC=http://www.intellicast.com/images/icons/" & J & "_wtext.jpg>&nbsp" & J & "<br>"
  Next J
  T.Close
  Set T = Nothing
  Set FSO = Nothing
End Sub

Public Sub WriteUpdateLOG(mUpdate As String)
  Set FSO = New FileSystemObject
  Set T = FSO.OpenTextFile(App.Path & "\Update History.log", ForAppending, True)
  T.WriteLine mUpdate
  T.Close
  Set T = Nothing
  Set FSO = Nothing
End Sub
