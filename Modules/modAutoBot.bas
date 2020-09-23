Attribute VB_Name = "modAutoBot"
Option Compare Text

Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Is Notch Running?
Global NotchRunning As Boolean
' Is Notch Learning?
Public NotchLearning As Boolean
Public LearningFrom As String
Public LearningWord As String

' Allows contextual sentence arranging
Public ContextualSentences As Boolean
' This determines when a random sentence is strung together.
' It's in Percentage form
Public ContextualRandomness As Integer
Global NewWords As Integer
' EnumSect lets us enumerate our INI section names
' and in return get the number of sections, so Notch
' can work out how many questions he has
Public Function EnumSect() As Integer
    Dim szBuf As String, Length As Integer
    Dim SectionArr() As String
    szBuf = String$(255, 0)
    Length = GetPrivateProfileSectionNames(szBuf, 255, IniFile(2))
    szBuf = Left$(szBuf, Length)
    SectionArr = Split(szBuf, vbNullChar)
    EnumSect = UBound(SectionArr)
End Function

Public Sub DoAutoBot(Question As String)

Dim BTmp As String

If NotchRunning = False Or LastFrom = "Notch" Then Exit Sub

TotPhrases = EnumSect - 2

For i = 1 To TotPhrases

BTmp = ReadText("Phrase" & i, "Question", 2)
Dim MultiQ() As String
BTmp = BTmp & "||"
MultiQ = SplitVB5(BTmp, "||")

For n = 0 To UBound(MultiQ)
'If Result(1) = "pm1" And InStr(Len(3 + "ø" + Len("N©H@-|-")), Question, "pm1") <= 0 Then Exit Sub


' WyldCard is our Wild Card character matching engine.
' A lot smaller and easier to understand than
' using the Regular Expressions library
If WyldCard(Question, MultiQ(n)) = True And Result(4) <> "Notch" Then

' This dirty little hack lets us 'Enum' the number
' of keys in our section. Only works if key starts
' with 'Answer'
Dim F As Integer
F = 1

Do Until ReadText("Phrase" & i, "Answer" & F, 2) = ""
F = F + 1
Loop

' Pick a random phrase. NEW HACK: Notch SHOULD not display blank phrases now.
' Loops until a valid message has been found. Usually 1-2 loops
Randomize
R = Int(Rnd * F) + 1
Do Until R > 0 And R <= F
R = Int(Rnd * F) + 1

Loop


If ReadText("Phrase" & i, "Broadcast", 2) = "True" Or ReadText("Phrase" & i, "Broadcast", 2) = "" Then
Broadcast ReadText("Phrase" & i, "Answer" & R, 2)
ElseIf ReadText("Phrase" & i, "Broadcast", 2) = "False" Then
sckUDP.SendData ReadText("Phrase" & i, "Answer" & R, 2)
End If

Exit Sub
End If
Next n
Next i

If NotchLearning = True And Result(1) = "msg" Or Result(1) = "pm1" And LastFrom <> "Notch" And LastFrom <> "" Then Try2Learn
Exit Sub

End Sub

Public Function Try2Learn()

' OK This is Notch's learning center. If someone sends a message to Notch, (msg or pm1)
' Notch is learning (NotchLearning = True), and he doesn't have the question
' in his DB, then he will learn it. He asks for a response, then when
' one is provided, he will let them know, write it into the DB, and reload
' the words to be used. Kinda simple... :|

' BTW, when Notch is learning. Only one person can teach at a time
' ie if Grayda is trying to teach Notch something, and he accidentally
' says 'Hi' to someone else, then Notch will interpret that as the response
' to his learning question. (Eg. Hi Notch. Response is: Hello dude!)

' Notch can't learn from himself, it would create
' a never ending loop!! (?)

If LastFrom = "Notch" Then Exit Function
If LearningWord = Result(3) Then Exit Function

' If we are currently learning something, then don't learn something else
If Trim(LearningFrom) = "" And WyldCard(Result(3), "Notch") = True Then
LearningWord = Result(3)
LearningFrom = LastFrom

WriteSect "Phrase" & EnumSect - 1, "Question=" & Result(3), 2
If Result(1) = "msg" Then
Broadcast "msgøNotchø+lastfrom+, I don't understand your question. Please provide me with a response!ø0øFalseøFalseø0"
' Double DoEvents so our message has time to reach the recipient before
' trying to do a "DoAutoBot" again
DoEvents
DoEvents

ElseIf Result(1) = "pm1" Then
Broadcast "pm1ø+lastfrom+ø+lastfrom+, I don't understand your question. Please provide me with a response!øNotch"
' Double DoEvents so our message has time to reach the recipient before
' trying to do a "DoAutoBot" again
DoEvents
DoEvents

End If
' Wait until we get a response
Exit Function
' Not learning anything (LearningFrom is blank)? then start learning!
ElseIf LearningFrom <> LastFrom Then

Exit Function
ElseIf LastFrom = LearningFrom Then
TotPhrases = EnumSect - 2


If Result(1) = "msg" Then
Broadcast "msgøNotchøThanks +lastfrom+! I now know what you are talking about!ø0øFalseøFalseø0"
' Double DoEvents so our message has time to reach the recipient before
' trying to do a "DoAutoBot" again
DoEvents
DoEvents

WriteString "Phrase" & TotPhrases, "Answer1", "msgøNotchø" & Result(3) & "ø0øFalseøFalseø0", 2
ElseIf Result(1) = "pm1" Then
Broadcast "pm1ø+lastfrom+øThanks +lastfrom+! I now know what you are talking about!øNotch"
' Double DoEvents so our message has time to reach the recipient before
' trying to do a "DoAutoBot" again
DoEvents
DoEvents

WriteString "Phrase" & TotPhrases, "Answer1", "pm1ø+lastfrom+ø" & Result(3) & "øNotch", 2
End If

NewWords = NewWords + 1
LearningFrom = ""
' Reload our Notch, will NEW words included!
OldINI = IniFile(2)
IniFile(2) = ""
IniFile(2) = OldINI
Exit Function
End If


End Function

' This is my custom WildCard system.
' It lets you use *'s to search for items
' Here is how it works:

' The Search Criteria (EG: Testing*Hello)
'   is SplitVB5 into an array.
' The first part (EG: Testing) is searched for
'   using the InStr command. If it is found, then
'   1 is added to the number of correct matches.
'   If it isn't found, nothing changes
' At the end, if the Number of matches is equal
'   to the number of items in the array, then
'   it returns as true, if not, then false

' If you can improve on this, then please let me know,
' by sending an e-mail to: firestorm_visual@hotmail.com
' but this fits my needs perfectly, so I don't think
' I'll improve on it :)


Public Function WyldCard(StringToCheck As String, SearchFor As String) As Boolean
On Error Resume Next

Dim Search() As String
Dim Matches As Integer
SearchFor = SearchFor & "*"
Search = SplitVB5(SearchFor, "*")

For n = LBound(Search) To UBound(Search)

RetPos = InStr(1, StringToCheck, Search(n), vbTextCompare)

If RetPos > 0 Then
Matches = Matches + 1
End If

Next n

If Matches = UBound(Search) + 1 Then
WyldCard = True
Else
WyldCard = False
End If


End Function


