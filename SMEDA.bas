Attribute VB_Name = "Module11"
'
' ProfessorF's Twitter Scraper
'
' Copyright (c) Nick V. Flor, 2014-2016, All rights reserved
'
' This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
' CC BY-SA => If you use this code for research, you must cite me in your paper references
'
' To view a copy of the license visit: http://creativecommons.org/licenses/by-sa/4.0/legalcode
' To view a summary of the license visit: http://creativecommons.org/licenses/by-sa/4.0/
'
' This material is based partly upon work supported by the National Science Foundation (NSF)
' under both ECCS - 1231046 and CMMI - 1635334. Any opinions, findings, and conclusions or recommendations expressed
' in this material are those of the author and do not necessarily reflect the views of the NSF.
'
' If you're not familiar with the Twitter API, you really shouldn't be looking
' ... at this code!
'
' History: 08Oct2016 10:02PM
'          16Jun2017 08:34PM - Bug fixes, added functions for viral research: countRTs, findDuplicateNames
'                              Note: many as Integers still need to be changed to as Long
'          19Jun2017 09:10AM - Added sub convertTwitterDateToExcelDate, useful for viral time deltas
'          20Jun2017 04:50PM - Added a function to check friendship, changed output of getAll for verified, geoenabled, hashtags
'          28Jun2017 05:07PM - Added a prototype getAllExtended
'          28Jun2017 10:19PM - Fixed bug in getAllExtended
'          29Jun2017 07:31AM - Fixed bug in getAllExtended, getRTs now displays URL of RT for easy access
'          29Jun2017 10:10PM - Fixed bug in countRT
'          30Jun2017 08:17AM - Fixed bug in genSocialEdges by creating getRTNameRegex
'          18Jul2017 05:27PM - Fixed GetAllExtendedSlowly (now works)
'          03Feb2018 01:57PM - Fixed crash when a user has all numbers as a name
'          17Feb2018 06:47PM - Created removeStopwords subroutine & regexMatch function
'          17Feb2018 09:06PM - Created a DocumentTermMatrix subroutine
'          24Mar2018 11:39AM - Added support for NGrams
'
Option Explicit
' IMPORTANT: YOU MUST OBTAIN CONSUMER KEY AND SECRET FROM TWITTER DEVELOPER ACCOUNT
Public Const consumer_key As String = ""
Public Const consumer_secret As String = ""
Public bearer_token As String
Public Const HINTERVAL As Long = 60 ' Bin size for histograms in minutes
Dim dpwords As New Dictionary ' positive sentiment words
Dim dnwords As New Dictionary ' negative sentiment words
Dim dswords As New Dictionary ' stop words
Dim dgwords As New Dictionary ' ngrams

'
' TwitterLogin: Logs you into twitter
' Entry:        consumer_key, consumer_secret SET ABOVE, get from Twitter
' Exit:         B2 contains "bearer_token" for the session
'
Private Sub TwitterLogin()
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As ScriptControl
Dim url As String
Dim credential As String
Dim byte_credential() As Byte
Dim base64_credential As String
Dim auth_head_val As String
Dim objXML As New MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement
Dim response As String
Dim json As Object


url = "https://api.twitter.com/oauth2/token"
credential = consumer_key + ":" + consumer_secret
byte_credential = StrConv(credential, vbFromUnicode)

Set objNode = objXML.createElement("b64")
objNode.DataType = "bin.base64"
objNode.nodeTypedValue = byte_credential
base64_credential = Replace(objNode.text, Chr(10), "") 'grr, after 72 chars, adds LF

auth_head_val = "Basic " + base64_credential
req.Open "POST", url, False
req.setRequestHeader "Authorization", auth_head_val
req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
req.send ("grant_type=client_credentials")

While req.readyState <> 4
    DoEvents
Wend
response = req.responseText
'
' Create Script Control
'
Set sc = CreateObjectx86("MSScriptControl.ScriptControl")

sc.Language = "javascript"
Set json = sc.Eval("(" + response + ")")


bearer_token = json.access_token
Range("B2") = bearer_token
'
' Close Script Control
'
CreateObjectx86 , True ' close mshta host window at the end
MsgBox ("Bearer Token: " + bearer_token)
End Sub
Private Function getRTName(t As String)
Dim colonloc, endloc As Long
Dim rts

rts = Mid(t, 1, 3)
If (rts = "RT ") Then
    colonloc = InStr(4, t, ":")
    If (colonloc <> 0) Then
        endloc = colonloc - 1
        getRTName = Mid(t, 5, endloc - 4)
    Else
        getRTName = ""
    End If
Else
    getRTName = ""
End If
End Function
Private Function getRTNameRegEx(t As String) 'regex version of above, which has a bug
Dim regex As New RegExp
Dim mc As MatchCollection
Dim m As Match
Dim n As String

regex.Pattern = "^RT @([^\s:]+)[\s:].*"
Set mc = regex.Execute(t)
n = ""
If mc.count > 0 Then
    Set m = mc(0)
    n = m.SubMatches(0)
End If
getRTNameRegEx = n
End Function
'
' getAll: Get All Tweets (as fast as possible)
'
' Entry:  Twitter search string in A1
'         Columns B & M Must be Text
' Exit:   Up to 45,000 Tweets scraped per the search string
'         If more than 45,000 tweets use GetAllSlowly to scrape remaining
'
Sub getAll()
Attribute getAll.VB_ProcData.VB_Invoke_Func = " \n14"
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As ScriptControl
Dim url As String
Dim response As String
Dim authorization_value As String
Dim parameters As String
Dim dnum As Long
Dim dcell As String
Dim last_id As String
Dim max_id As String
Dim x As Object
Dim text As String ' hack to prevent vba from capitalizing text to Text
Dim id As Long
Dim count As Long
Dim iname As Long
Dim r As String
Dim c As Long
Dim json As Object
Dim location 'hack to prevent vba from capitalizing location as Location
Dim zed As Long
Dim mentions, hashtags, medias As String
Dim en, um, uh, uu As Variant
Dim scount As Long
Dim stime As String
'
' Login if necesary
'
If Range("b2") = "" Then TwitterLogin
bearer_token = Range("b2")
authorization_value = "Bearer " + bearer_token ' assumes getTwitterToken called

'
' Create Script Control
'
Set sc = CreateObjectx86("MSScriptControl.ScriptControl")

'
' The Starting Cell of the Output
'
dnum = 3
dcell = "B"

'
' Initialize the tweet max_id and the tweet count to zero
'
max_id = "0"
Range("a3") = 0

'
' Main Loop
'
url = "https://api.twitter.com/1.1/search/tweets.json"
scount = 0 'status variables for long scrapes
stime = ""
Do
    '
    ' Set up parameters
    '
    last_id = max_id ' remember the very last tweet
    If (max_id = "0") Then
        parameters = Range("a1") + "&count=100" ' 100 is the max tweets you can grab
    Else
        parameters = Range("a1") + "&count=100&" + "max_id=" + max_id
    End If
    '
    ' Send the search request to twitter
    '
    req.Open "GET", url + "?" + parameters, False
    req.setRequestHeader "Authorization", authorization_value
    req.send
    '
    ' Wait until all the data is sent back
    '
    While req.readyState <> 4 ' #4 means all data received
        DoEvents
    Wend
    '
    ' Get the response object and put it into the json variable
    '
    response = req.responseText
    ' Range("a2") = response ' debug to check resposnes
    sc.Language = "javascript"
    Set json = sc.Eval("(" + response + ")")
'    Range("a2") = response ' debugging
    '
    ' Loop through every tweet in the json
    '
    For Each x In json.statuses
        If (max_id <> x.id_str) Then
            Range("B" + CStr(dnum)) = CStr(x.id_str)
            Range("C" + CStr(dnum)) = x.created_at
            Range("D" + CStr(dnum)) = x.user.screen_name
            Range("E" + CStr(dnum)) = x.user.created_at
            Range("F" + CStr(dnum)) = x.user.statuses_count
            Range("G" + CStr(dnum)) = x.user.favourites_count
            Range("H" + CStr(dnum)) = x.user.followers_count
            Range("I" + CStr(dnum)) = x.user.friends_count
            Range("J" + CStr(dnum)) = x.user.listed_count
            Range("K" + CStr(dnum)) = x.favorite_count
            Range("L" + CStr(dnum)) = x.retweet_count
            Range("M" + CStr(dnum)) = x.user.verified

            Range("N" + CStr(dnum)) = x.user.geo_enabled
            
            Range("O" + CStr(dnum)) = x.lang
            
            If (InStr(x.text, "=") <> 1) Then
                Range("P" + CStr(dnum)) = x.text
            Else
                Range("P" + CStr(dnum)) = "!" + x.text
            End If
            
            hashtags = ""
            For Each uh In x.entities.hashtags
                If (hashtags <> "") Then hashtags = hashtags + ","
                hashtags = hashtags + uh.text
            Next
            Range("Q" + CStr(dnum)) = hashtags

            mentions = ""
            For Each um In x.entities.user_mentions
                If (mentions <> "") Then mentions = mentions + ","
                mentions = mentions + um.screen_name
            Next
            Range("R" + CStr(dnum)) = mentions
            
            dnum = dnum + 1
            max_id = x.id_str
            If (stime <> CStr(x.created_at)) Then
                scount = 0
                stime = CStr(x.created_at)
            Else
                scount = scount + 1
            End If
            Application.StatusBar = CStr(dnum) + ":" + CStr(scount) + ":" + CStr(x.created_at)
        End If
        Range("A3") = dnum - 3
        DoEvents
    Next
    
    'sleepNow (2) ' If you don't want to use getAllSlowly, uncomment this, but unreliable
Loop Until (json.statuses = "" Or max_id = last_id)

CreateObjectx86 , True ' close mshta host window at the end
MsgBox "Get All Done"
End Sub
'
' getAllSlowly: Get All Tweets, but slowly to get around rate limit
'               Run this ONLY AFTER AND IF getAll crashes
'
' Entry:  Twitter search string in A1
'         Columns B & M Must be Text
'         Cursor positioned on last valid B-cell
' Exit:   All the tweets, subject to Excel's max row
'
Sub getAllSlowly()
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As ScriptControl
Dim url As String
Dim response As String
Dim authorization_value As String
Dim parameters As String
Dim dnum As Long
Dim dcell As String
Dim x As Object
Dim text As String ' hack to prevent vba from capitalizing text to Text
Dim id As Long
Dim count As Long
Dim iname As Long
Dim r As String
Dim c As Long
Dim max_id As String
Dim last_id As String
Dim json As Object
Dim scount As Long
Dim stime As String
Dim mentions, hashtags, medias As String
Dim um, uh, uu As Variant

'
' Login if necesary
'
If Range("b2") = "" Then TwitterLogin
bearer_token = Range("b2")
authorization_value = "Bearer " + bearer_token ' assumes getTwitterToken called

'
' Create Scripting object
'
Set sc = CreateObjectx86("MSScriptControl.ScriptControl")

'
' The Starting Cell of the Output
'
dnum = 3
dcell = "B"

'
' Initialize the tweet max_id and the tweet count to zero
'
max_id = Selection
dnum = ActiveCell.row + 1
'
' Main Loop
'
url = "https://api.twitter.com/1.1/search/tweets.json"
scount = 0 'status variables for long scrapes
stime = ""
Do
    '
    ' Set up parameters
    '

    last_id = max_id ' remember the very last tweet
    If (max_id = "0") Then
        parameters = Range("a1") + "&count=100" ' 100 is the max tweets you can grab
    Else
        parameters = Range("a1") + "&count=100&" + "max_id=" + max_id
    End If
    '
    ' Send the search request to twitter
    '
    req.Open "GET", url + "?" + parameters, False
    req.setRequestHeader "Authorization", authorization_value
    req.send
    '
    ' Wait until all the data is sent back
    '
    While req.readyState <> 4 ' #4 means all data received
        DoEvents
    Wend
    '
    ' Get the response object and put it into the json variable
    '
    response = req.responseText
    sc.Language = "javascript"
    Set json = sc.Eval("(" + response + ")")
    '
    ' Loop through every tweet in the json
    '
    For Each x In json.statuses
        If (max_id <> x.id_str) Then
            ' Same as GetAll (should be a function--future rev)
            Range("B" + CStr(dnum)) = CStr(x.id_str)
            Range("C" + CStr(dnum)) = x.created_at
            Range("D" + CStr(dnum)) = x.user.screen_name
            Range("E" + CStr(dnum)) = x.user.created_at
            Range("F" + CStr(dnum)) = x.user.statuses_count
            Range("G" + CStr(dnum)) = x.user.favourites_count
            Range("H" + CStr(dnum)) = x.user.followers_count
            Range("I" + CStr(dnum)) = x.user.friends_count
            Range("J" + CStr(dnum)) = x.user.listed_count
            Range("K" + CStr(dnum)) = x.favorite_count
            Range("L" + CStr(dnum)) = x.retweet_count
            Range("M" + CStr(dnum)) = x.user.verified
            Range("N" + CStr(dnum)) = x.user.geo_enabled
            
            Range("O" + CStr(dnum)) = x.lang
            If (InStr(x.text, "=") <> 1) Then
                Range("P" + CStr(dnum)) = x.text
            Else
                Range("P" + CStr(dnum)) = "!" + x.text
            End If
            
            hashtags = ""
            For Each uh In x.entities.hashtags
                If (hashtags <> "") Then hashtags = hashtags + ","
                hashtags = hashtags + uh.text
            Next
            
            Range("Q" + CStr(dnum)) = hashtags

            mentions = ""
            For Each um In x.entities.user_mentions
                If (mentions <> "") Then mentions = mentions + ","
                mentions = mentions + um.screen_name
            Next
            Range("R" + CStr(dnum)) = mentions
            
            dnum = dnum + 1
            max_id = x.id_str
            If (stime <> CStr(x.created_at)) Then
                scount = 0
                stime = CStr(x.created_at)
            Else
                scount = scount + 1
            End If
            Application.StatusBar = CStr(dnum) + ":" + CStr(scount) + ":" + CStr(x.created_at)
        End If
        Range("A3") = dnum - 3
        DoEvents
    Next
    
    sleepNow (2)
    
Loop Until (json.statuses = "" Or max_id = last_id)

CreateObjectx86 , True ' close mshta host window at the end

MsgBox "Get All Done"
End Sub
'
' getAllExtended: Get All Tweets (as fast as possible), extended version
'
' Entry:  Twitter search string in A1
'         Columns B & M Must be Text
' Exit:   Up to 45,000 Tweets scraped per the search string
'         If more than 45,000 tweets use GetAllSlowly to scrape remaining
'
Sub getAllExtended()
Attribute getAllExtended.VB_ProcData.VB_Invoke_Func = "e\n14"
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As ScriptControl
Dim url As String
Dim response As String
Dim authorization_value As String
Dim parameters As String
Dim dnum As Long
Dim dcell As String
Dim last_id As String
Dim max_id As String
Dim x As Object
Dim text As String ' hack to prevent vba from capitalizing text to Text
Dim id As Long
Dim count As Long
Dim iname As Long
Dim r As String
Dim c As Long
Dim json As Object
Dim location 'hack to prevent vba from capitalizing location as Location
Dim zed As Long
Dim mentions, hashtags, medias As String
Dim en, um, uh, uu As Variant
Dim scount As Long
Dim stime As String
'
' Login if necesary
'
If Range("b2") = "" Then TwitterLogin
bearer_token = Range("b2")
authorization_value = "Bearer " + bearer_token ' assumes getTwitterToken called

'
' Create Script Control
'
Set sc = CreateObjectx86("MSScriptControl.ScriptControl")

'
' The Starting Cell of the Output
'
dnum = 3
dcell = "B"

'
' Initialize the tweet max_id and the tweet count to zero
'
max_id = "0"
Range("a3") = 0

'
' Main Loop
'
url = "https://api.twitter.com/1.1/search/tweets.json"
scount = 0 'status variables for long scrapes
stime = ""
Do
    '
    ' Set up parameters, add tweet_mode=extended
    '
    last_id = max_id ' remember the very last tweet
    If (max_id = "0") Then
        parameters = Range("a1") + "&tweet_mode=extended" + "&count=100" ' 100 is the max tweets you can grab
    Else
        parameters = Range("a1") + "&tweet_mode=extended" + "&count=100&" + "max_id=" + max_id
    End If
    '
    ' Send the search request to twitter
    '
    req.Open "GET", url + "?" + parameters, False
    req.setRequestHeader "Authorization", authorization_value
    req.send
    '
    ' Wait until all the data is sent back
    '
    While req.readyState <> 4 ' #4 means all data received
        DoEvents
    Wend
    '
    ' Get the response object and put it into the json variable
    '
    response = req.responseText
    ' Range("a2") = response ' debug to check resposnes
    sc.Language = "javascript"
    Set json = sc.Eval("(" + response + ")")
    '
    ' Loop through every tweet in the json
    '
    For Each x In json.statuses
        If (max_id <> x.id_str) Then
            Range("B" + CStr(dnum)) = CStr(x.id_str)
            Range("C" + CStr(dnum)) = x.created_at
            Range("D" + CStr(dnum)) = x.user.screen_name
            Range("E" + CStr(dnum)) = x.user.created_at
            Range("F" + CStr(dnum)) = x.user.statuses_count
            Range("G" + CStr(dnum)) = x.user.favourites_count
            Range("H" + CStr(dnum)) = x.user.followers_count
            Range("I" + CStr(dnum)) = x.user.friends_count
            Range("J" + CStr(dnum)) = x.user.listed_count
            Range("K" + CStr(dnum)) = x.favorite_count
            Range("L" + CStr(dnum)) = x.retweet_count
            Range("M" + CStr(dnum)) = x.user.verified

            Range("N" + CStr(dnum)) = x.user.geo_enabled
            'medias = ""
            'For Each uu In x.extended_entities.media
            '    If (medias <> "") Then medias = medias + ","
            '    medias = medias + uu.media_url_https
            'Next
            'Range("N" + CStr(dnum)) = medias
            
            Range("O" + CStr(dnum)) = x.lang
            
            If (InStr(x.full_text, "=") <> 1) Then
                Range("P" + CStr(dnum)) = x.full_text
            Else
                Range("P" + CStr(dnum)) = "!" + x.full_text
            End If
            
            hashtags = ""
            For Each uh In x.entities.hashtags
                If (hashtags <> "") Then hashtags = hashtags + ","
                hashtags = hashtags + uh.text
            Next
            Range("Q" + CStr(dnum)) = hashtags

            mentions = ""
            For Each um In x.entities.user_mentions
                If (mentions <> "") Then mentions = mentions + ","
                mentions = mentions + um.screen_name
            Next
            Range("R" + CStr(dnum)) = mentions
            
            dnum = dnum + 1
            max_id = x.id_str
            If (stime <> CStr(x.created_at)) Then
                scount = 0
                stime = CStr(x.created_at)
            Else
                scount = scount + 1
            End If
            Application.StatusBar = CStr(dnum) + ":" + CStr(scount) + ":" + CStr(x.created_at)
        End If
        Range("A3") = dnum - 3
        DoEvents
    Next
    
    'sleepNow (2) ' If you don't want to use getAllSlowly, uncomment this, but unreliable
Loop Until (json.statuses = "" Or max_id = last_id)

CreateObjectx86 , True ' close mshta host window at the end
MsgBox "Get All Done"
End Sub
'
' getAllExtendedSlowly: Get All Tweets, but slowly to get around rate limit
'               Run this ONLY AFTER AND IF getAll crashes
'
' Entry:  Twitter search string in A1
'         Column B Must be Text
'         Cursor positioned on last valid B-cell
' Exit:   All the tweets, subject to Excel's max row
'
Sub getAllExtendedSlowly()
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As ScriptControl
Dim url As String
Dim response As String
Dim authorization_value As String
Dim parameters As String
Dim dnum As Long
Dim dcell As String
Dim last_id As String
Dim max_id As String
Dim x As Object
Dim text As String ' hack to prevent vba from capitalizing text to Text
Dim id As Long
Dim count As Long
Dim iname As Long
Dim r As String
Dim c As Long
Dim json As Object
Dim location 'hack to prevent vba from capitalizing location as Location
Dim zed As Long
Dim mentions, hashtags, medias As String
Dim en, um, uh, uu As Variant
Dim scount As Long
Dim stime As String
'
' Login if necesary
'
If Range("b2") = "" Then TwitterLogin
bearer_token = Range("b2")
authorization_value = "Bearer " + bearer_token ' assumes getTwitterToken called

'
' Create Script Control
'
Set sc = CreateObjectx86("MSScriptControl.ScriptControl")

'
' The Starting Cell of the Output
'
dnum = 3
dcell = "B"

'
' Initialize the tweet max_id and the tweet count to zero
'
max_id = Selection
dnum = ActiveCell.row + 1

'
' Main Loop
'
url = "https://api.twitter.com/1.1/search/tweets.json"
scount = 0 'status variables for long scrapes
stime = ""
Do
    '
    ' Set up parameters, add tweet_mode=extended
    '
    last_id = max_id ' remember the very last tweet
    If (max_id = "0") Then
        parameters = Range("a1") + "&tweet_mode=extended" + "&count=100" ' 100 is the max tweets you can grab
    Else
        parameters = Range("a1") + "&tweet_mode=extended" + "&count=100&" + "max_id=" + max_id
    End If
    '
    ' Send the search request to twitter
    '
    req.Open "GET", url + "?" + parameters, False
    req.setRequestHeader "Authorization", authorization_value
    req.send
    '
    ' Wait until all the data is sent back
    '
    While req.readyState <> 4 ' #4 means all data received
        DoEvents
    Wend
    '
    ' Get the response object and put it into the json variable
    '
    response = req.responseText
    ' Range("a2") = response ' debug to check resposnes
    sc.Language = "javascript"
    Set json = sc.Eval("(" + response + ")")
    '
    ' Loop through every tweet in the json
    '
    For Each x In json.statuses
        If (max_id <> x.id_str) Then
            Range("B" + CStr(dnum)) = CStr(x.id_str)
            Range("C" + CStr(dnum)) = x.created_at
            Range("D" + CStr(dnum)) = x.user.screen_name
            Range("E" + CStr(dnum)) = x.user.created_at
            Range("F" + CStr(dnum)) = x.user.statuses_count
            Range("G" + CStr(dnum)) = x.user.favourites_count
            Range("H" + CStr(dnum)) = x.user.followers_count
            Range("I" + CStr(dnum)) = x.user.friends_count
            Range("J" + CStr(dnum)) = x.user.listed_count
            Range("K" + CStr(dnum)) = x.favorite_count
            Range("L" + CStr(dnum)) = x.retweet_count
            Range("M" + CStr(dnum)) = x.user.verified

            Range("N" + CStr(dnum)) = x.user.geo_enabled
            'medias = ""
            'For Each uu In x.extended_entities.media
            '    If (medias <> "") Then medias = medias + ","
            '    medias = medias + uu.media_url_https
            'Next
            'Range("N" + CStr(dnum)) = medias
            
            Range("O" + CStr(dnum)) = x.lang
            
            If (InStr(x.full_text, "=") <> 1) Then
                Range("P" + CStr(dnum)) = x.full_text
            Else
                Range("P" + CStr(dnum)) = "!" + x.full_text
            End If
            
            hashtags = ""
            For Each uh In x.entities.hashtags
                If (hashtags <> "") Then hashtags = hashtags + ","
                hashtags = hashtags + uh.text
            Next
            Range("Q" + CStr(dnum)) = hashtags

            mentions = ""
            For Each um In x.entities.user_mentions
                If (mentions <> "") Then mentions = mentions + ","
                mentions = mentions + um.screen_name
            Next
            Range("R" + CStr(dnum)) = mentions
            
            dnum = dnum + 1
            max_id = x.id_str
            If (stime <> CStr(x.created_at)) Then
                scount = 0
                stime = CStr(x.created_at)
            Else
                scount = scount + 1
            End If
            Application.StatusBar = CStr(dnum) + ":" + CStr(scount) + ":" + CStr(x.created_at)
        End If
        Range("A3") = dnum - 3
        DoEvents
    Next
    
    sleepNow (2) ' If you don't want to use getAllExtendedSlowly, uncomment this, but unreliable
Loop Until (json.statuses = "" Or max_id = last_id)

CreateObjectx86 , True ' close mshta host window at the end
MsgBox "Get All Done"
End Sub
'
' getFriendStatus: Check the friendship status of the 100 OR LESS usernames highlighted
'
' Entry:  A columns of names selected
' Exit:   To the right of names, friend status
'
Function getFriendStatus(n1 As String, n2 As String)
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As ScriptControl
Dim url As String
Dim authorization_value As String
Dim parameters As String

Dim response As String
Dim json As Object

Dim x As Object
Dim text As String ' hack to prevent vba from capitalizing text to Text
Dim connections As String ' hack to prevent vba from capitalizing connections

Dim s As Range
Dim roff As Long
Dim names100 As String
Dim c As Variant
Dim sstatus


'
' Login if necesary
'
If Range("b2") = "" Then TwitterLogin
bearer_token = Range("b2")
authorization_value = "Bearer " + bearer_token ' assumes getTwitterToken called

'
' Create Script Control
'
Set sc = CreateObjectx86("MSScriptControl.ScriptControl")

'
' Initialize names to search for
'
Set s = Selection

'
' set Twitter API URL
'
url = "https://api.twitter.com/1.1/friendships/show.json"

parameters = "source_screen_name=" + n1 + "&target_screen_name=" + n2

'
' Send the search request to twitter
'
req.Open "GET", url + "?" + parameters, False
req.setRequestHeader "Authorization", authorization_value
req.send
'
' Wait until all the data is sent back
'
While req.readyState <> 4 ' #4 means all data received
    DoEvents
Wend
'
' Get the response object and put it into the json variable
'
response = req.responseText
'MsgBox response
sc.Language = "javascript"
Set json = sc.Eval("(" + response + ")")
'
' construct output value
'
sstatus = CStr(json.relationship.target.following) + "," + CStr(json.relationship.target.followed_by)

'
' Close scripting object
'
CreateObjectx86 , True ' close mshta host window at the end

getFriendStatus = sstatus

End Function
'
' getUser: Get information about a specified user based on user query string
'          I include this for a future revision of the scraper
'          It's unnecessary because I scrape this information in getAll
'
Private Sub getUser()
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As New ScriptControl
Dim url As String
Dim response As String
Dim authorization_value As String
Dim parameters As String
Dim dcell As String
Dim x As Object
Dim text As String ' hack to prevent vba from capitalizing text to Text
Dim id As Long
Dim name, description, status


If Range("b2") = "" Then TwitterLogin

bearer_token = Range("b2")

dnum = 3
dcell = "B"

url = "https://api.twitter.com/1.1/users/show.json"

authorization_value = "Bearer " + bearer_token ' assumes TwitterLogin successful
parameters = Range("a1")
req.Open "GET", url + "?" + parameters, False
req.setRequestHeader "Authorization", authorization_value
req.send

While req.readyState <> 4
    DoEvents
Wend
response = req.responseText
sc.Language = "javascript"
Set json = sc.Eval("(" + response + ")")
Range("followers") = json.followers_count
Range("following") = json.friends_count
Range("tweetname") = json.name ' causes problem
Range("favorites") = json.favourites_count
Range("description") = json.description
Range("nickname") = json.screen_name
Range("LastTweet") = json.status.text

End Sub
'
' getOneTweet: Get a single tweet specified in the query string
'              I include this for a future revision of the scraper
'              Very rarely do you look at just one tweet
'              ... and getAll gets this information and more
'
Private Sub getOneTweet()
Dim req As New XMLHTTP60 ' Don't forget Tools > References : Microsoft XML,v6.0
Dim sc As New ScriptControl
Dim url As String
Dim response As String
Dim authorization_value As String
Dim parameters As String
Dim dcell As String
Dim x As Object
Dim text As String ' hack to prevent vba from capitalizing text to Text
Dim id As Long
Dim name, description, status


If Range("b2") = "" Then TwitterLogin

bearer_token = Range("b2")

dnum = 3
dcell = "B"

url = "https://api.twitter.com/1.1/statuses/show.json"

authorization_value = "Bearer " + bearer_token ' assumes getTwitterToken called
parameters = Range("a1")
req.Open "GET", url + "?" + parameters, False
req.setRequestHeader "Authorization", authorization_value
req.send

While req.readyState <> 4
    DoEvents
Wend
response = req.responseText
sc.Language = "javascript"
Set json = sc.Eval("(" + response + ")")
Range("b4") = json.text
Range("b5") = json.favorite_count
Range("b6") = json.retweet_count
End Sub
'
' sleepNow: sleeps for the specified number of seconds
'
Sub sleepNow(s As Long)
Dim wakeup As Date

wakeup = DateAdd("s", s, Now)
Do
    DoEvents
Loop Until Now > wakeup
End Sub
'
' fixDate: Puts an Excel Date into a VBA date for calculation purposes
'
Function fixDate(ByVal d As String) As String
Dim dlen, x As Long
Dim dyear, ddate As String
' sample d=Sat Mar 14 00:00:01 +0000 2015
dyear = Mid(d, 27, 4) ' 2015
ddate = Mid(d, 5, 15) ' Mar 14 00:00:01
fixDate = Left(ddate, 7) + dyear + " " + Right(ddate, 8)
'Mar 14 2015 00:00:01
End Function


'
' calcMaxDate: Adds a fixed interval to a date for histogram purposes
'
' Entry: HINTERVAL (histogram interval in minutes) set at top of file
'
Function calcMaxDate(r As String) As Date
Dim firstdate As String
Dim maxdate As Date
firstdate = fixDate(r)
maxdate = CDate(firstdate)
'firstdate = CStr(Month(maxdate)) + "/" + CStr(Day(maxdate)) + "/" + CStr(Year(maxdate)) + " " + CStr(Hour(maxdate)) + ":" + CStr(Minute(maxdate)) + ":00"
firstdate = CStr(Month(maxdate)) + "/" + CStr(Day(maxdate)) + "/" + CStr(Year(maxdate)) + " " + CStr(Hour(maxdate)) + ":00:00"
maxdate = CDate(firstdate)
maxdate = DateTime.DateAdd("n", HINTERVAL, maxdate)
calcMaxDate = maxdate
End Function
'
' createTweetHistogram: Creates a histogram of values
'
' Entry:
'    Entire date column selected, this is currently column C
'    HINTERVAL set (at top of file, currently defaulted to 30 minutes)
'
' Exit:
'    Histogram bin table, below all tweets
'
Sub createTweetHistogram()
Dim maxdate As Date
Dim currdate As Date
Dim newdate As String
Dim x As Range
Dim firstdate, ffirstdate As String
Dim r As Range
Dim c As Long
Dim hr As Long
Dim drow As Long
Dim dcol As Long
Dim i As Long
Set r = Selection

maxdate = calcMaxDate(r.Cells(r.count, 1)) ' the first date
drow = r.row + r.count + 1 ' destination row
dcol = 1 ' start column for output date, Column A

c = 0
hr = 0
For i = r.count To 1 Step -1 ' Scraper grabs dates Hi to lo, so iterate backwards
    Set x = r(i, 1)
    newdate = fixDate(x)
    currdate = CDate(newdate)
    If (currdate <= maxdate) Then
        c = c + 1
    Else
        'MsgBox (c) ' write it out
        Cells(drow + hr, dcol) = DateTime.DateAdd("h", -0, maxdate)
        Cells(drow + hr, dcol + 1) = c
        c = 1 ' reset data
        hr = hr + 1
        maxdate = DateTime.DateAdd("n", HINTERVAL, maxdate)
    End If
    Cells(drow + hr, dcol) = DateTime.DateAdd("h", -0, maxdate)
    Cells(drow + hr, dcol + 1) = c
    DoEvents
Next

MsgBox ("createTweetHistogram: Done")

End Sub
'
' countUserTweets: Creates a table of users & tweets
'
' Entry:
'    Entire user column selected, this is currently column D
'
' Exit:
'    User count table, below all tweets
'
Sub countUserTweets()
Dim r As Range
Dim d As New Dictionary
Dim i, nrows As Long
Dim s As String
Dim x As Variant

Set r = Selection
nrows = r.Cells.Rows.count

For Each x In r
    s = CStr(x)
    If (d.Exists(s)) Then
        d(s) = d(s) + 1
    Else
        d(s) = 1
    End If
    DoEvents
Next
i = 2 ' leave a blank line at the end of the selection
For Each x In d.Keys
    r.Cells(nrows + i, 1) = x
    r.Cells(nrows + i, 2) = d(x)
    i = i + 1
    DoEvents
Next

MsgBox ("countUserTweets: Done")
End Sub
'
' countRTs: Counts positive and negative words in a tweet
'
' Entry: P column — non-filtered tweets selected in entirety
' Exit: End of RSTU columns will contain User, #RTs; a skipped line; Actualy Tweet, #RTs, who RTd, times, link to tweet
'
Sub countRTs()
Dim s As Range
Dim de As New Dictionary ' dictionary of entire tweet
Dim dj As New Dictionary ' dictionary of just tweeter
Dim dn As New Dictionary ' dictionary of names who RT'd
Dim dt As New Dictionary ' dictionary of times for who RT'd
Dim di As New Dictionary ' dictionary of ids for a RT
Dim regex As New RegExp
Dim mc As MatchCollection
Dim m As Object
' may change
Dim et, jt, nr, rt, ti As String ' et:entire tweet, jt:just tweeter, nr:name of retweeter, rt: retweet time
Dim doff, coff As Long
Dim c, k As Variant
Dim row As Long

' regex.Global = True
regex.IgnoreCase = True
regex.Pattern = "^RT @([^:]+):(.*)"

Set s = Selection
doff = s.count + 2
row = 1
For Each c In s ' go through range
    Set mc = regex.Execute(c)
    For Each m In mc
        nr = CStr(c.Cells(1, -11)) ' name of retweeter
        rt = CStr(c.Cells(1, -12)) ' retweet time
        ti = CStr(c.Cells(1, -13)) ' a retweet id (that links to original id)
        et = m.Value ' entire tweet
        jt = m.SubMatches(0) ' just the tweeter
        If de.Exists(et) Then
            de(et) = de(et) + 1
        Else
            de(et) = 1
        End If
        If dj.Exists(jt) Then
            dj(jt) = dj(jt) + 1
        Else
            dj(jt) = 1
        End If
        If dn.Exists(et) Then
            dn(et) = dn(et) + "," + nr
        Else
            dn(et) = nr
        End If
        If dt.Exists(et) Then
            dt(et) = dt(et) + "," + rt
        Else
            dt(et) = rt
        End If
        If di.Exists(et) Then
            di(et) = ti ' ultimately stores the last id of the first RT, which brings up orig tweet in Twitter
        Else
            di(et) = ti ' ultimately stores the last id of the first RT, which brings up orig tweet in Twitter
        End If
    Next

    Application.StatusBar = "Processing: " + CStr(row) + "/" + CStr(s.count)
    row = row + 1
    DoEvents
Next
' now output the tweeter RT counts
row = 1
coff = 2 ' we have to skip 2 because social network is to left
For Each k In dj.Keys
    s.Cells(doff, coff) = k
    s.Cells(doff, coff + 1) = dj(k)
    doff = doff + 1
    Application.StatusBar = "Phase 1/2: RT Usernames: " + CStr(row) + "/" + CStr(dj.count)
    row = row + 1
    DoEvents
Next

' now output the entire tweet counts
doff = doff + 1 ' skip another row
row = 1
For Each k In de.Keys
    s.Cells(doff, coff) = k
    s.Cells(doff, coff + 1) = de(k)
    s.Cells(doff, coff + 2) = dn(k)
    s.Cells(doff, coff + 3) = dt(k) ' may want to comment this out
    s.Cells(doff, coff + 4) = "https://www.twitter.com/statuses/" + di(k)
    
    doff = doff + 1
    Application.StatusBar = "Phase 2/2: RT Entire:" + CStr(row) + "/" + CStr(de.count)
    row = row + 1
    DoEvents
Next

MsgBox ("countRTs done!")
End Sub
'
' convertTwitterToExcelDates
' INPUT: column of Twitter dates
' OUTPUT: (immediately below column) skip a line, column of Excel Dates
' Twitter Date looks like: Fri Jun 16 01:09:15 +0000 2017
' Excel Date looks like: 6/16/2017 01:09:15
Sub convertTwitterToExcelDates()
Dim s As Range
Dim c As Variant
Dim mc As MatchCollection
Dim m As Match
Dim regex As New RegExp
Dim roff As Long
Dim newdate As String
Dim mcode As String

regex.Pattern = "^(.{3}) (.{3}) (.{2}) (.{8}) (.{5}) (.{4})"
Set s = Selection

roff = s.count + 2 ' row offset


For Each c In s
    Set mc = regex.Execute(c)
    mcode = "00"
    If (LCase(mc(0).SubMatches(1)) = "jan") Then mcode = "01"
    If (LCase(mc(0).SubMatches(1)) = "feb") Then mcode = "02"
    If (LCase(mc(0).SubMatches(1)) = "mar") Then mcode = "03"
    If (LCase(mc(0).SubMatches(1)) = "apr") Then mcode = "04"
    If (LCase(mc(0).SubMatches(1)) = "may") Then mcode = "05"
    If (LCase(mc(0).SubMatches(1)) = "jun") Then mcode = "06"
    If (LCase(mc(0).SubMatches(1)) = "jul") Then mcode = "07"
    If (LCase(mc(0).SubMatches(1)) = "aug") Then mcode = "08"
    If (LCase(mc(0).SubMatches(1)) = "sep") Then mcode = "09"
    If (LCase(mc(0).SubMatches(1)) = "oct") Then mcode = "10"
    If (LCase(mc(0).SubMatches(1)) = "nov") Then mcode = "11"
    If (LCase(mc(0).SubMatches(1)) = "dec") Then mcode = "12"
    
    newdate = mcode + "/" + mc(0).SubMatches(2) + "/" + mc(0).SubMatches(5) + " " + mc(0).SubMatches(3)
    s.Cells(roff, 1) = CDate(newdate)
    roff = roff + 1
    Application.StatusBar = "Converting " + CStr(roff - 2 - s.count) + "/" + CStr(s.count)
    DoEvents
Next
MsgBox "Finished: Convert Twitter to Excel Dates"
End Sub

'
' genEdges: Generate two kinds of social edges from Tweeter to mentioner & RT to Tweeter
'
' Entry: D column — selected in entirety (tweeter names)
' Exit: End of M,N column —
'                           tweeter*,mentioner* -- 1st kind of edge
'                           blank row (either graph the top or bottom edges)
'                           RT* to Tweeter* -- 2nd kind of edge
'
'
Private Sub genEdges()
Dim r As Range
Dim c, mc, ma As Variant
Dim i, dst, row As Long
Dim mname As Variant

    Set r = Selection
    dst = r.count + 2
    ' generate a Type 1 edge:from tweeter to mentioner, includes RTs
    row = 1
    For Each c In r
        mc = c.Cells(1, 15) ' 15 columns over contains a list of all users mentioned
        ma = Split(mc, ",")
        For Each mname In ma
            r.Cells(dst, 10).Value = c.Value
            r.Cells(dst, 11).Value = mname
            dst = dst + 1
        Next
        Application.StatusBar = "Outputting Type 1 Edges for Tweet #" + CStr(row) + "/" + CStr(r.count)
        row = row + 1
        DoEvents
    Next
    '
    ' generate a Type-2 edge from RT to mentioner
    ' I don't anticipate using this as the primary social network
    '
    dst = dst + 1
    row = 1
    For Each c In r
        mname = getRTName(c.Cells(1, 13))
        If (mname <> "") Then
            r.Cells(dst, 10).Value = mname
            r.Cells(dst, 11).Value = c
            dst = dst + 1
        End If
        Application.StatusBar = "Outputting Type 2 Edges for Tweet #" + CStr(row) + "/" + CStr(r.count)
        row = row + 1
        DoEvents
    Next
    
    MsgBox "genEdges Done"
End Sub
'
' stripText: Removes unwanted strings from the input.
'
' For performance reasons, if you run this on a large text column,
' … you should copy and paste-values so the function doesn't try to re-run
'
Function stripText(t As String)
Dim regex As New RegExp 'reference Microsoft VBScript Regular Expressions 5.5

'
' Replace carriage returns with spaces
'
regex.Global = True
regex.Pattern = "[\n\r]"
t = regex.Replace(t, " ")
'
' Replace URLS with [URL]
'
regex.Pattern = "(http[s]?:\/\/[^ ]*[ ]{1})|(http[s]?:[^ ]*$)" 'replace any URL with
t = regex.Replace(t, "[URL] ")
'
' Replace &amp with &
'
regex.Pattern = "&amp"
t = regex.Replace(t, "&")
'
' Remove anything that's not a letter, digit, or #@&[] (add more if needed)
' FUTURE REV: emoji's but > 0xffff
'
regex.Pattern = "[^\w\s#&@\[\]]" ' remove that's not digit or space or #@&[]
t = regex.Replace(t, "")
'
' Remove leading and trailing spaces (didn't work above)
'
regex.Pattern = "(^\s+)|(\s+$)" ' remove that's not digit or space or #@&[]
t = regex.Replace(t, "")

'
' Replace multiple spaces with a single space
'
regex.Pattern = "\s+"
t = regex.Replace(t, " ")

stripText = t
End Function
'
' filterColumn: Strips punctuation and URLs out of a cell
' Entry: A column — selected in its entirety
' Exit: A filtered column, output 4 columns over (if original column is P then output is in S)
'
Sub filterTweets()
Dim s As Range
Dim c As Variant
Dim row As Long

    Set s = Selection
    row = 1
    For Each c In s
        s.Cells(row, 4) = stripText(CStr(c)) ' filter and output text 4 columns over
        Application.StatusBar = row
        row = row + 1
        DoEvents
    Next
    MsgBox "filterTweets Done"
End Sub
'
' countWords: Count all unique words in a selected range column
'
' Entry: A column — selected in entirety AND stripped of all unnecessary punctuation (run stripText)
' Exit: Underneath that column, skips a line, then two columns (word, frequency)
'
'
Sub countWords()
Dim s As Range
Dim d As New Dictionary
Dim doff As Long
Dim words, word As Variant
Dim c, k As Variant
Dim row As Long

Set s = Selection
doff = s.count + 2
row = 1
For Each c In s ' go through range
    words = Split(c, " ")
    
    For Each word In words
        word = LCase(word)
        If (word = "") Then
            Dim xyzzy As Long
            xyzzy = 1
        End If
        If d.Exists(word) Then
            d(word) = d(word) + 1
        Else
            d(word) = 1
        End If
    Next
    Application.StatusBar = "Processing: " + CStr(row) + "/" + CStr(s.count)
    row = row + 1
    DoEvents
Next

row = 1
For Each k In d.Keys
    s.Cells(doff, 1) = k
    s.Cells(doff, 2) = d(k)
    doff = doff + 1
    Application.StatusBar = "Outputting: " + CStr(row) + "/" + CStr(d.count)
    row = row + 1
    DoEvents
Next

MsgBox ("countWords done!")
End Sub
'
' genWordEdges: Generate word edges and frequency
'
' Entry: A column — selected in entirety AND stripped of all unnecessary punctuation (run stripText)
' Exit: Underneath that column, skips a line, then three columns over, (edge1, edge2, frequency)
'
'
Sub genWordEdges()
Dim s As Range
Dim d As New Dictionary
Dim doff, wordcount As Long
Dim words, word, edge1, edge2 As Variant
Dim c, k As Variant
Dim i As Long
Dim row As Long

Set s = Selection
doff = s.count + 2
row = 1
For Each c In s ' go through range
    words = Split(c, " ")
    
    wordcount = UBound(words)
    For i = 0 To (wordcount - 1)
        edge1 = LCase(words(i))
        edge2 = LCase(words(i + 1))
        word = edge1 + "-" + edge2
        If d.Exists(word) Then
            d(word) = d(word) + 1
        Else
            d(word) = 1
        End If
        DoEvents
    Next
    Application.StatusBar = "Processing: " + CStr(row) + "/" + CStr(s.count)
    row = row + 1
Next

row = 1
For Each k In d.Keys
    word = Split(k, "-")
    s.Cells(doff, 4) = word(0)
    s.Cells(doff, 5) = word(1)
    s.Cells(doff, 6) = d(k)
    doff = doff + 1
    
    Application.StatusBar = "Outputing: " + CStr(row) + "/" + CStr(d.count)
    row = row + 1
    
    DoEvents
Next

MsgBox ("genWordEdges done!")

End Sub
'
' genSocialEdges: Generate two kinds of social edges from Tweeter to mentioner & RT to Tweeter
'
' Entry: D column — selected in entirety (tweeter names)
' Exit: End of M,N column —
'                           tweeter*,mentioner* -- 1st kind of edge
'                           blank row (either graph the top or bottom edges)
'                           RT* to Tweeter* -- 2nd kind of edge
'
'
Sub genSocialEdges()
Dim r As Range
Dim c, mc, ma, na, k As Variant
Dim i, dst, row As Long
Dim mname As Variant
Dim edge, node1, node2 As String
Dim d As New Dictionary

    Set r = Selection
    dst = r.count + 2
    ' generate a Type 1 edge:from tweeter to mentioner, includes RTs
    row = 1
    For Each c In r
        mc = c.Cells(1, 15) ' 15 columns over contains a list of all users mentioned
        ma = Split(mc, ",")
        For Each mname In ma
            node1 = c.Value
            node2 = mname
            edge = CStr(node1) + "-" + CStr(node2)
            If d.Exists(edge) Then
                d(edge) = d(edge) + 1
            Else
                d(edge) = 1
            End If
        Next
        Application.StatusBar = "Processing Type 1 Edges for Tweet #" + CStr(row) + "/" + CStr(r.count)
        row = row + 1
        DoEvents
    Next
    '
    ' Now output type 1 edges
    '
    row = 1
    For Each edge In d.Keys
        na = Split(edge, "-") ' node array
        node1 = na(0) ' split array is zero-based!
        node2 = na(1)
        r.Cells(dst, 10) = node1
        r.Cells(dst, 11) = node2
        r.Cells(dst, 12) = d(edge) ' count of edges
        
        Application.StatusBar = "Outputing Type 1 Edge: " + CStr(row) + "/" + CStr(d.count)
        row = row + 1
        DoEvents

        dst = dst + 1
    Next
    '
    ' generate a Type-2 edge from RT to mentioner
    ' I don't anticipate using this as the primary social network
    '
    d.RemoveAll ' clear out dictionary
    dst = dst + 1 ' skip a blank line for type 2 edges
    row = 1
    For Each c In r
        mname = getRTNameRegEx(c.Cells(1, 13))
        If (mname <> "") Then
            node1 = mname
            node2 = c
            edge = CStr(node1) + "-" + CStr(node2)
            If (d.Exists(edge)) Then
                d(edge) = d(edge) + 1
            Else
                d(edge) = 1
            End If
        End If
        Application.StatusBar = "Processing Type 2 Edges for Tweet #" + CStr(row) + "/" + CStr(r.count)
        row = row + 1
        DoEvents
    Next
    
    row = 1
    For Each k In d.Keys
        na = Split(k, "-")
        node1 = na(0)
        node2 = na(1)
        r.Cells(dst, 10) = node1
        r.Cells(dst, 11) = node2
        r.Cells(dst, 12) = d(k)
        
        Application.StatusBar = "Outputing Type 2 Edge: " + CStr(row) + "/" + CStr(d.count)
        row = row + 1
        DoEvents
        
        dst = dst + 1
    Next
    MsgBox "genSocialEdges Done"
End Sub
'
' genSentiment: Counts positive and negative words in a tweet
'
' Entry: S column — filtered tweets selected in entirety
' Exit: End of T,U,V column will contain # of positive words, # of negative words, difference, respectively
'
Sub genSentiment()
Dim r As Range
Dim c As Variant
Dim pos, neg, dif, row As Long

    Set r = Selection
    
    row = 1
    For Each c In r
        pos = countPos(c)
        neg = countNeg(c)
        dif = pos - neg
        c.Cells(1, 2) = pos
        c.Cells(1, 3) = neg
        c.Cells(1, 4) = dif
        Application.StatusBar = "Processing Tweet #" + CStr(row) + "/" + CStr(r.count)
        row = row + 1
        DoEvents
    Next
    MsgBox "genSentiment Done"
End Sub
'
' removeStopWords: Remove stop words
'
' Entry: P column — original tweets selected in their entirety
' Exit: End of P column contains cell with stopwords removed
'
Sub removeStopWords()
Dim r As Range
Dim c, stopword As Variant
Dim i, pos, neg, dif, row, scount As Long
Dim sw As Variant
Dim sr As Range
Dim doff As Long
Dim stopped, sword, rgx As String
Dim mc As MatchCollection

    Set r = Selection
    doff = r.count + 2 ' put results underneath selection
    
    ' first read all the stopwords into a dictionary
    If (dswords.count = 0) Then
        dswords.RemoveAll
        Set sw = Worksheets("SWords") ' Stop Words
        Set sr = sw.Range("A:A")
        scount = sr.End(xlDown).row
        
        For i = 1 To scount
            dswords(LCase(sr.Cells(i, 1))) = 0
        Next
    End If
    
   
    row = 1
    For Each c In r
        stopped = LCase(c)
        rgx = "[A-Za-z'-]+"
        Set mc = regexMatch(CStr(stopped), rgx)
        For Each stopword In mc
            sword = LCase(stopword)
            If (dswords.Exists(sword)) Then
                rgx = "(^" + sword + " | " + sword + " | " + sword + "$)"
                stopped = regexReplace(CStr(stopped), rgx, " ")
            End If
        Next
               
        'For Each stopword In dswords.Keys ' LONG WAY
        '    sword = CStr(stopword)
        '    rgx = "(^" + sword + " | " + sword + " | " + sword + "$)"
        '    stopped = regexReplace(CStr(stopped), rgx, " ")
        'Next
        r.Cells(doff, 1) = stopped
        
        Application.StatusBar = "Processing Tweet #" + CStr(row) + "/" + CStr(r.count)
        doff = doff + 1
        row = row + 1
        DoEvents
    Next
    MsgBox "removeStopwords Done"
End Sub
'
' createDocumentTermMatrix: Creates a document term matrix
'
' Entry: A highlighted column of statement
' Exit: new worksheet as a DocumentTermMatrix
'
Sub createDocumentTermMatrix()
Dim s As Range
Dim c, word, key As Variant
Dim a As Worksheet
Dim d As New Dictionary
Dim mc As MatchCollection
Dim rgx As String
Dim row, col, count As Long

'
' Grab all the words
'
Set s = Selection
For Each c In s
    rgx = "[^ ]+"
    Set mc = regexMatch(CStr(c), CStr(rgx))
    For Each word In mc
        If d.Exists(LCase(word)) = False Then
            d(LCase(word)) = 0
        End If
    Next
Next
'
' Create a new worksheet
'
ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.count)
Set a = ActiveSheet
'
' Write the terms as columns
'
col = 2
For Each key In d.Keys
    a.Cells(1, col) = CStr(key)
    col = col + 1
Next
row = 2
For Each c In s
    col = 1
    a.Cells(row, col) = row - 1
    col = col + 1
    For Each key In d.Keys
        count = regexCounter(CStr(c), "(^" + key + " | " + key + " | " + key + "$)")
        a.Cells(row, col) = count
        col = col + 1
    Next
    row = row + 1
Next

'a.Cells(1, 1) = "Hello World"
' a.Cells(2, 1) = "Yabadabba Doo"
End Sub
'
' replaceWithNGrams: replaces word combinations with their NGram equivalent
'
' Entry: P column — original tweets selected in their entirety
'        NGrams Tab — containing ngrams, like "Wake Up America"
' Exit: End of P column contains data with NGrams added (phrases have dashes, e.g., "Wake-Up-America")
'
Sub replaceWithNGrams()
Dim r As Range
Dim c, ng As Variant
Dim i, row, scount As Long
Dim sw As Variant
Dim sr As Range
Dim doff As Long
Dim tweet, sng, dashsng As String
Dim mc As MatchCollection

    Set r = Selection
    doff = r.count + 2 ' put results underneath selection
    
    ' first read all the ngrams into a dictionary
    If (dgwords.count = 0) Then
        dgwords.RemoveAll
        Set sw = Worksheets("NGrams") ' NGrams
        Set sr = sw.Range("A:A")
        scount = sr.End(xlDown).row
        
        For i = 1 To scount
            dgwords(LCase(sr.Cells(i, 1))) = 0
        Next
    End If
    
   
    row = 1
    For Each c In r
        tweet = LCase(c)
        For Each ng In dgwords.Keys
            sng = LCase(ng)
            dashsng = Replace(sng, " ", "-")
            tweet = Replace(tweet, sng, dashsng)
            'dashsng = regexReplace(CStr(sng), " ", "-")
            'tweet = regexReplace(CStr(tweet), CStr(sng), CStr(dashsng))
        Next
        r.Cells(doff, 1) = tweet
        
        Application.StatusBar = "Processing Tweet #" + CStr(row) + "/" + CStr(r.count)
        doff = doff + 1
        row = row + 1
        DoEvents
    Next
        
    MsgBox "replaceWithNGrams Done"
End Sub
Private Sub createSentimentDB()
Dim pw, nw As Variant
Dim pr, nr As Range
Dim pcount, ncount, i As Long
Dim pwords(), nwords()

If ((dpwords.count = 0) Or (dnwords.count = 0)) Then
    dpwords.RemoveAll
    dnwords.RemoveAll
    
    Set pw = Worksheets("PWords") ' social presence words or phrases
    Set nw = Worksheets("NWords") ' egocentric words or phrases
    Set pr = pw.Range("A:A")
    Set nr = nw.Range("A:A")
    
    pcount = pr.End(xlDown).row
    ncount = nr.End(xlDown).row
    
    ReDim pwords(pcount)
    ReDim nwords(ncount)
    
    For i = 1 To pcount
        dpwords(LCase(pr.Cells(i, 1))) = 0
    Next
    For i = 1 To ncount
        dnwords(LCase(nr.Cells(i, 1))) = 0
    Next
End If
    'MsgBox ("createSentimentDB finished")
End Sub

Function countPos(c)
Dim word, words As Variant
Dim numpos As Long

    createSentimentDB
    numpos = 0
    
    words = Split(c, " ")
    For Each word In words
        word = LCase(word)
        If dpwords.Exists(word) Then
            'MsgBox (word)
            numpos = numpos + 1
            ' save the word too
        End If
    Next
    countPos = numpos
End Function

Function countNeg(c)
Dim word, words As Variant
Dim numneg As Long

    createSentimentDB
    numneg = 0
    
    words = Split(c, " ")
    For Each word In words
        word = LCase(word)
        If dnwords.Exists(word) Then
            'MsgBox (word)
            numneg = numneg + 1
            ' save the word too
        End If
    Next
    countNeg = numneg
End Function
'
' CreateObjectx86 & CreateWindow are public domain code
' They allow use of the ScriptControl in 64-bit Excel
'
' CreateObjectx86: http://stackoverflow.com/questions/9725882/getting-scriptcontrol-to-work-with-excel-2010-x64/38134477
' CreateWindow: http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
'
Function CreateObjectx86(Optional sProgID, Optional bClose = False)

    Static oWnd As Object
    Dim bRunning As Boolean

    #If Win64 Then
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        If bClose Then
            If bRunning Then oWnd.Close
            Exit Function
        End If
        If Not bRunning Then
            Set oWnd = CreateWindow()
            oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
        End If
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        If bClose Then
            If bRunning Then oWnd.Close
            Exit Function
        End If
        If Not bRunning Then
            Set CreateObjectx86 = CreateObject(sProgID)
        End If
    #End If
End Function
Function CreateWindow()
    Dim sSignature, oShellWnd, oProc

    On Error Resume Next
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""about:<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.getProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        Next
    Loop

End Function
'
' findDuplicateNames
'   Input: two comma-separated lists of names
'   Output: a list of common names
'
Function findDuplicateNames(s1, s2) 's1, s2 are comma separated lists
Dim dict As New Dictionary
Dim a1, a2 As Variant ' array of names
Dim tn, n, cn As String ' trimmed name, name, common names
Dim key As Variant

a1 = Split(s1, ",")
a2 = Split(s2, ",")

For Each n In a1
tn = Trim(n)
    If dict.Exists(tn) Then
        dict(tn) = dict(tn) + 1
    Else
        dict(tn) = 1
    End If
Next

For Each n In a2
tn = Trim(n)
    If dict.Exists(tn) Then
        dict(tn) = dict(tn) + 1
    Else
        dict(tn) = 1
    End If
Next

cn = ""
For Each key In dict.Keys
    If (dict(key) > 1) Then
        If (cn = "") Then
            cn = key
        Else
            cn = cn + "," + key
        End If
    End If
Next
findDuplicateNames = cn

End Function
'
' calcTimeDeltas
'   input: a comma-separated list of times
'   output: a list of time deltas
'
Function calcTimeDeltas(s As String)
Dim regex As New RegExp
Dim mc As MatchCollection
Dim m As Match
Dim tarray As Variant
Dim t As Variant
Dim d, lastdate As Date
Dim delta As Long
Dim darray As String ' array of deltas
tarray = Split(s, ",")
lastdate = 0
darray = ""
For Each t In tarray
    regex.Pattern = "(\d{2}:\d{2}:\d{2})"
    Set mc = regex.Execute(t)
    For Each m In mc ' note: only 1 item
        d = CDate(m)
    Next
    If (lastdate = 0) Then
        lastdate = d
    Else
        delta = DateDiff("s", d, lastdate)
        If (darray = "") Then
            darray = CStr(delta)
        Else
            darray = CStr(delta) + "," + darray
        End If
    End If
    lastdate = d
Next
calcTimeDeltas = darray
End Function
Function regexCounter(s As String, p As String)
Dim regex As New RegExp
Dim m As MatchCollection

regex.Global = True
regex.IgnoreCase = True
regex.Pattern = p

Set m = regex.Execute(s)
regexCounter = m.count

End Function

Function regexReplace(s As String, p As String, r As String)
Dim regex As New RegExp
Dim ns As String

regex.Global = True
regex.IgnoreCase = True
regex.Pattern = p

ns = regex.Replace(s, r)
regexReplace = ns
End Function

Function regexMatch(s As String, p As String) As MatchCollection
Dim regex As New RegExp
Dim m As MatchCollection

regex.Global = True
regex.IgnoreCase = True
regex.Pattern = p

Set m = regex.Execute(s)
Set regexMatch = m
End Function
