Attribute VB_Name = "vBulletinApi"
Option Explicit

Private Const URL As String = "https://www.example.com/api.php"
Private Const API_KEY As String = "insert API key here"
Private Const CLIENT_NAME As String = "insert some client name here"
Private Const CLIENT_VERSION As String = "1.0"
Private Const PLATFORM_NAME As String = "Windows"
Private Const PLATFORM_VERSION As String = "10"
Private Const CLIENT_UNIQUE_ID As String = "insert random numeric client id here"

Private mUsername As String
Private mPassword As String
Private mAccessToken As String
Private mClientId As String
Private mSecret As String
Private mApiVersion As String
Private mLoggedIn As Boolean

Public Static Property Get username() As String
    username = mUsername
End Property

Public Static Property Let username(ByVal lUsername As String)
    mUsername = lUsername
End Property

Public Property Let Password(ByVal lPassword As String)
    mPassword = lPassword
End Property

Private Static Property Get Request() As MSXML2.XMLHTTP60
    Dim objRequest As MSXML2.XMLHTTP60
    If objRequest Is Nothing Then
        Set objRequest = New MSXML2.XMLHTTP60
    End If
    Set Request = objRequest
End Property

Private Property Get IsInitialized() As Boolean
    IsInitialized = (Len(mAccessToken) > 0)
End Property

Public Function UpdateGroupMemberships(lUserid As Integer, lMemberGroups As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Updates a user's additional group memberships.
' This won't touch the user's main group.
' It will overwrite any existing additional group memberships.
'
' Arguments:
'   lUserid         User's id.
'   lMemberGroups   Variant array of group ids for the user's additional group
'                   memberships.
'
' Returns:
'   True, if the update succeeded.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    UpdateGroupMemberships = False
    
    ' Log in if necessary
    If Not Login() Then
        Exit Function
    End If

    Dim qs As String
    qs = BuildApiQueryString("user.save")
        
    Dim bodyParams As Dictionary: Set bodyParams = New Dictionary
    With bodyParams
        .Add "userid", lUserid
        .Add "password", ""
        .Add "options", ""
        .Add "adminoptions", ""
        .Add "userfield", ""
    End With
        
    If IsNull(lMemberGroups) Then
        bodyParams.Add "user[membergroupids]", ""
    Else
        bodyParams.Add "user[membergroupids]", Join(lMemberGroups, ",")
    End If
    
    Dim body As String
    body = BuildQueryString(bodyParams)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL & "&" & qs, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send body
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to log in: [" & .Status & "] " & .responseText
            Exit Function
        Else
            ' Check whether the response contains JSON or not.
            If (Left(.responseText, 1) = "{") Then
                ' This API method only returns JSON in case of errors
                Dim response As Dictionary
                Set response = JsonConverter.ParseJson(.responseText)
                CheckResponseForErrors response
            Else
                ' In case of success vBulletin returns the user's id as
                ' plain text. We don't care about it, though.
                UpdateGroupMemberships = True
            End If
        End If
    End With
End Function

Public Function RegisterUser( _
        lUsername As String, _
        lPassword As String, _
        lEmail As String, _
        lLanguageId As Integer, _
        lGroupId As Integer, _
        lMemberGroups As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Registers a new user on the forum.
'
' Arguments:
'   lUsername       User's username.
'   lPassword       User's password.
'   lEmail          User's email address. Will be validated by this function.
'   lLanguageId     User's language id.
'   lGroupId        User's main group id.
'   lMemberGroups   Variant array of group ids for the user's additional group
'                   memberships.
'
' Returns:
'   The new user's id or 0 if the user could not be registered.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RegisterUser = 0
    
    ' Log in if necessary
    If Not Login() Then
        Exit Function
    End If

    Dim qs As String
    qs = BuildApiQueryString("user.save")
        
    Dim bodyParams As Dictionary: Set bodyParams = New Dictionary
    With bodyParams
        .Add "userid", 0
        .Add "user[username]", lUsername
        .Add "user[email]", lEmail
        .Add "user[usergroupid]", lGroupId
        .Add "user[languageid]", lLanguageId
        .Add "password", lPassword
        .Add "options", ""
        .Add "adminoptions", ""
        .Add "userfield", ""
    End With
        
    If Not IsNull(lMemberGroups) Then
        bodyParams.Add "user[membergroupids]", Join(lMemberGroups, ",")
    End If
    
    Dim body As String
    body = BuildQueryString(bodyParams)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL & "&" & qs, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send body
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to register user: [" & .Status & "] " & .responseText
            Exit Function
        Else
            ' Check whether the response contains JSON or not.
            If (Left(.responseText, 1) = "{") Then
                ' This API method only returns JSON in case of errors
                Dim response As Dictionary
                Set response = JsonConverter.ParseJson(.responseText)
                CheckResponseForErrors response
            Else
                ' In case of success vBulletin returns only the user's id
                RegisterUser = CInt(.responseText)
            End If
        End If
    End With
End Function

Public Function DeleteUser(lUserid As Integer) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Deletes a user from the forum.
'
' Arguments:
'   lUserid       User's id.
'
' Returns:
'   True, if the user was deleted.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    DeleteUser = False
    
    ' Log in if necessary
    If Not Login() Then
        Exit Function
    End If
    
    Dim qs As String
    qs = BuildApiQueryString("user.delete")
        
    Dim bodyParams As Dictionary: Set bodyParams = New Dictionary
    With bodyParams
        .Add "userid", lUserid
    End With
    
    Dim body As String
    body = BuildQueryString(bodyParams)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL & "&" & qs, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send body
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to delete user: [" & .Status & "] " & .responseText
            Exit Function
        Else
            ' Check whether the response contains JSON or not.
            If (Left(.responseText, 1) = "{") Then
                ' This API method only returns JSON in case of errors
                Dim response As Dictionary
                Set response = JsonConverter.ParseJson(.responseText)
                CheckResponseForErrors response
            Else
                ' In case of success vBulletin returns only the user's id
                DeleteUser = True
            End If
        End If
    End With
End Function

Public Function FetchUserById(lUserid As Integer) As Dictionary
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fetches the user date for the given user id.
'
' Arguments:
'   lUserid     Id of the user who's data to fetch.
'
' Returns:
'   A dictionary representing the JSON user object as returned by vBulletin.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set FetchUserById = Nothing
    
    ' Log in if necessary
    If Not Login() Then
        Exit Function
    End If
    
    Dim qs As String
    qs = BuildApiQueryString("user.fetchUserinfo")
        
    Dim bodyParams As Dictionary: Set bodyParams = New Dictionary
    With bodyParams
        .Add "userid", lUserid
    End With
    
    Dim body As String
    body = BuildQueryString(bodyParams)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL & "&" & qs, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send body
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to fetch user: [" & .Status & "] " & .responseText
            Exit Function
        Else
            Dim response As Dictionary
            Set response = JsonConverter.ParseJson(.responseText)
        End If
    End With
    
    If CheckResponseForErrors(response) Then
        Exit Function
    End If
    
    Set FetchUserById = response
End Function

Public Function FetchUserList() As Dictionary
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fetches the list of all registered users from the forum.
'
' Returns:
'   A dictionary representing the JSON list of members as returned by vBulletin.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set FetchUserList = Nothing
    
    ' Log in if necessary
    If Not Login() Then
        Exit Function
    End If
    
    Dim qs As String
    qs = BuildApiQueryString("user.memberList")
        
    Dim bodyParams As Dictionary: Set bodyParams = New Dictionary
    With bodyParams
        .Add "criteria[perpage]", 10000
    End With
    
    Dim body As String
    body = BuildQueryString(bodyParams)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL & "&" & qs, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send body
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to fetch user list: [" & .Status & "] " & .responseText
            Exit Function
        Else
            Dim response As Dictionary
            Set response = JsonConverter.ParseJson(.responseText)
        End If
    End With
    
    If CheckResponseForErrors(response) Then
        Exit Function
    End If
    
    Set FetchUserList = response("members")
End Function

Public Function SendMail(lTo As String, lSubject As String, lMessage As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sends an email through the vBulletin mailer.
'
' Arguments:
'   lTo         Email address to send to
'   lSubject    Email subject line
'   lMessage    Email body
'
' Returns:
'   True, if the email was sent successfully.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SendMail = False
    
    ' Log in if necessary
    If Not Login() Then
        Exit Function
    End If
    
    Dim qs As String
    qs = BuildApiQueryString("mail.send")
    
    Dim bodyParams As Dictionary: Set bodyParams = New Dictionary
    With bodyParams
        .Add "to", lTo
        .Add "subject", lSubject
        .Add "message", lMessage
    End With
    
    Dim body As String
    body = BuildQueryString(bodyParams)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL & "&" & qs, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send body
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to send mail: [" & .Status & "] " & .responseText
            Exit Function
        Else
            ' This API method always returns JSON
            Dim response As Dictionary
            Set response = JsonConverter.ParseJson(.responseText)
        End If
    End With
    
    If Not CheckResponseForErrors(response) Then
        SendMail = True
    End If
End Function

Private Function Login(Optional lForce As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logs in as the user in the Username property using the password in the
' Password property.
'
' Arguments:
'   lForce  Forces a new login even if we're already logged in.
'
' Returns:
'   True, if the login was successful.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Login = False
    
    ' We don't repeat a login unless forced to.
    If mLoggedIn And Not lForce Then
        Login = True
        Exit Function
    End If
    
    ' Initialize the API if necessary
    If Not Init() Then
        Exit Function
    End If
    
    Dim qs As String
    qs = BuildApiQueryString("user.login2")
        
    Dim bodyParams As Dictionary: Set bodyParams = New Dictionary
    With bodyParams
        .Add "username", mUsername
        .Add "passwords[password]", mPassword
        .Add "extraAuthInfo", ""
        .Add "logintype", "cplogin"
    End With
    
    Dim body As String
    body = BuildQueryString(bodyParams)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL & "&" & qs, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send body
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to log in: [" & .Status & "] " & .responseText
            Exit Function
        Else
            Dim response As Dictionary
            Set response = JsonConverter.ParseJson(.responseText)
        End If
    End With
    
    If CheckResponseForErrors(response) Then
        MsgBox "Login failed."
        Exit Function
    End If
    
    mLoggedIn = True
    Login = True
End Function

Private Function Init(Optional lForce As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Calls the api.init method on vBulletin and initializes access token, client
' id and secret.
' Must be called before any other methods are called.
'
' Arguments:
'   lForce  Forces a reinitialization even if we already have a valid
'           access token, secret and client id.
'
' Returns:
'   True, if initialization succeeded.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim retried As Boolean: retried = False
    
    Init = False
       
    ' If we're forcing a new init, then delete all the old info.
    If lForce Then
        mAccessToken = ""
        mClientId = ""
        mLoggedIn = False
    End If
    
    ' If we already possess an access token (and it wasn't deleted above using
    ' the lForce argument) then don't reinitialize
    If Len(Trim(mAccessToken)) > 0 Then
        Init = True
        Exit Function
    End If
    
retry: ' This is from where we retry once if the first init call failed.
    Dim query As Dictionary: Set query = New Dictionary
    With query
        .Add "api_m", "api.init"
        .Add "clientname", CLIENT_NAME
        .Add "clientversion", CLIENT_VERSION
        .Add "platformname", PLATFORM_NAME
        .Add "platformversion", PLATFORM_VERSION
        .Add "uniqueid", CLIENT_UNIQUE_ID
    End With
    
    ' If we already have a valid client id, then add it to the query.
    ' This would be the case if our access token had expired.
    If Len(Trim(mClientId)) > 0 Then
        query.Add "api_c", mClientId
    End If
    
    Dim qs As String
    qs = BuildQueryString(query)
    
    ' Perform the request and convert the response to an object
    With Request
        .Open "POST", URL, True
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send qs
        
        ' Wait until request is finished
        While .readyState <> 4
            DoEvents
        Wend
        
        If .Status <> 200 Then
            ' This would indicate serious server or network problems.
            ' Errors during the execution of our API call in vBulletin
            ' are handled later.
            MsgBox "Failed to initialize API: [" & .Status & "] " & .responseText
            Exit Function
        Else
            Dim response As Dictionary
            Set response = JsonConverter.ParseJson(.responseText)
        End If
    End With
    
    If CheckResponseForErrors(response) Then
        If mClientId = "" And Not retried Then
            ' This handles the API error "invalid_clientid". The outdated
            ' client id was already deleted in CheckResponseForErrors.
            ' Now we retry the init call once.
            retried = True
            MsgBox "Client id was invalid. Let's try again."
            GoTo retry
        Else
            ' Every other error or the second "invalid_clientid" error in
            ' sequence are given to the user to fix.
            MsgBox "API initialization failed. Its up to you to fix the problem."
            Exit Function
        End If
    End If
    
    ' Store all the data we need for subsequent API calls.
    mApiVersion = response("apiversion")
    mClientId = response("apiclientid")
    mSecret = response("secret")
    mAccessToken = response("apiaccesstoken")
    
    Init = True
End Function

Private Function BuildApiQueryString(lMethod As String) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Takes a vBulletin API method name and returns a signed query string to pass
' to the server.
'
' Arguments:
'   lMethod     The name of the API method to call.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim apiParams As Dictionary: Set apiParams = New Dictionary
    apiParams.Add "api_m", lMethod
    
    Dim qs As String
    Dim signature As String
    qs = BuildQueryString(apiParams)
    signature = MD5(qs & mAccessToken & mClientId & mSecret & API_KEY)

    With apiParams
        .Add "api_s", mAccessToken
        .Add "api_v", mApiVersion
        .Add "api_c", mClientId
        .Add "api_sig", signature
    End With
    
    BuildApiQueryString = BuildQueryString(apiParams)
End Function

Private Function BuildQueryString(ByRef lParams As Dictionary) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Takes a dictionary and forms a URL query string out of the contents.
'
' Arguments:
'   lParams     The dictionary with query parameters. Keys must be Strings.
'   lPartial    If False, the query string will begin with '?'.
'               If True, it begins with '&'.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim qs As String
    Dim key As Variant
    Dim parName As String
    Dim parValue As String
    
    For Each key In lParams.Keys
        parName = WorksheetFunction.EncodeURL(key)
        parValue = WorksheetFunction.EncodeURL(lParams(key))
        qs = qs & "&" & parName & "=" & parValue
    Next
    
    qs = Replace(qs, Find:="%20", Replace:="+")
    
    ' Return everything but the first character
    BuildQueryString = Right(qs, Len(qs) - 1)
End Function

Private Function CheckResponseForErrors(ByRef lResponse As Dictionary) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks whether the given response dictionary represents an error.
'
' Arguments:
'   lResponse   A response dictionary returned by an API call.
'
' Returns:
'   False, if the API call was successful. True, if it returned an error.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If there is no errors array in the response, then we're good.
    If Not lResponse.Exists("errors") Then
        CheckResponseForErrors = False
        Exit Function
    End If
    
    ' In any other case, there is an error
    CheckResponseForErrors = True
    
    Dim msg As String
    msg = "vBulletin responded with error '" & lResponse("errors")(1)(1) & "'." & vbCrLf & vbCrLf
    
    Select Case lResponse("errors")(1)(1)
    Case "invalid_clientid"
        mClientId = ""
        mAccessToken = ""
        mLoggedIn = False
        msg = msg & "The client id is no longer valid and was deleted. Please retry."
    Case "invalid_accesstoken"
        mAccessToken = ""
        msg = msg & "The access token has expired. Please retry."
    Case "invalid_api_signature"
        msg = msg & "The call signature is invalid. This must be a software bug. Feel free to debug the VBA code yourself."
    Case "missing_api_signature"
        msg = msg & "The call signature is mssing. This must be a software bug. Feel free to debug the VBA code yourself."
    Case "bbclosed"
        msg = msg & "Forum is closed."
    Case "toobusy"
        msg = msg & "Forum is temporarily closed due to server load average."
    Case "badlogin_logintypeusername"
        msg = msg & "Username or password is incorrect."
    Case "badlogin_strikes_logintypeusername"
        msg = msg & "Username or password is incorrect. Too many failed login attempts. Try again in 15 minutes."
    Case "invalid_user_specified"
        msg = msg & "User not found."
    Case "bademail"
        msg = msg & "The given email address is invalid."
    Case Else
        msg = msg & "Error unknown."
    End Select
    
    MsgBox msg
End Function