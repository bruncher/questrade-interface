Public Sub GetQuestradeData()

Application.ScreenUpdating = False

' connect to get authorization
Dim strHost As String
strHost = "https://login.questrade.com/oauth2/token?grant_type=refresh_token&refresh_token="

' token can be manually generated on Questrade and copied to cell B2
Dim strToken As String
strToken = Range("B2").Value

Dim strURL As String
strURL = strHost & strToken

' open and send http request to get authrorization
Dim httpReq As New WinHttpRequest
Dim strResp As String

httpReq.Open "GET", strURL, False
httpReq.Send

' get the response
strResp = httpReq.ResponseText

Debug.Print strResp

' find the refresh token
Dim strRefresh As String
Dim numParse As Integer

numParse = InStr(strResp, "refresh_token")
If numParse > 0 Then
    strRefresh = Mid(strResp, numParse + 16, 33)
Else
    strRefresh = ""
End If

' update the token in spreadsheet
Range("B2").Value = strRefresh

' exit subroutine if the request didn't work
If strRefresh = "" Then
    Debug.Print "Invalid request, exiting subroutine"
    Exit Sub
End If

' get the access token and update spreadsheet
Dim strAuth As String

numParse = InStr(strResp, "access_token")
If numParse > 0 Then
    strAuth = Mid(strResp, numParse + 15, 33)
Else
    strAuth = ""
End If

Range("B6").Value = strAuth

' get the api server and update spreadsheet
Dim strServer As String

numParse = InStr(strResp, "api_server")
If numParse > 0 Then
    strServer = Mid(strResp, numParse + 23, 5)
    strHost = "https://" & strServer & ".iq.questrade.com"
Else
    strServer = ""
    strHost = ""
End If

Range("B4").Value = strHost

' use the authorization code to get data
Dim strCommand As String
Dim strSubcom As String

' get basic account info
strCommand = "/v1/accounts"

strURL = strHost & strCommand

httpReq.Open "GET", strURL, False

httpReq.SetRequestHeader "Authorization", "Bearer " & strAuth

httpReq.Send

strResp = httpReq.ResponseText

Debug.Print strResp

' update account info in spreadsheet
Dim strType As String
Dim strAccount As String
Dim numParseEnd As Integer

' get account type
numParse = InStr(strResp, "type")
If numParse > 0 Then
    numParseEnd = InStr(Mid(strResp, numParse + 7), Chr(34))
    If numParseEnd > 0 Then
        strType = Mid(strResp, numParse + 7, numParseEnd - 1)
    Else
        strType = ""
    End If
Else
    strType = ""
End If

Range("B10").Value = strType

' get account number
numParse = InStr(strResp, "number")
If numParse > 0 Then
    numParseEnd = InStr(Mid(strResp, numParse + 9), Chr(34))
    If numParseEnd > 0 Then
        strAccount = "/" & Mid(strResp, numParse + 9, numParseEnd - 1)
    Else
        strAccount = ""
    End If
Else
    strAccounte = ""
End If

Range("B8").Value = strAccount

' get account balances
strSubcom = "/balances"

strURL = strHost & strCommand & strAccount & strSubcom

httpReq.Open "GET", strURL, False

httpReq.SetRequestHeader "Authorization", "Bearer " & strAuth

httpReq.Send

strResp = httpReq.ResponseText

Debug.Print strResp

' extract and print cash
Dim cashCAD, cashUSD As Variant

' extract CAD cash
numParse = InStr(strResp, "CAD")
If numParse > 0 Then
    numParseEnd = InStr(Mid(strResp, numParse + 12), ",")
    If numParseEnd > 0 Then
        cashCAD = Mid(strResp, numParse + 12, numParseEnd - 1)
    Else
        cashCAD = ""
    End If
Else
    cashCAD = ""
End If

Range("B12").Value = cashCAD

' extract USD cash
numParse = InStr(strResp, "USD")
If numParse > 0 Then
    numParseEnd = InStr(Mid(strResp, numParse + 12), ",")
    If numParseEnd > 0 Then
        cashUSD = Mid(strResp, numParse + 12, numParseEnd - 1)
    Else
        cashUSD = ""
    End If
Else
    cashUSD = ""
End If

Range("B14").Value = cashUSD

' get account positions
strSubcom = "/positions"

strURL = strHost & strCommand & strAccount & strSubcom

httpReq.Open "GET", strURL, False

httpReq.SetRequestHeader "Authorization", "Bearer " & strAuth

httpReq.Send

strResp = httpReq.ResponseText

Debug.Print strResp

End Sub
