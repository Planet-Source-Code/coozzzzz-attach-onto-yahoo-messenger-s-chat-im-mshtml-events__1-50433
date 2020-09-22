VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This concept was part of a larger one I was working on for Yahoo! Messenger
'which I grew bored of. Reason for posting it on PSC now in case others find
'interest in it. I originally began commenting each line but it seems best that
'I provide a resource for the entire DOM structure (MSHTML) for you to better
'understand the capabilities that this allows.
'
'1.) Project -> References -> Microsoft HTML Object Library
'2.) View -> Object Browser -> MSHTML
'3.) You should have all you need there. Go to MSDN if you need further help
'    on properties and methods listed.
Option Explicit

Private Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2

Private Declare Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, riid As UUID, ByVal wParam As Long, ppvObject As Any) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public WithEvents yDocument As HTMLDocument
Attribute yDocument.VB_VarHelpID = -1
Public WithEvents yDiv As HTMLDivElement
Attribute yDiv.VB_VarHelpID = -1
Public WithEvents yTable As HTMLTable
Attribute yTable.VB_VarHelpID = -1

Public yMessageType As String
Public Function Attach(ByVal strWindowTitle As String) As Boolean
    Dim lTemp1 As String, lTemp2 As Long, sTemp1 As String
    lTemp1 = FindWindow("IMClass", strWindowTitle)
    If lTemp1 > 0 Then
        lTemp2 = return_IEWindow(lTemp1)
        If lTemp2 > 0 Then
            Set yDocument = return_Document(lTemp2)
            If TypeName(yDocument) <> "Nothing" Then
                Set yDiv = yDocument.All.tags("div")(0)
                If TypeName(yDiv) <> "Nothing" Then
                    Select Case LCase(yDiv.id)
                        Case "imbody"
                            yMessageType = "chat"
                        Case "$im"
                            yMessageType = "im"
                        Case Else
                            Exit Function
                    End Select
                    Set yTable = yDocument.All.tags("table")(0)
                    If TypeName(yTable) = "Nothing" Then
                        sTemp1 = yDiv.innerHTML
                        yDiv.innerHTML = "<table>" & sTemp1 & "</table>"
                    End If
                    Set yTable = yDocument.All.tags("table")(0)
                    If TypeName(yTable) <> "Nothing" Then Attach = True
                End If
            End If
        End If
    End If
End Function
Public Sub ChangeBackgroundImage(ByVal strLocation As String)
    yDocument.body.Style.backgroundImage = "url(" & strLocation & ")"
    'Just to keep it clean
    yDocument.body.Style.backgroundRepeat = "no-repeat"
    yDocument.body.Style.backgroundPosition = "bottom right"
End Sub
Public Sub ChangeBackgroundColor(ByVal strColor As String)
    yDocument.body.Style.backgroundColor = strColor
End Sub
Private Function return_WindowClass(ByVal lHandle As Long) As String
    'Function Return: Window's Class Name
    Dim sBuff As String, iLen As Integer
    'Create a string buffer to store the window's class name
    sBuff = Space(256)
    'Call the GetClassName function which returns the length of the class name
    'which is buffered in sBuff
    iLen = GetClassName(lHandle, sBuff, 256)
    'Return the class name by length to avoid the spaces from Space(256)
    return_WindowClass = Left(sBuff, iLen)
End Function
Private Function return_IEWindow(ByVal lHandle As Long) As Long
    'Function Return: Handle to a possible "Internet Explorer_Server" Window
    Dim lTemp As Long
    'Get the child of the provided window handle (lHandle)
    lTemp = GetWindow(lHandle, GW_CHILD)
    'Loop until the child window of the current window is what we're looking for
    'or if our current window does not exist (lTemp=0)
    Do Until return_WindowClass(GetWindow(lTemp, GW_CHILD)) = "Internet Explorer_Server" Or lTemp = 0
        'Get the next window in relation to our current window (iTemp)
        lTemp = GetWindow(lTemp, GW_HWNDNEXT)
    Loop
    'Loop has finished, if we found our window...
    If lTemp > 0 Then
        'Return it to our function
        return_IEWindow = GetWindow(lTemp, GW_CHILD)
    Else
        'If not, return 0.. (this should be default anyways)
        return_IEWindow = 0
    End If
End Function
Private Function return_Document(ByVal lHandle As Long) As IHTMLDocument
    'Originally from (includes details):
    'http://www.mvps.org/emorcillo/cod/inet/iedom_vb6.htm
    Dim IID_IHTMLDocument As UUID, lRes As Long, lMsg As Long
    lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")
    Call SendMessageTimeout(lHandle, lMsg, 0, 0, &H2, 1000, lRes)
    If lRes Then
        With IID_IHTMLDocument
            .Data1 = &H626FC520
            .Data2 = &HA41E
            .Data3 = &H11CF
            .Data4(0) = &HA7
            .Data4(1) = &H31
            .Data4(2) = &H0
            .Data4(3) = &HA0
            .Data4(4) = &HC9
            .Data4(5) = &H8
            .Data4(6) = &H26
            .Data4(7) = &H37
        End With
        Call ObjectFromLresult(lRes, IID_IHTMLDocument, 0, return_Document)
    End If
End Function
Private Sub yTable_onresize()
    Dim strUsername As String, strMessage As String, strAdditional As String
    Dim spanElement1 As HTMLSpanElement, spanElement2 As HTMLSpanElement, aElement1 As HTMLAnchorElement
    If yMessageType = "im" Then
        Set spanElement1 = yDocument.All.tags("span")(yDocument.All.tags("span").length - 2)
        Set spanElement2 = yDocument.All.tags("span")(yDocument.All.tags("span").length - 1)
        If TypeName(spanElement1.getAttribute("n")) <> "Null" Then
            If TypeName(spanElement1) <> "Nothing" And TypeName(spanElement2) <> "Nothing" Then
                strUsername = spanElement1.getAttribute("n")
                strMessage = spanElement2.outerText
                Debug.Print strUsername & ":" & strMessage
                'Processing of Instant Messages, for now, change the background color
                spanElement2.Style.backgroundColor = "#c3c3c3"
            End If
        Else
            'Processing of Other Instant Messages
            strAdditional = spanElement2.outerText
        End If
    ElseIf yMessageType = "chat" Then
        Set aElement1 = yDocument.All.tags("a")(yDocument.All.tags("a").length - 1)
        Set spanElement2 = yDocument.All.tags("span")(yDocument.All.tags("span").length - 1)
        If TypeName(aElement1.getAttribute("username")) <> "Null" Then
            If TypeName(aElement1) <> "Nothing" And TypeName(spanElement2) <> "Nothing" Then
                strUsername = aElement1.getAttribute("username")
                strMessage = spanElement2.outerText
                Debug.Print strUsername & ":" & strMessage
                'Processing of Chat Messages, for now, change the background color
                spanElement2.Style.backgroundColor = "#c3c3c3"
            End If
        Else
            'Processing of Other Chat Messages
            strAdditional = spanElement2.outerText
        End If
    End If
End Sub
