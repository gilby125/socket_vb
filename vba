Option Explicit
Option Private Module
 
 
'reference Windows Sockets 2 - Windows applications _ Microsoft Docs
'http://msdn.microsoft.com/en-us/library/windows/desktop/ms740673(v=vs.85).aspx
Private Const INVALID_SOCKET = -1
Private Const WSADESCRIPTION_LEN = 256
Private Const SOCKET_ERROR As Long = -1 'const #define SOCKET_ERROR            (-1)
 
Private Enum AF
    AF_UNSPEC = 0
    AF_INET = 2
    AF_IPX = 6
    AF_APPLETALK = 16
    AF_NETBIOS = 17
    AF_INET6 = 23
    AF_IRDA = 26
    AF_BTH = 32
End Enum
 
Private Enum sock_type
    SOCK_STREAM = 1
    SOCK_DGRAM = 2
    SOCK_RAW = 3
    SOCK_RDM = 4
    SOCK_SEQPACKET = 5
End Enum
 
Private Enum Protocol
    IPPROTO_ICMP = 1
    IPPROTO_IGMP = 2
    BTHPROTO_RFCOMM = 3
    IPPROTO_TCP = 6
    IPPROTO_UDP = 17
    IPPROTO_ICMPV6 = 58
    IPPROTO_RM = 113
End Enum
 
'Private Type sockaddr
'    sa_family As Integer
'    sa_data(0 To 13) As Byte
'End Type
 
Private Type sockaddr_in
    sin_family As Integer
    sin_port(0 To 1) As Byte
    sin_addr(0 To 3) As Byte
    sin_zero(0 To 7) As Byte
End Type
 
 
'typedef UINT_PTR        SOCKET;
Private Type udtSOCKET
    pointer As Long
End Type
 
 
 
' typedef struct WSAData {
'  WORD           wVersion;
'  WORD           wHighVersion;
'  char           szDescription[WSADESCRIPTION_LEN+1];
'  char           szSystemStatus[WSASYS_STATUS_LEN+1];
'  unsigned short iMaxSockets;
'  unsigned short iMaxUdpDg;
'  char FAR       *lpVendorInfo;
'} WSADATA, *LPWSADATA;
 
Private Type udtWSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSADESCRIPTION_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
 
'int errorno = WSAGetLastError()
Private Declare Function WSAGetLastError Lib "Ws2_32" () As Integer
 
'   int WSAStartup(
'  __in   WORD wVersionRequested,
'  __out  LPWSADATA lpWSAData
');
Private Declare Function WSAStartup Lib "Ws2_32" _
    (ByVal wVersionRequested As Integer, ByRef lpWSAData As udtWSADATA) As winsockErrorCodes2
 
 
'    SOCKET WSAAPI socket(
'  __in  int af,
'  __in  int type,
'  __in  int protocol
');
 
Private Declare Function ws2_socket Lib "Ws2_32" Alias "socket" _
    (ByVal AF As Long, ByVal stype As Long, ByVal Protocol As Long) As LongPtr
 
Private Declare Function ws2_closesocket Lib "Ws2_32" Alias "closesocket" _
    (ByVal socket As Long) As Long
 
'int recv(
'  SOCKET s,
'  char   *buf,
'  int    len,
'  int    flags
');
Private Declare Function ws2_recv Lib "Ws2_32" Alias "recv" _
    (ByVal socket As Long, ByVal buf As LongPtr, _
     ByVal length As Long, ByVal flags As Long) As Long
 
'int WSAAPI connect(
'  SOCKET         s,
'  const sockaddr *name,
'  int            namelen
');
 
Private Declare Function ws2_connect Lib "Ws2_32" Alias "connect" _
    (ByVal S As LongPtr, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
 
'int WSAAPI send(
'  SOCKET     s,
'  const char *buf,
'  int        len,
'  int        flags
');
Private Declare Function ws2_send Lib "Ws2_32" Alias "send" _
    (ByVal S As LongPtr, ByVal buf As LongPtr, ByVal buflen As Long, ByVal flags As Long) As Long
 
 
Private Declare Function ws2_shutdown Lib "Ws2_32" Alias "shutdown" _
        (ByVal S As Long, ByVal how As Long) As Long
 
Private Declare Sub WSACleanup Lib "Ws2_32" ()
 
Private Enum eShutdownConstants
    SD_RECEIVE = 0  '#define SD_RECEIVE      0x00
    SD_SEND = 1     '#define SD_SEND         0x01
    SD_BOTH = 2     '#define SD_BOTH         0x02
End Enum
 
Private Sub TestWS2SendAndReceive()
 
    Dim sResponse As String
    If WS2SendAndReceive("KEYS *" & vbCrLf, sResponse) Then
        Debug.Print VBA.Join(RedisResponseToTypedVariable(sResponse), ";")
    End If
 
    If WS2SendAndReceive("GET foo" & vbCrLf, sResponse) Then
        Debug.Print RedisResponseToTypedVariable(sResponse)
    End If
 
    If WS2SendAndReceive("SET baz Barry" & vbCrLf, sResponse) Then
        Debug.Assert RedisResponseToTypedVariable(sResponse) = "OK"
    End If
 
    If WS2SendAndReceive("SET foo BAR" & vbCrLf, sResponse) Then
        Debug.Assert RedisResponseToTypedVariable(sResponse) = "OK"
    End If
 
    If WS2SendAndReceive("DEL baz" & vbCrLf, sResponse) Then
        Debug.Print RedisResponseToTypedVariable(sResponse)
    End If
 
 
    If WS2SendAndReceive("GET baz" & vbCrLf, sResponse) Then
        Debug.Print RedisResponseToTypedVariable(sResponse)
    End If
 
    If WS2SendAndReceive("GET foo" & vbCrLf, sResponse) Then
        Debug.Print RedisResponseToTypedVariable(sResponse)
    End If
 
 
    If WS2SendAndReceive("KEYS *" & vbCrLf, sResponse) Then
        Debug.Print VBA.Join(RedisResponseToTypedVariable(sResponse), ";")
    End If
 
 
    If WS2SendAndReceive("SET count 0" & vbCrLf, sResponse) Then
        Debug.Assert RedisResponseToTypedVariable(sResponse) = "OK"
    End If
 
    If WS2SendAndReceive("INCR count" & vbCrLf, sResponse) Then
        Debug.Assert RedisResponseToTypedVariable(sResponse) = "1"
    End If
 
 
 
 
End Sub
 
Public Function WS2SendAndReceive(ByVal sCommand As String, ByRef psResponse As String) As Boolean
    'https://docs.microsoft.com/en-gb/windows/desktop/api/winsock/nf-winsock-recv
 
    psResponse = ""
    '//----------------------
    '// Declare and initialize variables.
    Dim iResult As Integer : iResult = 0
    Dim wsaData As udtWSADATA
 
    Dim ConnectSocket As LongPtr
 
    Dim clientService As sockaddr_in
 
    Dim sendBuf() As Byte
    sendBuf = StrConv(sCommand, vbFromUnicode)
 
    Const recvbuflen As Long = 512
    Dim recvbuf(0 To recvbuflen - 1) As Byte
 
    '//----------------------
    '// Initialize Winsock
    Dim eResult As winsockErrorCodes2
    eResult = WSAStartup(&H202, wsaData)
    If eResult <> 0 Then
        Debug.Print "WSAStartup failed with error: " & eResult
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
 
    '//----------------------
    '// Create a SOCKET for connecting to server
    ConnectSocket = ws2_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If ConnectSocket = INVALID_SOCKET Then
        Dim eLastError As winsockErrorCodes2
        eLastError = WSAGetLastError()
        Debug.Print "socket failed with error: " & eLastError
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
 
    '//----------------------
    '// The sockaddr_in structure specifies the address family,
    '// IP address, and port of the server to be connected to.
    clientService.sin_family = AF_INET
 
    clientService.sin_addr(0) = 127
    clientService.sin_addr(1) = 0
    clientService.sin_addr(2) = 0
    clientService.sin_addr(3) = 1
 
    clientService.sin_port(1) = 235 '* 6379
    clientService.sin_port(0) = 24
 
    '//----------------------
    '// Connect to server.
 
    iResult = ws2_connect(ConnectSocket, clientService, LenB(clientService))
    If (iResult = SOCKET_ERROR) Then
 
        eLastError = WSAGetLastError()
 
        Debug.Print "connect failed with error: " & eLastError
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    '//----------------------
    '// Send an initial buffer
    Dim sendbuflen As Long
    sendbuflen = UBound(sendBuf) - LBound(sendBuf) + 1
    iResult = ws2_send(ConnectSocket, VarPtr(sendBuf(0)), sendbuflen, 0)
    If (iResult = SOCKET_ERROR) Then
        eLastError = WSAGetLastError()
        Debug.Print "send failed with error: " & eLastError
 
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    'Debug.Print "Bytes Sent: ", iResult
 
    '// shutdown the connection since no more data will be sent
    iResult = ws2_shutdown(ConnectSocket, SD_SEND)
    If (iResult = SOCKET_ERROR) Then
 
        eLastError = WSAGetLastError()
        Debug.Print "shutdown failed with error: " & eLastError
 
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    ' receive only one message (TODO handle when buffer is not large enough)
 
    iResult = ws2_recv(ConnectSocket, VarPtr(recvbuf(0)), recvbuflen, 0)
    If (iResult > 0) Then
        'Debug.Print "Bytes received: ", iResult
    ElseIf (iResult = 0) Then
        Debug.Print "Connection closed"
        WS2SendAndReceive = False
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        GoTo SingleExit
    Else
        eLastError = WSAGetLastError()
        Debug.Print "recv failed with error: " & eLastError
    End If
 
    psResponse = Left$(StrConv(recvbuf, vbUnicode), iResult)
 
    'Debug.Print psResponse
 
    '// close the socket
    iResult = ws2_closesocket(ConnectSocket)
    If (iResult = SOCKET_ERROR) Then
 
        eLastError = WSAGetLastError()
        Debug.Print "close failed with error: " & eLastError
 
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    Call WSACleanup
    WS2SendAndReceive = True
 
SingleExit:
    Exit Function
ErrHand:
 
End Function
 
Public Function RedisResponseToTypedVariable(ByVal sResponse As String)
 
    Dim lTotalLength As Long
    lTotalLength = Len(sResponse)
    Debug.Assert lTotalLength > 0
 
    Dim vSplitResponse As Variant
    vSplitResponse = VBA.Split(sResponse, vbCrLf)
 
    Dim lReponseLineCount As Long
    lReponseLineCount = UBound(vSplitResponse) - LBound(vSplitResponse)
 
    Select Case Left(sResponse, 1)
        Case "$"
 
            RedisResponseToTypedVariable = vSplitResponse(1)
 
        Case "+"
            RedisResponseToTypedVariable = Mid$(vSplitResponse(0), 2)
 
        Case ":"
            '* response is an integer
            RedisResponseToTypedVariable = CLng(Mid$(vSplitResponse(0), 2))
 
        Case "-"
            '* response is an error
            Err.Raise vbObjectError, , Mid$(vSplitResponse(0), 2)
        'Stop
        Case "*"
            '* multiple responses, build an array to return
            Dim lResponseCount As Long
            lResponseCount = CLng(Mid$(vSplitResponse(0), 2))
            If lResponseCount > 0 Then
                Debug.Assert lResponseCount = (lReponseLineCount - 1) / 2
        ReDim vReturn(0 To lResponseCount - 1)
                Dim lLoop As Long
                For lLoop = 0 To lResponseCount - 1
                    vReturn(lLoop) = vSplitResponse((lLoop + 1) * 2)
                Next lLoop
            End If
            RedisResponseToTypedVariable = vReturn
 
        Case Else
            Stop  '* this should not happen
    End Select
 
End Function
 
