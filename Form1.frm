VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Chingutalk 
      Caption         =   "친구톡 발송"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Alimtalk 
      Caption         =   "알림톡 발송"
      Height          =   495
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton MMS 
      Caption         =   "MMS 발송"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Balance 
      Caption         =   "잔액조회"
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton LMS 
      Caption         =   "LMS 발송"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton MessageList 
      Caption         =   "목록조회"
      Height          =   495
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton SMS 
      Caption         =   "SMS 발송"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "요청 결과가 디버그 창에 출력됩니다."
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Alimtalk_Click(Index As Integer)
    ' 텍스트 내용만 있는 간단한 알림톡
    Dim kakaoOptions1 As New Dictionary
    kakaoOptions1.Add "pfId", "KA01PF190626020502205cl0mYSoplA2"  ' 카카오톡채널 연동 후 발급받은 값을 사용해 주세요
    kakaoOptions1.Add "templateId", "KA01TP190626032036196g86q1RGN7D1"  ' 템플릿 등록 후 발급받은 값을 사용해 주세요
    
    Dim msg1 As New Dictionary
    msg1.Add "to", "01000000001"
    msg1.Add "from", "029302266"
    msg1.Add "text", "안녕하세요." & vbLf & "홍길동님 회원가입을 환영합니다."
    msg1.Add "kakaoOptions", kakaoOptions1
    
    ' 웹 링크 버튼이 하나 있는 알림톡
    Dim button1 As New Dictionary
    button1.Add "buttonName", "시작하기"
    button1.Add "buttonType", "WL"
    button1.Add "linkMo", "https://example.com"
    
    Dim kakaoOptions2 As New Dictionary
    kakaoOptions2.Add "pfId", "KA01PF190626020502205cl0mYSoplA2"
    kakaoOptions2.Add "templateId", "KA01TP190626032036196g86q1RGN7D2"
    kakaoOptions2.Add "buttons", Array(button1)
    
    Dim msg2 As New Dictionary
    msg2.Add "to", "01000000002"
    msg2.Add "from", "029302266"
    msg2.Add "text", "안녕하세요." & vbLf & "홍길동님 회원가입을 환영합니다."
    msg2.Add "kakaoOptions", kakaoOptions2
    

    ' 모든 종류의 버튼 예시
    Dim button2 As New Dictionary
    button2.Add "buttonName", "시작하기"
    button2.Add "buttonType", "WL"
    button2.Add "linkMo", "https://m.example.com"
    button2.Add "linkPc", "https://example.com"
    
    Dim button3 As New Dictionary
    button3.Add "buttonName", "앱실행"
    button3.Add "buttonType", "AL"
    button3.Add "linkAnd", "examplescheme://"  ' 안드로이드
    button3.Add "linkIos", "examplescheme://"  ' iOS
    
    Dim button4 As New Dictionary
    button4.Add "buttonName", "배송조회"
    button4.Add "buttonType", "DS"
    
    Dim button5 As New Dictionary
    button5.Add "buttonName", "봇키워드"
    button5.Add "buttonType", "BK" ' 챗봇에게 키워드를 전달합니다. 버튼이름의 키워드가 그대로 전달됩니다.
    
    Dim button6 As New Dictionary
    button6.Add "buttonName", "상담요청하기"
    button6.Add "buttonType", "MD" ' 상담요청하기 버튼을 누르면 수신 받은 알림톡 메시지가 상담원에게 그대로 전달됩니다.
    
    Dim kakaoOptions3 As New Dictionary
    kakaoOptions3.Add "pfId", "KA01PF190626020502205cl0mYSoplA2"
    kakaoOptions3.Add "templateId", "KA01TP190626032036196g86q1RGN7D3"
    kakaoOptions3.Add "buttons", Array(button2, button3, button4, button5, button6)
    
    Dim msg3 As New Dictionary
    msg3.Add "to", "01000000003"
    msg3.Add "from", "029302266"
    msg3.Add "text", "안녕하세요." & vbLf & "홍길동님 회원가입을 환영합니다." & vbLf & "아래 다양한 형식의 버튼을 통해 사용방법을 익히실 수 있습니다."
    msg3.Add "kakaoOptions", kakaoOptions3


    ' 1만건까지 추가 가능
    Dim Messages
    Messages = Array(msg1, msg2, msg3)
    
    Dim Body As New Dictionary
    Body.Add "messages", Messages
    
    Dim Response As WebResponse
    Set Response = Request("messages/v4/send-many", "POST", Body)
    If Response.StatusCode = WebStatusCode.Ok Then
        Debug.Print ("발송성공")
        ' JSON Object로 접근
        Debug.Print ("Group ID:" & Response.data("groupId"))
        Debug.Print ("Status:" & Response.data("status"))
        Debug.Print ("Total Count:" & Response.data("count")("total"))
        ' String 형식으로 접근 모든 내용 출력
        Debug.Print (Response.Content)
    Else
        Debug.Print ("발송실패")
        Debug.Print (Response.Content)
    End If
End Sub

Private Sub Balance_Click(Index As Integer)
    Dim Response As WebResponse
    Set Response = Request("cash/v1/balance", "GET")
    If Response.StatusCode = WebStatusCode.Ok Then
        ' JSON Object로 접근
        Debug.Print ("현재 잔액:" & Response.data("balance"))
        Debug.Print ("현재 포인트:" & Response.data("point"))
        ' String 형식으로 접근 모든 내용 출력
        Debug.Print (Response.Content)
        MsgBox ("현재 잔액은 " & Response.data("balance") & "원 그리고 " & Response.data("point") & "포인트가 남아있습니다.")
    Else
        Debug.Print ("잔액조회실패")
        Debug.Print (Response.Content)
    End If
End Sub

Private Sub Chingutalk_Click(Index As Integer)
    ' 텍스트 내용만 있는 간단한 친구톡
    Dim kakaoOptions1 As New Dictionary
    kakaoOptions1.Add "pfId", "KA01PF190626020502205cl0mYSoplA2"  ' 카카오톡채널 연동 후 발급받은 값을 사용해 주세요
    
    Dim msg1 As New Dictionary
    msg1.Add "to", "01000000001"
    msg1.Add "from", "029302266"
    msg1.Add "text", "광고를 포함하여 어떤 내용이든 입력 가능합니다."
    msg1.Add "kakaoOptions", kakaoOptions1
    
    ' 모든 종류의 버튼 예시
    Dim button1 As New Dictionary
    button1.Add "buttonName", "시작하기"
    button1.Add "buttonType", "WL"
    button1.Add "linkMo", "https://m.example.com"   ' 모바일 기기에서 보여지는 링크
    button1.Add "linkPc", "https://example.com"     ' PC에서 보여지는 링크
    
    Dim button2 As New Dictionary
    button2.Add "buttonName", "앱실행"
    button2.Add "buttonType", "AL"
    button2.Add "linkAnd", "https://example.com"  ' 안드로이드
    button2.Add "linkIos", "https://example.com"  ' iOS
    
    Dim button3 As New Dictionary
    button3.Add "buttonName", "봇키워드"
    button3.Add "buttonType", "BK" ' 챗봇에게 키워드를 전달합니다. 버튼이름의 키워드가 그대로 전달됩니다.
    
    Dim button4 As New Dictionary
    button4.Add "buttonName", "상담요청하기"
    button4.Add "buttonType", "MD" ' 상담요청하기 버튼을 누르면 수신 받은 알림톡 메시지가 상담원에게 그대로 전달됩니다.
    
    Dim kakaoOptions2 As New Dictionary
    kakaoOptions2.Add "pfId", "KA01PF190626020502205cl0mYSoplA2"
    kakaoOptions2.Add "buttons", Array(button1, button2, button3, button4)
    
    Dim msg2 As New Dictionary
    msg2.Add "to", "01000000002"
    msg2.Add "from", "029302266"
    msg2.Add "text", "광고를 포함하여 어떤 내용이든 입력 가능합니다."
    msg2.Add "kakaoOptions", kakaoOptions2
    
    
    ' 친구톡에 사용할 이미지 업로드
    Dim resp As WebResponse
    Dim imageId As String
    Set resp = UploadKakaoImage("testImage.jpg", "https://example.com")
    If resp.StatusCode <> 200 Then
        Debug.Print ("이미지 업로드 실패")
        Exit Sub
    End If
    Debug.Print (resp.Content)
    imageId = resp.data("fileId")
    
    ' 친구톡 이미지 발송
    Dim kakaoOptions3 As New Dictionary
    kakaoOptions3.Add "pfId", "KA01PF190626020502205cl0mYSoplA2"
    kakaoOptions3.Add "imageId", imageId
    
    Dim msg3 As New Dictionary
    msg3.Add "to", "01000000003"
    msg3.Add "from", "029302266"
    msg3.Add "text", "광고를 포함하여 어떤 내용이든 입력 가능합니다." & vbLf & "이미지를 터치하면 URL로 이동됩니다."
    msg3.Add "kakaoOptions", kakaoOptions3
    

    ' 1만건까지 추가 가능
    Dim Messages
    Messages = Array(msg1, msg2, msg3)
    
    Dim Body As New Dictionary
    Body.Add "messages", Messages
    
    Dim Response As WebResponse
    Set Response = Request("messages/v4/send-many", "POST", Body)
    If Response.StatusCode = WebStatusCode.Ok Then
        Debug.Print ("발송성공")
        ' JSON Object로 접근
        Debug.Print ("Group ID:" & Response.data("groupId"))
        Debug.Print ("Status:" & Response.data("status"))
        Debug.Print ("Total Count:" & Response.data("count")("total"))
        ' String 형식으로 접근 모든 내용 출력
        Debug.Print (Response.Content)
    Else
        Debug.Print ("발송실패")
        Debug.Print (Response.Content)
    End If
End Sub

Private Sub MessageList_Click(Index As Integer)
    Dim Response As WebResponse
    Set Response = Request("messages/v4/list", "GET")
    Debug.Print (Response.StatusDescription)
    Dim line As Integer
    Dim indent As Integer
    Debug.Print (Response.Content)
End Sub


Private Sub SMS_Click()
    Dim msg1 As New Dictionary
    msg1.Add "to", "01000000001"
    msg1.Add "from", "029302266"
    msg1.Add "text", "한글 45자, 영자 90자 이하 입력되면 자동으로 SMS타입의 메시지가 추가됩니다."
    
    Dim msg2 As New Dictionary
    msg2.Add "to", "01000000002"
    msg2.Add "from", "029302266"
    msg2.Add "text", "한글 45자, 영자 90자 이상 입력되면 자동으로 LMS타입의 문자메시자가 발송됩니다. 0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    ' 타입을 명시할 경우 text 길이가 한글 45 혹은 영자 90자를 넘을 경우 오류가 발생합니다.
    Dim msg3 As New Dictionary
    msg3.Add "type", "SMS" ' 타입을 SMS로 입력
    msg3.Add "to", "01000000003"
    msg3.Add "from", "029302266"
    msg3.Add "text", "SMS 타입에 한글 45자, 영자 90자 이상 입력되면 오류가 발생합니다. 0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    ' 1만건까지 추가 가능
    Dim Messages
    Messages = Array(msg1, msg2, msg3)
    
    Dim Body As New Dictionary
    Body.Add "messages", Messages
    
    Dim Response As WebResponse
    Set Response = Request("messages/v4/send-many", "POST", Body)
    If Response.StatusCode = WebStatusCode.Ok Then
        Debug.Print ("발송성공")
        ' JSON Object로 접근
        Debug.Print ("Group ID:" & Response.data("groupId"))
        Debug.Print ("Status:" & Response.data("status"))
        Debug.Print ("Total Count:" & Response.data("count")("total"))
        ' String 형식으로 접근 모든 내용 출력
        Debug.Print (Response.Content)
    Else
        Debug.Print ("발송실패")
        Debug.Print (Response.Content)
    End If
End Sub


Private Sub LMS_Click(Index As Integer)
    Dim msg1 As New Dictionary
    msg1.Add "to", "01000000001"
    msg1.Add "from", "029302266"
    msg1.Add "text", "한글 45자, 영자 90자 이상 입력되면 자동으로 LMS타입의 문자메시자가 발송됩니다. 0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    Dim msg2 As New Dictionary
    msg2.Add "to", "01000000002"
    msg2.Add "from", "029302266"
    msg2.Add "subject", "LMS 제목" ' 제목을 지정할 수 있습니다.
    msg2.Add "text", "한글 45자, 영자 90자 이상 입력되면 자동으로 LMS타입의 문자메시자가 발송됩니다. 0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    Dim msg3 As New Dictionary
    msg3.Add "type", "LMS" ' 타입을 명시할 수 있습니다.
    msg3.Add "to", "01000000003"
    msg3.Add "from", "029302266"
    msg3.Add "text", "내용이 짧아도 LMS로 발송됩니다."
   
    Dim msg4 As New Dictionary
    msg4.Add "to", "01000000004"
    msg4.Add "from", "029302266"
    msg4.Add "text", "한글 45자, 영자 90자 이하는 자동으로 SMS타입의 문자가 발송됩니다."

    ' 1만건까지 추가 가능
    Dim Messages
    Messages = Array(msg1, msg2, msg3, msg4)
    
    Dim Body As New Dictionary
    Body.Add "messages", Messages
    
    Dim Response As WebResponse
    Set Response = Request("messages/v4/send-many", "POST", Body)
    If Response.StatusCode = WebStatusCode.Ok Then
        Debug.Print ("발송성공")
        ' JSON Object로 접근
        Debug.Print ("Group ID:" & Response.data("groupId"))
        Debug.Print ("Status:" & Response.data("status"))
        Debug.Print ("Total Count:" & Response.data("count")("total"))
        ' String 형식으로 접근 모든 내용 출력
        Debug.Print (Response.Content)
    Else
        Debug.Print ("발송실패")
        Debug.Print (Response.Content)
    End If

End Sub


Private Sub MMS_Click(Index As Integer)
    On Error GoTo Handler
    Dim resp As WebResponse
    Dim imageId As String
    Set resp = UploadImage("testImage.jpg")
    If resp.StatusCode <> 200 Then
        Exit Sub
    End If
    
    Debug.Print (resp.Content)
    imageId = resp.data("fileId")
    
    Dim msg1 As New Dictionary
    msg1.Add "to", "01000000001"
    msg1.Add "from", "029302266"
    msg1.Add "subject", "MMS 제목"
    msg1.Add "text", "이미지 아이디가 입력되면 MMS로 발송됩니다."
    msg1.Add "imageId", imageId
    
    Dim msg2 As New Dictionary
    msg2.Add "to", "01000000002"
    msg2.Add "from", "029302266"
    msg2.Add "subject", "MMS 제목"
    msg2.Add "text", "동일한 이미지 아이디가 입력되면 동일한 이미지가 MMS로 발송됩니다."
    msg2.Add "imageId", imageId
    

    ' 1만건까지 추가 가능
    Dim Messages
    Messages = Array(msg1, msg2)
    
    Dim Body As New Dictionary
    Body.Add "messages", Messages
    
    Dim Response As WebResponse
    Set Response = Request("messages/v4/send-many", "POST", Body)
    If Response.StatusCode = WebStatusCode.Ok Then
        Debug.Print ("발송성공")
        ' JSON Object로 접근
        Debug.Print ("Group ID:" & Response.data("groupId"))
        Debug.Print ("Status:" & Response.data("status"))
        Debug.Print ("Total Count:" & Response.data("count")("total"))
        ' String 형식으로 접근 모든 내용 출력
        Debug.Print (Response.Content)
    Else
        Debug.Print ("발송실패")
        Debug.Print (Response.Content)
    End If
Handler:
    Debug.Print "Error " & Err.Number & Err.Source & Err.Description
End Sub

