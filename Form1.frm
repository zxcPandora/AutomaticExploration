VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "POST"
   ClientHeight    =   9060
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   13815
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command16 
      Caption         =   "开始"
      Height          =   300
      Left            =   9000
      TabIndex        =   89
      Top             =   5280
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox Text42 
      Height          =   270
      Left            =   12840
      TabIndex        =   88
      Text            =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      Caption         =   "启动星图"
      Height          =   375
      Left            =   4800
      TabIndex        =   87
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text41 
      Enabled         =   0   'False
      Height          =   270
      Left            =   7560
      TabIndex        =   85
      Text            =   "290"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer4 
      Left            =   11880
      Top             =   4800
   End
   Begin VB.TextBox Text40 
      Height          =   270
      Left            =   10800
      TabIndex        =   84
      Text            =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "停止"
      Height          =   255
      Left            =   10200
      TabIndex        =   83
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "AutoBattle"
      Height          =   255
      Left            =   11280
      TabIndex        =   82
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text38 
      Height          =   270
      Left            =   6120
      TabIndex        =   81
      Text            =   "1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text35 
      Height          =   270
      Left            =   4320
      TabIndex        =   79
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text34 
      Height          =   270
      Left            =   2520
      TabIndex        =   77
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text32 
      Height          =   270
      Left            =   960
      TabIndex        =   75
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "洗地模式"
      Height          =   375
      Left            =   7800
      TabIndex        =   72
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text31 
      Height          =   270
      Left            =   5040
      TabIndex        =   71
      Text            =   "210="
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text13 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   11640
      TabIndex        =   68
      Text            =   "1"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text39 
      Height          =   270
      Left            =   12240
      TabIndex        =   66
      Text            =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   360
      Top             =   4560
   End
   Begin VB.TextBox Text37 
      Enabled         =   0   'False
      Height          =   270
      Left            =   13200
      TabIndex        =   65
      Text            =   "0"
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox Text36 
      Height          =   270
      Left            =   6840
      TabIndex        =   63
      Text            =   "5000"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text33 
      Height          =   270
      Left            =   960
      TabIndex        =   61
      Text            =   """"
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Send"
      Height          =   375
      Left            =   11160
      TabIndex        =   60
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "T1"
      Height          =   375
      Left            =   10440
      TabIndex        =   59
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text30 
      Enabled         =   0   'False
      Height          =   270
      Left            =   5160
      TabIndex        =   54
      Text            =   "0"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text29 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   3000
      TabIndex        =   52
      Text            =   "5000"
      Top             =   4680
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   12240
      Top             =   3720
   End
   Begin VB.TextBox Text28 
      Enabled         =   0   'False
      Height          =   270
      Left            =   12120
      TabIndex        =   49
      Text            =   "未登录"
      Top             =   4320
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动重连"
      Height          =   255
      Left            =   11280
      TabIndex        =   48
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text27 
      Height          =   270
      Left            =   1440
      TabIndex        =   47
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text26 
      Height          =   270
      Left            =   8040
      TabIndex        =   46
      Text            =   "Text26"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text25 
      Height          =   270
      Left            =   7320
      TabIndex        =   45
      Text            =   "Text25"
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Left            =   12480
      Top             =   4320
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   255
      Left            =   5880
      TabIndex        =   44
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "get"
      Height          =   375
      Left            =   5160
      TabIndex        =   43
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text24 
      Height          =   270
      Left            =   3600
      TabIndex        =   42
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text23 
      Height          =   270
      Left            =   3240
      TabIndex        =   41
      Text            =   """"
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text22 
      Height          =   270
      Left            =   2640
      TabIndex        =   40
      Text            =   """ value"
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   255
      Left            =   3240
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   300
      Left            =   1800
      TabIndex        =   38
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   1440
      TabIndex        =   37
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text21 
      Height          =   270
      Left            =   2160
      TabIndex        =   36
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text20 
      Height          =   270
      Left            =   1800
      TabIndex        =   35
      Text            =   "险</"
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text19 
      Height          =   270
      Left            =   960
      TabIndex        =   34
      Text            =   "float:right;"">"
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      Height          =   270
      Left            =   9120
      TabIndex        =   33
      Text            =   "290"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text17 
      Enabled         =   0   'False
      Height          =   270
      Left            =   9120
      TabIndex        =   32
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text16 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   3840
      TabIndex        =   31
      Text            =   "1"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   10080
      TabIndex        =   30
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   375
      Left            =   10080
      TabIndex        =   29
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2880
      TabIndex        =   27
      Text            =   "1"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "Auto"
      Height          =   375
      Left            =   6000
      TabIndex        =   26
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   4080
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8655
      Left            =   13680
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   15266
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   1560
      TabIndex        =   24
      Text            =   "gzip, deflate"
      Top             =   3720
      Width           =   12015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送"
      Height          =   375
      Left            =   5880
      TabIndex        =   21
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   5040
      Width           =   13455
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   960
      TabIndex        =   19
      Top             =   3360
      Width           =   12615
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   720
      TabIndex        =   17
      Top             =   3000
      Width           =   12855
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1320
      TabIndex        =   15
      Text            =   "no-cache"
      Top             =   2640
      Width           =   12255
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1080
      TabIndex        =   13
      Text            =   "Keep-Alive"
      Top             =   2280
      Width           =   12495
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   1920
      Width           =   12255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   720
      TabIndex        =   9
      Text            =   $"Form1.frx":0000
      Top             =   1560
      Width           =   12855
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1200
      TabIndex        =   7
      Text            =   "application/x-www-form-urlencoded"
      Top             =   1200
      Width           =   12375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1440
      TabIndex        =   5
      Text            =   "zh-cn"
      Top             =   840
      Width           =   12135
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   12735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Text            =   "http://ol.lstyxl.com/index.php?page=login"
      Top             =   120
      Width           =   13215
   End
   Begin VB.Label Label29 
      Caption         =   "已出发舰队："
      Height          =   255
      Left            =   7560
      TabIndex        =   86
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label26 
      Caption         =   "攻击航道数："
      Height          =   255
      Left            =   6120
      TabIndex        =   80
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "球位："
      Height          =   255
      Left            =   4320
      TabIndex        =   78
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "星系："
      Height          =   255
      Left            =   2520
      TabIndex        =   76
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "银河系："
      Height          =   375
      Left            =   960
      TabIndex        =   74
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "目标星球："
      Height          =   375
      Left            =   120
      TabIndex        =   73
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "舰队数量："
      Height          =   255
      Left            =   10800
      TabIndex        =   70
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "舰船选择："
      Height          =   255
      Left            =   8040
      TabIndex        =   67
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label28 
      Caption         =   "顶号次数："
      Height          =   615
      Left            =   13200
      TabIndex        =   64
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label27 
      Caption         =   "自动重连时间："
      Height          =   375
      Left            =   5640
      TabIndex        =   62
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label22 
      Caption         =   "用户名："
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "密码："
      Height          =   255
      Left            =   1440
      TabIndex        =   57
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   "星系："
      Height          =   255
      Left            =   3840
      TabIndex        =   56
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "银河系："
      Height          =   255
      Left            =   2880
      TabIndex        =   55
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "网络波动次数："
      Height          =   255
      Left            =   3960
      TabIndex        =   53
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "错误延时时间（毫秒）："
      Height          =   255
      Left            =   1080
      TabIndex        =   51
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "登录状态："
      Height          =   255
      Left            =   11280
      TabIndex        =   50
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Accept-Encoding:"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "返回文本："
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Post Data:"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Cookie:"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Cache-Control:"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Connection:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Content-Length:"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Accept:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Content-Type:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Accept-Language:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Referer:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Const xStr As String = "0123456789"
Public cs As Integer
Public szText As String
Public szFindStrBegin As String
Public szFindStrEnd As String
Public nBegin As Long
Public nEnd As Long
Public nLength As Long
Public szMyText As String
Public ssText As String
Public ssFindStrBegin As String
Public ssFindStrEnd As String
Public sBegin As Long
Public sEnd As Long
Public sLength As Long
Public ssMyText As String
Public Sub TimeDelay(ByVal PauseSecond As Single)
' Attribute TimeDelay.VB_Description = "延时"
Dim Star, PauseTime
Star = Timer
PauseTime = PauseSecond
Do While Timer < Star + PauseTime
DoEvents
Loop
End Sub
Private Sub Combo1_Click()
If Combo1.Text = "小型运输舰" Then
Text31.Text = "202="
End If
If Combo1.Text = "大型运输舰" Then
Text31.Text = "203="
End If
If Combo1.Text = "狂热者级战机" Then
Text31.Text = "204="
End If
If Combo1.Text = "争斗者级战机" Then
Text31.Text = "205="
End If
If Combo1.Text = "远征级护卫舰" Then
Text31.Text = "206="
End If
If Combo1.Text = "边疆级壁垒舰" Then
Text31.Text = "207="
End If
If Combo1.Text = "殖民船" Then
Text31.Text = "208="
End If
If Combo1.Text = "摆渡人级回收舰" Then
Text31.Text = "209="
End If
If Combo1.Text = "探针" Then
Text31.Text = "210="
End If
If Combo1.Text = "重型隐匿轰炸舰" Then
Text31.Text = "211="
End If
If Combo1.Text = "永恒级战略堡垒舰" Then
Text31.Text = "213="
End If
If Combo1.Text = "星球要塞" Then
Text31.Text = "214="
End If
If Combo1.Text = "灾变级战舰" Then
Text31.Text = "215="
End If
If Combo1.Text = "君临者主宰舰" Then
Text31.Text = "216="
End If
If Combo1.Text = "信仰级运输舰" Then
Text31.Text = "217="
End If
If Combo1.Text = "幽能死星" Then
Text31.Text = "218="
End If
If Combo1.Text = "阿瓦隆级拖曳舰" Then
Text31.Text = "219="
End If
If Combo1.Text = "科考舰" Then
Text31.Text = "220="
End If
End Sub
Private Sub Command1_Click()
Dim XMLObject As XMLHTTP, SendStr As String '声明
Set XMLObject = CreateObject("Microsoft.XMLHTTP") '设置对象
If Text9.Text = "" Then '检查是否存在cookie(不存在cookie)
SendStr = Text10.Text
murl = Text1.Text
mReferer = Text2.Text
mAcceptLanguage = Text3.Text
mContentType = Text4.Text
mAccept = Text5.Text
mConnection = Text7.Text
mCacheControl = Text8.Text
mAcceptEncoding = Text12.Text '输入文本转换变量
XMLObject.open "POST", murl, False '目标网址
XMLObject.setRequestHeader "Referer", mReferer '来源的页面链接
XMLObject.setRequestHeader "CONTENT-TYPE", mContentType '代表发送端（客户端|服务器）发送的实体数据的数据类型[定义网络文件的类型和网页的编码，决定文件接收方将以什么形式、什么编码读取这个文件]
XMLObject.setRequestHeader "CONTENT-LENGTH", Len(SendStr) '内容的真实字节数
XMLObject.setRequestHeader "Accept-Language", mAcceptLanguage '客户端接收的语言类型
XMLObject.setRequestHeader "Accept", mAccept '发送端（客户端）希望接受的数据类型
XMLObject.setRequestHeader "Connection", mConnection '客户端和服务端的连接关系
XMLObject.setRequestHeader "Cache-Control", mCacheControl '服务端是否禁止客户端缓存页面数据
XMLObject.setRequestHeader "Accept-Encoding", mAcceptEncoding '(客户端能接收的压缩数据的类型)
XMLObject.sEnd SendStr '发送数据
Text11.Text = XMLObject.responseText '显示返回结果
Set XMLObject = Nothing '清除对象
Else '(存在cookie)
SendStr = Text10.Text
murl = Text1.Text
mReferer = Text2.Text
mAcceptLanguage = Text3.Text
mContentType = Text4.Text
mAccept = Text5.Text
mConnection = Text7.Text
mCacheControl = Text8.Text
mCookie = Text9.Text
mAcceptEncoding = Text12.Text '输入文本转换变量
XMLObject.open "POST", "murl", False '目标网址
XMLObject.setRequestHeader "Referer", mReferer '来源的页面链接
XMLObject.setRequestHeader "CONTENT-TYPE", mContentType '代表发送端（客户端|服务器）发送的实体数据的数据类型[定义网络文件的类型和网页的编码，决定文件接收方将以什么形式、什么编码读取这个文件]
XMLObject.setRequestHeader "CONTENT-LENGTH", Len(SendStr) '内容的真实字节数
XMLObject.setRequestHeader "Accept-Language", mAcceptLanguage '客户端接收的语言类型
XMLObject.setRequestHeader "Accept", mAccept '发送端（客户端）希望接受的数据类型
XMLObject.setRequestHeader "Connection", mConnection '客户端和服务端的连接关系
XMLObject.setRequestHeader "Cache-Control", mCacheControl '服务端是否禁止客户端缓存页面数据
XMLObject.setRequestHeader "Cookie", mCookie '增加的cookie内容[客户端暂存服务端的信息]
XMLObject.setRequestHeader "Accept-Encoding", mAcceptEncoding '(客户端能接收的压缩数据的类型)
XMLObject.sEnd SendStr '发送数据
Text11.Text = XMLObject.responseText '显示返回结果
Set XMLObject = Nothing '清除对象
End If
End Sub
Private Sub Command11_Click()
Text40.Text = 0
Text11.Visible = False
Label18.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
Label29.Visible = True
Text32.Visible = True
Text34.Visible = True
Text35.Visible = True
Text38.Visible = True
Text41.Visible = True
Command16.Visible = True
Command14.Visible = True
End Sub
Private Sub Command12_Click()
Call Command9_Click
Call Command8_Click
End Sub
Private Sub Command13_Click()
If Text40.Text = 1 Then
Exit Sub
End If
If Text40.Text = 0 Then
Do
If Text40.Text = 1 Then
Exit Sub
End If
Text1.Text = "http://ol.lstyxl.com/game.php?page=fleetStep1"
Text2.Text = "http://ol.lstyxl.com/game.php?page=fleetTable"
Text10.Text = "galaxy=1&system=2&planet=3&type=1&target_mission=0&ship" + Text31.Text + Text13.Text
Call Command1_Click
TimeDelay (1)
Text1.Text = "http://ol.lstyxl.com/game.php?page=fleetStep2"
Text2.Text = "http://ol.lstyxl.com/game.php?page=fleetStep1"
szFindStrBegin = "token" '定义要查找的字符串开头
szFindStrEnd = "fleet_group" '定义要查找的字符串结尾
szText = Text11.Text '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
nBegin = InStr(szText, szFindStrBegin) '找开头字符串
If nBegin > 0 Then '必须有能找到开头了才继续
nEnd = InStr(nBegin, szText, szFindStrEnd) '找结尾字符串
If nEnd > nBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
nLength = nEnd - nBegin + Len(szFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
nLength = nEnd - nBegin - Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
nBegin = nBegin + Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
szMyText = Mid(szText, nBegin, nLength) '取出“before then.”到 "test" 中间的东西
szText = Replace(szMyText, Text22.Text, "")
End If
End If
szFindStrBegin = "=" '定义要查找的字符串开头
szFindStrEnd = ">" '定义要查找的字符串结尾
nBegin = InStr(szText, szFindStrBegin) '找开头字符串
If nBegin > 0 Then '必须有能找到开头了才继续
nEnd = InStr(nBegin, szText, szFindStrEnd) '找结尾字符串
If nEnd > nBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
nLength = nEnd - nBegin + Len(szFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
nLength = nEnd - nBegin - Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
nBegin = nBegin + Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
szMyText = Mid(szText, nBegin, nLength) '取出“before then.”到 "test" 中间的东西
szMyText = Replace(szMyText, Text23.Text, "")
Text24.Text = szMyText
Text10.Text = "token=" + szMyText + "&fleet_group=0&target_mission=0&galaxy=" + Text32.Text + "&system=" + Text34.Text + "&planet=" + Text35.Text + "&type=1&speed=10"
Call Command1_Click
TimeDelay (1)
End If
End If
Text1.Text = "http://ol.lstyxl.com/game.php?page=fleetStep3"
Text2.Text = "http://ol.lstyxl.com/game.php?page=fleetStep2"
Text10.Text = "token=" + Text24.Text + "&mission=1&metal=&crystal=&deuterium=&staytime=1"
Call Command1_Click
If UBound(Split(Text11.Text, "该行星的舰船数量小于输入的舰船数量")) > 0 Then
End
End If
TimeDelay (1)
Call Command9_Click
Call Command8_Click
Text41.Text = UBound(Split(Text11.Text, Text32.Text + ":" + Text34.Text + ":" + Text35.Text))
If UBound(Split(Text11.Text, "首页")) > 0 And Check1.Value = 1 Then
Text1.Text = "http://ol.lstyxl.com/index.php?page=login"
Text2.Text = ""
Text10.Text = "uni=1&username=" + Text15.Text + "&password=" + Text27.Text
Call Command1_Click
Call Command13_Click
End If
If UBound(Split(Text11.Text, "首页")) > 0 And Check1.Value = 0 Then
End
End If
Loop While UBound(Split(Text11.Text, Text32.Text + ":" + Text34.Text + ":" + Text35.Text)) < Text38.Text
End If
End Sub
Private Sub Command14_Click()
Text40.Text = 1
Text11.Visible = True
Label18.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label29.Visible = False
Text32.Visible = False
Text34.Visible = False
Text35.Visible = False
Text38.Visible = False
Text41.Visible = False
Command16.Visible = False
Command14.Visible = False
End Sub
Private Sub Command15_Click()
Shell (App.Path & "\星图.exe")
End Sub

Private Sub Command16_Click()
Text40.Text = 0
Call Command13_Click
End Sub
Private Sub Command2_Click()
Text42.Text = 1
End Sub
Private Sub Command3_Click()
Text42.Text = 0
Call Command6_Click
End Sub
Private Sub Command6_Click()
If Text42.Text = 1 Then
Exit Sub
End If
If Text42.Text = 0 Then
Do
If Text42.Text = 1 Then
Exit Sub
End If
Text1.Text = "http://ol.lstyxl.com/game.php?page=fleetStep1"
Text2.Text = "http://ol.lstyxl.com/game.php?page=fleetTable"
Text10.Text = "galaxy=1&system=2&planet=3&type=1&target_mission=0&ship" + Text31.Text + Text13.Text
Call Command1_Click
TimeDelay (1)
Text1.Text = "http://ol.lstyxl.com/game.php?page=fleetStep2"
Text2.Text = "http://ol.lstyxl.com/game.php?page=fleetStep1"
szFindStrBegin = "token" '定义要查找的字符串开头
szFindStrEnd = "fleet_group" '定义要查找的字符串结尾
szText = Text11.Text '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
nBegin = InStr(szText, szFindStrBegin) '找开头字符串
If nBegin > 0 Then '必须有能找到开头了才继续
nEnd = InStr(nBegin, szText, szFindStrEnd) '找结尾字符串
If nEnd > nBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
nLength = nEnd - nBegin + Len(szFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
nLength = nEnd - nBegin - Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
nBegin = nBegin + Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
szMyText = Mid(szText, nBegin, nLength) '取出“before then.”到 "test" 中间的东西
szText = Replace(szMyText, Text22.Text, "")
End If
End If
szFindStrBegin = "=" '定义要查找的字符串开头
szFindStrEnd = ">" '定义要查找的字符串结尾
nBegin = InStr(szText, szFindStrBegin) '找开头字符串
If nBegin > 0 Then '必须有能找到开头了才继续
nEnd = InStr(nBegin, szText, szFindStrEnd) '找结尾字符串
If nEnd > nBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
nLength = nEnd - nBegin + Len(szFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
nLength = nEnd - nBegin - Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
nBegin = nBegin + Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
szMyText = Mid(szText, nBegin, nLength) '取出“before then.”到 "test" 中间的东西
szMyText = Replace(szMyText, Text23.Text, "")
Text24.Text = szMyText
Text10.Text = "token=" + szMyText + "&fleet_group=0&target_mission=0&galaxy=" + Text14.Text + "&system=" + Text16.Text + "&planet=16&type=1&speed=10"
Call Command1_Click
TimeDelay (1)
End If
End If
Text1.Text = "http://ol.lstyxl.com/game.php?page=fleetStep3"
Text2.Text = "http://ol.lstyxl.com/game.php?page=fleetStep2"
Text10.Text = "token=" + Text24.Text + "&mission=15&metal=&crystal=&deuterium=&staytime=1"
Call Command1_Click
If UBound(Split(Text11.Text, "该行星的舰船数量小于输入的舰船数量")) > 0 Then
End
End If
TimeDelay (1)
Call Command9_Click
Call Command8_Click
If UBound(Split(Text11.Text, "武器")) > 0 Then
Text17.Text = UBound(Split(Text11.Text, "探险")) - 1
ssFindStrBegin = Text19.Text '定义要查找的字符串开头
ssFindStrEnd = Text20.Text '定义要查找的字符串结尾
ssText = Text11.Text '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
sBegin = InStr(ssText, ssFindStrBegin) '找开头字符串
If sBegin > 0 Then '必须有能找到开头了才继续
sEnd = InStr(sBegin, ssText, ssFindStrEnd) '找结尾字符串
If sEnd > sBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
sLength = sEnd - sBegin + Len(ssFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
sLength = sEnd - sBegin - Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
sBegin = sBegin + Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
ssMyText = Mid(ssText, sBegin, sLength) '取出“before then.”到 "test" 中间的东西
Text21.Text = ssMyText
End If
End If
ssFindStrBegin = "/ " '定义要查找的字符串开头
ssFindStrEnd = " 探" '定义要查找的字符串结尾
ssText = Text21.Text '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
sBegin = InStr(ssText, ssFindStrBegin) '找开头字符串
If sBegin > 0 Then '必须有能找到开头了才继续
sEnd = InStr(sBegin, ssText, ssFindStrEnd) '找结尾字符串
If sEnd > sBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
sLength = sEnd - sBegin + Len(ssFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
sLength = sEnd - sBegin - Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
sBegin = sBegin + Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
ssMyText = Mid(ssText, sBegin, sLength) '取出“before then.”到 "test" 中间的东西
Text18.Text = ssMyText
End If
End If
End If
If UBound(Split(Text11.Text, "首页")) > 0 And Check1.Value = 1 Then
Text1.Text = "http://ol.lstyxl.com/index.php?page=login"
Text2.Text = ""
Text10.Text = "uni=1&username=" + Text15.Text + "&password=" + Text27.Text
Call Command1_Click
Call Command3_Click
End If
If UBound(Split(Text11.Text, "首页")) > 0 And Check1.Value = 0 Then
End
End If
Loop While Val(Text17.Text) <> Val(Text18.Text)
End If
End Sub
Private Sub Command8_Click()
On Error GoTo errorcheck
Dim XMLObject As XMLHTTP, SendStr As String '声明
Set XMLObject = CreateObject("Microsoft.XMLHTTP") '设置对象
SendStr = Text10.Text
murl = Text1.Text
mReferer = Text2.Text
mAcceptLanguage = Text3.Text
mContentType = Text4.Text
mAccept = Text5.Text
mConnection = Text7.Text
mCacheControl = Text8.Text
mAcceptEncoding = Text12.Text '输入文本转换变量
XMLObject.open "GET", murl, False '目标网址
XMLObject.setRequestHeader "Referer", mReferer '来源的页面链接
XMLObject.setRequestHeader "CONTENT-TYPE", mContentType '代表发送端（客户端|服务器）发送的实体数据的数据类型[定义网络文件的类型和网页的编码，决定文件接收方将以什么形式、什么编码读取这个文件]
XMLObject.setRequestHeader "CONTENT-LENGTH", Len(SendStr) '内容的真实字节数
XMLObject.setRequestHeader "Accept-Language", mAcceptLanguage '客户端接收的语言类型
XMLObject.setRequestHeader "Accept", mAccept '发送端（客户端）希望接受的数据类型
XMLObject.setRequestHeader "Connection", mConnection '客户端和服务端的连接关系
XMLObject.setRequestHeader "Cache-Control", mCacheControl '服务端是否禁止客户端缓存页面数据
XMLObject.setRequestHeader "Accept-Encoding", mAcceptEncoding '(客户端能接收的压缩数据的类型)
XMLObject.sEnd SendStr '发送数据
Text11.Text = XMLObject.responseText '显示返回结果
Set XMLObject = Nothing '清除对象
Exit Sub
errorcheck:
Timer2.Interval = 0
Timer1.Interval = 0
Dim Savetime As Double
Savetime = timeGetTime '记下开始时的时间
While timeGetTime < Savetime + Text29.Text '循环等待
DoEvents '转让控制权，以便让操作系统处理其它的事件。
Wend
Text30.Text = Text30.Text + 1
Timer2.Interval = 1800
Timer1.Interval = 1800
End Sub
Private Sub Command9_Click()
Text1.Text = "http://ol.lstyxl.com/game.php?page=fleetTable"
Text2.Text = "http://ol.lstyxl.com/game.php"
Text10.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Text10_Change()
Text6.Text = Len(Text10.Text) '内容转换为真实字节数
End Sub
Private Sub Text13_LostFocus()
If Val(Text13.Text) <= 0 Then
Text13.Text = "1"
Text1.SetFocus
End If
End Sub
Private Sub Text15_Change()
Text10.Text = "uni=1&username=" + Text15.Text + "&password=" + Text27.Text
End Sub
Private Sub Text17_Change()
If Val(Text17.Text) <> Val(Text18.Text) Then
Timer2.Interval = 0
Timer1.Interval = 0
TimeDelay (1)
Call Command6_Click
End If
If Val(Text17.Text) = Val(Text18.Text) Then
Timer2.Interval = 1800
Timer1.Interval = 1800
End If
End Sub
Private Sub Text27_Change()
Text10.Text = "uni=1&username=" + Text15.Text + "&password=" + Text27.Text
End Sub
Private Sub Text41_Change()
If Val(Text41.Text) <> Val(Text38.Text) Then
Timer2.Interval = 0
Timer4.Interval = 0
TimeDelay (1)
Call Command13_Click
End If
If Val(Text41.Text) = Val(Text38.Text) Then
Timer2.Interval = 1800
Timer4.Interval = 1800
End If
End Sub
Private Sub Timer1_Timer()
Text25.Text = Val(Text17.Text)
Text26.Text = Val(Text18.Text)
If UBound(Split(Text11.Text, "武器")) > 0 Then
Text17.Text = UBound(Split(Text11.Text, "探险")) - 1
ssFindStrBegin = Text19.Text '定义要查找的字符串开头
ssFindStrEnd = Text20.Text '定义要查找的字符串结尾
ssText = Text11.Text '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
sBegin = InStr(ssText, ssFindStrBegin) '找开头字符串
If sBegin > 0 Then '必须有能找到开头了才继续
sEnd = InStr(sBegin, ssText, ssFindStrEnd) '找结尾字符串
If sEnd > sBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
sLength = sEnd - sBegin + Len(ssFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
sLength = sEnd - sBegin - Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
sBegin = sBegin + Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
ssMyText = Mid(ssText, sBegin, sLength) '取出“before then.”到 "test" 中间的东西
Text21.Text = ssMyText
End If
End If
ssFindStrBegin = "/ " '定义要查找的字符串开头
ssFindStrEnd = " 探" '定义要查找的字符串结尾
ssText = Text21.Text '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
sBegin = InStr(ssText, ssFindStrBegin) '找开头字符串
If sBegin > 0 Then '必须有能找到开头了才继续
sEnd = InStr(sBegin, ssText, ssFindStrEnd) '找结尾字符串
If sEnd > sBegin Then '结尾必须比开头的位置大
'包含查找的字符串模式，注释掉下面的2行
sLength = sEnd - sBegin + Len(ssFindStrEnd) '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
'不包含查找的字符串模式
sLength = sEnd - sBegin - Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
sBegin = sBegin + Len(ssFindStrBegin) '如果不包含查找的字符串，用这2行
ssMyText = Mid(ssText, sBegin, sLength) '取出“before then.”到 "test" 中间的东西
Text18.Text = ssMyText
End If
End If
End If
End Sub
Private Sub Timer2_Timer()
Call Command9_Click
Call Command8_Click
End Sub
Private Sub Timer3_Timer()
If UBound(Split(Text11.Text, "首页")) > 0 Then
Text28.Text = "未登录"
End If
If UBound(Split(Text11.Text, "星球概况")) > 0 Then
Text28.Text = "已登录"
End If
End Sub
Private Sub Timer4_Timer()
Text41.Text = UBound(Split(Text11.Text, Text32.Text + ":" + Text34.Text + ":" + Text35.Text))
End Sub
Private Sub Timer7_Timer()
If UBound(Split(Text10.Text, "no-js")) > 0 Then
Call Command3_Click
End If
End Sub
Private Sub WebBrowser1_DownloadBegin()
    WebBrowser1.Silent = True
End Sub
Private Sub WebBrowser1_DownloadComplete()
    WebBrowser1.Silent = True
End Sub
Private Sub Form_Load()
If App.PrevInstance Then
MsgBox "本程序已打开，请关闭后在打开"
End
End If
Combo1.AddItem "小型运输舰"
Combo1.AddItem "大型运输舰"
Combo1.AddItem "狂热者级战机"
Combo1.AddItem "争斗者级战机"
Combo1.AddItem "远征级护卫舰"
Combo1.AddItem "边疆级壁垒舰"
Combo1.AddItem "殖民船"
Combo1.AddItem "摆渡人级回收舰"
Combo1.AddItem "探针"
Combo1.AddItem "重型隐匿轰炸舰"
Combo1.AddItem "永恒级战略堡垒舰"
Combo1.AddItem "星球要塞"
Combo1.AddItem "灾变级战舰"
Combo1.AddItem "君临者主宰舰"
Combo1.AddItem "信仰级运输舰"
Combo1.AddItem "幽能死星"
Combo1.AddItem "阿瓦隆级拖曳舰"
Combo1.AddItem "科考舰"
Combo1.Text = Combo1.List(8)
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
'只能输入数字
KeyAscii = IIf(InStr(xStr & Chr(8), Chr(KeyAscii)), KeyAscii, 0)
End Sub
Private Sub Text29_KeyPress(KeyAscii As Integer)
'只能输入数字
KeyAscii = IIf(InStr(xStr & Chr(8), Chr(KeyAscii)), KeyAscii, 0)
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
'只能输入数字
KeyAscii = IIf(InStr(xStr & Chr(8), Chr(KeyAscii)), KeyAscii, 0)
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
'只能输入数字
KeyAscii = IIf(InStr(xStr & Chr(8), Chr(KeyAscii)), KeyAscii, 0)
End Sub
Private Sub Text38_KeyPress(KeyAscii As Integer)
'只能输入数字
KeyAscii = IIf(InStr(xStr & Chr(8), Chr(KeyAscii)), KeyAscii, 0)
End Sub
