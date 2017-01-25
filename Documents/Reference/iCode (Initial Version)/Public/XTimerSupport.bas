Attribute VB_Name = "XTimerSupport"
Option Explicit
'================================================
' 注意！当调试此工程时不要按结束按钮！
'   在中断模式时，不要对工程进行编辑，
'   那样将导致工程被重置！
'
' 这个模块具有危险性因为它使用
'   SetTimer API 和 AddressOf 操作符
'   来设置一个代码计时器。一旦这样
'   计时器被设置，此系统在您返回到
'   设计时之后将继续调用 TimerProc 函数事件。
'
' 因为 TimerProc 在设计时不可用，
'   这个程序在 Visual Basic 中将
'   产生一个程序故障。
'
' 当调试这个模块时，您需要确定
'   在返回到设计时之前所有的系统
'   计时器都已经停止(使用 KillTimer)。
'   您可以从立即窗口中调用
'   SCRUB 来完成此操作。
'
' 回调计时器具有固有的危险性。
'   使用计时器控件对于您多数的开发进程
'   将更加的安全，只有到最后才切换到
'   回调计时器。
'==================================================

' 当使用更多的计时器时，数组 maxti
'   的合计尺寸将增大。(参阅下面的 'MoreRoom:' 代码。)
Const MAXTIMERINCREMEMT = 5

Private Type XTIMERINFO                                                         ' Hungarian xti
    xt As XTimer
    id As Long
    blnReentered As Boolean
End Type

Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerProc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

' maxti 是一个活动的 XTimer 对象的数组。使用用户
' -----   定义类型的数组来代替一个集合对象
'   的原因是当我们产生 XTimer 对象的 Tick
'   事件时获得早期绑定。
Private maxti() As XTIMERINFO
'
' mintMaxTimers 告诉我们在任何给定的时间时
' -------------   数组 maxti 由多大。
Private mintMaxTimers As Integer

' BeginTimer function 当 XTimer 的间隔属性被
' -------------------   设置成一个新的非零值时
'   被一个 XTimer 对象所调用。
'
' 此函数使 API 调用产生一个请求来设置
'   计时器。如果计时器被成功的创建，
'   此函数放置到 XTimer 对象的引用到
'   数组 maxti。这个引用将用于调用
'   产生 XTimer 的 Tick 事件的方法。
Public Function BeginTimer(ByVal xt As XTimer, ByVal Interval As Long)
    Dim lngTimerID As Long
    Dim intTimerNumber As Integer
    
    lngTimerID = SetTimer(0, 0, Interval, AddressOf TimerProc)
    ' 成功的条件时从 SetTimer 返回一个非零值。如果我们不能
    '   获得一个计时器，则产生一个错误。
    If lngTimerID = 0 Then Err.Raise vbObjectError + 31013, , "没有可用的计时器"
    
    ' 下面的循环在数组 maxti 中定位出第一个
    '   可用的空位。如果超过了上一级绑定，
    '   产生错误并且将数组加大。(如果
    '   您编译这个 DLL 为本地代码，不要关闭
    '   绑定检查！)
    For intTimerNumber = 1 To mintMaxTimers
        If maxti(intTimerNumber).id = 0 Then Exit For
    Next
    '
    ' 如果没有找到空位，增加数组的尺寸。
    If intTimerNumber > mintMaxTimers Then
        mintMaxTimers = mintMaxTimers + MAXTIMERINCREMEMT
        ReDim Preserve maxti(1 To mintMaxTimers)
    End If
    '
    ' 当产生 XTimer 对象的 Tick 事件时，
    '   保存一个到使用的引用。
    Set maxti(intTimerNumber).xt = xt
    '
    ' 保存 SetTimer API 返回的的计时器 id ，
    '   并且返回到 XTimer 对象的值。
    maxti(intTimerNumber).id = lngTimerID
    maxti(intTimerNumber).blnReentered = False
    BeginTimer = lngTimerID
End Function

' TimerProc 是计时器过程，在您的某个计时器关闭时
' ---------   系统将调用它。
'
' 重要信息 -- 因为这个过程必须在标准模块中，
'   您的所有计时器对象都必须可以共享它。
'   这意味着过程必须标识出哪一个计时器
'   已经关闭。这个操作通过查找数组
'   maxti 的计时器 ID 来完成(idEvent)。
'
' 如果这个子过程声明错误，将产生程序故障！
'   这是使用要求回调函数的 API
'   的危险处之一。
'
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
    Dim intCt As Integer
    
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = idEvent Then
            ' 如果这个事件的一个早期实例
            '   仍然在进行时，不要产生此事件。
            If maxti(intCt).blnReentered Then Exit Sub
            ' blnReentered 标志将阻塞这个
            '   事件的未来实例直到
            '   当前的事例完成。
            maxti(intCt).blnReentered = True
            On Error Resume Next
            ' 为适当的 XTimer 对象产生一个 Tick 事件。
            maxti(intCt).xt.RaiseTick
            If Err.Number <> 0 Then
                ' 如果发生错误，XTimer
                '   将负责终止除第一个运行的
                '   外的所有计时器。清除
                '   孤立计时器，来防止
                '   以后产生 GP 故障。
                KillTimer 0, idEvent
                maxti(intCt).id = 0
                '
                ' 释放到 XTimer 对象的引用。
                Set maxti(intCt).xt = Nothing
            End If
            '
            ' 允许这个事件再次进入 TimerProc。
            maxti(intCt).blnReentered = False
            Exit Sub
        End If
    Next
    ' 下面的代码行是一个失败的保护，
    '   在 XTimer 无故被释放而 Windows
    '   系统计时器还没有将它铲除。
    '
    ' 执行也可能到达这个点因为一个
    '   已知的 NT 3.51 的错误，因此您可能
    '   在执行了 KillTimer API 后收到
    '   一个外部的计时器事件。
    KillTimer 0, idEvent
End Sub

' EndTimer 过程被 XTimer 调用，每当
' ------------------   Enabled 属性被设置为
'   False，以及当需要一个新的计时器间隔。
'   没有办法来重新设置系统计时器，所以更改
'   间隔的唯一方法是铲除现存的计时器并且
'   调用 BeginTimer 来启动一个新的计时器。
'
Public Sub EndTimer(ByVal xt As XTimer)
    Dim lngTimerID As Long
    Dim intCt As Integer
    
    ' 询问 XTimer 的 TimerID，以至我们可以查找
    '   有关正确的 XTIMERINFO 的数组。(您可以
    '   查找 XTimer 自身的引用，使用
    '   Is 操作符来同带有 maxti(intCt).xt 的 xt 进行
    '   比较，但那样将降低速度。)
    lngTimerID = xt.TimerID
    '
    ' 如果计时器 ID 是零，EndTimer 调用将出错。
    If lngTimerID = 0 Then Exit Sub
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = lngTimerID Then
            ' 铲除系统计时器。
            KillTimer 0, lngTimerID
            '
            ' 释放对 XTimer 对象的引用。
            Set maxti(intCt).xt = Nothing
            '
            ' 清除 ID，释放空位以备新的活动的计时器使用。
            maxti(intCt).id = 0
            Exit Sub
        End If
    Next
End Sub

' Scrub 过程是一个仅仅为了调试目的的安全的阀：
' ---------------   如果当 XTimer 对象活动时，
'   您不得不结束这个工程，从立即窗口中调用 Scrub 。
'   这将调用  KillTimer 来铲除所有的系统计时器，
'   使开发环境可以安全地放回到设计模式。
'
Public Sub Scrub()
    Dim intCt As Integer
    ' 铲除仍然活动的计时器。
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id <> 0 Then KillTimer 0, maxti(intCt).id
    Next
End Sub
