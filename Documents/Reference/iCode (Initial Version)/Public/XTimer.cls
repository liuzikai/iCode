VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'===============================================
' 注意！当调试这个工程时不要按“结束”按钮！
'   请参阅在 XTimerSupport 模块(XTimerS.bas)
'   顶部的解释。
'===============================================

' XTimer 属性的私有存储体：
Private mlngTimerID As Long
Private mlngInterval As Long
Private mblnEnabled As Boolean

' XTimer的唯一事件是 Tick。XTimer 的 Tick 事件
'   没有任何参数(排除参数可以加速
'   事件的苗条性)，但它不能成为您希望
'   提供但没有提供参数的原因。
Event Tick()

' TimerID property 被 EndTimer 过程所要求，
' ----------------   其目的是在活动的计时器
'   的支持模块数组中快速定位计时器。
'
' 对于客户来说，使用这个属性没有意义，
'   索引它被声明为 Friend 而非 Public。
'
Friend Property Get TimerID() As Long
    TimerID = mlngTimerID
End Property

' Enabled 属性打开和关闭计时器。它通过
' -----------  去除系统的计时器来实现，
'   因为没有办法暂停系统计时器。
'   如果他们存在，则正在运行。
'
Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
'
Public Property Let Enabled(ByVal NewValue As Boolean)
    ' 如果对属性的状态没有进行更改，
    '   则退出。这个避免了当已经
    '   有一个系统计时器正在运行时，
    '   启动第二个系统计时器，等等。
    If NewValue = mblnEnabled Then Exit Property
    '
    ' 保存新的属性设置。
    mblnEnabled = NewValue
    '
    ' 如果间隔为零，则计时器已经
    '   停止。不要启动它。
    If mlngInterval = 0 Then Exit Property
    '
    ' 打开和关闭计时器。
    If mblnEnabled Then
        Debug.Assert mlngTimerID = 0
        mlngTimerID = BeginTimer(Me, mlngInterval)
    Else
        ' 下列的代码是必须的，因为
        '   XTimer 可以通过两个方法
        '   关闭系统计时器：Enabled = False，
        '   或者 Interval = 0。
        If mlngTimerID <> 0 Then
            Call EndTimer(Me)
            mlngTimerID = 0
        End If
    End If
End Property

' Interval 属性需要做很多事情除了设置
' -----------------   计时器的间隔。如果 XTimer
'   是 enabled，并且间隔被从零到非零进行了更改
'   那么系统计时器必须被启动。
'   同样地，如果间隔被更改为零，
'   系统计时器必须被停止。
'
' 无论间隔是否更改此属性让过程
'   都停止一个系统计时器并且启动另一个系统计时器。
'   这是因为无法更改一个系统计时器的间隔。
'
Public Property Get Interval() As Long
    Interval = mlngInterval
End Property
'
Public Property Let Interval(ByVal NewInterval As Long)
    ' 如果新的间隔值与旧的一致，
    '   没有必要做任何工作。
    If NewInterval = mlngInterval Then Exit Property
    '
    ' 保存新的值。
    mlngInterval = NewInterval
    '
    ' 如果 XTimer 是活动的，mlngTimerID 为非零。
    '   在这种情况下，旧的系统计时器必须在新的
    '   启动之前终止。
    If mlngTimerID <> 0 Then
        Call EndTimer(Me)
        mlngTimerID = 0
    End If
    '
    ' 如果新的间隔为零，那么 XTimer
    '   成为非活动的，不管当前的
    '   Enabled 属性为何值。如果新的间隔
    '   不为零，并且 Enabled 属性为 True，
    '   那么一个新的系统计时器被启动，并且
    '   它的 ID 被存储在 mlngTimerID 中。
    If (NewInterval <> 0) And mblnEnabled Then
        mlngTimerID = BeginTimer(Me, NewInterval)
    End If
End Property

' RaiseTick 被所支持的模块所调用，当
' ----------------   系统计时器事件为
'   这个 XTimer 对象的系统计时器产生的时侯。
'
' 实现的细节：您可能期待声明这个
'   方法为 Friend 代替 Public，因为不
'   需要让客户来调用 RaiseTick。然而，
'   RaiseTick 被决定声明为 Public，
'   因为 XTimer 在 Tick 事件然被处理时
'   可能被释放。当对象的一个公有方法在堆栈中时，
'   它不能被终止，但是在堆栈中的它的
'   Friend 方法可以被终止。如果此对象
'   在 Friend 方法返回之前终止 (如果
'   客户在 XTimer 的 Tick 事件中执行大量
'   代码，这种情况可能会发生)，结果是产生一个 GPF 。
'   (注意这个高度的不寻常情景取决于外部的事件；
'   它在一般的使用 Friend 函数中不会发生。)
'
Public Sub RaiseTick()
    RaiseEvent Tick
End Sub

Private Sub Class_Terminate()
    ' 当客户释放它的最后一个对 XTimer 对象的引用时
    '   它将离开 -- 但仅仅当
    '   XTimer 的 Enabled 属性为 False 的时侯，
    '   或者它的 Interval 属性为 True!
    '
    ' 这是因为当 XTimer 的系统计时器
    '   正在运行时，XTimerSupport 模块
    '   为了产生它的 Tick 事件，
    '   不得不使用一个到 XTimer 的引用。
    '   那么客户在释放 XTimer 对象前
    '   禁用它们，将削弱系统计时器!
    '
    ' 这些削弱的系统计时器直到
    '   XTimers 部件关闭时才能被恢复
    '   -- 那就是说，当客户使用的
    '   DLL 关闭时，此 DLL 将不被
    '   卸载直到所有 XTimer 对象被释放。
    '   因为对公有对象的引用 (在这种
    '   情况下，它们被 XTimerSupport 所保存)
    '   将阻止 DLL 的卸载。
    '
    ' 那么为什么要麻烦地清除终止事件中的系统
    '   计时器呢？因为当
    '   DLL 被关闭后，所有到
    '   XTimer 对象的引用将被清除
    '   -- 并且 XTimer 将获得它的终止
    '   事件。系统计时器在此时将被破坏。
    On Error Resume Next
    If mlngTimerID <> 0 Then KillTimer 0, mlngTimerID
    '
    ' 下列工作是 XTimer 将要做的如果
    '   它非要释放关闭的 DLL 的优先级
    'If mlngTimerID <> 0 Then Call EndTimer(Me)
End Sub

Public Sub Start()
    Me.Enabled = True
End Sub

Public Sub Pause()
    Me.Enabled = False
End Sub
