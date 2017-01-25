VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccelerationMotionTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const Defult_TimerInterval As Long = 20
Private Const Defult_Time As Long = 300
Private Const Defult_VMinPerV0 As Double = 0.1

Event Timer(ByVal Value As Double)

Private m_StartValue As Double
Private m_EndValue As Double
Private m_Time As Long
Private m_VMinPerV0 As Double

Public Enum AMTimer_Direction
    Positive = 1
    None = 0
    Negative = -1
End Enum

Private m_Direction As AMTimer_Direction

Private m_TimerInterval As Long

Private WithEvents m_Timer As XTimer
Attribute m_Timer.VB_VarHelpID = -1

Private s As Double, V0 As Double, a As Double, t As Long, tMax As Long

Public Sub Strat()
    s = Me.EndValue - Me.StartValue
    tMax = Me.Time / Me.TimerInterval
    
    '     Vt + V0                  2s                 2s
    's = --------- * t => Vt+V0 = ---- => V0 = -----------------
    '        2                      t           (1+VMinPerV0)*t
    
    V0 = s * 2 / ((1 + Me.VMinPerV0) * tMax)
    a = ((Me.VMinPerV0 - 1) * V0) / tMax
    
    m_Direction = Positive
    
    t = 0
    
    m_Timer.Enabled = True
End Sub

Public Sub RollBack()
    s = Me.StartValue - Me.EndValue
    tMax = Me.Time / Me.TimerInterval
    
    '     Vt + V0                  2s                 2s
    's = --------- * t => Vt+V0 = ---- => V0 = -----------------
    '        2                      t           (1+VMinPerV0)*t
    
    V0 = s * 2 / ((1 + Me.VMinPerV0) * tMax)
    a = ((Me.VMinPerV0 - 1) * V0) / tMax
    
    m_Direction = Negative
    
    t = 0
    
    m_Timer.Enabled = True
End Sub

Public Sub Pause()
    m_Timer.Enabled = False
End Sub

Public Property Get TimerInterval() As Long
    TimerInterval = m_TimerInterval
End Property

Public Property Let TimerInterval(Value As Long)
    m_TimerInterval = Value
End Property

Public Property Get Direction() As AMTimer_Direction
    Direction = m_Direction
End Property

Public Property Let Direction(Value As AMTimer_Direction)
    If m_Direction = Value Then Exit Property
    If Value = Positive Then
        Me.Strat
    ElseIf Value = None Then
        Me.Pause
    ElseIf Value = Negative Then
        Me.RollBack
    End If
End Property

Public Property Get VMinPerV0() As Double
    VMinPerV0 = m_VMinPerV0
End Property

Public Property Let VMinPerV0(Value As Double)
    m_VMinPerV0 = Value
End Property

Public Property Get Time() As Long
    Time = m_Time
End Property

Public Property Let Time(Value As Long)
    m_Time = Value
End Property

Public Property Get StartValue() As Double
    StartValue = m_StartValue
End Property

Public Property Let StartValue(Value As Double)
    m_StartValue = Value
End Property

Public Property Get EndValue() As Double
    EndValue = m_EndValue
End Property

Public Property Let EndValue(Value As Double)
    m_EndValue = Value
End Property

Private Sub Class_Initialize()
    m_Time = Defult_Time
    m_TimerInterval = Defult_TimerInterval
    m_VMinPerV0 = Defult_VMinPerV0
    Set m_Timer = New XTimer
    m_Timer.Interval = m_TimerInterval
End Sub

Private Sub Class_Terminate()
    Set m_Timer = Nothing
End Sub

Private Sub m_Timer_Tick()
    t = t + 1
    If t >= tMax Then
        m_Timer.Enabled = False
        Dim EV As Double
        If Direction = Positive Then
            EV = Me.EndValue
        ElseIf Direction = Negative Then
            EV = Me.StartValue
        End If
        RaiseEvent Timer(EV)
    Else
        Dim SV As Double
        If Direction = Positive Then
            SV = Me.StartValue
        ElseIf Direction = Negative Then
            SV = Me.EndValue
        End If
        RaiseEvent Timer(SV + V0 * t + 0.5 * a * t * t)
    End If
End Sub
