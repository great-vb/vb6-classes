VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'大整数结构体
Private Type LARGE_INTEGER

    LowPart As Long
    HighPart As Long

End Type

'获取时间计数器计数值
Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long

'获取震动频率：一个大整数
Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

'用于将大整数的字节复制到Currency类型，完成转换的工作
Private Declare Sub CopyMemory _
                Lib "kernel32.dll" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

'初始时间
Private initTime  As LARGE_INTEGER

'每秒震动次数
Private freqCount As LARGE_INTEGER

'采样时振动器震动计数值
Private logTime   As LARGE_INTEGER

'每秒震动次数的Currency类型数据
Private fFreq     As Currency

'将一个长整型转换成高精度型数据
Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function

'类初始化
Private Sub Class_Initialize()

    '获取振荡器的震荡频率
    If QueryPerformanceFrequency(freqCount) > 0 Then
        If QueryPerformanceCounter(initTime) > 0 Then
        Else
            Err.Raise 11, , "记录震荡次数出错！"
        End If

    Else
        Err.Raise 11, , "读取震荡频率出错！"
    End If

    fFreq = LargeIntToCurrency(freqCount)
End Sub

'重置计数器
Public Sub ReInit()
    Call Class_Initialize
End Sub

'重置计数器，并返回从上次重置计数器开始共流失的秒数。
Public Function ReInitGet() As Currency

    Dim cuStart As Currency

    Dim cuEnd   As Currency

    QueryPerformanceCounter logTime
    cuStart = LargeIntToCurrency(initTime)
    cuEnd = LargeIntToCurrency(logTime)
    ReInitGet = (cuEnd - cuStart) / fFreq
    QueryPerformanceCounter initTime
End Function

'返回从上次重置计数器开始共流失的秒数，但不重置计数器
Public Function SmoothGet() As Currency

    Dim cuStart As Currency

    Dim cuEnd   As Currency

    QueryPerformanceCounter logTime
    cuStart = LargeIntToCurrency(initTime)
    cuEnd = LargeIntToCurrency(logTime)
    SmoothGet = (cuEnd - cuStart) / fFreq
End Function

