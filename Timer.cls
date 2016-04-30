VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'開始時間と終了時間のインスタンスを作成します

'サンプルコード
'Dim stopwatch As New Timer: stopwatch.StartTimer
'Dim i As Long              '処理
'For i = 0 To 1000000000    '処理
'Next			    '処理
'MsgBox stopwatch.StopTimer.GetString


Dim StartTime As Variant
Dim StopTime As Variant


Public Function StartTimer() As Timer
    StartTime = Time
    Set StopTimer = Me
End Function

Public Function StopTimer() As Timer
    StopTime = Time
    Set StopTimer = Me
End Function

Public Function GetString() As Variant
    
    StopTime = StopTime - StartTime
    
    GetString = Minute(StopTime) & "分" & Second(StopTime) & "秒"
    
End Function

Public Function GetVariant() As Variant
    
    StopTime = StopTime - StartTime
    
    GetVariant = StopTime
    
End Function
