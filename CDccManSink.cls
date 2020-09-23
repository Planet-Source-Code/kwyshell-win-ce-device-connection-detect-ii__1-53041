VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDccManSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private m_objDccMan As CDccMan
Private m_lContext As Long

Implements IDccManSink

Private Sub Class_Initialize()

    ' Kwyshell. 2004.04.09 Install Com Wrap for Sink Interface
    MDccManSinkImpl.DccManSinkInitialize Me
    
    Set m_objDccMan = New CDccMan
    m_objDccMan.Advise Me, m_lContext

End Sub

Private Sub Class_Terminate()
        
    m_objDccMan.Unadvise m_lContext
    Set m_objDccMan = Nothing

    ' Kwyshell. 2004.04.09 UnInstall Com Wrap for Sink Interface
    MDccManSinkImpl.DccManSinkFinalialize Me

End Sub

Public Sub ShowCommSettings()
    m_objDccMan.ShowCommSettings
End Sub

Private Sub IDccManSink_OnLogActive()
    MsgBox "OnLogActive", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogAnswered()
    MsgBox "OnLogAnswered", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogDisconnection()
    MsgBox "OnLogDisconnection", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogError()
    MsgBox "OnLogError", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogInactive()
    MsgBox "OnLogInactive", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogIpAddr(ByVal dwIpAddr As Long)
    MsgBox "OnLogIpAddr IP: " & dwIpAddr
End Sub

Private Sub IDccManSink_OnLogListen()
    MsgBox "OnLogListen", vbOKOnly, "Device Connection Detector"
End Sub

Private Sub IDccManSink_OnLogTerminated()
    MsgBox "OnLogTerminated", vbOKOnly, "Device Connection Detector"
End Sub

