Attribute VB_Name = "MDccManSinkImpl"
'*************************************************************************
' IDccManSink Module Base Implememtation.
' ------------------------------------------------------------------------
' Author:   Kwyshell
' Homepage: http://home.kimo.com.tw/kwyshell/
' Email:    kwyshell@yahoo.com.tw
' Date:     2004.04.09
' ------------------------------------------------------------------------
' Because VB is not a thread safety program. Trying to do anything at
' different thread is dangerous.
' You can not try to call any thing about COM, Win32 API functions so
' that you have nothing to do for that.
' The only way to run the code without dangerous is trying to deal
' everything at original thread.
'*************************************************************************

Option Explicit

' Private m_objVTableList As New Collection
Private m_lFnOnLogIpAddr As Long
Private m_lFnOnLogTerminated As Long
Private m_lFnOnLogActive As Long
Private m_lFnOnLogInactive As Long
Private m_lFnOnLogAnswered As Long
Private m_lFnOnLogListen As Long
Private m_lFnOnLogDisconnection As Long
Private m_lFnOnLogError As Long

Private Enum EDccManCallID
    eCallIDSink_OnLogIpAddr = 0
    eCallIDSink_OnLogTerminated = 1
    eCallIDSink_OnLogActive = 2
    eCallIDSink_OnLogInactive = 3
    eCallIDSink_OnLogAnswered = 4
    eCallIDSink_OnLogListen = 5
    eCallIDSink_OnLogDisconnection = 6
    eCallIDSink_OnLogError = 7
End Enum

' APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Simple Call Implementation
Private Const m_clMaxParameters As Long = 100
Private m_alParameters(m_clMaxParameters) As Long
Private m_lPP As Long
Private m_bCallLock As Boolean

Public Sub DccManSinkInitialize(ByVal this As IDccManSink)

    Dim lObjAddress As Long
    
    ' Object Address
    lObjAddress = ObjPtr(this)

    ' Replace VTable Function Entries
    m_lFnOnLogIpAddr = _
        ReplaceVtableEntry(lObjAddress, &H4, AddressOf IDccManSink_OnLogIpAddr)
    m_lFnOnLogTerminated = _
        ReplaceVtableEntry(lObjAddress, &H5, AddressOf IDccManSink_OnLogTerminated)
    m_lFnOnLogActive = _
        ReplaceVtableEntry(lObjAddress, &H6, AddressOf IDccManSink_OnLogActive)
    m_lFnOnLogInactive = _
        ReplaceVtableEntry(lObjAddress, &H7, AddressOf IDccManSink_OnLogInactive)
    m_lFnOnLogAnswered = _
        ReplaceVtableEntry(lObjAddress, &H8, AddressOf IDccManSink_OnLogAnswered)
    m_lFnOnLogListen = _
        ReplaceVtableEntry(lObjAddress, &H9, AddressOf IDccManSink_OnLogListen)
    m_lFnOnLogDisconnection = _
        ReplaceVtableEntry(lObjAddress, &HA, AddressOf IDccManSink_OnLogDisconnection)
    m_lFnOnLogError = _
        ReplaceVtableEntry(lObjAddress, &HB, AddressOf IDccManSink_OnLogError)
        
End Sub

Public Sub DccManSinkFinalialize(ByVal this As IDccManSink)
   
    Dim lObjAddress As Long
            
    ' Object Address
    lObjAddress = ObjPtr(this)
    
    ' Restore VTable Function Entries
    m_lFnOnLogIpAddr = _
        ReplaceVtableEntry(lObjAddress, &H4, m_lFnOnLogIpAddr)
    m_lFnOnLogTerminated = _
        ReplaceVtableEntry(lObjAddress, &H5, m_lFnOnLogTerminated)
    m_lFnOnLogActive = _
        ReplaceVtableEntry(lObjAddress, &H6, m_lFnOnLogActive)
    m_lFnOnLogInactive = _
        ReplaceVtableEntry(lObjAddress, &H7, m_lFnOnLogInactive)
    m_lFnOnLogAnswered = _
        ReplaceVtableEntry(lObjAddress, &H8, m_lFnOnLogAnswered)
    m_lFnOnLogListen = _
        ReplaceVtableEntry(lObjAddress, &H9, m_lFnOnLogListen)
    m_lFnOnLogDisconnection = _
        ReplaceVtableEntry(lObjAddress, &HA, m_lFnOnLogDisconnection)
    m_lFnOnLogError = _
        ReplaceVtableEntry(lObjAddress, &HB, m_lFnOnLogError)

End Sub

Private Function DccManSink_SafetyHandler(p1 As Long, p2 As Long, p3 As Long, p4 As Long) As Long

    ' Retrieve this Pointer
    ' Call this md before recoverying vtable
#If 0 Then
    Dim objDccMainSink As IDccManSink
    Dim objDummy As IDccManSink
    
    If p1 = 0 Then
        DccManSink_SafetyHandler = 0
        Exit Function
    End If
    
    CopyMemory objDummy, ByVal p1, 4
    Set objDccMainSink = objDummy
    CopyMemory objDummy, 0&, 4
    
    If Not (objDccMainSink Is Nothing) Then
        DccManSink_SafetyHandler = 0
        Exit Function
    End If
#End If

    Select Case p2
        Case eCallIDSink_OnLogIpAddr
            MsgBox "OnLogIpAddr IP: " & p3, vbOKOnly, "Device Connection Detector"
        Case eCallIDSink_OnLogTerminated
            MsgBox "OnLogTerminated", vbOKOnly, "Device Connection Detector"
        Case eCallIDSink_OnLogActive
            MsgBox "OnLogActive", vbOKOnly, "Device Connection Detector"
        Case eCallIDSink_OnLogInactive
            MsgBox "OnLogInactive", vbOKOnly, "Device Connection Detector"
        Case eCallIDSink_OnLogAnswered
            MsgBox "OnLogAnswered", vbOKOnly, "Device Connection Detector"
        Case eCallIDSink_OnLogListen
            MsgBox "OnLogListen", vbOKOnly, "Device Connection Detector"
        Case eCallIDSink_OnLogDisconnection
            MsgBox "OnLogDisconnection", vbOKOnly, "Device Connection Detector"
        Case eCallIDSink_OnLogError
            MsgBox "OnLogError", vbOKOnly, "Device Connection Detector"
    End Select
    
    DccManSink_SafetyHandler = 0
    
#If 0 Then
    Set objDccMainSink = Nothing
#End If

End Function

' Simple version for processing commands
Public Function DccManSinkImpl_CallProc()

    Dim i As Integer
    
    ' Detect CallLock
    If m_bCallLock Or m_lPP <= 0 Then
        Exit Function
    End If
    
    ' Do Parameters
    For i = 0 To m_lPP Step 4
        DccManSink_SafetyHandler m_alParameters(i), m_alParameters(i + 1), _
            m_alParameters(i + 2), m_alParameters(i + 3)
    Next
    
    m_lPP = 0

End Function

' Kwyshell. Never Call this in ComWrap Function
Private Function SCDCallParams(param1 As Long, Optional param2 As Long = 0, Optional param3 As Long = 0, Optional param4 As Long = 0) As Long

    ' No API call here. We cannot do anything for synchronization. Simple Lock only
    m_bCallLock = True

    If m_lPP <= m_clMaxParameters - 4 Then
        m_alParameters(m_lPP) = param1
        m_alParameters(m_lPP + 1) = param2
        m_alParameters(m_lPP + 2) = param3
        m_alParameters(m_lPP + 3) = param4
        m_lPP = m_lPP + 4
    End If
    
    m_bCallLock = False

End Function

Private Function IDccManSink_OnLogIpAddr(ByVal this As IDccManSink, ByVal dwIpAddr As Long) As Long
    IDccManSink_OnLogIpAddr = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogIpAddr, dwIpAddr)
End Function

Private Function IDccManSink_OnLogTerminated(ByVal this As IDccManSink) As Long
    IDccManSink_OnLogTerminated = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogTerminated)
End Function

Private Function IDccManSink_OnLogActive(ByVal this As IDccManSink) As Long
    IDccManSink_OnLogActive = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogActive)
End Function

Private Function IDccManSink_OnLogInactive(ByVal this As IDccManSink) As Long
    IDccManSink_OnLogInactive = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogInactive)
End Function

Private Function IDccManSink_OnLogAnswered(ByVal this As IDccManSink) As Long
    IDccManSink_OnLogAnswered = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogAnswered)
End Function

Private Function IDccManSink_OnLogListen(ByVal this As IDccManSink) As Long
    IDccManSink_OnLogListen = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogListen)
End Function

Private Function IDccManSink_OnLogDisconnection(ByVal this As IDccManSink) As Long
    IDccManSink_OnLogDisconnection = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogDisconnection)
End Function

Private Function IDccManSink_OnLogError(ByVal this As IDccManSink) As Long
    IDccManSink_OnLogError = SCDCallParams(ObjPtr(this), eCallIDSink_OnLogError)
End Function





