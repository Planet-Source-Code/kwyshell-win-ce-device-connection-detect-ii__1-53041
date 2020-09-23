Attribute VB_Name = "MComVBWrap"
'*************************************************************************
' COM-VB Stuff Support Module
' ------------------------------------------------------------------------
' Author:   Kwyshell
' Homepage: http://home.kimo.com.tw/kwyshell/
' Email:    kwyshell@yahoo.com.tw
' Date:     2004.04.09
'*************************************************************************

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Const PAGE_EXECUTE_READWRITE = &H40

' Define a object size
Private Const CVW_OBJECT_SIZE = 4

Public Function ReplaceVtableEntry(pObj As Long, EntryNumber As Integer, ByVal lpfn As Long) As Long

    Dim lOldAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long

    CopyMemory lpVtableHead, ByVal pObj, 4
    lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4
    
    If lpfn <> lOldAddr Then
        Call VirtualProtect(ByVal lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
        CopyMemory ByVal lpfnAddr, lpfn, 4
        Call VirtualProtect(ByVal lpfnAddr, 4, lOldProtect, lOldProtect)
    End If
    
    ReplaceVtableEntry = lOldAddr

End Function

Public Function ObjectToStr(ByVal objRef As Object)

    Dim lObjAdd As Long
    
    lObjAdd = ObjPtr(objRef)
    ObjectToStr = "&H" & Hex(lObjAdd)

End Function
