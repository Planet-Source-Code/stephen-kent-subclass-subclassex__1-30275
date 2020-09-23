VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* clsSubClass.cls - SubClass Class Module Definition & Code
'* ******************************************************
'*
'* Revision:    1.0.0.0
'*
'* ******************************************************
'* Copyright (C) 2001
'* Stephen Kent.
'* ******************************************************
'*
'* Created by:  Stephen Kent
'*
'* Created on:  2001-12-28
'*
'* Project:     SubClass
'*
'* Description: This is the class which basically does all
'*              the grunt work involved in subclassing.
'*
'* Version control information
'*
'*      Revision:   1.0.0.0
'*      Date:       2001-12-30
'*      Modtime:    09:28 AM
'*      Author:     Stephen Kent
'*
'* ******************************************************
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private m_bSubClassed As Boolean
Private m_bTerminating As Boolean
Private m_lhWnd As Long
Private m_lPrevWindowProc As Long

'Generic Message Event (Called for all messages including WM_NULL)
Public Event Message(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)

Public Function CallBackProc(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim bUserHandled As Boolean
    Dim lRetValue As Long

    bUserHandled = False
    lRetValue = 0
    'Check to make sure the specified hwnd is a valid window and it matches our handle
    If IsWindow(hWnd) And (hWnd = m_lhWnd) Then
        'have we subclassed the window?
        If m_bSubClassed Then
            'We have subclassed so we need to hand an event back to the developer
            '   bUserHandled is to determine if the developer handled then message
            '   lRetValue determines what value to send back with the message (only if user handled message)
            RaiseEvent Message(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
            'Did user fail to handle message?
            If Not bUserHandled Then
                'Yes so we need to call the default handler for the window
                CallBackProc = CallWindowProc(m_lPrevWindowProc, hWnd, lMsg, wParam, lParam)
            Else
                'No, User fully handled message so we return the value they gave us.
                CallBackProc = lRetValue
            End If
        Else
            'We haven't subclassed so send the message on to the window without further processing.
            CallBackProc = SendMessage(hWnd, lMsg, wParam, lParam)
        End If
    End If
End Function

'Sub-Procedure to start subclassing an object
Public Sub Attach()
    'Check that we haven't already subclassed and the the window handle is valid.
    If (Not (m_bSubClassed)) And (IsWindow(m_lhWnd)) Then
        'Set the new message handler and save the address of the old one.
        m_lPrevWindowProc = SetWindowLong(m_lhWnd, GWL_WNDPROC, AddressOf lCallBackProc)
        'Mark that we've subclassed the window
        m_bSubClassed = True
    End If
End Sub

'Sub-Procedure to stop subclassing an object
Public Sub Detach()
    'Check that we have already subclassed and the the window handle is valid.
    If m_bSubClassed And (IsWindow(m_lhWnd)) Then
        'Restore the old message handler.
        SetWindowLong m_lhWnd, GWL_WNDPROC, m_lPrevWindowProc
        'Mark that we no longer have control of the window
        m_bSubClassed = False
    End If
End Sub

Private Sub Class_Initialize()
    'Basic Initialization (no really needed but a good habit)
    m_lhWnd = 0
    m_bSubClassed = False
    m_lPrevWindowProc = 0
    m_bTerminating = False
    'Add this object into the subclass collection so that the message
    '   handler knows which object to send messages to.
    AddSubClassObject Me
End Sub

Private Sub Class_Terminate()
    'Mark that we're terminating (so Remove knows which one to remove)
    m_bTerminating = True
    'Remove ourselves from the subclass collection to prevent errors
    '   and speed up response time for other objects.
    RemoveSubClassObject Me
    'Finally un-subclass the window (auto-checks to make sure it was subclassed)
    Detach
End Sub

Public Property Get SubClassed() As Boolean
    SubClassed = m_bSubClassed
End Property

Public Property Get Terminating() As Boolean
    Terminating = m_bTerminating
End Property

Public Property Get hWnd() As Long
    hWnd = m_lhWnd
End Property

Public Property Let hWnd(ByVal lNewhWnd As Long)
    'if window handles don't match then change
    If lNewhWnd <> m_lhWnd Then
        'Un-subclass the window because the hwnd is changing (auto-checks to make sure it was subclassed)
        Detach
        'Assign the new window handle
        m_lhWnd = lNewhWnd
    End If
End Property
