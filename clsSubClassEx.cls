VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubClassEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* clsSubClassEx.cls - SubClassEx Class Module Definition & Code
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
'*              This class also provides extra events so
'*              developer has to do less work.
'*
'* Version control information
'*
'*      Revision:   1.0.0.0
'*      Date:       2001-12-30
'*      Modtime:    09:29 AM
'*      Author:     Stephen Kent
'*
'* ******************************************************
Option Explicit

'Make sure we implement this interface because this is the
'   interface used for all operations in the sub class module.
Implements clsSubClass

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

'Single message type events (some common targets for subclassing)
Public Event Move(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Size(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SetFocus(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event KillFocus(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Enable(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Paint(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event EraseBackground(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SysColorChange(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event ShowWindow(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SetCursor(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event PaintIcon(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event IconEraseBackground(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event DrawItem(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MeasureItem(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event DeleteItem(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event VirtualKeyToItem(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event CharToItem(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SetFont(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event GetFont(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event QueryDragIcon(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event CompareItem(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event WindowPosChanging(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event WindowPosChanged(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Notify(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Help(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NotifyFormat(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event StyleChanged(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event DisplayChange(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event GetIcon(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SetIcon(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientMouseMove(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientLeftButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientLeftButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientLeftButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientRightButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientRightButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientRightButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientMiddleButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientMiddleButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NonClientMiddleButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event KeyDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event KeyUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Char(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SysKeyDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SysKeyUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SysChar(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Command(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event SysCommand(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Timer(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event HorizontalScroll(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event VerticalScroll(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event InitMenu(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event InitMenuPopup(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MenuSelect(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MenuChar(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event EnterIdle(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MenuRightButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MenuDrag(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event UnInitMenuPopup(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MenuCommand(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MouseMove(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event LeftButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event LeftButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event LeftButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event RightButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event RightButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event RightButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MiddleButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MiddleButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MiddleButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MouseWheel(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event ParentNotify(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event EnterMenuLoop(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event ExitMenuLoop(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event NextMenu(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Sizing(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event CaptureChanged(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Moving(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDIRestore(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDINext(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDIMaximize(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDITile(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDICascade(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDIIconArrange(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDIGetActive(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDISetMenu(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event EnterSizeMove(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event ExitSizeMove(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event DropFiles(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MDIRefreshMenu(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MouseHover(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event MouseLeave(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Cut(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Copy(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Paste(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Clear(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event Undo(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event PaletteChanged(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event HotKey(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event PrintClient(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
Public Event UserDefinedMessage(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)

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
            'NOTE:  In SubClassEx we also hand back events for common messages seperately.
            '   Thus if a message is a common target message it will generate 2 events:
            '       It's specific Message Event
            '       The Generic Message Event
            Select Case lMsg
                Case WM_MOVE
                    RaiseEvent Move(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SIZE
                    RaiseEvent Size(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SETFOCUS
                    RaiseEvent SetFocus(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_KILLFOCUS
                    RaiseEvent KillFocus(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_ENABLE
                    RaiseEvent Enable(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_PAINT
                    RaiseEvent Paint(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_ERASEBKGND
                    RaiseEvent EraseBackground(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SYSCOLORCHANGE
                    RaiseEvent SysColorChange(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SHOWWINDOW
                    RaiseEvent ShowWindow(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SETCURSOR
                    RaiseEvent SetCursor(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_PAINTICON
                    RaiseEvent PaintIcon(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_ICONERASEBKGND
                    RaiseEvent IconEraseBackground(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_DRAWITEM
                    RaiseEvent DrawItem(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MEASUREITEM
                    RaiseEvent MeasureItem(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_DELETEITEM
                    RaiseEvent DeleteItem(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_VKEYTOITEM
                    RaiseEvent VirtualKeyToItem(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_CHARTOITEM
                    RaiseEvent CharToItem(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SETFONT
                    RaiseEvent SetFont(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_GETFONT
                    RaiseEvent GetFont(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_QUERYDRAGICON
                    RaiseEvent QueryDragIcon(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_COMPAREITEM
                    RaiseEvent CompareItem(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_WINDOWPOSCHANGING
                    RaiseEvent WindowPosChanging(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_WINDOWPOSCHANGED
                    RaiseEvent WindowPosChanged(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NOTIFY
                    RaiseEvent Notify(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_HELP
                    RaiseEvent Help(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NOTIFYFORMAT
                    RaiseEvent NotifyFormat(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_STYLECHANGED
                    RaiseEvent StyleChanged(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_DISPLAYCHANGE
                    RaiseEvent DisplayChange(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_GETICON
                    RaiseEvent GetIcon(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SETICON
                    RaiseEvent SetIcon(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCMOUSEMOVE
                    RaiseEvent NonClientMouseMove(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCLBUTTONDOWN
                    RaiseEvent NonClientLeftButtonDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCLBUTTONUP
                    RaiseEvent NonClientLeftButtonUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCLBUTTONDBLCLK
                    RaiseEvent NonClientLeftButtonDoubleClick(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCRBUTTONDOWN
                    RaiseEvent NonClientRightButtonDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCRBUTTONUP
                    RaiseEvent NonClientRightButtonUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCRBUTTONDBLCLK
                    RaiseEvent NonClientRightButtonDoubleClick(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCMBUTTONDOWN
                    RaiseEvent NonClientMiddleButtonDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCMBUTTONUP
                    RaiseEvent NonClientMiddleButtonUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NCMBUTTONDBLCLK
                    RaiseEvent NonClientMiddleButtonDoubleClick(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_KEYDOWN
                    RaiseEvent KeyDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_KEYUP
                    RaiseEvent KeyUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_CHAR
                    RaiseEvent Char(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SYSKEYDOWN
                    RaiseEvent SysKeyDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SYSKEYUP
                    RaiseEvent SysKeyUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SYSCHAR
                    RaiseEvent SysChar(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_COMMAND
                    RaiseEvent Command(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SYSCOMMAND
                    RaiseEvent SysCommand(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_TIMER
                    RaiseEvent Timer(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_HSCROLL
                    RaiseEvent HorizontalScroll(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_VSCROLL
                    RaiseEvent VerticalScroll(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_INITMENU
                    RaiseEvent InitMenu(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_INITMENUPOPUP
                    RaiseEvent InitMenuPopup(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MENUSELECT
                    RaiseEvent MenuSelect(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MENUCHAR
                    RaiseEvent MenuChar(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_ENTERIDLE
                    RaiseEvent EnterIdle(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MENURBUTTONUP
                    RaiseEvent MenuRightButtonUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MENUDRAG
                    RaiseEvent MenuDrag(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_UNINITMENUPOPUP
                    RaiseEvent UnInitMenuPopup(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MENUCOMMAND
                    RaiseEvent MenuCommand(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MOUSEMOVE
                    RaiseEvent MouseMove(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_LBUTTONDOWN
                    RaiseEvent LeftButtonDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_LBUTTONUP
                    RaiseEvent LeftButtonUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_LBUTTONDBLCLK
                    RaiseEvent LeftButtonDoubleClick(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_RBUTTONDOWN
                    RaiseEvent RightButtonDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_RBUTTONUP
                    RaiseEvent RightButtonUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_RBUTTONDBLCLK
                    RaiseEvent RightButtonDoubleClick(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MBUTTONDOWN
                    RaiseEvent MiddleButtonDown(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MBUTTONUP
                    RaiseEvent MiddleButtonUp(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MBUTTONDBLCLK
                    RaiseEvent MiddleButtonDoubleClick(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MOUSEWHEEL
                    RaiseEvent MouseWheel(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_PARENTNOTIFY
                    RaiseEvent ParentNotify(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_ENTERMENULOOP
                    RaiseEvent EnterMenuLoop(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_EXITMENULOOP
                    RaiseEvent ExitMenuLoop(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_NEXTMENU
                    RaiseEvent NextMenu(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_SIZING
                    RaiseEvent Sizing(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_CAPTURECHANGED
                    RaiseEvent CaptureChanged(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MOVING
                    RaiseEvent Moving(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDIRESTORE
                    RaiseEvent MDIRestore(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDINEXT
                    RaiseEvent MDINext(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDIMAXIMIZE
                    RaiseEvent MDIMaximize(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDITILE
                    RaiseEvent MDITile(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDICASCADE
                    RaiseEvent MDICascade(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDIICONARRANGE
                    RaiseEvent MDIIconArrange(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDIGETACTIVE
                    RaiseEvent MDIGetActive(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDISETMENU
                    RaiseEvent MDISetMenu(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_ENTERSIZEMOVE
                    RaiseEvent EnterSizeMove(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_EXITSIZEMOVE
                    RaiseEvent ExitSizeMove(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_DROPFILES
                    RaiseEvent DropFiles(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MDIREFRESHMENU
                    RaiseEvent MDIRefreshMenu(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MOUSEHOVER
                    RaiseEvent MouseHover(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_MOUSELEAVE
                    RaiseEvent MouseLeave(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_CUT
                    RaiseEvent Cut(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_COPY
                    RaiseEvent Copy(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_PASTE
                    RaiseEvent Paste(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_CLEAR
                    RaiseEvent Clear(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_UNDO
                    RaiseEvent Undo(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_PALETTECHANGED
                    RaiseEvent PaletteChanged(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_HOTKEY
                    RaiseEvent HotKey(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case WM_PRINTCLIENT
                    RaiseEvent PrintClient(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
                Case Is > WM_USER
                    RaiseEvent UserDefinedMessage(hWnd, lMsg, wParam, lParam, bUserHandled, lRetValue)
            End Select
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

'=====================================
'This section defines the implemented interface
'   It just re-routes the calls to the standard ones.
Private Sub clsSubClass_Attach()
    Attach
End Sub

Private Function clsSubClass_CallBackProc(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    clsSubClass_CallBackProc = CallBackProc(hWnd, lMsg, wParam, lParam)
End Function

Private Sub clsSubClass_Detach()
    Detach
End Sub

Private Property Let clsSubClass_hWnd(ByVal RHS As Long)
    hWnd = RHS
End Property

Private Property Get clsSubClass_hWnd() As Long
    clsSubClass_hWnd = hWnd
End Property

Private Property Get clsSubClass_SubClassed() As Boolean
    clsSubClass_SubClassed = SubClassed
End Property

Private Property Get clsSubClass_Terminating() As Boolean
    clsSubClass_Terminating = Terminating
End Property
