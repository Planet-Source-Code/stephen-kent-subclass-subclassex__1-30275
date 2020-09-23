Attribute VB_Name = "modSubClass"
'* modSubClass.mod - Sub Class Module for SubClass Tracking/Routing
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
'* Description: This is the code required to do callback
'*              processing and tracking/routing of messages
'*              to the appropriate subclass object
'*
'* Version control information
'*
'*      Revision:   1.0.0.0
'*      Date:       2001-12-30
'*      Modtime:    09:26 AM
'*      Author:     Stephen Kent
'*
'* ******************************************************
Option Explicit

'Lots of message constants (All I could find that were defined)
Public Const WM_NULL                         As Long = &H0
Public Const WM_CREATE                       As Long = &H1
Public Const WM_DESTROY                      As Long = &H2
Public Const WM_MOVE                         As Long = &H3
Public Const WM_SIZE                         As Long = &H5
Public Const WM_ACTIVATE                     As Long = &H6
Public Const WM_SETFOCUS                     As Long = &H7
Public Const WM_KILLFOCUS                    As Long = &H8
Public Const WM_ENABLE                       As Long = &HA
Public Const WM_SETREDRAW                    As Long = &HB
Public Const WM_SETTEXT                      As Long = &HC
Public Const WM_GETTEXT                      As Long = &HD
Public Const WM_GETTEXTLENGTH                As Long = &HE
Public Const WM_PAINT                        As Long = &HF
Public Const WM_CLOSE                        As Long = &H10
Public Const WM_QUERYENDSESSION              As Long = &H11
Public Const WM_QUIT                         As Long = &H12
Public Const WM_QUERYOPEN                    As Long = &H13
Public Const WM_ERASEBKGND                   As Long = &H14
Public Const WM_SYSCOLORCHANGE               As Long = &H15
Public Const WM_ENDSESSION                   As Long = &H16
Public Const WM_SHOWWINDOW                   As Long = &H18
Public Const WM_WININICHANGE                 As Long = &H1A
Public Const WM_SETTINGCHANGE                As Long = WM_WININICHANGE
Public Const WM_DEVMODECHANGE                As Long = &H1B
Public Const WM_ACTIVATEAPP                  As Long = &H1C
Public Const WM_FONTCHANGE                   As Long = &H1D
Public Const WM_TIMECHANGE                   As Long = &H1E
Public Const WM_CANCELMODE                   As Long = &H1F
Public Const WM_SETCURSOR                    As Long = &H20
Public Const WM_MOUSEACTIVATE                As Long = &H21
Public Const WM_CHILDACTIVATE                As Long = &H22
Public Const WM_QUEUESYNC                    As Long = &H23
Public Const WM_GETMINMAXINFO                As Long = &H24
Public Const WM_PAINTICON                    As Long = &H26
Public Const WM_ICONERASEBKGND               As Long = &H27
Public Const WM_NEXTDLGCTL                   As Long = &H28
Public Const WM_SPOOLERSTATUS                As Long = &H2A
Public Const WM_DRAWITEM                     As Long = &H2B
Public Const WM_MEASUREITEM                  As Long = &H2C
Public Const WM_DELETEITEM                   As Long = &H2D
Public Const WM_VKEYTOITEM                   As Long = &H2E
Public Const WM_CHARTOITEM                   As Long = &H2F
Public Const WM_SETFONT                      As Long = &H30
Public Const WM_GETFONT                      As Long = &H31
Public Const WM_SETHOTKEY                    As Long = &H32
Public Const WM_GETHOTKEY                    As Long = &H33
Public Const WM_QUERYDRAGICON                As Long = &H37
Public Const WM_COMPAREITEM                  As Long = &H39
Public Const WM_GETOBJECT                    As Long = &H3D
Public Const WM_COMPACTING                   As Long = &H41
Public Const WM_COMMNOTIFY                   As Long = &H44    'no longer suported
Public Const WM_WINDOWPOSCHANGING            As Long = &H46
Public Const WM_WINDOWPOSCHANGED             As Long = &H47
Public Const WM_POWER                        As Long = &H48
Public Const WM_COPYDATA                     As Long = &H4A
Public Const WM_CANCELJOURNAL                As Long = &H4B
Public Const WM_NOTIFY                       As Long = &H4E
Public Const WM_INPUTLANGCHANGEREQUEST       As Long = &H50
Public Const WM_INPUTLANGCHANGE              As Long = &H51
Public Const WM_TCARD                        As Long = &H52
Public Const WM_HELP                         As Long = &H53
Public Const WM_USERCHANGED                  As Long = &H54
Public Const WM_NOTIFYFORMAT                 As Long = &H55
Public Const WM_CONTEXTMENU                  As Long = &H7B
Public Const WM_STYLECHANGING                As Long = &H7C
Public Const WM_STYLECHANGED                 As Long = &H7D
Public Const WM_DISPLAYCHANGE                As Long = &H7E
Public Const WM_GETICON                      As Long = &H7F
Public Const WM_SETICON                      As Long = &H80
Public Const WM_NCCREATE                     As Long = &H81
Public Const WM_NCDESTROY                    As Long = &H82
Public Const WM_NCCALCSIZE                   As Long = &H83
Public Const WM_NCHITTEST                    As Long = &H84
Public Const WM_NCPAINT                      As Long = &H85
Public Const WM_NCACTIVATE                   As Long = &H86
Public Const WM_GETDLGCODE                   As Long = &H87
Public Const WM_SYNCPAINT                    As Long = &H88
Public Const WM_NCMOUSEMOVE                  As Long = &HA0
Public Const WM_NCLBUTTONDOWN                As Long = &HA1
Public Const WM_NCLBUTTONUP                  As Long = &HA2
Public Const WM_NCLBUTTONDBLCLK              As Long = &HA3
Public Const WM_NCRBUTTONDOWN                As Long = &HA4
Public Const WM_NCRBUTTONUP                  As Long = &HA5
Public Const WM_NCRBUTTONDBLCLK              As Long = &HA6
Public Const WM_NCMBUTTONDOWN                As Long = &HA7
Public Const WM_NCMBUTTONUP                  As Long = &HA8
Public Const WM_NCMBUTTONDBLCLK              As Long = &HA9
Public Const WM_KEYFIRST                     As Long = &H100
Public Const WM_KEYDOWN                      As Long = &H100
Public Const WM_KEYUP                        As Long = &H101
Public Const WM_CHAR                         As Long = &H102
Public Const WM_DEADCHAR                     As Long = &H103
Public Const WM_SYSKEYDOWN                   As Long = &H104
Public Const WM_SYSKEYUP                     As Long = &H105
Public Const WM_SYSCHAR                      As Long = &H106
Public Const WM_SYSDEADCHAR                  As Long = &H107
Public Const WM_KEYLAST                      As Long = &H108
Public Const WM_IME_STARTCOMPOSITION         As Long = &H10D
Public Const WM_IME_ENDCOMPOSITION           As Long = &H10E
Public Const WM_IME_COMPOSITION              As Long = &H10F
Public Const WM_IME_KEYLAST                  As Long = &H10F
Public Const WM_INITDIALOG                   As Long = &H110
Public Const WM_COMMAND                      As Long = &H111
Public Const WM_SYSCOMMAND                   As Long = &H112
Public Const WM_TIMER                        As Long = &H113
Public Const WM_HSCROLL                      As Long = &H114
Public Const WM_VSCROLL                      As Long = &H115
Public Const WM_INITMENU                     As Long = &H116
Public Const WM_INITMENUPOPUP                As Long = &H117
Public Const WM_MENUSELECT                   As Long = &H11F
Public Const WM_MENUCHAR                     As Long = &H120
Public Const WM_ENTERIDLE                    As Long = &H121
Public Const WM_MENURBUTTONUP                As Long = &H122
Public Const WM_MENUDRAG                     As Long = &H123
Public Const WM_MENUGETOBJECT                As Long = &H124
Public Const WM_UNINITMENUPOPUP              As Long = &H125
Public Const WM_MENUCOMMAND                  As Long = &H126
Public Const WM_CTLCOLORMSGBOX               As Long = &H132
Public Const WM_CTLCOLOREDIT                 As Long = &H133
Public Const WM_CTLCOLORLISTBOX              As Long = &H134
Public Const WM_CTLCOLORBTN                  As Long = &H135
Public Const WM_CTLCOLORDLG                  As Long = &H136
Public Const WM_CTLCOLORSCROLLBAR            As Long = &H137
Public Const WM_CTLCOLORSTATIC               As Long = &H138
Public Const WM_MOUSEFIRST                   As Long = &H200
Public Const WM_MOUSEMOVE                    As Long = &H200
Public Const WM_LBUTTONDOWN                  As Long = &H201
Public Const WM_LBUTTONUP                    As Long = &H202
Public Const WM_LBUTTONDBLCLK                As Long = &H203
Public Const WM_RBUTTONDOWN                  As Long = &H204
Public Const WM_RBUTTONUP                    As Long = &H205
Public Const WM_RBUTTONDBLCLK                As Long = &H206
Public Const WM_MBUTTONDOWN                  As Long = &H207
Public Const WM_MBUTTONUP                    As Long = &H208
Public Const WM_MBUTTONDBLCLK                As Long = &H209
Public Const WM_MOUSEWHEEL                   As Long = &H20A
Public Const WM_MOUSELAST                    As Long = &H20A
Public Const WM_PARENTNOTIFY                 As Long = &H210
Public Const WM_ENTERMENULOOP                As Long = &H211
Public Const WM_EXITMENULOOP                 As Long = &H212
Public Const WM_NEXTMENU                     As Long = &H213
Public Const WM_SIZING                       As Long = &H214
Public Const WM_CAPTURECHANGED               As Long = &H215
Public Const WM_MOVING                       As Long = &H216
Public Const WM_POWERBROADCAST               As Long = &H218
Public Const WM_DEVICECHANGE                 As Long = &H219
Public Const WM_MDICREATE                    As Long = &H220
Public Const WM_MDIDESTROY                   As Long = &H221
Public Const WM_MDIACTIVATE                  As Long = &H222
Public Const WM_MDIRESTORE                   As Long = &H223
Public Const WM_MDINEXT                      As Long = &H224
Public Const WM_MDIMAXIMIZE                  As Long = &H225
Public Const WM_MDITILE                      As Long = &H226
Public Const WM_MDICASCADE                   As Long = &H227
Public Const WM_MDIICONARRANGE               As Long = &H228
Public Const WM_MDIGETACTIVE                 As Long = &H229
Public Const WM_MDISETMENU                   As Long = &H230
Public Const WM_ENTERSIZEMOVE                As Long = &H231
Public Const WM_EXITSIZEMOVE                 As Long = &H232
Public Const WM_DROPFILES                    As Long = &H233
Public Const WM_MDIREFRESHMENU               As Long = &H234
Public Const WM_IME_SETCONTEXT               As Long = &H281
Public Const WM_IME_NOTIFY                   As Long = &H282
Public Const WM_IME_CONTROL                  As Long = &H283
Public Const WM_IME_COMPOSITIONFULL          As Long = &H284
Public Const WM_IME_SELECT                   As Long = &H285
Public Const WM_IME_CHAR                     As Long = &H286
Public Const WM_IME_REQUEST                  As Long = &H288
Public Const WM_IME_KEYDOWN                  As Long = &H290
Public Const WM_IME_KEYUP                    As Long = &H291
Public Const WM_MOUSEHOVER                   As Long = &H2A1
Public Const WM_MOUSELEAVE                   As Long = &H2A3
Public Const WM_CUT                          As Long = &H300
Public Const WM_COPY                         As Long = &H301
Public Const WM_PASTE                        As Long = &H302
Public Const WM_CLEAR                        As Long = &H303
Public Const WM_UNDO                         As Long = &H304
Public Const WM_RENDERFORMAT                 As Long = &H305
Public Const WM_RENDERALLFORMATS             As Long = &H306
Public Const WM_DESTROYCLIPBOARD             As Long = &H307
Public Const WM_DRAWCLIPBOARD                As Long = &H308
Public Const WM_PAINTCLIPBOARD               As Long = &H309
Public Const WM_VSCROLLCLIPBOARD             As Long = &H30A
Public Const WM_SIZECLIPBOARD                As Long = &H30B
Public Const WM_ASKCBFORMATNAME              As Long = &H30C
Public Const WM_CHANGECBCHAIN                As Long = &H30D
Public Const WM_HSCROLLCLIPBOARD             As Long = &H30E
Public Const WM_QUERYNEWPALETTE              As Long = &H30F
Public Const WM_PALETTEISCHANGING            As Long = &H310
Public Const WM_PALETTECHANGED               As Long = &H311
Public Const WM_HOTKEY                       As Long = &H312
Public Const WM_PRINT                        As Long = &H317
Public Const WM_PRINTCLIENT                  As Long = &H318
Public Const WM_HANDHELDFIRST                As Long = &H358
Public Const WM_HANDHELDLAST                 As Long = &H35F
Public Const WM_AFXFIRST                     As Long = &H360
Public Const WM_AFXLAST                      As Long = &H37F
Public Const WM_PENWINFIRST                  As Long = &H380
Public Const WM_PENWINLAST                   As Long = &H38F
Public Const WM_APP                          As Long = 32768

'Messages above this number (except WM_APP) are private messages
Public Const WM_USER                         As Long = &H400

'Variable to hold the collection of subclass objects for message routing
Private m_colSubClassObjects As Collection

'This is the call back routine that will replace the window's message routine
Public Function lCallBackProc(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim oSubClassObject As clsSubClass

    'Check to see if we have our subclass object collection
    If Not (m_colSubClassObjects Is Nothing) Then
        'Loop through all our subclass objects
        For Each oSubClassObject In m_colSubClassObjects
            'Check to see that the object has the window handle and has subclassed it.
            If (oSubClassObject.hWnd = hWnd) And (oSubClassObject.SubClassed) Then
                'Pass the call to the sub class object for processing
                lCallBackProc = oSubClassObject.CallBackProc(hWnd, lMsg, wParam, lParam)
                'We don't want any crashes so only the first object that
                '   subclassed a window will get that message. (Otherwise
                '   it can create an infinite message loop are really screw
                '   things up)
                Exit For
            End If
        Next
    Else
        'No collection so we're stumped - Return False because there's nothing we can do.
        lCallBackProc = False
    End If
End Function

Public Sub AddSubClassObject(oSubClassObject As clsSubClass)
    'Check to see that we have a collection already in existence
    If m_colSubClassObjects Is Nothing Then
        'Nope, Then create one
        Set m_colSubClassObjects = New Collection
    End If
    'Add to the collection
    m_colSubClassObjects.Add oSubClassObject
End Sub

Public Sub RemoveSubClassObject(oSubClassObject As clsSubClass)
    Dim lIndex As Long

    'Check to make sure we have a collection to remove from
    If Not (m_colSubClassObjects Is Nothing) Then
        'loop through all entries until we find a match
        For lIndex = 1 To m_colSubClassObjects.Count
            'Check to see that the hwnd matches and the object is terminating (only terminating when trying to remove itself from the collection)
            If (m_colSubClassObjects(lIndex).hWnd = oSubClassObject.hWnd) And (m_colSubClassObjects(lIndex).Terminating) Then
                'We found a match so remove it and exit the loop
                m_colSubClassObjects.Remove lIndex
                Exit For
            End If
        Next
        'Check to see if there are more objects in collection
        If m_colSubClassObjects.Count = 0 Then
            'No, so destroy the collection
            Set m_colSubClassObjects = Nothing
        End If
    End If
End Sub
