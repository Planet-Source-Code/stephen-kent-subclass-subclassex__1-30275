VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SubClassing Example"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAllowRightButton 
      Caption         =   "Allow Right Button Events"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   2970
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowMiddleButton 
      Caption         =   "Allow Middle Button Events"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   2700
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowLeftButton 
      Caption         =   "Allow Left Button Events"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2430
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowMouseMove 
      Caption         =   "Allow Mouse Move Event"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.Frame frLastEventInfo 
      Caption         =   "Information from Last Event:"
      Height          =   1485
      Left            =   2310
      TabIndex        =   8
      Top             =   0
      Width           =   2235
      Begin VB.Label lblLastEvent3 
         Caption         =   "Event:"
         Height          =   225
         Left            =   90
         TabIndex        =   15
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblLastEvent2 
         Caption         =   "Event:"
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   990
         Width           =   2055
      End
      Begin VB.Label lblLastEvent1 
         Caption         =   "Event:"
         Height          =   225
         Left            =   90
         TabIndex        =   13
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label lblLast3Events 
         Caption         =   "Last 3 Mouse Events Fired:"
         Height          =   225
         Left            =   60
         TabIndex        =   12
         Top             =   570
         Width           =   2085
      End
      Begin VB.Label lblMouseY 
         Caption         =   "Mouse Y Pos:"
         Height          =   225
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblMouseX 
         Caption         =   "Mouse X Pos:"
         Height          =   225
         Left            =   60
         TabIndex        =   9
         Top             =   180
         Width           =   2055
      End
   End
   Begin VB.CheckBox chkAllowSize 
      Caption         =   "Allow Size"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1890
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowMove 
      Caption         =   "Allow Move"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1620
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowSystemMenu 
      Caption         =   "Allow System Menu"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1350
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowContextMenu 
      Caption         =   "Allow Context Menu"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowRestore 
      Caption         =   "Allow Restore"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   810
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowMinimize 
      Caption         =   "Allow Minimize"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowMaximize 
      Caption         =   "Allow Maximize"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   270
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CheckBox chkAllowUnload 
      Caption         =   "Allow Unload"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.Line linBottomBorder 
      X1              =   2310
      X2              =   4500
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line linRightBorder 
      X1              =   4500
      X2              =   4500
      Y1              =   3180
      Y2              =   1500
   End
   Begin VB.Line linLeftBorder 
      X1              =   2310
      X2              =   2310
      Y1              =   1740
      Y2              =   3180
   End
   Begin VB.Label lblTestArea 
      Caption         =   "Form Test Area Below:"
      Height          =   225
      Left            =   2310
      TabIndex        =   19
      Top             =   1530
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SC_SIZE         As Long = 61440
Private Const SC_MOVE         As Long = 61456
Private Const SC_MINIMIZE     As Long = 61472
Private Const SC_MAXIMIZE     As Long = 61488
Private Const SC_CLOSE        As Long = 61536
Private Const SC_KEYMENU      As Long = 61696
Private Const SC_RESTORE      As Long = 61728

'Our sub class object.  WithEvents is vital if you are to receive events
'   from the subclass object.
Private WithEvents oSubClass As clsSubClassEx
Attribute oSubClass.VB_VarHelpID = -1

Private Sub Form_Click()
    'Show our information so user knows event took place
    AddEvent "Click"
End Sub

Private Sub Form_DblClick()
    'Show our information so user knows event took place
    AddEvent "Double Click"
End Sub

Private Sub Form_Load()
    'Create our new subclassing object
    Set oSubClass = New clsSubClassEx
    'We want to subclass this form so pass in the forms hWnd
    oSubClass.hWnd = Me.hWnd
    'Start Subclassing.
    oSubClass.Attach
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Show our information so user knows event took place
    lblMouseX.Caption = "Mouse X Pos: " & CStr(X)
    lblMouseY.Caption = "Mouse Y Pos: " & CStr(Y)
    AddEvent "MouseDown"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Show our information so user knows event took place
    lblMouseX.Caption = "Mouse X Pos: " & CStr(X)
    lblMouseY.Caption = "Mouse Y Pos: " & CStr(Y)
    AddEvent "MouseMove"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Show our information so user knows event took place
    lblMouseX.Caption = "Mouse X Pos: " & CStr(X)
    lblMouseY.Caption = "Mouse Y Pos: " & CStr(Y)
    AddEvent "MouseUp"
End Sub

Private Sub AddEvent(ByVal strEvent As String)
    'Move the event down the list and add a new one to the top of the list
    'The oldest one is simply over written.
    lblLastEvent3.Caption = lblLastEvent2.Caption
    lblLastEvent2.Caption = lblLastEvent1.Caption
    lblLastEvent1.Caption = "Event: " & strEvent
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oSubClass.Detach    'We're finished so stop subclassing
End Sub

Private Sub oSubClass_LeftButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Left Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowLeftButton.Value)
End Sub

Private Sub oSubClass_LeftButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Left Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowLeftButton.Value)
End Sub

Private Sub oSubClass_LeftButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Left Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowLeftButton.Value)
End Sub

Private Sub oSubClass_Message(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    Debug.Print CStr(hWnd), Hex(lMsg), CStr(wParam), CStr(lParam)
    'If message hasn't been handled then account for Context Menu
    If (lMsg = WM_CONTEXTMENU) And Not (bHandled) Then bHandled = Not CBool(chkAllowContextMenu.Value)
End Sub

Private Sub oSubClass_MiddleButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Middle Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowMiddleButton.Value)
End Sub

Private Sub oSubClass_MiddleButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Middle Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowMiddleButton.Value)
End Sub

Private Sub oSubClass_MiddleButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Middle Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowMiddleButton.Value)
End Sub

Private Sub oSubClass_MouseMove(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Mouse Move Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowMouseMove.Value)
End Sub

Private Sub oSubClass_RightButtonDoubleClick(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Right Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowRightButton.Value)
End Sub

Private Sub oSubClass_RightButtonDown(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Right Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowRightButton.Value)
End Sub

Private Sub oSubClass_RightButtonUp(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message hasn't been handled then account for Right Button Allowability
    If Not bHandled Then bHandled = Not CBool(chkAllowRightButton.Value)
End Sub

Private Sub oSubClass_SysCommand(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturnValue As Long)
    'If the message has already been handled then we don't have anything to do
    If Not bHandled Then
        'Select which command we received.
        'NOTE: some have multiple values that are not constants
        '   this is because there are no constants defined, but
        '   those values are used by the system.
        Select Case wParam
            Case SC_CLOSE   'Close Window/Unload Window
                bHandled = Not CBool(chkAllowUnload.Value)
            Case SC_MAXIMIZE, 61490 'Maximize
                bHandled = Not CBool(chkAllowMaximize.Value)
            Case SC_MINIMIZE 'Minimize
                bHandled = Not CBool(chkAllowMinimize.Value)
            Case SC_RESTORE, 61730  'Restore
                bHandled = Not CBool(chkAllowRestore.Value)
            Case SC_KEYMENU, 61587  'System Menu
                bHandled = Not CBool(chkAllowSystemMenu.Value)
            Case SC_MOVE, 61458 'Move
                bHandled = Not CBool(chkAllowMove.Value)
            Case SC_SIZE, 61448, 61446, 61443, 61442, 61441 'Size
                bHandled = Not CBool(chkAllowSize.Value)
        End Select
    End If
End Sub
