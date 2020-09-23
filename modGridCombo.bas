Attribute VB_Name = "modGridCombo"
Option Explicit
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
'Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Declare Function ValidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'Public Declare Function GetForegroundWindow Lib "user32" () As Long
'Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long


Public Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Public Const SPI_GETWORKAREA = 48
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const WM_SETTEXT = &HC
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_CAPTURECHANGED = &H215
Public Const WM_USER = &H400
Public Const WM_SELECTNOW = WM_USER + 100
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_NCHITTEST = &H84
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = -4
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000
Public Const HTCLIENT = 1
Public Const SB_ENDSCROLL = 8
Public Const DFC_SCROLL = 3
Public Const DFCS_SCROLLDOWN = &H1
Public Const DFCS_PUSHED = &H200
Public Const DFCS_FLAT = &H4000
Public Const DFCS_INACTIVE = &H100
Public Const WSVB_LOCKED = 2048
Public Const WMDG_NOTIFY = WM_USER + 20
'Public Const SWP_NOMOVE = &H2
'Public Const HWND_NOTOPMOST = -2
'Public Const WM_ENABLE = &HA
'Public Const WM_RBUTTONDOWN = &H204
'Public Const WM_MBUTTONDOWN = &H207
'Public Const WM_KILLFOCUS = &H8
'Public Const WM_SETFOCUS = &H7
'Public Const WM_KEYDOWN = &H100
'Public Const WM_NCLBUTTONDOWN = &HA1
'Public Const WM_SHOWWINDOW = &H18
'Public Const WM_PAGEUP = 33
'Public Const WM_PAGEDOWN = 34
'Public Const WM_ARROWDOWN = 40
'Public Const WM_ARROWUP = 38
'Public Const WM_ENTER = 13
'Public Const WM_ESCAPE = 27
'Public Const WM_DESTROY = &H2
'Public Const WM_PAINT = &HF
'Public Const WM_KEYUP = &H101
'Public Const SW_SHOWNOACTIVATE = 4
'Public Const SW_SHOWNA = 8
'Public Const HTHSCROLL = 6
'Public Const HTVSCROLL = 7
'Public Const GW_CHILD = 5

Public Type pDataStruct
    IsGridHooked As Boolean
    Selection As Boolean
    hwndEditBox As Long
    hwndParent As Long
    hwndGrid As Long
    PrevGridProc As Long
    DataGridCntl As DataGrid
    Top As Integer
    BColumn As Integer
    PictureCtl As PictureBox
    AbsolutePos As Integer
End Type
Private Scrolling As Integer

Public pData As pDataStruct

Function DataGridWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim bCallOriginalProc As Boolean, lResult As Long
    Dim LowWord As Long, HIWORD As Long
    bCallOriginalProc = True
    lResult = 0
On Error GoTo DataGridWindowProc_Err
    Select Case uMsg
        Case WM_LBUTTONDOWN
            Call GetHiLoWord(lParam, LowWord, HIWORD)
            Call DataGridCntl_OnLButtonDown(hw, wParam, LowWord, HIWORD)
        Case WM_LBUTTONUP
            Call GetHiLoWord(lParam, LowWord, HIWORD)
            Call DataGridCntl_OnLButtonUp(hw, wParam, LowWord, HIWORD)
        Case WM_MOUSEMOVE
            Call GetHiLoWord(lParam, LowWord, HIWORD)
            Call DataGridCntl_OnMouseMove(hw, wParam, LowWord, HIWORD)
        Case WM_CAPTURECHANGED
            Call DataGridCntl_OnCaptureChanged(hw, wParam)
        Case WM_VSCROLL
            Call GetHiLoWord(wParam, LowWord, HIWORD)
            lResult = CallWindowProc(pData.PrevGridProc, hw, uMsg, wParam, lParam)
            Call DataGridCntl_OnVScroll(hw, LowWord, HIWORD, lParam)
            bCallOriginalProc = False
        Case WM_HSCROLL
            Call GetHiLoWord(wParam, LowWord, HIWORD)
            Call DataGridCntl_OnHScroll(hw, LowWord, HIWORD, lParam)
        Case WM_SELECTNOW
            Call DataGridCntl_OnSelectNow(hw)
        Case WM_LBUTTONDBLCLK
            Call DataGridCntl_OnLButtonDown(hw, wParam, LowWord, HIWORD)
        Case WMDG_NOTIFY
            Scrolling = 0
            Debug.Print "Scroll Notify"
            bCallOriginalProc = False
    End Select
    If bCallOriginalProc = True Then
        lResult = CallWindowProc(pData.PrevGridProc, hw, uMsg, wParam, lParam)
    End If
DataGridWindowProc_Exit:
    DataGridWindowProc = lResult
    Exit Function
DataGridWindowProc_Err:
    MsgBox Err.Description, vbCritical, "DataGridWindowProc"
    lResult = 0
    Resume DataGridWindowProc_Exit
End Function

Private Sub DataGridCntl_OnLButtonDown(hw As Long, wParam As Long, x As Long, y As Long)
    Dim Rc As RECT, pt As POINTAPI, nHitTest As Long
    Dim ScrHnd As Long, LngParam As Long, i As Integer
On Error GoTo DataGridCntl_OnLButtonDown_Err
    Call GetCursorPos(pt)
    LngParam = MakeLParm(pt.x, pt.y)
    nHitTest = CallWindowProc(ByVal pData.PrevGridProc, ByVal hw, WM_NCHITTEST, 0, ByVal LngParam)
    If nHitTest = HTCLIENT Then
        ScrHnd = WindowFromPoint(ByVal pt.x, ByVal pt.y)
        If ScrHnd <> pData.hwndGrid Then
            If ScreenToClient(ByVal ScrHnd, pt) <> 0 Then
                LngParam = MakeLParm(pt.x, pt.y)
                Scrolling = 1
                Call PostMessage(ByVal ScrHnd, ByVal WM_LBUTTONDOWN, ByVal 1, ByVal LngParam)
            End If
        Else
            On Error Resume Next
            i = pData.DataGridCntl.RowBookmark(pData.DataGridCntl.RowContaining((Screen.TwipsPerPixelY * (y - Rc.Top))))
            i = Err.Number
            If i <> 0 Then
                Call ReleaseCapture
            End If
            On Error GoTo DataGridCntl_OnLButtonDown_Err
        End If
    Else
        Call GetWindowRect(ByVal hw, Rc)
        If PtInRect(Rc, ByVal pt.x, ByVal pt.y) = 0 Then
            Call ReleaseCapture
        End If
    End If
DataGridCntl_OnLButtonDown_Exit:
    Exit Sub
DataGridCntl_OnLButtonDown_Err:
    MsgBox Err.Description, vbCritical, "DataGridCntl_OnLButtonDown"
    Resume DataGridCntl_OnLButtonDown_Exit
End Sub

Private Sub DataGridCntl_OnLButtonUp(hwnd As Long, wParam As Long, x As Long, y As Long)
    Dim Rc As RECT, pt As POINTAPI
    Dim LngParam As Long, i As Integer, nHitTest As Long, StrText As String
On Error GoTo DataGridCntl_OnLButtonUp_Err
    ' Check if the point is on the client rectangle. If it is,
    ' & there is a valid selection, release the mouse capture &
    ' let the window get hidden.
    Call GetClientRect(ByVal hwnd, Rc)
    If PtInRect(Rc, ByVal x, ByVal y) <> 0 Then
        If pData.DataGridCntl.ColContaining(Screen.TwipsPerPixelY * (x - Rc.Left)) >= 0 Then
            On Error Resume Next
            pData.DataGridCntl.Bookmark = pData.DataGridCntl.RowBookmark(pData.DataGridCntl.RowContaining((Screen.TwipsPerPixelY * (y - Rc.Top))))
            i = Err.Number
            If i <> 0 Then
                Call ReleaseCapture
            Else
                If Not (GetWindowLong(pData.hwndEditBox, GWL_STYLE) And WSVB_LOCKED) = WSVB_LOCKED Then
                    pData.AbsolutePos = GetRecNo(pData.DataGridCntl.DataSource)
                    StrText = pData.DataGridCntl.Columns(pData.BColumn).Text
                    Call ReleaseCapture
                    pData.Selection = True
                    Call SendMessage(ByVal pData.hwndEditBox, ByVal WM_SETTEXT, ByVal 0, ByVal CStr(StrText))
                    pData.Selection = False
                Else
                    Call ReleaseCapture
                End If
            End If
            On Error GoTo DataGridCntl_OnLButtonUp_Err
        End If
    End If
DataGridCntl_OnLButtonUp_Exit:
    Exit Sub
DataGridCntl_OnLButtonUp_Err:
    MsgBox Err.Description, vbCritical, "DataGridCntl_OnLButtonUp"
    Resume DataGridCntl_OnLButtonUp_Exit
End Sub

Private Sub DataGridCntl_OnSelectNow(hwnd As Long)
On Error GoTo DataGridCntl_OnSelectNow_Err
    Call ReleaseCapture
DataGridCntl_OnSelectNow_Exit:
    Exit Sub
DataGridCntl_OnSelectNow_Err:
    MsgBox Err.Description, vbCritical, "DataGridCntl_OnSelectNow"
    Resume DataGridCntl_OnSelectNow_Exit
End Sub

Private Sub DataGridCntl_OnMouseMove(hwnd As Long, wParam As Long, x As Long, y As Long)
    Dim Rc As RECT
On Error GoTo DataGridCntl_OnMouseMove_Err
    ' Keep track of the mouse movement & change the selected item.
    Call GetClientRect(ByVal hwnd, Rc)
    If PtInRect(Rc, ByVal x, ByVal y) > 0 Then
    
    End If
DataGridCntl_OnMouseMove_Exit:
    Exit Sub
DataGridCntl_OnMouseMove_Err:
    MsgBox Err.Description, vbCritical, "DataGridCntl_OnMouseMove"
    Resume DataGridCntl_OnMouseMove_Exit
End Sub

Private Sub DataGridCntl_OnCaptureChanged(hw As Long, wParam As Long)
    ' Close the list box if the scrollbar is not in ACTION &
    ' the window is visible.
On Error GoTo DataGridCntl_OnCaptureChanged_Err
    If Scrolling = 0 Then
        If IsWindowVisible(ByVal pData.hwndGrid) <> 0 Then
            Call HideDropDown
        End If
    End If
DataGridCntl_OnCaptureChanged_Exit:
    Exit Sub
DataGridCntl_OnCaptureChanged_Err:
    MsgBox Err.Description, vbCritical, "DataGridCntl_OnCaptureChanged"
    Resume DataGridCntl_OnCaptureChanged_Exit
End Sub

Private Sub DataGridCntl_OnVScroll(hwnd As Long, nScrollCode As Long, nPos As Long, hWndScrollbar As Long)
On Error GoTo DataGridCntl_OnVScroll_Err
    Select Case nScrollCode
        Case SB_ENDSCROLL
            If (Scrolling <> 0) Then
                ' When the scrollbar goes out of ACTION, we need to
                ' gain mouse capture again. Remember to record that the scrollbar
                ' is NOT in ACTION.
                Scrolling = 0
                Call SetCapture(ByVal hwnd)
            End If
    End Select
DataGridCntl_OnVScroll_Exit:
    Exit Sub
DataGridCntl_OnVScroll_Err:
    MsgBox Err.Description, vbCritical, "DataGridCntl_OnVScroll"
    Resume DataGridCntl_OnVScroll_Exit
End Sub

Private Sub DataGridCntl_OnHScroll(hwnd As Long, nScrollCode As Long, nPos As Long, hWndScrollbar As Long)
On Error GoTo DataGridCntl_OnHScroll_Err
    Select Case nScrollCode
        Case SB_ENDSCROLL
            If Scrolling <> 0 Then
                'When the scrollbar goes out of ACTION, we need to
                'gain mouse capture again. Remember to record that the scrollbar
                'is NOT in ACTION.
                Scrolling = 0
                Call SetCapture(ByVal hwnd)
            End If
    End Select
DataGridCntl_OnHScroll_Exit:
    Exit Sub
DataGridCntl_OnHScroll_Err:
    MsgBox Err.Description, vbCritical, "DataGridCntl_OnHScroll"
    Resume DataGridCntl_OnHScroll_Exit
End Sub

Private Function GetHiLoWord(lParam As Long, TheLowWord As Long, TheHiWord As Long)
On Error GoTo GetHiLoWord_Err
    ' This is the LOWORD of the lParam:
    TheLowWord = lParam And &HFFFF&
    ' LOWORD now equals 65,535 or &HFFFF
    ' This is the HIWORD of the lParam:
    TheHiWord = lParam \ &H10000 And &HFFFF&
    ' HIWORD now equals 30,583 or &H7777
    GetHiLoWord = 1
GetHiLoWord_Exit:
    Exit Function
GetHiLoWord_Err:
    MsgBox Err.Description, vbCritical, "GetHiLoWord"
    Resume GetHiLoWord_Exit
End Function

Private Function MakeLParm(LOWORD As Long, HIWORD As Long) As Long
    Dim lng As Long
On Error GoTo MakeLParm_Err
    'Make sure Upper Bytes are 0000
    lng = LOWORD And &HFFFF&
    'First Make sure HiWord is 0000 at the begining 4 bytes then
    'Miltiply by H10000 to shift the bytes to the left then or them with LOWORD
    ' to form the HI and LOWORD
    lng = ((HIWORD And &HFFFF&) * &H10000) Or lng
    
    MakeLParm = lng
MakeLParm_Exit:
    Exit Function
MakeLParm_Err:
    MsgBox Err.Description, vbCritical, "MakeLParm"
    MakeLParm = 0
    Resume MakeLParm_Exit
End Function

Public Sub HideDropDown()
On Error GoTo HideDropDown_Err
    Call ReleaseCapture     ' Make sure that there is not any other dropdowns
    
    If pData.IsGridHooked Then Call SetWindowLong(ByVal pData.hwndGrid, ByVal GWL_WNDPROC, ByVal pData.PrevGridProc)
    pData.IsGridHooked = False
    
    If Not pData.PictureCtl Is Nothing Then
        pData.PictureCtl.Refresh
    End If
    Call ShowWindow(ByVal pData.hwndGrid, ByVal SW_HIDE)
    DoEvents
    
    If Not pData.DataGridCntl Is Nothing Then
        pData.DataGridCntl.Visible = False
    End If
    Call SetWindowPos(ByVal pData.hwndGrid, ByVal HWND_BOTTOM, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal SWP_HIDEWINDOW)
    Call SetParent(ByVal pData.hwndGrid, ByVal pData.hwndParent)
    If Not pData.DataGridCntl Is Nothing Then
        pData.DataGridCntl.Top = pData.Top
        pData.DataGridCntl.Enabled = False
    End If
    pData.Selection = False
    
    Call ReleaseCapture
    Set pData.DataGridCntl = Nothing
    Set pData.PictureCtl = Nothing
HideDropDown_Exit:
    Exit Sub
HideDropDown_Err:
    MsgBox Err.Description, vbCritical, "HideDropDown"
    Resume HideDropDown_Exit
End Sub

Public Function GetRecNo(Rs As ADODB.Recordset) As Long
On Error GoTo GetRecNo_Err
    GetRecNo = Rs.AbsolutePosition
GetRecNo_Exit:
    Exit Function
GetRecNo_Err:
    MsgBox Err.Description, vbCritical, "GetRecNo"
    Resume GetRecNo_Exit
End Function

Public Sub RemDGBookMark(DG As DataGrid)
On Error GoTo RemDGBookMark_Err
    Do While DG.SelBookmarks.Count > 0
        DG.SelBookmarks.Remove 0
    Loop
RemDGBookMark_Exit:
    Exit Sub
RemDGBookMark_Err:
    Err.Raise Err.Number, "RemDGBookMark", Err.Description
End Sub

Public Function NoOfRecs(Rs As ADODB.Recordset) As Integer
On Error GoTo NoOfRecs_Err
    If Rs Is Nothing Then
        NoOfRecs = 0
    Else
        NoOfRecs = Rs.RecordCount
    End If
NoOfRecs_Exit:
    Exit Function
NoOfRecs_Err:
    MsgBox Err.Description, vbCritical, "NoOfRecs"
    Resume NoOfRecs_Exit
End Function

