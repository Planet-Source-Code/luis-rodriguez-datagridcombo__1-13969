VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.UserControl GridCombo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   ScaleHeight     =   2685
   ScaleWidth      =   3180
   ToolboxBitmap   =   "GridCombo.ctx":0000
   Begin MSDataGridLib.DataGrid DataGridCntl 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PictCtl 
      BackColor       =   &H80000005&
      DrawStyle       =   5  'Transparent
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   2295
      Begin VB.TextBox DGComboTextBox 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "GridCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:

Private Const m_def_DGStyle = 0
Private Const m_def_MousePointer = 0
Private Const m_def_BoundColumn = 0
Private Const m_def_OldValue = ""
Private Const m_def_HeaderHeight = 210
Private Const m_def_gridMaxRows = 6

'Property Variables:
Private m_ChangeSource As Integer
Private m_ColumnHeaders As Boolean
Private m_HideColumns As String
Private m_DGStyle As Integer
Private m_MousePointer As Integer
Private m_BoundColumn As Integer
Private m_OldValue As String
Private m_OldText As String
Private m_AbsolutePost As Integer
Private m_Enabled As Boolean
Private m_GridOriginalProc As Long

Public Enum DGStyles
    dgsFlexible = 0
    dgsNonFlexible = 1
End Enum

'Event Declarations:
Event DropDown()
Event Selection()
Event ClearText()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)



'========================================================================
' Public Functions
'========================================================================
Public Function IsDropDown() As Boolean
On Error GoTo IsDropDown_Err
    IsDropDown = (IsWindowVisible(DataGridCntl.hwnd) <> 0)
IsDropDown_Exit:
    Exit Function
IsDropDown_Err:
    Err.Raise Err.Number, "IsDropDown", Err.Description
End Function

Public Sub ShowDropDown(Show As Integer)
On Error GoTo ShowDropDown_Err
    If Show = 0 And IsWindowVisible(DataGridCntl.hwnd) = 0 Then
        GoTo ShowDropDown_Exit
    ElseIf Show = 1 And (IsWindowVisible(DataGridCntl.hwnd) <> 0 Or m_Enabled = False) Then
        GoTo ShowDropDown_Exit
    End If
    
    Call InitpData
    If Show Then
        RaiseEvent DropDown     ' Raise the DropDown Event
        If NoOfRecs(DataGridCntl.DataSource) = 0 Then
            DataGridCntl.Height = m_def_HeaderHeight
            DataGridCntl.ScrollBars = dbgNone
        Else
            DataGridCntl.ScrollBars = dbgAutomatic
        End If
        Call SetCurrentItem
        If IsWindowVisible(ByVal DGComboTextBox.hwnd) <> 0 Then
            DGComboTextBox.SetFocus
        End If
        Call PrepareDataGrid
        If Not GetWindowLong(DataGridCntl.hwnd, ByVal GWL_WNDPROC) = GetTheAddressOf(AddressOf DataGridWindowProc) Then
            pData.PrevGridProc = SetWindowLong(ByVal DataGridCntl.hwnd, ByVal GWL_WNDPROC, AddressOf DataGridWindowProc)
            m_GridOriginalProc = pData.PrevGridProc
            pData.IsGridHooked = True
            
            Call SendMessage(DataGridCntl.hwnd, ByVal WMDG_NOTIFY, 0, "")
        End If
        Call SetCapture(ByVal DataGridCntl.hwnd)
        Call ShowWindow(ByVal DataGridCntl.hwnd, ByVal SW_SHOW)
    Else
        Call HideDropDown
    End If
ShowDropDown_Exit:
    Exit Sub
ShowDropDown_Err:
    Call ReleaseCapture
    MsgBox Err.Description, vbCritical, "ShowDropDown"
    Resume ShowDropDown_Exit
End Sub

Public Sub SetListItem(Item As String, Exact As Boolean)
    Dim m_Rs As ADODB.Recordset
On Error GoTo SetListItem_Err
    If NoOfRecs(DataGridCntl.DataSource) > 0 Then
        Set m_Rs = DataGridCntl.DataSource
        m_Rs.MoveFirst
        ' This needs to be modified for other type of fields
        If m_Rs(m_BoundColumn).Type = adInteger Then
            m_Rs.Find m_Rs(m_BoundColumn).Name & " = " & Val(Item) & ""
        Else
            If Exact Then
                m_Rs.Find m_Rs(m_BoundColumn).Name & " = '" & Item & "'"
            Else
                m_Rs.Find "[" & m_Rs(m_BoundColumn).Name & "] LIKE '" & IIf(Item = "", " ", Item) & "%'"
            End If
        End If
        If Not m_Rs.EOF Then
            If Not (Exact And m_Rs(m_BoundColumn) <> Item) Then
                m_AbsolutePost = m_Rs.AbsolutePosition
                DGComboTextBox.Text = m_Rs.Fields(m_BoundColumn)
                If DGComboTextBox.Text <> m_OldText Then
                    m_OldText = DGComboTextBox.Text
                    RaiseEvent Selection
                End If
                If m_AbsolutePost <> -1 Then m_Rs.AbsolutePosition = m_AbsolutePost
            End If
        Else
            Call RemDGBookMark(DataGridCntl)
        End If
    End If
SetListItem_Exit:
    Set m_Rs = Nothing
    Exit Sub
SetListItem_Err:
    MsgBox Err.Description, vbCritical, "SetListItem"
    Resume SetListItem_Exit
End Sub

Public Sub Clear()
On Error GoTo Clear_Err
    Set DataGridCntl.DataSource = Nothing
    DGComboTextBox.Text = ""
    m_AbsolutePost = -1
Clear_Exit:
    Exit Sub
Clear_Err:
    MsgBox Err.Description, vbCritical, "Clear"
    Resume Clear_Exit
End Sub
'========================================================================
' Public Functions
'========================================================================








'========================================================================
' DataGridCntl Events
'========================================================================
Private Sub DataGridCntl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo DataGridCntl_MouseMove_Err
    If Not DGComboTextBox.Locked Then
        If NoOfRecs(DataGridCntl.DataSource) > 0 Then
            '-- Make sure that the Position is within Grid Boundaries --
            If DataGridCntl.RowContaining(y) >= 0 And DataGridCntl.ColContaining(x) >= 0 And x > 0 And y > 0 And y <= DataGridCntl.Height Then
                If y > (DataGridCntl.Top + DataGridCntl.Height) Or x > (DataGridCntl.Left + DataGridCntl.Width) Then
                    GoTo DataGridCntl_MouseMove_Exit
                End If
                If DataGridCntl.SelBookmarks.Count > 0 Then
                    ' -- See If Highlighted bookmark is equal to current --
                    If DataGridCntl.SelBookmarks(0) = DataGridCntl.RowBookmark(DataGridCntl.RowContaining(y)) Then
                        GoTo DataGridCntl_MouseMove_Exit
                    End If
                End If
                Call RemDGBookMark(DataGridCntl)
                DataGridCntl.SelBookmarks.Add DataGridCntl.RowBookmark(DataGridCntl.RowContaining(y))
            End If
        End If
    End If
DataGridCntl_MouseMove_Exit:
    Exit Sub
DataGridCntl_MouseMove_Err:
    MsgBox Err.Description, vbCritical, "DataGrid_MouseMove"
    Resume DataGridCntl_MouseMove_Exit
End Sub
'========================================================================
' DataGridCntl Events
'========================================================================










'========================================================================
' DGComboTextBox Events
'========================================================================
Private Sub DGComboTextBox_GotFocus()
On Error GoTo DGComboTextBox_GotFocus_Err
    DGComboTextBox.SelStart = 0
    DGComboTextBox.SelLength = Len(DGComboTextBox.Text)
DGComboTextBox_GotFocus_Exit:
    Exit Sub
DGComboTextBox_GotFocus_Err:
    MsgBox Err.Description, vbCritical, "DGComboTextBox_GotFocus"
    Resume DGComboTextBox_GotFocus_Exit
End Sub

Private Sub DGComboTextBox_LostFocus()
On Error GoTo DGComboTextBox_LostFocus_Err
    If GetFocus() <> UserControl.DGComboTextBox.hwnd And GetFocus() <> UserControl.DataGridCntl.hwnd Then
        '-- If Control Looses focus Hide the Drop Down --
        Call Me.ShowDropDown(0)
    End If
DGComboTextBox_LostFocus_Exit:
    Exit Sub
DGComboTextBox_LostFocus_Err:
    MsgBox Err.Description, vbCritical, "DGComboTextBox_LostFocus"
    Resume DGComboTextBox_LostFocus_Exit
End Sub

Private Sub DGComboTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
On Error GoTo DGComboTextBox_KeyDown_Err
    If m_DGStyle <> 0 And DGComboTextBox.Locked = False Then
        If Not KeyCode = vbKeyBack Then
            Call TextKeyDown(KeyCode)
        Else
            DGComboTextBox.SelStart = IIf(DGComboTextBox.SelStart = 0, 0, DGComboTextBox.SelStart - 1)
            KeyCode = 0
        End If
    End If
DGComboTextBox_KeyDown_Exit:
    Exit Sub
DGComboTextBox_KeyDown_Err:
    MsgBox Err.Description, vbCritical, "DGComboTextBox_KeyDown"
    Resume DGComboTextBox_KeyDown_Exit
End Sub

Private Sub DGComboTextBox_KeyPress(KeyAscii As Integer)
    Dim iLen As Integer, iIndex As Integer
    Dim m_Rs As ADODB.Recordset
    Dim StrText As String
On Error GoTo DGComboTextBox_KeyPress_Err
    RaiseEvent KeyPress(KeyAscii)
    If m_DGStyle = 0 Or DGComboTextBox.Locked Then
        GoTo DGComboTextBox_KeyPress_Exit
    End If
    If DGComboTextBox.SelLength = 0 Then
        iIndex = DGComboTextBox.SelStart
        StrText = Left(DGComboTextBox.Text, DGComboTextBox.SelStart) & UCase(Chr(KeyAscii)) & Mid(DGComboTextBox.Text, DGComboTextBox.SelStart + 1)
        iLen = Len(DGComboTextBox.Text) - DGComboTextBox.SelStart
    Else
        iIndex = DGComboTextBox.SelStart
        If iIndex > 0 Then
            StrText = Left(DGComboTextBox.Text, DGComboTextBox.SelStart) & UCase(Chr(KeyAscii)) & Mid(DGComboTextBox.Text, DGComboTextBox.SelStart + DGComboTextBox.SelLength + 1)
        Else
            StrText = UCase(Chr(KeyAscii)) & Mid(DGComboTextBox.Text, DGComboTextBox.SelStart + DGComboTextBox.SelLength + 1)
        End If
        iLen = Len(DGComboTextBox.Text) - DGComboTextBox.SelLength
    End If
    If NoOfRecs(DataGridCntl.DataSource) > 0 Then
        Set m_Rs = DataGridCntl.DataSource
        m_Rs.MoveFirst
        m_Rs.Find m_Rs(m_BoundColumn).Name & " LIKE '" & StrText & "%'"
        If Not m_Rs.EOF Then
            m_AbsolutePost = m_Rs.AbsolutePosition
            Call RemDGBookMark(DataGridCntl)
            DataGridCntl.SelBookmarks.Add DataGridCntl.Bookmark
            'm_ChangeSource = 0
            pData.Selection = True
            DGComboTextBox.Text = m_Rs.Fields(m_BoundColumn)
            DGComboTextBox.SelStart = iIndex + 1
            DGComboTextBox.SelLength = Len(DGComboTextBox.Text) - (iLen + 1)
            pData.Selection = False
            'm_ChangeSource = 1
            If DGComboTextBox.Text <> m_OldText Then
                m_OldText = DGComboTextBox.Text
                RaiseEvent Selection
            End If
        End If
        If m_AbsolutePost <> -1 Then m_Rs.AbsolutePosition = m_AbsolutePost
    End If
DGComboTextBox_KeyPress_Exit:
    KeyAscii = 0
    Set m_Rs = Nothing
    Exit Sub
DGComboTextBox_KeyPress_Err:
    MsgBox Err.Description, vbCritical, "DGComboTextBox_KeyPress"
    Resume DGComboTextBox_KeyPress_Exit
End Sub

Private Sub DGComboTextBox_Change()
On Error GoTo DGComboTextBox_Change_Err
    If m_DGStyle = 0 And pData.Selection = False Then
        GoTo DGComboTextBox_Change_Exit
    End If
    If m_ChangeSource = 1 Then      ' m_ChangeSource?
        Call SetRecordField
        GoTo DGComboTextBox_Change_Exit
    End If
    If DGComboTextBox.DataChanged = False Then
        GoTo DGComboTextBox_Change_Exit
    End If
    If Trim(m_OldText) <> Trim(DGComboTextBox.Text) Then
        '-- Did the selection changed ? --
        If pData.Selection Then
            m_OldText = DGComboTextBox.Text
            pData.Selection = False
            m_AbsolutePost = GetRecNo(DataGridCntl.DataSource)
            Call SetRecordField
            DGComboTextBox.SelStart = 0
            DGComboTextBox.SelLength = Len(DGComboTextBox.Text)
            RaiseEvent Selection
        Else
            If Trim(DGComboTextBox.Text) <> "" Then
                If Not FindListItem(DGComboTextBox.Text) Then
                    DGComboTextBox.Text = ""
                Else
                    m_OldText = DGComboTextBox.Text
                    m_AbsolutePost = GetRecNo(DataGridCntl.DataSource)
                End If
                Call SetRecordField
            Else
                m_OldText = ""
                m_AbsolutePost = -1
                Call SetRecordField
                RaiseEvent ClearText
            End If
        End If
    End If
DGComboTextBox_Change_Exit:
    Exit Sub
DGComboTextBox_Change_Err:
    Err.Raise 380
End Sub
'========================================================================
' DGComboTextBox Events
'========================================================================









'========================================================================
' Picture Frame Events
'========================================================================
Private Sub PictCtl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo PictCtl_MouseDown_Err
    If m_Enabled And (x > UserControl.DGComboTextBox.Width) Then
        Call DrawFrameControl(ByVal UserControl.PictCtl.hdc, MakeRect((DGComboTextBox.Width / Screen.TwipsPerPixelX) + 1, 0, GetSystemMetrics(ByVal SM_CXHTHUMB), (DGComboTextBox.Height / Screen.TwipsPerPixelY) + 2), DFC_SCROLL, DFCS_SCROLLDOWN Or DFCS_PUSHED Or DFCS_FLAT)
        DGComboTextBox.SetFocus     '{ Set the Focus to the TextBox and highlight the text }
        DGComboTextBox.SelLength = Len(DGComboTextBox.Text)
        DoEvents
        Call Me.ShowDropDown(1)     '{ Show the DropDown DataGrid }
    End If
PictCtl_MouseDown_Exit:
    Exit Sub
PictCtl_MouseDown_Err:
    MsgBox Err.Description, vbCritical, "PictCtl_MouseDown"
    Resume PictCtl_MouseDown_Exit
End Sub

Private Sub PictCtl_Paint()
On Error GoTo PictCtl_Paint_Err
    ' Show the DropDown button Disabled only at run time
    If Not m_Enabled And UserControl.Ambient.UserMode Then
        Call DrawFrameControl(ByVal PictCtl.hdc, MakeRect((DGComboTextBox.Width / Screen.TwipsPerPixelX) + 1, 0, GetSystemMetrics(ByVal SM_CXHTHUMB), (DGComboTextBox.Height / Screen.TwipsPerPixelY) + 2), DFC_SCROLL, DFCS_SCROLLDOWN Or DFCS_INACTIVE)
    Else
        Call DrawFrameControl(ByVal PictCtl.hdc, MakeRect((DGComboTextBox.Width / Screen.TwipsPerPixelX) + 1, 0, GetSystemMetrics(ByVal SM_CXHTHUMB), (DGComboTextBox.Height / Screen.TwipsPerPixelY) + 2), DFC_SCROLL, DFCS_SCROLLDOWN)
    End If
PictCtl_Paint_Exit:
    Exit Sub
PictCtl_Paint_Err:
    MsgBox Err.Description, vbCritical, "PictCtl_Paint"
    Resume PictCtl_Paint_Exit
End Sub
'========================================================================
' Picture Frame Events
'========================================================================






'========================================================================
' UserControl Events
'========================================================================
Private Sub UserControl_Initialize()
On Error GoTo UserControl_Initialize_Err
    DataGridCntl.Visible = False
    m_ChangeSource = 0
    m_Enabled = True
    Call SetWindowLong(ByVal DataGridCntl.hwnd, ByVal GWL_EXSTYLE, ByVal WS_EX_TOOLWINDOW)
    Call SetWindowLong(ByVal DataGridCntl.hwnd, ByVal GWL_STYLE, ByVal (WS_CHILD Or WS_BORDER))
    m_AbsolutePost = -1
UserControl_Initialize_Exit:
    Exit Sub
UserControl_Initialize_Err:
    MsgBox Err.Description, vbCritical, "UserControl_Initialize"
    Resume UserControl_Initialize_Exit
End Sub

Private Sub UserControl_Terminate()
On Error GoTo UserControl_Terminate_Err
    If GetWindowLong(DataGridCntl.hwnd, ByVal GWL_WNDPROC) = GetTheAddressOf(AddressOf DataGridWindowProc) Then
        Call SetWindowLong(ByVal DataGridCntl.hwnd, ByVal GWL_WNDPROC, ByVal m_GridOriginalProc)
    End If
    Call SetParent(ByVal DataGridCntl.hwnd, UserControl.hwnd)
    Set DataGridCntl.DataSource = Nothing
UserControl_Terminate_Exit:
    Exit Sub
UserControl_Terminate_Err:
    MsgBox Err.Description, vbCritical, "UserControl_Terminate"
    Resume UserControl_Terminate_Exit
End Sub

Private Sub UserControl_Resize()
On Error GoTo UserControl_Resize_Err
    If IsWindowVisible(DataGridCntl.hwnd) = 0 Then
        If UserControl.Height < 315 Then
            UserControl.Height = 315
        End If
        
        PictCtl.Top = 0
        PictCtl.Left = 0
        PictCtl.Width = UserControl.Width
        PictCtl.Height = UserControl.Height

        DGComboTextBox.Top = 1 * Screen.TwipsPerPixelY
        DGComboTextBox.Left = 1 * Screen.TwipsPerPixelX
        DGComboTextBox.Width = UserControl.Width - ((GetSystemMetrics(ByVal SM_CXHTHUMB) + 5) * Screen.TwipsPerPixelX)
        DGComboTextBox.Height = UserControl.Height - (6 * Screen.TwipsPerPixelY)
    End If
UserControl_Resize_Exit:
    Exit Sub
UserControl_Resize_Err:
    MsgBox Err.Description, vbCritical, "UserControl_Resize"
    Resume UserControl_Resize_Exit
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo UserControl_ReadProperties_Err
    DGComboTextBox.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    DGComboTextBox.Locked = PropBag.ReadProperty("Locked", False)
    DGComboTextBox.Text = PropBag.ReadProperty("Text", "")
    m_OldValue = PropBag.ReadProperty("OldValue", m_def_OldValue)
    m_BoundColumn = PropBag.ReadProperty("BoundColumn", m_def_BoundColumn)
    m_ColumnHeaders = PropBag.ReadProperty("ColumnHeaders", True)
    DGComboTextBox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    m_HideColumns = PropBag.ReadProperty("HideColumns", "")
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    DGComboTextBox.DataMember = PropBag.ReadProperty("DataMember", "")
    DGComboTextBox.DataField = PropBag.ReadProperty("DataField", "")
    m_DGStyle = PropBag.ReadProperty("DGStyle", m_def_DGStyle)
UserControl_ReadProperties_Exit:
    Exit Sub
UserControl_ReadProperties_Err:
    MsgBox Err.Description, vbCritical, "UserControl_ReadProperties"
    Resume UserControl_ReadProperties_Exit
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo UserControl_WriteProperties_Err

    Call PropBag.WriteProperty("Alignment", DGComboTextBox.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Locked", DGComboTextBox.Locked, False)
    Call PropBag.WriteProperty("Text", DGComboTextBox.Text, "")
    Call PropBag.WriteProperty("OldValue", m_OldValue, m_def_OldValue)
    Call PropBag.WriteProperty("BoundColumn", m_BoundColumn, m_def_BoundColumn)
    Call PropBag.WriteProperty("ColumnHeaders", m_ColumnHeaders, True)
    Call PropBag.WriteProperty("ForeColor", DGComboTextBox.ForeColor, &H80000008)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("DataSource", DGComboTextBox.DataSource, Nothing)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("DataMember", DGComboTextBox.DataMember, "")
    Call PropBag.WriteProperty("DataField", DGComboTextBox.DataField, "")
    Call PropBag.WriteProperty("DGStyle", m_DGStyle, m_def_DGStyle)
    Call PropBag.WriteProperty("HideColumns", m_HideColumns, "")
UserControl_WriteProperties_Exit:
    Exit Sub
UserControl_WriteProperties_Err:
    MsgBox Err.Description, vbCritical, "UserControl_WriteProperties"
    Resume UserControl_WriteProperties_Exit
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error GoTo UserControl_InitProperties_Err
    m_OldValue = m_def_OldValue
    m_OldValue = m_def_OldValue
    m_BoundColumn = m_def_BoundColumn
    m_ColumnHeaders = True
    m_MousePointer = m_def_MousePointer
    m_DGStyle = m_def_DGStyle
UserControl_InitProperties_Exit:
    Exit Sub
UserControl_InitProperties_Err:
    MsgBox Err.Description, vbCritical, "UserControl_InitProperties"
    Resume UserControl_InitProperties_Exit
End Sub
'========================================================================
' UserControl Events
'========================================================================






'========================================================================
' UserControl Properties
'========================================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    UserControl.DGComboTextBox.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get HideColumns() As String
    HideColumns = m_HideColumns
End Property

Public Property Let HideColumns(ByVal New_HideColumns As String)
    m_HideColumns = New_HideColumns
    PropertyChanged "HideColumns"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get ListCount() As Integer
    ListCount = NoOfRecs(DataGridCntl.DataSource)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DGComboTextBox,DGComboTextBox,-1,Locked
Public Property Get Locked() As Boolean
    Locked = DGComboTextBox.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    DGComboTextBox.Locked = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DGComboTextBox,DGComboTextBox,-1,Text
Public Property Get Text() As String
    Text = DGComboTextBox.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_ChangeSource = 1
    DGComboTextBox.Text = New_Text
    m_OldText = New_Text
    m_ChangeSource = 0
    PropertyChanged "Text"
End Property

Public Property Set RowSource(New_RowSource As ADODB.Recordset)
    Dim m_Rs As ADODB.Recordset
    If Not New_RowSource Is Nothing Then
        'Set m_Rs = New_RowSource.Clone(adLockReadOnly)
        Set m_Rs = New_RowSource
        'm_Rs.Filter = New_RowSource.Filter
        'm_Rs.Sort = New_RowSource.Sort
        DataGridCntl.ClearFields
        DataGridCntl.Height = ((m_def_gridMaxRows * (DataGridCntl.RowHeight + 15)) + m_def_HeaderHeight)
        Set DataGridCntl.DataSource = Nothing
        Set DataGridCntl.DataSource = m_Rs
        If m_Rs.RecordCount > 0 Then m_Rs.MoveFirst
        Call ResizeGrid
        Set m_Rs = Nothing
    Else
        Set DataGridCntl.DataSource = Nothing
    End If
End Property

Public Property Get RowSource() As ADODB.Recordset
    Set RowSource = DataGridCntl.DataSource
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,""
Public Property Get OldValue() As String
    OldValue = m_OldValue
End Property

Public Property Let OldValue(ByVal New_OldValue As String)
    m_OldValue = New_OldValue
    PropertyChanged "OldValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get BoundColumn() As Integer
    BoundColumn = m_BoundColumn
End Property

Public Property Let BoundColumn(ByVal New_BoundColumn As Integer)
    m_BoundColumn = New_BoundColumn
    PropertyChanged "BoundColumn"
End Property

Public Property Get ColumnHeaders() As Boolean
    ColumnHeaders = m_ColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal New_ColumnHeaders As Boolean)
    m_ColumnHeaders = New_ColumnHeaders
    DataGridCntl.ColumnHeaders = m_ColumnHeaders
    PropertyChanged "ColumnHeaders"
End Property

Public Property Get Columns() As Columns
    Set Columns = DataGridCntl.Columns
    Call ResizeGrid
End Property

Public Property Let Columns(ByVal New_Columns As Columns)
    Dim i As Integer
    For i = 0 To DataGridCntl.Columns.Count - 1
        DataGridCntl.Columns(i) = New_Columns(i)
    Next i
    Call ResizeGrid
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DGComboTextBox,DGComboTextBox,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = DGComboTextBox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    DGComboTextBox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MousePointer() As Integer
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get DataField() As String
    DataField = CStr(DGComboTextBox.DataField)
End Property

Public Property Let DataField(ByVal New_DataField As String)
    DGComboTextBox.DataField = New_DataField
    PropertyChanged "DataField"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DGComboTextBox,DGComboTextBox,-1,DataSource
Public Property Get DataSource() As DataSource
    Set DataSource = DGComboTextBox.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As DataSource)
    Set DGComboTextBox.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DGComboTextBox,DGComboTextBox,-1,DataMember
Public Property Get DataMember() As String
    DataMember = DGComboTextBox.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    DGComboTextBox.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

Public Property Get DGStyle() As DGStyles
    DGStyle = m_DGStyle
End Property

Public Property Let DGStyle(ByVal New_DGStyle As DGStyles)
    If Ambient.UserMode Then Err.Raise 382
    m_DGStyle = New_DGStyle
    PropertyChanged "DGStyle"
End Property

Public Property Get ColumnValue(Index As Integer) As String
    If NoOfRecs(DataGridCntl.DataSource) <= 0 Then
        ColumnValue = ""
    ElseIf DGComboTextBox = "" Then
        ColumnValue = ""
    Else
        ColumnValue = DataGridCntl.Columns(Index).Value
    End If
End Property
'========================================================================
' UserControl Properties
'========================================================================












'========================================================================
' Private Functions
'========================================================================
Private Sub InitpData()
On Error GoTo InitpData_Err
    Call ReleaseCapture
    Set pData.DataGridCntl = DataGridCntl
    Set pData.PictureCtl = PictCtl
    pData.IsGridHooked = False
    pData.AbsolutePos = m_AbsolutePost
    pData.hwndEditBox = UserControl.DGComboTextBox.hwnd
    pData.hwndParent = UserControl.hwnd
    pData.hwndGrid = UserControl.DataGridCntl.hwnd
    pData.PrevGridProc = 0
    pData.Selection = False
    pData.Top = DGComboTextBox.Top + DGComboTextBox.Height + 100
    pData.BColumn = m_BoundColumn
InitpData_Exit:
    Exit Sub
InitpData_Err:
    MsgBox Err.Description, vbCritical, "InitpData"
    Resume InitpData_Exit
End Sub

Private Sub ResizeGrid()
    Dim LngWidth As Long, i As Integer
    Dim ScrollBarWidth As Long
On Error GoTo ResizeGrid_Err
    LngWidth = 0
    Call ResizeColumns
    For i = 0 To DataGridCntl.Columns.Count - 1
        LngWidth = LngWidth + DataGridCntl.Columns(i).Width
    Next i
    ScrollBarWidth = 0
    DataGridCntl.Height = ((m_def_gridMaxRows * (DataGridCntl.RowHeight + 15)) + m_def_HeaderHeight)
    If NoOfRecs(DataGridCntl.DataSource) > m_def_gridMaxRows Then
        ScrollBarWidth = Screen.TwipsPerPixelX * (GetSystemMetrics(ByVal SM_CXHTHUMB))
    Else
        DataGridCntl.Height = (NoOfRecs(DataGridCntl.DataSource) * (DataGridCntl.RowHeight + 15)) + m_def_HeaderHeight
    End If
    DataGridCntl.ColumnHeaders = m_ColumnHeaders
    If DataGridCntl.ColumnHeaders = False Then
        DataGridCntl.Height = DataGridCntl.Height - m_def_HeaderHeight
    End If
    DataGridCntl.Width = LngWidth + ScrollBarWidth
    If DataGridCntl.Width < UserControl.Width Then
        DataGridCntl.Columns(DataGridCntl.Columns.Count - 1).Width = DataGridCntl.Columns(DataGridCntl.Columns.Count - 1).Width + (UserControl.Width - DataGridCntl.Width)
        DataGridCntl.Width = UserControl.Width
    End If
ResizeGrid_Exit:
    Exit Sub
ResizeGrid_Err:
    MsgBox Err.Description, vbCritical, "ResizeGrid"
    Resume ResizeGrid_Exit
End Sub

Private Sub ResizeColumns()
    Dim i As Integer, ArrFieldsToHide As Variant
    Dim CharacterWidth As Integer
    Dim m_Rs As ADODB.Recordset
On Error GoTo ResizeColumns_Err
    Set m_Rs = DataGridCntl.DataSource
    ArrFieldsToHide = Split(m_HideColumns, ",")
    CharacterWidth = UserControl.TextWidth("A")
    
    If Not m_Rs Is Nothing Then
        For i = 0 To m_Rs.Fields.Count - 1
            If FoundInArray(ArrFieldsToHide, i) Then
                DataGridCntl.Columns(i).Width = 0
            Else
                If m_Rs.Fields(i).Type = adChar Then
                    If DataGridCntl.Columns(i).Width < (m_Rs.Fields(i).DefinedSize * CharacterWidth) Then
                        DataGridCntl.Columns(i).Width = m_Rs.Fields(i).DefinedSize * CharacterWidth
                    End If
                End If
            End If
        Next i
    End If
ResizeColumns_Exit:
    Set m_Rs = Nothing
    Exit Sub
ResizeColumns_Err:
    MsgBox Err.Description, vbCritical, "ResizeColumns"
    Resume ResizeColumns_Exit
End Sub

Private Function FoundInArray(Arr As Variant, Index As Integer) As Boolean
    Dim i As Integer
On Error GoTo FoundInArray_Err
    FoundInArray = False
    For i = 0 To UBound(Arr)
        If Arr(i) = Index Then
            FoundInArray = True
            Exit For
        End If
    Next i
FoundInArray_Exit:
    Exit Function
FoundInArray_Err:
    Resume FoundInArray_Exit
End Function

Private Function GetRECTForGrid() As RECT
On Error GoTo GetRECTForGrid_Err
    Dim nScreenH As Long, nScreenW As Long
    Dim nBottom As Long, nRight As Long
    Dim RcGrid As RECT, RcScrn As RECT
    Dim Rc As RECT, RcCntl As RECT
    
    '{ Get the UserControl coordinates and the DataGrid Coordinates }
    Call GetWindowRect(ByVal UserControl.hwnd, RcCntl)
    '{ Get the Screen Coordinates }
    Call SystemParametersInfo(ByVal SPI_GETWORKAREA, ByVal 0, RcScrn, ByVal 0)
    nScreenH = RcScrn.Bottom - RcScrn.Top   '{ Screen Height    }
    Rc.Bottom = (DataGridCntl.Height / Screen.TwipsPerPixelY)
    Rc.Top = RcCntl.Bottom + 1
    nBottom = Rc.Top + Rc.Bottom
    '{ Is Grid going to go out of the screen ? }
    If (Rc.Top + Rc.Bottom) > RcScrn.Bottom Then
        '{ Yes, then Move Grid to the top of the Control }
        nBottom = RcCntl.Top - 1
        Rc.Top = nBottom - Rc.Bottom
        If Rc.Top < 0 Then
            Rc.Top = 0
        End If
    End If
    Rc.Right = (DataGridCntl.Width / Screen.TwipsPerPixelX)
    Rc.Left = RcCntl.Left
    nRight = Rc.Left + Rc.Right
    '{ Is Data Grid Going to be out of Screen ? }
    Do While ((Rc.Left + Rc.Right) > RcScrn.Right) And Rc.Left >= 0
        Rc.Left = Rc.Left - 1
    Loop
    GetRECTForGrid = Rc
GetRECTForGrid_Exit:
    Exit Function
GetRECTForGrid_Err:
    MsgBox Err.Description, vbCritical, "GetRECTForGrid"
    Resume GetRECTForGrid_Exit
End Function

Private Function DSRecordSet(Rs As ADODB.Recordset) As ADODB.Recordset
On Error GoTo DSRecordSet_Err
    Set DSRecordSet = Rs
DSRecordSet_Exit:
    Exit Function
DSRecordSet_Err:
    Resume DSRecordSet_Exit
End Function

Private Sub SetRecordField()
On Error GoTo SetRecordField_Err
    If Not DGComboTextBox.DataSource Is Nothing Then
        DSRecordSet(DGComboTextBox.DataSource).Fields(DGComboTextBox.DataField).Value = DGComboTextBox.Text
    End If
SetRecordField_Exit:
    Exit Sub
SetRecordField_Err:
    Resume SetRecordField_Exit
End Sub

Private Sub SetCurrentItem()
On Error GoTo SetCurrentItem_Err
    If DGComboTextBox.Text <> "" Then
        If FindListItem(DGComboTextBox.Text) Then
            Call RemDGBookMark(DataGridCntl)
            DataGridCntl.SelBookmarks.Add DataGridCntl.Bookmark
            m_AbsolutePost = GetRecNo(DataGridCntl.DataSource)
        End If
    End If
SetCurrentItem_Exit:
    Exit Sub
SetCurrentItem_Err:
    MsgBox Err.Description, vbCritical, "SetCurrentItem"
    Resume SetCurrentItem_Exit
End Sub

Private Sub PrepareDataGrid()
On Error GoTo PrepareDataGrid_Err
    Dim Rc As RECT
    
    With DataGridCntl
        .Enabled = True
        .ClearSelCols
        .ColumnHeaders = (NoOfRecs(.DataSource) <> 0)
    End With
    Rc = GetRECTForGrid()
    Call MoveWindow(ByVal DataGridCntl.hwnd, ByVal Rc.Left, ByVal Rc.Top, ByVal Rc.Right, ByVal Rc.Bottom, ByVal 1)
    DataGridCntl.Visible = True
    Call ShowWindow(ByVal DataGridCntl.hwnd, ByVal SW_HIDE)
    Call SetParent(ByVal DataGridCntl.hwnd, ByVal GetDesktopWindow)
    Call SetWindowPos(ByVal DataGridCntl.hwnd, ByVal HWND_TOPMOST, ByVal Rc.Left, ByVal Rc.Top, ByVal Rc.Right, ByVal Rc.Bottom, ByVal (SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE))
PrepareDataGrid_Exit:
    Exit Sub
PrepareDataGrid_Err:
    MsgBox Err.Description, vbCritical, "PrepareDataGrid"
    Resume PrepareDataGrid_Exit
End Sub

Private Function FindListItem(Item As String) As Boolean
    Dim m_Rs As ADODB.Recordset
On Error GoTo FindListItem_Err
    If NoOfRecs(DataGridCntl.DataSource) > 0 Then
        Set m_Rs = DataGridCntl.DataSource
        'If m_Rs(m_BoundColumn) <> CStr(Item) Then
            m_Rs.MoveFirst
            m_Rs.Find "[" & m_Rs(m_BoundColumn).Name & "] = '" & CStr(Item) & "'"
        'End If
        FindListItem = Not m_Rs.EOF
    Else
        FindListItem = False
    End If
FindListItem_Exit:
    Set m_Rs = Nothing
    Exit Function
FindListItem_Err:
    FindListItem = False
    MsgBox Err.Description, vbCritical, "FindListItem"
    Resume FindListItem_Exit
End Function

Private Function MakeRect(l As Long, t As Long, w As Long, h As Long) As RECT
On Error GoTo MakeRect_Err
   With MakeRect
      .Left = l
      .Top = t
      .Right = l + w
      .Bottom = t + h
   End With
MakeRect_Exit:
    Exit Function
MakeRect_Err:
    MsgBox Err.Description, vbCritical, "MakeRect"
    Resume MakeRect_Exit
End Function

Private Sub TextKeyDown(KeyCode As Integer)
    Dim FirstTime As Boolean
    Dim m_Rs As ADODB.Recordset
On Error GoTo TextKeyDown_Err
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        If IsWindowVisible(DataGridCntl.hwnd) <> 0 Then
            Call Me.ShowDropDown(0)
        End If
        KeyCode = 0
        GoTo TextKeyDown_Exit
    End If
    If NoOfRecs(DataGridCntl.DataSource) = 0 Then
        GoTo TextKeyDown_Exit
    End If
    'Debug.Print m_AbsolutePost
    'GoTo TextKeyDown_Exit

    Set m_Rs = DataGridCntl.DataSource
    FirstTime = False
    If IsWindowVisible(DataGridCntl.hwnd) <> 0 Then
        If DataGridCntl.SelBookmarks.Count > 0 Then
            m_Rs.Bookmark = DataGridCntl.SelBookmarks(0)
        Else
            m_Rs.MoveFirst
            FirstTime = True
        End If
    Else
        If m_AbsolutePost <> -1 Then
            m_Rs.AbsolutePosition = m_AbsolutePost
        Else
            m_Rs.MoveFirst
            FirstTime = True
        End If
    End If
    
    Select Case KeyCode
        Case vbKeyUp
            If Not m_Rs.BOF Then
                m_Rs.MovePrevious
                If FirstTime Then m_Rs.MoveFirst
                If Not m_Rs.BOF Then
                    pData.Selection = True
                    DGComboTextBox.Text = CStr(m_Rs.Fields(m_BoundColumn))
                    pData.Selection = False
                    Call RemDGBookMark(DataGridCntl)
                    DataGridCntl.SelBookmarks.Add m_Rs.Bookmark
                End If
            End If
            KeyCode = 0
        Case vbKeyDown
            If Not m_Rs.EOF Then
                m_Rs.MoveNext
                If FirstTime Then m_Rs.MoveFirst
                If Not m_Rs.EOF Then
                    Call RemDGBookMark(DataGridCntl)
                    DataGridCntl.SelBookmarks.Add m_Rs.Bookmark
                    pData.Selection = True
                    DGComboTextBox.Text = CStr(m_Rs.Fields(m_BoundColumn))
                    pData.Selection = False
                End If
            End If
            KeyCode = 0
        Case vbKeyPageUp
            If Not m_Rs.BOF Then
                If IsWindowVisible(DataGridCntl.hwnd) <> 0 Then
                    Call LockWindowUpdate(ByVal DataGridCntl.hwnd)
                End If
                m_Rs.Move -(DataGridCntl.VisibleRows - 1)
                If m_Rs.BOF Then
                    m_Rs.MoveFirst
                End If
                pData.Selection = True
                DGComboTextBox.Text = CStr(m_Rs.Fields(m_BoundColumn))
                pData.Selection = False
                m_AbsolutePost = m_Rs.AbsolutePosition
                Call RemDGBookMark(DataGridCntl)
                If IsWindowVisible(DataGridCntl.hwnd) <> 0 Then
                    Call LockWindowUpdate(ByVal 0)
                End If
                DataGridCntl.SelBookmarks.Add m_Rs.Bookmark
            End If
        Case vbKeyPageDown
            If Not m_Rs.EOF Then
                If IsWindowVisible(DataGridCntl.hwnd) <> 0 Then
                    Call LockWindowUpdate(ByVal DataGridCntl.hwnd)
                End If
                m_Rs.Move (DataGridCntl.VisibleRows - 1)
                If m_Rs.EOF Then
                    m_Rs.MoveLast
                End If
                pData.Selection = True
                DGComboTextBox.Text = CStr(m_Rs.Fields(m_BoundColumn))
                pData.Selection = False
                m_AbsolutePost = m_Rs.AbsolutePosition
                Call RemDGBookMark(DataGridCntl)
                If IsWindowVisible(DataGridCntl.hwnd) <> 0 Then
                    Call LockWindowUpdate(ByVal 0)
                End If
                DataGridCntl.SelBookmarks.Add m_Rs.Bookmark
            End If
            KeyCode = 0
    End Select
    If m_AbsolutePost <> -1 Then
        m_Rs.AbsolutePosition = m_AbsolutePost
    End If
TextKeyDown_Exit:
    Call LockWindowUpdate(ByVal 0)
    Set m_Rs = Nothing
    Exit Sub
TextKeyDown_Err:
    MsgBox Err.Description, vbCritical, "TextKeyDown"
    Resume TextKeyDown_Exit
End Sub

Private Function GetTheAddressOf(lng As Long) As Long
On Error GoTo GetTheAddressOf_Err
    GetTheAddressOf = lng
GetTheAddressOf_Exit:
    Exit Function
GetTheAddressOf_Err:
    MsgBox Err.Description, vbCritical, "GetTheAddressOf"
    Resume GetTheAddressOf_Exit
End Function
'========================================================================
' Private Functions
'========================================================================
