Attribute VB_Name = "mod_subclass"
Option Explicit
    
'GDI32 Declerations
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function TextOutBStr Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpString As Any, ByVal nCount As Long) As Long

'Kernel32 Declerations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

'User32 Declerations
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Rect Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Draw Item Structure Type
Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

'Owner Draw Control Types
Const ODT_LISTBOX = 2

'Standard Windows Style
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000
Const WS_BORDER = &H800000

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

'List Box Styles:
Public Const LBS_MULTIPLESEL = &H8
Public Const LBS_OWNERDRAWFIXED = &H10

'ListBox Notification Messages:
Const LBN_SELCHANGE = 1
Const LBN_DBLCLK = 2
Const LBN_SELCANCEL = 3
Const LBN_SETFOCUS = 4
Const LBN_KILLFOCUS = 5

Const LB_SETCURSEL = &H186
Const LB_GETCURSEL = &H188
Const LB_GETTEXT = &H189
Const LB_GETTEXTLEN = &H18A
Const LB_SETCARETINDEX = &H19E
Const LB_GETCARETINDEX = &H19F
Const LB_ITEMFROMPOINT = &H1A9

'The Windows Standard Color Constants
Const COLOR_WINDOW = 5
Const COLOR_WINDOWTEXT = 8
Const COLOR_HIGHLIGHT = 13
Const COLOR_HIGHLIGHTTEXT = 14
Const COLOR_GRAYTEXT = 17

'Notify Message
Private Const WM_NOTIFY& = &H4E
Public Const WM_COMMAND = &H111
Const WM_DRAWITEM = &H2B

'Mouse Constants
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONDBLCLK = &H203

' Windows Messages Related To Keyboard
Const WM_KEYDOWN = &H100
Const WM_CHAR = &H102

Const WM_GETFONT = &H31
Const WM_PAINT = &HF
Const WM_ERASEBKGND = &H14

'Constants Relating To The Selected Status Of List Item:
Const ODA_FOCUS = &H4
Const ODS_FOCUS = &H10
Const ODS_SELECTED = &H1

Public Const GW_OWNER = 4
Public Const GWL_WNDPROC = (-4)

Const TwoPower16 = 2 ^ 16

Private lpListBox As ListBox
Private lpParent As Form
'\\ The above can be changed in case the parent is not a form. Any container
'\\ control will do. (ie: frame, picturebox, form, etc, etc)

Public lpHBitmap(3) As StdPicture

Public oldWndProc As Long

Private LBProc1 As Long
Private m_Hooked_LBhWnd As Long
Private m_LBHwnd As Long
Private m_Hooked_hWnd As Long
Private m_bCustomDraw As Boolean

Public Function List_Set(lpLB As ListBox) As Boolean
    
    On Error GoTo errset
        Set lpListBox = lpLB
        m_LBHwnd = lpListBox.hwnd
        List_Set = True
        Exit Function

errset:

End Function

Public Function List_SetParent(lpLBBoss As Form) As Boolean
    
    On Error GoTo errset
        Set lpParent = lpLBBoss
        List_SetParent = True
        Exit Function

errset:

End Function

Private Function List_Subclass_Proc3(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
'\\ This is for subclassing the parent, so that we can listen to listbox's
'\\ notification messages.
    
Dim iHw As Integer, iLW As Integer
Dim lCurind As Long
    
    Select Case Msg
        
        Case WM_COMMAND
            
            If lParam = m_LBHwnd Then
                LongInt2Int wParam, iHw, iLW
                
                Select Case (iHw)
                    Case LBN_SELCANCEL
                        lCurind = SendMessage(lParam, LB_GETCURSEL, 0, ByVal 0&)
                    
                    Case LBN_DBLCLK
                
                End Select
            
            End If
        
        Case WM_DRAWITEM
            
            If List_DrawItem(lParam) = 0 Then
                '\\ We have handled the painting, so dont pass on to
                '\\ default window procedure.

                List_Subclass_Proc3 = 0
                Exit Function
            
            End If
        
        Case Else
    
    End Select
    
    List_Subclass_Proc3 = CallWindowProc(oldWndProc, hwnd, Msg, wParam, lParam)
    
End Function

Public Sub Hook_Set()
    
    If lpListBox Is Nothing Then
        MsgBox "List box pointer not set. Wrong sequence of calls"
        Exit Sub
    End If
    
    If lpParent Is Nothing Then
        MsgBox "parent pointer for the LB not set. Wrong sequence of calls"
        Exit Sub
    End If
    
    Hook_SetList lpListBox.hwnd, True
    Hook_SetParent lpParent.hwnd, True

End Sub
Public Sub Hook_Unset()
    
    Hook_SetList vbNull, False
    Hook_SetParent vbNull, False

End Sub

Public Sub Hook_SetParent(ByVal hwnd As Long, b As Boolean)
    
    If b Then
        oldWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf List_Subclass_Proc3)
        m_Hooked_hWnd = hwnd
        m_bCustomDraw = True
    Else
        Call SetWindowLong(m_Hooked_hWnd, GWL_WNDPROC, oldWndProc)
        m_Hooked_hWnd = 0
    End If

End Sub

Public Sub Hook_SetList(ByVal LBhWnd As Long, b As Boolean)
    
    If b Then
        LBProc1 = SetWindowLong(LBhWnd, GWL_WNDPROC, AddressOf List_Subclass_Proc4)
        m_Hooked_LBhWnd = LBhWnd
    Else
        Call SetWindowLong(m_Hooked_LBhWnd, GWL_WNDPROC, LBProc1)
        m_Hooked_LBhWnd = 0
    End If

End Sub

Private Function List_DrawItem(ByVal lParam As Long) As Integer
    
'\\ Item Draw Notification Event handler for List Box.

Dim drawstruct As DRAWITEMSTRUCT
Dim szBuf(256) As Byte
    
    CopyMemory drawstruct, ByVal lParam, Len(drawstruct)
    
Dim i As Integer
Dim hbrGray As Long, hbrback As Long, szListStr As String ' * 256
Dim crback As Long, crtext As Long, lbuflen As Long
    
    Select Case (drawstruct.CtlType)
            
            Case ODT_LISTBOX:
                lbuflen = SendMessage(drawstruct.hwndItem, LB_GETTEXTLEN, drawstruct.itemID, ByVal 0&)
                lbuflen = SendMessage(drawstruct.hwndItem, LB_GETTEXT, drawstruct.itemID, szBuf(0))
                
                i = drawstruct.itemID
            
                If (drawstruct.itemState And ODS_FOCUS) = ODS_FOCUS Then
                    '\\ Set background and text colors for selected items.
                    crback = RGB(82, 151, 249) '\\ Background color
                    crtext = RGB(0, 0, 0) '\\ Font color
                Else
                    crback = lpListBox.BackColor
                    crtext = lpListBox.ForeColor
                End If
            
                If (drawstruct.itemState And ODS_FOCUS) = ODS_FOCUS Then
                    crtext = RGB(0, 0, 0)
                End If
            
                '\\ Fill item rectangle with background color.
                hbrback = CreateSolidBrush(crback)
                FillRect drawstruct.hdc, drawstruct.rcItem, hbrback
                DeleteObject hbrback
                
                '\\ Set current background and text colors.
                SetBkColor drawstruct.hdc, crback
                SetTextColor drawstruct.hdc, crtext
                
                '\\ TextOut uses current background and text colors.
                TextOutBStr drawstruct.hdc, drawstruct.rcItem.Left, drawstruct.rcItem.Top, szBuf(0), lbuflen
            
                '\\ If enabled item has the input focus, call DrawFocusRect to
                '\\ set or clear the focus rectangle.
                If (drawstruct.itemState And ODS_FOCUS) Then
                    DrawFocusRect drawstruct.hdc, drawstruct.rcItem
                End If
            
                List_DrawItem = 0
        
        Case Else '\\ VERY IMPORTANT! - We are not handling others.

            
            List_DrawItem = 1
    
    End Select

End Function

Public Function LongInt2Int(ByVal lLongInt As Long, ByRef iHiWord As Integer, ByRef iLowWord As Integer) As Boolean

Dim tmpHW As Integer, tmpLW As Integer
    
    CopyMemory tmpLW, lLongInt, Len(tmpLW)
    tmpHW = (lLongInt / TwoPower16)
    iHiWord = tmpHW
    iLowWord = tmpLW

End Function

Public Function MakeLParam(ByVal iHiWord As Integer, ByVal iLowWord As Integer) As Long
    
    MakeLParam = (iHiWord * TwoPower16) + iLowWord

End Function

Private Function List_Subclass_Proc4(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'\\ This is to sub class the List box itself.

Dim iHw As Integer, iLW As Integer
Dim lCurind As Long
    
    Select Case Msg
        Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
            LongInt2Int lParam, iHw, iLW
            lCurind = SendMessage(hwnd, LB_ITEMFROMPOINT, ByVal 0, ByVal lParam)
            
        Case WM_KEYDOWN
            LongInt2Int wParam, iHw, iLW
            
            Select Case (iLW)
                
                Case vbKeyDown
                    
                    '\\ Notice that we are still letting the Focus rect be
                    '\\ drawn by default window proc. We are just changing the
                    '\\ Caret item /focus item.

                    lCurind = SendMessage(hwnd, LB_GETCARETINDEX, 0, ByVal 0&)
                    
                    If ((lCurind + 1) Mod 3) = 0 Then
                        lCurind = SendMessage(hwnd, LB_SETCARETINDEX, lCurind + 1, ByVal 0&)
                    End If
                    
                    lCurind = SendMessage(hwnd, LB_GETCURSEL, 0, ByVal 0&)
                    
                    If ((lCurind + 1) Mod 3) = 0 Then
                        lCurind = SendMessage(hwnd, LB_SETCURSEL, lCurind + 1, ByVal 0&)
                    End If
                
                Case vbKeyUp
                    
                    lCurind = SendMessage(hwnd, LB_GETCARETINDEX, 0, ByVal 0&)
                    
                    If ((lCurind - 1) Mod 3) = 0 Then
                        lCurind = SendMessage(hwnd, LB_SETCARETINDEX, lCurind - 1, ByVal 0&)
                    End If
                    
                    lCurind = SendMessage(hwnd, LB_GETCURSEL, 0, ByVal 0&)
                    
                    If ((lCurind - 1) Mod 3) = 0 Then
                        lCurind = SendMessage(hwnd, LB_SETCURSEL, lCurind - 1, ByVal 0&)
                    End If
                
                '\\ These two are also to be handled.
                Case vbKeyPageUp
                Case vbKeyPageDown
                    
            End Select
        
        Case Else
    
    End Select
    
    List_Subclass_Proc4 = CallWindowProc(LBProc1, hwnd, Msg, wParam, lParam)
    
End Function
