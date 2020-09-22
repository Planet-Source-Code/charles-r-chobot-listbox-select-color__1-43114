VERSION 5.00
Begin VB.Form frm_subclass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listbox Subclassing - Select Color"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst_subclass 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2760
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frm_subclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'\\ Topic: Alphablending
'\\ Author: Charles R. Chobot [aka Kurupt]
'\\ Coded In: Visual Basic 6 [EE]

'\\ Notes: This example was coded because I have grown tired of
'          not being able to set a select color for a listbox
'          and having to use the s

'\\ Credits: Part of this code was ported from a C++ example found
'            on MSDN.

'\\          I would like to thank:
'                 Microsoft Developers Network : http://www.msdn.com

Option Explicit

Private m_hooked As Boolean

Private Sub Form_Activate()

Dim lrtn As Long
    
    lrtn = GetWindowLong(lst_subclass.hwnd, GWL_STYLE)
    If (lrtn And LBS_OWNERDRAWFIXED) = LBS_OWNERDRAWFIXED Then
    End If
    
    If (lrtn And LBS_MULTIPLESEL) = LBS_MULTIPLESEL Then
        lrtn = lrtn Xor LBS_MULTIPLESEL
    End If
    lrtn = SetWindowLong(lst_subclass.hwnd, GWL_STYLE, lrtn)
    
End Sub

Private Sub Form_Load()
    
Dim i As Integer

    If App.PrevInstance = True Then
        MsgBox "There is an instance already running. Cant run one more"
        Unload Me
    Else

    Do
    DoEvents
        For i = 1 To 10
            lst_subclass.AddItem ("Hello there!")
        Next i
    Loop Until lst_subclass.ListCount = 10
    
        mod_subclass.List_Set lst_subclass
        Hook_SetParent Me.hwnd, True
        m_hooked = True
    
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mod_subclass.Hook_Unset

End Sub
