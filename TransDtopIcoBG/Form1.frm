VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   375
   ClientLeft      =   4665
   ClientTop       =   5400
   ClientWidth     =   1935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Make Transparent"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by :: em ::® 2000
'http://www.em.f2s.com
'I searched the net hi n low for VB code to make the desktop icon
'text background transparent but to no avail, so I wrote this
'This is, to my knowledge, the First EVER VB code on how to
'achieve this =)
'hope you learn something & give credit where it's due
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex%) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const COLOR_BACKGROUND = 1
Private Const LVM_FIRST = &H1000 ' ListView messages
Private Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Private Const CLR_NONE = &HFFFFFFFF

Private Sub Command1_Click()
Dim bRet As Boolean
Dim Progman As Long
Dim SHELLDLLDefView As Long
Dim SysListView32 As Long

bRet = False
Progman = FindWindow("Progman", "Program Manager")
If Progman <> 0 Then
 SHELLDLLDefView = FindWindowEx(Progman, 0&, "SHELLDLL_DefView", vbNullString)
  If SHELLDLLDefView <> 0 Then
   SysListView32 = FindWindowEx(SHELLDLLDefView, 0&, "SysListView32", vbNullString)
    If SysListView32 <> 0 Then
     If (ListView_GetTextBkColor(SysListView32) <> CLR_NONE) Then
      bRet = ListView_SetTextBkColor(SysListView32, CLR_NONE)
     Else
      Call ListView_SetTextBkColor(SysListView32, GetSysColor(COLOR_BACKGROUND))
     End If
     Call InvalidateRect(SysListView32, ByVal 0&, True)
     Call UpdateWindow(SysListView32)
     If bRet Then
      Command1.Caption = "Make Coloured"
     Else
      Command1.Caption = "Make Transparent"
     End If
    End If
  End If
End If

End Sub

Private Function ListView_SetTextBkColor(hwnd As Long, clrTextBk As Long) As Boolean
Dim lRet As Long

lRet = SendMessage((hwnd), LVM_SETTEXTBKCOLOR, 0&, clrTextBk)
If lRet = 0 Then
 ListView_SetTextBkColor = False
Else
 ListView_SetTextBkColor = True
End If
End Function

Private Function ListView_GetTextBkColor(hwnd As Long) As Long
ListView_GetTextBkColor = SendMessage((hwnd), LVM_GETTEXTBKCOLOR, 0, 0)
End Function
