VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox chkModalFrame 
      Caption         =   "Modal frame"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox chkStaticEdge 
      Caption         =   "Static edge"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkClientEdge 
      Caption         =   "Client edge"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*
'This code works only for CommandButton,PictureBox,TextBox,Frame and ListBox
'It can work for other controls such as CheckBox and OptionButton but you'll
'need to extend more styles
'*/

Private Declare Function GetWindowLong& Lib "user32" _
 Alias "GetWindowLongA" (ByVal hwnd As Long, _
 ByVal nIndex As Long)

Private Declare Function SetWindowLong& Lib "user32" _
 Alias "SetWindowLongA" (ByVal hwnd As Long, _
 ByVal nIndex As Long, ByVal dwNewLong As Long)

Private Declare Function SetWindowPos& Lib "user32" _
 (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
 ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
 ByVal cy As Long, ByVal wFLAGS As Long)

Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200

Private Const wFLAGS = _
    SWP_NOMOVE Or _
    SWP_NOSIZE Or _
    SWP_NOOWNERZORDER Or _
    SWP_NOZORDER Or _
    SWP_FRAMECHANGED

Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_DLGMODALFRAME = &H1
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000

Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Sub chkClientEdge_Click()
Dim lRet As Long
With Command1
    'Get the window extended style
    lRet = GetWindowLong(.hwnd, GWL_EXSTYLE)
    lRet = IIf(chkClientEdge.Value = vbChecked, lRet Or WS_EX_CLIENTEDGE, lRet And Not WS_EX_CLIENTEDGE)
    Call reDrawWindow(.hwnd, lRet)
End With
End Sub

Private Sub chkModalFrame_Click()
Dim lRet As Long
With Command3
    'Get the window extended style
    lRet = GetWindowLong(.hwnd, GWL_EXSTYLE)
    lRet = IIf(chkModalFrame.Value = vbChecked, lRet Or WS_EX_DLGMODALFRAME, lRet And Not WS_EX_DLGMODALFRAME)
    Call reDrawWindow(.hwnd, lRet)
End With
End Sub

Private Sub chkStaticEdge_Click()
Dim lRet As Long
With Command2
    'Get the window extended style
    lRet = GetWindowLong(.hwnd, GWL_EXSTYLE)
    lRet = IIf(chkStaticEdge.Value = vbChecked, lRet Or WS_EX_STATICEDGE, lRet And Not WS_EX_STATICEDGE)
    Call reDrawWindow(.hwnd, lRet)
End With
End Sub

Private Sub reDrawWindow(lhWnd As Long, lRet As Long)
    'Set the new extended style
    SetWindowLong lhWnd, GWL_EXSTYLE, lRet
    'Refresh the window
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, wFLAGS
End Sub
