VERSION 5.00
Begin VB.Form FrmDD 
   Caption         =   "Missile Command"
   ClientHeight    =   3510
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   6900
   Icon            =   "FrmDD.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5160
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePlay 
         Caption         =   "&Play"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FrmDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

    If blnFirstTime Then
        
        blnFirstTime = False
        
        DoIt
        
    End If

End Sub

Private Sub Form_Load()

Dim w%, h%
    
    gstrPath = App.Path & "\"
    
    blnFirstTime = True
    
    'w% = (1440 * 8) + 200
    'h% = (1440 * 6) + 1440
    
    w% = (1440 * 7) + 200
    h% = (1440 * 4) + 1440

    Me.Move (Screen.Width - w%) / 2, (Screen.Height - h%) / 2, w%, h%
    Pic.Move 30, 30, w% - 200, h% - 1440

    Me.Show
    
    Label1(0).Left = 100
    Label1(0).Top = Pic.Top + Pic.Height + 100
    
    Label1(1).Left = 100
    Label1(1).Top = Pic.Top + Pic.Height + 100 + 310
    
    DoEvents
    
End Sub

Private Sub Form_Resize()

    If blnAllowResize Then
        ReSize
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    blnAllowResize = False
    blnQuit = True

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

'Private Sub mnuFileOptions_Click()

    'frmOptions.Show vbModal

'End Sub

Private Sub mnuFilePlay_Click()

    DoIt

End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    mMouseX = x
    mMouseY = Y
    mMouseButton = Button       '1=left, 2=middle, 3=right

End Sub
Private Sub ReSize()

    Pic.Move 30, 30, Me.Width - 200, Me.Height - 1440
    Pic.Cls
    Pic.ScaleWidth = MaxX%
    Pic.ScaleHeight = MaxY%
    
    Label1(0).Left = 100
    Label1(0).Top = Pic.Top + Pic.Height + 100
    
    Label1(1).Left = 100
    Label1(1).Top = Pic.Top + Pic.Height + 100 + 310
    
    ResetTargets Targets, gstrPath & tFile, Pic

End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 2 Then
    
        mMouseX = x
        mMouseY = Y
        mMouseButton = 1 'left
    
    End If

End Sub
