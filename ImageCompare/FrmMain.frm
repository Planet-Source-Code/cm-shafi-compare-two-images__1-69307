VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Compare Two Images -by Cm.Shafi"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12120
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   12015
      TabIndex        =   8
      Top             =   5400
      Width           =   12015
      Begin VB.CommandButton CmdClear 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Clear all"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdAction 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Compare Both Images"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   13
         Top             =   0
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Case Sensitive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton CmdRefresh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Clear Result"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10920
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdBrowse1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Browse Image 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton CmdBrowse2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Browse Image 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   9
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   0
         TabIndex        =   16
         Top             =   660
         Width           =   960
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   0
         Top             =   600
         Width           =   12015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Match:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1920
         TabIndex        =   15
         Top             =   660
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Match:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   8400
         TabIndex        =   14
         Top             =   660
         Width           =   1335
      End
   End
   Begin VB.PictureBox PicBk 
      Height          =   5055
      Index           =   3
      Left            =   6120
      ScaleHeight     =   4995
      ScaleWidth      =   5100
      TabIndex        =   3
      Top             =   6600
      Width           =   5160
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   3
         Left            =   1680
         MouseIcon       =   "FrmMain.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   127
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can move image by draging the mouse"
         Height          =   195
         Index           =   3
         Left            =   1305
         TabIndex        =   21
         Top             =   0
         Width           =   3060
      End
   End
   Begin VB.PictureBox PicBk 
      Height          =   5055
      Index           =   2
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   5100
      TabIndex        =   2
      Top             =   6360
      Width           =   5160
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   2
         Left            =   1560
         MouseIcon       =   "FrmMain.frx":0614
         MousePointer    =   99  'Custom
         ScaleHeight     =   127
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can move image by draging the mouse"
         Height          =   195
         Index           =   2
         Left            =   945
         TabIndex        =   20
         Top             =   0
         Width           =   3060
      End
   End
   Begin VB.PictureBox PicBk 
      Height          =   5055
      Index           =   1
      Left            =   5760
      ScaleHeight     =   4995
      ScaleWidth      =   5100
      TabIndex        =   1
      Top             =   120
      Width           =   5160
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3945
         Index           =   1
         Left            =   2040
         MouseIcon       =   "FrmMain.frx":091E
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":0C28
         ScaleHeight     =   261
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   172
         TabIndex        =   5
         Top             =   480
         Width           =   2610
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can move image by draging the mouse"
         Height          =   195
         Index           =   1
         Left            =   1785
         TabIndex        =   19
         Top             =   0
         Width           =   3060
      End
   End
   Begin MSComDlg.CommonDialog dlgPicture 
      Left            =   -120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicBk 
      Height          =   5055
      Index           =   0
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   5100
      TabIndex        =   0
      Top             =   120
      Width           =   5160
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3945
         Index           =   0
         Left            =   1080
         MouseIcon       =   "FrmMain.frx":51D3
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":54DD
         ScaleHeight     =   261
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   172
         TabIndex        =   4
         Top             =   480
         Width           =   2610
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can move image by draging the mouse"
         Height          =   195
         Index           =   0
         Left            =   705
         TabIndex        =   18
         Top             =   0
         Width           =   3060
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Cm.Shafi
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Dim dXM(3) As Long, DYM(3) As Long
Dim isStart As Boolean
Private Sub CmdAction_Click()
 Dim Xp As Long, Yp As Long, Xd As Long, Yd As Long, i As Long, j As Long
    Dim P1, P2
    Dim a1, a2
    Dim G As Long, M As Long
    Pic(0).ScaleMode = vbPixels         'to Pixel mode
    Pic(1).ScaleMode = vbPixels

    For i = 2 To 3
        Set Pic(i).Picture = LoadPicture() 'clear Previous result
        Pic(i).Cls
    Next
    Label1.Caption = "Match:" & " %"
    Label2.Caption = "Not Match:" & " %"
    If Pic(0).Height <= Pic(1).Height Then
        Pic(2).Height = Pic(0).Height
        Pic(3).Height = Pic(1).Height
    Else
        Pic(2).Height = Pic(1).Height
        Pic(3).Height = Pic(0).Height
    End If
    If Pic(0).Width <= Pic(1).Width Then
        Pic(2).Width = Pic(0).Width
        Pic(3).Width = Pic(1).Width
    Else
        Pic(2).Width = Pic(1).Width
        Pic(3).Width = Pic(0).Width
    End If
     For i = 0 To 3
        Pic(i).Move (PicBk(i).Width - Pic(i).Width) / 2, (PicBk(i).Height - Pic(i).Height) / 2 ' set default position
    Next
    DoEvents
    Xp = Pic(0).ScaleWidth
    Yp = Pic(0).ScaleHeight
     a1 = (Xp) * (Yp)
    Pic(2).Cls
     Pic(3).Cls
    For i = 0 To Xp - 1
        For j = 0 To Yp - 1
            P1 = GetPixel(Pic(0).hDC, i, j)     'Get colour from images pixel by pixel
            P2 = GetPixel(Pic(1).hDC, i, j)
            If P1 = P2 Then
                G = G + 1
                a2 = SetPixel(Pic(2).hDC, i, j, P2) 'Set pixel to result's image
                Else
                If (P1 < (P2 + 1000000) And P1 > (P2 - 1000000)) And Check1.Value = 0 Then
                G = G + 1
                a2 = SetPixel(Pic(2).hDC, i, j, P2)
                Else
                a2 = SetPixel(Pic(3).hDC, i, j, P2)
                M = M + 1
                End If
            End If
'            Label1.Caption = "Match:" & Round((G * 100) / a1, 4) & " %"
'            Label2.Caption = "Not Match:" & Round((M * 100) / a1, 4) & " %"
        Next
        Label1.Caption = "Match:" & Round((G * 100) / a1, 4) & " %"
        Label2.Caption = "Not Match:" & Round((M * 100) / a1, 4) & " %"
    Next
'    Label1.Caption = "Match:" & Round((G * 100) / a1, 4) & " %"
'    Label2.Caption = "Not Match:" & Round((M * 100) / a1, 4) & " %"
    
End Sub


Private Sub CmdBrowse1_Click()
Dim sFile As String
Dim Bcolor As ColorConstants
    dlgPicture.Flags = _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNExplorer
    dlgPicture.CancelError = True
    dlgPicture.Filter = "Graphics Files|*.bmp;*.ico;*.jpg;*.gif"

    On Error Resume Next
    dlgPicture.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If
   sFile = dlgPicture.FileName
Set Pic(0).Picture = LoadPicture(sFile)

End Sub

Private Sub CmdBrowse2_Click()
Dim sFile As String
Dim Bcolor As ColorConstants
    dlgPicture.Flags = _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNExplorer
    dlgPicture.CancelError = True
    dlgPicture.Filter = "Graphics Files|*.bmp;*.ico;*.jpg;*.gif"

    On Error Resume Next
    dlgPicture.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If
   sFile = dlgPicture.FileName
Set Pic(1).Picture = LoadPicture(sFile)

End Sub

Private Sub Cmdclear_Click()
    Dim i As Integer
    For i = 0 To 3
        Set Pic(i).Picture = LoadPicture()
        Pic(i).Cls
    Next
    Label1.Caption = "Match:" & " %"
    Label2.Caption = "Not Match:" & " %"
    
End Sub

Private Sub CmdRefresh_Click()
    Dim i As Integer
    For i = 2 To 3
        Set Pic(i).Picture = LoadPicture()
        Pic(i).Cls
    Next
    Label1.Caption = "Match:" & " %"
    Label2.Caption = "Not Match:" & " %"

End Sub

Private Sub Form_Load()
Dim i As Integer
    Me.Move 0, 0, Screen.Width, Screen.Height
    
    PicBk(0).Move 0, 0, (Me.Width / 2) - 60, (Me.Height - 1500) / 2
    PicBk(1).Move PicBk(0).Width + 60, 0, (Me.Width / 2) - 60, (Me.Height - 1500) / 2
    PicBk(2).Move 0, PicBk(0).Height + 1000, (Me.Width / 2) - 60, (Me.Height - 1500) / 2
    PicBk(3).Move PicBk(0).Width + 60, PicBk(0).Height + 1000, (Me.Width / 2) - 60, (Me.Height - 1500) / 2
    Picture1.Move 0, PicBk(0).Height, Me.Width, 1000
    CmdBrowse1.Left = (PicBk(0).Width - CmdBrowse1.Width) / 2
    CmdBrowse2.Left = PicBk(0).Width + 60 + (PicBk(1).Width - CmdBrowse2.Width) / 2
    Label1.Left = (PicBk(0).Width - Label1.Width) / 2
    Label2.Left = PicBk(0).Width + 60 + (PicBk(1).Width - Label2.Width) / 2
    CmdAction.Left = (Me.Width - CmdAction.Width) / 2
    CmdRefresh.Left = (Me.Width - CmdRefresh.Width) - 260
    Check1.Left = (Me.Width - Check1.Width) / 2
    Shape1.Left = 0
    Shape1.Width = Me.Width
    For i = 0 To 3
        Pic(i).Move (PicBk(i).Width - Pic(i).Width) / 2, (PicBk(i).Height - Pic(i).Height) / 2
        Label4(i).Left = (PicBk(i).Width - Label4(i).Width) / 2
    Next
    
End Sub



Private Sub Pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    isStart = True
    dXM(Index) = X
    DYM(Index) = Y
End Sub

Private Sub Pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isStart Then
        Pic(Index).Left = Pic(Index).Left + (X - dXM(Index))
        Pic(Index).Top = Pic(Index).Top + (Y - DYM(Index))
        Me.Refresh
    End If
End Sub

Private Sub Pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    isStart = False
End Sub
