VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TexCube"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":3A0A
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "Hello World"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer tmrProcess 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   2280
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bilinear Filter"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox picTex 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1035
      Left            =   120
      Picture         =   "frmMain.frx":E4A4C
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cTimer  As New clsTiming
Dim OriginX As Long
Dim OriginY As Long

Private Sub Command1_Click()
    
    picTex.Cls
    picTex.Print Text1.Text
    'picTex.Line (0, 15)-(60, 15), vbBlue
    'picTex.Circle (30, 30), 20, vbRed
    picTex.Refresh
    Clear dibTex
    dibTex.Width = picTex.ScaleWidth
    dibTex.Height = picTex.ScaleHeight
    dibTex.hBmp = picTex.Image.handle
    CreateArrayFromPicbox dibTex
    Changed = True
    
End Sub

Private Sub Form_Load()
    
    OriginX = Me.ScaleWidth / 2
    OriginY = Me.ScaleHeight / 2
    
'Canvas
    dibCanvas.Width = Me.ScaleWidth
    dibCanvas.Height = Me.ScaleHeight
    dibCanvas.hBmp = Me.Image.handle
    CreateArrayFromPicbox dibCanvas

'Image,text,geo
    Command1_Click
    
    CreateCube
    Check1.Value = vbChecked
    Changed = True
    Me.Show
    tmrProcess.Enabled = True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button Then
        Changed = True
        rX = X: rY = Y
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Changed = True
        With Meshs
            .Rotation.X = Y - rY: .Rotation.X = .Rotation.X Mod 360
            .Rotation.Y = X - rX: .Rotation.Y = .Rotation.Y Mod 360
        End With
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    tmrProcess.Enabled = False
    Changed = False
    Set cTimer = Nothing
    Clear dibTex
    Clear dibCanvas
    Erase Buffer
    End

End Sub

Private Sub Check1_Click()
    
    Changed = True

End Sub


Private Sub tmrProcess_Timer()
    
    Dim idx As Integer
    
    DoEvents
    With Meshs
        UpdateMeshParameters
        If Changed = True Then
            cTimer.Reset
            For idx = 0 To .numVert
                .Vertices(idx).VectorsT = MatrixMultVertex(.WorldMatrix, .Vertices(idx).Vectors)
                .Screen(idx).X = .Vertices(idx).VectorsT.X + OriginX
                .Screen(idx).Y = .Vertices(idx).VectorsT.Y + OriginY
            Next idx
            Buffer = dibCanvas.bi.Bits
            Render
            SetBits Me.hDC, dibCanvas, Buffer
            Me.Refresh
            Changed = False
            Me.Caption = cTimer.Elapsed
        End If
    End With

End Sub
