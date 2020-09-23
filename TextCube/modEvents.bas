Attribute VB_Name = "modEvents"

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public rX As Single
Public rY As Single
Public rZ As Single

Private Function State(Key As KeyCodeConstants) As Boolean
    
    State = (GetKeyState(Key) And &H8000)

End Function

Public Sub UpdateMeshParameters()

    Const Value = 0.3

    With Meshs
'X: Reset meshs position
        If State(vbKeyX) Then
            Changed = True
            Call ResetMeshParameters
        End If
'Esc: Quit
        If State(vbKeyEscape) Then Unload frmMain
'R: Mesh Rotate
        If State(vbKeyR) Then
            Changed = True
            If State(vbKeyUp) Then
                rX = rX + Value
                .Rotation.X = .Rotation.X + rX
                .Rotation.X = .Rotation.X Mod 360
            End If
            If State(vbKeyDown) Then
                rX = rX - Value
                .Rotation.X = .Rotation.X + rX
                .Rotation.X = .Rotation.X Mod 360
            End If
            If State(vbKeyLeft) Then
                rY = rY - Value
                .Rotation.Y = .Rotation.Y + rY
                .Rotation.Y = .Rotation.Y Mod 360
            End If
            If State(vbKeyRight) Then
                rY = rY + Value
                .Rotation.Y = .Rotation.Y + rY
                .Rotation.Y = .Rotation.Y Mod 360
            End If
            If State(vbKeyPageUp) Then
                rZ = rZ - Value
                .Rotation.Z = .Rotation.Z + rZ
                .Rotation.Z = .Rotation.Z Mod 360
            End If
            If State(vbKeyPageDown) Then
                rZ = rZ + Value
                .Rotation.Z = .Rotation.Z + rZ
                .Rotation.Z = .Rotation.Z Mod 360
            End If
        End If
'Changed=True
        If Changed = True Then .WorldMatrix = MatrixWorld()
        
    End With
    
End Sub

Public Sub ResetMeshParameters()

    With Meshs
        .Rotation.X = 0
        .Rotation.Y = 0
        .Rotation.Z = 0
'        .Translation.x = 0
'        .Translation.y = 0
'        .Translation.Z = 0
        .Scales = VectorSet(6, 6, 6)
        rX = 0
        rY = 0
        rZ = 0
    End With
    
End Sub

