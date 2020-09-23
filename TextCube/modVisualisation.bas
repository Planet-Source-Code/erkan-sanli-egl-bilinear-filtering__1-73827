Attribute VB_Name = "modVisualisation"
Option Explicit

Const ApproachVal As Single = 0.000001

Private Type TEXEL
    Y1      As Long
    Y2      As Long
    M1      As MAPCOORS
    M2      As MAPCOORS
    Used    As Boolean
End Type

Dim Texels() As TEXEL

Public Sub Render()
    
    Dim idx     As Integer
    Dim minX    As Long
    Dim maxX    As Long
    Dim P1      As POINTAPI
    Dim P2      As POINTAPI
    Dim P3      As POINTAPI
    Dim M1      As MAPCOORS
    Dim M2      As MAPCOORS
    Dim M3      As MAPCOORS
    
    If VisibleFaces < 0 Then Exit Sub
    
    With Meshs
        For idx = 0 To UBound(.FaceV)
            P1 = .Screen(.Faces(.FaceV(idx).Index).A)
            P2 = .Screen(.Faces(.FaceV(idx).Index).B)
            P3 = .Screen(.Faces(.FaceV(idx).Index).C)
            M1 = .TVerts(.TFaces(.FaceV(idx).Index).A)
            M2 = .TVerts(.TFaces(.FaceV(idx).Index).B)
            M3 = .TVerts(.TFaces(.FaceV(idx).Index).C)
            
            minX = IIf(P1.X < P2.X, P1.X, P2.X)
            If P3.X < minX Then minX = P3.X
            maxX = IIf(P1.X > P2.X, P1.X, P2.X)
            If P3.X > maxX Then maxX = P3.X
            ReDim Texels(minX To maxX)
            
            LineerInterpolateTex P1, P2, M1, M2
            LineerInterpolateTex P2, P3, M2, M3
            LineerInterpolateTex P3, P1, M3, M1
            
            'Limits Width
            If minX < 0 Then minX = 0
            If maxX > dibCanvas.Width - 1 Then maxX = dibCanvas.Width - 1
            
            For minX = minX To maxX
                FillTexLine minX
            Next
        Next
    End With
End Sub

Private Sub LineerInterpolateTex(P1 As POINTAPI, P2 As POINTAPI, M1 As MAPCOORS, M2 As MAPCOORS)

    Dim DeltaX  As Long
    Dim X1      As Long
    Dim X2      As Long
    Dim Y1      As Single
    Dim Y2      As Single
    Dim MM1     As MAPCOORS
    Dim MM2     As MAPCOORS
    Dim StepY   As Single
    Dim StepU   As Single
    Dim StepV   As Single
    
    X1 = P1.X
    X2 = P2.X
    Y1 = P1.Y
    Y2 = P2.Y
    MM1 = M1
    MM2 = M2
    
    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Div(Y2 - Y1, DeltaX)
        StepU = Div(MM2.U - MM1.U, DeltaX)
        StepV = Div(MM2.V - MM1.V, DeltaX)
        For X1 = X1 To X2
            With Texels(X1)
                If .Used Then
                    If .Y1 < Fix(Y1) Then .Y1 = Fix(Y1): .M1 = MM1
                    If .Y2 > Fix(Y1) Then .Y2 = Fix(Y1): .M2 = MM1
                Else
                    .Y1 = Fix(Y1): .M1 = MM1
                    .Y2 = Fix(Y1): .M2 = MM1
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
            MM1.U = MM1.U + StepU
            MM1.V = MM1.V + StepV
        Next
    Else
        DeltaX = X1 - X2
        StepY = Div(Y1 - Y2, DeltaX)
        StepU = Div(MM1.U - MM2.U, DeltaX)
        StepV = Div(MM1.V - MM2.V, DeltaX)
        For X2 = X2 To X1
            With Texels(X2)
                If .Used Then
                    If .Y1 < Fix(Y2) Then .Y1 = Fix(Y2): .M1 = MM2
                    If .Y2 > Fix(Y2) Then .Y2 = Fix(Y2): .M2 = MM2
                Else
                    .Y1 = Fix(Y2): .M1 = MM2
                    .Y2 = Fix(Y2): .M2 = MM2
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
            MM2.U = MM2.U + StepU
            MM2.V = MM2.V + StepV
        Next
   End If
End Sub

Private Sub FillTexLine(X As Long)
    
    Dim DeltaY  As Long
    Dim minY    As Long
    Dim maxY    As Long
    Dim StepU   As Single
    Dim StepV   As Single
    
    On Error Resume Next
    With Texels(X)
        DeltaY = .Y1 - .Y2
        StepU = Div(.M1.U - .M2.U, DeltaY)
        StepV = Div(.M1.V - .M2.V, DeltaY)
        
        'Limits Height
        If .Y2 < 0 Then
            minY = 0
            .M2.U = .M2.U + (StepU * Abs(.Y2))
            .M2.V = .M2.V + (StepV * Abs(.Y2))
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > dibCanvas.Height - 1, dibCanvas.Height - 1, .Y1)
        
        If frmMain.Check1.Value = vbChecked Then
            For minY = minY To maxY
                Bilinear X, minY, .M2
                .M2.U = .M2.U + StepU
                .M2.V = .M2.V + StepV
            Next
        Else
            For minY = minY To maxY
                Buffer(X, minY) = dibTex.bi.Bits(.M2.U, .M2.V)
                .M2.U = .M2.U + StepU
                .M2.V = .M2.V + StepV
            Next
        End If
    End With
    
End Sub

Private Function Div(R1 As Single, ByVal R2 As Single) As Single
    
    If R2 = 0 Then R2 = ApproachVal
    Div = R1 / R2

End Function

Private Sub Bilinear(X As Long, Y As Long, M As MAPCOORS)

    Dim U1       As Long
    Dim V1       As Long
    Dim U2       As Long
    Dim V2       As Long
    Dim fU       As Single 'Fraction U
    Dim fV       As Single 'Fraction V
    Dim osfU     As Single 'OneSubtractFraction U
    Dim osfV     As Single 'OneSubtractFraction V
    Dim fUfV     As Single 'Fraction U Mult. Fraction V
    Dim fUosfV   As Single '...
    Dim osfUfV   As Single '...
    Dim osfUosfV As Single '...
    
    On Error Resume Next
    
    U1 = Fix(M.U)
    V1 = Fix(M.V)
    U2 = U1 + 1
    V2 = V1 + 1
    fU = M.U - U1
    fV = M.V - V1
    osfU = 1 - fU
    osfV = 1 - fV
    fUfV = fU * fV
    fUosfV = fU * osfV
    osfUfV = osfU * fV
    osfUosfV = osfU * osfV

    Buffer(X, Y).R = fUfV * dibTex.bi.Bits(U2, V2).R + fUosfV * dibTex.bi.Bits(U2, V1).R + _
                      osfUfV * dibTex.bi.Bits(U1, V2).R + osfUosfV * dibTex.bi.Bits(U1, V1).R

    Buffer(X, Y).G = fUfV * dibTex.bi.Bits(U2, V2).G + fUosfV * dibTex.bi.Bits(U2, V1).G + _
                      osfUfV * dibTex.bi.Bits(U1, V2).G + osfUosfV * dibTex.bi.Bits(U1, V1).G

    Buffer(X, Y).B = fUfV * dibTex.bi.Bits(U2, V2).B + fUosfV * dibTex.bi.Bits(U2, V1).B + _
                      osfUfV * dibTex.bi.Bits(U1, V2).B + osfUosfV * dibTex.bi.Bits(U1, V1).B

End Sub
