Attribute VB_Name = "mod3DMath"
Option Explicit

Public Const sPIDiv180 As Single = 0.0174533 'PI / 180


Public Function MatrixMultVertex(M As MATRIX, V As VECTOR) As VECTOR

    MatrixMultVertex.X = M.rc11 * V.X + M.rc12 * V.Y + M.rc13 * V.Z '+ M.rc14
    MatrixMultVertex.Y = M.rc21 * V.X + M.rc22 * V.Y + M.rc23 * V.Z '+ M.rc24
    MatrixMultVertex.Z = M.rc31 * V.X + M.rc32 * V.Y + M.rc33 * V.Z '+ M.rc34
    'MatrixMultVertex.W = 1

End Function

Public Function MatrixWorld() As MATRIX
    
    Dim CosX As Single
    Dim SinX As Single
    Dim CosY As Single
    Dim SinY As Single
    Dim CosZ As Single
    Dim SinZ As Single
    
    With Meshs
        With .Rotation
            CosX = Cos(.X * sPIDiv180)
            SinX = Sin(.X * sPIDiv180)
            CosY = Cos(.Y * sPIDiv180)
            SinY = Sin(.Y * sPIDiv180)
            CosZ = Cos(.Z * sPIDiv180)
            SinZ = Sin(.Z * sPIDiv180)
        End With
        MatrixWorld.rc11 = .Scales.X * CosY * CosZ
        MatrixWorld.rc12 = .Scales.Y * (SinX * SinY * CosZ + CosX * -SinZ)
        MatrixWorld.rc13 = .Scales.Z * (CosX * SinY * CosZ + SinX * SinZ)
        MatrixWorld.rc14 = 0 '.Translation.x
        MatrixWorld.rc21 = .Scales.X * CosY * SinZ
        MatrixWorld.rc22 = .Scales.Y * (SinX * SinY * SinZ + CosX * CosZ)
        MatrixWorld.rc23 = .Scales.Z * (CosX * SinY * SinZ + -SinX * CosZ)
        MatrixWorld.rc24 = 0 '.Translation.y
        MatrixWorld.rc31 = .Scales.X * -SinY
        MatrixWorld.rc32 = .Scales.Y * SinX * CosY
        MatrixWorld.rc33 = .Scales.Z * CosX * CosY
        MatrixWorld.rc34 = 0 '.Translation.Z
        'MatrixWorld.rc41 = 0
        'MatrixWorld.rc42 = 0
        'MatrixWorld.rc43 = 0
        'MatrixWorld.rc44 = 1
    End With
    
End Function

Public Function VectorSet(X As Single, Y As Single, Z As Single) As VECTOR

    VectorSet.X = X
    VectorSet.Y = Y
    VectorSet.Z = Z
    VectorSet.W = 1
    
End Function

Public Function FaceVisible(V1 As VECTOR, V2 As VECTOR, V3 As VECTOR) As Boolean

    FaceVisible = ((V3.X - V2.X) * (V1.Y - V2.Y) - (V3.Y - V2.Y) * (V1.X - V2.X) > 0)

End Function

           

