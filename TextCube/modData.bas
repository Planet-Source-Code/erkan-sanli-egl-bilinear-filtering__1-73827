Attribute VB_Name = "modData"
Option Explicit

Public Type VECTOR
    X               As Single
    Y               As Single
    Z               As Single
    W               As Single
End Type

Public Type VERTEX
    Vectors         As VECTOR
    VectorsT        As VECTOR
End Type

Public Type FACE
    A               As Integer
    B               As Integer
    C               As Integer
End Type

Public Type ORDER
    Value          As Single
    Index          As Integer
End Type

Public Type POINTAPI
    X               As Long
    Y               As Long
End Type

Public Type MAPCOORS
    U               As Single
    V               As Single
End Type

Public Type MATRIX
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type

Public Type MESH
    numVert         As Long
    numFace         As Integer
    numTVert        As Long
    numTFace        As Integer
    Vertices()      As VERTEX
    Faces()         As FACE
    Screen()        As POINTAPI
    FaceV()         As ORDER
    TVerts()        As MAPCOORS
    TFaces()        As FACE
    
    Rotation        As VECTOR
    Scales          As VECTOR
    WorldMatrix     As MATRIX
End Type

Public Meshs        As MESH
Public dibTex       As DIB
Public dibCanvas    As DIB
Public Buffer()     As COLORRGBA_BYTE
Public Changed      As Boolean

Public Sub CreateCube()
    
    Dim idx         As Integer
    Dim texWidth    As Integer
    Dim texHeight   As Integer
    Dim L           As Single
    
    L = 25
    With Meshs
        .numVert = 7
        .numTVert = 11
        .numFace = 11
        .numTFace = 11
        ReDim .Vertices(.numVert)
        ReDim .TVerts(.numTVert)
        ReDim .Screen(.numVert)
        ReDim .Faces(.numFace)
        ReDim .TFaces(.numTFace)
        .Vertices(0).Vectors = VectorSet(-L, -L, -L)
        .Vertices(1).Vectors = VectorSet(L, -L, -L)
        .Vertices(2).Vectors = VectorSet(-L, L, -L)
        .Vertices(3).Vectors = VectorSet(L, L, -L)
        .Vertices(4).Vectors = VectorSet(-L, -L, L)
        .Vertices(5).Vectors = VectorSet(L, -L, L)
        .Vertices(6).Vectors = VectorSet(-L, L, L)
        .Vertices(7).Vectors = VectorSet(L, L, L)
         
        .Faces(0).A = 0:        .Faces(0).B = 2:        .Faces(0).C = 3
        .Faces(1).A = 3:        .Faces(1).B = 1:        .Faces(1).C = 0
        .Faces(2).A = 4:        .Faces(2).B = 5:        .Faces(2).C = 7
        .Faces(3).A = 7:        .Faces(3).B = 6:        .Faces(3).C = 4
        .Faces(4).A = 0:        .Faces(4).B = 1:        .Faces(4).C = 5
        .Faces(5).A = 5:        .Faces(5).B = 4:        .Faces(5).C = 0
        .Faces(6).A = 1:        .Faces(6).B = 3:        .Faces(6).C = 7
        .Faces(7).A = 7:        .Faces(7).B = 5:        .Faces(7).C = 1
        .Faces(8).A = 3:        .Faces(8).B = 2:        .Faces(8).C = 6
        .Faces(9).A = 6:        .Faces(9).B = 7:        .Faces(9).C = 3
        .Faces(10).A = 2:       .Faces(10).B = 0:       .Faces(10).C = 4
        .Faces(11).A = 4:       .Faces(11).B = 6:       .Faces(11).C = 2
        
        .TVerts(0).U = 0:       .TVerts(0).V = 0
        .TVerts(1).U = 1:       .TVerts(1).V = 0
        .TVerts(2).U = 0:       .TVerts(2).V = 1
        .TVerts(3).U = 1:       .TVerts(3).V = 1
        .TVerts(4).U = 0:       .TVerts(4).V = 0
        .TVerts(5).U = 1:       .TVerts(5).V = 0
        .TVerts(6).U = 0:       .TVerts(6).V = 1
        .TVerts(7).U = 1:       .TVerts(7).V = 1
        .TVerts(8).U = 0:       .TVerts(8).V = 0
        .TVerts(9).U = 1:       .TVerts(9).V = 0
        .TVerts(10).U = 0:      .TVerts(10).V = 1
        .TVerts(11).U = 1:      .TVerts(11).V = 1
        
        .TFaces(0).A = 9:       .TFaces(0).B = 11:      .TFaces(0).C = 10
        .TFaces(1).A = 10:      .TFaces(1).B = 8:       .TFaces(1).C = 9
        .TFaces(2).A = 8:       .TFaces(2).B = 9:       .TFaces(2).C = 11
        .TFaces(3).A = 11:      .TFaces(3).B = 10:      .TFaces(3).C = 8
        .TFaces(4).A = 4:       .TFaces(4).B = 5:       .TFaces(4).C = 7
        .TFaces(5).A = 7:       .TFaces(5).B = 6:       .TFaces(5).C = 4
        .TFaces(6).A = 0:       .TFaces(6).B = 1:       .TFaces(6).C = 3
        .TFaces(7).A = 3:       .TFaces(7).B = 2:       .TFaces(7).C = 0
        .TFaces(8).A = 4:       .TFaces(8).B = 5:       .TFaces(8).C = 7
        .TFaces(9).A = 7:       .TFaces(9).B = 6:       .TFaces(9).C = 4
        .TFaces(10).A = 0:      .TFaces(10).B = 1:      .TFaces(10).C = 3
        .TFaces(11).A = 3:      .TFaces(11).B = 2:      .TFaces(11).C = 0
        
        ResetMeshParameters
        texWidth = dibTex.Width - 1
        texHeight = dibTex.Height - 1
        For idx = 0 To .numTVert
            .TVerts(idx).U = .TVerts(idx).U * texWidth
            .TVerts(idx).V = .TVerts(idx).V * texHeight
        Next
    End With
    
End Sub
