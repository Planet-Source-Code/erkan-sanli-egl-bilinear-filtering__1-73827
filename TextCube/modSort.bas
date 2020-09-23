Attribute VB_Name = "modSort"
Option Explicit

Public Function VisibleFaces() As Integer
    
    Dim i   As Integer
    Dim iV  As Integer
    
    With Meshs
        iV = -1
        Erase .FaceV
        For i = 0 To .numFace
             If FaceVisible(.Vertices(.Faces(i).A).VectorsT, _
                            .Vertices(.Faces(i).B).VectorsT, _
                            .Vertices(.Faces(i).C).VectorsT) Then
                iV = iV + 1
                ReDim Preserve .FaceV(iV)
                .FaceV(iV).Value = (.Vertices(.Faces(i).A).VectorsT.Z + _
                                     .Vertices(.Faces(i).B).VectorsT.Z + _
                                     .Vertices(.Faces(i).C).VectorsT.Z)
                .FaceV(iV).Index = i
            End If
        Next
        If iV > -1 Then SortFaces 0, iV
        VisibleFaces = iV
    End With

End Function

Private Sub SortFaces(ByVal First As Integer, ByVal Last As Integer)

    Dim FirstIdx    As Integer
    Dim MidIdx      As Integer
    Dim LastIdx     As Integer
    Dim MidVal      As Single
    Dim TempOrder   As ORDER
    
    If (First < Last) Then
        With Meshs
            MidIdx = (First + Last) * 0.5
            MidVal = .FaceV(MidIdx).Value
            FirstIdx = First
            LastIdx = Last
            Do
                Do While .FaceV(FirstIdx).Value < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While .FaceV(LastIdx).Value > MidVal
                    LastIdx = LastIdx - 1
                Loop
                If (FirstIdx <= LastIdx) Then
                    TempOrder = .FaceV(LastIdx)
                    .FaceV(LastIdx) = .FaceV(FirstIdx)
                    .FaceV(FirstIdx) = TempOrder
                    FirstIdx = FirstIdx + 1
                    LastIdx = LastIdx - 1
                End If
            Loop Until FirstIdx > LastIdx

            If (LastIdx <= MidIdx) Then
                SortFaces First, LastIdx
                SortFaces FirstIdx, Last
            Else
                SortFaces FirstIdx, Last
                SortFaces First, LastIdx
            End If
        End With
    End If

End Sub
