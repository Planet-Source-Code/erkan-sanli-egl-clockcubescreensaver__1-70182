Attribute VB_Name = "modSort"
Option Explicit

Public Function VisibleFaces() As Integer
    
    Dim i       As Integer
    Dim iV      As Integer
    With Mesh1
        iV = -1
        Erase .FaceV
        Select Case Params.Mask
            Case vbChecked
                For i = 0 To .NumFaces
                    iV = iV + 1
                    ReDim Preserve .FaceV(iV)
                    .FaceV(iV).ZValue = (.VerticesT(.Faces(i).a).Z) + _
                                        (.VerticesT(.Faces(i).B).Z) + _
                                        (.VerticesT(.Faces(i).C).Z) + _
                                        (.VerticesT(.Faces(i).D).Z)
                    .FaceV(iV).iVisible = i
                Next
            Case Else
                For i = 0 To .NumFaces
                    'If (DotProduct(.Faces(i).NormalT, Camera) > 0) Then
                    If (.Faces(i).NormalT.Z > 0) Then
                        iV = iV + 1
                        ReDim Preserve .FaceV(iV)
                        .FaceV(iV).ZValue = (.VerticesT(.Faces(i).a).Z)
                        .FaceV(iV).iVisible = i
                    End If
                Next
        End Select
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
        With Mesh1
            MidIdx = (First + Last) \ 2
            MidVal = .FaceV(MidIdx).ZValue
            FirstIdx = First
            LastIdx = Last
            Do
                Do While .FaceV(FirstIdx).ZValue < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While .FaceV(LastIdx).ZValue > MidVal
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
