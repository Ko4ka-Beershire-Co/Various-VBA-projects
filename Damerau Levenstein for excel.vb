Public Function WeightedDL(source As String, target As String) As Double

    Dim deleteCost As Double
    Dim insertCost As Double
    Dim replaceCost As Double
    Dim swapCost As Double

    deleteCost = 1
    insertCost = 1
    replaceCost = 1
    swapCost = 1

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    If Len(source) = 0 Then
        WeightedDL = Len(target) * insertCost
        Exit Function
    End If

    If Len(target) = 0 Then
        WeightedDL = Len(source) * deleteCost
        Exit Function
    End If

    Dim table() As Double
    ReDim table(Len(source), Len(target))

    Dim sourceIndexByCharacter() As Variant
    ReDim sourceIndexByCharacter(0 To 1, 0 To Len(source) - 1) As Variant

    If Left(source, 1) <> Left(target, 1) Then
        table(0, 0) = Application.Min(replaceCost, (deleteCost + insertCost))
    End If

    sourceIndexByCharacter(0, 0) = Left(source, 1)
    sourceIndexByCharacter(1, 0) = 0

    Dim deleteDistance As Double
    Dim insertDistance As Double
    Dim matchDistance As Double

    For i = 1 To Len(source) - 1

        deleteDistance = table(i - 1, 0) + deleteCost
        insertDistance = ((i + 1) * deleteCost) + insertCost

        If Mid(source, i + 1, 1) = Left(target, 1) Then
            matchDistance = (i * deleteCost) + 0
        Else
            matchDistance = (i * deleteCost) + replaceCost
        End If

        table(i, 0) = Application.Min(Application.Min(deleteDistance, insertDistance), matchDistance)
    Next

    For j = 1 To Len(target) - 1

        deleteDistance = table(0, j - 1) + insertCost
        insertDistance = ((j + 1) * insertCost) + deleteCost

        If Left(source, 1) = Mid(target, j + 1, 1) Then
            matchDistance = (j * insertCost) + 0
        Else
            matchDistance = (j * insertCost) + replaceCost
        End If

        table(0, j) = Application.Min(Application.Min(deleteDistance, insertDistance), matchDistance)
    Next

    For i = 1 To Len(source) - 1

        Dim maxSourceLetterMatchIndex As Integer

        If Mid(source, i + 1, 1) = Left(target, 1) Then
            maxSourceLetterMatchIndex = 0
        Else
            maxSourceLetterMatchIndex = -1
        End If

        For j = 1 To Len(target) - 1

            Dim candidateSwapIndex As Integer
            candidateSwapIndex = -1

            For k = 0 To UBound(sourceIndexByCharacter, 2)
                If sourceIndexByCharacter(0, k) = Mid(target, j + 1, 1) Then candidateSwapIndex = sourceIndexByCharacter(1, k)
            Next

            Dim jSwap As Integer
            jSwap = maxSourceLetterMatchIndex

            deleteDistance = table(i - 1, j) + deleteCost
            insertDistance = table(i, j - 1) + insertCost
            matchDistance = table(i - 1, j - 1)

            If Mid(source, i + 1, 1) <> Mid(target, j + 1, 1) Then
                matchDistance = matchDistance + replaceCost
            Else
                maxSourceLetterMatchIndex = j
            End If

            Dim swapDistance As Double

            If candidateSwapIndex <> -1 And jSwap <> -1 Then

                Dim iSwap As Integer
                iSwap = candidateSwapIndex

                Dim preSwapCost
                If iSwap = 0 And jSwap = 0 Then
                    preSwapCost = 0
                Else
                    preSwapCost = table(Application.Max(0, iSwap - 1), Application.Max(0, jSwap - 1))
                End If

                swapDistance = preSwapCost + ((i - iSwap - 1) * deleteCost) + ((j - jSwap - 1) * insertCost) + swapCost

            Else
                swapDistance = 500
            End If

            table(i, j) = Application.Min(Application.Min(Application.Min(deleteDistance, insertDistance), matchDistance), swapDistance)

        Next

        sourceIndexByCharacter(0, i) = Mid(source, i + 1, 1)
        sourceIndexByCharacter(1, i) = i

    Next

    WeightedDL = table(Len(source) - 1, Len(target) - 1)

End Function
