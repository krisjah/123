Sub CreatePenguinShow()
    ' This VBA macro creates and animates penguins walking in lockstep in Excel
    ' Instructions:
    ' 1. Open Excel and create a new workbook
    ' 2. Press Alt+F11 to open VBA Editor
    ' 3. Insert > Module
    ' 4. Paste this entire code
    ' 5. Press Alt+F11 to return to Excel
    ' 6. Press Alt+F8, select "CreatePenguinShow" and click Run

    Dim ws As Worksheet
    Dim penguins(1 To 5) As Shape
    Dim i As Integer

    ' Set up the worksheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Tab.Color = RGB(135, 206, 235) ' Sky blue

    ' Create title
    ws.Range("A1:Z1").Merge
    ws.Range("A1").Value = "🐧 PENGUINS WALKING IN LOCKSTEP 🐧"
    ws.Range("A1").Font.Size = 24
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").Interior.Color = RGB(176, 224, 230)

    ' Create ground
    ws.Range("A20:Z30").Interior.Color = RGB(240, 255, 255)
    ws.Range("A20:Z20").Borders(xlEdgeTop).Weight = xlThick
    ws.Range("A20:Z20").Borders(xlEdgeTop).Color = RGB(176, 224, 230)

    ' Delete any existing shapes
    For Each shp In ws.Shapes
        If Left(shp.Name, 7) = "Penguin" Then shp.Delete
    Next

    ' Create 5 penguins using shapes
    For i = 1 To 5
        Set penguins(i) = ws.Shapes.AddShape(msoShapeOval, _
            30 + (i - 1) * 80, 300, 50, 70)

        With penguins(i)
            .Name = "Penguin" & i
            .Fill.ForeColor.RGB = RGB(44, 62, 80) ' Dark blue-gray
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Weight = 2
        End With

        ' Add belly (white oval)
        Dim belly As Shape
        Set belly = ws.Shapes.AddShape(msoShapeOval, _
            penguins(i).Left + 10, penguins(i).Top + 20, 30, 40)
        With belly
            .Name = "Belly" & i
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Visible = msoFalse
        End With

        ' Add head (circle)
        Dim head As Shape
        Set head = ws.Shapes.AddShape(msoShapeOval, _
            penguins(i).Left + 5, penguins(i).Top - 30, 40, 40)
        With head
            .Name = "Head" & i
            .Fill.ForeColor.RGB = RGB(44, 62, 80)
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Weight = 2
        End With

        ' Add face (white oval)
        Dim face As Shape
        Set face = ws.Shapes.AddShape(msoShapeOval, _
            head.Left + 8, head.Top + 8, 24, 24)
        With face
            .Name = "Face" & i
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Visible = msoFalse
        End With

        ' Add beak (triangle)
        Dim beak As Shape
        Set beak = ws.Shapes.AddShape(msoShapeIsoscelesTriangle, _
            face.Left + 8, face.Top + 15, 8, 10)
        With beak
            .Name = "Beak" & i
            .Fill.ForeColor.RGB = RGB(255, 165, 0) ' Orange
            .Line.Visible = msoFalse
            .Rotation = 180
        End With

        ' Add eyes (small circles)
        Dim eyeL As Shape, eyeR As Shape
        Set eyeL = ws.Shapes.AddShape(msoShapeOval, _
            face.Left + 5, face.Top + 6, 4, 4)
        Set eyeR = ws.Shapes.AddShape(msoShapeOval, _
            face.Left + 15, face.Top + 6, 4, 4)

        With eyeL
            .Name = "EyeL" & i
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Visible = msoFalse
        End With

        With eyeR
            .Name = "EyeR" & i
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Visible = msoFalse
        End With

        ' Add feet
        Dim footL As Shape, footR As Shape
        Set footL = ws.Shapes.AddShape(msoShapeOval, _
            penguins(i).Left + 5, penguins(i).Top + 70, 15, 8)
        Set footR = ws.Shapes.AddShape(msoShapeOval, _
            penguins(i).Left + 30, penguins(i).Top + 70, 15, 8)

        With footL
            .Name = "FootL" & i
            .Fill.ForeColor.RGB = RGB(255, 165, 0)
            .Line.Visible = msoFalse
        End With

        With footR
            .Name = "FootR" & i
            .Fill.ForeColor.RGB = RGB(255, 165, 0)
            .Line.Visible = msoFalse
        End With
    Next i

    MsgBox "Penguins created! Now run 'AnimatePenguins' to see them walk in lockstep!", vbInformation
End Sub

Sub AnimatePenguins()
    ' Animates the penguins walking across the screen in lockstep
    Dim ws As Worksheet
    Dim i As Integer, step As Integer
    Dim allShapes As Collection
    Dim moveDistance As Single
    Dim waddleOffset As Single

    Set ws = ActiveSheet
    moveDistance = 5 ' pixels to move each step

    ' Animate for 100 steps
    For step = 1 To 100
        ' Move all penguin parts together
        For i = 1 To 5
            ' Calculate waddle effect (alternating up/down)
            If step Mod 2 = 0 Then
                waddleOffset = 2
            Else
                waddleOffset = -2
            End If

            ' Move each penguin's components
            On Error Resume Next
            ws.Shapes("Penguin" & i).IncrementLeft moveDistance
            ws.Shapes("Penguin" & i).IncrementTop waddleOffset

            ws.Shapes("Belly" & i).IncrementLeft moveDistance
            ws.Shapes("Belly" & i).IncrementTop waddleOffset

            ws.Shapes("Head" & i).IncrementLeft moveDistance
            ws.Shapes("Head" & i).IncrementTop waddleOffset

            ws.Shapes("Face" & i).IncrementLeft moveDistance
            ws.Shapes("Face" & i).IncrementTop waddleOffset

            ws.Shapes("Beak" & i).IncrementLeft moveDistance
            ws.Shapes("Beak" & i).IncrementTop waddleOffset

            ws.Shapes("EyeL" & i).IncrementLeft moveDistance
            ws.Shapes("EyeL" & i).IncrementTop waddleOffset

            ws.Shapes("EyeR" & i).IncrementLeft moveDistance
            ws.Shapes("EyeR" & i).IncrementTop waddleOffset

            ws.Shapes("FootL" & i).IncrementLeft moveDistance
            ws.Shapes("FootL" & i).IncrementTop waddleOffset

            ws.Shapes("FootR" & i).IncrementLeft moveDistance
            ws.Shapes("FootR" & i).IncrementTop waddleOffset
            On Error GoTo 0
        Next i

        ' Small delay to make animation visible
        Application.Wait Now + TimeValue("00:00:00.05")
        DoEvents
    Next step

    MsgBox "The penguins have finished their march! Click 'ResetPenguins' to start over.", vbInformation
End Sub

Sub ResetPenguins()
    ' Reset penguin positions to start
    Application.ScreenUpdating = False
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Application.ScreenUpdating = True
    CreatePenguinShow
End Sub
