Public Class frmMain

   Private mobjFormBitmap As Bitmap
   Private mobjBitmapGraphics As Graphics
   Private mintFormWidth As Int16
   Private mintFormHeight As Int16
   Private mintBoardWidth As Int16
   Private mintBoardUnits As Int16
   Private mintBoardLeft As Int16
   Private mintBoardTop As Int16
   Private mblnBarriers(,) As Boolean
   Private mintShortestPath(,) As Int16
   Private mintPathCost(,) As Int32
   Private mobjRandom As Random

   Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

      Static blnDoneOnce As Boolean = False

      If Not blnDoneOnce Then
         blnDoneOnce = True
         mintFormWidth = Me.Width
         mintFormHeight = Me.Height
         mobjFormBitmap = New Bitmap(mintFormWidth, mintFormHeight, Me.CreateGraphics())
         mobjBitmapGraphics = Graphics.FromImage(mobjFormBitmap)
         mobjBitmapGraphics.FillRectangle(Brushes.White, 0, 0, mintFormWidth, mintFormHeight)
         mintBoardWidth = 400
         mintBoardUnits = 20
         ReDim mblnBarriers(mintBoardUnits - 1, mintBoardUnits - 1)
         ReDim mintShortestPath((mintBoardUnits ^ 2) - 1, (mintBoardUnits ^ 2) - 1)
         ReDim mintPathCost((mintBoardUnits ^ 2) - 1, (mintBoardUnits ^ 2) - 1)
         DrawBoard()
         mobjRandom = New Random
      End If

   End Sub

   Private Sub frmMain_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

      e.Graphics.DrawImage(mobjFormBitmap, 0, 0)

   End Sub


   Private Sub DrawBoard()

      Dim intX As Int16
      Dim intY As Int16

      mintBoardLeft = (mintFormWidth / 2) - (mintBoardWidth / 2)
      mintBoardTop = (mintFormHeight / 2) - (mintBoardWidth / 2)

      mobjBitmapGraphics.DrawRectangle(Pens.Black, mintBoardLeft, mintBoardTop, mintBoardWidth, mintBoardWidth)
      For intLines As Int16 = 1 To mintBoardUnits - 1
         intX = mintBoardLeft + (mintBoardWidth * intLines / mintBoardUnits)
         intY = mintBoardTop + (mintBoardWidth * intLines / mintBoardUnits)
         mobjBitmapGraphics.DrawLine(Pens.Black, intX, mintBoardTop, intX, mintBoardTop + mintBoardWidth)
         mobjBitmapGraphics.DrawLine(Pens.Black, mintBoardLeft, intY, mintBoardLeft + mintBoardWidth, intY)
      Next

      Me.Invalidate()

   End Sub


   Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)

      ' to remove the flickering

   End Sub

   Private Sub frmMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown

      Dim intX As Int16
      Dim intY As Int16

      intX = Math.Floor((e.X - mintBoardLeft) / (mintBoardWidth / mintBoardUnits))
      intY = Math.Floor((e.Y - mintBoardTop) / (mintBoardWidth / mintBoardUnits))
      If intX >= 0 And intX < mintBoardUnits And intY >= 0 And intY < mintBoardUnits Then
         mblnBarriers(intX, intY) = True
         PaintBlock(intX, intY, Pens.Black)
      End If

   End Sub

   Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

      If e.KeyCode = Keys.Return Then
         BuildShortestPaths()
         FindShortestPath()
      End If

   End Sub

   Private Sub BuildShortestPaths()

      Dim intBlockFrom As Int16
      Dim intBlockTo As Int16
      Dim intBlockThrough As Int16
      Dim intCost1 As Int32
      Dim intCost2 As Int32
      Dim intCostTotal As Int32

      For intBlockFrom = 0 To (mintBoardUnits ^ 2) - 1
         For intBlockTo = 0 To (mintBoardUnits ^ 2) - 1
            mintPathCost(intBlockFrom, intBlockTo) = GetCost(intBlockFrom, intBlockTo)
            mintShortestPath(intBlockFrom, intBlockTo) = intBlockFrom
         Next
      Next

      For intBlockThrough = 0 To (mintBoardUnits ^ 2) - 1
         For intBlockFrom = 0 To (mintBoardUnits ^ 2) - 1
            intCost1 = mintPathCost(intBlockFrom, intBlockThrough)
            If intCost1 < 1000000 And intBlockFrom <> intBlockThrough Then
               For intBlockTo = 0 To (mintBoardUnits ^ 2) - 1
                  intCost2 = mintPathCost(intBlockThrough, intBlockTo)
                  intCostTotal = mintPathCost(intBlockFrom, intBlockTo)
                  If intCost2 < 1000000 And intBlockTo <> intBlockThrough Then
                     If intCost1 + intCost2 < intCostTotal Then
                        mintPathCost(intBlockFrom, intBlockTo) = intCost1 + intCost2
                        mintShortestPath(intBlockFrom, intBlockTo) = mintShortestPath(intBlockThrough, intBlockTo)
                     ElseIf intCost1 + intCost2 = intCostTotal And mobjRandom.NextDouble < 0.5 Then
                        mintPathCost(intBlockFrom, intBlockTo) = intCost1 + intCost2
                        mintShortestPath(intBlockFrom, intBlockTo) = mintShortestPath(intBlockThrough, intBlockTo)
                     End If
                  End If
               Next
            End If
         Next
      Next

   End Sub

   Private Function GetCost(ByVal pintNode1 As Int16, ByVal pintNode2 As Int16) As Int32

      Dim intX1 As Int16
      Dim intY1 As Int16
      Dim intX2 As Int16
      Dim intY2 As Int16

      intX1 = pintNode1 Mod mintBoardUnits
      intY1 = mintBoardUnits - Math.Floor(pintNode1 / mintBoardUnits) - 1
      intX2 = pintNode2 Mod mintBoardUnits
      intY2 = mintBoardUnits - Math.Floor(pintNode2 / mintBoardUnits) - 1

      If mblnBarriers(intX1, intY1) = False And mblnBarriers(intX2, intY2) = False Then
         If Math.Abs(intX1 - intX2) + Math.Abs(intY1 - intY2) = 1 Then
            Return 1
         Else
            Return 1000000
         End If
      Else
         Return 1000000
      End If

   End Function

   Private Sub FindShortestPath()

      Dim intCurrentNode As Int16

      intCurrentNode = (mintBoardUnits ^ 2) - 1
      PaintBlock(intCurrentNode Mod mintBoardUnits, mintBoardUnits - Math.Floor(intCurrentNode / mintBoardUnits) - 1, Pens.Blue)

      Do While intCurrentNode <> 0
         intCurrentNode = mintShortestPath(0, intCurrentNode)
         PaintBlock(intCurrentNode Mod mintBoardUnits, mintBoardUnits - Math.Floor(intCurrentNode / mintBoardUnits) - 1, Pens.Blue)
      Loop

   End Sub

   Private Sub PaintBlock(ByVal pintX As Int16, ByVal pintY As Int16, ByVal pobjPen As Pen)

      Dim intX1Coordinate As Int16
      Dim intX2Coordinate As Int16

      For intYCoordinate As Int16 = mintBoardTop + (mintBoardWidth * pintY / mintBoardUnits) + 1 To mintBoardTop + (mintBoardWidth * (pintY + 1) / mintBoardUnits) - 1
         intX1Coordinate = mintBoardLeft + (mintBoardWidth * pintX / mintBoardUnits) + 1
         intX2Coordinate = mintBoardLeft + (mintBoardWidth * (pintX + 1) / mintBoardUnits) - 1
         mobjBitmapGraphics.DrawLine(pobjPen, intX1Coordinate, intYCoordinate, intX2Coordinate, intYCoordinate)
         Me.Invalidate()
         Application.DoEvents()
      Next

   End Sub

End Class




