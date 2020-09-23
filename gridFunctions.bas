Attribute VB_Name = "gridFunctions"

Public Sub AutosizeGridColumns(ByRef msFG As MSFlexGrid, ByVal MaxRowsToParse As Integer, ByVal MaxColWidth As Integer, frm As Form)
    Dim I, J As Integer
    Dim txtString As String
    Dim intTempWidth, BiggestWidth As Integer
    Dim intRows As Integer
    Const intPadding = 150


    With msFG


        For I = 0 To .Cols - 1
            ' Loops through every column
            .Col = I
            ' Set the active colunm
            intRows = .Rows
            ' Set the number of rows
            If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
            ' If there are more rows of data, reset
            ' intRows to the MaxRowsToParse constant
            '
            
            intBiggestWidth = 0
            ' Reset some values to 0


            For J = 0 To intRows - 1
                ' check up to MaxRowsToParse # of rows a
                '     nd obtain
                ' the greatest width of the cell content
                '     s
                
                .Row = J
      
                txtString = .Text
                intTempWidth = frm.TextWidth(txtString) + intPadding
                ' The intPadding constant compensates fo
                '     r text insets
                ' You can adjust this value above as des
                '     ired.
                
                If intTempWidth > intBiggestWidth Then intBiggestWidth = intTempWidth
                ' Reset intBiggestWidth to the intMaxCol
                '     Width value if necessary
            Next J
            .ColWidth(I) = intBiggestWidth
        Next I
        ' Now check to see if the columns aren't
        '     as wide as the grid itself.
        ' If not, determine the difference and e
        '     xpand each column proportionately
        ' to fill the grid
        intTempWidth = 0
        


        For I = 0 To .Cols - 1
            intTempWidth = intTempWidth + .ColWidth(I)
            ' Add up the width of all the columns
        Next I
        


        If intTempWidth < msFG.Width Then
            ' Compate the width of the columns to th
            '     e width of the grid control
            ' and if necessary expand the columns.
            intTempWidth = Fix((msFG.Width - intTempWidth) / .Cols)
            ' Determine the amount od width expansio
            '     n needed by each column


            For I = 0 To .Cols - 1
                .ColWidth(I) = .ColWidth(I) + intTempWidth
                ' add the necessary width to each column
                '
                
            Next I
        End If
    End With
End Sub
