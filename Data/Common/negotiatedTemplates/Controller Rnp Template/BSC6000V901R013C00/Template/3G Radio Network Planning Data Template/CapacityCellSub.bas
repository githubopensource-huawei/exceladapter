Attribute VB_Name = "CapacityCellSub"






Public Sub createDelFreqBar()
    Dim delChooseBar As CommandBar
    Dim delFreqStyle As CommandBarButton
    Dim BatchdelFreqStyle As CommandBarButton
    Dim BatchdelChooseBar As CommandBar
    
    'Call deleteCellBar
    
    Dim actSheetName As String
    actSheetName = ThisWorkbook.ActiveSheet.Name
    CELL_SHEET_NAME = actSheetName
    CELL_TYPE = cellSheetType(actSheetName)
      
      If CELL_TYPE = 0 Then
            Set delChooseBar = Application.CommandBars.Add(DELFREQ_BAR_NAME, msoBarTop)
            Dim delbarDescLbl As String
            delbarDescLbl = "DeleteFreq"
            With delChooseBar
               .Protection = msoBarNoResize
               .Visible = True
               Set delFreqStyle = .Controls.Add(Type:=msoControlButton)
               With delFreqStyle
                    .Style = msoButtonIconAndCaption
                    .caption = getResByKey(delbarDescLbl)
                    .TooltipText = getResByKey(delbarDescLbl)
                    .OnAction = "deleteFrequency"
                    .FaceId = 186
                    .Enabled = True
                End With
              End With
                        Set BatchdelChooseBar = Application.CommandBars.Add(BATCH_DELFREQ_BAR_NAME, msoBarTop)
                        Dim BatchdelbarDescLbl As String
                        BatchdelbarDescLbl = "BatchDeleteFreq"
                        With BatchdelChooseBar
                                        .Protection = msoBarNoResize
                                        .Visible = True
                                        Set BatchdelFreqStyle = .Controls.Add(Type:=msoControlButton)
                                        With BatchdelFreqStyle
                                                .Style = msoButtonIconAndCaption
                                                .caption = getResByKey(BatchdelbarDescLbl)
                                                .TooltipText = getResByKey(BatchdelbarDescLbl)
                                                .OnAction = "BatchdeleteFrequency"
                                                .FaceId = 186
                                                .Enabled = True
                                        End With
                        End With
          
        End If
End Sub

