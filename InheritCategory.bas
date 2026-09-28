        On Error Resume Next
        objTable.Columns.Add "Categories"
        On Error GoTo ErrorHandler
        
        ' Sort to get the most recently updated categories first (optional but good practice)
        objTable.Sort "[ReceivedTime]", True
        
        ' Move through the conversation tree
        Do Until objTable.EndOfTable
            Set objRow = objTable.GetNextRow()
            
            ' Check if this row has data in the Categories column
            If Not IsNull(objRow("Categories")) Then
                If objRow("Categories") <> "" Then
                    
                    ' 6. EXCLUDE THE CURRENT MESSAGE
                    ' Don't copy from itself (protects against timing/race conditions)
                    If objRow("EntryID") <> objNewMail.EntryID Then
                        strCategory = objRow("Categories")
                        Exit Do ' We found our category, stop searching!
                    End If
                    
                End If
            End If
        Loop
        
        ' 8. TIMING & RACE CONDITION PROTECTION
        ' Apply the category and explicitly Save.
        ' If a Rule moves the item simultaneously, the ErrorHandler catches the lock conflict gracefully.
        If strCategory <> "" Then
            objNewMail.Categories = strCategory
            objNewMail.Save
        End If
    End If

ExitRoutine:
    ' Clean up objects
    Set objRow = Nothing
    Set objTable = Nothing
    Set objConversation = Nothing
    Exit Sub

ErrorHandler:
    ' If Outlook Rules moved the message or an object locked out, exit quietly without throwing a debug error to the user.
    Resume ExitRoutine
End Sub
