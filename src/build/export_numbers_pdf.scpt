on run argv
    if (count of argv) < 3 then
        error "Usage: osascript script.scpt <doc_filepath> <pdf_filepath> <has_pivot_table> [pivot_table_name] [metrics_table_name]"
    end if
    
    set docFilePath to item 1 of argv
    set pdfFilePath to item 2 of argv
    set hasPivotTable to item 3 of argv
    
    set pivotTableName to ""
    if (count of argv) >= 4 then
        set pivotTableName to item 4 of argv
    end if

    set metricsTableName to ""
    if (count of argv) >= 5 then
        set metricsTableName to item 5 of argv
    end if

    tell application "Numbers"
        activate
        set targetDoc to open POSIX file docFilePath
        
        -- Wait for the document window to finish loading
        set waitCounter to 0
        repeat until exists front window
            delay 0.2
            set waitCounter to waitCounter + 1
            if waitCounter > 50 then
                error "Timeout waiting for Numbers window to appear."
            end if
        end repeat
        
        -- Only execute the refresh logic if the document actually has a pivot table
        if hasPivotTable is "true" then
            tell targetDoc
                tell active sheet
                    -- Force formula recalculation if a metrics table was provided
                    if metricsTableName is not "" then
                        try
                            tell table metricsTableName
                                repeat with i from 1 to count of cells
                                    try
                                        set f to formula of cell i
                                        if f is not missing value then
                                            -- Rewriting the formula triggers the calculation engine
                                            set value of cell i to f
                                        end if
                                    end try
                                end repeat
                            end tell
                        end try
                    end if
                    
                    -- Select a cell inside the Pivot Table so the Refresh menu item is enabled
                    if pivotTableName is not "" then
                        try
                            set selection range of table pivotTableName to range "A1" of table pivotTableName
                        end try
                    end if
                end tell
            end tell
        end if
    end tell

    if hasPivotTable is "true" then
        -- Trigger 'Refresh Pivot Table' from the Organise/Organize menu
        tell application "System Events"
            tell process "Numbers"
                set frontmost to true
                if exists menu bar item "Organise" of menu bar 1 then
                    set orgMenu to menu 1 of menu bar item "Organise" of menu bar 1
                else
                    set orgMenu to menu 1 of menu bar item "Organize" of menu bar 1
                end if
                
                -- Wait briefly for Numbers to enable the menu item after selection
                repeat 15 times
                    if enabled of menu item "Refresh Pivot Table" of orgMenu then
                        click menu item "Refresh Pivot Table" of orgMenu
                        exit repeat
                    end if
                    delay 0.2
                end repeat
            end tell
        end tell

        -- Allow the pivot table layout and calculation engine to settle before export
        delay 0.5
    end if

    tell application "Numbers"
        export front document to POSIX file pdfFilePath as PDF
        close front document saving no
        quit
    end tell
end run