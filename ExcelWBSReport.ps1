
function generateMainForm {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    # Create main form
    $main_form = New-Object System.Windows.Forms.Form
    

    $main_form.Text = 'BAE AHR BHR Report'
    $main_form.Width = 822
    $main_form.Height = 506
    $main_form.AutoSize = $true
    $main_form.MaximizeBox = $false
    $main_form.MinimizeBox = $true
    $main_form.FormBorderStyle = 'FixedDialog'


    # Create Title
    $title = New-Object System.Windows.Forms.Label

    #Set properties
    $title.Text = 'BAE AHR BHR Report'
    $title.Font = 'Calibri, 16pt, style=Bold'
    $title.Size = '300,39'
    $title.Location = '12,9'

    #Add to form
    $main_form.Controls.Add($title)

    #Create first line break
    $titleLineBreak = New-Object System.Windows.Forms.Label

    #properties
    $titleLineBreak.Text = ''
    $titleLineBreak.AutoSize = $false
    $titleLineBreak.Location = '22,62'
    $titleLineBreak.BorderStyle = 'FixedSingle'
    $titleLineBreak.Size='750,2'

    $main_form.Controls.Add($titleLineBreak)

    # Create 'Step 1' label
    $step1Label = New-Object System.Windows.Forms.Label

    #properties
    $step1Label.Text = 'Step 1 - Select WBS Code Tracker'
    $step1Label.AutoSize = $true
    $step1Label.Font='Calibri,10pt'
    $step1Label.Location='21,75'
    $step1Label.Size='171,20'

    $main_form.Controls.Add($step1Label)

    # Create first 'Browse' text box to show the file the user has chosen
    $trackerTextBox = New-Object System.Windows.Forms.TextBox

    #properties
    $trackerTextBox.Location = '22,110'
    $trackerTextBox.Multiline = $true
    $trackerTextBox.Font = 'Calibri,12pt'
    $trackerTextBox.Size = '499,39'
    $trackerTextBox.ReadOnly = $true
    $trackerTextBox.BackColor = 'White'

    $main_form.Controls.Add($trackerTextBox)

    # Create 'Browse' button associated with the first text box
    $trackerBrowseButton = New-Object System.Windows.Forms.Button

    #properties
    $trackerBrowseButton.Location = '546,110'
    $trackerBrowseButton.Size = '169,39'
    $trackerBrowseButton.Text = 'Browse'
    $trackerBrowseButton.Font = 'Calibri,10pt'

    $main_form.Controls.Add($trackerBrowseButton)

    $trackerBrowseButton.Add_Click(
        {
            $trackerFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $trackerFileDialog.Filter = "Excel files (*.xls*) | *.xls*"
            $trackerFileDialog.ShowDialog() | Out-Null
            $trackerTextBox.Text = $trackerFileDialog.FileName
            
        }
    )

    # Create line break underneath step 1
    $step1LineBreak = New-Object System.Windows.Forms.Label

    #properties
    $step1LineBreak.Text = ''
    $step1LineBreak.AutoSize = $false
    $step1LineBreak.Location = '22,176'
    $step1LineBreak.BorderStyle = 'FixedSingle'
    $step1LineBreak.Size='750,2'

    $main_form.Controls.Add($step1LineBreak)

    # Create 'Step 2' label
    $step2Label = New-Object System.Windows.Forms.Label

    #properties
    $step2Label.Text = 'Step 2 - Select Weekly Report'
    $step2Label.AutoSize = $true
    $step2Label.Font='Calibri,10pt'
    $step2Label.Location='22,195'
    $step2Label.Size='171,20'

    $main_form.Controls.Add($step2Label)

    #Create second 'Browse' text box
    $reportTextBox = New-Object System.Windows.Forms.TextBox

    #properties
    $reportTextBox.Location = '22,230'
    $reportTextBox.Multiline = $true
    $reportTextBox.Font = 'Calibri,12pt'
    $reportTextBox.Size = '499,39'
    $reportTextBox.ReadOnly = $true
    $reportTextBox.BackColor = 'White'

    $main_form.Controls.Add($reportTextBox)

    # Create 'Browse' button associated with the second text box
    $reportBrowseButton = New-Object System.Windows.Forms.Button

    #properties
    $reportBrowseButton.Location = '546,230'
    $reportBrowseButton.Size = '169,39'
    $reportBrowseButton.Text = 'Browse'
    $reportBrowseButton.Font = 'Calibri,10pt'

    $main_form.Controls.Add($reportBrowseButton)

    $reportBrowseButton.Add_Click(
        {
            $reportFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $reportFileDialog.Filter = "Excel files (*.xls*) | *.xls*"
            $reportFileDialog.ShowDialog() | Out-Null
            $reportTextBox.Text = $reportFileDialog.FileName
        }
    )

    # Create line break underneath step 2
    $step2LineBreak = New-Object System.Windows.Forms.Label

    #properties
    $step2LineBreak.Text = ''
    $step2LineBreak.AutoSize = $false
    $step2LineBreak.Location = '22,297'
    $step2LineBreak.BorderStyle = 'FixedSingle'
    $step2LineBreak.Size='750,2'

    $main_form.Controls.Add($step2LineBreak)

    # Create 'Step 3' label
    $step3Label = New-Object System.Windows.Forms.Label

    #properties
    $step3Label.Text = 'Step 3 - Check PO''s Algorithm'
    $step3Label.AutoSize = $true
    $step3Label.Font='Calibri,10pt'
    $step3Label.Location='22,319'
    $step3Label.Size='171,20'

    $main_form.Controls.Add($step3Label)


    # Create button to run script
    $produceReport = New-Object System.Windows.Forms.Button

    #properties
    $produceReport.Text = 'Produce Report'
    $produceReport.AutoSize = $false
    $produceReport.Font='Calibri,10pt'
    $produceReport.Location='22,357'
    $produceReport.Size='167,42'

    $main_form.Controls.Add($produceReport)

    $progress = New-Object System.Windows.Forms.ProgressBar

    $progress.Location = '200,357'
    $progress.Size = '516,42'
    $progress.TabIndex = 4
    $main_form.Controls.Add($progress)
    $produceReport.Add_Click(
        {

            $result = [System.Windows.Forms.MessageBox]::Show('Are you sure you would like to generate the report?' , "Confirmation" , 4)
            if ($result -eq 'Yes') {
                if ([string]::IsNullOrEmpty($trackerTextBox.Text) -or [string]::IsNullOrEmpty($reportTextBox.Text)) {
                    [System.Windows.Forms.MessageBox]::Show('MISSING EXCEL SHEET: Please ensure both excel sheets are entered in steps 1 & 2.')
                } else {
                    
                    $progress.Value = 1

                    $progress.Maximum = 100
                    $progress.Step = 10
                    $progress.Value = 2

                    $produceReport.Enabled = $false
                    $produceReport.ImageIndex = 0

                    $trackerBrowseButton.Enabled = $false
                    $trackerBrowseButton.ImageIndex = 0

                    $reportBrowseButton.Enabled = $false
                    $reportBrowseButton.ImageIndex = 0
                   
                    mainReportGeneration $trackerTextBox.Text $reportTextBox.Text $progress
                    
                    $progress.Value = 0

                    $produceReport.Enabled = $true
                    $produceReport.ImageIndex = -1

                    $trackerBrowseButton.Enabled = $true
                    $trackerBrowseButton.ImageIndex = -1

                    $reportBrowseButton.Enabled = $true
                    $reportBrowseButton.ImageIndex = -1
                         
                }
            } 
            
        }
    ) # end producereport button click
   
    $helpButton = New-Object System.Windows.Forms.Button

    $helpButton.Location = '712,12'
    $helpButton.Size = '86,32'
    $helpButton.Text = 'Help'

    $main_form.Controls.Add($helpButton)

    $helpButton.Add_Click(
        {
            [System.Windows.Forms.MessageBox]::Show('Script is assuming the lookup values on the weekly spreadsheet are held in sheet number 4. Script takes approx ~40 seconds. If getting unexpected results ensure excel documents have been entered in the correct order. Any further problems contact rkitcher@dxc.com','Help')
        }
    )
    
    
    #Show
    $main_form.ShowDialog()
} #End generateMainForm

function addColumn {
    param($WBSCells)

    Try {
        #Add new column with date
        $WBSCells.item(1, $newCol ) = 'Remaining ' + (Get-Date -Format "dd/MM/yyyy")
        $WBSCells.item(1, $newCol ).Font.Size = 12
        $WBSCells.item(1, $newCol ).Font.Bold = $true
        $WBSCells.item(1, $newCol ).HorizontalAlignment = -4131
        $WBSCells.item(1, $newCol ).EntireColumn.Autofit() | Out-Null
        
        
    } Catch {
        #Catch any errors
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    
        displayErrorMessage $ErrorMessage $FailedItem

        
    }
    
} #End addColumn

function setBorders {
    $borderReference = ''
    $borderReference += convertColIndexToName($newCol)
    $borderReference += '1'

    $endBorderReference = ''
    $endBorderReference += convertColIndexToName($newCol)
    $endBorderReference += getLastUsedRow($wsWBSCodes)

    $borderRange = $wsWBSCodes.Range($borderReference, $endBorderReference)
    $borderRange.Borders.LineStyle = 1
    $borderRange.Borders.Weight = 2
}

#Search for PO in BAE sheet
function findPO {
    param($PO)
    Try {
        #If PO is empty or not a PO
        if ($PO -eq 'TBC' -or [string]::IsNullOrEmpty($PO)) {
        
        } else {
            $foundItem = $BAECells.Find($PO) #similar to CTRL+F function

            #if PO is in sheet
            if ($foundItem) {
                #write-Host $foundItem.Value2 'found at:' $foundItem.Row ',' $foundItem.Column
                return $True, $foundItem.Row    #return row of found PO     
            }
    }
    } Catch {
        #Catch any errors
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    
        displayErrorMessage $ErrorMessage $FailedItem

        
    }
    
} #End findPO

#Changes the index format of a column into the string equivalent 
function convertColIndexToName {
    Param($num)

    Try {
    
        $loopCount = 0
        $conversion = ''
        $alphabet = 'A', 'B', 'C', 'D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'

        #algorithm to convert numbers into letter format
        # e.g. 13 = 'M'
        # 27 = 'AA'
        # 39 = 'AM'
        while ($true) {
            if ($num -le 26) {
                break
            } else {
                $num -= 26
                $loopCount += 1
            }
        }

        if ($loopCount -gt 0) {
            $conversion += $alphabet[$loopCount - 1]
        } 

        #build string
        $conversion += $alphabet[$num - 1]


        return $conversion #return new format
    } Catch {
        #Catch any errors
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    
        displayErrorMessage $ErrorMessage $FailedItem

        
    }
} #End convertColIndexToName

function displayErrorMessage {
    param ($errorMessage, $failedItem)
    [System.Windows.Forms.MessageBox]::Show("Unexpected error occured: $FailedItem. Error Message $ErrorMessage")
    
}


#Returns last used column
function getLastUsedColumn {
    param($ws) 
    Try {
        return $ws.UsedRange.columns.count
    } Catch {
        #Catch any errors
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    
        displayErrorMessage $ErrorMessage $FailedItem

        
    }
} #End getLastUsedColumn

#Returns last used row
function getLastUsedRow {
    param($ws)
    Try {
        return $ws.UsedRange.rows.count
    } Catch {
        #Catch any errors
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    
        displayErrorMessage $ErrorMessage $FailedItem

        
    }
} #End getLastUsedRow

function mainReportGeneration {
    param($WBSFilepath, $BAEFilepath, $progress)

    Try {
        $progress.Value = 10
        #create new excel object
        $XL = New-Object -ComObject Excel.Application
        $XLworkbooks = $XL.Workbooks

        #Connect to excel files
        #$BAEFilepath = 'C:\Users\rkitcher.EAD\Documents\Automation project\New BAE AHR BHR Report Week 26 Summary FileV2\New BAE AHR BHR Report Week 26 Summary FileV2.xlsb'
        #$WBSFilepath = 'C:\Users\rkitcher.EAD\Documents\Automation project\WBS Codes 0407.xlsx'
        #$WBSFilepath = 144321

        #write-host 'WBS Filepath: ' $WBSFilepath 
        #write-host 'BAE Filepath: ' $BAEFilepath 

        #Init Workbooks
        $wbBAEReport = $XLworkbooks.Open($BAEFilepath)
        $wbWBSCodes = $XLworkbooks.Open($WBSFilepath)

        #Make Excel visible
        $XL.Visible = $false

        #Get sheets to avoid using two dots with COM objects
        $BAEsheets = $wbBAEReport.Worksheets
        $WBSsheets = $wbWBSCodes.Worksheets

        #Init worksheets
        $wsWBSCodes = $WBSsheets.Item(1)
        $wsBAEReport = $BAEsheets.Item((findCorrectSheet))

        #Get cells
        $WBSCells = $wsWBSCodes.Cells
        $BAECells = $wsBAEReport.Cells

        #New column index that we created
        $newCol = getLastUsedColumn($wsWBSCodes) 
        $newCol += 1

        $progress.Value = 20
        

        #Add new column
        addColumn($WBSCells)
        
        

        $progress.Value = 55
        #For each PO in the sheet
        For ($row = 2; $row -le (getLastUsedRow($wsWBSCodes)); $row++) {
    
            $cell = $WBSCells.Item($row, 2).Value2 #get PO value
            $isFound, $foundRow = findPO($cell) #Find PO from other workbook, if found, return true and row of found PO
            $isValueFound = $False 
            
            #if a PO was found in the other sheet
            if ($isFound) {
                #While the remaining to spend value is not found
                while (!$isValueFound) {
                    
                    $valueSpent = $BAECells.Item($foundRow, "BA").Value2 #get contents of cell
            
                    #if empty, increment row
                    if ([string]::IsNullOrEmpty($valueSpent)) {
                        $foundRow += 1

                    #if not empty, we want this value
                    } else {
                        #Write-Host 'Value: ' $valueSpent
                        $progress.Value = 75
                        #Create cell reference
                        $tempFromReference = ''
                        $tempFromReference += "BA"
                        $tempFromReference += $foundRow

                        #Create another cell reference
                        $tempToReference = ''
                        $tempToReference += convertColIndexToName($newCol) #convert col index to string
                        $tempToReference += $row

                
                        #write-host $tempReference

                        #Choose location to copy from
                        $copyFromRange = $wsBAEReport.Range($tempFromReference)
                        $copyFromRange.Copy() | Out-Null

                        #paste destination
                        $copyToRange = $wsWBSCodes.Range($tempToReference)
                        $wsWBSCodes.Paste($copyToRange)
                        $WBSCells.Item($row, $newCol).Interior.ColorIndex = 0
                        
                        

                        $isValueFound = $True #stop condition
                    }
                }
            }
            
            setBorders
            
    
        }
        $progress.Value = 90
        $XL.Visible = $true
        
    } catch {
        #Catch any errors
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    
        displayErrorMessage $ErrorMessage $FailedItem
        
    }

    
} #End mainReportGeneration

function hideConsoleWindow {
    
    # Hide PowerShell Console
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)

}

function findCorrectSheet {
    foreach ($singleSheet in $BAESheets) {

        ### COLUMNS MAY GET ADDED/REMOVED CHANGE CONDITION BELOW IF THIS IS THE CASE

        $sheetCells = $singleSheet.Cells
        try {
            if ($sheetCells.Item('1', 'BA').Value2 -eq 'Remaining Spend' -or $sheetCells.Item('1', 'BB').Value2 -eq 'Remaining Spend' -and $singleSheet.Name -ne '£1 Rates') {
                return $singleSheet.Name
            }
        } catch {
            continue
        }
    }
}

###
### MAIN CODE
###

Try {
    Add-Type -assembly System.Windows.Forms

    hideConsoleWindow

    #Generate main GUI
    generateMainForm

} Catch {

    #Catch any errors
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    
    displayErrorMessage $ErrorMessage $FailedItem
    
}

