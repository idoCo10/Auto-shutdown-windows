# Script for Shutting down Windows after 3 hours of inactivity (mouse&keyboard movements) while saving Word, Excel and Notepad documents:

# For automate the task creation run the script "Create task Auto Shutdown.ps1" as administrator and then run the task from task schedualer.
# If you don't have the script you can follow the instructions bellow.


# To allow Powershell scripts run one time this command as Administrator:
# Set-ExecutionPolicy RemoteSigned -Force

# To manually create the task - Open Task Schedualer:
# Create Task
# General: Choose "Run only when user is logged on", Check "Run with highest privileges".
# Triggers: Create New, Begin the task - At log on, Choose "Any User", Check "Enable".
# Action: Create New, Choose "Start a program", Program/script: powershell.exe, Add arguments (optional): -WindowStyle Hidden -ExecutionPolicy Bypass -File C:\Users\Tech\Documents\Auto-Shutdown.ps1 (Don't forget to change the Username).
# Conditions: uncheck all.
# Settings: Check only "Allow task to be run on demand".
# Run the task.








# Function to get the last input time
function Get-LastInputTime {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class IdleTime {
        [DllImport("user32.dll")]
        public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
        public struct LASTINPUTINFO {
            public uint cbSize;
            public uint dwTime;
        }
        public static uint GetIdleTime() {
            LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
            GetLastInputInfo(ref lastInputInfo);
            return ((uint)Environment.TickCount - lastInputInfo.dwTime) / 1000;
        }
    }
"@

    [IdleTime]::GetIdleTime()
}

# Ensure Windows Forms Assembly is loaded for SendKeys
Add-Type -AssemblyName System.Windows.Forms


# Function to save open Word files
function Save-OpenWordDocuments {
    # Ensure the necessary types are available
    Add-Type -AssemblyName Microsoft.Office.Interop.Word

    try {
        # Get the running Word application
        $wordApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")

        if ($wordApp) {
            Write-Host "Saving open Word documents..."

            # Iterate over each open document and save
            foreach ($document in $wordApp.Documents) {
                try {
                    # Save the document with its existing filename
                    $document.Save() 
                    Write-Host "Saved document: $($document.Name)"
                } catch {
                    Write-Host "Failed to save document: $($_.Exception.Message)"
                }
            }
            
            # Optionally close Word after saving
            $wordApp.Quit()  # Uncomment this if you want to close Word as well
        } else {
            Write-Host "No running instance of Word found."
        }
    } catch {
        Write-Host "Error accessing Word: $($_.Exception.Message)"
    }
}

# Function to save open Excle files
function Save-OpenExcelFiles {
    # Ensure the necessary types are available
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel

    try {
        # Get the running Excel application
        $excelApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")

        if ($excelApp) {
            Write-Host "Saving open Excel workbooks..."

            # Iterate over each open workbook and save
            foreach ($workbook in $excelApp.Workbooks) {
                try {
                    $workbook.Save()  # Save the workbook
                } catch {
                    Write-Host "Failed to save workbook: $($_.Exception.Message)"
                }
            }
            
            # Optionally close Excel after saving
            $excelApp.Quit()  # close Excel
        } else {
            Write-Host "No running instance of Excel found."
        }
    } catch {
        Write-Host "Error accessing Excel: $($_.Exception.Message)"
    }
}


# Function to save open Notepad documents
function Save-OpenNotepadFiles {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WindowsInterop {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetForegroundWindow();
    }
"@

    # Get all running Notepad processes
    $notepadApps = Get-Process notepad -ErrorAction SilentlyContinue
    if ($notepadApps) {
        Write-Host "Saving open Notepad text documents..."

        foreach ($notepadApp in $notepadApps) {
            $windowHandle = $notepadApp.MainWindowHandle

            if ($windowHandle -ne [IntPtr]::Zero) {
                # Bring the Notepad window to the foreground
                [WindowsInterop]::SetForegroundWindow($windowHandle)
                
                # Give a short delay to ensure the window is focused
                Start-Sleep -Milliseconds 500

                # Send CTRL+S to save the document
                [System.Windows.Forms.SendKeys]::SendWait("^s")
                
                # Optionally wait for a brief period to handle potential save dialogs
                Start-Sleep -Seconds 1
            }
        }
    }
}



# Monitor for inactivity for 3 hours (10800 seconds)
$timeout = 10800

while ($true) {
    $idleTime = Get-LastInputTime
    if ($idleTime -ge $timeout) {
        Write-Host "No activity detected for 3 hours. Saving documents and shutting down..."

        # Save open Excel and Notepad files
        Save-OpenWordDocuments
        Save-OpenExcelFiles
        Save-OpenNotepadFiles

        # Shut down the computer
        Stop-Computer -Force
        break
    }
    Start-Sleep -Seconds 1
}
