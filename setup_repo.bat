@echo off
setlocal enabledelayedexpansion
title SQCART Installation

:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [91mPlease run this installer as Administrator[0m
    echo Right-click the file and select "Run as administrator"
    pause
    exit /b 1
)

:: Set installation paths
set "INSTALL_DIR=%PROGRAMFILES%\SQCART"
set "DOCS_DIR=%USERPROFILE%\Documents\SQCART"

:: ASCII art and welcome message
cls
echo [92m
echo  _____ _____ _____ _____ _____ _____ 
echo ^|   __^|     ^|     ^|  _  ^|     ^|_   _^|
echo ^|__   ^|  ^|  ^|   --^|     ^|  ^|  ^| ^| ^| 
echo ^|_____^|__^|__^|_____^|__^|__^|_____^| ^|_^| 
echo [0m
echo Welcome to SQCART Installation
echo Supplier Quality Corrective Action Tool
echo.
echo This will install SQCART on your computer.
echo Installation directory: %INSTALL_DIR%
echo.
echo Press any key to begin installation or Ctrl+C to cancel...
pause >nul

:: Progress bar function
:ProgressBar
set /a "total=7"
set /a "current=0"
call :UpdateProgress

:: Create directories
set /a "current+=1"
echo Creating directories...
call :UpdateProgress
mkdir "%INSTALL_DIR%" 2>nul
mkdir "%DOCS_DIR%" 2>nul
mkdir "%DOCS_DIR%\Reports" 2>nul

:: Check Excel installation
set /a "current+=1"
echo Checking Excel installation...
call :UpdateProgress
powershell -Command "if (!(Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')) { Write-Host '[91mWarning: Excel 2016 or newer not found. Please install Microsoft Excel.[0m'; exit 1 }" >nul 2>&1

:: Create Excel template
set /a "current+=1"
echo Creating Excel template...
call :UpdateProgress
powershell -Command "& {
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    
    # Add sheets
    $inputSheet = $workbook.Sheets.Add()
    $inputSheet.Name = 'Supplier Input Form'
    $analysisSheet = $workbook.Sheets.Add()
    $analysisSheet.Name = 'AS9100 Compliance Check'
    $reportSheet = $workbook.Sheets.Add()
    $reportSheet.Name = 'Report'
    
    # Format Supplier Input Form
    $inputSheet.Activate()
    $inputSheet.Range('A1').Value = 'Supplier Information'
    $inputSheet.Range('A1').Font.Bold = $true
    $inputSheet.Range('A1').Font.Size = 14
    
    # Add dropdown lists and validation
    $inputSheet.Range('B5').Validation.Add(1, 1, 1, 'Quality,Manufacturing,Engineering')
    
    # Add VBA code
    $vbaCode = @'
    Option Explicit
    
    Private Sub Workbook_Open()
        Call InitializeSQCART
    End Sub
    
    Private Sub InitializeSQCART()
        On Error Resume Next
        
        ' Show welcome message
        MsgBox "Welcome to SQCART" & vbNewLine & _
               "Supplier Quality Corrective Action Tool" & vbNewLine & _
               "Version 1.0", vbInformation
               
        ' Initialize forms
        Call ClearAllForms
        Call SetupValidation
        
        ' Add Export to PDF button
        Call AddPDFButton
        
        Sheets("Supplier Input Form").Activate
        Range("A2").Select
    End Sub
    
    Private Sub ClearAllForms()
        Sheets("Supplier Input Form").Range("A2:Z100").ClearContents
        Sheets("AS9100 Compliance Check").Range("A2:Z100").ClearContents
    End Sub
    
    Private Sub SetupValidation()
        With Sheets("Supplier Input Form")
            .Range("TeamRoles").Validation.Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:="Team Lead,Quality Engineer,Process Engineer"
        End With
    End Sub

    Public Sub ExportToPDF()
        On Error GoTo ErrorHandler
        
        Dim filePath As String
        Dim defaultPath As String
        defaultPath = Environ$("USERPROFILE") & "\Documents\SQCART\Reports\"
        
        ' Create Reports folder if it doesn't exist
        If Dir(defaultPath, vbDirectory) = "" Then
            MkDir defaultPath
        End If
        
        ' Generate default filename with timestamp
        filePath = defaultPath & "SQCART_Report_" & _
                  Format(Now, "yyyy-mm-dd_HHmmss") & ".pdf"
        
        ' Show save dialog
        With Application.FileDialog(msoFileDialogSaveAs)
            .InitialFileName = filePath
            .FilterIndex = 1
            .Title = "Save SQCART Report as PDF"
            .Filters.Clear
            .Filters.Add "PDF Files", "*.pdf"
            
            If .Show = True Then
                filePath = .SelectedItems(1)
                
                ' Export active sheet to PDF
                ActiveSheet.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=filePath, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, _
                    OpenAfterPublish:=True
                
                MsgBox "Report exported successfully to:" & vbNewLine & filePath, _
                       vbInformation, "Export Complete"
            End If
        End With
        
        Exit Sub
        
ErrorHandler:
        MsgBox "Error exporting to PDF: " & Err.Description, _
               vbCritical, "Export Error"
    End Sub

    Private Sub AddPDFButton()
        On Error Resume Next
        
        ' Add a button to the Quick Access Toolbar
        With Application.CommandBars("Quick Access Toolbar")
            .Controls.Add(Type:=msoControlButton, _
                         Before:=1).OnAction = "ExportToPDF"
            With .Controls(1)
                .FaceId = 940  ' PDF icon
                .Caption = "Export to PDF"
                .TooltipText = "Export SQCART Report to PDF"
            End With
        End With
    End Sub
'@
    
    $vbaProject = $workbook.VBProject
    $vbaModule = $vbaProject.VBComponents.Add(1)
    $vbaModule.CodeModule.AddFromString($vbaCode)
    
    # Save and close
    $workbook.SaveAs("$env:INSTALL_DIR\SQCART_Template.xlsm", 52)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}"

:: Create shortcuts
set /a "current+=1"
echo Creating shortcuts...
call :UpdateProgress
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%USERPROFILE%\Desktop\SQCART.lnk'); $s.TargetPath = '%INSTALL_DIR%\SQCART_Template.xlsm'; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.Description = 'Supplier Quality Corrective Action Tool'; $s.Save()"
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\SQCART.lnk'); $s.TargetPath = '%INSTALL_DIR%\SQCART_Template.xlsm'; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.Description = 'Supplier Quality Corrective Action Tool'; $s.Save()"

:: Create uninstaller
set /a "current+=1"
echo Creating uninstaller...
call :UpdateProgress
(
    echo @echo off
    echo title SQCART Uninstaller
    echo echo Uninstalling SQCART...
    echo rmdir /S /Q "%INSTALL_DIR%"
    echo del "%USERPROFILE%\Desktop\SQCART.lnk"
    echo del "%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\SQCART.lnk"
    echo echo.
    echo echo SQCART has been uninstalled successfully.
    echo pause
) > "%INSTALL_DIR%\uninstall.bat"

:: Final setup
set /a "current+=1"
echo Finalizing installation...
call :UpdateProgress

:: Complete
set /a "current=%total%"
call :UpdateProgress
echo.
echo [92mInstallation completed successfully![0m
echo.
echo You can now:
echo 1. Use the desktop shortcut to launch SQCART
echo 2. Find SQCART in the Start Menu
echo 3. Access SQCART in: %INSTALL_DIR%
echo 4. Export reports to PDF using the new PDF button
echo.
echo Reports will be saved to: %DOCS_DIR%\Reports
echo.
echo Press any key to exit...
pause >nul
exit /b 0

:UpdateProgress
set /a "percent=100*%current%/%total%"
set "progress="
for /l %%i in (1,1,50) do (
    set /a "val=%%i*2"
    if !val! leq !percent! (
        set "progress=!progress!█"
    ) else (
        set "progress=!progress!░"
    )
)
echo [%progress%] !percent!%%
goto :eof