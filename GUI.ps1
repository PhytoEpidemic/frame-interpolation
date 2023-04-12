Add-Type -AssemblyName System.Windows.Forms

function IsProcessRunning {
	param([System.Diagnostics.Process]$process)


	if ($null -ne $process) {
		try {
			$runningProcess = Get-Process -Id $process.Id -ErrorAction Stop
			return $true
		}
		catch {
			return $false
		}
	}

	return $false
}


Function ChooseFolder($Message) {
    $FolderBrowse = New-Object System.Windows.Forms.OpenFileDialog -Property @{ValidateNames = $false;CheckFileExists = $false;RestoreDirectory = $true;FileName = $Message;}
    $null = $FolderBrowse.ShowDialog()
    $FolderName = Split-Path -Path $FolderBrowse.FileName
    
	return $FolderName
}

function ConvertImageTo32Bit($inputFile, $outputFile) {
    $originalImage = [System.Drawing.Image]::FromFile($inputFile)
    $imageWidth = $originalImage.Width
    $imageHeight = $originalImage.Height

    $newImage = New-Object System.Drawing.Bitmap($imageWidth, $imageHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($newImage)
    $graphics.DrawImage($originalImage, 0, 0, $imageWidth, $imageHeight)
    
    $newImage.Save($outputFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $originalImage.Dispose()
    $newImage.Dispose()
    $graphics.Dispose()
}

function SetBlackPixelsToTransparent($inputFile, $outputFile) {
    $originalImage = [System.Drawing.Image]::FromFile($inputFile)
    $imageWidth = $originalImage.Width
    $imageHeight = $originalImage.Height

    $newImage = New-Object System.Drawing.Bitmap($imageWidth, $imageHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($newImage)
    $graphics.DrawImage($originalImage, 0, 0, $imageWidth, $imageHeight)

    $blackColor = [System.Drawing.Color]::FromArgb(255, 10, 10, 10)
    $transparentColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 0)

    for ($x = 0; $x -lt $imageWidth; $x++) {
        for ($y = 0; $y -lt $imageHeight; $y++) {
            $pixelColor = $newImage.GetPixel($x, $y)

            if ($pixelColor.A -le $blackColor.A -and
                $pixelColor.R -le $blackColor.R -and
                $pixelColor.G -le $blackColor.G -and
                $pixelColor.B -le $blackColor.B) {

                $newImage.SetPixel($x, $y, $transparentColor)
            }
        }
    }


    $newImage.Save($outputFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $originalImage.Dispose()
    $newImage.Dispose()
    $graphics.Dispose()
}


function SetTransparentPixelsToBlack($inputFile, $outputFile) {
    $originalImage = [System.Drawing.Image]::FromFile($inputFile)
    $imageWidth = $originalImage.Width
    $imageHeight = $originalImage.Height

    $newImage = New-Object System.Drawing.Bitmap($imageWidth, $imageHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($newImage)
    $graphics.DrawImage($originalImage, 0, 0, $imageWidth, $imageHeight)

    $blackColor = [System.Drawing.Color]::FromArgb(0, 255, 255, 255)
    $transparentColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 0)

    for ($x = 0; $x -lt $imageWidth; $x++) {
        for ($y = 0; $y -lt $imageHeight; $y++) {
            $pixelColor = $newImage.GetPixel($x, $y)

            if ($pixelColor.A -le $blackColor.A -and
                $pixelColor.R -le $blackColor.R -and
                $pixelColor.G -le $blackColor.G -and
                $pixelColor.B -le $blackColor.B) {

                $newImage.SetPixel($x, $y, $transparentColor)
            }
        }
    }


    $newImage.Save($outputFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $originalImage.Dispose()
    $newImage.Dispose()
    $graphics.Dispose()
}

function SetBlackPixelsToNotBlack($inputFile, $outputFile) {
    $originalImage = [System.Drawing.Image]::FromFile($inputFile)
    $imageWidth = $originalImage.Width
    $imageHeight = $originalImage.Height

    $newImage = New-Object System.Drawing.Bitmap($imageWidth, $imageHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($newImage)
    $graphics.DrawImage($originalImage, 0, 0, $imageWidth, $imageHeight)

    $blackColor = [System.Drawing.Color]::FromArgb(255, 10, 10, 10)
    $transparentColor = [System.Drawing.Color]::FromArgb(255, 21, 21, 21)

    for ($x = 0; $x -lt $imageWidth; $x++) {
        for ($y = 0; $y -lt $imageHeight; $y++) {
            $pixelColor = $newImage.GetPixel($x, $y)

            if ($pixelColor.A -eq $blackColor.A -and
                $pixelColor.R -le $blackColor.R -and
                $pixelColor.G -le $blackColor.G -and
                $pixelColor.B -le $blackColor.B) {

                $newImage.SetPixel($x, $y, $transparentColor)
            }
        }
    }


    $newImage.Save($outputFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $originalImage.Dispose()
    $newImage.Dispose()
    $graphics.Dispose()
}

function ProcessImageFolder($folderPath) {
    $imageFiles = Get-ChildItem -Path $folderPath -Filter "*.png"

    foreach ($imageFile in $imageFiles) {
        $inputFile = $imageFile.FullName
        $outputFile = [System.IO.Path]::Combine($folderPath, "temp_" + $imageFile.Name)

        SetBlackPixelsToTransparent $inputFile $outputFile

        Remove-Item $inputFile
        Rename-Item $outputFile $inputFile
    }
}

function ProcessImageFolderFixBackbround($folderPath) {
    $imageFiles = Get-ChildItem -Path $folderPath -Filter "*.png"

    foreach ($imageFile in $imageFiles) {
        $inputFile = $imageFile.FullName
        $outputFile = [System.IO.Path]::Combine($folderPath, "temp_" + $imageFile.Name)

        SetTransparentPixelsToBlack $inputFile $outputFile

        Remove-Item $inputFile
        Rename-Item $outputFile $inputFile
    }
}

function ProcessImageFolderRemoveFullBlack($folderPath) {
    $imageFiles = Get-ChildItem -Path $folderPath -Filter "*.png"

    foreach ($imageFile in $imageFiles) {
        $inputFile = $imageFile.FullName
        $outputFile = [System.IO.Path]::Combine($folderPath, "temp_" + $imageFile.Name)

        SetBlackPixelsToNotBlack $inputFile $outputFile

        Remove-Item $inputFile
        Rename-Item $outputFile $inputFile
    }
}


# Create picture box
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBox.Size = New-Object System.Drawing.Size(640,480)
$pictureBox.Location = New-Object System.Drawing.Point(20,80)
#$panel.Controls.Add($pictureBox)



# Load sprite sheet image
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Image Files (*.png, *.jpg, *.bmp)|*.png;*.jpg;*.bmp"
$openFileDialog.Title = "Select a Sprite Sheet"
$global:ImageObject = ""


# Create numeric up-down controls for columns and rows
$columnsLabel = New-Object System.Windows.Forms.Label
$columnsLabel.Text = 'Columns:'
$columnsLabel.Location = New-Object System.Drawing.Point(660,10)
$columnsLabel.AutoSize = $true

$rowsLabel = New-Object System.Windows.Forms.Label
$rowsLabel.Text = 'Rows:'
$rowsLabel.Location = New-Object System.Drawing.Point(660,50)
$rowsLabel.AutoSize = $true

$columnsNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$columnsNumericUpDown.Location = New-Object System.Drawing.Point(720,10)
$columnsNumericUpDown.Minimum = 1

$rowsNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$rowsNumericUpDown.Location = New-Object System.Drawing.Point(720,50)
$rowsNumericUpDown.Minimum = 1


# Draw grid function
function DrawGrid($columns, $rows) {
    $originalImage = $global:ImageObject.Clone() # Create a clone of the original image to prevent modifying the original

    $graphics = [System.Drawing.Graphics]::FromImage($originalImage)

    # Set grid line color to semi-transparent black
    $gridColor = [System.Drawing.Color]::FromArgb(128, [System.Drawing.Color]::Black)
    $pen = New-Object System.Drawing.Pen($gridColor)
    $pen.Width = 3
    for ($i = 1; $i -lt $columns; $i++) {
        $x = ($originalImage.Width / $columns) * $i
        $graphics.DrawLine($pen, $x, 0, $x, $originalImage.Height)
    }

    for ($i = 1; $i -lt $rows; $i++) {
        $y = ($originalImage.Height / $rows) * $i
        $graphics.DrawLine($pen, 0, $y, $originalImage.Width, $y)
    }

    $pictureBox.Image = $originalImage
}





# Create apply button
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = 'Save Images'
$applyButton.AutoSize = $true
$applyButton.Location = New-Object System.Drawing.Point(660,90)
$applyButton.Add_Click({
    $columns = $columnsNumericUpDown.Value
    $rows = $rowsNumericUpDown.Value
    $OutputFolder = ChooseFolder "Output folder"
    if ($OutputFolder) {
        CropImageWithFFmpeg $global:ImageFile ($OutputFolder) $columns $rows
        ProcessImageFolderFixBackbround $OutputFolder
        ProcessImageFolderRemoveFullBlack $OutputFolder
    }
    
})



# Event handler for columnsNumericUpDown value change
$columnsNumericUpDown_ValueChanged = {
    $columns = $columnsNumericUpDown.Value
    $rows = $rowsNumericUpDown.Value
    DrawGrid $columns $rows
}

# Event handler for rowsNumericUpDown value change
$rowsNumericUpDown_ValueChanged = {
    $columns = $columnsNumericUpDown.Value
    $rows = $rowsNumericUpDown.Value
    DrawGrid $columns $rows
}

# Add event handlers to columnsNumericUpDown and rowsNumericUpDown
$columnsNumericUpDown.add_ValueChanged($columnsNumericUpDown_ValueChanged)
$rowsNumericUpDown.add_ValueChanged($rowsNumericUpDown_ValueChanged)




function IsAnaconda3Installed() {
    $condaExeName = "conda"
    $defaultInstallPaths = @(
        "${Env:UserProfile}\Anaconda3\Scripts",
        "${Env:UserProfile}\AppData\Local\Continuum\anaconda3\Scripts"
    )

    # Check if conda is in the system PATH
    $condaInPath = (Get-Command -ErrorAction SilentlyContinue $condaExeName) -ne $null

    if ($condaInPath) {
        return $true
    }

    # Check default installation directories
    foreach ($installPath in $defaultInstallPaths) {
        $condaExePath = Join-Path $installPath $condaExeName
        if (Test-Path $condaExePath) {
            return $true
        }
    }

    return $false
}
function AddControlsToForm($form, $controls) {
    foreach ($control in $controls) {
        $form.Controls.Add($control)
    }
}

function RemoveControlsFromForm($form, $controls) {
    foreach ($control in $controls) {
        $form.Controls.Remove($control)
    }
}

function CropImageWithFFmpeg($inputFile, $outputPath, $columns, $rows) {
    $outputDir = [System.IO.Path]::GetDirectoryName($outputPath)
    if (-not (Test-Path $outputPath)) {
        New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
    }

    $image = [System.Drawing.Image]::FromFile($inputFile)
    $imageWidth = $image.Width
    $imageHeight = $image.Height

    $tileWidth = [int]($imageWidth / $columns)
    $tileHeight = [int]($imageHeight / $rows)

    $processes = @()
    for ($row = 0; $row -lt $rows; $row++) {
        for ($column = 0; $column -lt $columns; $column++) {
            $x = $tileWidth * $column
            $y = $tileHeight * $row

            $outputFile = [System.IO.Path]::Combine($outputPath, "tile_${row}_${column}.png")

            $arguments = "-y -i "+"`""+$inputFile+"`""+ " -vf "+"crop=${tileWidth}:${tileHeight}:${x}:${y}"+ " `""+$outputFile+"`""

            $process = Start-Process -FilePath "ffmpeg" -ArgumentList $arguments -NoNewWindow -PassThru
            $processes += $process
        }
    }
    foreach ($process in $processes) {
        $process.WaitForExit()
    }
}



$Form = New-Object System.Windows.Forms.Form
$Form.Text = "FILM Toolset"

# Create ComboBox
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 40)
$comboBox.Items.AddRange(@('Sprite sheet', 'Two Images', 'Video', 'Folder of images'))
$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Form.Controls.Add($comboBox)

# Event handler for ComboBox value change
$comboBox_SelectedIndexChanged = {
    RemoveControlsFromForm $Form @(
        $pictureBox,
        $columnsLabel,
        $rowsLabel,
        $columnsNumericUpDown,
        $rowsNumericUpDown,
        $Interpolate,
        $applyButton,
        $InputOne,
        $InputTwo
    )
    switch ($comboBox.SelectedItem) {
        'Sprite sheet' {
            # Code for Sprite sheet
            $result = $openFileDialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $global:ImageFile = $openFileDialog.FileName
                $global:ImageObject = [System.Drawing.Image]::FromFile($global:ImageFile)
                $pictureBox.Image = $global:ImageObject
                AddControlsToForm $Form @(
                    $pictureBox,
                    $columnsLabel,
                    $rowsLabel,
                    $columnsNumericUpDown,
                    $rowsNumericUpDown,
                    $applyButton
                )
            }
            #$form.Controls.Add($applyButton)
            
        }
        'Two Images' {
            # Code for Two Images
            AddControlsToForm $Form @(
                    $Interpolate,
                    $InputOne,
                    $InputTwo
                )
        }
        'Video' {
            AddControlsToForm $Form @(
                    $Interpolate,
                    $InputOne
                )
        }
        'Folder of images' {
            AddControlsToForm $Form @(
                    $Interpolate,
                    $InputOne
                )
        }
    }
}

$comboBox.add_SelectedIndexChanged($comboBox_SelectedIndexChanged)



$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "Directory path:"
$Label1.AutoSize = $true
$Label1.Location = New-Object System.Drawing.Point(10, 20)
#$Form.Controls.Add($Label1)

$TextBox1 = New-Object System.Windows.Forms.TextBox
$TextBox1.Location = New-Object System.Drawing.Point(120, 18)
$TextBox1.Size = New-Object System.Drawing.Size(200, 20)
$TextBox1.Add_TextChanged({
    $DirectoryPath = $TextBox1.Text

    if (Test-Path $DirectoryPath -PathType Container) {
        $ImageExtensions = @("*.jpg", "*.jpeg", "*.png")
        $VideoExtensions = @("*.mp4", "*.avi", "*.mkv","*.gif")

        $ImageFiles = $ImageExtensions | ForEach-Object { Get-ChildItem -Path $DirectoryPath -Filter $_ -File -ErrorAction SilentlyContinue }
        $VideoFiles = $VideoExtensions | ForEach-Object { Get-ChildItem -Path $DirectoryPath -Filter $_ -File -ErrorAction SilentlyContinue }

        if ($ImageFiles -or $VideoFiles) {
            # Valid directory with image or video files
            $TextBox1.ForeColor = [System.Drawing.Color]::Black
        } else {
            # Valid directory but no image or video files
            $TextBox1.ForeColor = [System.Drawing.Color]::Orange
        }
    } else {
        # Invalid directory
        $TextBox1.ForeColor = [System.Drawing.Color]::Red
    }
})
#$Form.Controls.Add($TextBox1)


$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Text = "Browse"
$BrowseButton.Location = New-Object System.Drawing.Point(330, 17)
$BrowseButton.Size = New-Object System.Drawing.Size(80, 23)
$BrowseButton.Add_Click({
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($FolderBrowserDialog.ShowDialog() -eq "OK") {
        $TextBox1.Text = $FolderBrowserDialog.SelectedPath
    }
})
#$Form.Controls.Add($BrowseButton)

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Number of passes (1-6):"
$Label2.AutoSize = $true
$Label2.Location = New-Object System.Drawing.Point(10, 18)
$Form.Controls.Add($Label2)

$NumericUpDown1 = New-Object System.Windows.Forms.NumericUpDown
$NumericUpDown1.Location = New-Object System.Drawing.Point(140, 16)
$NumericUpDown1.Minimum = 1
$NumericUpDown1.Maximum = 6
$NumericUpDown1.Add_TextChanged({
    $CurrentValue = $NumericUpDown1.Value
    if ($CurrentValue -lt $NumericUpDown1.Minimum) {
        $NumericUpDown1.Value = $NumericUpDown1.Minimum
    } elseif ($CurrentValue -gt $NumericUpDown1.Maximum) {
        $NumericUpDown1.Value = $NumericUpDown1.Maximum
    }
})
$Form.Controls.Add($NumericUpDown1)


$Interpolate = New-Object System.Windows.Forms.Button
$Interpolate.Text = "Interpolate"
$Interpolate.Location = New-Object System.Drawing.Point(280, 20)
$Interpolate.Size = New-Object System.Drawing.Size(100, 50)
$Interpolate.Add_Click({
    $Interpolate.Enabled = $false
    $SetupEnv = "call environment.bat"
    if (IsAnaconda3Installed) {
        $SetupEnv = "call conda activate envr\frame_interpolation"
    }
    if ($InputOne.BackgroundImage -ne $null -and $InputTwo.BackgroundImage -ne $null) {
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "PNG Files|*.png"
        $SaveFileDialog.Title = "Save interpolated image"
        
        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $OutputPath = $SaveFileDialog.FileName

            $Script = "
            @echo off
            title FILM
            cd %~dp0
            "+$SetupEnv+"
            cd %~dp0frame-interpolation
            python -m eval.interpolator_test ^
   --frame1 `"$($InputOne.Tag)`" ^
   --frame2 `"$($InputTwo.Tag)`" ^
   --model_path pretrained_models/film_net/Style/saved_model ^
   --output_frame `"$OutputPath`"
"
        }
    } else {
        
        
        $global:DirOfFrames = $InputOne.Tag
        # Check if the file is a video file
        if ($InputOne.Tag -and ([System.IO.Path]::GetExtension($InputOne.Tag) -imatch "\.(mp4|avi|mkv|mov|gif)$")) {
            $videoFile = $InputOne.Tag
            $videoFolderPath = [System.IO.Path]::GetDirectoryName($videoFile)
            $videoFileName = [System.IO.Path]::GetFileNameWithoutExtension($videoFile)
            $outputFolderName = $videoFileName + "_frames"
            $outputFolderPath = Join-Path $videoFolderPath $outputFolderName
        
            # Extract frames using SplitVideoToFrames function
            SplitVideoToFrames $videoFile $outputFolderPath
            $global:DirOfFrames = $outputFolderPath
        }

        $Number = $NumericUpDown1.Value

        $Script = "
@echo off
title FILM
cd %~dp0
"+$SetupEnv+"
cd %~dp0frame-interpolation

python -m eval.interpolator_cli ^
   --pattern `"$($global:DirOfFrames)`" ^
   --model_path pretrained_models/film_net/Style/saved_model ^
   --times_to_interpolate $Number
"
    }

    if ($Script) {
        $Script | Out-File -FilePath "run2.bat" -Encoding ASCII
        $global:process = Start-Process "run2.bat" -PassThru
        $ProcessTracker.Start()
    } else {
        $Interpolate.Enabled = $true
    }
})


$ProcessTracker = New-Object System.Windows.Forms.Timer
$ProcessTracker.Interval = 1000
$ProcessTracker.Add_Tick({
    $isRunning = (IsProcessRunning -process $global:process)
    if (-Not $isRunning) {
       
        $ProcessTracker.Stop()
        $InputFile = $InputOne.Tag
            $InputFilePath = [System.IO.Path]::GetDirectoryName($InputFile)
            $InputFileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
            $outputFileName = $InputFileName + "_interp"+ [System.IO.Path]::GetExtension($InputFile)
            $outputFilePath = Join-Path $InputFilePath $outputFileName
        
        if ($InputOne.Tag -and ([System.IO.Path]::GetExtension($InputOne.Tag) -imatch "\.(mp4|avi|mkv|mov|gif)$")) {
            
            EncodeFramesToVideo ($global:DirOfFrames + "\interpolated_frames") (GetUniquePath $outputFilePath) (GetVideoFramerate $InputFile)
        }
        ProcessImageFolder ($global:DirOfFrames + "\interpolated_frames")
         #$TextBox1.Enabled = $true
        #$BrowseButton.Enabled = $true
        #$NumericUpDown1.Enabled = $true
        $Interpolate.Enabled = $true
    } else {
        #$TextBox1.Enabled = $false
        #$BrowseButton.Enabled = $false
        #$NumericUpDown1.Enabled = $false
        $Interpolate.Enabled = $false
    }
})
function GetUniquePath([string]$inputPath) {
    if (-not (Test-Path $inputPath)) {
        return $inputPath
    }

    $directory = [System.IO.Path]::GetDirectoryName($inputPath)
    $filenameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
    $extension = [System.IO.Path]::GetExtension($inputPath)

    $counter = 2
    $uniquePathFound = $false
    $newPath = $inputPath
    while (-not $uniquePathFound) {
        $newFilename = "$filenameWithoutExtension ($counter)$extension"
        $newPath = Join-Path $directory $newFilename

        if (-not (Test-Path $newPath)) {
            $uniquePathFound = $true
        } else {
            $counter++
        }
        
    }

    return $newPath
}


function GetVideoFramerate([string]$inputVideoPath) {
    $ffprobeExePath = "ffprobe" # Change this path if FFprobe is not in the system PATH

    # Execute FFprobe to get video stream information
    $videoInfo = (Start-Process -FilePath $ffprobeExePath -ArgumentList "-v", "error", "-select_streams", "v:0", "-show_entries", "stream=r_frame_rate", "-of", "default=noprint_wrappers=1:nokey=1", ("`""+($inputVideoPath)+"`"") -NoNewWindow -Wait -RedirectStandardOutput "FFprobeOutput.txt" -PassThru).ExitCode

    if ($videoInfo -eq 0) {
        $frameRate = Get-Content "FFprobeOutput.txt"
        Remove-Item "FFprobeOutput.txt"
        return (Invoke-Expression $frameRate)
    } else {
        Write-Host "Error getting video framerate"
        return $null
    }
}


function SplitVideoToFrames([string]$inputVideoPath, [string]$outputFolderPath) {
    if (-not (Test-Path $outputFolderPath)) {
        New-Item -ItemType Directory -Path $outputFolderPath | Out-Null
    }

    $ffmpegExePath = "ffmpeg" # Change this path if FFmpeg is not in the system PATH
    $outputPattern = Join-Path $outputFolderPath "frame%04d.png"

    Start-Process -FilePath $ffmpegExePath -ArgumentList "-i", ("`""+$inputVideoPath+"`""), ("`""+$outputPattern+"`"") -NoNewWindow -Wait
}


function EncodeFramesToVideo([string]$inputFramePattern, [string]$outputVideoPath, [int]$frameRate) {
    $ffmpegExePath = "ffmpeg" # Change this path if FFmpeg is not in the system PATH

    Start-Process -FilePath $ffmpegExePath -ArgumentList "-framerate", $frameRate, "-i", ("`""+($inputFramePattern + "\frame_%3d.png")+"`""), "-pix_fmt", "yuv420p", ("`""+$outputVideoPath+"`"") -NoNewWindow -Wait
}


function Button_DragEnter([object]$sender, [System.Windows.Forms.DragEventArgs]$e) {
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
}

function Button_DragDrop([object]$sender, [System.Windows.Forms.DragEventArgs]$e) {
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    $firstFile = $files[0]

    if (Test-Path $firstFile -PathType Container) {
        $sender.Text = Split-Path $firstFile -Leaf
        $sender.BackgroundImage = $null
        $sender.Tag = $firstFile
    } elseif (Test-Path $firstFile -PathType Leaf) {
        $extension = [System.IO.Path]::GetExtension($firstFile)

        if ($extension -imatch "\.(jpg|jpeg|png)$") {
            $sender.Text = ""
            $sender.BackgroundImageLayout = "Zoom"

            $image = [System.Drawing.Image]::FromFile($firstFile)
            $sender.BackgroundImage = $image
            $sender.Tag = $firstFile
        } elseif ($extension -imatch "\.(mp4|avi|mkv|mov|gif)$") {
            $sender.Text = Split-Path $firstFile -Leaf
            $sender.BackgroundImage = $null
            $sender.Tag = $firstFile
        } else {
            $sender.Text = "Invalid file format"
            $sender.BackgroundImage = $null
            $sender.Tag = $null
        }
    }
}



$InputOne = New-Object System.Windows.Forms.Button
$InputOne.Text = "Drop image or folder here"
$InputOne.Location = New-Object System.Drawing.Point(30, 80)
$InputOne.Size = New-Object System.Drawing.Size(200, 200)
$InputOne.AllowDrop = $true


$InputOne.Add_DragEnter({
    Button_DragEnter $InputOne $_
})
$InputOne.Add_DragDrop({
    Button_DragDrop $InputOne $_
})


$InputTwo = New-Object System.Windows.Forms.Button
$InputTwo.Text = "Drop image or folder here"
$InputTwo.Location = New-Object System.Drawing.Point(250, 80)
$InputTwo.Size = New-Object System.Drawing.Size(200, 200)
$InputTwo.AllowDrop = $true
$InputTwo.Add_DragEnter({
    Button_DragEnter $InputTwo $_
})
$InputTwo.Add_DragDrop({
    Button_DragDrop $InputTwo $_
})







$Form.StartPosition = "CenterScreen"
$Form.AutoSize = $true
$Form.AutoSizeMode = "GrowAndShrink"
$Form.ShowDialog()
$ProcessTracker.Stop()