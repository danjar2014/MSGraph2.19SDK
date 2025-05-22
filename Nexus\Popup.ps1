# Ensure types are only defined once
if (-not ([System.Management.Automation.PSTypeName]'KeyboardHook').Type) {
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class KeyboardHook {
            public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
            private static LowLevelKeyboardProc _proc = HookCallback;
            private static IntPtr _hookID = IntPtr.Zero;

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr GetModuleHandle(string lpModuleName);

            public const int WH_KEYBOARD_LL = 13;
            public const int WM_KEYDOWN = 0x0100;
            public const int WM_KEYUP = 0x0101;
            public const int VK_LWIN = 0x5B;
            public const int VK_RWIN = 0x5C;
            public const int VK_D = 0x44;
            public const int VK_TAB = 0x09;

            private static bool winPressed = false;

            public static void SetHook() {
                using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
                using (var curModule = curProcess.MainModule) {
                    _hookID = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }

            public static void Unhook() {
                UnhookWindowsHookEx(_hookID);
            }

            private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
                if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN) {
                    int vkCode = Marshal.ReadInt32(lParam);

                    // Track Win key state
                    if (vkCode == VK_LWIN || vkCode == VK_RWIN) {
                        winPressed = true;
                        return (IntPtr)1; // Block the Win key
                    }

                    // Block Ctrl + Win + D (new desktop) and Win + Tab (task view)
                    if ((winPressed && vkCode == VK_D) || (winPressed && vkCode == VK_TAB)) {
                        return (IntPtr)1; // Block the key combination
                    }
                } else if (wParam == (IntPtr)WM_KEYUP) {
                    int vkCode = Marshal.ReadInt32(lParam);
                    if (vkCode == VK_LWIN || vkCode == VK_RWIN) {
                        winPressed = false;
                    }
                }

                return CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }
"@
}

if (-not ([System.Management.Automation.PSTypeName]'Win32').Type) {
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            [DllImport("user32.dll")]
            public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

            [DllImport("user32.dll")]
            public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

            [DllImport("user32.dll")]
            public static extern bool IsWindowVisible(IntPtr hWnd);

            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            public const int SW_MINIMIZE = 6;
            public const int SW_RESTORE = 9;
        }

        public struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
"@
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Function to write to log file
function Write-Log {
    param (
        [string]$Message
    )

    $logFile = "C:\_Logfiles\SetPreBootPinBitlocker.log"
    $logDir = "C:\_Logfiles"

    try {
        # Create log directory if it doesn't exist
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }

        # Create log entry with timestamp
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] $Message"

        # Append to log file
        Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
    } catch {
        # If logging fails, write to console as fallback
        Write-Host "Failed to write to log file: $_"
        Write-Host $logEntry
    }
}

# Log the start of the script
Write-Log "Starting Popup.ps1"

# Set the keyboard hook to block Win key and new desktop creation
[KeyboardHook]::SetHook()

# Detect the system culture
$culture = (Get-Culture).Name
$isFrench = $culture -eq "fr-FR"

# Retrieve the complexity level (default: Low)
$ComplexityLevel = $env:PinComplexityLevel
if (-not $ComplexityLevel) {
    $ComplexityLevel = "Low"
}
if ($ComplexityLevel -notin @("Low", "Medium", "High")) {
    $ComplexityLevel = "Low"
}

# Global variable to allow closing the application
$script:allowClose = $false

# List to store minimized windows
$minimizedWindows = @()

# General configuration with explicit integer conversion
$screenWidth = [int][System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [int][System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
$margin = 50
$halfWidth = [int]($screenWidth / 2)

# Log screen dimensions
Write-Log "Primary screen dimensions: Width=$screenWidth, Height=$screenHeight"

# Get the bounds of the primary screen
$primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
$primaryBounds = $primaryScreen.Bounds

# Function to minimize windows on secondary screens
function Minimize-SecondaryScreenWindows {
    $callback = {
        param($hWnd, $lParam)

        # Check if the window is visible
        if (-not [Win32]::IsWindowVisible($hWnd)) {
            return $true
        }

        # Get the window's rectangle
        $rect = New-Object RECT
        if (-not [Win32]::GetWindowRect($hWnd, [ref]$rect)) {
            return $true
        }

        # Check if the window is on a secondary screen
        $windowCenterX = ($rect.Left + $rect.Right) / 2
        $windowCenterY = ($rect.Top + $rect.Bottom) / 2

        # If the window's center is outside the primary screen bounds, minimize it
        if ($windowCenterX -lt $primaryBounds.Left -or $windowCenterX -gt $primaryBounds.Right -or
            $windowCenterY -lt $primaryBounds.Top -or $windowCenterY -gt $primaryBounds.Bottom) {
            Write-Log "Minimizing window at X=$windowCenterX, Y=$windowCenterY (Handle: $hWnd)"
            [Win32]::ShowWindow($hWnd, [Win32]::SW_MINIMIZE)
            $script:minimizedWindows += $hWnd
        }

        return $true
    }

    # Convert the callback to a delegate
    $delegate = [Win32+EnumWindowsProc]$callback

    # Enumerate all windows
    [Win32]::EnumWindows($delegate, [IntPtr]::Zero)
}

# Function to restore minimized windows
function Restore-MinimizedWindows {
    foreach ($hWnd in $minimizedWindows) {
        Write-Log "Restoring window (Handle: $hWnd)"
        [Win32]::ShowWindow($hWnd, [Win32]::SW_RESTORE)
    }
}

# Minimize windows on secondary screens
Minimize-SecondaryScreenWindows

# Create the forms for the main interface
$formPage1 = New-Object System.Windows.Forms.Form
$formPage2 = New-Object System.Windows.Forms.Form

foreach ($form in @($formPage1, $formPage2)) {
    $form.FormBorderStyle = 'None'
    $form.WindowState = 'Maximized'
    $form.ControlBox = $false
    $form.TopMost = $true
}

# Handle the FormClosing event to block Alt+F4 unless $allowClose is true
foreach ($form in @($formPage1, $formPage2)) {
    $form.Add_FormClosing({
        if (-not $script:allowClose) {
            $_.Cancel = $true
        }
    })
}

# Load the background image for the main interface
$backgroundImagePath = "C:\Temp\BitLockerPinSetup\Wallpaper.png"
$backgroundImage = [System.Drawing.Image]::FromFile($backgroundImagePath)

foreach ($form in @($formPage1, $formPage2)) {
    $form.BackgroundImage = $backgroundImage
    $form.BackgroundImageLayout = 'Stretch'
}

# Paths to the images
$imagePath1 = "C:\Temp\BitLockerPinSetup\SetBitLockerPin.png"
$imagePath2 = "C:\Temp\BitLockerPinSetup\PIN-W11-BitLocker-0.png"
$bitmap1 = [System.Drawing.Image]::FromFile($imagePath1)
$bitmap2 = [System.Drawing.Image]::FromFile($imagePath2)

# Suspend rendering during control creation
$formPage1.SuspendLayout()
$formPage2.SuspendLayout()

# PAGE 1 - Explanation
# Semi-transparent panel for the left half (image)
$panelLeftPage1 = New-Object System.Windows.Forms.Panel
$panelLeftPage1.Location = New-Object System.Drawing.Point(0, 0)
$panelLeftPage1.Size = New-Object System.Drawing.Size($halfWidth, $screenHeight)
$panelLeftPage1.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)

# Semi-transparent panel for the right half (text and image)
$panelRightPage1 = New-Object System.Windows.Forms.Panel
$panelRightPage1.Location = New-Object System.Drawing.Point($halfWidth, 0)
$panelRightPage1.Size = New-Object System.Drawing.Size($halfWidth, $screenHeight)
$panelRightPage1.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)

# Main image (centered in the left half)
$pictureBox1 = New-Object System.Windows.Forms.PictureBox
$pictureBox1.Image = $bitmap1
$pictureBox1.SizeMode = 'StretchImage'
$pictureBox1.Size = New-Object System.Drawing.Size(260, 260)
$pic1X = [int](($halfWidth - 260) / 2)
$pic1Y = [int](($screenHeight - 260) / 2)
$pictureBox1.Location = New-Object System.Drawing.Point($pic1X, $pic1Y)

# Explanatory text (centered in the right half, increased size)
$labelExplanation1 = New-Object System.Windows.Forms.Label
if ($isFrench) {
    $labelExplanation1.Text = "Votre organisation requiert un code PIN BitLocker pour sécuriser votre disque.`nÀ chaque démarrage du PC, vous devrez entrer ce PIN pour accéder à votre système.`nCela permet d'empêcher tout accès non autorisé."
} else {
    $labelExplanation1.Text = "Your organization requires a BitLocker PIN to secure your drive.`nAt each PC startup, you will need to enter this PIN to access your system.`nThis helps prevent unauthorized access."
}
$labelExplanation1.Font = New-Object System.Drawing.Font("Calibri Light", 20)
$labelExplanation1.Size = New-Object System.Drawing.Size(($halfWidth - $margin * 2), 200)
$labelExpX = [int](($halfWidth - $labelExplanation1.Size.Width) / 2) - 50
$labelExpY = [int](($screenHeight - 400) / 2)
$labelExplanation1.Location = New-Object System.Drawing.Point($labelExpX, $labelExpY)

# Second image (below the text, centered in the right half)
$pictureBox2 = New-Object System.Windows.Forms.PictureBox
$pictureBox2.Image = $bitmap2
$pictureBox2.SizeMode = 'StretchImage'
$pictureBox2.Size = New-Object System.Drawing.Size(650, 260)
$pic2X = [int](($halfWidth - 500) / 2) - 200
$pic2Y = $labelExpY + 150
$pictureBox2.Location = New-Object System.Drawing.Point($pic2X, $pic2Y)

# "Next" button (same style as OK on Page 2, reduced size by 30%, adjusted position)
$buttonNext = New-Object System.Windows.Forms.Button
if ($isFrench) {
    $buttonNext.Text = "Suivant"
} else {
    $buttonNext.Text = "Next"
}
$buttonNext.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25)
$buttonNext.Size = New-Object System.Drawing.Size(140, 42)
$btnNextX = [int]($halfWidth - 140 - $margin - 80)
$btnNextY = [int]($screenHeight - 42 - $margin - 50)
$buttonNext.Location = New-Object System.Drawing.Point($btnNextX, $btnNextY)
$buttonNext.UseVisualStyleBackColor = $true

# Hidden clickable area to close the application (top-left corner, 1x1 pixel, using a Label)
$hiddenAreaPage1 = New-Object System.Windows.Forms.Label
$hiddenAreaPage1.Size = New-Object System.Drawing.Size(1, 1)
$hiddenAreaPage1.Location = New-Object System.Drawing.Point(5, 0)
$hiddenAreaPage1.BackColor = [System.Drawing.Color]::Transparent
$hiddenAreaPage1.Text = ""
$hiddenAreaPage1.Add_Click({
    Write-Log "Hidden area clicked on Page 1 - Closing application"
    $script:allowClose = $true
    $formPage1.Close()
    Restore-MinimizedWindows
    [KeyboardHook]::Unhook()
})

# Add elements to the panels
$panelLeftPage1.Controls.Add($pictureBox1)
$panelRightPage1.Controls.AddRange(@($labelExplanation1, $pictureBox2, $buttonNext))
$formPage1.Controls.Add($hiddenAreaPage1)
$hiddenAreaPage1.BringToFront()

# Add panels to the form
$formPage1.Controls.AddRange(@($panelLeftPage1, $panelRightPage1))

# PAGE 2 - PIN Entry (enlarged area by 30%, including text)
# Semi-transparent panel for the left half (image)
$panelLeftPage2 = New-Object System.Windows.Forms.Panel
$panelLeftPage2.Location = New-Object System.Drawing.Point(0, 0)
$panelLeftPage2.Size = New-Object System.Drawing.Size($halfWidth, $screenHeight)
$panelLeftPage2.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)

# Semi-transparent panel for the right half (form)
$panelRightPage2 = New-Object System.Windows.Forms.Panel
$panelRightPage2.Location = New-Object System.Drawing.Point($halfWidth, 0)
$panelRightPage2.Size = New-Object System.Drawing.Size($halfWidth, $screenHeight)
$panelRightPage2.BackColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)

# Main image (centered in the left half, same as Page 1)
$pictureBox1Clone = New-Object System.Windows.Forms.PictureBox
$pictureBox1Clone.Image = $bitmap1
$pictureBox1Clone.SizeMode = 'StretchImage'
$pictureBox1Clone.Size = New-Object System.Drawing.Size(260, 260)
$pictureBox1Clone.Location = New-Object System.Drawing.Point($pic1X, $pic1Y)

# PIN entry form
$labelPINIsNotEqual = New-Object System.Windows.Forms.Label
$labelRetypePIN = New-Object System.Windows.Forms.Label
$labelNewPIN = New-Object System.Windows.Forms.Label
$labelChoosePin = New-Object System.Windows.Forms.Label
$buttonSetPIN = New-Object System.Windows.Forms.Button
$labelSetBLtartupPin = New-Object System.Windows.Forms.Label
$textboxRetypedPin = New-Object System.Windows.Forms.TextBox
$textboxNewPin = New-Object System.Windows.Forms.TextBox

# Positioning in the right half (forced to integer, adjusted for the new size)
$inputStartX = [int](($halfWidth - 579) / 2)
$inputStartY = [int](($screenHeight - 352) / 2)

# Hidden clickable area to close the application (top-left corner, 1x1 pixel, using a Label)
$hiddenAreaPage2 = New-Object System.Windows.Forms.Label
$hiddenAreaPage2.Size = New-Object System.Drawing.Size(1, 1)
$hiddenAreaPage2.Location = New-Object System.Drawing.Point(5, 0)
$hiddenAreaPage2.BackColor = [System.Drawing.Color]::Transparent
$hiddenAreaPage2.Text = ""
$hiddenAreaPage2.Add_Click({
    Write-Log "Hidden area clicked on Page 2 - Closing application"
    $script:allowClose = $true
    $formPage2.Close()
    $formPage1.Close()
    Restore-MinimizedWindows
    [KeyboardHook]::Unhook()
})

# Function to test PIN security with complexity levels
function Test-PinSecurity {
    param (
        [String]$PinString,
        [String]$ComplexityLevel = "High"
    )

    $array = $PinString.ToCharArray()

    # Low level: Minimal checks
    # 1. Check for digit repetition (e.g., 1111, 2222)
    $isRepetitive = $true
    $firstChar = $array[0]
    foreach ($char in $array) {
        if ($char -ne $firstChar) {
            $isRepetitive = $false
            break
        }
    }
    if ($isRepetitive) { return $true }

    # 2. Check for incremental sequences (e.g., 1234)
    $isIncremental = $true
    for ($i = 0; $i -lt $array.Count - 1; $i++) {
        if ([int]$array[$i] + 1 -eq [int]$array[$i + 1]) { continue } else { $isIncremental = $false; break }
    }
    if ($isIncremental) { return $true }

    # 3. Check for decremental sequences (e.g., 4321)
    $isDecremental = $true
    for ($i = 0; $i -lt $array.Count - 1; $i++) {
        if ([int]$array[$i] - 1 -eq [int]$array[$i + 1]) { continue } else { $isDecremental = $false; break }
    }
    if ($isDecremental) { return $true }

    if ($ComplexityLevel -eq "Low") { return $false }

    # Medium level: Additional checks
    # 4. Check for alternating sequences (e.g., 1212, 3434)
    $isAlternating = $true
    for ($i = 0; $i -lt $array.Count; $i++) {
        if ($i % 2 -eq 0 -and $array[$i] -ne $array[0]) { $isAlternating = $false; break }
        if ($i % 2 -eq 1 -and $array[$i] -ne $array[1]) { $isAlternating = $false; break }
    }
    if ($isAlternating -and $array[0] -ne $array[1]) { return $true }

    # 5. Check for identical digit blocks (e.g., 1122, 3344)
    $isBlockRepetitive = $true
    $blockSize = $array.Count / 2
    if ($array.Count % 2 -eq 0) {
        for ($i = 0; $i -lt $blockSize; $i++) {
            if ($array[$i] -ne $array[0]) { $isBlockRepetitive = $false; break }
        }
        for ($i = $blockSize; $i -lt $array.Count; $i++) {
            if ($array[$i] -ne $array[$blockSize]) { $isBlockRepetitive = $false; break }
        }
        if ($isBlockRepetitive -and $array[0] -ne $array[$blockSize]) { return $true }
    }

    # 6. Check for character diversity (at least 3 different digits)
    $distinctDigits = ($array | Sort-Object | Get-Unique).Count
    if ($distinctDigits -lt 3) { return $true }

    if ($ComplexityLevel -eq "Medium") { return $false }

    # High level: Maximum checks
    # 7. Check for mathematical sequences (e.g., 1357, 2468)
    $isOddSequence = $true
    $isEvenSequence = $true
    for ($i = 0; $i -lt $array.Count; $i++) {
        $digit = [int]$array[$i]
        if ($i % 2 -eq 0) {
            if ($digit -ne [int]$array[0] + ($i * 2)) { $isOddSequence = $false }
            if ($digit -ne [int]$array[0] + ($i * 2)) { $isEvenSequence = $false }
        } else {
            if ($digit -ne [int]$array[1] + (($i - 1) * 2)) { $isOddSequence = $false }
            if ($digit -ne [int]$array[1] + (($i - 1) * 2)) { $isEvenSequence = $false }
        }
    }
    if ($isOddSequence -or $isEvenSequence) { return $true }

    # 8. Check for partial consecutive sequences (at least 3 consecutive digits)
    for ($i = 0; $i -lt $array.Count - 2; $i++) {
        $isPartialIncremental = $true
        $isPartialDecremental = $true
        for ($j = $i; $j -lt $i + 2 -and $j -lt $array.Count - 1; $j++) {
            if ([int]$array[$j] + 1 -ne [int]$array[$j + 1]) { $isPartialIncremental = $false }
            if ([int]$array[$j] - 1 -ne [int]$array[$j + 1]) { $isPartialDecremental = $false }
        }
        if ($isPartialIncremental -or $isPartialDecremental) { return $true }
    }

    return $false
}

# Event handlers
$formBitLockerStartupPIN_Load = {
    Write-Log "Form Page 2 loaded"
    $formPage2.Activate()
    $textboxNewPin.Focus()
    try { $global:MinimumPIN = Get-ItemPropertyValue HKLM:\SOFTWARE\Policies\Microsoft\FVE -Name MinimumPIN -ErrorAction SilentlyContinue } catch { }
    try { $global:EnhancedPIN = Get-ItemPropertyValue HKLM:\SOFTWARE\Policies\Microsoft\FVE -Name UseEnhancedPin -ErrorAction SilentlyContinue } catch { }
    if ($isFrench) {
        $characters = "chiffres"
        if ($global:EnhancedPIN -eq 1) { $characters = "caractères" }
    } else {
        $characters = "digits"
        if ($global:EnhancedPIN -eq 1) { $characters = "characters" }
    }
    if ($global:MinimumPIN -isnot [int] -or $global:MinimumPIN -lt 4) { $global:MinimumPIN = 6 }
    if ($isFrench) {
        $labelChoosePin.Text = "Choisissez un PIN de $global:MinimumPIN à 20 $characters."
    } else {
        $labelChoosePin.Text = "Choose a PIN from $global:MinimumPIN to 20 $characters."
    }
    $labelChoosePin.Font = New-Object System.Drawing.Font("Calibri Light", 17.0)
}

$buttonSetPIN_Click = {
    Write-Log "OK button clicked"
    # Check PIN length (between $global:MinimumPIN and 20 digits)
    if ($textboxNewPin.Text.Length -lt $global:MinimumPIN) {
        $labelPINIsNotEqual.ForeColor = 'Red'
        if ($isFrench) {
            $labelPINIsNotEqual.Text = "Le PIN n'est pas assez long"
            Write-Log "PIN validation failed: Le PIN n'est pas assez long"
        } else {
            $labelPINIsNotEqual.Text = "The PIN is not long enough"
            Write-Log "PIN validation failed: The PIN is not long enough"
        }
        $labelPINIsNotEqual.Visible = $true
        return
    }
    if ($textboxNewPin.Text.Length -gt 20) {
        $labelPINIsNotEqual.ForeColor = 'Red'
        if ($isFrench) {
            $labelPINIsNotEqual.Text = "Le PIN est trop long (maximum 20 chiffres)"
            Write-Log "PIN validation failed: Le PIN est trop long (maximum 20 chiffres)"
        } else {
            $labelPINIsNotEqual.Text = "The PIN is too long (maximum 20 digits)"
            Write-Log "PIN validation failed: The PIN is too long (maximum 20 digits)"
        }
        $labelPINIsNotEqual.Visible = $true
        return
    }
    if ($textboxNewPin.Text.Length -eq 0) {
        $labelPINIsNotEqual.ForeColor = 'Red'
        if ($isFrench) {
            $labelPINIsNotEqual.Text = "Le PIN ne peut pas être vide"
            Write-Log "PIN validation failed: Le PIN ne peut pas être vide"
        } else {
            $labelPINIsNotEqual.Text = "The PIN cannot be empty"
            Write-Log "PIN validation failed: The PIN cannot be empty"
        }
        $labelPINIsNotEqual.Visible = $true
        return
    }
    if ($global:EnhancedPIN -eq "" -or $global:EnhancedPIN -eq $null -or $global:EnhancedPIN -eq 0) {
        if ($textboxNewPin.Text -NotMatch "^[\d\.]+$") {
            $labelPINIsNotEqual.ForeColor = 'Red'
            if ($isFrench) {
                $labelPINIsNotEqual.Text = "Seuls les chiffres sont autorisés"
                Write-Log "PIN validation failed: Seuls les chiffres sont autorisés"
            } else {
                $labelPINIsNotEqual.Text = "Only digits are allowed"
                Write-Log "PIN validation failed: Only digits are allowed"
            }
            $labelPINIsNotEqual.Visible = $true
            return
        }
    }
    if ($textboxNewPin.Text -eq $textboxRetypedPin.Text) {
        if (Test-PinSecurity -PinString $textboxNewPin.Text -ComplexityLevel $ComplexityLevel) {
            $labelPINIsNotEqual.ForeColor = 'Red'
            if ($isFrench) {
                $labelPINIsNotEqual.Text = "Le PIN est trop simple"
                Write-Log "PIN validation failed: Le PIN est trop simple"
            } else {
                $labelPINIsNotEqual.Text = "The PIN is too simple"
                Write-Log "PIN validation failed: The PIN is too simple"
            }
            $labelPINIsNotEqual.Visible = $true
            return
        }
        $labelPINIsNotEqual.Visible = $false
        $key = (43,155,164,59,21,127,28,43,81,18,198,145,127,51,72,55,39,23,228,166,146,237,41,131,176,14,4,67,230,81,212,214)
        $secure = ConvertTo-SecureString $textboxNewPin.Text -AsPlainText -Force
        $encodedText = ConvertFrom-SecureString -SecureString $secure -Key $key
        $pathPINFile = Join-Path -Path "$env:SystemRoot\tracing" -ChildPath "168ba6df825678e4da1a.tmp"
        Out-File -FilePath $pathPINFile -InputObject $encodedText -Force
        if ($isFrench) {
            [System.Windows.Forms.MessageBox]::Show("PIN BitLocker défini avec succès !", "Confirmation", "OK", "Information")
            Write-Log "PIN BitLocker défini avec succès !"
        } else {
            [System.Windows.Forms.MessageBox]::Show("BitLocker PIN set successfully!", "Confirmation", "OK", "Information")
            Write-Log "BitLocker PIN set successfully!"
        }
        $script:allowClose = $true
        $formPage2.Close()
        $formPage1.Close()
        Restore-MinimizedWindows
        [KeyboardHook]::Unhook()
    }
    else {
        $labelPINIsNotEqual.ForeColor = 'Red'
        if ($isFrench) {
            $labelPINIsNotEqual.Text = "Les PIN ne correspondent pas"
            Write-Log "PIN validation failed: Les PIN ne correspondent pas"
        } else {
            $labelPINIsNotEqual.Text = "The PINs do not match"
            Write-Log "PIN validation failed: The PINs do not match"
        }
        $labelPINIsNotEqual.Visible = $true
    }
}

$textboxRetypedPin_KeyUp = [System.Windows.Forms.KeyEventHandler]{ if ($_.KeyCode -eq 'Enter') { $buttonSetPIN_Click.Invoke() } }
$textboxNewPin_KeyUp = [System.Windows.Forms.KeyEventHandler]{ if ($_.KeyCode -eq 'Enter') { $buttonSetPIN_Click.Invoke() } }

$Form_Cleanup_FormClosed = {
    Write-Log "Form Page 2 closed"
    try {
        $buttonSetPIN.remove_Click($buttonSetPIN_Click)
        $textboxRetypedPin.remove_KeyUp($textboxRetypedPin_KeyUp)
        $textboxNewPin.remove_KeyUp($textboxNewPin_KeyUp)
        $formPage2.remove_Load($formBitLockerStartupPIN_Load)
        $formPage2.remove_FormClosed($Form_Cleanup_FormClosed)
    }
    catch { Out-Null }
}

# Configure controls (sizes and positions adjusted for 30% increase, including text)
if ($isFrench) {
    $labelSetBLtartupPin.Text = "Configurer un PIN BitLocker"
} else {
    $labelSetBLtartupPin.Text = "Set a BitLocker PIN"
}
$labelSetBLtartupPin.Font = New-Object System.Drawing.Font("Calibri Light", 25.0)
$labelSetBLtartupPin.ForeColor = 'MediumBlue'
$labelSetBLtartupPin.Location = New-Object System.Drawing.Point([int]($inputStartX + 25 * 1.3), [int]($inputStartY + 17 * 1.3))
$labelSetBLtartupPin.AutoSize = $true

$labelChoosePin.Location = New-Object System.Drawing.Point([int]($inputStartX + 26 * 1.3), [int]($inputStartY + 60 * 1.3))
$labelChoosePin.AutoSize = $true

if ($isFrench) {
    $labelNewPIN.Text = "Nouveau PIN :"
} else {
    $labelNewPIN.Text = "New PIN:"
}
$labelNewPIN.Font = New-Object System.Drawing.Font("Calibri Light", 17.0)
$labelNewPIN.Location = New-Object System.Drawing.Point([int]($inputStartX + 26 * 1.3), [int]($inputStartY + 105 * 1.3))
$labelNewPIN.AutoSize = $true

$textboxNewPin.Location = New-Object System.Drawing.Point([int]($inputStartX + 165 * 1.3), [int]($inputStartY + 102 * 1.3))
$textboxNewPin.Size = New-Object System.Drawing.Size(278, 30)
$textboxNewPin.UseSystemPasswordChar = $true
$textboxNewPin.Font = New-Object System.Drawing.Font("Calibri Light", 17.0)

if ($isFrench) {
    $labelRetypePIN.Text = "Confirmer le PIN :"
} else {
    $labelRetypePIN.Text = "Confirm PIN:"
}
$labelRetypePIN.Location = New-Object System.Drawing.Point([int]($inputStartX + 26 * 1.3), [int]($inputStartY + 146 * 1.3))
$labelRetypePIN.AutoSize = $true
$labelRetypePIN.Font = New-Object System.Drawing.Font("Calibri Light", 17.0)

$textboxRetypedPin.Location = New-Object System.Drawing.Point([int]($inputStartX + 165 * 1.3), [int]($inputStartY + 143 * 1.3))
$textboxRetypedPin.Size = New-Object System.Drawing.Size(278, 30)
$textboxRetypedPin.UseSystemPasswordChar = $true
$textboxRetypedPin.Font = New-Object System.Drawing.Font("Calibri Light", 17.0)

$labelPINIsNotEqual.Location = New-Object System.Drawing.Point([int]($inputStartX + 300 * 1.3), [int]($inputStartY + 166 * 1.3))
$labelPINIsNotEqual.AutoSize = $true
$labelPINIsNotEqual.ForeColor = 'Red'
$labelPINIsNotEqual.Visible = $false
$labelPINIsNotEqual.Font = New-Object System.Drawing.Font("Calibri Light", 17.0)

$buttonSetPIN.Text = "OK"
$buttonSetPIN.Location = New-Object System.Drawing.Point([int]($inputStartX + 165 * 1.3 + (278 - 130) / 2), [int]($inputStartY + 180 * 1.3))
$buttonSetPIN.Size = New-Object System.Drawing.Size(130, 39)
$buttonSetPIN.Font = New-Object System.Drawing.Font("Calibri Light", 17.0)
$buttonSetPIN.add_Click($buttonSetPIN_Click)

# Add elements to the panels
$panelLeftPage2.Controls.Add($pictureBox1Clone)
$panelRightPage2.Controls.AddRange(@($labelPINIsNotEqual, $labelRetypePIN, $labelNewPIN, $labelChoosePin, $labelSetBLtartupPin, $textboxRetypedPin, $textboxNewPin, $buttonSetPIN))
$formPage2.Controls.Add($hiddenAreaPage2)
$hiddenAreaPage2.BringToFront()

# Add panels to the form
$formPage2.Controls.AddRange(@($panelLeftPage2, $panelRightPage2))

# Resume rendering after control creation
$formPage1.ResumeLayout($false)
$formPage2.ResumeLayout($false)

# Add event handlers
$formPage2.add_Load($formBitLockerStartupPIN_Load)
$formPage2.add_FormClosed($Form_Cleanup_FormClosed)
$textboxNewPin.add_KeyUp($textboxNewPin_KeyUp)
$textboxRetypedPin.add_KeyUp($textboxRetypedPin_KeyUp)

# Handle Page 1 events
$buttonNext.Add_Click({
    Write-Log "Next button clicked - Showing Page 2"
    $formPage1.Close()
    $formPage2.ShowDialog()
})

# Initial display with focus
Write-Log "Showing main form (Page 1)"
$formPage1.ShowDialog()
