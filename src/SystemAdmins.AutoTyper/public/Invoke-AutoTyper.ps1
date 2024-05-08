#Requires -version 5.1;

<#
.SYNOPSIS
  A small WinForm program made in PowerShell to send key strokes to the active window.

.DESCRIPTION
  This program uses runspaces and WinForm to automatically send keystrokes to the active window. The program can be used to send keystrokes to a game, a chat window, or any other application that accepts keyboard input.
  Usually it is used to automate repetitive tasks, such as typing configuration, passwords and other information into interfaces such as iLO, RDP etc.

.Example
   .\Invoke-AutoTyper.ps1;

.NOTES
  Version:        1.0
  Author:         Alex Hansen (ath@systemadmins.com)
  Creation Date:  07-05-2024
  Purpose/Change: Initial script development.
#>

#region begin boostrap
############### Parameters - Start ###############

[cmdletbinding()]
param
(
)

############### Parameters - End ###############
#endregion

#region begin bootstrap
############### Bootstrap - Start ###############

# Add required assemblies.
$null = [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms');
$null = [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing');
$null = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic');

# Add required namespaces for hiding PowerShell window when launching.
Add-Type -Name Window -Namespace Console -MemberDefinition @'
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'@;

# Set the AppID for a window in PowerShel   l (seperate from the main script). Set icon in taskbar.
$definitionSetAppIdForWindow = @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class PSAppID
{
    // https://emoacht.wordpress.com/2012/11/14/csharp-appusermodelid/
    // IPropertyStore Interface
    [ComImport,
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    private interface IPropertyStore
    {
        uint GetCount([Out] out uint cProps);
        uint GetAt([In] uint iProp, out PropertyKey pkey);
        uint GetValue([In] ref PropertyKey key, [Out] PropVariant pv);
        uint SetValue([In] ref PropertyKey key, [In] PropVariant pv);
        uint Commit();
    }


    // PropertyKey Structure
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PropertyKey
    {
        private Guid formatId;    // Unique GUID for property
        private Int32 propertyId; // Property identifier (PID)

        public Guid FormatId
        {
            get
            {
                return formatId;
            }
        }

        public Int32 PropertyId
        {
            get
            {
                return propertyId;
            }
        }

        public PropertyKey(Guid formatId, Int32 propertyId)
        {
            this.formatId = formatId;
            this.propertyId = propertyId;
        }

        public PropertyKey(string formatId, Int32 propertyId)
        {
            this.formatId = new Guid(formatId);
            this.propertyId = propertyId;
        }

    }


    // PropVariant Class (only for string value)
    [StructLayout(LayoutKind.Explicit)]
    public class PropVariant : IDisposable
    {
        [FieldOffset(0)]
        ushort valueType;     // Value type

        // [FieldOffset(2)]
        // ushort wReserved1; // Reserved field
        // [FieldOffset(4)]
        // ushort wReserved2; // Reserved field
        // [FieldOffset(6)]
        // ushort wReserved3; // Reserved field

        [FieldOffset(8)]
        IntPtr ptr;           // Value


        // Value type (System.Runtime.InteropServices.VarEnum)
        public VarEnum VarType
        {
            get { return (VarEnum)valueType; }
            set { valueType = (ushort)value; }
        }

        public bool IsNullOrEmpty
        {
            get
            {
                return (valueType == (ushort)VarEnum.VT_EMPTY ||
                        valueType == (ushort)VarEnum.VT_NULL);
            }
        }

        // Value (only for string value)
        public string Value
        {
            get
            {
                return Marshal.PtrToStringUni(ptr);
            }
        }


        public PropVariant()
        { }

        public PropVariant(string value)
        {
            if (value == null)
                throw new ArgumentException("Failed to set value.");

            valueType = (ushort)VarEnum.VT_LPWSTR;
            ptr = Marshal.StringToCoTaskMemUni(value);
        }

        ~PropVariant()
        {
            Dispose();
        }

        public void Dispose()
        {
            PropVariantClear(this);
            GC.SuppressFinalize(this);
        }

    }

    [DllImport("Ole32.dll", PreserveSig = false)]
    private extern static void PropVariantClear([In, Out] PropVariant pvar);


    [DllImport("shell32.dll")]
    private static extern int SHGetPropertyStoreForWindow(
        IntPtr hwnd,
        ref Guid iid /*IID_IPropertyStore*/,
        [Out(), MarshalAs(UnmanagedType.Interface)] out IPropertyStore propertyStore);

    public static void SetAppIdForWindow(int handle, string AppId)
    {
        Guid iid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");
        IPropertyStore prop;
        int result1 = SHGetPropertyStoreForWindow((IntPtr)handle, ref iid, out prop);

        // Name = System.AppUserModel.ID
        // ShellPKey = PKEY_AppUserModel_ID
        // FormatID = 9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3
        // PropID = 5
        // Type = String (VT_LPWSTR)
        PropertyKey AppUserModelIDKey = new PropertyKey("{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 5);

        PropVariant pv = new PropVariant(AppId);

        uint result2 = prop.SetValue(ref AppUserModelIDKey, pv);

        Marshal.ReleaseComObject(prop);
    }
}
'@;

# Try to add the type definition.
try
{
    # Add the type definition.
    Add-Type -TypeDefinition $definitionSetAppIdForWindow -ErrorAction SilentlyContinue;
}
# Something went wrong importing the type.
catch
{
    # Write to log.
    Write-Verbose ('Failed to import type definition. Error: {0}' -f $_.Exception.Message);
}


############### Bootstrap - End ###############
#endregion

#region begin variables
############### Variables - Start ###############

# Object array to save the input.
$Script:saved = New-Object System.Collections.ArrayList;

# Object array to store runspaces.
$Script:runspaces = New-Object System.Collections.ArrayList;

############### Variables - End ###############
#endregion

#region begin functions
############### Functions - Start ###############

# Hide the console window.
function Hide-Console
{
    # Get the console window.
    $consolePtr = [Console.Window]::GetConsoleWindow();

    # Hide the console window.
    [void][Console.Window]::ShowWindow($consolePtr, 0);
}

# Function to send key strokes.
function Send-KeyboardInput
{
    [CmdletBinding()]
    [OutputType([void])]
    param
    (
        # The input string to send.
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InputString,

        # The delay between each key press.
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int]$DelayBeforeTypingInSeconds = 5,

        # The delay between each key press.
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int]$DelayInMiliseconds = 200
    )

    # Start sleep.
    Start-Sleep -Seconds $DelayBeforeTypingInSeconds;

    # Convert the input string to a char array.
    [char[]]$charArray = $InputString.ToCharArray();

    # Foreach character in the array, send the key.
    foreach ($char in $charArray)
    {
        # Delay between each key press.
        Start-Sleep -Milliseconds $DelayInMiliseconds;

        # Send the key.
        [System.Windows.Forms.SendKeys]::SendWait($char);
    }
}

# Function to create a new runspace.
function Get-PowerShellRunspace
{
    [OutputType([System.Management.Automation.PowerShell])]
    param
    (
    )

    # Create a new session state.
    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault();

    # Get all functions.
    $functions = Get-ChildItem -Path 'Function:\' -Force;

    # Foreach function.
    foreach ($function in $functions)
    {
        # If the function have a ":" inside (such as drives).
        if ($function.Name -like '*:*')
        {
            # Continue.
            continue;
        }

        # If Source is not null.
        if ($null -ne $function.ScriptBlock.Source)
        {
            # Continue.
            continue;
        }

        # Try to add the function to the session state.
        try
        {
            # Add function to the session state.
            $sessionState.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry $function.Name, $function.Definition));
        }
        catch
        {
            # Write error.
            Write-Error -Message $_.Exception.Message;
        }
    }

    # Create a new runspace pool.
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 5, $sessionState, $Host);

    # Open the runspace pool.
    $runspacePool.Open();

    # Create a new runspace.
    $runspace = [powershell]::Create();

    # Assign the runspace pool to the runspace.
    $runspace.RunspacePool = $runspacePool;

    # Return the runspace.
    return $runspace;
}

# Function to present GUI.
function Show-Form
{
    [CmdletBinding()]
    [OutputType([void])]
    param
    (
    )

    #region begin WinFormObjects
    # Create WinForm objects.
    $form = New-Object Windows.Forms.Form;
    $notifyIcon = New-Object Windows.Forms.NotifyIcon;
    $statusStrip = New-Object Windows.Forms.StatusStrip;
    $statusStripLabel = New-Object Windows.Forms.ToolStripStatusLabel;
    $tabControl = New-Object Windows.Forms.TabControl;
    $tabPageInsert = New-Object Windows.Forms.TabPage;
    $tabPageSaved = New-Object Windows.Forms.TabPage;
    $tabPageSettings = New-Object Windows.Forms.TabPage;
    $buttonTypeSaveToFile = New-Object Windows.Forms.Button;
    $buttonTypeClear = New-Object Windows.Forms.Button;
    $buttonTypeSave = New-Object Windows.Forms.Button;
    $buttonTypeSend = New-Object Windows.Forms.Button;
    $buttonTypeCancel = New-Object Windows.Forms.Button;
    $buttonSavedExport = New-Object Windows.Forms.Button;
    $buttonSavedImport = New-Object Windows.Forms.Button;
    $buttonSavedCopyToClipboard = New-Object Windows.Forms.Button;
    $richTextBoxInsert = New-Object Windows.Forms.RichTextBox;
    $richTextBoxSaved = New-Object Windows.Forms.RichTextBox;
    $textBoxDelayKey = New-Object Windows.Forms.TextBox;
    $textBoxDelayWait = New-Object Windows.Forms.TextBox;
    $labelDelayKey = New-Object Windows.Forms.Label;
    $labelDelayWait = New-Object Windows.Forms.Label;
    $splitContainerSaved = New-Object Windows.Forms.SplitContainer;
    $listBoxSaved = New-Object Windows.Forms.ListBox;
    $groupBoxSettingsDelay = New-Object Windows.Forms.GroupBox;
    $contextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip
    #endregion

    #region begin Icon
    # Convert Base64 string to bitmap to use as an icon.
    $iconBase64 = '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';
    $iconBytes = [Convert]::FromBase64String($iconBase64);
    $iconStream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length);
    $iconBitmap = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($iconStream).GetHIcon()));
    #endregion

    #region begin General
    # Set form properties.
    $form.Text = 'AutoTyper';
    $form.Size = New-Object Drawing.Size @(1162, 770);
    $form.StartPosition = 'CenterScreen';
    $form.FormBorderStyle = 'FixedSingle';
    $form.MaximizeBox = $false;
    $form.MinimizeBox = $true;
    $form.Name = 'formMain';
    $form.Icon = $iconBitmap;

    # Set notify icon properties.
    $notifyIcon.Icon = $iconBitmap;
    $notifyIcon.Visible = $true;
    $notifyIcon.Text = 'AutoTyper';

    # Set tab control properties.
    $tabControl.Dock = 'Fill';
    $tabControl.Name = 'tabControl';
    $tabControl.Size = New-Object Drawing.Size @(1162, 712);
    $tabControl.Location = New-Object Drawing.Point @(0, 0);

    # Set status strip properties.
    $statusStrip.Dock = 'Bottom';
    $statusStrip.Text = 'Status';
    $statusStrip.Name = 'statusStrip';
    $statusStrip.Size = New-Object Drawing.Size @(1148, 22);
    $statusStrip.Location = New-Object Drawing.Point @(0, 712);
    $null = $statusStrip.Items.Add($statusStripLabel);

    # Set status strip label properties.
    $statusStripLabel.Text = 'Ready';
    $statusStripLabel.Name = 'statusStripLabel';

    #endregion

    #region begin trayArea
    $contextMenuItemOpen = $contextMenuStrip.Items.Add('Open');
    $contextMenuItemExit = $contextMenuStrip.Items.Add('Exit');
    #endregion

    #region begin tabPageInsert
    # Set tab page for "tabPageInsert" properties.
    $tabPageInsert.Text = 'Typer';
    $tabPageInsert.Name = 'tabPageInsert';
    $tabPageInsert.Location = New-Object Drawing.Point @(4, 24);
    $tabPageInsert.Size = New-Object Drawing.Size @(1140, 684);
    $tabPageInsert.Padding = New-Object System.Windows.Forms.Padding(3);
    $tabPageInsert.UseVisualStyleBackColor = $true;

    # Set richTextBoxInsert properties.
    $richTextBoxInsert.Location = New-Object Drawing.Point @(6, 6);
    $richTextBoxInsert.Size = New-Object Drawing.Size @(1035, 675);
    $richTextBoxInsert.Name = 'richTextBoxInsert';
    $richTextBoxInsert.Text = '<insert text that should be typed here>';
    $richTextBoxInsert.ScrollBars = 'Vertical';
    $richTextBoxInsert.TabIndex = 0;

    # Set buttonSaveToFile properties.
    $buttonTypeSaveToFile.Text = 'Save to File';
    $buttonTypeSaveToFile.Size = New-Object Drawing.Size @(87, 32);
    $buttonTypeSaveToFile.Location = New-Object Drawing.Point @(1047, 120);
    $buttonTypeSaveToFile.Name = 'buttonTypeSaveToFile';
    $buttonTypeSaveToFile.UseVisualStyleBackColor = $true;
    $buttonTypeSaveToFile.TabIndex = 4;

    # Set buttonClear properties.
    $buttonTypeClear.Text = 'Clear';
    $buttonTypeClear.Size = New-Object Drawing.Size @(87, 32);
    $buttonTypeClear.Location = New-Object Drawing.Point @(1047, 158);
    $buttonTypeClear.Name = 'buttonTypeClear';
    $buttonTypeClear.UseVisualStyleBackColor = $true;
    $buttonTypeClear.TabIndex = 2;

    # Set buttonSave properties.
    $buttonTypeSave.Text = 'Save';
    $buttonTypeSave.Size = New-Object Drawing.Size @(87, 32);
    $buttonTypeSave.Location = New-Object Drawing.Point @(1047, 82);
    $buttonTypeSave.Name = 'buttonTypeSave';
    $buttonTypeSave.UseVisualStyleBackColor = $true;
    $buttonTypeSave.TabIndex = 3;

    # Set buttonSend properties.
    $buttonTypeSend.Text = 'Send';
    $buttonTypeSend.Size = New-Object Drawing.Size @(87, 32);
    $buttonTypeSend.Location = New-Object Drawing.Point @(1047, 6);
    $buttonTypeSend.Name = 'buttonTypeSend';
    $buttonTypeSend.UseVisualStyleBackColor = $true;
    $buttonTypeSend.TabIndex = 1;

    # Set buttonTypeCancel properties.
    $buttonTypeCancel.Text = 'Cancel';
    $buttonTypeCancel.Size = New-Object Drawing.Size @(87, 32);
    $buttonTypeCancel.Location = New-Object Drawing.Point @(1047, 44);
    $buttonTypeCancel.Name = 'buttonTypeCancel';
    $buttonTypeCancel.UseVisualStyleBackColor = $true;
    #endregion

    #region begin tabPageSaved
    # Set tab page for "tabPageSaved" properties.
    $tabPageSaved.Text = 'Saved';
    $tabPageSaved.Name = 'tabPageSaved';
    $tabPageSaved.Location = New-Object Drawing.Point @(4, 24);
    $tabPageSaved.Size = New-Object Drawing.Size @(1140, 684);
    $tabPageSaved.Padding = New-Object System.Windows.Forms.Padding(3);
    $tabPageSaved.UseVisualStyleBackColor = $true;

    # Set splitContainerSaved properties.
    $splitContainerSaved.Name = 'splitContainerSaved';
    $splitContainerSaved.Location = New-Object Drawing.Point @(6, 6);
    $splitContainerSaved.Size = New-Object Drawing.Size @(1035, 672);
    $splitContainerSaved.SplitterDistance = 206;
    $splitContainerSaved.IsSplitterFixed = $false;

    # Set listBoxSaved properties.
    $listBoxSaved.Name = 'listBoxSaved';
    $listBoxSaved.Size = New-Object Drawing.Size @(206, 672);
    $listBoxSaved.ItemHeight = 15;
    $listBoxSaved.Location = New-Object Drawing.Point @(0, 0);
    $listBoxSaved.SelectionMode = 'One';
    $listBoxSaved.Dock = 'Fill';
    $listBoxSaved.FormattingEnabled = $true;

    # Set richTextBoxSaved properties.
    $richTextBoxSaved.Name = 'richTextBoxSaved';
    $richTextBoxSaved.Size = New-Object Drawing.Size @(825, 672);
    $richTextBoxSaved.Location = New-Object Drawing.Point @(0, 0);
    $richTextBoxSaved.ScrollBars = 'Vertical';
    $richTextBoxSaved.Dock = 'Fill';
    $richTextBoxSaved.Text = '';
    $richTextBoxSaved.ReadOnly = $true;

    # Set buttonSavedExport properties.
    $buttonSavedExport.Text = 'Export';
    $buttonSavedExport.Size = New-Object Drawing.Size @(87, 32);
    $buttonSavedExport.Location = New-Object Drawing.Point @(1047, 6);
    $buttonSavedExport.Name = 'buttonSavedExport';
    $buttonSavedExport.UseVisualStyleBackColor = $true;

    # Set buttonSavedImport properties.
    $buttonSavedImport.Text = 'Import';
    $buttonSavedImport.Size = New-Object Drawing.Size @(87, 32);
    $buttonSavedImport.Location = New-Object Drawing.Point @(1047, 44);
    $buttonSavedImport.Name = 'buttonSavedImport';
    $buttonSavedImport.UseVisualStyleBackColor = $true;

    # Set buttonSavedCopyToClipboard properties.
    $buttonSavedCopyToClipboard.Text = 'Copy to Clipboard';
    $buttonSavedCopyToClipboard.Size = New-Object Drawing.Size @(87, 32);
    $buttonSavedCopyToClipboard.Location = New-Object Drawing.Point @(1047, 82);
    $buttonSavedCopyToClipboard.Name = 'buttonSavedCopyToClipboard';
    $buttonSavedCopyToClipboard.UseVisualStyleBackColor = $true;
    #endregion

    #region begin tabPageSettings
    # Set tab page for "tabPageSettings" properties.
    $tabPageSettings.Text = 'Settings';
    $tabPageSettings.Name = 'tabPageSettings';
    $tabPageSettings.Location = New-Object Drawing.Point @(4, 24);
    $tabPageSettings.Size = New-Object Drawing.Size @(1140, 684);
    $tabPageSettings.Padding = New-Object System.Windows.Forms.Padding(3);
    $tabPageSettings.UseVisualStyleBackColor = $true;

    # Set groupBoxSettingsDelay properties.
    $groupBoxSettingsDelay.Text = 'Delays';
    $groupBoxSettingsDelay.Size = New-Object Drawing.Size @(300, 95);
    $groupBoxSettingsDelay.Location = New-Object Drawing.Point @(8, 6);

    # Set labelDelayWait properties.
    $labelDelayWait.Text = 'Delay before typing (seconds):';
    $labelDelayWait.Size = New-Object Drawing.Size @(167, 15);
    $labelDelayWait.Location = New-Object Drawing.Point @(6, 29);
    $labelDelayWait.Name = 'labelDelayWait';

    # Set labelDelayKey properties.
    $labelDelayKey.Text = 'Delay between each key (miliseconds):';
    $labelDelayKey.Size = New-Object Drawing.Size @(210, 15);
    $labelDelayKey.Location = New-Object Drawing.Point @(6, 60);
    $labelDelayKey.Name = 'labelDelayKey';

    # Set textBoxDelayWait properties.
    $textBoxDelayWait.Size = New-Object Drawing.Size @(67, 23);
    $textBoxDelayWait.Location = New-Object Drawing.Point @(222, 26);
    $textBoxDelayWait.Name = 'textBoxDelayWait';
    $textBoxDelayWait.Text = '5';

    # Set textBoxDelayKey properties.
    $textBoxDelayKey.Size = New-Object Drawing.Size @(67, 23);
    $textBoxDelayKey.Location = New-Object Drawing.Point @(222, 57);
    $textBoxDelayKey.Name = 'textBoxDelayKey';
    $textBoxDelayKey.Text = '5';
    #endregion

    #region begin AddControls
    # Add elements togroupBoxSettingsDelay.
    $groupBoxSettingsDelay.Controls.Add($labelDelayWait);
    $groupBoxSettingsDelay.Controls.Add($labelDelayKey);
    $groupBoxSettingsDelay.Controls.Add($textBoxDelayWait);
    $groupBoxSettingsDelay.Controls.Add($textBoxDelayKey);

    # Add elements to splitContainerSaved.
    $splitContainerSaved.Panel1.Controls.Add($listBoxSaved);
    $splitContainerSaved.Panel2.Controls.Add($richTextBoxSaved);

    # Add elements to tabPageInsert.
    $tabPageInsert.Controls.Add($richTextBoxInsert);
    $tabPageInsert.Controls.Add($buttonTypeSaveToFile);
    $tabPageInsert.Controls.Add($buttonTypeClear);
    $tabPageInsert.Controls.Add($buttonTypeSave);
    $tabPageInsert.Controls.Add($buttonTypeSend);
    $tabPageInsert.Controls.Add($buttonTypeCancel);

    # Add elements to tabPageSaved.
    $tabPageSaved.Controls.Add($buttonSavedExport);
    $tabPageSaved.Controls.Add($buttonSavedImport);
    $tabPageSaved.Controls.Add($buttonSavedCopyToClipboard);
    $tabPageSaved.Controls.Add($splitContainerSaved);

    # Add elements to tabPageSettings.
    $tabPageSettings.Controls.Add($groupBoxSettingsDelay);

    # Add elements to tabControl.
    $tabControl.TabPages.Add($tabPageInsert);
    $tabControl.TabPages.Add($tabPageSaved);
    $tabControl.TabPages.Add($tabPageSettings);

    # Add elements to contextMenuStrip.
    $notifyIcon.ContextMenuStrip = $contextMenuStrip;

    # Add elements to form.
    $form.Controls.Add($tabControl);
    $form.Controls.Add($statusStrip);
    #endregion

    #region begin AddEventHandlers
    # Add event handler for Open context menu item.
    $contextMenuItemOpen.Add_Click(
        {
            # Show the form.
            $form.Visible = $true;

            # Bring to the front.
            $form.BringToFront();

            # Set the form to normal state.
            $form.WindowState = 'Normal';
        }
    );

    # Add event handler for Exit context menu item.
    $contextMenuItemExit.Add_Click(
        {
            # Close the form.
            $null = $form.Close();

            # Dispose the form.
            $null = $form.Dispose();

            # Dispose the notify icon.
            $null = $notifyIcon.Dispose();
        }
    );

    # Add event handler for form closing.
    $form.Add_FormClosing(
        {
            # Close all runspace in the list.
            foreach ($runspace in $Script:runspaces)
            {
                # Close the runspace.
                $null = $runspace.Dispose();
            }
        }
    );

    # Add event handler for clicking on button Send.
    $buttonTypeSend.Add_Click(
        {
            # Get the text from the rich text box.
            $textRichTextBoxInsert = $richTextBoxInsert.Text;

            # Get the delay before typing.
            $delayWait = $textBoxDelayWait.Text;

            # Get the delay between each key.
            $delayKey = $textBoxDelayKey.Text;

            # Create a new runspace.
            $runspace = Get-PowerShellRunspace;

            # Create script block.
            $sendKeyScriptBlock = {
                [CmdletBinding()]
                param
                (
                    # The input string to send.
                    [Parameter(Mandatory = $true)]
                    [ValidateNotNullOrEmpty()]
                    [string]$InputString,

                    # The delay between each key press.
                    [Parameter(Mandatory = $false)]
                    [ValidateNotNullOrEmpty()]
                    [int]$DelayBeforeTypingInSeconds = 5,

                    # The delay between each key press.
                    [Parameter(Mandatory = $false)]
                    [ValidateNotNullOrEmpty()]
                    [int]$DelayInMiliseconds = 200,

                    # The reference to the form.
                    [Parameter(Mandatory = $true)]
                    [ValidateNotNullOrEmpty()]
                    [System.Windows.Forms.Form]$Form
                )

                # Loop for X seconds before typing.
                for ($i = 0; $i -lt $DelayBeforeTypingInSeconds; $i++)
                {
                    # Update the status.
                    $Form.Controls['statusStrip'].Items[0].Text = ('Waiting {0} seconds before typing' -f ($DelayBeforeTypingInSeconds - $i));

                    # Sleep for some seconds
                    Start-Sleep -Seconds 1;
                }

                # Update the status.
                $Form.Controls['statusStrip'].Items[0].Text = 'Typing now';

                # Send keyboard input.
                Send-KeyboardInput -InputString $InputString -DelayInMiliseconds $DelayInMiliseconds -DelayBeforeTypingInSeconds 0;

                # Update the status.
                $Form.Controls['statusStrip'].Items[0].Text = 'Ready';
            };

            # Add the script block with arguments to the runspace.
            $null = $runspace.AddScript($sendKeyScriptBlock);
            $null = $runspace.AddArgument([ref]$textRichTextBoxInsert);
            $null = $runspace.AddArgument([ref]$delayWait);
            $null = $runspace.AddArgument([ref]$delayKey);
            $null = $runspace.AddParameter('Form', $form);

            # Invoke the runspace(s).
            $null = $runspace.BeginInvoke();

            # Unlock the GUI.
            [System.Windows.Forms.Application]::DoEvents();

            # Add runspace to the list.
            $null = $Script:runspaces.Add($runspace);
        }
    );

    # Add event handler for clicking on button Clear.
    $buttonTypeClear.Add_Click(
        {
            # Clear the text from the rich text box.
            $richTextBoxInsert.Clear();
        }
    );

    # Make sure only integers are allowed in textBoxDelayWait.
    $textBoxDelayWait.Add_KeyPress(
        {
            if (-not [char]::IsControl($_.KeyChar) -and -not [char]::IsDigit($_.KeyChar))
            {
                $_.Handled = $true;
            }
        }
    );

    # Make sure only integers are allowed in the textBoxDelayKey.
    $textBoxDelayKey.Add_KeyPress(
        {
            if (-not [char]::IsControl($_.KeyChar) -and -not [char]::IsDigit($_.KeyChar))
            {
                $_.Handled = $true;
            }
        }
    );

    # Add event handler for clicking on button Save.
    $buttonTypeSave.Add_Click(
        {
            # Enter name of the saved item.
            $name = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a name for the saved item:', 'Save', (Get-Date).ToString('yyyyMMddHHmmss'));

            # If name is empty.
            if ([string]::IsNullOrEmpty($name))
            {
                # Show a message box.
                [System.Windows.Forms.MessageBox]::Show('Name cannot be empty.', 'Save', 'OK', 'Error');

                # Return from event handler.
                return [void];
            }

            # Generate ID.
            $id = [System.Guid]::NewGuid().ToString();

            # Full name.
            $fullName = ('{0} ({1})' -f $name, $id);

            # Get the text from the rich text box.
            $textRichTextBoxInsert = $richTextBoxInsert.Text;

            # Add to object array saved.
            $null = $Script:saved.Add(
                [PSCustomObject]@{
                    Id       = $id;
                    FullName = $fullName;
                    Name     = $name;
                    Text     = $textRichTextBoxInsert;
                    DateTime = (Get-Date);
                }
            );

            # Add to list box.
            $null = $listBoxSaved.Items.Add($fullName);

            # Set the text in the status strip.
            $statusStripLabel.Text = ("Saved item '{0}'" -f $name);
        }
    );

    # Add event handler when selecting an item in the list box.
    $listBoxSaved.Add_SelectedIndexChanged(
        {
            # Get the selected item.
            $selectedItem = $listBoxSaved.SelectedItem;

            # Get the object from the saved array.
            $selectedObject = $Script:saved | Where-Object { $_.FullName -eq $selectedItem };

            # Set the text in the rich text box.
            $richTextBoxSaved.Text = $selectedObject.Text;
        }
    );

    # Add event handler for clicking on button Save To File.
    $buttonTypeSaveToFile.Add_Click(
        {
            # Create a save file dialog.
            $saveFileDialog = New-Object -TypeName 'System.Windows.Forms.SaveFileDialog';

            # Set the title.
            $saveFileDialog.Title = 'Save to File';

            # Set the filter.
            $saveFileDialog.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*';

            # Set the initial directory.
            $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop');

            # Show the dialog.
            $saveFileDialog.ShowDialog();

            # Get the file name.
            $fileName = $saveFileDialog.FileName;

            # Get the text from the rich text box.
            $textRichTextBoxInsert = $richTextBoxInsert.Text;

            # Save the text to the file.
            $textRichTextBoxInsert | Set-Content -Path $fileName -Force -Encoding UTF8;

            # Set the text in the status strip.
            $statusStripLabel.Text = ("Saved input to file '{0}'" -f $fileName);
        }
    );

    # Add event handler for clicking on button Export.
    $buttonSavedExport.Add_Click(
        {
            # If saved array is empty.
            if ($Script:saved.Count -eq 0)
            {
                # Show a message box.
                [System.Windows.Forms.MessageBox]::Show('There is nothing to export.', 'Export', 'OK', 'Information');

                # Return from event handler.
                return [void];
            }

            # Create a save file dialog.
            $saveFileDialog = New-Object -TypeName 'System.Windows.Forms.SaveFileDialog';

            # Set the title.
            $saveFileDialog.Title = 'Export';

            # Set the filter.
            $saveFileDialog.Filter = 'AutoTyper (*.autotyper)|*.autotyper|All files (*.*)|*.*';

            # Set the initial directory.
            $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop');

            # Show the dialog.
            $saveFileDialog.ShowDialog();

            # Get the file name.
            $fileName = $saveFileDialog.FileName;

            # Convert the saved array to JSON.
            $jsonSaved = $Script:saved | ConvertTo-Json -Depth 100;

            # Try to save the file.
            Try
            {
                # Save to the file.
                $jsonSaved | Set-Content -Path $fileName -Force -Encoding UTF8;

                # Show a message box.
                [System.Windows.Forms.MessageBox]::Show('Exported successfully.', 'Export', 'OK', 'Information');

                # Set the text in the status strip.
                $statusStripLabel.Text = ("Exported saved items to file '{0}'" -f $fileName);
            }
            catch
            {
                # Show a message box.
                [System.Windows.Forms.MessageBox]::Show('Export failed.', 'Export', 'OK', 'Error');
            }
        }
    );

    # Add event handler for clicking on button Import.
    $buttonSavedImport.Add_Click(
        {
            # Create an open file dialog.
            $openFileDialog = New-Object -TypeName 'System.Windows.Forms.OpenFileDialog';

            # Set the title.
            $openFileDialog.Title = 'Import';

            # Set the filter.
            $openFileDialog.Filter = 'AutoTyper (*.autotyper)|*.autotyper|All files (*.*)|*.*';

            # Set the initial directory.
            $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop');

            # Show the dialog.
            $openFileDialog.ShowDialog();

            # Get the file name.
            $fileName = $openFileDialog.FileName;

            # If the file name is empty.
            if ([string]::IsNullOrEmpty($fileName))
            {
                # Return from event handler.
                return [void];
            }

            # Try to read the file.
            Try
            {
                # Read the file.
                $jsonSaved = Get-Content -Path $fileName -Raw;

                # Convert the JSON to an object array.
                $saved = $jsonSaved | ConvertFrom-Json;

                # Clear the list box.
                $listBoxSaved.Items.Clear();

                # Clear the saved array.
                $Script:saved.Clear();

                # Add the items to the list box.
                foreach ($item in $saved)
                {
                    # If all the properties are not present.
                    if (-not $item.PSObject.Properties.Match('Id') -or
                        -not $item.PSObject.Properties.Match('FullName') -or
                        -not $item.PSObject.Properties.Match('Name') -or
                        -not $item.PSObject.Properties.Match('Text') -or
                        -not $item.PSObject.Properties.Match('DateTime'))
                    {
                        # Continue to the next item.
                        continue;
                    }

                    # Add to object array saved.
                    $null = $Script:saved.Add(
                        [PSCustomObject]@{
                            Id       = $item.Id;
                            FullName = $item.FullName;
                            Name     = $item.Name;
                            Text     = $item.Text;
                            DateTime = $item.DateTime;
                        }
                    );

                    # Add to list box.
                    $null = $listBoxSaved.Items.Add($item.FullName);
                }

                # Show a message box.
                [System.Windows.Forms.MessageBox]::Show('Imported successfully.', 'Import', 'OK', 'Information');

                # Set the text in the status strip.
                $statusStripLabel.Text = ("Imported saved items successfully from file '{0}'" -f $fileName);
            }
            catch
            {
                # Show a message box.
                [System.Windows.Forms.MessageBox]::Show('Import failed.', 'Import', 'OK', 'Error');
            }
        }
    );

    # Add event handler for clicking on button Copy to Clipboard.
    $buttonSavedCopyToClipboard.Add_Click(
        {
            # Get the selected item.
            $selectedItem = $listBoxSaved.SelectedItem;

            # Get the object from the saved array.
            $selectedObject = $Script:saved | Where-Object { $_.FullName -eq $selectedItem };

            # If the selected object is null.
            if (-not $selectedObject)
            {
                # Show a message box.
                [System.Windows.Forms.MessageBox]::Show('No item selected.', 'Copy to Clipboard', 'OK', 'Error');

                # Return from event handler.
                return [void];
            }

            # Set the text in the status strip.
            $statusStripLabel.Text = 'Copied to clipboard';

            # Set the text in the clipboard.
            [System.Windows.Forms.Clipboard]::SetText($selectedObject.Text);

            # Show a message box.
            [System.Windows.Forms.MessageBox]::Show('Copied to clipboard.', 'Copy to Clipboard', 'OK', 'Information');
        }
    );

    # Add event handler for clicking on button Cancel.
    $buttonTypeCancel.Add_Click(
        {
            # Set the text in the status strip.
            $statusStripLabel.Text = 'Cancelled';

            # Close all runspace in the list.
            foreach ($runspace in $Script:runspaces)
            {
                # Close the runspace.
                $runspace.Dispose();
            }

            # Run garbage collection.
            [System.GC]::Collect();
        }
    );
    #endregion

    # Force garbage collection just to start slightly lower RAM usage.
    [System.GC]::Collect();

    # Bring the form to the front.
    $form.BringToFront();

    # Always show tray icon.
    $notifyIcon.Visible = $true;

    # Set app id for window.
    $null = [PSAppID]::SetAppIdForWindow($form.Handle, 'AutoTyper.App');

    # Hide console.
    $null = Hide-Console;

    # Display the form.
    $null = $form.ShowDialog();

    # Close all runspace in the list.
    foreach ($runspace in $Script:runspaces)
    {
        # Close the runspace.
        $null = $runspace.Dispose();
    }

    # Dispose the notify icon.
    $notifyIcon.Dispose();

    # Dispose the form.
    $form.Dispose();
}

############### Functions - End ###############
#endregion

#region begin main
############### Main - Start ###############

# Present form.
Show-Form;

############### Main - End ###############
#endregion

#region begin finalize
############### Finalize - Start ###############

############### Finalize - End ###############
#endregion
