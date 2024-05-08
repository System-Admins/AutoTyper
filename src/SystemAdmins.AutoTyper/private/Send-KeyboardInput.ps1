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