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