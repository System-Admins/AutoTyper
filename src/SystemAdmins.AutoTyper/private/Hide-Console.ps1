# Hide the console window.
function Hide-Console
{
    # Get the console window.
    $consolePtr = [Console.Window]::GetConsoleWindow();

    # Hide the console window.
    [void][Console.Window]::ShowWindow($consolePtr, 0);
}