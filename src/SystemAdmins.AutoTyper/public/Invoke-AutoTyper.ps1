function Invoke-AutoTyper
{
    <#
    .SYNOPSIS
    A small WinForm program made in PowerShell to send key strokes to the active window.

    .DESCRIPTION
    This program uses runspaces and WinForm to automatically send keystrokes to the active window. The program can be used to send keystrokes to a game, a chat window, or any other application that accepts keyboard input.
    Usually it is used to automate repetitive tasks, such as typing configuration, passwords and other information into interfaces such as iLO, RDP etc.

    .Example
    Invoke-AutoTyper

    .NOTES
    Version:        1.0
    Author:         Alex Hansen (ath@systemadmins.com)
    Creation Date:  07-05-2024
    Purpose/Change: Initial script development.
    #>

    [cmdletbinding()]
    param
    (
    )

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

    # Object array to save the input.
    $Script:saved = New-Object System.Collections.ArrayList;

    # Object array to store runspaces.
    $Script:runspaces = New-Object System.Collections.ArrayList;

    # Present form.
    Show-Form;
}