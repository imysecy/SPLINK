using System;
using System.Runtime.InteropServices;

namespace Interfaces
{
    [ComImport,
     GuidAttribute("6d5140c1-7436-11ce-8034-00aa006009fa"),
     InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown),
     ComVisible(false)]
    public interface IServiceProvider
    {
        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int QueryService(ref Guid guidService, ref Guid riid, out IntPtr
        ppvObject);
    } 
}
