using System;
using System.Runtime.InteropServices;

namespace Interfaces
{
    /// <summary>
    /// <c>IAuthenticate</c> Interface
    /// Implements User Authentication  Interface
    /// </summary>
    [ComImport, GuidAttribute("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown),
    ComVisible(false)]
    public interface IAuthenticate
    {
        /// <summary>
        /// <c>Authenticate</c>
        /// </summary>
        /// <param name="phwnd"></param>
        /// <param name="pszUsername"></param>
        /// <param name="pszPassword"></param>
        /// <returns></returns>
        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int Authenticate(ref IntPtr phwnd,
        ref IntPtr pszUsername,
        ref IntPtr pszPassword
        );
    } 
}
