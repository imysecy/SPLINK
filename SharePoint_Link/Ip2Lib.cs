using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace com.softwarekey.ClientLib.InstantPLUS
{
    /// <summary>Class for calling the native 32 bit Instant PLUS DLL.</summary>
    public class Ip2Lib32Managed
    {
        /// <summary>DEPRECATED: Use Ip2LibManaged for flags instead.</summary>
        public const Int32 FLAGS_NONE = Ip2LibManaged.FLAGS_NONE;
        public const Int32 FLAGS_USE_THREAD = Ip2LibManaged.FLAGS_USE_THREAD;
        public const Int32 RESULT_FILE_NOT_FOUND = Ip2LibManaged.RESULT_FILE_NOT_FOUND;
        public const Int32 RESULT_CANNOT_CREATE_THREAD = Ip2LibManaged.RESULT_CANNOT_CREATE_THREAD;
        public const Int32 RESULT_INCOMPATIBLE_FILE_VERSION = Ip2LibManaged.RESULT_INCOMPATIBLE_FILE_VERSION;

        /// <summary>
        /// (DEPRECATED: Use Ip2LibManaged instead) Makes the call to the native Instant PLUS library.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="key">Decryption Key.</param>
        /// <param name="path">Path to XML config file.</param>
        /// <returns>Int32</returns>
        public static Int32 CallInstantPLUS(Int32 flags, string key, string path)
        {
            return Ip2LibManaged.CallInstantPLUS(flags, key, path);
        }

        [DllImport("IP2Lib32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 CallIpEx(IntPtr hWnd, Int32 flags, string key, string path, string res1, string res2, string res3, Int32 res4, Int32 res5, Int32 res6);

        [DllImport("IP2Lib32.dll", EntryPoint = "n1", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFGetString(Int32 flags, Int32 var_no, StringBuilder buffer);

        [DllImport("IP2Lib32.dll", EntryPoint = "n2", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFGetNum(Int32 flags, Int32 var_no, ref int value);

        [DllImport("IP2Lib32.dll", EntryPoint = "n3", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFGetDate(Int32 flags, Int32 var_no, ref int month_hours, ref int day_minutes, ref int year_seconds);

        [DllImport("IP2Lib32.dll", EntryPoint = "n4", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFSetString(Int32 flags, Int32 var_no, string buffer);

        [DllImport("IP2Lib32.dll", EntryPoint = "n5", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFSetNum(Int32 flags, Int32 var_no, Int32 value);

        [DllImport("IP2Lib32.dll", EntryPoint = "n6", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFSetDate(Int32 flags, Int32 var_no, Int32 month_hours, Int32 day_minutes, Int32 year_seconds);

        [DllImport("IP2Lib32.dll", EntryPoint = "n7", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_Activate(Int32 flags, Int32 action, Int32 action_flags, IntPtr hWnd);

        [DllImport("IP2Lib32.dll", EntryPoint = "n8", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_TrialRunsLeft(Int32 flags, ref int runsleft);

        [DllImport("IP2Lib32.dll", EntryPoint = "n9", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_TrialDaysLeft(Int32 flags, ref int daysleft);

        [DllImport("IP2Lib32.dll", EntryPoint = "n10", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_IsTrial(Int32 flags);

        [DllImport("IP2Lib32.dll", EntryPoint = "n17", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr WR_GetLFHandle(ref Int32 lRetVal);

        [DllImport("IP2Lib32.dll", EntryPoint = "n18", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetLicenseID(ref int licenseid);

        [DllImport("IP2Lib32.dll", EntryPoint = "n19", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetPassword(StringBuilder password);

        [DllImport("IP2Lib32.dll", EntryPoint = "n25", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_DeactivateLicense(Int32 flags);

        [DllImport("IP2Lib32.dll", EntryPoint = "r1", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_Close(Int32 flags);

        [DllImport("IP2Lib32.dll", EntryPoint = "r2", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_IsJustActivated(Int32 flags);

        [DllImport("IP2Lib32.dll", EntryPoint = "r4", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetProductID(Int32 flags, ref int number);


        [DllImport("IP2Lib32.dll", EntryPoint = "r5", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetProdOptionID(Int32 flags, ref int number);

        [DllImport("IP2Lib32.dll", EntryPoint = "r14", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_ExpireType(Int32 flags, StringBuilder str1);

        [DllImport("IP2Lib32.dll", EntryPoint = "r15", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetPurchaseUrl(Int32 flags, StringBuilder str1);

        [DllImport("IP2Lib32.dll", EntryPoint = "r3", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_DeactivateInstallation(Int32 flags, IntPtr hWnd);
    }

    /// <summary>Class for calling the native 64 bit Instant PLUS DLL</summary>
    public class Ip2Lib64Managed
    {
        [DllImport("IP2Lib64.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 CallIpEx(IntPtr hWnd, Int32 flags, string key, string path, string res1, string res2, string res3, Int32 res4, Int32 res5, Int32 res6);

        [DllImport("IP2Lib64.dll", EntryPoint = "n1", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFGetString(Int32 flags, Int32 var_no, StringBuilder buffer);

        [DllImport("IP2Lib64.dll", EntryPoint = "n2", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFGetNum(Int32 flags, Int32 var_no, ref int value);

        [DllImport("IP2Lib64.dll", EntryPoint = "n3", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFGetDate(Int32 flags, Int32 var_no, ref int month_hours, ref int day_minutes, ref int year_seconds);

        [DllImport("IP2Lib64.dll", EntryPoint = "n4", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFSetString(Int32 flags, Int32 var_no, string buffer);

        [DllImport("IP2Lib64.dll", EntryPoint = "n5", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFSetNum(Int32 flags, Int32 var_no, Int32 value);

        [DllImport("IP2Lib64.dll", EntryPoint = "n6", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_LFSetDate(Int32 flags, Int32 var_no, Int32 month_hours, Int32 day_minutes, Int32 year_seconds);

        [DllImport("IP2Lib64.dll", EntryPoint = "n7", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_Activate(Int32 flags, Int32 action, Int32 action_flags, IntPtr hWnd);

        [DllImport("IP2Lib64.dll", EntryPoint = "n8", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_TrialRunsLeft(Int32 flags, ref int runsleft);

        [DllImport("IP2Lib64.dll", EntryPoint = "n9", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_TrialDaysLeft(Int32 flags, ref int daysleft);

        [DllImport("IP2Lib64.dll", EntryPoint = "n10", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_IsTrial(Int32 flags);

        [DllImport("IP2Lib64.dll", EntryPoint = "n17", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr WR_GetLFHandle(ref Int32 lRetVal);

        [DllImport("IP2Lib64.dll", EntryPoint = "n18", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetLicenseID(ref Int32 licenseid);

        [DllImport("IP2Lib64.dll", EntryPoint = "n19", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetPassword(StringBuilder password);

        [DllImport("IP2Lib64.dll", EntryPoint = "n25", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_DeactivateLicense(Int32 flags);

        [DllImport("IP2Lib64.dll", EntryPoint = "r1", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_Close(Int32 flags);

        [DllImport("IP2Lib64.dll", EntryPoint = "r2", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_IsJustActivated(Int32 flags);

        [DllImport("IP2Lib64.dll", EntryPoint = "r4", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetProductID(Int32 flags, ref int number);

        [DllImport("IP2Lib64.dll", EntryPoint = "r5", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetProdOptionID(Int32 flags, ref int number);

        [DllImport("IP2Lib64.dll", EntryPoint = "r14", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_ExpireType(Int32 flags, StringBuilder str1);

        [DllImport("IP2Lib64.dll", EntryPoint = "r15", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_GetPurchaseUrl(Int32 flags, StringBuilder str1);

        [DllImport("IP2Lib64.dll", EntryPoint = "r3", CallingConvention = CallingConvention.StdCall)]
        public static extern Int32 WR_DeactivateInstallation(Int32 flags, IntPtr hWnd);
    }

    public class Ip2LibManaged
    {
        /// <summary>Target processor architecture enumeration.</summary>
        public enum TargetArchitecture
        {
            x86 = 0,
            x64 = 1
        };

        /// <summary>Constant flags.</summary>
        public const Int32 FLAGS_NONE = 0;
        public const Int32 FLAGS_USE_THREAD = 1;
        public const Int32 RESULT_FILE_NOT_FOUND = 7;
        public const Int32 RESULT_CANNOT_CREATE_THREAD = -99;
        public const Int32 RESULT_INCOMPATIBLE_FILE_VERSION = -98;

        /// <summary>
        ///  Constants for exception messages
        /// </summary>
        private const string EXCEPTION_STARTUP_TITLE = "Start-up Error";
        private const string EXCEPTION_STARTUP_BADIMAGEFORMAT_MESSAGE = "The application encountered an error during start-up.  This error is most likely occurring because the application has not been configured to target only x86 architecture.  Please contact technical support for assistance.  The error description is as follows:\n\n";
        private const string EXCEPTION_STARTUP_DLLNOTFOUNDEXCEPTION = "The application encountered an error during start-up.  This error is most likely occurring because a file this program requires to run correctly is not present.  Please contact technical support for assistance.  The error description is as follows:\n\n";
        private const string EXCEPTION_STARTUP_OTHER = "The application encountered an error during start-up.  Please contact technical support for assistance.  The error description is as follows:\n\n";
        private const string EXCEPTION_STARTUP_RESULT_FILE_NOT_FOUND = "The application encountered an error during start-up because a required file is missing.  Please contact technical support for further assistance.";
        private const string EXCEPTION_STARTUP_INCOMPATIBLE_FILE_VERSION = "Incompatible XML file version";
        private const string EXCEPTION_RUNTIME_TITLE = "Error";
        private const string EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE = "The application encountered an error.  This error is most likely occurring because the application has not been configured to target only x86 architecture.  Please contact technical support for assistance.  The error description is as follows:\n\n";
        private const string EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION = "The application encountered an error.  This error is most likely occurring because a file this program requires to run correctly is not present.  Please contact technical support for assistance.  The error description is as follows:\n\n";
        private const string EXCEPTION_RUNTIME_OTHER = "The application encountered an error.  Please contact technical support for assistance.  The error description is as follows:\n\n";

        /// <summary>
        ///  Makes the call to the native Instant PLUS library. Use CallInstantPLUSEx() for additional arguments.
        /// </summary>
        /// <param name="flags">Flags.</param>
        /// <param name="key">Decryption Key.</param>
        /// <param name="path">Path to XML config file.</param>
        /// <returns>Int32.</returns>
        public static Int32 CallInstantPLUS(Int32 flags, string key, string path)
        {
            return CallInstantPLUSEx(IntPtr.Zero, flags, key, path, "", "", "", 0, 0, 0);
        }

        /// <summary>
        /// Makes the call to the native Instant PLUS library.
        /// </summary>
        /// <param name="hWnd">A handle to the application's window.</param>
        /// <param name="flags">Flags.</param>
        /// <param name="key">Decryption Key.</param>
        /// <param name="path">Path to XML config file.</param>
        /// <param name="res1">Reserved.</param>
        /// <param name="res2">Reserved.</param>
        /// <param name="res3">Reserved.</param>
        /// <param name="res4">Reserved.</param>
        /// <param name="res5">Reserved.</param>
        /// <param name="res6">Reserved.</param>
        /// <returns>Int32.</returns>
        public static Int32 CallInstantPLUSEx(IntPtr hWnd, Int32 flags, string key, string path, string res1, string res2, string res3, Int32 res4, Int32 res5, Int32 res6)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.CallIpEx(hWnd, flags, key, path, res1, res2, res3, res4, res5, res6);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.CallIpEx(hWnd, flags, key, path, res1, res2, res3, res4, res5, res6);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_STARTUP_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_STARTUP_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_STARTUP_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_STARTUP_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_STARTUP_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_STARTUP_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (RESULT_FILE_NOT_FOUND == result)
            {
                MessageBox.Show(EXCEPTION_STARTUP_RESULT_FILE_NOT_FOUND.Replace("\\n", Environment.NewLine), EXCEPTION_STARTUP_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (RESULT_INCOMPATIBLE_FILE_VERSION == result)
            {
                MessageBox.Show(EXCEPTION_STARTUP_INCOMPATIBLE_FILE_VERSION.Replace("\\n", Environment.NewLine), EXCEPTION_STARTUP_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// Determines if the current assembly is running under a 32 bit process or a 64 bit process.
        /// </summary>
        /// <returns>TargetArchitecture.</returns>
        public static TargetArchitecture GetTargetArchitecture()
        {
            TargetArchitecture arch = TargetArchitecture.x86;

            switch (IntPtr.Size)
            {
                case 8:
                    arch = TargetArchitecture.x64;
                    break;
                case 4:
                default:
                    arch = TargetArchitecture.x86;
                    break;
            }

            return arch;
        }

        /// <summary>
        /// This function is used to call into the wrapper from the application in order to perform an activation.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="action">Reserved - Always use 1.</param>
        /// <param name="action_flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="hWnd">A handle to the application's window.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_Activate(Int32 flags, Int32 action, Int32 action_flags, IntPtr hWnd)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_Activate(flags, action, action_flags, hWnd);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_Activate(flags, action, action_flags, hWnd);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// Shuts down the Instant PLUS API.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_Close(Int32 flags)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_Close(flags);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_Close(flags);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to obtain the string values of the variables stored in the License File. 
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="var_no">Determines which variable in the license file is being retrieved.</param>
        /// <param name="buffer">The location in which to place the variable data.</param>
        /// <returns>Int32.</returns>

        public static Int32 WR_LFGetString(Int32 flags, Int32 var_no, StringBuilder buffer)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_LFGetString(flags, var_no, buffer);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_LFGetString(flags, var_no, buffer);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// Use this function to obtain the values of the numeric variables stored in the License File. 
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="var_no">Determines which variable in the license file is being retrieved.</param>
        /// <param name="value">The location in which to place the variable data.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_LFGetNum(Int32 flags, Int32 var_no, ref int value)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_LFGetNum(flags, var_no, ref value);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_LFGetNum(flags, var_no, ref value);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to obtain the values of the date and time variables stored in the License File.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="var_no">Determines which variable is being retrieved.</param>
        /// <param name="month_hours">Returns the month or hours of the variable.</param>
        /// <param name="day_minutes">Returns the day or minutes of the variable.</param>
        /// <param name="year_seconds">Returns the year or seconds of the variable.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_LFGetDate(Int32 flags, Int32 var_no, ref int month_hours, ref int day_minutes, ref int year_seconds)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_LFGetDate(flags, var_no, ref month_hours, ref day_minutes, ref year_seconds);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_LFGetDate(flags, var_no, ref month_hours, ref day_minutes, ref year_seconds);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// This function is used to set the string values of the variables stored in the License File.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="var_no">Determines which variable in the license file is being set.</param>
        /// <param name="buffer">The value to which to set the License File variable.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_LFSetString(Int32 flags, Int32 var_no, string buffer)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_LFSetString(flags, var_no, buffer);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_LFSetString(flags, var_no, buffer);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to set the values of the numeric variables in the License File. 
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="var_no">Determines which variable in the license file is being set.</param>
        /// <param name="value">The value to write into the license file.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_LFSetNum(Int32 flags, Int32 var_no, Int32 value)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_LFSetNum(flags, var_no, value);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_LFSetNum(flags, var_no, value);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// This function is used to change the date and time values of the variables stored in the License File.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="var_no">Determines which variable is being set.</param>
        /// <param name="month_hours">Sets the month or hours of the variable in the license file.</param>
        /// <param name="day_minutes">Sets the day or minutes of the variable in the license file.</param>
        /// <param name="year_seconds">Sets the year or seconds of the variable in the license file.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_LFSetDate(Int32 flags, Int32 var_no, Int32 month_hours, Int32 day_minutes, Int32 year_seconds)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_LFSetDate(flags, var_no, month_hours, day_minutes, year_seconds);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_LFSetDate(flags, var_no, month_hours, day_minutes, year_seconds);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// This function retrieves the number of executions remaining in a trial. 
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="runsleft">Variable to place the number of executions remaining in the trial.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_TrialRunsLeft(Int32 flags, ref int runsleft)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_TrialRunsLeft(flags, ref runsleft);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_TrialRunsLeft(flags, ref runsleft);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// This function retrieves the number of days remaining in a trial.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="daysleft">Variable to place the number of days remaining in the trial.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_TrialDaysLeft(Int32 flags, ref int daysleft)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_TrialDaysLeft(flags, ref daysleft);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_TrialDaysLeft(flags, ref daysleft);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// This function returns the trial status of an application.  
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_IsTrial(Int32 flags)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_IsTrial(flags);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_IsTrial(flags);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to retrieve the handle to the license file.  This handle may then be used with any of the Protection PLUS API functions.
        /// </summary>
        /// <param name="lfhandle">The output parameter for result.</param>
        /// <returns>IntPtr.</returns>
        public static IntPtr WR_GetLFHandle(ref Int32 lRetVal)
        {
            IntPtr lfhandle = IntPtr.Zero;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        lfhandle = Ip2Lib64Managed.WR_GetLFHandle(ref lRetVal);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        lfhandle = Ip2Lib32Managed.WR_GetLFHandle(ref lRetVal);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return lfhandle;
        }

        /// <summary>
        /// Retrieves the License ID from the license file.  If the application has yet to be activated the License ID will be set to 0.
        /// </summary>
        /// <param name="licenseid">Variable to place the License ID.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_GetLicenseID(ref Int32 licenseid)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_GetLicenseID(ref licenseid);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_GetLicenseID(ref licenseid);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Retrieves the activation password from the license file.  If the application has yet to be activated the password will be blank.
        /// </summary>
        /// <param name="password">Variable to place the activation password.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_GetPassword(StringBuilder password)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_GetPassword(password);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_GetPassword(password);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// This function will reset a license to a deactivated status.  Upon next execution, the user will need to reactive the application.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_DeactivateLicense(Int32 flags)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_DeactivateLicense(flags);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_DeactivateLicense(flags);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to determine if the application was just activated during the current execution.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_IsJustActivated(Int32 flags)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_IsJustActivated(flags);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_IsJustActivated(flags);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to retrieve the SOLO Product ID from the wrapper.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="number">Variable to place the SOLO Product ID.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_GetProductID(Int32 flags, ref int number)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_GetProductID(flags, ref number);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_GetProductID(flags, ref number);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to retrieve the SOLO Product Option ID from the wrapper.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="number">Variable to place the SOLO Product Option ID.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_GetProdOptionID(Int32 flags, ref int number)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_GetProdOptionID(flags, ref number);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_GetProdOptionID(flags, ref number);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Use this function to retrieve the expiration type from the license file. 
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="str1">Variable to place the expire type of the application.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_ExpireType(Int32 flags, StringBuilder str1)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_ExpireType(flags, str1);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_ExpireType(flags, str1);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Retrieves the url used to purchase the application.
        /// </summary>
        /// <param name="flags">Reserved - Always use zero or FLAGS_NONE.</param>
        /// <param name="str1">Variable to place the url used for purchasing.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_GetPurchaseUrl(Int32 flags, StringBuilder str1)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_GetPurchaseUrl(flags, str1);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_GetPurchaseUrl(flags, str1);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
        /// <summary>
        /// Deactivates installation using SOLO Server's XML Activation Service.
        /// </summary>
        /// <param name="flags">Flags.</param>
        /// <param name="hwnd">A handle to the application's window.</param>
        /// <returns>Int32.</returns>
        public static Int32 WR_DeactivateInstallation(Int32 flags, IntPtr hWnd)
        {
            Int32 result = 0;

            try
            {
                switch (GetTargetArchitecture())
                {
                    case TargetArchitecture.x64:
                        //Run using the 64 bit DLL
                        result = Ip2Lib64Managed.WR_DeactivateInstallation(flags, hWnd);
                        break;
                    case TargetArchitecture.x86:
                    default:
                        //Run using the 32 bit DLL
                        result = Ip2Lib32Managed.WR_DeactivateInstallation(flags, hWnd);
                        break;
                }
            }
            catch (BadImageFormatException bEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_BADIMAGEFORMAT_MESSAGE.Replace("\\n", Environment.NewLine) + bEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DllNotFoundException dEx)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_DLLNOTFOUNDEXCEPTION.Replace("\\n", Environment.NewLine) + dEx.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(EXCEPTION_RUNTIME_OTHER.Replace("\\n", Environment.NewLine) + ex.Message, EXCEPTION_RUNTIME_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
    }
}