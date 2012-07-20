using System;
using System.Windows.Forms;
namespace Utility
{
    /// <summary>
    /// <c>EncodingAndDecoding</c> class
    /// implements the functions to encode and decode the data
    /// it also implements the showmessage function to display message to user
    /// </summary>
    static class EncodingAndDecoding
    {
        #region This Will Encode The Given String

        /// <summary>
        /// <c>Base64Encode</c> member function
        /// This function encodes data in base64
        /// </summary>
        /// <param name="sData"></param>
        /// <returns></returns>
        public static string Base64Encode(string sData)
        {
            try
            {
                byte[] encData_byte = new byte[sData.Length];
                encData_byte = System.Text.Encoding.UTF8.GetBytes(sData);
                string encodedData = Convert.ToBase64String(encData_byte);
                return encodedData;
            }
            catch (Exception ex)
            {
                throw new Exception("Error in base64Encode" + ex.Message);
            }


        }
        #endregion

        #region This Will Decode The Given String

        /// <summary>
        /// <c>Base64Decode</c> member function
        /// this member function decodes  the specific string from base64
        /// </summary>
        /// <param name="sData"></param>
        /// <returns></returns>
        public static string Base64Decode(string sData)
        {
            System.Text.UTF8Encoding encoder = new System.Text.UTF8Encoding();
            System.Text.Decoder utf8Decode = encoder.GetDecoder();
            byte[] todecode_byte = Convert.FromBase64String(sData);
            int charCount = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length);
            char[] decoded_char = new char[charCount];
            utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0);
            string result = new String(decoded_char);
            return result;
        }
        #endregion


        /// <summary>
        /// <c>ShowMessageBox</c> member function
        /// this member function display message to user
        /// </summary>
        /// <param name="callingMethodName"></param>
        /// <param name="message"></param>
        /// <param name="icon"></param>
        public static void ShowMessageBox(string callingMethodName, string message, MessageBoxIcon icon)
        {
            //if (icon == MessageBoxIcon.Error)
            //{
            //    MessageBox.Show(callingMethodName);
            //}

            MessageBox.Show(message, "ITOPIA", MessageBoxButtons.OK, icon);
        }
    }
}
