using System;
using System.Text;
using System.Security.Cryptography;

namespace SharePoint_Link.Utility
{
    /// <summary>
    /// <c>HashingClass</c> class
    /// HashingClass implements properties and functions to calculate hash values from data
    /// </summary>
    class HashingClass
    {

        /// <summary>
        /// <c>hashedemailbody</c> member field of type string
        /// holds hashed value of email body
        /// </summary>
        private static string hashedemailbody = string.Empty;

        /// <summary>
        /// <c>mailsubject</c> member field of type string
        /// holds hashed value of email subject
        /// </summary>
        private static string mailsubject = string.Empty;

        /// <summary>
        /// <c>modifieddate</c> member field
        /// holds  hashed value email date 
        /// </summary>
        private static string modifieddate = string.Empty;

        /// <summary>
        /// <c>Modifieddate</c> member property
        /// encapsulates  modifieddate
        /// </summary>
        public static string Modifieddate
        {
            get { return HashingClass.modifieddate; }
            set { HashingClass.modifieddate = value; }
        }

        /// <summary>
        /// <c>Mailsubject</c>member property 
        /// encapsulates mailsubject member field
        /// </summary>
        public static string Mailsubject
        {
            get { return HashingClass.mailsubject; }
            set { HashingClass.mailsubject = value; }
        }

        /// <summary>
        /// <c>Hashedemailbody</c> member property 
        /// encapsulates  hashedemailbody
        /// </summary>
        public static string Hashedemailbody
        {
            get { return HashingClass.hashedemailbody; }
            set { HashingClass.hashedemailbody = value; }
        }


        /// <summary>
        /// <c>ComputeHash</c> member function
        /// calculates hash value
        /// </summary>
        /// <param name="plainText"></param>
        /// <param name="hashAlgorithm"></param>
        /// <param name="saltBytes"></param>
        /// <returns></returns>
        public static string ComputeHash(string plainText, string hashAlgorithm, byte[] saltBytes)
        {
            // If salt is not specified, generate it on the fly.
            if (saltBytes == null)
            {
                // Define min and max salt sizes.
                int minSaltSize = 4;
                int maxSaltSize = 8;

                // Generate a random number for the size of the salt.
                Random random = new Random();
                int saltSize = random.Next(minSaltSize, maxSaltSize);

                // Allocate a byte array, which will hold the salt.
                saltBytes = new byte[saltSize];

                // Initialize a random number generator.
                RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider();

                // Fill the salt with cryptographically strong byte values.
                rng.GetNonZeroBytes(saltBytes);

            }

            // Convert plain text into a byte array.
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            // Allocate array, which will hold plain text and salt.
            byte[] plainTextWithSaltBytes =
                    new byte[plainTextBytes.Length + saltBytes.Length];

            // Copy plain text bytes into resulting array.
            for (int i = 0; i < plainTextBytes.Length; i++)
                plainTextWithSaltBytes[i] = plainTextBytes[i];

            // Append salt bytes to the resulting array.
            for (int i = 0; i < saltBytes.Length; i++)
                plainTextWithSaltBytes[plainTextBytes.Length + i] = saltBytes[i];

            // Because we support multiple hashing algorithms, we must define
            // hash object as a common (abstract) base class. We will specify the
            // actual hashing algorithm class later during object creation.
            HashAlgorithm hash;

            // Make sure hashing algorithm name is specified.
            if (hashAlgorithm == null)
                hashAlgorithm = "";

            // Initialize appropriate hashing algorithm class.
            switch (hashAlgorithm.ToUpper())
            {
                case "SHA1":
                    hash = new SHA1Managed();
                    break;

                case "SHA256":
                    hash = new SHA256Managed();
                    break;

                case "SHA384":
                    hash = new SHA384Managed();
                    break;

                case "SHA512":
                    hash = new SHA512Managed();
                    break;

                default:
                    hash = new MD5CryptoServiceProvider();
                    break;
            }

            // Compute hash value of our plain text with appended salt.
            byte[] hashBytes = hash.ComputeHash(plainTextWithSaltBytes);

            // Create array which will hold hash and original salt bytes.
            byte[] hashWithSaltBytes = new byte[hashBytes.Length +
                                                saltBytes.Length];

            // Copy hash bytes into resulting array.
            for (int i = 0; i < hashBytes.Length; i++)
                hashWithSaltBytes[i] = hashBytes[i];

            // Append salt bytes to the result.
            for (int i = 0; i < saltBytes.Length; i++)
                hashWithSaltBytes[hashBytes.Length + i] = saltBytes[i];

            // Convert result into a base64-encoded string.
            string hashValue = Convert.ToBase64String(hashWithSaltBytes);

            // Return the result.
            return hashValue;


        }


        /// <summary>
        /// <c>ComputeHashWithoutSalt</c> member function
        /// this function calls <c>ComputeHash</c> member function to calculate hash value of provided  data
        /// </summary>
        /// <param name="plainText"></param>
        /// <param name="hashAlgorithm"></param>
        /// <returns></returns>
        public static string ComputeHashWithoutSalt(string plainText, string hashAlgorithm)
        {
            string result = "";
            try
            {
                // Convert plain text into a byte array.
                byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
                HashAlgorithm hash;

                // Make sure hashing algorithm name is specified.
                if (hashAlgorithm == null)
                    hashAlgorithm = "";

                // Initialize appropriate hashing algorithm class.
                switch (hashAlgorithm.ToUpper())
                {
                    case "SHA1":
                        hash = new SHA1Managed();
                        break;

                    case "SHA256":
                        hash = new SHA256Managed();
                        break;

                    case "SHA384":
                        hash = new SHA384Managed();
                        break;

                    case "SHA512":
                        hash = new SHA512Managed();
                        break;

                    default:
                        hash = new MD5CryptoServiceProvider();
                        break;
                }

                // Compute hash value of our plain text with appended salt.
                byte[] hashBytes = hash.ComputeHash(plainTextBytes);
                result = Convert.ToBase64String(hashBytes);
            }
            catch (Exception)
            {


            }

            ListWebClass.Log("Calculated Hash 10:" + result, false);
            while (result.IndexOf('/') != -1)
            {
                try
                {
                    result = result.Remove(result.IndexOf('/'), 1);
                }
                catch (Exception)
                {


                }
            }
            ListWebClass.Log("Updated Hash 10:" + result, false);
            return result;
        }


        /// <summary>
        /// <c>VerifyHash</c> verify hash value
        /// this membe function is not currently used
        /// </summary>
        /// <param name="plainText"></param>
        /// <param name="hashAlgorithm"></param>
        /// <param name="hashValue"></param>
        /// <returns></returns>
        public static bool VerifyHash(string plainText, string hashAlgorithm, string hashValue)
        {
            // Convert base64-encoded hash value into a byte array.
            byte[] hashWithSaltBytes = Convert.FromBase64String(hashValue);

            // We must know size of hash (without salt).
            int hashSizeInBits, hashSizeInBytes;

            // Make sure that hashing algorithm name is specified.
            if (hashAlgorithm == null)
                hashAlgorithm = "";

            // Size of hash is based on the specified algorithm.
            switch (hashAlgorithm.ToUpper())
            {
                case "SHA1":
                    hashSizeInBits = 160;
                    break;

                case "SHA256":
                    hashSizeInBits = 256;
                    break;

                case "SHA384":
                    hashSizeInBits = 384;
                    break;

                case "SHA512":
                    hashSizeInBits = 512;
                    break;

                default: // Must be MD5
                    hashSizeInBits = 128;
                    break;
            }

            // Convert size of hash from bits to bytes.
            hashSizeInBytes = hashSizeInBits / 8;

            // Make sure that the specified hash value is long enough.
            if (hashWithSaltBytes.Length < hashSizeInBytes)
                return false;

            // Allocate array to hold original salt bytes retrieved from hash.
            byte[] saltBytes = new byte[hashWithSaltBytes.Length -
                                        hashSizeInBytes];

            // Copy salt from the end of the hash to the new array.
            for (int i = 0; i < saltBytes.Length; i++)
                saltBytes[i] = hashWithSaltBytes[hashSizeInBytes + i];

            // Compute a new hash string.
            string expectedHashString =
                        ComputeHash(plainText, hashAlgorithm, saltBytes);

            // If the computed hash matches the specified hash,
            // the plain text value must be correct.
            return (hashValue == expectedHashString);
        }

    }
}
