using System.Security.Cryptography.X509Certificates;
using System.Net;

namespace Interfaces
{
    public class TrustAllCertificatePolicy : System.Net.ICertificatePolicy
    {
        public TrustAllCertificatePolicy() { }
        public bool CheckValidationResult(ServicePoint sp,
            X509Certificate cert,
            WebRequest req,
            int problem)
        {
            return true;
        }

    }
}
