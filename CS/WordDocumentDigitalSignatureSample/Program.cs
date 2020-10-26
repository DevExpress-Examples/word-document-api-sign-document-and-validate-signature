using DevExpress.Office.DigitalSignatures;
using DevExpress.Office.Tsp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentDigitalSignatureSample
{
    class Program
    {
        static void Main(string[] args)
        {
            SignDocument("Template.docx");
        }
        static void SignDocument(string path)
        {
            string output = "Template_signed.docx";
            DocumentSigner documentSigner = new DocumentSigner();
            documentSigner.Sign(path, output, CreateSignatureOptions(), CreateSignatureInfo());
            Process.Start(output);
        }
    
        static SignatureOptions CreateSignatureOptions()
        {
            X509Certificate2 certificate = new X509Certificate2("Certificate/SignDemo.pfx", "dxdemo");
            Uri tsaServer = new Uri("https://freetsa.org/tsr");
            HashAlgorithmType hashAlgorithm = HashAlgorithmType.SHA256;
            SignatureOptions options = new SignatureOptions();
            options.Certificate = certificate;

            if (tsaServer != null)
                options.TsaClient = new TsaClient(tsaServer, HashAlgorithmType.SHA256);

            X509ChainPolicy policy = new X509ChainPolicy();
            policy.RevocationMode = X509RevocationMode.NoCheck;
            policy.RevocationFlag = X509RevocationFlag.ExcludeRoot;
            policy.VerificationFlags |= X509VerificationFlags.AllowUnknownCertificateAuthority | 
                X509VerificationFlags.IgnoreCertificateAuthorityRevocationUnknown;
            options.CertificatePolicy = policy;
            options.TimestampCertificatePolicy = policy;
            options.SignatureFlags &= ~SignatureFlags.ValidateCertificate;
            options.CertificateKeyUsageFlags = X509KeyUsageFlags.None;
            options.DigestMethod = hashAlgorithm;
            return options;
        }
        static SignatureInfo CreateSignatureInfo()
        {
            SignatureInfo signatureInfo = new SignatureInfo();
            string role = "Sales Representative";
            string comments = "Demo Digital Signature";
            string country = "USA";
            string state = "WA";
            string city = "Seattle";
            string address1 = "507 - 20th Ave. E.";
            string address2 = "Apt. 2A";
            string postalCode = "98122";

            signatureInfo.CommitmentType = CommitmentType.ProofOfApproval;
            signatureInfo.Time = DateTime.UtcNow;
            signatureInfo.ClaimedRoles.Clear();
            signatureInfo.ClaimedRoles.Add(role);
            signatureInfo.Country = country;
            signatureInfo.City = city;
            signatureInfo.StateOrProvince = state;
            signatureInfo.Address1 = address1;
            signatureInfo.Address2 = address2;
            signatureInfo.PostalCode = postalCode;
            signatureInfo.Comments = comments;
            return signatureInfo;
        }
    
    }
}
