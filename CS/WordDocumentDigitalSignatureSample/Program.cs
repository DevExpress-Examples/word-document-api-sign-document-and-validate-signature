using DevExpress.Office.DigitalSignatures;
using DevExpress.Office.Tsp;
using System;
using System.Diagnostics;
using System.Net;
using System.Security.Cryptography.X509Certificates;

namespace WordDocumentDigitalSignatureSample
{
    class Program
    {
        static string input = "Template.docx";
        static string output = "Template_signed.docx";
        static void Main(string[] args)
        {
            //Enable all security protocols:
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 
                | SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11;
 
            SignDocument(input);
            ValidateSignature(output);
        }
        static void SignDocument(string path)
        {            
            //Sign a document and save the result:
            DocumentSigner documentSigner = new DocumentSigner();
            documentSigner.Sign(path, output, CreateSignatureOptions(), CreateSignatureInfo());
        }
        //Specify signature options
        static SignatureOptions CreateSignatureOptions()
        {
            X509Certificate2 certificate = new X509Certificate2("Certificate/SignDemo.pfx", "dxdemo");
            Uri tsaServer = new Uri("https://freetsa.org/tsr");
            SignatureOptions options = new SignatureOptions();
            options.Certificate = certificate;
            if (tsaServer != null)
                options.TsaClient = new TsaClient(tsaServer, HashAlgorithmType.SHA256);

            //In this example, certificate validation is skipped
            options.SignatureFlags &= ~SignatureFlags.ValidateCertificate;
            options.CertificateKeyUsageFlags = X509KeyUsageFlags.None;
            options.DigestMethod = HashAlgorithmType.SHA256;

            X509ChainPolicy policy = new X509ChainPolicy();
            policy.RevocationMode = X509RevocationMode.NoCheck;
            policy.RevocationFlag = X509RevocationFlag.ExcludeRoot;
            policy.VerificationFlags |= X509VerificationFlags.AllowUnknownCertificateAuthority |
                X509VerificationFlags.IgnoreCertificateAuthorityRevocationUnknown;
            options.CertificatePolicy = policy;
            options.TimestampCertificatePolicy = policy;
            return options;
        }
        
        //Specify signature information:
        static SignatureInfo CreateSignatureInfo()
        {
            SignatureInfo signatureInfo = new SignatureInfo();
            signatureInfo.CommitmentType = CommitmentType.ProofOfApproval;
            signatureInfo.Time = DateTime.UtcNow;
            signatureInfo.ClaimedRoles.Clear();
            signatureInfo.ClaimedRoles.Add("Sales Representative");
            signatureInfo.Country = "USA";
            signatureInfo.City = "Seattle";
            signatureInfo.StateOrProvince = "WA";
            signatureInfo.Address1 = "507 - 20th Ave. E.";
            signatureInfo.Address2 = "Apt. 2A";
            signatureInfo.PostalCode = "98122";
            signatureInfo.Comments = "Demo Digital Signature";
            return signatureInfo;
        }

        private static void ValidateSignature(string path)
        {
            DocumentSigner validator = new DocumentSigner();

            SignatureValidationOptions validationOptions = new SignatureValidationOptions();
            
            //In this example, signature and timestamp certificate validation is skipped
            validationOptions.ValidationFlags = ~ValidationFlags.ValidateSignatureCertificate & ~ValidationFlags.ValidateTimestampCertificate;

            //Validate the signature:
            PackageSignatureValidation signatureValidation = validator.Validate(path, validationOptions);
            string validationMessage = signatureValidation.ResultMessage;
            
            //Check validation result and show information in the console:
            switch (signatureValidation.Result)
            {
                case PackageSignatureValidationResult.Valid:
                    Console.WriteLine(validationMessage); Console.ReadKey();
                    Process.Start(path);
                    break;

                case PackageSignatureValidationResult.SignaturesNotFound:
                    Console.WriteLine(validationMessage);
                    break;

                case PackageSignatureValidationResult.Invalid:
                case PackageSignatureValidationResult.PartiallyValid:
                    var failedCheckDetails = signatureValidation.Items[0].FailedCheckDetails;
                    Console.WriteLine(validationMessage);
                    int i = 1;
                    foreach (SignatureCheckResult checkResult in failedCheckDetails)
                    {
                        Console.WriteLine(String.Format("Validation details {0}: \r\n" +
                            "{1} failed, Info: {2} \r\n", i, checkResult.CheckType, checkResult.Info));
                        i++;
                    }
                    Console.ReadKey();
                    break;
            }
        }
    }
}
