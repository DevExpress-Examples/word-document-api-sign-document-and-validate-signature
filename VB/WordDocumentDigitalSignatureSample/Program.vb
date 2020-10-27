Imports DevExpress.Office.DigitalSignatures
Imports DevExpress.Office.Tsp
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Threading.Tasks

Namespace WordDocumentDigitalSignatureSample
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			SignDocument("Template.docx")
		End Sub
		Private Shared Sub SignDocument(ByVal path As String)
			Dim output As String = "Template_signed.docx"
			Dim documentSigner As New DocumentSigner()
			documentSigner.Sign(path, output, CreateSignatureOptions(), CreateSignatureInfo())
			Process.Start(output)
		End Sub

		Private Shared Function CreateSignatureOptions() As SignatureOptions
			Dim certificate As New X509Certificate2("Certificate/SignDemo.pfx", "dxdemo")
			Dim tsaServer As New Uri("https://freetsa.org/tsr")
			Dim hashAlgorithm As HashAlgorithmType = HashAlgorithmType.SHA256
			Dim options As New SignatureOptions()
			options.Certificate = certificate

			If tsaServer IsNot Nothing Then
				options.TsaClient = New TsaClient(tsaServer, HashAlgorithmType.SHA256)
			End If

			Dim policy As New X509ChainPolicy()
			policy.RevocationMode = X509RevocationMode.NoCheck
			policy.RevocationFlag = X509RevocationFlag.ExcludeRoot
			policy.VerificationFlags = policy.VerificationFlags Or X509VerificationFlags.AllowUnknownCertificateAuthority Or X509VerificationFlags.IgnoreCertificateAuthorityRevocationUnknown
			options.CertificatePolicy = policy
			options.TimestampCertificatePolicy = policy
			options.SignatureFlags = options.SignatureFlags And Not SignatureFlags.ValidateCertificate
			options.CertificateKeyUsageFlags = X509KeyUsageFlags.None
			options.DigestMethod = hashAlgorithm
			Return options
		End Function
		Private Shared Function CreateSignatureInfo() As SignatureInfo
			Dim signatureInfo As New SignatureInfo()
			signatureInfo.CommitmentType = CommitmentType.ProofOfApproval
			signatureInfo.Time = DateTime.UtcNow
			signatureInfo.ClaimedRoles.Clear()
			signatureInfo.ClaimedRoles.Add("Sales Representative")
			signatureInfo.Country = "USA"
			signatureInfo.City = "Seattle"
			signatureInfo.StateOrProvince = "WA"
			signatureInfo.Address1 = "507 - 20th Ave. E."
			signatureInfo.Address2 = "Apt. 2A"
			signatureInfo.PostalCode = "98122"
			signatureInfo.Comments = "Demo Digital Signature"
			Return signatureInfo
		End Function

	End Class
End Namespace
