Imports DevExpress.Office.DigitalSignatures
Imports DevExpress.Office.Tsp
Imports System
Imports System.Diagnostics
Imports System.Net
Imports System.Security.Cryptography.X509Certificates

Namespace WordDocumentDigitalSignatureSample
	Friend Class Program
		Private Shared input As String = "Template.docx"
		Private Shared output As String = "Template_signed.docx"
		Shared Sub Main(ByVal args() As String)
			'Enable all security protocols:
			ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11

			SignDocument(input)
			ValidateSignature(output)
		End Sub
		Private Shared Sub SignDocument(ByVal path As String)
			'Sign a document and save the result:
			Dim documentSigner As New DocumentSigner()
			documentSigner.Sign(path, output, CreateSignatureOptions(), CreateSignatureInfo())
		End Sub
		'Specify signature options
		Private Shared Function CreateSignatureOptions() As SignatureOptions
			Dim certificate As New X509Certificate2("Certificate/SignDemo.pfx", "dxdemo")
			Dim tsaServer As New Uri("https://freetsa.org/tsr")
			Dim options As New SignatureOptions()
			options.Certificate = certificate
			If tsaServer IsNot Nothing Then
				options.TsaClient = New TsaClient(tsaServer, HashAlgorithmType.SHA256)
			End If

			'In this example, certificate validation is skipped
			options.SignatureFlags = options.SignatureFlags And Not SignatureFlags.ValidateCertificate
			options.CertificateKeyUsageFlags = X509KeyUsageFlags.None
			options.DigestMethod = HashAlgorithmType.SHA256

			Dim policy As New X509ChainPolicy()
			policy.RevocationMode = X509RevocationMode.NoCheck
			policy.RevocationFlag = X509RevocationFlag.ExcludeRoot
			policy.VerificationFlags = policy.VerificationFlags Or X509VerificationFlags.AllowUnknownCertificateAuthority Or X509VerificationFlags.IgnoreCertificateAuthorityRevocationUnknown
			options.CertificatePolicy = policy
			options.TimestampCertificatePolicy = policy
			Return options
		End Function

		'Specify signature information:
		Private Shared Function CreateSignatureInfo() As SignatureInfo
			Dim signatureInfo As New SignatureInfo()
			signatureInfo.CommitmentType = CommitmentType.ProofOfApproval
			signatureInfo.Time = Date.UtcNow
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

		Private Shared Sub ValidateSignature(ByVal path As String)
			Dim validator As New DocumentSigner()

			Dim validationOptions As New SignatureValidationOptions()

			'In this example, signature and timestamp certificate validation is skipped
			validationOptions.ValidationFlags = (Not ValidationFlags.ValidateSignatureCertificate) And Not ValidationFlags.ValidateTimestampCertificate

			'Validate the signature:
			Dim signatureValidation As PackageSignatureValidation = validator.Validate(path, validationOptions)
			Dim validationMessage As String = signatureValidation.ResultMessage

			'Check validation result and show information in the console:
			Select Case signatureValidation.Result
				Case PackageSignatureValidationResult.Valid
					Console.WriteLine(validationMessage)
					Console.ReadKey()
					Process.Start(path)

				Case PackageSignatureValidationResult.SignaturesNotFound
					Console.WriteLine(validationMessage)

				Case PackageSignatureValidationResult.Invalid, PackageSignatureValidationResult.PartiallyValid
					Dim failedCheckDetails = signatureValidation.Items(0).FailedCheckDetails
					Console.WriteLine(validationMessage)
					Dim i As Integer = 1
					For Each checkResult As SignatureCheckResult In failedCheckDetails
						Console.WriteLine(String.Format("Validation details {0}: " & ControlChars.CrLf & "{1} failed, Info: {2} " & ControlChars.CrLf, i, checkResult.CheckType, checkResult.Info))
						i += 1
					Next checkResult
					Console.ReadKey()
			End Select
		End Sub
	End Class
End Namespace
