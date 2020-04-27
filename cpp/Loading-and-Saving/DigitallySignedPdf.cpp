#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <security/cryptography/x509_certificates/x509_certificate_2.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfDigitalSignatureDetails.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace System::Security::Cryptography::X509Certificates;

void DigitallySignedPdf()
{
    // ExStart:DigitallySignedPdf
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetInputDataDir_LoadingAndSaving();

	if (!System::IO::File::Exists(inputDataDir + u"signature.pfx"))
    {
        std::cout << "Certificate file does not exist." << std::endl;
        return;
    }
    
    // Create a simple document from scratch.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Test Signed PDF.");
    
    // Load the certificate from disk.
    // The other constructor overloads can be used to load certificates from different locations.
    System::SharedPtr<X509Certificate2> cert = System::MakeObject<X509Certificate2>(inputDataDir + u"signature.pfx", u"signature");
    
    
    // Pass the certificate and details to the save options class to sign with.
    System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
    options->set_DigitalSignatureDetails(System::MakeObject<PdfDigitalSignatureDetails>());
    
    
    System::String outputPath = outputDataDir + u"Document.Signed_out.pdf";
    // Save the document as PDF.
    doc->Save(outputPath, options);
    // ExEnd:DigitallySignedPdf
    std::cout << "Digitally signed PDF file created successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
}