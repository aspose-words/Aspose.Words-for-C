#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/date_time.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfDigitalSignatureDetails.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void DigitallySignedPdfUsingCertificateHolder()
{
    // ExStart:DigitallySignedPdfUsingCertificateHolder
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    // Create a simple document from scratch.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Test Signed PDF.");

    System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
    options->set_DigitalSignatureDetails(System::MakeObject<PdfDigitalSignatureDetails>(CertificateHolder::Create(inputDataDir + u"CioSrv1.pfx", u"cinD96..arellA"), u"reason", u"location", System::DateTime::get_Now()));

    System::String outputPath = outputDataDir + u"DigitallySignedPdfUsingCertificateHolder.pdf";
    doc->Save(outputPath, options);
    // ExEnd:DigitallySignedPdfUsingCertificateHolder
    std::cout << "Digitally signed PDF file created successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
}
