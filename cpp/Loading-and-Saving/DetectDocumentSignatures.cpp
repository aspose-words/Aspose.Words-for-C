#include "stdafx.h"
#include "examples.h"

#include <system/io/path.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>

using namespace Aspose::Words;

void DetectDocumentSignatures()
{
    std::cout << "DetectDocumentSignatures example started." << std::endl;
    // ExStart:DetectDocumentSignatures
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();

    // The path to the document which is to be processed.
    System::String filePath = inputDataDir + u"Document.Signed.docx";

    System::SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(filePath);
    if (info->get_HasDigitalSignature())
    {
        std::cout << "Document " << System::IO::Path::GetFileName(filePath).ToUtf8String() << " has digital signatures, they will be lost if you open/save this document with Aspose.Words." << std::endl;
    }
    // ExEnd:DetectDocumentSignatures
    std::cout << "DetectDocumentSignatures example finished." << std::endl << std::endl;
}