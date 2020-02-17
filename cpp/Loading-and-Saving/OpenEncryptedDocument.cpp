#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>

using namespace Aspose::Words;

void OpenEncryptedDocument()
{
    std::cout << "OpenEncryptedDocument example started." << std::endl;
    // ExStart:OpenEncryptedDocument
    // The path to the documents directory.
    System::String dataDir = GetInputDataDir_LoadingAndSaving();

    // Loads encrypted document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"LoadEncrypted.docx", System::MakeObject<LoadOptions>(u"aspose"));

    // ExEnd:OpenEncryptedDocument
    std::cout << "Encrypted document loaded successfully." << std::endl;
    std::cout << "OpenEncryptedDocument example finished." << std::endl << std::endl;
}