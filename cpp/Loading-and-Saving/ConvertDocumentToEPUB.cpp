#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/HtmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void ConvertUsingDefaultSaveOption(System::String const &dataDir)
    {
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.EpubConversion.doc");
        // Save the document in EPUB format.
        doc->Save(dataDir + GetOutputFilePath(u"ConvertDocumentToEPUB.ConvertUsingDefaultSaveOption.epub"));
    }
}

void ConvertDocumentToEPUB()
{
    std::cout << "ConvertDocumentToEPUB example started." << std::endl;
    // ExStart:ConvertDocumentToEPUB
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();

    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.EpubConversion.doc");

    // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
    // How the output document is saved.
    System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();

    // Specify the desired encoding.
    saveOptions->set_Encoding(System::Text::Encoding::get_UTF8());

    // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
    // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
    // HTML files greater than a certain size e.g 300kb.
    saveOptions->set_DocumentSplitCriteria(DocumentSplitCriteria::HeadingParagraph);

    // Specify that we want to export document properties.
    saveOptions->set_ExportDocumentProperties(true);

    // Specify that we want to save in EPUB format.
    saveOptions->set_SaveFormat(SaveFormat::Epub);

    // Export the document as an EPUB file.
    doc->Save(dataDir + GetOutputFilePath(u"ConvertDocumentToEPUB.epub"), saveOptions);
    // ExEnd:ConvertDocumentToEPUB

    ConvertUsingDefaultSaveOption(dataDir);
    std::cout << "Document converted to EPUB successfully." << std::endl;
    std::cout << "ConvertDocumentToEPUB example finished." << std::endl << std::endl;
}