#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/HtmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void ConvertUsingDefaultSaveOption(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.EpubConversion.doc");
        // Save the document in EPUB format.
        doc->Save(outputDataDir + u"ConvertDocumentToEPUB.ConvertUsingDefaultSaveOption.epub");
    }

    void ConvertWithSpecifiedSaveOptions(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ConvertDocumentToEPUB
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.EpubConversion.doc");

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
        doc->Save(outputDataDir + u"ConvertDocumentToEPUB.ConvertDocumentToEPUB.epub", saveOptions);
        // ExEnd:ConvertDocumentToEPUB

        std::cout << "Document converted to EPUB successfully." << std::endl;
    }
}

void ConvertDocumentToEPUB()
{
    std::cout << "ConvertDocumentToEPUB example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    ConvertUsingDefaultSaveOption(inputDataDir, outputDataDir);
    ConvertWithSpecifiedSaveOptions(inputDataDir, outputDataDir);

    std::cout << "ConvertDocumentToEPUB example finished." << std::endl << std::endl;
}