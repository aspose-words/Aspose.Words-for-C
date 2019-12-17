#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/HtmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void ConvertDocumentToHtmlWithRoundtrip()
{
    std::cout << "ConvertDocumentToHtmlWithRoundtrip example started." << std::endl;
    // ExStart:ConvertDocumentToHtmlWithRoundtrip
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (doc).doc");

    System::SharedPtr<HtmlSaveOptions> options = System::MakeObject<HtmlSaveOptions>();
    // HtmlSaveOptions.ExportRoundtripInformation property specifies
    // Whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
    // Default value is true for HTML and false for MHTML and EPUB.
    options->set_ExportRoundtripInformation(true);

    doc->Save(outputDataDir + u"ConvertDocumentToHtmlWithRoundtrip.html", options);
    // ExEnd:ConvertDocumentToHtmlWithRoundtrip
    std::cout << "Document converted to html with roundtrip informations successfully." << std::endl;
    std::cout << "ConvertDocumentToHtmlWithRoundtrip example finished." << std::endl << std::endl;
}