#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/PdfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void DisplayDocTitleInWindowTitlebar(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:DisplayDocTitleInWindowTitlebar
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Template.doc");

        System::SharedPtr<PdfSaveOptions> saveOptions = System::MakeObject<PdfSaveOptions>();
        saveOptions->set_DisplayDocTitle(true);

        System::String outputPath = outputDataDir + u"Doc2Pdf.DisplayDocTitleInWindowTitlebar.pdf";

        // Save the document in PDF format.
        doc->Save(outputPath, saveOptions);
        // ExEnd:DisplayDocTitleInWindowTitlebar
        std::cout << "Document converted to PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void Doc2PdfImpl(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:Doc2Pdf
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Template.doc");

        System::String outputPath = outputDataDir + u"Doc2Pdf.Doc2PdfImpl.pdf";
        // Save the document in PDF format.
        doc->Save(outputPath);
        // ExEnd:Doc2Pdf
        std::cout << "Document converted to PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void Doc2Pdf()
{
    std::cout << "Doc2Pdf example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_QuickStart();
    System::String outputDataDir = GetOutputDataDir_QuickStart();

    DisplayDocTitleInWindowTitlebar(inputDataDir, outputDataDir);
    Doc2PdfImpl(inputDataDir, outputDataDir);
    std::cout << "Doc2Pdf example finished." << std::endl << std::endl;
}