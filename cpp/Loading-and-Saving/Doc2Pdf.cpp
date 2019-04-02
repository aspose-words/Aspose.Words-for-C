#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/PdfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void DisplayDocTitleInWindowTitlebar(System::String const &dataDir)
    {
        // ExStart:DisplayDocTitleInWindowTitlebar
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Template.doc");

        System::SharedPtr<PdfSaveOptions> saveOptions = System::MakeObject<PdfSaveOptions>();
        saveOptions->set_DisplayDocTitle(true);

        System::String outputPath = dataDir + GetOutputFilePath(u"Doc2Pdf.DisplayDocTitleInWindowTitlebar.pdf");

        // Save the document in PDF format.
        doc->Save(outputPath, saveOptions);
        // ExEnd:DisplayDocTitleInWindowTitlebar
        std::cout << "Document converted to PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

}

void Doc2Pdf()
{
    std::cout << "Doc2Pdf example started." << std::endl;
    // ExStart:Doc2Pdf
    // The path to the documents directory.
    System::String dataDir = GetDataDir_QuickStart();

    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Template.doc");

    System::String outputPath = dataDir + GetOutputFilePath(u"Doc2Pdf.pdf");
    // Save the document in PDF format.
    doc->Save(outputPath);
    // ExEnd:Doc2Pdf
    std::cout << "Document converted to PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    DisplayDocTitleInWindowTitlebar(dataDir);
    std::cout << "Doc2Pdf example finished." << std::endl << std::endl;
}