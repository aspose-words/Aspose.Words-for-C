#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Importing/ImportFormatMode.h>

using namespace Aspose::Words;

void UpdatePageLayout()
{
    std::cout << "UpdatePageLayout example started." << std::endl;
    // ExStart:UpdatePageLayout
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document 
    // Is appended then any changes made after will not be reflected in the rendered output.
    dstDoc->UpdatePageLayout();

    // Join the documents.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

    // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
    // If not called again the appended document will not appear in the output of the next rendering.
    dstDoc->UpdatePageLayout();

    System::String outputPath = dataDir = dataDir + GetOutputFilePath(u"UpdatePageLayout.pdf");
    // Save the joined document to PDF.
    dstDoc->Save(outputPath);
    // ExEnd:UpdatePageLayout
    std::cout << "Document appended successfully with updated page layout after appending the document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "UpdatePageLayout example finished." << std::endl << std::endl;
}