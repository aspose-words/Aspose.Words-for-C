#include "stdafx.h"
#include "examples.h"

#include <Model/Sections/Section.h>
#include <Model/Sections/HeaderFooterCollection.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void UnlinkHeadersFooters()
{
    std::cout << "UnlinkHeadersFooters example started." << std::endl;
    // ExStart:UnlinkHeadersFooters
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
    // From the destination document.
    srcDoc->get_FirstSection()->get_HeadersFooters()->LinkToPrevious(false);

    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = dataDir + GetOutputFilePath(u"UnlinkHeadersFooters.doc");
    dstDoc->Save(outputPath);
    // ExEnd:UnlinkHeadersFooters
    std::cout << "Document appended successfully with unlinked header footers." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "UnlinkHeadersFooters example finished." << std::endl << std::endl;
}
