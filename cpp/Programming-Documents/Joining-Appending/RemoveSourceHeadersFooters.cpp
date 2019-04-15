#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/HeaderFooterCollection.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void RemoveSourceHeadersFooters()
{
    std::cout << "RemoveSourceHeadersFooters example started." << std::endl;
    // ExStart:RemoveSourceHeadersFooters
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Remove the headers and footers from each of the sections in the source document.
    for (System::SharedPtr<Section> section : System::IterateOver<System::SharedPtr<Section>>(srcDoc->get_Sections()))
    {
        section->ClearHeadersFooters();
    }

    // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
    // For HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
    // Document. This should set to false to avoid this behavior.
    srcDoc->get_FirstSection()->get_HeadersFooters()->LinkToPrevious(false);

    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = dataDir + GetOutputFilePath(u"RemoveSourceHeadersFooters.doc");
    dstDoc->Save(outputPath);
    // ExEnd:RemoveSourceHeadersFooters
    std::cout << "Document appended successfully with source header footers removed." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveSourceHeadersFooters example finished." << std::endl << std::endl;
}
