#include "stdafx.h"
#include "examples.h"

#include <Model/Sections/SectionStart.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void JoinNewPage()
{
    std::cout << "JoinNewPage example started." << std::endl;
    // ExStart:JoinNewPage
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");

    // Set the appended document to start on a new page.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::NewPage);

    // Append the source document using the original styles found in the source document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = outputDataDir + u"JoinNewPage.doc";
    dstDoc->Save(outputPath);
    // ExEnd:JoinNewPage
    std::cout << "Document appended successfully with new section start." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "JoinNewPage example finished." << std::endl << std::endl;
}
