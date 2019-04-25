#include "stdafx.h"
#include "examples.h"

#include <Model/Sections/SectionStart.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void ListKeepSourceFormatting()
{
    std::cout << "ListKeepSourceFormatting example started." << std::endl;
    // ExStart:ListKeepSourceFormatting
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.DestinationList.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.SourceList.doc");

    // Append the content of the document so it flows continuously.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);

    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = dataDir + GetOutputFilePath(u"ListKeepSourceFormatting.doc");
    dstDoc->Save(outputPath);
    // ExEnd:ListKeepSourceFormatting
    std::cout << "Document appended successfully with lists retaining source formatting." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ListKeepSourceFormatting example finished." << std::endl << std::endl;
}
