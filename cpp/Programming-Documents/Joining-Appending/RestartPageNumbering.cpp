#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Sections/SectionStart.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words; 

void RestartPageNumbering()
{
    std::cout << "RestartPageNumbering example started." << std::endl;
    // ExStart:RestartPageNumbering
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");

    // Set the appended document to appear on the next page.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
    // Restart the page numbering for the document to be appended.
    srcDoc->get_FirstSection()->get_PageSetup()->set_RestartPageNumbering(true);

    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = outputDataDir + u"RestartPageNumbering.doc";
    dstDoc->Save(outputPath);
    // ExEnd:RestartPageNumbering
    std::cout << "Document appended successfully with restart page numbering." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RestartPageNumbering example finished." << std::endl << std::endl;
}