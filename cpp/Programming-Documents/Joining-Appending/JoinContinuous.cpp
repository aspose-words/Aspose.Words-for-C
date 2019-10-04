#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Sections/SectionStart.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void JoinContinuous()
{
    std::cout << "JoinContinuous example started." << std::endl;
    // ExStart:JoinContinuous
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");

    // Make the document appear straight after the destination documents content.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);

    // Append the source document using the original styles found in the source document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = outputDataDir + u"JoinContinuous.doc";
    dstDoc->Save(outputPath);
    // ExEnd:JoinContinuous
    std::cout << "Document appended successfully with continous section start." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "JoinContinuous example finished." << std::endl << std::endl;
}
