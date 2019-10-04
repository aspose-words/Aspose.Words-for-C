#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionStart.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void KeepSourceTogether()
{
    std::cout << "KeepSourceTogether example started." << std::endl;
    // ExStart:KeepSourceTogether
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.DestinationList.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");

    // Set the source document to appear straight after the destination document's content.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);

    // Iterate through all sections in the source document.
    for (System::SharedPtr<Paragraph> para : System::IterateOver<System::SharedPtr<Paragraph>>(srcDoc->GetChildNodes(NodeType::Paragraph, true)))
    {
        para->get_ParagraphFormat()->set_KeepWithNext(true);
    }

    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = outputDataDir + u"KeepSourceTogether.doc";
    dstDoc->Save(outputPath);
    // ExEnd:KeepSourceTogether
    std::cout << "Document appended successfully while keeping the content from splitting across two pages." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "KeepSourceTogether example finished." << std::endl << std::endl;
}
