#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/Paragraph.h>
#include <Model/Sections/SectionStart.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void KeepSourceTogether()
{
    std::cout << "KeepSourceTogether example started." << std::endl;
    // ExStart:KeepSourceTogether
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.DestinationList.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Set the source document to appear straight after the destination document's content.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::Continuous);

    // Iterate through all sections in the source document.
    auto para_enumerator = srcDoc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->GetEnumerator();
    System::SharedPtr<Paragraph> para;
    while (para_enumerator->MoveNext() && (para = System::DynamicCast<Paragraph>(para_enumerator->get_Current()), true))
    {
        para->get_ParagraphFormat()->set_KeepWithNext(true);
    }

    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = dataDir + GetOutputFilePath(u"KeepSourceTogether.doc");
    dstDoc->Save(outputPath);
    // ExEnd:KeepSourceTogether
    std::cout << "Document appended successfully while keeping the content from splitting across two pages." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "KeepSourceTogether example finished." << std::endl << std::endl;
}
