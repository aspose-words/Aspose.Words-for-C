#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <Model/Text/Paragraph.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>

#include "Common.h"

using namespace Aspose::Words;

void ExtractContentBetweenParagraphs()
{
    std::cout << "ExtractContentBetweenParagraphs example started." << std::endl;
    // ExStart:ExtractContentBetweenParagraphs
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");

    // Gather the nodes. The GetChild method uses 0-based index
    System::SharedPtr<Paragraph> startPara = System::DynamicCast<Paragraph>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::Paragraph, 6, true));
    System::SharedPtr<Paragraph> endPara = System::DynamicCast<Paragraph>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::Paragraph, 10, true));
    // Extract the content between these nodes in the document. Include these markers in the extraction.
    auto extractedNodes = ExtractContent(startPara, endPara, true);

    // Insert the content into a new separate document and save it to disk.
    System::SharedPtr<Document> dstDoc = GenerateDocument(doc, extractedNodes);
    System::String outputPath = dataDir + GetOutputFilePath(u"ExtractContentBetweenParagraphs.doc");
    dstDoc->Save(outputPath);
    // ExEnd:ExtractContentBetweenParagraphs
    std::cout << "Extracted content betweenn the paragraphs successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExtractContentBetweenParagraphs example finished." << std::endl << std::endl;
}
