#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <Model/Text/Paragraph.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>

#include "Common.h"

using namespace Aspose::Words;

void ExtractContentBetweenParagraphStyles()
{
    std::cout << "ExtractContentBetweenParagraphStyles example started." << std::endl;
    // ExStart:ExtractContentBetweenParagraphStyles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");

    // Gather a list of the paragraphs using the respective heading styles.
    auto parasStyleHeading1 = ParagraphsByStyleName(doc, u"Heading 1");
    auto parasStyleHeading3 = ParagraphsByStyleName(doc, u"Heading 3");

    // Use the first instance of the paragraphs with those styles.
    System::SharedPtr<Node> startPara1 = parasStyleHeading1[0];
    System::SharedPtr<Node> endPara1 = parasStyleHeading3[0];

    // Extract the content between these nodes in the document. Don't include these markers in the extraction.
    auto extractedNodes = ExtractContent(startPara1, endPara1, false);

    // Insert the content into a new separate document and save it to disk.
    System::SharedPtr<Document> dstDoc = GenerateDocument(doc, extractedNodes);
    System::String outputPath = dataDir + GetOutputFilePath(u"ExtractContentBetweenParagraphStyles.doc");
    dstDoc->Save(outputPath);
    // ExEnd:ExtractContentBetweenParagraphStyles
    std::cout << "Extracted content betweenn the paragraph styles successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExtractContentBetweenParagraphStyles example finished." << std::endl << std::endl;
}
