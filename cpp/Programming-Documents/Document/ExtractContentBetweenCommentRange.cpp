#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <Model/Text/CommentRangeStart.h>
#include <Model/Text/CommentRangeEnd.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>

#include "Common.h"

using namespace Aspose::Words;


void ExtractContentBetweenCommentRange()
{
    // ExStart:ExtractContentBetweenCommentRange
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    
    // This is a quick way of getting both comment nodes.
    // Your code should have a proper method of retrieving each corresponding start and end node.
    System::SharedPtr<CommentRangeStart> commentStart = System::DynamicCast<Aspose::Words::CommentRangeStart>(doc->GetChild(Aspose::Words::NodeType::CommentRangeStart, 0, true));
    System::SharedPtr<CommentRangeEnd> commentEnd = System::DynamicCast<Aspose::Words::CommentRangeEnd>(doc->GetChild(Aspose::Words::NodeType::CommentRangeEnd, 0, true));
    
    // Firstly extract the content between these nodes including the comment as well. 
    auto extractedNodesInclusive = ExtractContent(commentStart, commentEnd, true);
    System::SharedPtr<Document> dstDoc = GenerateDocument(doc, extractedNodesInclusive);
    dstDoc->Save(dataDir + GetOutputFilePath(u"ExtractContentBetweenCommentRange.Inclusive.doc"));
    
    // Secondly extract the content between these nodes without the comment.
    auto extractedNodesExclusive = ExtractContent(commentStart, commentEnd, false);
    dstDoc = GenerateDocument(doc, extractedNodesExclusive);
    dataDir += GetOutputFilePath(u"ExtractContentBetweenCommentRange.Exclusive.doc");
    dstDoc->Save(dataDir);
    // ExEnd:ExtractContentBetweenCommentRange
    std::cout << "\nExtracted content between the comment range successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
