#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/CommentCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void CommentReply()
{
    std::cout << "CommentReply example started." << std::endl;
    // ExStart:AddRemoveCommentReply
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithComments();
    System::String outputDataDir = GetOutputDataDir_WorkingWithComments();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    System::SharedPtr<Comment> comment = System::DynamicCast<Comment>(doc->GetChild(NodeType::Comment, 0, true));

    //Remove the reply
    comment->RemoveReply(comment->get_Replies()->idx_get(0));

    //Add a reply to comment
    comment->AddReply(u"John Doe", u"JD", System::DateTime(2017, 9, 25, 12, 15, 0), u"New reply");

    System::String outputPath = outputDataDir + u"CommentReply.doc";

    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:AddRemoveCommentReply
    std::cout << "Comment's reply is removed successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CommentReply example started." << std::endl << std::endl;
}
