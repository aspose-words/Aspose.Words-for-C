#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/date_time.h>
#include <Model/Text/CommentCollection.h>
#include <Model/Text/Comment.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void CommentReply()
{
    std::cout << "CommentReply example started." << std::endl;
    // ExStart:AddRemoveCommentReply
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithComments();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    System::SharedPtr<Comment> comment = System::DynamicCast<Comment>(doc->GetChild(NodeType::Comment, 0, true));

    //Remove the reply
    comment->RemoveReply(comment->get_Replies()->idx_get(0));

    //Add a reply to comment
    comment->AddReply(u"John Doe", u"JD", System::DateTime(2017, 9, 25, 12, 15, 0), u"New reply");

    System::String outputPath = dataDir + GetOutputFilePath(u"CommentReply.doc");

    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:AddRemoveCommentReply
    std::cout << "Comment's reply is removed successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CommentReply example started." << std::endl << std::endl;
}
