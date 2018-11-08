#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <system/collections/ienumerator.h>
#include <Model/Text/CommentCollection.h>
#include <Model/Text/Comment.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;

namespace
{

    // ExStart:ExtractComments
    std::vector<System::String> ExtractComments(const System::SharedPtr<Document>& doc)
    {
        std::vector<System::String> collectedComments;
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);

        // Look through all comments and gather information about them.
        auto comment_enumerator = comments->GetEnumerator();
        System::SharedPtr<Comment> comment;
        while (comment_enumerator->MoveNext() && (comment = System::DynamicCast<Comment>(comment_enumerator->get_Current()), true))
        {
            collectedComments.push_back(comment->get_Author() + u" " + comment->get_DateTime() + u" " + comment->GetText());
        }

        return collectedComments;
    }
    // ExEnd:ExtractComments

    // ExStart:ExtractCommentsByAuthor
    std::vector<System::String> ExtractComments(const System::SharedPtr<Document>& doc, const System::String& authorName)
    {
        std::vector<System::String> collectedComments;
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);

        // Look through all comments and gather information about those written by the authorName author.
        auto comment_enumerator = comments->GetEnumerator();
        System::SharedPtr<Comment> comment;
        while (comment_enumerator->MoveNext() && (comment = System::DynamicCast<Comment>(comment_enumerator->get_Current()), true))
        {
            if (comment->get_Author() == authorName)
            {
                collectedComments.push_back(comment->get_Author() + u" " + comment->get_DateTime() + u" " + comment->GetText());
            }
        }

        return collectedComments;
    }
    // ExEnd:ExtractCommentsByAuthor

    // ExStart:RemoveComments
    void RemoveComments(const System::SharedPtr<Document>& doc)
    {
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);
        // Remove all comments.
        comments->Clear();
    }
    // ExEnd:RemoveComments

    // ExStart:RemoveCommentsByAuthor
    void RemoveComments(const System::SharedPtr<Document>& doc, const System::String& authorName)
    {
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);
        // Look through all comments and remove those written by the authorName author.
        for (int32_t i = comments->get_Count() - 1; i >= 0; i--)
        {
            System::SharedPtr<Comment> comment = System::DynamicCast<Comment>(comments->idx_get(i));
            if (comment->get_Author() == authorName)
            {
                comment->Remove();
            }
        }
    }
    // ExEnd:RemoveCommentsByAuthor

    // ExStart:CommentResolvedandReplies
    void CommentResolvedandReplies(const System::SharedPtr<Document>& doc)
    {
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);
        System::SharedPtr<Comment> parentComment = System::DynamicCast<Comment>(comments->idx_get(0));

        auto childComment_enumerator = parentComment->get_Replies()->GetEnumerator();
        System::SharedPtr<Comment> childComment;
        while (childComment_enumerator->MoveNext() && (childComment = System::DynamicCast<Comment>(childComment_enumerator->get_Current()), true))
        {
            // Get comment parent and status.
            std::cout << childComment->get_Ancestor()->get_Id() << std::endl << childComment->get_Done() << std::endl;
            // And update comment Done mark.
            childComment->set_Done(true);
        }
    }
    // ExEnd:CommentResolvedandReplies
}

void ProcessComments()
{
    std::cout << "ProcessComments example started." << std::endl;
    // ExStart:ProcessComments
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithComments();
    
    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    
    // Extract the information about the comments of all the authors.

    for (auto& comment : ExtractComments(doc))
    {
        std::cout << comment.ToUtf8String();
    }

    // Remove comments by the "pm" author.
    RemoveComments(doc, u"pm");
    std::cout << "Comments from \"pm\" are removed!" << std::endl;
    
    // Extract the information about the comments of the "ks" author.
    
    for (auto& comment: ExtractComments(doc, u"ks"))
    {
        std::cout << comment.ToUtf8String();
    }

    //Read the comment's reply and resolve them.
    CommentResolvedandReplies(doc);
    
    // Remove all comments.
    RemoveComments(doc);
    std::cout << "All comments are removed!" << std::endl;
    
    System::String outputPath = dataDir + GetOutputFilePath(u"ProcessComments.doc");
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:ProcessComments
    std::cout << "Comments extracted and removed successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ProcessComments example finished." << std::endl << std::endl;
}
