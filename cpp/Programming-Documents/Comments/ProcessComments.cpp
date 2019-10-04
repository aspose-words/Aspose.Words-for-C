#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Text/CommentCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

using namespace Aspose::Words;

namespace
{
    // ExStart:ExtractComments
    std::vector<System::String> ExtractComments(const System::SharedPtr<Document>& doc)
    {
        std::vector<System::String> collectedComments;
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);

        // Look through all comments and gather information about them.
        for (System::SharedPtr<Comment> comment : System::IterateOver<System::SharedPtr<Comment>>(comments))
        {
            collectedComments.push_back(comment->get_Author() + u" " + comment->get_DateTime() + u" " + System::StaticCast<Node>(comment)->ToString(SaveFormat::Text));
        }

        return collectedComments;
    }
    // ExEnd:ExtractComments

    // ExStart:ExtractCommentsByAuthor
    std::vector<System::String> ExtractComments(const System::SharedPtr<Document>& doc, const System::String& authorName)
    {
        std::vector<System::String> collectedComments;
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);

        // Look through all comments and gather information about those written by the authorName author.
        for (System::SharedPtr<Comment> comment : System::IterateOver<System::SharedPtr<Comment>>(comments))
        {
            if (comment->get_Author() == authorName)
            {
                collectedComments.push_back(comment->get_Author() + u" " + comment->get_DateTime() + u" " + System::StaticCast<Node>(comment)->ToString(SaveFormat::Text));
            }
        }

        return collectedComments;
    }
    // ExEnd:ExtractCommentsByAuthor

    // ExStart:RemoveComments
    void RemoveComments(const System::SharedPtr<Document>& doc)
    {
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);
        // Remove all comments.
        comments->Clear();
    }
    // ExEnd:RemoveComments

    // ExStart:RemoveCommentsByAuthor
    void RemoveComments(const System::SharedPtr<Document>& doc, const System::String& authorName)
    {
        // Collect all comments in the document
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);
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
        System::SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);
        System::SharedPtr<Comment> parentComment = System::DynamicCast<Comment>(comments->idx_get(0));

        for (System::SharedPtr<Comment> childComment : System::IterateOver<System::SharedPtr<Comment>>(parentComment->get_Replies()))
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
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithComments();
    System::String outputDataDir = GetOutputDataDir_WorkingWithComments();
    
    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    
    // Extract the information about the comments of all the authors.
    for (System::String const &comment : ExtractComments(doc))
    {
        std::cout << comment.ToUtf8String();
    }

    // Remove comments by the "pm" author.
    RemoveComments(doc, u"pm");
    std::cout << "Comments from \"pm\" are removed!" << std::endl;
    
    // Extract the information about the comments of the "ks" author.
    for (System::String const &comment: ExtractComments(doc, u"ks"))
    {
        std::cout << comment.ToUtf8String();
    }

    //Read the comment's reply and resolve them.
    CommentResolvedandReplies(doc);

    // Remove all comments.
    RemoveComments(doc);
    std::cout << "All comments are removed!" << std::endl;
    
    System::String outputPath = outputDataDir + u"ProcessComments.doc";
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:ProcessComments
    std::cout << "Comments extracted and removed successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ProcessComments example finished." << std::endl << std::endl;
}
