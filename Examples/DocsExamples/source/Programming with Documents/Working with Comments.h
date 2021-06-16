#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Comment.h>
#include <Aspose.Words.Cpp/CommentCollection.h>
#include <Aspose.Words.Cpp/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/CommentRangeStart.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithComments : public DocsExamplesBase
{
public:
    void AddComments()
    {
        //ExStart:AddComments
        //ExStart:CreateSimpleDocumentUsingDocumentBuilder
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Some text is added.");
        //ExEnd:CreateSimpleDocumentUsingDocumentBuilder

        auto comment = MakeObject<Comment>(doc, u"Awais Hafeez", u"AH", System::DateTime::get_Today());

        builder->get_CurrentParagraph()->AppendChild(comment);

        comment->get_Paragraphs()->Add(MakeObject<Paragraph>(doc));
        comment->get_FirstParagraph()->get_Runs()->Add(MakeObject<Run>(doc, u"Comment text."));

        doc->Save(ArtifactsDir + u"WorkingWithComments.AddComments.docx");
        //ExEnd:AddComments
    }

    void AnchorComment()
    {
        //ExStart:AnchorComment
        auto doc = MakeObject<Document>();

        auto para1 = MakeObject<Paragraph>(doc);
        auto run1 = MakeObject<Run>(doc, u"Some ");
        auto run2 = MakeObject<Run>(doc, u"text ");
        para1->AppendChild(run1);
        para1->AppendChild(run2);
        doc->get_FirstSection()->get_Body()->AppendChild(para1);

        auto para2 = MakeObject<Paragraph>(doc);
        auto run3 = MakeObject<Run>(doc, u"is ");
        auto run4 = MakeObject<Run>(doc, u"added ");
        para2->AppendChild(run3);
        para2->AppendChild(run4);
        doc->get_FirstSection()->get_Body()->AppendChild(para2);

        auto comment = MakeObject<Comment>(doc, u"Awais Hafeez", u"AH", System::DateTime::get_Today());
        comment->get_Paragraphs()->Add(MakeObject<Paragraph>(doc));
        comment->get_FirstParagraph()->get_Runs()->Add(MakeObject<Run>(doc, u"Comment text."));

        auto commentRangeStart = MakeObject<CommentRangeStart>(doc, comment->get_Id());
        auto commentRangeEnd = MakeObject<CommentRangeEnd>(doc, comment->get_Id());

        run1->get_ParentNode()->InsertAfter(commentRangeStart, run1);
        run3->get_ParentNode()->InsertAfter(commentRangeEnd, run3);
        commentRangeEnd->get_ParentNode()->InsertAfter(comment, commentRangeEnd);

        doc->Save(ArtifactsDir + u"WorkingWithComments.AnchorComment.doc");
        //ExEnd:AnchorComment
    }

    void AddRemoveCommentReply()
    {
        //ExStart:AddRemoveCommentReply
        auto doc = MakeObject<Document>(MyDir + u"Comments.docx");

        auto comment = System::DynamicCast<Comment>(doc->GetChild(NodeType::Comment, 0, true));
        comment->RemoveReply(comment->get_Replies()->idx_get(0));

        comment->AddReply(u"John Doe", u"JD", System::DateTime(2017, 9, 25, 12, 15, 0), u"New reply");

        doc->Save(ArtifactsDir + u"WorkingWithComments.AddRemoveCommentReply.docx");
        //ExEnd:AddRemoveCommentReply
    }

    void ProcessComments()
    {
        //ExStart:ProcessComments
        auto doc = MakeObject<Document>(MyDir + u"Comments.docx");

        // Extract the information about the comments of all the authors.
        for (const auto& comment : ExtractComments(doc))
        {
            std::cout << comment;
        }

        // Remove comments by the "pm" author.
        RemoveComments(doc, u"pm");
        std::cout << "Comments from \"pm\" are removed!" << std::endl;

        // Extract the information about the comments of the "ks" author.
        for (const auto& comment : ExtractComments(doc, u"ks"))
        {
            std::cout << comment;
        }

        // Read the comment's reply and resolve them.
        CommentResolvedAndReplies(doc);

        // Remove all comments.
        RemoveComments(doc);
        std::cout << "All comments are removed!" << std::endl;

        doc->Save(ArtifactsDir + u"WorkingWithComments.ProcessComments.docx");
        //ExEnd:ProcessComments
    }

private:
    SharedPtr<System::Collections::Generic::List<String>> ExtractComments(SharedPtr<Document> doc)
    {
        SharedPtr<System::Collections::Generic::List<String>> collectedComments = MakeObject<System::Collections::Generic::List<String>>();
        SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);

        for (const auto& comment : System::IterateOver<Comment>(comments))
        {
            collectedComments->Add(comment->get_Author() + u" " + comment->get_DateTime() + u" " + comment->ToString(SaveFormat::Text));
        }

        return collectedComments;
    }

    SharedPtr<System::Collections::Generic::List<String>> ExtractComments(SharedPtr<Document> doc, String authorName)
    {
        SharedPtr<System::Collections::Generic::List<String>> collectedComments = MakeObject<System::Collections::Generic::List<String>>();
        SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);

        for (const auto& comment : System::IterateOver<Comment>(comments))
        {
            if (comment->get_Author() == authorName)
            {
                collectedComments->Add(comment->get_Author() + u" " + comment->get_DateTime() + u" " + comment->ToString(SaveFormat::Text));
            }
        }

        return collectedComments;
    }

    void RemoveComments(SharedPtr<Document> doc)
    {
        SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);

        comments->Clear();
    }

    void RemoveComments(SharedPtr<Document> doc, String authorName)
    {
        SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);

        // Look through all comments and remove those written by the authorName.
        for (int i = comments->get_Count() - 1; i >= 0; i--)
        {
            auto comment = System::DynamicCast<Comment>(comments->idx_get(i));
            if (comment->get_Author() == authorName)
            {
                comment->Remove();
            }
        }
    }

    void CommentResolvedAndReplies(SharedPtr<Document> doc)
    {
        SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);

        auto parentComment = System::DynamicCast<Comment>(comments->idx_get(0));
        for (const auto& childComment : System::IterateOver<Comment>(parentComment->get_Replies()))
        {
            // Get comment parent and status.
            std::cout << childComment->get_Ancestor()->get_Id() << std::endl;
            std::cout << System::Convert::ToString(childComment->get_Done()) << std::endl;

            // And update comment Done mark.
            childComment->set_Done(true);
        }
    }
};

}} // namespace DocsExamples::Programming_with_Documents
