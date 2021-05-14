#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Comment.h>
#include <Aspose.Words.Cpp/CommentCollection.h>
#include <Aspose.Words.Cpp/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <system/collections/ienumerable.h>
#include <system/date_time.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace ApiExamples {

class ExComment : public ApiExampleBase
{
public:
    void AddCommentWithReply()
    {
        //ExStart
        //ExFor:Comment
        //ExFor:Comment.SetText(String)
        //ExFor:Comment.AddReply(String, String, DateTime, String)
        //ExSummary:Shows how to add a comment to a document, and then reply to it.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto comment = MakeObject<Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
        comment->SetText(u"My comment.");

        // Place the comment at a node in the document's body.
        // This comment will show up at the location of its paragraph,
        // outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
        builder->get_CurrentParagraph()->AppendChild(comment);

        // Add a reply, which will show up under its parent comment.
        comment->AddReply(u"Joe Bloggs", u"J.B.", System::DateTime::get_Now(), u"New reply");

        // Comments and replies are both Comment nodes.
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Comment, true)->get_Count());

        // Comments that do not reply to other comments are "top-level". They have no ancestor comments.
        ASSERT_TRUE(comment->get_Ancestor() == nullptr);

        // Replies have an ancestor top-level comment.
        ASPOSE_ASSERT_EQ(comment, comment->get_Replies()->idx_get(0)->get_Ancestor());

        doc->Save(ArtifactsDir + u"Comment.AddCommentWithReply.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Comment.AddCommentWithReply.docx");
        auto docComment = System::DynamicCast<Comment>(doc->GetChild(NodeType::Comment, 0, true));

        ASSERT_EQ(1, docComment->get_Count());
        ASSERT_EQ(1, comment->get_Replies()->get_Count());

        ASSERT_EQ(u"\u0005My comment.\r", docComment->GetText());
        ASSERT_EQ(u"\u0005New reply\r", docComment->get_Replies()->idx_get(0)->GetText());
    }

    void PrintAllComments()
    {
        //ExStart
        //ExFor:Comment.Ancestor
        //ExFor:Comment.Author
        //ExFor:Comment.Replies
        //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
        //ExSummary:Shows how to print all of a document's comments and their replies.
        auto doc = MakeObject<Document>(MyDir + u"Comments.docx");

        SharedPtr<NodeCollection> comments = doc->GetChildNodes(NodeType::Comment, true);
        ASSERT_EQ(12, comments->get_Count());
        //ExSkip

        // If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
        // Print all top-level comments along with any replies they may have.
        auto hasNullAncestor = [](SharedPtr<Comment> c)
        {
            return c->get_Ancestor() == nullptr;
        };

        for (auto comment : System::IterateOver(comments->LINQ_OfType<SharedPtr<Comment>>()->LINQ_Where(hasNullAncestor)))
        {
            std::cout << "Top-level comment:" << std::endl;
            std::cout << "\t\"" << comment->GetText().Trim() << "\", by " << comment->get_Author() << std::endl;
            std::cout << "Has " << comment->get_Replies()->get_Count() << " replies" << std::endl;
            for (auto commentReply : System::IterateOver<Comment>(comment->get_Replies()))
            {
                std::cout << "\t\"" << commentReply->GetText().Trim() << "\", by " << commentReply->get_Author() << std::endl;
            }
            std::cout << std::endl;
        }
        //ExEnd
    }

    void RemoveCommentReplies()
    {
        //ExStart
        //ExFor:Comment.RemoveAllReplies
        //ExFor:Comment.RemoveReply(Comment)
        //ExFor:CommentCollection.Item(Int32)
        //ExSummary:Shows how to remove comment replies.
        auto doc = MakeObject<Document>();

        auto comment = MakeObject<Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
        comment->SetText(u"My comment.");

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(comment);

        comment->AddReply(u"Joe Bloggs", u"J.B.", System::DateTime::get_Now(), u"New reply");
        comment->AddReply(u"Joe Bloggs", u"J.B.", System::DateTime::get_Now(), u"Another reply");

        ASSERT_EQ(2, comment->get_Replies()->LINQ_Count());

        // Below are two ways of removing replies from a comment.
        // 1 -  Use the "RemoveReply" method to remove replies from a comment individually:
        comment->RemoveReply(comment->get_Replies()->idx_get(0));

        ASSERT_EQ(1, comment->get_Replies()->LINQ_Count());

        // 2 -  Use the "RemoveAllReplies" method to remove all replies from a comment at once:
        comment->RemoveAllReplies();

        ASSERT_EQ(0, comment->get_Replies()->LINQ_Count());
        //ExEnd
    }

    void Done()
    {
        //ExStart
        //ExFor:Comment.Done
        //ExFor:CommentCollection
        //ExSummary:Shows how to mark a comment as "done".
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Helo world!");

        // Insert a comment to point out an error.
        auto comment = MakeObject<Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
        comment->SetText(u"Fix the spelling error!");
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(comment);

        // Comments have a "Done" flag, which is set to "false" by default.
        // If a comment suggests that we make a change within the document,
        // we can apply the change, and then also set the "Done" flag afterwards to indicate the correction.
        ASSERT_FALSE(comment->get_Done());

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Hello world!");
        comment->set_Done(true);

        // Comments that are "done" will differentiate themselves
        // from ones that are not "done" with a faded text color.
        comment = MakeObject<Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
        comment->SetText(u"Add text to this paragraph.");
        builder->get_CurrentParagraph()->AppendChild(comment);

        doc->Save(ArtifactsDir + u"Comment.Done.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Comment.Done.docx");
        comment = System::DynamicCast<Comment>(doc->GetChildNodes(NodeType::Comment, true)->idx_get(0));

        ASSERT_TRUE(comment->get_Done());
        ASSERT_EQ(u"\u0005Fix the spelling error!", comment->GetText().Trim());
        ASSERT_EQ(u"Hello world!", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Text());
    }

    //ExStart
    //ExFor:Comment.Done
    //ExFor:Comment.#ctor(DocumentBase)
    //ExFor:Comment.Accept(DocumentVisitor)
    //ExFor:Comment.DateTime
    //ExFor:Comment.Id
    //ExFor:Comment.Initial
    //ExFor:CommentRangeEnd
    //ExFor:CommentRangeEnd.#ctor(DocumentBase,Int32)
    //ExFor:CommentRangeEnd.Accept(DocumentVisitor)
    //ExFor:CommentRangeEnd.Id
    //ExFor:CommentRangeStart
    //ExFor:CommentRangeStart.#ctor(DocumentBase,Int32)
    //ExFor:CommentRangeStart.Accept(DocumentVisitor)
    //ExFor:CommentRangeStart.Id
    //ExSummary:Shows how print the contents of all comments and their comment ranges using a document visitor.
    void CreateCommentsAndPrintAllInfo()
    {
        auto doc = MakeObject<Document>();

        auto newComment = MakeObject<Comment>(doc);
        newComment->set_Author(u"VDeryushev");
        newComment->set_Initial(u"VD");
        newComment->set_DateTime(System::DateTime::get_Now());

        newComment->SetText(u"Comment regarding text.");

        // Add text to the document, warp it in a comment range, and then add your comment.
        SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        para->AppendChild(MakeObject<CommentRangeStart>(doc, newComment->get_Id()));
        para->AppendChild(MakeObject<Run>(doc, u"Commented text."));
        para->AppendChild(MakeObject<CommentRangeEnd>(doc, newComment->get_Id()));
        para->AppendChild(newComment);

        // Add two replies to the comment.
        newComment->AddReply(u"John Doe", u"JD", System::DateTime::get_Now(), u"New reply.");
        newComment->AddReply(u"John Doe", u"JD", System::DateTime::get_Now(), u"Another reply.");

        PrintAllCommentInfo(doc->GetChildNodes(NodeType::Comment, true));
    }

    /// <summary>
    /// Iterates over every top-level comment and prints its comment range, contents, and replies.
    /// </summary>
    static void PrintAllCommentInfo(SharedPtr<NodeCollection> comments)
    {
        auto commentVisitor = MakeObject<ExComment::CommentInfoPrinter>();

        // Iterate over all top-level comments. Unlike reply-type comments, top-level comments have no ancestor.
        std::function<bool(SharedPtr<Node> c)> haveNoAncestor = [](SharedPtr<Node> c)
        {
            return (System::DynamicCast<Comment>(c))->get_Ancestor() == nullptr;
        };

        for (auto comment : System::IterateOver<Comment>(comments->LINQ_Where(haveNoAncestor)))
        {
            // First, visit the start of the comment range.
            auto commentRangeStart = System::DynamicCast<CommentRangeStart>(comment->get_PreviousSibling()->get_PreviousSibling()->get_PreviousSibling());
            commentRangeStart->Accept(commentVisitor);

            // Then, visit the comment, and any replies that it may have.
            comment->Accept(commentVisitor);

            for (auto reply : System::IterateOver<Comment>(comment->get_Replies()))
            {
                reply->Accept(commentVisitor);
            }

            // Finally, visit the end of the comment range, and then print the visitor's text contents.
            auto commentRangeEnd = System::DynamicCast<CommentRangeEnd>(comment->get_PreviousSibling());
            commentRangeEnd->Accept(commentVisitor);

            std::cout << commentVisitor->GetText() << std::endl;
        }
    }

    /// <summary>
    /// Prints information and contents of all comments and comment ranges encountered in the document.
    /// </summary>
    class CommentInfoPrinter : public DocumentVisitor
    {
    public:
        CommentInfoPrinter() : mVisitorIsInsideComment(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideComment = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        String GetText()
        {
            return mBuilder->ToString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (mVisitorIsInsideComment)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->get_Text() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        VisitorAction VisitCommentRangeStart(SharedPtr<CommentRangeStart> commentRangeStart) override
        {
            IndentAndAppendLine(String(u"[Comment range start] ID: ") + commentRangeStart->get_Id());
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        VisitorAction VisitCommentRangeEnd(SharedPtr<CommentRangeEnd> commentRangeEnd) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(String(u"[Comment range end] ID: ") + commentRangeEnd->get_Id() + u"\n");
            mVisitorIsInsideComment = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        VisitorAction VisitCommentStart(SharedPtr<Comment> comment) override
        {
            IndentAndAppendLine(
                String::Format(u"[Comment start] For comment range ID {0}, By {1} on {2}", comment->get_Id(), comment->get_Author(), comment->get_DateTime()));
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when the visiting of a Comment node is ended in the document.
        /// </summary>
        VisitorAction VisitCommentEnd(SharedPtr<Comment> comment) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Comment end]");
            mVisitorIsInsideComment = false;

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideComment;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd
};

} // namespace ApiExamples
