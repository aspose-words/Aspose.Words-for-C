// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExComment.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/CommentCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3466528438u, ::ApiExamples::ExComment::CommentInfoPrinter, ThisTypeBaseTypesInfo);

ExComment::CommentInfoPrinter::CommentInfoPrinter() : mVisitorIsInsideComment(false), mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideComment = false;
}

String ExComment::CommentInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideComment)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->get_Text() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentRangeStart(SharedPtr<Aspose::Words::CommentRangeStart> commentRangeStart)
{
    IndentAndAppendLine(String(u"[Comment range start] ID: ") + commentRangeStart->get_Id());
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentRangeEnd(SharedPtr<Aspose::Words::CommentRangeEnd> commentRangeEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(String(u"[Comment range end] ID: ") + commentRangeEnd->get_Id() + u"\n");
    mVisitorIsInsideComment = false;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentStart(SharedPtr<Aspose::Words::Comment> comment)
{
    IndentAndAppendLine(String::Format(u"[Comment start] For comment range ID {0}, By {1} on {2}",comment->get_Id(),comment->get_Author(),comment->get_DateTime()));
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentEnd(SharedPtr<Aspose::Words::Comment> comment)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Comment end]");
    mVisitorIsInsideComment = false;

    return Aspose::Words::VisitorAction::Continue;
}

void ExComment::CommentInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExComment::CommentInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExComment::CommentInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

SharedPtr<Aspose::Words::Comment> ExComment::CreateComment(SharedPtr<Aspose::Words::Document> doc, String author, String initials, System::DateTime dateTime, String text)
{
    auto newComment = MakeObject<Comment>(doc);
    newComment->set_Author(author);
    newComment->set_Initial(initials);
    newComment->set_DateTime(dateTime);
    newComment->SetText(text);

    return newComment;
}

SharedPtr<System::Collections::Generic::List<SharedPtr<Aspose::Words::Comment>>> ExComment::ExtractComments(SharedPtr<Aspose::Words::Document> doc)
{
    SharedPtr<System::Collections::Generic::List<SharedPtr<Comment>>> collectedComments = MakeObject<System::Collections::Generic::List<SharedPtr<Comment>>>();

    SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);

    for (auto comment : System::IterateOver<Comment>(comments))
    {
        // All replies have ancestor, so we will add this check
        if (comment->get_Ancestor() == nullptr)
        {
            collectedComments->Add(comment);
        }
    }

    return collectedComments;
}

void ExComment::PrintAllCommentInfo(SharedPtr<System::Collections::Generic::List<SharedPtr<Aspose::Words::Comment>>> comments)
{
    // Create an object that inherits from the DocumentVisitor class
    auto commentVisitor = MakeObject<ExComment::CommentInfoPrinter>();

    // Get the enumerator from the document's comment collection and iterate over the comments
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Comment>>> enumerator = comments->GetEnumerator();
        while (enumerator->MoveNext())
        {
            SharedPtr<Comment> currentComment = enumerator->get_Current();

            // Accept our DocumentVisitor it to print information about our comments
            if (currentComment != nullptr)
            {
                // Get CommentRangeStart from our current comment and construct its information
                auto commentRangeStart = System::DynamicCast<Aspose::Words::CommentRangeStart>(currentComment->get_PreviousSibling()->get_PreviousSibling()->get_PreviousSibling());
                commentRangeStart->Accept(commentVisitor);

                // Construct current comment information
                currentComment->Accept(commentVisitor);

                // Get CommentRangeEnd from our current comment and construct its information
                auto commentRangeEnd = System::DynamicCast<Aspose::Words::CommentRangeEnd>(currentComment->get_PreviousSibling());
                commentRangeEnd->Accept(commentVisitor);
            }
        }

        // Output of all information received
        System::Console::WriteLine(commentVisitor->GetText());
    }
}

namespace gtest_test
{

class ExComment : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExComment> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExComment>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExComment> ExComment::s_instance;

} // namespace gtest_test

void ExComment::AddCommentWithReply()
{
    //ExStart
    //ExFor:Comment
    //ExFor:Comment.SetText(String)
    //ExFor:Comment.AddReply(String, String, DateTime, String)
    //ExSummary:Shows how to add a comment with a reply to a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create new comment
    auto newComment = MakeObject<Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
    newComment->SetText(u"My comment.");

    // Add this comment to a document node
    builder->get_CurrentParagraph()->AppendChild(newComment);

    // Add comment reply
    newComment->AddReply(u"John Doe", u"JD", System::DateTime(2017, 9, 25, 12, 15, 0), u"New reply");
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    auto docComment = System::DynamicCast<Aspose::Words::Comment>(doc->GetChild(Aspose::Words::NodeType::Comment, 0, true));

    ASSERT_EQ(1, docComment->get_Count());
    ASSERT_EQ(1, newComment->get_Replies()->get_Count());

    ASSERT_EQ(u"\u0005My comment.\r", docComment->GetText());
    ASSERT_EQ(u"\u0005New reply\r", docComment->get_Replies()->idx_get(0)->GetText());
}

namespace gtest_test
{

TEST_F(ExComment, AddCommentWithReply)
{
    s_instance->AddCommentWithReply();
}

} // namespace gtest_test

void ExComment::GetAllCommentsAndReplies()
{
    //ExStart
    //ExFor:Comment.Ancestor
    //ExFor:Comment.Author
    //ExFor:Comment.Replies
    //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
    //ExSummary:Shows how to get all comments with all replies.
    auto doc = MakeObject<Document>(MyDir + u"Comments.docx");

    // Get all comment from the document
    SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);
    ASSERT_EQ(12, comments->get_Count());
    //ExSkip

    // For all comments and replies we identify comment level and info about it
    for (auto comment : System::IterateOver(comments->LINQ_OfType<SharedPtr<Comment> >()))
    {
        if (comment->get_Ancestor() == nullptr)
        {
            System::Console::WriteLine(u"\nThis is a top-level comment");
            System::Console::WriteLine(String(u"Comment author: ") + comment->get_Author());
            System::Console::WriteLine(String(u"Comment text: ") + comment->GetText());

            for (auto commentReply : System::IterateOver(comment->get_Replies()->LINQ_OfType<SharedPtr<Comment> >()))
            {
                System::Console::WriteLine(u"\n\tThis is a comment reply");
                System::Console::WriteLine(String(u"\tReply author: ") + commentReply->get_Author());
                System::Console::WriteLine(String(u"\tReply text: ") + commentReply->GetText());
            }
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExComment, GetAllCommentsAndReplies)
{
    s_instance->GetAllCommentsAndReplies();
}

} // namespace gtest_test

void ExComment::RemoveCommentReplies()
{
    //ExStart
    //ExFor:Comment.RemoveAllReplies
    //ExSummary:Shows how to remove comment replies.
    auto doc = MakeObject<Document>(MyDir + u"Comments.docx");

    SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);
    auto comment = System::DynamicCast<Aspose::Words::Comment>(comments->idx_get(0));
    ASSERT_EQ(2, comment->get_Replies()->LINQ_Count());
    //ExSkip

    comment->RemoveAllReplies();
    ASSERT_EQ(0, comment->get_Replies()->LINQ_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExComment, RemoveCommentReplies)
{
    s_instance->RemoveCommentReplies();
}

} // namespace gtest_test

void ExComment::RemoveCommentReply()
{
    //ExStart
    //ExFor:Comment.RemoveReply(Comment)
    //ExFor:CommentCollection.Item(Int32)
    //ExSummary:Shows how to remove specific comment reply.
    auto doc = MakeObject<Document>(MyDir + u"Comments.docx");

    SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);

    auto parentComment = System::DynamicCast<Aspose::Words::Comment>(comments->idx_get(0));
    SharedPtr<CommentCollection> repliesCollection = parentComment->get_Replies();
    ASSERT_EQ(2, parentComment->get_Replies()->LINQ_Count());
    //ExSkip

    // Remove the first reply to comment
    parentComment->RemoveReply(repliesCollection->idx_get(0));
    ASSERT_EQ(1, parentComment->get_Replies()->LINQ_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExComment, RemoveCommentReply)
{
    s_instance->RemoveCommentReply();
}

} // namespace gtest_test

void ExComment::MarkCommentRepliesDone()
{
    //ExStart
    //ExFor:Comment.Done
    //ExFor:CommentCollection
    //ExSummary:Shows how to mark comment as Done.
    auto doc = MakeObject<Document>(MyDir + u"Comments.docx");

    SharedPtr<NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);

    auto comment = System::DynamicCast<Aspose::Words::Comment>(comments->idx_get(0));
    SharedPtr<CommentCollection> repliesCollection = comment->get_Replies();

    for (auto childComment : System::IterateOver<Comment>(repliesCollection))
    {
        if (!childComment->get_Done())
        {
            // Update comment reply Done mark
            childComment->set_Done(true);
        }
    }
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    comment = System::DynamicCast<Aspose::Words::Comment>(doc->GetChildNodes(Aspose::Words::NodeType::Comment, true)->idx_get(0));
    repliesCollection = comment->get_Replies();

    for (auto childComment : System::IterateOver<Comment>(repliesCollection))
    {
        ASSERT_TRUE(childComment->get_Done());
    }
}

namespace gtest_test
{

TEST_F(ExComment, MarkCommentRepliesDone)
{
    s_instance->MarkCommentRepliesDone();
}

} // namespace gtest_test

void ExComment::CreateCommentsAndPrintAllInfo()
{
    auto doc = MakeObject<Document>();
    doc->RemoveAllChildren();

    auto sect = System::DynamicCast<Aspose::Words::Section>(doc->AppendChild(MakeObject<Section>(doc)));
    auto body = System::DynamicCast<Aspose::Words::Body>(sect->AppendChild(MakeObject<Body>(doc)));

    // Create a commented text with several comment replies
    for (int i = 0; i <= 10; i++)
    {
        SharedPtr<Comment> newComment = CreateComment(doc, u"VDeryushev", u"VD", System::DateTime::get_Now(), String(u"My test comment ") + i);

        auto para = System::DynamicCast<Aspose::Words::Paragraph>(body->AppendChild(MakeObject<Paragraph>(doc)));
        para->AppendChild(MakeObject<CommentRangeStart>(doc, newComment->get_Id()));
        para->AppendChild(MakeObject<Run>(doc, String(u"Commented text ") + i));
        para->AppendChild(MakeObject<CommentRangeEnd>(doc, newComment->get_Id()));
        para->AppendChild(newComment);

        for (int y = 0; y <= 2; y++)
        {
            newComment->AddReply(u"John Doe", u"JD", System::DateTime::get_Now(), String(u"New reply ") + y);
        }
    }

    // Look at information of our comments
    PrintAllCommentInfo(ExtractComments(doc));
}

namespace gtest_test
{

TEST_F(ExComment, CreateCommentsAndPrintAllInfo)
{
    s_instance->CreateCommentsAndPrintAllInfo();
}

} // namespace gtest_test

} // namespace ApiExamples
