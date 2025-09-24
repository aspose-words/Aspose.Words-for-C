// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExComment.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/CommentCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1319098108u, ::Aspose::Words::ApiExamples::ExComment::CommentInfoPrinter, ThisTypeBaseTypesInfo);

ExComment::CommentInfoPrinter::CommentInfoPrinter() : mVisitorIsInsideComment(false), mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideComment = false;
}

System::String ExComment::CommentInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideComment)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->get_Text() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentRangeStart(System::SharedPtr<Aspose::Words::CommentRangeStart> commentRangeStart)
{
    IndentAndAppendLine(System::String(u"[Comment range start] ID: ") + commentRangeStart->get_Id());
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentRangeEnd(System::SharedPtr<Aspose::Words::CommentRangeEnd> commentRangeEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(System::String(u"[Comment range end] ID: ") + commentRangeEnd->get_Id() + u"\n");
    mVisitorIsInsideComment = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentStart(System::SharedPtr<Aspose::Words::Comment> comment)
{
    IndentAndAppendLine(System::String::Format(u"[Comment start] For comment range ID {0}, By {1} on {2}", comment->get_Id(), comment->get_Author(), comment->get_DateTime()));
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;
    
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExComment::CommentInfoPrinter::VisitCommentEnd(System::SharedPtr<Aspose::Words::Comment> comment)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Comment end]");
    mVisitorIsInsideComment = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExComment::CommentInfoPrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}


RTTI_INFO_IMPL_HASH(2258469839u, ::Aspose::Words::ApiExamples::ExComment, ThisTypeBaseTypesInfo);

void ExComment::PrintAllCommentInfo(System::SharedPtr<Aspose::Words::NodeCollection> comments)
{
    auto commentVisitor = System::MakeObject<Aspose::Words::ApiExamples::ExComment::CommentInfoPrinter>();
    
    // Iterate over all top-level comments. Unlike reply-type comments, top-level comments have no ancestor.
    for (auto&& comment : System::IterateOver<Aspose::Words::Comment>(comments->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> c)>>([](System::SharedPtr<Aspose::Words::Node> c) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Comment>(c))->get_Ancestor() == nullptr;
    })))->LINQ_ToList()))
    {
        // First, visit the start of the comment range.
        auto commentRangeStart = System::ExplicitCast<Aspose::Words::CommentRangeStart>(comment->get_PreviousSibling()->get_PreviousSibling()->get_PreviousSibling());
        commentRangeStart->Accept(commentVisitor);
        
        // Then, visit the comment, and any replies that it may have.
        comment->Accept(commentVisitor);
        // Visit only start of the comment.
        comment->AcceptStart(commentVisitor);
        // Visit only end of the comment.
        comment->AcceptEnd(commentVisitor);
        
        for (auto&& reply : System::IterateOver<Aspose::Words::Comment>(comment->get_Replies()))
        {
            reply->Accept(commentVisitor);
        }
        
        // Finally, visit the end of the comment range, and then print the visitor's text contents.
        auto commentRangeEnd = System::ExplicitCast<Aspose::Words::CommentRangeEnd>(comment->get_PreviousSibling());
        commentRangeEnd->Accept(commentVisitor);
        
        std::cout << commentVisitor->GetText() << std::endl;
    }
}


namespace gtest_test
{

class ExComment : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExComment> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExComment>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExComment> ExComment::s_instance;

} // namespace gtest_test

void ExComment::AddCommentWithReply()
{
    //ExStart
    //ExFor:Comment
    //ExFor:Comment.SetText(String)
    //ExFor:Comment.AddReply(String, String, DateTime, String)
    //ExSummary:Shows how to add a comment to a document, and then reply to it.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto comment = System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
    comment->SetText(u"My comment.");
    
    // Place the comment at a node in the document's body.
    // This comment will show up at the location of its paragraph,
    // outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
    builder->get_CurrentParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(comment);
    
    // Add a reply, which will show up under its parent comment.
    comment->AddReply(u"Joe Bloggs", u"J.B.", System::DateTime::get_Now(), u"New reply");
    
    // Comments and replies are both Comment nodes.
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Comment, true)->get_Count());
    
    // Comments that do not reply to other comments are "top-level". They have no ancestor comments.
    ASSERT_TRUE(System::TestTools::IsNull(comment->get_Ancestor()));
    
    // Replies have an ancestor top-level comment.
    ASPOSE_ASSERT_EQ(comment, comment->get_Replies()->idx_get(0)->get_Ancestor());
    
    doc->Save(get_ArtifactsDir() + u"Comment.AddCommentWithReply.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Comment.AddCommentWithReply.docx");
    auto docComment = System::ExplicitCast<Aspose::Words::Comment>(doc->GetChild(Aspose::Words::NodeType::Comment, 0, true));
    
    ASSERT_EQ(1, docComment->get_Count());
    ASSERT_EQ(1, comment->get_Replies()->get_Count());
    
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

void ExComment::PrintAllComments()
{
    //ExStart
    //ExFor:Comment.Ancestor
    //ExFor:Comment.Author
    //ExFor:Comment.Replies
    //ExFor:CompositeNode.GetEnumerator
    //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
    //ExSummary:Shows how to print all of a document's comments and their replies.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Comments.docx");
    
    System::SharedPtr<Aspose::Words::NodeCollection> comments = doc->GetChildNodes(Aspose::Words::NodeType::Comment, true);
    ASSERT_EQ(12, comments->get_Count());
    //ExSkip
    
    // If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
    // Print all top-level comments along with any replies they may have.
    for (auto&& comment : comments->LINQ_OfType<System::SharedPtr<Aspose::Words::Comment> >()->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Comment>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Comment> c)>>([](System::SharedPtr<Aspose::Words::Comment> c) -> bool
    {
        return c->get_Ancestor() == nullptr;
    })))->LINQ_ToList())
    {
        std::cout << "Top-level comment:" << std::endl;
        std::cout << System::String::Format(u"\t\"{0}\", by {1}", comment->GetText().Trim(), comment->get_Author()) << std::endl;
        std::cout << System::String::Format(u"Has {0} replies", comment->get_Replies()->get_Count()) << std::endl;
        for (auto&& commentReply : System::IterateOver<Aspose::Words::Comment>(comment->get_Replies()))
        {
            std::cout << System::String::Format(u"\t\"{0}\", by {1}", commentReply->GetText().Trim(), commentReply->get_Author()) << std::endl;
        }
        std::cout << std::endl;
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExComment, PrintAllComments)
{
    s_instance->PrintAllComments();
}

} // namespace gtest_test

void ExComment::RemoveCommentReplies()
{
    //ExStart
    //ExFor:Comment.RemoveAllReplies
    //ExFor:Comment.RemoveReply(Comment)
    //ExFor:CommentCollection.Item(Int32)
    //ExSummary:Shows how to remove comment replies.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto comment = System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
    comment->SetText(u"My comment.");
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(comment);
    
    comment->AddReply(u"Joe Bloggs", u"J.B.", System::DateTime::get_Now(), u"New reply");
    comment->AddReply(u"Joe Bloggs", u"J.B.", System::DateTime::get_Now(), u"Another reply");
    
    ASSERT_EQ(2, comment->get_Replies()->get_Count());
    
    // Below are two ways of removing replies from a comment.
    // 1 -  Use the "RemoveReply" method to remove replies from a comment individually:
    comment->RemoveReply(comment->get_Replies()->idx_get(0));
    
    ASSERT_EQ(1, comment->get_Replies()->get_Count());
    
    // 2 -  Use the "RemoveAllReplies" method to remove all replies from a comment at once:
    comment->RemoveAllReplies();
    
    ASSERT_EQ(0, comment->get_Replies()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExComment, RemoveCommentReplies)
{
    s_instance->RemoveCommentReplies();
}

} // namespace gtest_test

void ExComment::Done()
{
    //ExStart
    //ExFor:Comment.Done
    //ExFor:CommentCollection
    //ExSummary:Shows how to mark a comment as "done".
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Helo world!");
    
    // Insert a comment to point out an error. 
    auto comment = System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
    comment->SetText(u"Fix the spelling error!");
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(comment);
    
    // Comments have a "Done" flag, which is set to "false" by default. 
    // If a comment suggests that we make a change within the document,
    // we can apply the change, and then also set the "Done" flag afterwards to indicate the correction.
    ASSERT_FALSE(comment->get_Done());
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Hello world!");
    comment->set_Done(true);
    
    // Comments that are "done" will differentiate themselves
    // from ones that are not "done" with a faded text color.
    comment = System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
    comment->SetText(u"Add text to this paragraph.");
    builder->get_CurrentParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(comment);
    
    doc->Save(get_ArtifactsDir() + u"Comment.Done.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Comment.Done.docx");
    comment = System::ExplicitCast<Aspose::Words::Comment>(doc->GetChildNodes(Aspose::Words::NodeType::Comment, true)->idx_get(0));
    
    ASSERT_TRUE(comment->get_Done());
    ASSERT_EQ(u"\u0005Fix the spelling error!", comment->GetText().Trim());
    ASSERT_EQ(u"Hello world!", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Text());
}

namespace gtest_test
{

TEST_F(ExComment, Done)
{
    s_instance->Done();
}

} // namespace gtest_test

void ExComment::CreateCommentsAndPrintAllInfo()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto newComment = System::MakeObject<Aspose::Words::Comment>(doc);
    newComment->set_Author(u"VDeryushev");
    newComment->set_Initial(u"VD");
    newComment->set_DateTime(System::DateTime::get_Now());
    
    newComment->SetText(u"Comment regarding text.");
    
    // Add text to the document, warp it in a comment range, and then add your comment.
    System::SharedPtr<Aspose::Words::Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    para->AppendChild<System::SharedPtr<Aspose::Words::CommentRangeStart>>(System::MakeObject<Aspose::Words::CommentRangeStart>(doc, newComment->get_Id()));
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Commented text."));
    para->AppendChild<System::SharedPtr<Aspose::Words::CommentRangeEnd>>(System::MakeObject<Aspose::Words::CommentRangeEnd>(doc, newComment->get_Id()));
    para->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(newComment);
    
    // Add two replies to the comment.
    newComment->AddReply(u"John Doe", u"JD", System::DateTime::get_Now(), u"New reply.");
    newComment->AddReply(u"John Doe", u"JD", System::DateTime::get_Now(), u"Another reply.");
    
    PrintAllCommentInfo(doc->GetChildNodes(Aspose::Words::NodeType::Comment, true));
}

namespace gtest_test
{

TEST_F(ExComment, CreateCommentsAndPrintAllInfo)
{
    s_instance->CreateCommentsAndPrintAllInfo();
}

} // namespace gtest_test

void ExComment::UtcDateTime()
{
    //ExStart:UtcDateTime
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:Comment.DateTimeUtc
    //ExSummary:Shows how to get UTC date and time.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::DateTime dateTime = System::DateTime::get_Now();
    auto comment = System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"J.D.", dateTime);
    comment->SetText(u"My comment.");
    
    builder->get_CurrentParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(comment);
    
    doc->Save(get_ArtifactsDir() + u"Comment.UtcDateTime.docx");
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Comment.UtcDateTime.docx");
    
    comment = System::ExplicitCast<Aspose::Words::Comment>(doc->GetChild(Aspose::Words::NodeType::Comment, 0, true));
    // DateTimeUtc return data without milliseconds.
    ASSERT_EQ(dateTime.ToUniversalTime().ToString(u"yyyy-MM-dd hh:mm:ss"), comment->get_DateTimeUtc().ToString(u"yyyy-MM-dd hh:mm:ss"));
    //ExEnd:UtcDateTime
}

namespace gtest_test
{

TEST_F(ExComment, UtcDateTime)
{
    s_instance->UtcDateTime();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
