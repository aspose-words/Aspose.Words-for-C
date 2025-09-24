// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRevision.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/collections/ienumerator.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroupCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroup.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldDate.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Document/RevisionsView.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Comparing/ComparisonTargetType.h>
#include <Aspose.Words.Cpp/Model/Comparing/CompareOptions.h>
#include <Aspose.Words.Cpp/Model/Comparing/AdvancedCompareOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/ShowInBalloons.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionTextEffect.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionColor.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Comparing;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4004401701u, ::Aspose::Words::ApiExamples::ExRevision::RevisionCriteria, ThisTypeBaseTypesInfo);

ExRevision::RevisionCriteria::RevisionCriteria(System::String authorName, Aspose::Words::RevisionType revisionType)
    : RevisionType(((Aspose::Words::RevisionType)0))
{
    AuthorName = authorName;
    RevisionType = revisionType;
}

bool ExRevision::RevisionCriteria::IsMatch(System::SharedPtr<Aspose::Words::Revision> revision)
{
    return revision->get_Author() == AuthorName && revision->get_RevisionType() == RevisionType;
}


RTTI_INFO_IMPL_HASH(363893867u, ::Aspose::Words::ApiExamples::ExRevision, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExRevision : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExRevision> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExRevision>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExRevision> ExRevision::s_instance;

} // namespace gtest_test

void ExRevision::Revisions()
{
    //ExStart
    //ExFor:Revision
    //ExFor:Revision.Accept
    //ExFor:Revision.Author
    //ExFor:Revision.DateTime
    //ExFor:Revision.Group
    //ExFor:Revision.Reject
    //ExFor:Revision.RevisionType
    //ExFor:RevisionCollection
    //ExFor:RevisionCollection.Item(Int32)
    //ExFor:RevisionCollection.Count
    //ExFor:RevisionType
    //ExFor:Document.HasRevisions
    //ExFor:Document.TrackRevisions
    //ExFor:Document.Revisions
    //ExSummary:Shows how to work with revisions in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Normal editing of the document does not count as a revision.
    builder->Write(u"This does not count as a revision. ");
    
    ASSERT_FALSE(doc->get_HasRevisions());
    
    // To register our edits as revisions, we need to declare an author, and then start tracking them.
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
    
    builder->Write(u"This is revision #1. ");
    
    ASSERT_TRUE(doc->get_HasRevisions());
    ASSERT_EQ(1, doc->get_Revisions()->get_Count());
    
    // This flag corresponds to the "Review" -> "Tracking" -> "Track Changes" option in Microsoft Word.
    // The "StartTrackRevisions" method does not affect its value,
    // and the document is tracking revisions programmatically despite it having a value of "false".
    // If we open this document using Microsoft Word, it will not be tracking revisions.
    ASSERT_FALSE(doc->get_TrackRevisions());
    
    // We have added text using the document builder, so the first revision is an insertion-type revision.
    System::SharedPtr<Aspose::Words::Revision> revision = doc->get_Revisions()->idx_get(0);
    ASSERT_EQ(u"John Doe", revision->get_Author());
    ASSERT_EQ(u"This is revision #1. ", revision->get_ParentNode()->GetText());
    ASSERT_EQ(Aspose::Words::RevisionType::Insertion, revision->get_RevisionType());
    ASSERT_EQ(revision->get_DateTime().get_Date(), System::DateTime::get_Now().get_Date());
    ASPOSE_ASSERT_EQ(doc->get_Revisions()->get_Groups()->idx_get(0), revision->get_Group());
    
    // Remove a run to create a deletion-type revision.
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->Remove();
    
    // Adding a new revision places it at the beginning of the revision collection.
    ASSERT_EQ(Aspose::Words::RevisionType::Deletion, doc->get_Revisions()->idx_get(0)->get_RevisionType());
    ASSERT_EQ(2, doc->get_Revisions()->get_Count());
    
    // Insert revisions show up in the document body even before we accept/reject the revision.
    // Rejecting the revision will remove its nodes from the body. Conversely, nodes that make up delete revisions
    // also linger in the document until we accept the revision.
    ASSERT_EQ(u"This does not count as a revision. This is revision #1.", doc->GetText().Trim());
    
    // Accepting the delete revision will remove its parent node from the paragraph text
    // and then remove the collection's revision itself.
    doc->get_Revisions()->idx_get(0)->Accept();
    
    ASSERT_EQ(1, doc->get_Revisions()->get_Count());
    ASSERT_EQ(u"This is revision #1.", doc->GetText().Trim());
    
    builder->Writeln(u"");
    builder->Write(u"This is revision #2.");
    
    // Now move the node to create a moving revision type.
    System::SharedPtr<Aspose::Words::Node> node = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1);
    System::SharedPtr<Aspose::Words::Node> endNode = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_NextSibling();
    System::SharedPtr<Aspose::Words::Node> referenceNode = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0);
    
    while (node != endNode)
    {
        System::SharedPtr<Aspose::Words::Node> nextNode = node->get_NextSibling();
        doc->get_FirstSection()->get_Body()->InsertBefore<System::SharedPtr<Aspose::Words::Node>>(node, referenceNode);
        node = nextNode;
    }
    
    ASSERT_EQ(Aspose::Words::RevisionType::Moving, doc->get_Revisions()->idx_get(0)->get_RevisionType());
    ASSERT_EQ(8, doc->get_Revisions()->get_Count());
    ASSERT_EQ(u"This is revision #2.\rThis is revision #1. \rThis is revision #2.", doc->GetText().Trim());
    
    // The moving revision is now at index 1. Reject the revision to discard its contents.
    doc->get_Revisions()->idx_get(1)->Reject();
    
    ASSERT_EQ(6, doc->get_Revisions()->get_Count());
    ASSERT_EQ(u"This is revision #1. \rThis is revision #2.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, Revisions)
{
    s_instance->Revisions();
}

} // namespace gtest_test

void ExRevision::RevisionCollection()
{
    //ExStart
    //ExFor:Revision.ParentStyle
    //ExFor:RevisionCollection.GetEnumerator
    //ExFor:RevisionCollection.Groups
    //ExFor:RevisionCollection.RejectAll
    //ExFor:RevisionGroupCollection.GetEnumerator
    //ExSummary:Shows how to work with a document's collection of revisions.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revisions.docx");
    System::SharedPtr<Aspose::Words::RevisionCollection> revisions = doc->get_Revisions();
    
    // This collection itself has a collection of revision groups.
    // Each group is a sequence of adjacent revisions.
    ASSERT_EQ(7, revisions->get_Groups()->get_Count());
    //ExSkip
    std::cout << System::String::Format(u"{0} revision groups:", revisions->get_Groups()->get_Count()) << std::endl;
    
    // Iterate over the collection of groups and print the text that the revision concerns.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::RevisionGroup>>> e = revisions->get_Groups()->GetEnumerator();
        while (e->MoveNext())
        {
            std::cout << (System::String::Format(u"\tGroup type \"{0}\", ", e->get_Current()->get_RevisionType()) + System::String::Format(u"author: {0}, contents: [{1}]", e->get_Current()->get_Author(), e->get_Current()->get_Text().Trim())) << std::endl;
        }
    }
    
    // Each Run that a revision affects gets a corresponding Revision object.
    // The revisions' collection is considerably larger than the condensed form we printed above,
    // depending on how many Runs we have segmented the document into during Microsoft Word editing.
    ASSERT_EQ(11, revisions->get_Count());
    //ExSkip
    std::cout << System::String::Format(u"\n{0} revisions:", revisions->get_Count()) << std::endl;
    
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Revision>>> e = revisions->GetEnumerator();
        while (e->MoveNext())
        {
            // A StyleDefinitionChange strictly affects styles and not document nodes. This means the "ParentStyle"
            // property will always be in use, while the ParentNode will always be null.
            // Since all other changes affect nodes, ParentNode will conversely be in use, and ParentStyle will be null.
            if (e->get_Current()->get_RevisionType() == Aspose::Words::RevisionType::StyleDefinitionChange)
            {
                std::cout << (System::String::Format(u"\tRevision type \"{0}\", ", e->get_Current()->get_RevisionType()) + System::String::Format(u"author: {0}, style: [{1}]", e->get_Current()->get_Author(), e->get_Current()->get_ParentStyle()->get_Name())) << std::endl;
            }
            else
            {
                std::cout << (System::String::Format(u"\tRevision type \"{0}\", ", e->get_Current()->get_RevisionType()) + System::String::Format(u"author: {0}, contents: [{1}]", e->get_Current()->get_Author(), e->get_Current()->get_ParentNode()->GetText().Trim())) << std::endl;
            }
        }
    }
    
    // Reject all revisions via the collection, reverting the document to its original form.
    revisions->RejectAll();
    
    ASSERT_EQ(0, revisions->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, RevisionCollection)
{
    s_instance->RevisionCollection();
}

} // namespace gtest_test

void ExRevision::GetInfoAboutRevisionsInRevisionGroups()
{
    //ExStart
    //ExFor:RevisionGroup
    //ExFor:RevisionGroup.Author
    //ExFor:RevisionGroup.RevisionType
    //ExFor:RevisionGroup.Text
    //ExFor:RevisionGroupCollection
    //ExFor:RevisionGroupCollection.Count
    //ExSummary:Shows how to print info about a group of revisions in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revisions.docx");
    
    ASSERT_EQ(7, doc->get_Revisions()->get_Groups()->get_Count());
    
    for (auto&& group : doc->get_Revisions()->get_Groups())
    {
        std::cout << System::String::Format(u"Revision author: {0}; Revision type: {1} \n\tRevision text: {2}", group->get_Author(), group->get_RevisionType(), group->get_Text()) << std::endl;
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, GetInfoAboutRevisionsInRevisionGroups)
{
    s_instance->GetInfoAboutRevisionsInRevisionGroups();
}

} // namespace gtest_test

void ExRevision::GetSpecificRevisionGroup()
{
    //ExStart
    //ExFor:RevisionGroupCollection
    //ExFor:RevisionGroupCollection.Item(Int32)
    //ExSummary:Shows how to get a group of revisions in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revisions.docx");
    
    System::SharedPtr<Aspose::Words::RevisionGroup> revisionGroup = doc->get_Revisions()->get_Groups()->idx_get(0);
    //ExEnd
    
    ASSERT_EQ(Aspose::Words::RevisionType::Deletion, revisionGroup->get_RevisionType());
    ASSERT_EQ(u"Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ", revisionGroup->get_Text());
}

namespace gtest_test
{

TEST_F(ExRevision, GetSpecificRevisionGroup)
{
    s_instance->GetSpecificRevisionGroup();
}

} // namespace gtest_test

void ExRevision::ShowRevisionBalloons()
{
    //ExStart
    //ExFor:RevisionOptions.ShowInBalloons
    //ExSummary:Shows how to display revisions in balloons.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revisions.docx");
    
    // By default, text that is a revision has a different color to differentiate it from the other non-revision text.
    // Set a revision option to show more details about each revision in a balloon on the page's right margin.
    doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowInBalloons(Aspose::Words::Layout::ShowInBalloons::FormatAndDelete);
    doc->Save(get_ArtifactsDir() + u"Revision.ShowRevisionBalloons.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, ShowRevisionBalloons)
{
    s_instance->ShowRevisionBalloons();
}

} // namespace gtest_test

void ExRevision::RevisionOptions()
{
    //ExStart
    //ExFor:ShowInBalloons
    //ExFor:RevisionOptions.ShowInBalloons
    //ExFor:RevisionOptions.CommentColor
    //ExFor:RevisionOptions.DeletedTextColor
    //ExFor:RevisionOptions.DeletedTextEffect
    //ExFor:RevisionOptions.InsertedTextEffect
    //ExFor:RevisionOptions.MovedFromTextColor
    //ExFor:RevisionOptions.MovedFromTextEffect
    //ExFor:RevisionOptions.MovedToTextColor
    //ExFor:RevisionOptions.MovedToTextEffect
    //ExFor:RevisionOptions.RevisedPropertiesColor
    //ExFor:RevisionOptions.RevisedPropertiesEffect
    //ExFor:RevisionOptions.RevisionBarsColor
    //ExFor:RevisionOptions.RevisionBarsWidth
    //ExFor:RevisionOptions.ShowOriginalRevision
    //ExFor:RevisionOptions.ShowRevisionMarks
    //ExFor:RevisionTextEffect
    //ExSummary:Shows how to modify the appearance of revisions.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revisions.docx");
    
    // Get the RevisionOptions object that controls the appearance of revisions.
    System::SharedPtr<Aspose::Words::Layout::RevisionOptions> revisionOptions = doc->get_LayoutOptions()->get_RevisionOptions();
    
    // Render insertion revisions in green and italic.
    revisionOptions->set_InsertedTextColor(Aspose::Words::Layout::RevisionColor::Green);
    revisionOptions->set_InsertedTextEffect(Aspose::Words::Layout::RevisionTextEffect::Italic);
    
    // Render deletion revisions in red and bold.
    revisionOptions->set_DeletedTextColor(Aspose::Words::Layout::RevisionColor::Red);
    revisionOptions->set_DeletedTextEffect(Aspose::Words::Layout::RevisionTextEffect::Bold);
    
    // The same text will appear twice in a movement revision:
    // once at the departure point and once at the arrival destination.
    // Render the text at the moved-from revision yellow with a double strike through
    // and double-underlined blue at the moved-to revision.
    revisionOptions->set_MovedFromTextColor(Aspose::Words::Layout::RevisionColor::Yellow);
    revisionOptions->set_MovedFromTextEffect(Aspose::Words::Layout::RevisionTextEffect::DoubleStrikeThrough);
    revisionOptions->set_MovedToTextColor(Aspose::Words::Layout::RevisionColor::ClassicBlue);
    revisionOptions->set_MovedToTextEffect(Aspose::Words::Layout::RevisionTextEffect::DoubleUnderline);
    
    // Render format revisions in dark red and bold.
    revisionOptions->set_RevisedPropertiesColor(Aspose::Words::Layout::RevisionColor::DarkRed);
    revisionOptions->set_RevisedPropertiesEffect(Aspose::Words::Layout::RevisionTextEffect::Bold);
    
    // Place a thick dark blue bar on the left side of the page next to lines affected by revisions.
    revisionOptions->set_RevisionBarsColor(Aspose::Words::Layout::RevisionColor::DarkBlue);
    revisionOptions->set_RevisionBarsWidth(15.0f);
    
    // Show revision marks and original text.
    revisionOptions->set_ShowOriginalRevision(true);
    revisionOptions->set_ShowRevisionMarks(true);
    
    // Get movement, deletion, formatting revisions, and comments to show up in green balloons
    // on the right side of the page.
    revisionOptions->set_ShowInBalloons(Aspose::Words::Layout::ShowInBalloons::Format);
    revisionOptions->set_CommentColor(Aspose::Words::Layout::RevisionColor::BrightGreen);
    
    // These features are only applicable to formats such as .pdf or .jpg.
    doc->Save(get_ArtifactsDir() + u"Revision.RevisionOptions.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, RevisionOptions)
{
    s_instance->RevisionOptions();
}

} // namespace gtest_test

void ExRevision::RevisionSpecifiedCriteria()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"This does not count as a revision. ");
    
    // To register our edits as revisions, we need to declare an author, and then start tracking them.
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
    builder->Write(u"This is insertion revision #1. ");
    doc->StopTrackRevisions();
    
    doc->StartTrackRevisions(u"Jane Doe", System::DateTime::get_Now());
    builder->Write(u"This is insertion revision #2. ");
    // Remove a run "This does not count as a revision.".
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->Remove();
    doc->StopTrackRevisions();
    
    ASSERT_EQ(3, doc->get_Revisions()->get_Count());
    // We have two revisions from different authors, so we need to accept only one.
    doc->get_Revisions()->Accept(System::MakeObject<Aspose::Words::ApiExamples::ExRevision::RevisionCriteria>(u"John Doe", Aspose::Words::RevisionType::Insertion));
    ASSERT_EQ(2, doc->get_Revisions()->get_Count());
    // Reject revision with different author name and revision type.
    doc->get_Revisions()->Reject(System::MakeObject<Aspose::Words::ApiExamples::ExRevision::RevisionCriteria>(u"Jane Doe", Aspose::Words::RevisionType::Deletion));
    ASSERT_EQ(1, doc->get_Revisions()->get_Count());
    
    doc->Save(get_ArtifactsDir() + u"Revision.RevisionSpecifiedCriteria.docx");
}

namespace gtest_test
{

TEST_F(ExRevision, RevisionSpecifiedCriteria)
{
    s_instance->RevisionSpecifiedCriteria();
}

} // namespace gtest_test

void ExRevision::TrackRevisions()
{
    //ExStart
    //ExFor:Document.StartTrackRevisions(String)
    //ExFor:Document.StartTrackRevisions(String, DateTime)
    //ExFor:Document.StopTrackRevisions
    //ExSummary:Shows how to track revisions while editing a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Editing a document usually does not count as a revision until we begin tracking them.
    builder->Write(u"Hello world! ");
    
    ASSERT_EQ(0, doc->get_Revisions()->get_Count());
    ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->get_IsInsertRevision());
    
    doc->StartTrackRevisions(u"John Doe");
    
    builder->Write(u"Hello again! ");
    
    ASSERT_EQ(1, doc->get_Revisions()->get_Count());
    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(1)->get_IsInsertRevision());
    ASSERT_EQ(u"John Doe", doc->get_Revisions()->idx_get(0)->get_Author());
    ASSERT_TRUE((System::DateTime::get_Now() - doc->get_Revisions()->idx_get(0)->get_DateTime()).get_Milliseconds() <= 10);
    
    // Stop tracking revisions to not count any future edits as revisions.
    doc->StopTrackRevisions();
    builder->Write(u"Hello again! ");
    
    ASSERT_EQ(1, doc->get_Revisions()->get_Count());
    ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(2)->get_IsInsertRevision());
    
    // Creating revisions gives them a date and time of the operation.
    // We can disable this by passing DateTime.MinValue when we start tracking revisions.
    doc->StartTrackRevisions(u"John Doe", System::DateTime::MinValue);
    builder->Write(u"Hello again! ");
    
    ASSERT_EQ(2, doc->get_Revisions()->get_Count());
    ASSERT_EQ(u"John Doe", doc->get_Revisions()->idx_get(1)->get_Author());
    ASSERT_EQ(System::DateTime::MinValue, doc->get_Revisions()->idx_get(1)->get_DateTime());
    
    // We can accept/reject these revisions programmatically
    // by calling methods such as Document.AcceptAllRevisions, or each revision's Accept method.
    // In Microsoft Word, we can process them manually via "Review" -> "Changes".
    doc->Save(get_ArtifactsDir() + u"Revision.StartTrackRevisions.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, TrackRevisions)
{
    s_instance->TrackRevisions();
}

} // namespace gtest_test

void ExRevision::AcceptAllRevisions()
{
    //ExStart
    //ExFor:Document.AcceptAllRevisions
    //ExSummary:Shows how to accept all tracking changes in the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Edit the document while tracking changes to create a few revisions.
    doc->StartTrackRevisions(u"John Doe");
    builder->Write(u"Hello world! ");
    builder->Write(u"Hello again! ");
    builder->Write(u"This is another revision.");
    doc->StopTrackRevisions();
    
    ASSERT_EQ(3, doc->get_Revisions()->get_Count());
    
    // We can iterate through every revision and accept/reject it as a part of our document.
    // If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
    doc->AcceptAllRevisions();
    
    ASSERT_EQ(0, doc->get_Revisions()->get_Count());
    ASSERT_EQ(u"Hello world! Hello again! This is another revision.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, AcceptAllRevisions)
{
    s_instance->AcceptAllRevisions();
}

} // namespace gtest_test

void ExRevision::GetRevisedPropertiesOfList()
{
    //ExStart
    //ExFor:RevisionsView
    //ExFor:Document.RevisionsView
    //ExSummary:Shows how to switch between the revised and the original view of a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revisions at list levels.docx");
    doc->UpdateListLabels();
    
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    ASSERT_EQ(u"1.", paragraphs->idx_get(0)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"a.", paragraphs->idx_get(1)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(System::String::Empty, paragraphs->idx_get(2)->get_ListLabel()->get_LabelString());
    
    // View the document object as if all the revisions are accepted. Currently supports list labels.
    doc->set_RevisionsView(Aspose::Words::RevisionsView::Final);
    
    ASSERT_EQ(System::String::Empty, paragraphs->idx_get(0)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"1.", paragraphs->idx_get(1)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"a.", paragraphs->idx_get(2)->get_ListLabel()->get_LabelString());
    //ExEnd
    
    doc->set_RevisionsView(Aspose::Words::RevisionsView::Original);
    doc->AcceptAllRevisions();
    
    ASSERT_EQ(u"a.", paragraphs->idx_get(0)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(System::String::Empty, paragraphs->idx_get(1)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"b.", paragraphs->idx_get(2)->get_ListLabel()->get_LabelString());
}

namespace gtest_test
{

TEST_F(ExRevision, GetRevisedPropertiesOfList)
{
    s_instance->GetRevisedPropertiesOfList();
}

} // namespace gtest_test

void ExRevision::Compare()
{
    //ExStart
    //ExFor:Document.Compare(Document, String, DateTime)
    //ExFor:RevisionCollection.AcceptAll
    //ExSummary:Shows how to compare documents.
    auto docOriginal = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(docOriginal);
    builder->Writeln(u"This is the original document.");
    
    auto docEdited = System::MakeObject<Aspose::Words::Document>();
    builder = System::MakeObject<Aspose::Words::DocumentBuilder>(docEdited);
    builder->Writeln(u"This is the edited document.");
    
    // Comparing documents with revisions will throw an exception.
    if (docOriginal->get_Revisions()->get_Count() == 0 && docEdited->get_Revisions()->get_Count() == 0)
    {
        docOriginal->Compare(docEdited, u"authorName", System::DateTime::get_Now());
    }
    
    // After the comparison, the original document will gain a new revision
    // for every element that is different in the edited document.
    ASSERT_EQ(2, docOriginal->get_Revisions()->get_Count());
    //ExSkip
    for (auto&& r : System::IterateOver(docOriginal->get_Revisions()))
    {
        std::cout << System::String::Format(u"Revision type: {0}, on a node of type \"{1}\"", r->get_RevisionType(), r->get_ParentNode()->get_NodeType()) << std::endl;
        std::cout << System::String::Format(u"\tChanged text: \"{0}\"", r->get_ParentNode()->GetText()) << std::endl;
    }
    
    // Accepting these revisions will transform the original document into the edited document.
    docOriginal->get_Revisions()->AcceptAll();
    
    ASSERT_EQ(docOriginal->GetText(), docEdited->GetText());
    //ExEnd
    
    docOriginal = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(docOriginal);
    ASSERT_EQ(0, docOriginal->get_Revisions()->get_Count());
}

namespace gtest_test
{

TEST_F(ExRevision, Compare)
{
    s_instance->Compare();
}

} // namespace gtest_test

void ExRevision::CompareDocumentWithRevisions()
{
    auto doc1 = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc1);
    builder->Writeln(u"Hello world! This text is not a revision.");
    
    auto docWithRevision = System::MakeObject<Aspose::Words::Document>();
    builder = System::MakeObject<Aspose::Words::DocumentBuilder>(docWithRevision);
    
    docWithRevision->StartTrackRevisions(u"John Doe");
    builder->Writeln(u"This is a revision.");
    
    ASSERT_THROW(static_cast<std::function<void()>>([&docWithRevision, &doc1]() -> void
    {
        docWithRevision->Compare(doc1, u"John Doe", System::DateTime::get_Now());
    })(), System::InvalidOperationException);
}

namespace gtest_test
{

TEST_F(ExRevision, CompareDocumentWithRevisions)
{
    s_instance->CompareDocumentWithRevisions();
}

} // namespace gtest_test

void ExRevision::CompareOptions()
{
    //ExStart
    //ExFor:CompareOptions
    //ExFor:CompareOptions.CompareMoves
    //ExFor:CompareOptions.IgnoreFormatting
    //ExFor:CompareOptions.IgnoreCaseChanges
    //ExFor:CompareOptions.IgnoreComments
    //ExFor:CompareOptions.IgnoreTables
    //ExFor:CompareOptions.IgnoreFields
    //ExFor:CompareOptions.IgnoreFootnotes
    //ExFor:CompareOptions.IgnoreTextboxes
    //ExFor:CompareOptions.IgnoreHeadersAndFooters
    //ExFor:CompareOptions.Target
    //ExFor:ComparisonTargetType
    //ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
    //ExSummary:Shows how to filter specific types of document elements when making a comparison.
    // Create the original document and populate it with various kinds of elements.
    auto docOriginal = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(docOriginal);
    
    // Paragraph text referenced with an endnote:
    builder->Writeln(u"Hello world! This is the first paragraph.");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Original endnote text.");
    
    // Table:
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Original cell 1 text");
    builder->InsertCell();
    builder->Write(u"Original cell 2 text");
    builder->EndTable();
    
    // Textbox:
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 150, 20);
    builder->MoveTo(textBox->get_FirstParagraph());
    builder->Write(u"Original textbox contents");
    
    // DATE field:
    builder->MoveTo(docOriginal->get_FirstSection()->get_Body()->AppendParagraph(u""));
    builder->InsertField(u" DATE ");
    
    // Comment:
    auto newComment = System::MakeObject<Aspose::Words::Comment>(docOriginal, u"John Doe", u"J.D.", System::DateTime::get_Now());
    newComment->SetText(u"Original comment.");
    builder->get_CurrentParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(newComment);
    
    // Header:
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"Original header contents.");
    
    // Create a clone of our document and perform a quick edit on each of the cloned document's elements.
    auto docEdited = System::ExplicitCast<Aspose::Words::Document>(System::ExplicitCast<Aspose::Words::Node>(docOriginal)->Clone(true));
    System::SharedPtr<Aspose::Words::Paragraph> firstParagraph = docEdited->get_FirstSection()->get_Body()->get_FirstParagraph();
    
    firstParagraph->get_Runs()->idx_get(0)->set_Text(u"hello world! this is the first paragraph, after editing.");
    firstParagraph->get_ParagraphFormat()->set_Style(docEdited->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Heading1));
    (System::ExplicitCast<Aspose::Words::Notes::Footnote>(docEdited->GetChild(Aspose::Words::NodeType::Footnote, 0, true)))->get_FirstParagraph()->get_Runs()->idx_get(1)->set_Text(u"Edited endnote text.");
    (System::ExplicitCast<Aspose::Words::Tables::Table>(docEdited->GetChild(Aspose::Words::NodeType::Table, 0, true)))->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited Cell 2 contents");
    (System::ExplicitCast<Aspose::Words::Drawing::Shape>(docEdited->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited textbox contents");
    (System::ExplicitCast<Aspose::Words::Fields::FieldDate>(docEdited->get_Range()->get_Fields()->idx_get(0)))->set_UseLunarCalendar(true);
    (System::ExplicitCast<Aspose::Words::Comment>(docEdited->GetChild(Aspose::Words::NodeType::Comment, 0, true)))->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited comment.");
    docEdited->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited header contents.");
    
    // Comparing documents creates a revision for every edit in the edited document.
    // A CompareOptions object has a series of flags that can suppress revisions
    // on each respective type of element, effectively ignoring their change.
    auto compareOptions = System::MakeObject<Aspose::Words::Comparing::CompareOptions>();
    compareOptions->set_CompareMoves(false);
    compareOptions->set_IgnoreFormatting(false);
    compareOptions->set_IgnoreCaseChanges(false);
    compareOptions->set_IgnoreComments(false);
    compareOptions->set_IgnoreTables(false);
    compareOptions->set_IgnoreFields(false);
    compareOptions->set_IgnoreFootnotes(false);
    compareOptions->set_IgnoreTextboxes(false);
    compareOptions->set_IgnoreHeadersAndFooters(false);
    compareOptions->set_Target(Aspose::Words::Comparing::ComparisonTargetType::New);
    
    docOriginal->Compare(docEdited, u"John Doe", System::DateTime::get_Now(), compareOptions);
    docOriginal->Save(get_ArtifactsDir() + u"Revision.CompareOptions.docx");
    //ExEnd
    
    docOriginal = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Revision.CompareOptions.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"OriginalEdited endnote text.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(docOriginal->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
}

namespace gtest_test
{

TEST_F(ExRevision, CompareOptions)
{
    s_instance->CompareOptions();
}

} // namespace gtest_test

void ExRevision::IgnoreDmlUniqueId(bool isIgnoreDmlUniqueId)
{
    //ExStart
    //ExFor:CompareOptions.AdvancedOptions
    //ExFor:AdvancedCompareOptions.IgnoreDmlUniqueId
    //ExFor:CompareOptions.IgnoreDmlUniqueId
    //ExSummary:Shows how to compare documents ignoring DML unique ID.
    auto docA = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DML unique ID original.docx");
    auto docB = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DML unique ID compare.docx");
    
    // By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
    // If we are ignoring DML's unique ID, and revisions count were 0.
    auto compareOptions = System::MakeObject<Aspose::Words::Comparing::CompareOptions>();
    compareOptions->get_AdvancedOptions()->set_IgnoreDmlUniqueId(isIgnoreDmlUniqueId);
    
    docA->Compare(docB, u"Aspose.Words", System::DateTime::get_Now(), compareOptions);
    
    ASSERT_EQ(isIgnoreDmlUniqueId ? 0 : 2, docA->get_Revisions()->get_Count());
    //ExEnd
}

namespace gtest_test
{

using ExRevision_IgnoreDmlUniqueId_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRevision::IgnoreDmlUniqueId)>::type;

struct ExRevision_IgnoreDmlUniqueId : public ExRevision, public Aspose::Words::ApiExamples::ExRevision, public ::testing::WithParamInterface<ExRevision_IgnoreDmlUniqueId_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRevision_IgnoreDmlUniqueId, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreDmlUniqueId(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRevision_IgnoreDmlUniqueId, ::testing::ValuesIn(ExRevision_IgnoreDmlUniqueId::TestCases()));

} // namespace gtest_test

void ExRevision::LayoutOptionsRevisions()
{
    //ExStart
    //ExFor:Document.LayoutOptions
    //ExFor:LayoutOptions
    //ExFor:LayoutOptions.RevisionOptions
    //ExFor:RevisionColor
    //ExFor:RevisionOptions
    //ExFor:RevisionOptions.InsertedTextColor
    //ExFor:RevisionOptions.ShowRevisionBars
    //ExFor:RevisionOptions.RevisionBarsPosition
    //ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a revision, then change the color of all revisions to green.
    builder->Writeln(u"This is not a revision.");
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
    ASSERT_EQ(Aspose::Words::Layout::RevisionColor::ByAuthor, doc->get_LayoutOptions()->get_RevisionOptions()->get_InsertedTextColor());
    //ExSkip
    ASSERT_TRUE(doc->get_LayoutOptions()->get_RevisionOptions()->get_ShowRevisionBars());
    //ExSkip
    builder->Writeln(u"This is a revision.");
    doc->StopTrackRevisions();
    builder->Writeln(u"This is not a revision.");
    
    // Remove the bar that appears to the left of every revised line.
    doc->get_LayoutOptions()->get_RevisionOptions()->set_InsertedTextColor(Aspose::Words::Layout::RevisionColor::BrightGreen);
    doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowRevisionBars(false);
    doc->get_LayoutOptions()->get_RevisionOptions()->set_RevisionBarsPosition(Aspose::Words::Drawing::HorizontalAlignment::Right);
    
    doc->Save(get_ArtifactsDir() + u"Revision.LayoutOptionsRevisions.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRevision, LayoutOptionsRevisions)
{
    s_instance->LayoutOptionsRevisions();
}

} // namespace gtest_test

void ExRevision::GranularityCompareOption(Aspose::Words::Comparing::Granularity granularity)
{
    //ExStart
    //ExFor:CompareOptions.Granularity
    //ExFor:Granularity
    //ExSummary:Shows to specify a granularity while comparing documents.
    auto docA = System::MakeObject<Aspose::Words::Document>();
    auto builderA = System::MakeObject<Aspose::Words::DocumentBuilder>(docA);
    builderA->Writeln(u"Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");
    
    auto docB = System::MakeObject<Aspose::Words::Document>();
    auto builderB = System::MakeObject<Aspose::Words::DocumentBuilder>(docB);
    builderB->Writeln(u"Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");
    
    // Specify whether changes are tracking
    // by character ('Granularity.CharLevel'), or by word ('Granularity.WordLevel').
    auto compareOptions = System::MakeObject<Aspose::Words::Comparing::CompareOptions>();
    compareOptions->set_Granularity(granularity);
    
    docA->Compare(docB, u"author", System::DateTime::get_Now(), compareOptions);
    
    // The first document's collection of revision groups contains all the differences between documents.
    System::SharedPtr<Aspose::Words::RevisionGroupCollection> groups = docA->get_Revisions()->get_Groups();
    ASSERT_EQ(5, groups->get_Count());
    //ExEnd
    
    if (granularity == Aspose::Words::Comparing::Granularity::CharLevel)
    {
        ASSERT_EQ(Aspose::Words::RevisionType::Deletion, groups->idx_get(0)->get_RevisionType());
        ASSERT_EQ(u"Alpha ", groups->idx_get(0)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Deletion, groups->idx_get(1)->get_RevisionType());
        ASSERT_EQ(u",", groups->idx_get(1)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Insertion, groups->idx_get(2)->get_RevisionType());
        ASSERT_EQ(u"s", groups->idx_get(2)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Insertion, groups->idx_get(3)->get_RevisionType());
        ASSERT_EQ(u"- \"", groups->idx_get(3)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Insertion, groups->idx_get(4)->get_RevisionType());
        ASSERT_EQ(u"\"", groups->idx_get(4)->get_Text());
    }
    else
    {
        ASSERT_EQ(Aspose::Words::RevisionType::Deletion, groups->idx_get(0)->get_RevisionType());
        ASSERT_EQ(u"Alpha Lorem", groups->idx_get(0)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Deletion, groups->idx_get(1)->get_RevisionType());
        ASSERT_EQ(u",", groups->idx_get(1)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Insertion, groups->idx_get(2)->get_RevisionType());
        ASSERT_EQ(u"Lorems", groups->idx_get(2)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Insertion, groups->idx_get(3)->get_RevisionType());
        ASSERT_EQ(u"- \"", groups->idx_get(3)->get_Text());
        
        ASSERT_EQ(Aspose::Words::RevisionType::Insertion, groups->idx_get(4)->get_RevisionType());
        ASSERT_EQ(u"\"", groups->idx_get(4)->get_Text());
    }
}

namespace gtest_test
{

using ExRevision_GranularityCompareOption_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRevision::GranularityCompareOption)>::type;

struct ExRevision_GranularityCompareOption : public ExRevision, public Aspose::Words::ApiExamples::ExRevision, public ::testing::WithParamInterface<ExRevision_GranularityCompareOption_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Comparing::Granularity::CharLevel),
            std::make_tuple(Aspose::Words::Comparing::Granularity::WordLevel),
        };
    }
};

TEST_P(ExRevision_GranularityCompareOption, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->GranularityCompareOption(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRevision_GranularityCompareOption, ::testing::ValuesIn(ExRevision_GranularityCompareOption::TestCases()));

} // namespace gtest_test

void ExRevision::IgnoreStoreItemId()
{
    //ExStart:IgnoreStoreItemId
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:AdvancedCompareOptions
    //ExFor:AdvancedCompareOptions.IgnoreStoreItemId
    //ExSummary:Shows how to compare SDT with same content but different store item id.
    auto docA = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document with SDT 1.docx");
    auto docB = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document with SDT 2.docx");
    
    // Configure options to compare SDT with same content but different store item id.
    auto compareOptions = System::MakeObject<Aspose::Words::Comparing::CompareOptions>();
    compareOptions->get_AdvancedOptions()->set_IgnoreStoreItemId(false);
    
    docA->Compare(docB, u"user", System::DateTime::get_Now(), compareOptions);
    ASSERT_EQ(8, docA->get_Revisions()->get_Count());
    
    compareOptions->get_AdvancedOptions()->set_IgnoreStoreItemId(true);
    
    docA->get_Revisions()->RejectAll();
    docA->Compare(docB, u"user", System::DateTime::get_Now(), compareOptions);
    ASSERT_EQ(0, docA->get_Revisions()->get_Count());
    //ExEnd:IgnoreStoreItemId
}

namespace gtest_test
{

TEST_F(ExRevision, IgnoreStoreItemId)
{
    s_instance->IgnoreStoreItemId();
}

} // namespace gtest_test

void ExRevision::RevisionCellColor()
{
    //ExStart:RevisionCellColor
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:RevisionOptions.InsertCellColor
    //ExFor:RevisionOptions.DeleteCellColor
    //ExSummary:Shows how to work with insert/delete cell revision color.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Cell revisions.docx");
    
    doc->get_LayoutOptions()->get_RevisionOptions()->set_InsertCellColor(Aspose::Words::Layout::RevisionColor::LightBlue);
    doc->get_LayoutOptions()->get_RevisionOptions()->set_DeleteCellColor(Aspose::Words::Layout::RevisionColor::DarkRed);
    
    doc->Save(get_ArtifactsDir() + u"Revision.RevisionCellColor.pdf");
    //ExEnd:RevisionCellColor
}

namespace gtest_test
{

TEST_F(ExRevision, RevisionCellColor)
{
    s_instance->RevisionCellColor();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
