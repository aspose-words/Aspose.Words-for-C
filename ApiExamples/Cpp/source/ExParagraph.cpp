// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExParagraph.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/console.h>
#include <system/collections/ienumerable.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/TabStopCollection.h>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/FrameFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionType.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/Revision.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "TestUtil.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Drawing;
namespace ApiExamples {

void ExParagraph::InsertFieldUsingFieldType(SharedPtr<Aspose::Words::Document> doc, Aspose::Words::Fields::FieldType fieldType, bool updateField, SharedPtr<Aspose::Words::Node> refNode, bool isAfter, int paraIndex)
{
    SharedPtr<Paragraph> para = DocumentHelper::GetParagraph(doc, paraIndex);
    para->InsertField(fieldType, updateField, refNode, isAfter);
}

void ExParagraph::InsertFieldUsingFieldCode(SharedPtr<Aspose::Words::Document> doc, String fieldCode, SharedPtr<Aspose::Words::Node> refNode, bool isAfter, int paraIndex)
{
    SharedPtr<Paragraph> para = DocumentHelper::GetParagraph(doc, paraIndex);
    para->InsertField(fieldCode, refNode, isAfter);
}

void ExParagraph::InsertFieldUsingFieldCodeFieldString(SharedPtr<Aspose::Words::Document> doc, String fieldCode, String fieldValue, SharedPtr<Aspose::Words::Node> refNode, bool isAfter, int paraIndex)
{
    SharedPtr<Paragraph> para = DocumentHelper::GetParagraph(doc, paraIndex);
    para->InsertField(fieldCode, fieldValue, refNode, isAfter);
}

namespace gtest_test
{

class ExParagraph : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExParagraph> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExParagraph>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExParagraph> ExParagraph::s_instance;

} // namespace gtest_test

void ExParagraph::DocumentBuilderInsertParagraph()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertParagraph
    //ExFor:ParagraphFormat.FirstLineIndent
    //ExFor:ParagraphFormat.Alignment
    //ExFor:ParagraphFormat.KeepTogether
    //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndAlpha
    //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndDigit
    //ExFor:Paragraph.IsEndOfDocument
    //ExSummary:Shows how to insert a paragraph into the document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Specify font formatting
    SharedPtr<Font> font = builder->get_Font();
    font->set_Size(16);
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_Blue());
    font->set_Name(u"Arial");
    font->set_Underline(Aspose::Words::Underline::Dash);

    // Specify paragraph formatting
    SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
    paragraphFormat->set_FirstLineIndent(8);
    paragraphFormat->set_Alignment(Aspose::Words::ParagraphAlignment::Justify);
    paragraphFormat->set_AddSpaceBetweenFarEastAndAlpha(true);
    paragraphFormat->set_AddSpaceBetweenFarEastAndDigit(true);
    paragraphFormat->set_KeepTogether(true);

    // Using Writeln() ends the paragraph after writing and makes a new one, while Write() stays on the same paragraph
    builder->Writeln(u"A whole paragraph.");

    // We can use this flag to ensure that we're at the end of the document
    ASSERT_TRUE(builder->get_CurrentParagraph()->get_IsEndOfDocument());
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

    ASPOSE_ASSERT_EQ(8, paragraph->get_ParagraphFormat()->get_FirstLineIndent());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Justify, paragraph->get_ParagraphFormat()->get_Alignment());
    ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_AddSpaceBetweenFarEastAndAlpha());
    ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_AddSpaceBetweenFarEastAndDigit());
    ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_KeepTogether());
    ASSERT_EQ(u"A whole paragraph.", paragraph->GetText().Trim());

    SharedPtr<Font> runFont = paragraph->get_Runs()->idx_get(0)->get_Font();

    ASPOSE_ASSERT_EQ(16.0, runFont->get_Size());
    ASSERT_TRUE(runFont->get_Bold());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), runFont->get_Color().ToArgb());
    ASSERT_EQ(u"Arial", runFont->get_Name());
    ASSERT_EQ(Aspose::Words::Underline::Dash, runFont->get_Underline());
}

namespace gtest_test
{

TEST_F(ExParagraph, DocumentBuilderInsertParagraph)
{
    s_instance->DocumentBuilderInsertParagraph();
}

} // namespace gtest_test

void ExParagraph::InsertField()
{
    //ExStart
    //ExFor:Paragraph.AppendField(FieldType, Boolean)
    //ExFor:Paragraph.AppendField(String)
    //ExFor:Paragraph.AppendField(String, String)
    //ExFor:Paragraph.InsertField(string, Node, bool)
    //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
    //ExFor:Paragraph.InsertField(string, string, Node, bool)
    //ExSummary:Shows how to insert fields in different ways.
    // Create a blank document and get its first paragraph
    auto doc = MakeObject<Document>();
    SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

    // Choose a DATE field by FieldType, append it to the end of the paragraph and update it
    para->AppendField(Aspose::Words::Fields::FieldType::FieldDate, true);

    // Append a TIME field using a field code
    para->AppendField(u" TIME  \\@ \"HH:mm:ss\" ");

    // Append a QUOTE field that will display a placeholder value until it is updated manually in Microsoft Word
    // or programmatically with Document.UpdateFields() or Field.Update()
    para->AppendField(u" QUOTE \"Real value\"", u"Placeholder value");

    // We can choose a node in the paragraph and insert a field
    // before or after that node instead of appending it to the end of a paragraph
    para = doc->get_FirstSection()->get_Body()->AppendParagraph(u"");
    auto run = MakeObject<Run>(doc);
    run->set_Text(u" My Run. ");
    para->AppendChild(run);

    // Insert an AUTHOR field into the paragraph and place it before the run we created
    doc->get_BuiltInDocumentProperties()->idx_get(u"Author")->set_Value(System::ObjectExt::Box<String>(u"John Doe"));
    para->InsertField(Aspose::Words::Fields::FieldType::FieldAuthor, true, run, false);

    // Insert another field designated by field code before the run
    para->InsertField(u" QUOTE \"Real value\" ", run, false);

    // Insert another field with a place holder value and place it after the run
    para->InsertField(u" QUOTE \"Real value\"", u" Placeholder value. ", run, true);

    doc->Save(ArtifactsDir + u"Paragraph.InsertField.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Paragraph.InsertField.docx");

    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldDate, u" DATE ", System::DateTime::get_Now(), doc->get_Range()->get_Fields()->idx_get(0), System::TimeSpan(0, 0, 0, 0));
    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldTime, u" TIME  \\@ \"HH:mm:ss\" ", System::DateTime::get_Now(), doc->get_Range()->get_Fields()->idx_get(1), System::TimeSpan(0, 0, 0, 5));
    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldQuote, u" QUOTE \"Real value\"", u"Placeholder value", doc->get_Range()->get_Fields()->idx_get(2));
    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldAuthor, u" AUTHOR ", u"John Doe", doc->get_Range()->get_Fields()->idx_get(3));
    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldQuote, u" QUOTE \"Real value\" ", u"Real value", doc->get_Range()->get_Fields()->idx_get(4));
    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldQuote, u" QUOTE \"Real value\"", u" Placeholder value. ", doc->get_Range()->get_Fields()->idx_get(5));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertField)
{
    s_instance->InsertField();
}

} // namespace gtest_test

void ExParagraph::InsertFieldBeforeTextInParagraph()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    InsertFieldUsingFieldCode(doc, u" AUTHOR ", nullptr, false, 1);

    ASSERT_EQ(u"\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldBeforeTextInParagraph)
{
    s_instance->InsertFieldBeforeTextInParagraph();
}

} // namespace gtest_test

void ExParagraph::InsertFieldAfterTextInParagraph()
{
    String date = System::DateTime::get_Today().ToString(u"d");

    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    InsertFieldUsingFieldCode(doc, u" DATE ", nullptr, true, 1);

    ASSERT_EQ(String::Format(u"Hello World!\u0013 DATE \u0014{0}\u0015\r",date), DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldAfterTextInParagraph)
{
    s_instance->InsertFieldAfterTextInParagraph();
}

} // namespace gtest_test

void ExParagraph::InsertFieldBeforeTextInParagraphWithoutUpdateField()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    InsertFieldUsingFieldType(doc, Aspose::Words::Fields::FieldType::FieldAuthor, false, nullptr, false, 1);

    ASSERT_EQ(u"\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldBeforeTextInParagraphWithoutUpdateField)
{
    s_instance->InsertFieldBeforeTextInParagraphWithoutUpdateField();
}

} // namespace gtest_test

void ExParagraph::InsertFieldAfterTextInParagraphWithoutUpdateField()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    InsertFieldUsingFieldType(doc, Aspose::Words::Fields::FieldType::FieldAuthor, false, nullptr, true, 1);

    ASSERT_EQ(u"Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldAfterTextInParagraphWithoutUpdateField)
{
    s_instance->InsertFieldAfterTextInParagraphWithoutUpdateField();
}

} // namespace gtest_test

void ExParagraph::InsertFieldWithoutSeparator()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    InsertFieldUsingFieldType(doc, Aspose::Words::Fields::FieldType::FieldListNum, true, nullptr, false, 1);

    ASSERT_EQ(u"\u0013 LISTNUM \u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldWithoutSeparator)
{
    s_instance->InsertFieldWithoutSeparator();
}

} // namespace gtest_test

void ExParagraph::InsertFieldBeforeParagraphWithoutDocumentAuthor()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();
    doc->get_BuiltInDocumentProperties()->set_Author(u"");

    InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", nullptr, nullptr, false, 1);

    ASSERT_EQ(u"\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldBeforeParagraphWithoutDocumentAuthor)
{
    s_instance->InsertFieldBeforeParagraphWithoutDocumentAuthor();
}

} // namespace gtest_test

void ExParagraph::InsertFieldAfterParagraphWithoutChangingDocumentAuthor()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", nullptr, nullptr, true, 1);

    ASSERT_EQ(u"Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldAfterParagraphWithoutChangingDocumentAuthor)
{
    s_instance->InsertFieldAfterParagraphWithoutChangingDocumentAuthor();
}

} // namespace gtest_test

void ExParagraph::InsertFieldBeforeRunText()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    //Add some text into the paragraph
    SharedPtr<Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 1);

    InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", u"Test Field Value", run, false, 1);

    ASSERT_EQ(u"Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldBeforeRunText)
{
    s_instance->InsertFieldBeforeRunText();
}

} // namespace gtest_test

void ExParagraph::InsertFieldAfterRunText()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    // Add some text into the paragraph
    SharedPtr<Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 1);

    InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", u"", run, true, 1);

    ASSERT_EQ(u"Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldAfterRunText)
{
    s_instance->InsertFieldAfterRunText();
}

} // namespace gtest_test

void ExParagraph::InsertFieldEmptyParagraphWithoutUpdateField()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentWithoutDummyText();

    InsertFieldUsingFieldType(doc, Aspose::Words::Fields::FieldType::FieldAuthor, false, nullptr, false, 1);

    ASSERT_EQ(u"\u0013 AUTHOR \u0014\u0015\f", DocumentHelper::GetParagraphText(doc, 1));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldEmptyParagraphWithoutUpdateField)
{
    s_instance->InsertFieldEmptyParagraphWithoutUpdateField();
}

} // namespace gtest_test

void ExParagraph::InsertFieldEmptyParagraphWithUpdateField()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentWithoutDummyText();

    InsertFieldUsingFieldType(doc, Aspose::Words::Fields::FieldType::FieldAuthor, true, nullptr, false, 0);

    ASSERT_EQ(u"\u0013 AUTHOR \u0014Test Author\u0015\r", DocumentHelper::GetParagraphText(doc, 0));
}

namespace gtest_test
{

TEST_F(ExParagraph, InsertFieldEmptyParagraphWithUpdateField)
{
    s_instance->InsertFieldEmptyParagraphWithUpdateField();
}

} // namespace gtest_test

void ExParagraph::CompositeNodeChildren()
{
    //ExStart
    //ExFor:CompositeNode.Count
    //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
    //ExFor:CompositeNode.InsertAfter(Node, Node)
    //ExFor:CompositeNode.InsertBefore(Node, Node)
    //ExFor:CompositeNode.PrependChild(Node)
    //ExFor:Paragraph.GetText
    //ExFor:Run
    //ExSummary:Shows how to add, update and delete child nodes from a CompositeNode's child collection.
    auto doc = MakeObject<Document>();

    // An empty document has one paragraph by default
    ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());

    // A paragraph is a composite node because it can contain runs, which are another type of node
    SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    auto paragraphText = MakeObject<Run>(doc, u"Initial text. ");
    paragraph->AppendChild(paragraphText);

    // We will place these 3 children into the main text of our paragraph
    auto run1 = MakeObject<Run>(doc, u"Run 1. ");
    auto run2 = MakeObject<Run>(doc, u"Run 2. ");
    auto run3 = MakeObject<Run>(doc, u"Run 3. ");

    // We initialized them but not in our paragraph yet
    ASSERT_EQ(u"Initial text.", paragraph->GetText().Trim());

    // Insert run2 before initial paragraph text. This will be at the start of the paragraph
    paragraph->InsertBefore(run2, paragraphText);

    // Insert run3 after initial paragraph text. This will be at the end of the paragraph
    paragraph->InsertAfter(run3, paragraphText);

    // Insert run1 before every other child node. run2 was the start of the paragraph, now it will be run1
    paragraph->PrependChild(run1);

    ASSERT_EQ(u"Run 1. Run 2. Initial text. Run 3.", paragraph->GetText().Trim());
    ASSERT_EQ(4, paragraph->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // Access the child node collection and update/delete children
    (System::DynamicCast<Aspose::Words::Run>(paragraph->GetChildNodes(Aspose::Words::NodeType::Run, true)->idx_get(1)))->set_Text(u"Updated run 2. ");
    paragraph->GetChildNodes(Aspose::Words::NodeType::Run, true)->Remove(paragraphText);

    ASSERT_EQ(u"Run 1. Updated run 2. Run 3.", paragraph->GetText().Trim());
    ASSERT_EQ(3, paragraph->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExParagraph, CompositeNodeChildren)
{
    s_instance->CompositeNodeChildren();
}

} // namespace gtest_test

void ExParagraph::RevisionHistory()
{
    //ExStart
    //ExFor:Paragraph.IsMoveFromRevision
    //ExFor:Paragraph.IsMoveToRevision
    //ExFor:ParagraphCollection
    //ExFor:ParagraphCollection.Item(Int32)
    //ExFor:Story.Paragraphs
    //ExSummary:Shows how to get paragraph that was moved (deleted/inserted) in Microsoft Word while change tracking was enabled.
    auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

    // There are two sets of move revisions in this document
    // One moves a small part of a paragraph, while the other moves a whole paragraph
    // Paragraph.IsMoveFromRevision/IsMoveToRevision will only be true if a whole paragraph is moved, as in the latter case
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    for (int i = 0; i < paragraphs->get_Count(); i++)
    {
        if (paragraphs->idx_get(i)->get_IsMoveFromRevision())
        {
            System::Console::WriteLine(u"The paragraph {0} has been moved (deleted).", System::ObjectExt::Box<int>(i));
        }
        if (paragraphs->idx_get(i)->get_IsMoveToRevision())
        {
            System::Console::WriteLine(u"The paragraph {0} has been moved (inserted).", System::ObjectExt::Box<int>(i));
        }
    }
    //ExEnd

    ASSERT_EQ(11, doc->get_Revisions()->LINQ_Count());

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> r)> _local_func_0 = [](SharedPtr<Revision> r)
    {
        return r->get_RevisionType() == Aspose::Words::RevisionType::Moving;
    };

    ASSERT_EQ(6, doc->get_Revisions()->LINQ_Count(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_0)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Node> p)> _local_func_1 = [](SharedPtr<Node> p)
    {
        return (System::DynamicCast<Aspose::Words::Paragraph>(p))->get_IsMoveFromRevision();
    };

    ASSERT_EQ(1, paragraphs->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_1)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Node> p)> _local_func_2 = [](SharedPtr<Node> p)
    {
        return (System::DynamicCast<Aspose::Words::Paragraph>(p))->get_IsMoveToRevision();
    };

    ASSERT_EQ(1, paragraphs->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_2)));
}

namespace gtest_test
{

TEST_F(ExParagraph, RevisionHistory)
{
    s_instance->RevisionHistory();
}

} // namespace gtest_test

void ExParagraph::GetFormatRevision()
{
    //ExStart
    //ExFor:Paragraph.IsFormatRevision
    //ExSummary:Shows how to get information about whether this object was formatted in Microsoft Word while change tracking was enabled
    auto doc = MakeObject<Document>(MyDir + u"Format revision.docx");

    // This paragraph's formatting was changed while revisions were being tracked
    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_IsFormatRevision());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExParagraph, GetFormatRevision)
{
    s_instance->GetFormatRevision();
}

} // namespace gtest_test

void ExParagraph::GetFrameProperties()
{
    //ExStart
    //ExFor:Paragraph.FrameFormat
    //ExFor:FrameFormat
    //ExFor:FrameFormat.IsFrame
    //ExFor:FrameFormat.Width
    //ExFor:FrameFormat.Height
    //ExFor:FrameFormat.HeightRule
    //ExFor:FrameFormat.HorizontalAlignment
    //ExFor:FrameFormat.VerticalAlignment
    //ExFor:FrameFormat.HorizontalPosition
    //ExFor:FrameFormat.RelativeHorizontalPosition
    //ExFor:FrameFormat.HorizontalDistanceFromText
    //ExFor:FrameFormat.VerticalPosition
    //ExFor:FrameFormat.RelativeVerticalPosition
    //ExFor:FrameFormat.VerticalDistanceFromText
    //ExSummary:Shows how to get information about formatting properties of paragraphs that are frames.
    auto doc = MakeObject<Document>(MyDir + u"Paragraph frame.docx");

    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    for (auto paragraph : System::IterateOver(paragraphs->LINQ_OfType<SharedPtr<Paragraph>>()->LINQ_Where([](SharedPtr<Paragraph> p) { return p->get_FrameFormat()->get_IsFrame(); })))
    {
        System::Console::WriteLine(String(u"Width: ") + paragraph->get_FrameFormat()->get_Width());
        System::Console::WriteLine(String(u"Height: ") + paragraph->get_FrameFormat()->get_Height());
        System::Console::WriteLine(String(u"HeightRule: ") + System::ObjectExt::ToString(paragraph->get_FrameFormat()->get_HeightRule()));
        System::Console::WriteLine(String(u"HorizontalAlignment: ") + System::ObjectExt::ToString(paragraph->get_FrameFormat()->get_HorizontalAlignment()));
        System::Console::WriteLine(String(u"VerticalAlignment: ") + System::ObjectExt::ToString(paragraph->get_FrameFormat()->get_VerticalAlignment()));
        System::Console::WriteLine(String(u"HorizontalPosition: ") + paragraph->get_FrameFormat()->get_HorizontalPosition());
        System::Console::WriteLine(String(u"RelativeHorizontalPosition: ") + System::ObjectExt::ToString(paragraph->get_FrameFormat()->get_RelativeHorizontalPosition()));
        System::Console::WriteLine(String(u"HorizontalDistanceFromText: ") + paragraph->get_FrameFormat()->get_HorizontalDistanceFromText());
        System::Console::WriteLine(String(u"VerticalPosition: ") + paragraph->get_FrameFormat()->get_VerticalPosition());
        System::Console::WriteLine(String(u"RelativeVerticalPosition: ") + System::ObjectExt::ToString(paragraph->get_FrameFormat()->get_RelativeVerticalPosition()));
        System::Console::WriteLine(String(u"VerticalDistanceFromText: ") + paragraph->get_FrameFormat()->get_VerticalDistanceFromText());
    }
    //ExEnd

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Paragraph> p)> _local_func_4 = [](SharedPtr<Paragraph> p)
    {
        return p->get_FrameFormat()->get_IsFrame();
    };

    for (auto paragraph : System::IterateOver(paragraphs->LINQ_OfType<SharedPtr<Paragraph> >()->LINQ_Where(static_cast<System::Func<SharedPtr<Paragraph>, bool>>(_local_func_4))))
    {
        ASPOSE_ASSERT_EQ(233.3, paragraph->get_FrameFormat()->get_Width());
        ASPOSE_ASSERT_EQ(138.8, paragraph->get_FrameFormat()->get_Height());
        ASPOSE_ASSERT_EQ(34.05, paragraph->get_FrameFormat()->get_HorizontalPosition());
        ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Page, paragraph->get_FrameFormat()->get_RelativeHorizontalPosition());
        ASPOSE_ASSERT_EQ(9, paragraph->get_FrameFormat()->get_HorizontalDistanceFromText());
        ASPOSE_ASSERT_EQ(20.5, paragraph->get_FrameFormat()->get_VerticalPosition());
        ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, paragraph->get_FrameFormat()->get_RelativeVerticalPosition());
        ASPOSE_ASSERT_EQ(0, paragraph->get_FrameFormat()->get_VerticalDistanceFromText());
    }
}

namespace gtest_test
{

TEST_F(ExParagraph, GetFrameProperties)
{
    s_instance->GetFrameProperties();
}

} // namespace gtest_test

void ExParagraph::IsRevision()
{
    //ExStart
    //ExFor:Paragraph.IsDeleteRevision
    //ExFor:Paragraph.IsInsertRevision
    //ExSummary:Shows how to work with revision paragraphs.
    auto doc = MakeObject<Document>();
    SharedPtr<Body> body = doc->get_FirstSection()->get_Body();
    SharedPtr<Paragraph> para = body->get_FirstParagraph();

    // Add text to the first paragraph, then add two more paragraphs
    para->AppendChild(MakeObject<Run>(doc, u"Paragraph 1. "));
    body->AppendParagraph(u"Paragraph 2. ");
    body->AppendParagraph(u"Paragraph 3. ");

    // We have three paragraphs, none of which registered as any type of revision
    // If we add/remove any content in the document while tracking revisions,
    // they will be displayed as such in the document and can be accepted/rejected
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());

    // This paragraph is a revision and will have the according "IsInsertRevision" flag set
    para = body->AppendParagraph(u"Paragraph 4. ");
    ASSERT_TRUE(para->get_IsInsertRevision());

    // Get the document's paragraph collection and remove a paragraph
    SharedPtr<ParagraphCollection> paragraphs = body->get_Paragraphs();
    ASSERT_EQ(4, paragraphs->get_Count());
    para = paragraphs->idx_get(2);
    para->Remove();

    // Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
    // and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions
    ASSERT_EQ(4, paragraphs->get_Count());
    ASSERT_TRUE(para->get_IsDeleteRevision());

    // The delete revision paragraph is removed once we accept changes
    doc->AcceptAllRevisions();
    ASSERT_EQ(3, paragraphs->get_Count());
    ASSERT_TRUE(para->get_Count() == 0);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExParagraph, IsRevision)
{
    s_instance->IsRevision();
}

} // namespace gtest_test

void ExParagraph::BreakIsStyleSeparator()
{
    //ExStart
    //ExFor:Paragraph.BreakIsStyleSeparator
    //ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
    // Create a blank document and insert a table of contents field
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertTableOfContents(u"\\o \\h \\z \\u");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    // Insert a paragraph with a style that will be picked up as an entry in the TOC
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);

    // Both these strings are on the same line and same paragraph and will therefore show up on the same TOC entry
    builder->Write(u"Heading 1. ");
    builder->Write(u"Will appear in the TOC. ");

    // Any text on a new line that does not have a heading style will not register as a TOC entry
    // If we insert a style separator, we can write more text on the same line
    // and use a different style without it showing up in the TOC
    // If we use a heading type style afterwards, we can draw two TOC entries from one line of document text
    builder->InsertStyleSeparator();
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Quote);
    builder->Write(u"Won't appear in the TOC. ");

    // This flag is set to true for such paragraphs
    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_BreakIsStyleSeparator());

    // Update the TOC and save the document
    doc->UpdateFields();
    doc->Save(ArtifactsDir + u"Paragraph.BreakIsStyleSeparator.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Paragraph.BreakIsStyleSeparator.docx");

    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldTOC, u"TOC \\o \\h \\z \\u", u"\u0013 HYPERLINK \\l \"_Toc256000000\" \u0014Heading 1. Will appear in the TOC.\t\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\u0015\r", doc->get_Range()->get_Fields()->idx_get(0));
    ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_BreakIsStyleSeparator());
}

namespace gtest_test
{

TEST_F(ExParagraph, BreakIsStyleSeparator)
{
    s_instance->BreakIsStyleSeparator();
}

} // namespace gtest_test

void ExParagraph::TabStops()
{
    //ExStart
    //ExFor:Paragraph.GetEffectiveTabStops
    //ExSummary:Shows how to set custom tab stops.
    auto doc = MakeObject<Document>();
    SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

    // If there are no tab stops in this collection, while we are in this paragraph
    // the cursor will jump 36 points each time we press the Tab key in Microsoft Word
    ASSERT_EQ(0, doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetEffectiveTabStops()->get_Length());

    // We can add custom tab stops in Microsoft Word if we enable the ruler via the view tab
    // Each unit on that ruler is two default tab stops, which is 72 points
    // Those tab stops can be programmatically added to the paragraph like this
    SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();
    format->get_TabStops()->Add(72, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dots);
    format->get_TabStops()->Add(216, Aspose::Words::TabAlignment::Center, Aspose::Words::TabLeader::Dashes);
    format->get_TabStops()->Add(360, Aspose::Words::TabAlignment::Right, Aspose::Words::TabLeader::Line);

    // These tab stops are added to this collection, and can also be seen by enabling the ruler mentioned above
    ASSERT_EQ(3, para->GetEffectiveTabStops()->get_Length());

    // Add a Run with tab characters that will snap the text to our TabStop positions and save the document
    para->AppendChild(MakeObject<Run>(doc, u"\tTab 1\tTab 2\tTab 3"));
    doc->Save(ArtifactsDir + u"Paragraph.TabStops.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Paragraph.TabStops.docx");
    format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();

    TestUtil::VerifyTabStop(72.0, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dots, false, format->get_TabStops()->idx_get(0));
    TestUtil::VerifyTabStop(216.0, Aspose::Words::TabAlignment::Center, Aspose::Words::TabLeader::Dashes, false, format->get_TabStops()->idx_get(1));
    TestUtil::VerifyTabStop(360.0, Aspose::Words::TabAlignment::Right, Aspose::Words::TabLeader::Line, false, format->get_TabStops()->idx_get(2));
}

namespace gtest_test
{

TEST_F(ExParagraph, TabStops)
{
    s_instance->TabStops();
}

} // namespace gtest_test

void ExParagraph::JoinRuns()
{
    //ExStart
    //ExFor:Paragraph.JoinRunsWithSameFormatting
    //ExSummary:Shows how to simplify paragraphs by merging superfluous runs.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a few small runs into the document
    builder->Write(u"Run 1. ");
    builder->Write(u"Run 2. ");
    builder->Write(u"Run 3. ");
    builder->Write(u"Run 4. ");

    // The Paragraph may look like it's in once piece in Microsoft Word,
    // but it is fragmented into several Runs, which leaves room for optimization
    // A big run may be split into many smaller runs with the same formatting
    // if we keep splitting up a piece of text while manually editing it in Microsoft Word
    SharedPtr<Paragraph> para = builder->get_CurrentParagraph();
    ASSERT_EQ(4, para->get_Runs()->get_Count());

    // Change the style of the last run to something different from the first three
    para->get_Runs()->idx_get(3)->get_Font()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Emphasis);

    // We can run the JoinRunsWithSameFormatting() method to merge similar Runs
    // This method also returns the number of joins that occured during the merge
    // Two merges occured to combine Runs 1-3, while Run 4 was left out because it has an incompatible style
    ASSERT_EQ(2, para->JoinRunsWithSameFormatting());

    // The paragraph has been simplified to two runs
    ASSERT_EQ(2, para->get_Runs()->get_Count());
    ASSERT_EQ(u"Run 1. Run 2. Run 3. ", para->get_Runs()->idx_get(0)->get_Text());
    ASSERT_EQ(u"Run 4. ", para->get_Runs()->idx_get(1)->get_Text());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExParagraph, JoinRuns)
{
    s_instance->JoinRuns();
}

} // namespace gtest_test

} // namespace ApiExamples
