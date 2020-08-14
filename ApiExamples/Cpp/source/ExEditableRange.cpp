// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExEditableRange.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/console.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditorType.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeStart.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRange.h>
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

RTTI_INFO_IMPL_HASH(1663377590u, ::ApiExamples::ExEditableRange::EditableRangeInfoPrinter, ThisTypeBaseTypesInfo);

ExEditableRange::EditableRangeInfoPrinter::EditableRangeInfoPrinter() : mInsideEditableRange(false)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
}

String ExEditableRange::EditableRangeInfoPrinter::ToText()
{
    return mBuilder->ToString();
}

void ExEditableRange::EditableRangeInfoPrinter::Reset()
{
    mBuilder->Clear();
    mInsideEditableRange = false;
}

Aspose::Words::VisitorAction ExEditableRange::EditableRangeInfoPrinter::VisitEditableRangeStart(SharedPtr<Aspose::Words::EditableRangeStart> editableRangeStart)
{
    mBuilder->AppendLine(u" -- Editable range found! -- ");
    mBuilder->AppendLine(String(u"\tID: ") + editableRangeStart->get_Id());
    mBuilder->AppendLine(String(u"\tUser: ") + editableRangeStart->get_EditableRange()->get_SingleUser());
    mBuilder->AppendLine(u"\tContents: ");

    mInsideEditableRange = true;

    // Let the visitor continue visiting other nodes
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExEditableRange::EditableRangeInfoPrinter::VisitEditableRangeEnd(SharedPtr<Aspose::Words::EditableRangeEnd> editableRangeEnd)
{
    mBuilder->AppendLine(u" -- End of editable range -- ");

    mInsideEditableRange = false;

    // Let the visitor continue visiting other nodes
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExEditableRange::EditableRangeInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mInsideEditableRange)
    {
        mBuilder->AppendLine(String(u"\t\"") + run->get_Text() + u"\"");
    }

    // Let the visitor continue visiting other nodes
    return Aspose::Words::VisitorAction::Continue;
}

System::Object::shared_members_type ApiExamples::ExEditableRange::EditableRangeInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExEditableRange::EditableRangeInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

void ExEditableRange::TestCreateEditableRanges(SharedPtr<Aspose::Words::Document> doc, SharedPtr<ExEditableRange::EditableRangeInfoPrinter> visitor)
{
    SharedPtr<NodeCollection> editableRangeStarts = doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeStart, true);

    ASSERT_EQ(2, editableRangeStarts->get_Count());
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeEnd, true)->get_Count());

    SharedPtr<EditableRange> range = (System::DynamicCast<Aspose::Words::EditableRangeStart>(editableRangeStarts->idx_get(0)))->get_EditableRange();

    ASSERT_EQ(0, range->get_Id());
    ASSERT_EQ(u"john.doe@myoffice.com", range->get_SingleUser());
    ASSERT_EQ(Aspose::Words::EditorType::Unspecified, range->get_EditorGroup());

    range = (System::DynamicCast<Aspose::Words::EditableRangeStart>(editableRangeStarts->idx_get(1)))->get_EditableRange();

    ASSERT_EQ(1, range->get_Id());
    ASSERT_EQ(u"jane.doe@myoffice.com", range->get_SingleUser());
    ASSERT_EQ(Aspose::Words::EditorType::Unspecified, range->get_EditorGroup());

    String visitorText = visitor->ToText();

    ASSERT_TRUE(visitorText.Contains(u"Paragraph inside first editable range"));
    ASSERT_TRUE(visitorText.Contains(u"Paragraph inside second editable range"));
}

namespace gtest_test
{

class ExEditableRange : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExEditableRange> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExEditableRange>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExEditableRange> ExEditableRange::s_instance;

} // namespace gtest_test

void ExEditableRange::RemovesEditableRange()
{
    //ExStart
    //ExFor:EditableRange.Remove
    //ExSummary:Shows how to remove an editable range from a document.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create an EditableRange so we can remove it. Does not have to be well-formed
    SharedPtr<EditableRangeStart> edRange1Start = builder->StartEditableRange();
    SharedPtr<EditableRange> editableRange1 = edRange1Start->get_EditableRange();
    builder->Writeln(u"Paragraph inside editable range");
    SharedPtr<EditableRangeEnd> edRange1End = builder->EndEditableRange();
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeStart, true)->get_Count());
    //ExSkip
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeEnd, true)->get_Count());
    //ExSkip
    ASSERT_EQ(0, edRange1Start->get_EditableRange()->get_Id());
    //ExSkip
    ASSERT_EQ(u"", edRange1Start->get_EditableRange()->get_SingleUser());
    //ExSkip

    // Remove the range that was just made
    editableRange1->Remove();
    //ExEnd

    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeStart, true)->get_Count());
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeEnd, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExEditableRange, RemovesEditableRange)
{
    s_instance->RemovesEditableRange();
}

} // namespace gtest_test

void ExEditableRange::CreateEditableRanges()
{
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Start an editable range
    SharedPtr<EditableRangeStart> edRange1Start = builder->StartEditableRange();

    // An EditableRange object is created for the EditableRangeStart that we just made
    SharedPtr<EditableRange> editableRange1 = edRange1Start->get_EditableRange();

    // Put something inside the editable range
    builder->Writeln(u"Paragraph inside first editable range");

    // An editable range is well-formed if it has a start and an end
    // Multiple editable ranges can be nested and overlapping
    SharedPtr<EditableRangeEnd> edRange1End = builder->EndEditableRange();

    // Explicitly state which EditableRangeStart a new EditableRangeEnd should be paired with
    SharedPtr<EditableRangeStart> edRange2Start = builder->StartEditableRange();
    builder->Writeln(u"Paragraph inside second editable range");
    SharedPtr<EditableRange> editableRange2 = edRange2Start->get_EditableRange();
    SharedPtr<EditableRangeEnd> edRange2End = builder->EndEditableRange(edRange2Start);

    // Editable range starts and ends have their own respective node types
    ASSERT_EQ(Aspose::Words::NodeType::EditableRangeStart, edRange1Start->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::EditableRangeEnd, edRange1End->get_NodeType());

    // Editable range IDs are unique and set automatically
    ASSERT_EQ(0, editableRange1->get_Id());
    ASSERT_EQ(1, editableRange2->get_Id());

    // Editable range starts and ends always belong to a range
    ASPOSE_ASSERT_EQ(edRange1Start, editableRange1->get_EditableRangeStart());
    ASPOSE_ASSERT_EQ(edRange1End, editableRange1->get_EditableRangeEnd());

    // They also inherit the ID of the entire editable range that they belong to
    ASSERT_EQ(editableRange1->get_Id(), edRange1Start->get_Id());
    ASSERT_EQ(editableRange1->get_Id(), edRange1End->get_Id());
    ASSERT_EQ(editableRange2->get_Id(), edRange2Start->get_EditableRange()->get_Id());
    ASSERT_EQ(editableRange2->get_Id(), edRange2End->get_EditableRangeStart()->get_EditableRange()->get_Id());

    // If the editable range was found in a document, it will probably have something in the single user property
    // But if we make one programmatically, the property is empty by default
    ASSERT_EQ(String::Empty, editableRange1->get_SingleUser());

    // We have to set it ourselves if we want the ranges to belong to somebody
    editableRange1->set_SingleUser(u"john.doe@myoffice.com");
    editableRange2->set_SingleUser(u"jane.doe@myoffice.com");

    // Initialize a custom visitor for editable ranges that will print their contents
    auto editableRangeReader = MakeObject<ExEditableRange::EditableRangeInfoPrinter>();

    // Both the start and end of an editable range can accept visitors, but not the editable range itself
    edRange1Start->Accept(editableRangeReader);
    edRange2End->Accept(editableRangeReader);

    // Or, if we want to go over all the editable ranges in a document, we can get the document to accept the visitor
    editableRangeReader->Reset();
    doc->Accept(editableRangeReader);

    System::Console::WriteLine(editableRangeReader->ToText());
    TestCreateEditableRanges(doc, editableRangeReader);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExEditableRange, CreateEditableRanges)
{
    s_instance->CreateEditableRanges();
}

} // namespace gtest_test

void ExEditableRange::IncorrectStructureException()
{
    auto doc = MakeObject<Document>();

    auto builder = MakeObject<DocumentBuilder>(doc);

    // Checking that isn't valid structure for the current document

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&builder]()
    {
        builder->EndEditableRange();
    };

    ASSERT_THROW(_local_func_0(), System::InvalidOperationException);

    builder->StartEditableRange();
}

namespace gtest_test
{

TEST_F(ExEditableRange, IncorrectStructureException)
{
    s_instance->IncorrectStructureException();
}

} // namespace gtest_test

void ExEditableRange::IncorrectStructureDoNotAdded()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();
    auto builder = MakeObject<DocumentBuilder>(doc);

    //ExStart
    //ExFor:EditableRange.EditorGroup
    //ExFor:EditorType
    //ExSummary:Shows how to add editing group for editable ranges
    SharedPtr<EditableRangeStart> startRange1 = builder->StartEditableRange();

    builder->Writeln(u"EditableRange_1_1");
    builder->Writeln(u"EditableRange_1_2");

    // Sets the editor for editable range region
    startRange1->get_EditableRange()->set_EditorGroup(Aspose::Words::EditorType::Everyone);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);

    // Assert that it's not valid structure and editable ranges aren't added to the current document
    SharedPtr<NodeCollection> startNodes = doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeStart, true);
    ASSERT_EQ(0, startNodes->get_Count());

    SharedPtr<NodeCollection> endNodes = doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeEnd, true);
    ASSERT_EQ(0, endNodes->get_Count());
}

namespace gtest_test
{

TEST_F(ExEditableRange, IncorrectStructureDoNotAdded)
{
    s_instance->IncorrectStructureDoNotAdded();
}

} // namespace gtest_test

} // namespace ApiExamples
