// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExEditableRange.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/exceptions.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditorType.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRange.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"
#include "DocumentHelper.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2432718826u, ::Aspose::Words::ApiExamples::ExEditableRange::EditableRangePrinter, ThisTypeBaseTypesInfo);

ExEditableRange::EditableRangePrinter::EditableRangePrinter() : mInsideEditableRange(false)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
}

System::String ExEditableRange::EditableRangePrinter::ToText()
{
    return mBuilder->ToString();
}

void ExEditableRange::EditableRangePrinter::Reset()
{
    mBuilder->Clear();
    mInsideEditableRange = false;
}

Aspose::Words::VisitorAction ExEditableRange::EditableRangePrinter::VisitEditableRangeStart(System::SharedPtr<Aspose::Words::EditableRangeStart> editableRangeStart)
{
    mBuilder->AppendLine(u" -- Editable range found! -- ");
    mBuilder->AppendLine(System::String(u"\tID:\t\t") + editableRangeStart->get_Id());
    if (editableRangeStart->get_EditableRange()->get_SingleUser() == System::String::Empty)
    {
        mBuilder->AppendLine(System::String(u"\tGroup:\t") + System::ObjectExt::ToString(editableRangeStart->get_EditableRange()->get_EditorGroup()));
    }
    else
    {
        mBuilder->AppendLine(System::String(u"\tUser:\t") + editableRangeStart->get_EditableRange()->get_SingleUser());
    }
    mBuilder->AppendLine(u"\tContents:");
    
    mInsideEditableRange = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExEditableRange::EditableRangePrinter::VisitEditableRangeEnd(System::SharedPtr<Aspose::Words::EditableRangeEnd> editableRangeEnd)
{
    mBuilder->AppendLine(u" -- End of editable range --\n");
    
    mInsideEditableRange = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExEditableRange::EditableRangePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mInsideEditableRange)
    {
        mBuilder->AppendLine(System::String(u"\t\"") + run->get_Text() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}


RTTI_INFO_IMPL_HASH(3993618825u, ::Aspose::Words::ApiExamples::ExEditableRange, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExEditableRange : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExEditableRange> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExEditableRange>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExEditableRange> ExEditableRange::s_instance;

} // namespace gtest_test

void ExEditableRange::CreateAndRemove()
{
    //ExStart
    //ExFor:DocumentBuilder.StartEditableRange
    //ExFor:DocumentBuilder.EndEditableRange
    //ExFor:EditableRange
    //ExFor:EditableRange.EditableRangeEnd
    //ExFor:EditableRange.EditableRangeStart
    //ExFor:EditableRange.Id
    //ExFor:EditableRange.Remove
    //ExFor:EditableRangeEnd.EditableRangeStart
    //ExFor:EditableRangeEnd.Id
    //ExFor:EditableRangeEnd.NodeType
    //ExFor:EditableRangeStart.EditableRange
    //ExFor:EditableRangeStart.Id
    //ExFor:EditableRangeStart.NodeType
    //ExSummary:Shows how to work with an editable range.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->Protect(Aspose::Words::ProtectionType::ReadOnly, u"MyPassword");
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(System::String(u"Hello world! Since we have set the document's protection level to read-only,") + u" we cannot edit this paragraph without the password.");
    
    // Editable ranges allow us to leave parts of protected documents open for editing.
    System::SharedPtr<Aspose::Words::EditableRangeStart> editableRangeStart = builder->StartEditableRange();
    builder->Writeln(u"This paragraph is inside an editable range, and can be edited.");
    System::SharedPtr<Aspose::Words::EditableRangeEnd> editableRangeEnd = builder->EndEditableRange();
    
    // A well-formed editable range has a start node, and end node.
    // These nodes have matching IDs and encompass editable nodes.
    System::SharedPtr<Aspose::Words::EditableRange> editableRange = editableRangeStart->get_EditableRange();
    
    ASSERT_EQ(editableRangeStart->get_Id(), editableRange->get_Id());
    ASSERT_EQ(editableRangeEnd->get_Id(), editableRange->get_Id());
    
    // Different parts of the editable range link to each other.
    ASSERT_EQ(editableRangeStart->get_Id(), editableRange->get_EditableRangeStart()->get_Id());
    ASSERT_EQ(editableRangeStart->get_Id(), editableRangeEnd->get_EditableRangeStart()->get_Id());
    ASSERT_EQ(editableRange->get_Id(), editableRangeStart->get_EditableRange()->get_Id());
    ASSERT_EQ(editableRangeEnd->get_Id(), editableRange->get_EditableRangeEnd()->get_Id());
    
    // We can access the node types of each part like this. The editable range itself is not a node,
    // but an entity which consists of a start, an end, and their enclosed contents.
    ASSERT_EQ(Aspose::Words::NodeType::EditableRangeStart, editableRangeStart->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::EditableRangeEnd, editableRangeEnd->get_NodeType());
    
    builder->Writeln(u"This paragraph is outside the editable range, and cannot be edited.");
    
    doc->Save(get_ArtifactsDir() + u"EditableRange.CreateAndRemove.docx");
    
    // Remove an editable range. All the nodes that were inside the range will remain intact.
    editableRange->Remove();
    //ExEnd
    
    ASSERT_EQ(System::String(u"Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r") + u"This paragraph is inside an editable range, and can be edited.\r" + u"This paragraph is outside the editable range, and cannot be edited.", doc->GetText().Trim());
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeStart, true)->get_Count());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"EditableRange.CreateAndRemove.docx");
    
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, doc->get_ProtectionType());
    ASSERT_EQ(System::String(u"Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r") + u"This paragraph is inside an editable range, and can be edited.\r" + u"This paragraph is outside the editable range, and cannot be edited.", doc->GetText().Trim());
    
    editableRange = (System::ExplicitCast<Aspose::Words::EditableRangeStart>(doc->GetChild(Aspose::Words::NodeType::EditableRangeStart, 0, true)))->get_EditableRange();
    
    Aspose::Words::ApiExamples::TestUtil::VerifyEditableRange(0, System::String::Empty, Aspose::Words::EditorType::Unspecified, editableRange);
}

namespace gtest_test
{

TEST_F(ExEditableRange, CreateAndRemove)
{
    s_instance->CreateAndRemove();
}

} // namespace gtest_test

void ExEditableRange::Nested()
{
    //ExStart
    //ExFor:DocumentBuilder.StartEditableRange
    //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
    //ExFor:EditableRange.EditorGroup
    //ExSummary:Shows how to create nested editable ranges.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->Protect(Aspose::Words::ProtectionType::ReadOnly, u"MyPassword");
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(System::String(u"Hello world! Since we have set the document's protection level to read-only, ") + u"we cannot edit this paragraph without the password.");
    
    // Create two nested editable ranges.
    System::SharedPtr<Aspose::Words::EditableRangeStart> outerEditableRangeStart = builder->StartEditableRange();
    builder->Writeln(u"This paragraph inside the outer editable range and can be edited.");
    
    System::SharedPtr<Aspose::Words::EditableRangeStart> innerEditableRangeStart = builder->StartEditableRange();
    builder->Writeln(u"This paragraph inside both the outer and inner editable ranges and can be edited.");
    
    // Currently, the document builder's node insertion cursor is in more than one ongoing editable range.
    // When we want to end an editable range in this situation,
    // we need to specify which of the ranges we wish to end by passing its EditableRangeStart node.
    builder->EndEditableRange(innerEditableRangeStart);
    
    builder->Writeln(u"This paragraph inside the outer editable range and can be edited.");
    
    builder->EndEditableRange(outerEditableRangeStart);
    
    builder->Writeln(u"This paragraph is outside any editable ranges, and cannot be edited.");
    
    // If a region of text has two overlapping editable ranges with specified groups,
    // the combined group of users excluded by both groups are prevented from editing it.
    outerEditableRangeStart->get_EditableRange()->set_EditorGroup(Aspose::Words::EditorType::Everyone);
    innerEditableRangeStart->get_EditableRange()->set_EditorGroup(Aspose::Words::EditorType::Contributors);
    
    doc->Save(get_ArtifactsDir() + u"EditableRange.Nested.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"EditableRange.Nested.docx");
    
    ASSERT_EQ(System::String(u"Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r") + u"This paragraph inside the outer editable range and can be edited.\r" + u"This paragraph inside both the outer and inner editable ranges and can be edited.\r" + u"This paragraph inside the outer editable range and can be edited.\r" + u"This paragraph is outside any editable ranges, and cannot be edited.", doc->GetText().Trim());
    
    System::SharedPtr<Aspose::Words::EditableRange> editableRange = (System::ExplicitCast<Aspose::Words::EditableRangeStart>(doc->GetChild(Aspose::Words::NodeType::EditableRangeStart, 0, true)))->get_EditableRange();
    
    Aspose::Words::ApiExamples::TestUtil::VerifyEditableRange(0, System::String::Empty, Aspose::Words::EditorType::Everyone, editableRange);
    
    editableRange = (System::ExplicitCast<Aspose::Words::EditableRangeStart>(doc->GetChild(Aspose::Words::NodeType::EditableRangeStart, 1, true)))->get_EditableRange();
    
    Aspose::Words::ApiExamples::TestUtil::VerifyEditableRange(1, System::String::Empty, Aspose::Words::EditorType::Contributors, editableRange);
}

namespace gtest_test
{

TEST_F(ExEditableRange, Nested)
{
    s_instance->Nested();
}

} // namespace gtest_test

void ExEditableRange::Visitor()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->Protect(Aspose::Words::ProtectionType::ReadOnly, u"MyPassword");
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(System::String(u"Hello world! Since we have set the document's protection level to read-only,") + u" we cannot edit this paragraph without the password.");
    
    // When we write-protect documents, editable ranges allow us to pick specific areas that users may edit.
    // There are two mutually exclusive ways to narrow down the list of allowed editors.
    // 1 -  Specify a user:
    System::SharedPtr<Aspose::Words::EditableRange> editableRange = builder->StartEditableRange()->get_EditableRange();
    editableRange->set_SingleUser(u"john.doe@myoffice.com");
    builder->Writeln(System::String::Format(u"This paragraph is inside the first editable range, can only be edited by {0}.", editableRange->get_SingleUser()));
    builder->EndEditableRange();
    
    ASSERT_EQ(Aspose::Words::EditorType::Unspecified, editableRange->get_EditorGroup());
    
    // 2 -  Specify a group that allowed users are associated with:
    editableRange = builder->StartEditableRange()->get_EditableRange();
    editableRange->set_EditorGroup(Aspose::Words::EditorType::Administrators);
    builder->Writeln(System::String::Format(u"This paragraph is inside the first editable range, can only be edited by {0}.", editableRange->get_EditorGroup()));
    builder->EndEditableRange();
    
    ASSERT_EQ(System::String::Empty, editableRange->get_SingleUser());
    
    builder->Writeln(u"This paragraph is outside the editable range, and cannot be edited by anybody.");
    
    // Print details and contents of every editable range in the document.
    auto editableRangePrinter = System::MakeObject<Aspose::Words::ApiExamples::ExEditableRange::EditableRangePrinter>();
    
    doc->Accept(editableRangePrinter);
    
    std::cout << editableRangePrinter->ToText() << std::endl;
}

namespace gtest_test
{

TEST_F(ExEditableRange, Visitor)
{
    s_instance->Visitor();
}

} // namespace gtest_test

void ExEditableRange::IncorrectStructureException()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Assert that isn't valid structure for the current document.
    ASSERT_THROW(static_cast<std::function<void()>>([&builder]() -> void
    {
        builder->EndEditableRange();
    })(), System::InvalidOperationException);
    
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
    System::SharedPtr<Aspose::Words::Document> doc = Aspose::Words::ApiExamples::DocumentHelper::CreateDocumentFillWithDummyText();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::EditableRangeStart> startRange1 = builder->StartEditableRange();
    
    builder->Writeln(u"EditableRange_1_1");
    builder->Writeln(u"EditableRange_1_2");
    
    startRange1->get_EditableRange()->set_EditorGroup(Aspose::Words::EditorType::Everyone);
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    // Assert that it's not valid structure and editable ranges aren't added to the current document.
    System::SharedPtr<Aspose::Words::NodeCollection> startNodes = doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeStart, true);
    ASSERT_EQ(0, startNodes->get_Count());
    
    System::SharedPtr<Aspose::Words::NodeCollection> endNodes = doc->GetChildNodes(Aspose::Words::NodeType::EditableRangeEnd, true);
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
} // namespace Words
} // namespace Aspose
