#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRange.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeStart.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditorType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/exceptions.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace ApiExamples {

class ExEditableRange : public ApiExampleBase
{
public:
    void CreateAndRemove()
    {
        //ExStart
        //ExFor:DocumentBuilder.StartEditableRange
        //ExFor:DocumentBuilder.EndEditableRange()
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
        auto doc = MakeObject<Document>();
        doc->Protect(ProtectionType::ReadOnly, u"MyPassword");

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(String(u"Hello world! Since we have set the document's protection level to read-only,") +
                         u" we cannot edit this paragraph without the password.");

        // Editable ranges allow us to leave parts of protected documents open for editing.
        SharedPtr<EditableRangeStart> editableRangeStart = builder->StartEditableRange();
        builder->Writeln(u"This paragraph is inside an editable range, and can be edited.");
        SharedPtr<EditableRangeEnd> editableRangeEnd = builder->EndEditableRange();

        // A well-formed editable range has a start node, and end node.
        // These nodes have matching IDs and encompass editable nodes.
        SharedPtr<EditableRange> editableRange = editableRangeStart->get_EditableRange();

        ASSERT_EQ(editableRangeStart->get_Id(), editableRange->get_Id());
        ASSERT_EQ(editableRangeEnd->get_Id(), editableRange->get_Id());

        // Different parts of the editable range link to each other.
        ASSERT_EQ(editableRangeStart->get_Id(), editableRange->get_EditableRangeStart()->get_Id());
        ASSERT_EQ(editableRangeStart->get_Id(), editableRangeEnd->get_EditableRangeStart()->get_Id());
        ASSERT_EQ(editableRange->get_Id(), editableRangeStart->get_EditableRange()->get_Id());
        ASSERT_EQ(editableRangeEnd->get_Id(), editableRange->get_EditableRangeEnd()->get_Id());

        // We can access the node types of each part like this. The editable range itself is not a node,
        // but an entity which consists of a start, an end, and their enclosed contents.
        ASSERT_EQ(NodeType::EditableRangeStart, editableRangeStart->get_NodeType());
        ASSERT_EQ(NodeType::EditableRangeEnd, editableRangeEnd->get_NodeType());

        builder->Writeln(u"This paragraph is outside the editable range, and cannot be edited.");

        doc->Save(ArtifactsDir + u"EditableRange.CreateAndRemove.docx");

        // Remove an editable range. All the nodes that were inside the range will remain intact.
        editableRange->Remove();
        //ExEnd

        ASSERT_EQ(
            String(u"Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r") +
                u"This paragraph is inside an editable range, and can be edited.\r" + u"This paragraph is outside the editable range, and cannot be edited.",
            doc->GetText().Trim());
        ASSERT_EQ(0, doc->GetChildNodes(NodeType::EditableRangeStart, true)->get_Count());

        doc = MakeObject<Document>(ArtifactsDir + u"EditableRange.CreateAndRemove.docx");

        ASSERT_EQ(ProtectionType::ReadOnly, doc->get_ProtectionType());
        ASSERT_EQ(
            String(u"Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r") +
                u"This paragraph is inside an editable range, and can be edited.\r" + u"This paragraph is outside the editable range, and cannot be edited.",
            doc->GetText().Trim());

        editableRange = (System::DynamicCast<EditableRangeStart>(doc->GetChild(NodeType::EditableRangeStart, 0, true)))->get_EditableRange();

        TestUtil::VerifyEditableRange(0, String::Empty, EditorType::Unspecified, editableRange);
    }

    void Nested()
    {
        //ExStart
        //ExFor:DocumentBuilder.StartEditableRange
        //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
        //ExFor:EditableRange.EditorGroup
        //ExSummary:Shows how to create nested editable ranges.
        auto doc = MakeObject<Document>();
        doc->Protect(ProtectionType::ReadOnly, u"MyPassword");

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(String(u"Hello world! Since we have set the document's protection level to read-only, ") +
                         u"we cannot edit this paragraph without the password.");

        // Create two nested editable ranges.
        SharedPtr<EditableRangeStart> outerEditableRangeStart = builder->StartEditableRange();
        builder->Writeln(u"This paragraph inside the outer editable range and can be edited.");

        SharedPtr<EditableRangeStart> innerEditableRangeStart = builder->StartEditableRange();
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
        outerEditableRangeStart->get_EditableRange()->set_EditorGroup(EditorType::Everyone);
        innerEditableRangeStart->get_EditableRange()->set_EditorGroup(EditorType::Contributors);

        doc->Save(ArtifactsDir + u"EditableRange.Nested.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"EditableRange.Nested.docx");

        ASSERT_EQ(
            String(u"Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r") +
                u"This paragraph inside the outer editable range and can be edited.\r" +
                u"This paragraph inside both the outer and inner editable ranges and can be edited.\r" +
                u"This paragraph inside the outer editable range and can be edited.\r" +
                u"This paragraph is outside any editable ranges, and cannot be edited.",
            doc->GetText().Trim());

        SharedPtr<EditableRange> editableRange =
            (System::DynamicCast<EditableRangeStart>(doc->GetChild(NodeType::EditableRangeStart, 0, true)))->get_EditableRange();

        TestUtil::VerifyEditableRange(0, String::Empty, EditorType::Everyone, editableRange);

        editableRange = (System::DynamicCast<EditableRangeStart>(doc->GetChild(NodeType::EditableRangeStart, 1, true)))->get_EditableRange();

        TestUtil::VerifyEditableRange(1, String::Empty, EditorType::Contributors, editableRange);
    }

    //ExStart
    //ExFor:EditableRange
    //ExFor:EditableRange.EditorGroup
    //ExFor:EditableRange.SingleUser
    //ExFor:EditableRangeEnd
    //ExFor:EditableRangeEnd.Accept(DocumentVisitor)
    //ExFor:EditableRangeStart
    //ExFor:EditableRangeStart.Accept(DocumentVisitor)
    //ExFor:EditorType
    //ExSummary:Shows how to limit the editing rights of editable ranges to a specific group/user.
    void Visitor()
    {
        auto doc = MakeObject<Document>();
        doc->Protect(ProtectionType::ReadOnly, u"MyPassword");

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(String(u"Hello world! Since we have set the document's protection level to read-only,") +
                         u" we cannot edit this paragraph without the password.");

        // When we write-protect documents, editable ranges allow us to pick specific areas that users may edit.
        // There are two mutually exclusive ways to narrow down the list of allowed editors.
        // 1 -  Specify a user:
        SharedPtr<EditableRange> editableRange = builder->StartEditableRange()->get_EditableRange();
        editableRange->set_SingleUser(u"john.doe@myoffice.com");
        builder->Writeln(String::Format(u"This paragraph is inside the first editable range, can only be edited by {0}.", editableRange->get_SingleUser()));
        builder->EndEditableRange();

        ASSERT_EQ(EditorType::Unspecified, editableRange->get_EditorGroup());

        // 2 -  Specify a group that allowed users are associated with:
        editableRange = builder->StartEditableRange()->get_EditableRange();
        editableRange->set_EditorGroup(EditorType::Administrators);
        builder->Writeln(String::Format(u"This paragraph is inside the first editable range, can only be edited by {0}.", editableRange->get_EditorGroup()));
        builder->EndEditableRange();

        ASSERT_EQ(String::Empty, editableRange->get_SingleUser());

        builder->Writeln(u"This paragraph is outside the editable range, and cannot be edited by anybody.");

        // Print details and contents of every editable range in the document.
        auto editableRangePrinter = MakeObject<ExEditableRange::EditableRangePrinter>();

        doc->Accept(editableRangePrinter);

        std::cout << editableRangePrinter->ToText() << std::endl;
    }

    /// <summary>
    /// Collects properties and contents of visited editable ranges in a string.
    /// </summary>
    class EditableRangePrinter : public DocumentVisitor
    {
    public:
        EditableRangePrinter() : mInsideEditableRange(false)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
        }

        String ToText()
        {
            return mBuilder->ToString();
        }

        void Reset()
        {
            mBuilder->Clear();
            mInsideEditableRange = false;
        }

        /// <summary>
        /// Called when an EditableRangeStart node is encountered in the document.
        /// </summary>
        VisitorAction VisitEditableRangeStart(SharedPtr<EditableRangeStart> editableRangeStart) override
        {
            mBuilder->AppendLine(u" -- Editable range found! -- ");
            mBuilder->AppendLine(String(u"\tID:\t\t") + editableRangeStart->get_Id());
            if (editableRangeStart->get_EditableRange()->get_SingleUser() == String::Empty)
            {
                mBuilder->AppendLine(String(u"\tGroup:\t") + System::ObjectExt::ToString(editableRangeStart->get_EditableRange()->get_EditorGroup()));
            }
            else
            {
                mBuilder->AppendLine(String(u"\tUser:\t") + editableRangeStart->get_EditableRange()->get_SingleUser());
            }
            mBuilder->AppendLine(u"\tContents:");

            mInsideEditableRange = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when an EditableRangeEnd node is encountered in the document.
        /// </summary>
        VisitorAction VisitEditableRangeEnd(SharedPtr<EditableRangeEnd> editableRangeEnd) override
        {
            mBuilder->AppendLine(u" -- End of editable range --\n");

            mInsideEditableRange = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document. This visitor only records runs that are inside editable ranges.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (mInsideEditableRange)
            {
                mBuilder->AppendLine(String(u"\t\"") + run->get_Text() + u"\"");
            }

            return VisitorAction::Continue;
        }

    private:
        bool mInsideEditableRange;
        SharedPtr<System::Text::StringBuilder> mBuilder;
    };
    //ExEnd

    void IncorrectStructureException()
    {
        auto doc = MakeObject<Document>();

        auto builder = MakeObject<DocumentBuilder>(doc);

        // Assert that isn't valid structure for the current document.
        ASSERT_THROW(
            static_cast<std::function<SharedPtr<EditableRangeEnd>()>>([&builder]() -> SharedPtr<EditableRangeEnd> { return builder->EndEditableRange(); })(),
            System::InvalidOperationException);

        builder->StartEditableRange();
    }

    void IncorrectStructureDoNotAdded()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<EditableRangeStart> startRange1 = builder->StartEditableRange();

        builder->Writeln(u"EditableRange_1_1");
        builder->Writeln(u"EditableRange_1_2");

        startRange1->get_EditableRange()->set_EditorGroup(EditorType::Everyone);
        doc = DocumentHelper::SaveOpen(doc);

        // Assert that it's not valid structure and editable ranges aren't added to the current document.
        SharedPtr<NodeCollection> startNodes = doc->GetChildNodes(NodeType::EditableRangeStart, true);
        ASSERT_EQ(0, startNodes->get_Count());

        SharedPtr<NodeCollection> endNodes = doc->GetChildNodes(NodeType::EditableRangeEnd, true);
        ASSERT_EQ(0, endNodes->get_Count());
    }
};

} // namespace ApiExamples
