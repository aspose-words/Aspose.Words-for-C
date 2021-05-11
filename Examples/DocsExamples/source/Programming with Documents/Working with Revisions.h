#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Layout/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/RevisionOptions.h>
#include <Aspose.Words.Cpp/Layout/ShowInBalloons.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/MeasurementUnits.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Revision.h>
#include <Aspose.Words.Cpp/RevisionCollection.h>
#include <Aspose.Words.Cpp/RevisionGroup.h>
#include <Aspose.Words.Cpp/RevisionGroupCollection.h>
#include <Aspose.Words.Cpp/RevisionType.h>
#include <Aspose.Words.Cpp/RevisionsView.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Layout;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithRevisions : public DocsExamplesBase
{
public:
    void AcceptRevisions()
    {
        //ExStart:AcceptAllRevisions
        auto doc = MakeObject<Document>();
        SharedPtr<Body> body = doc->get_FirstSection()->get_Body();
        SharedPtr<Paragraph> para = body->get_FirstParagraph();

        // Add text to the first paragraph, then add two more paragraphs.
        para->AppendChild(MakeObject<Run>(doc, u"Paragraph 1. "));
        body->AppendParagraph(u"Paragraph 2. ");
        body->AppendParagraph(u"Paragraph 3. ");

        // We have three paragraphs, none of which registered as any type of revision
        // If we add/remove any content in the document while tracking revisions,
        // they will be displayed as such in the document and can be accepted/rejected.
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());

        // This paragraph is a revision and will have the according "IsInsertRevision" flag set.
        para = body->AppendParagraph(u"Paragraph 4. ");
        ASSERT_TRUE(para->get_IsInsertRevision());

        // Get the document's paragraph collection and remove a paragraph.
        SharedPtr<ParagraphCollection> paragraphs = body->get_Paragraphs();
        ASSERT_EQ(4, paragraphs->get_Count());
        para = paragraphs->idx_get(2);
        para->Remove();

        // Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
        // and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions.
        ASSERT_EQ(4, paragraphs->get_Count());
        ASSERT_TRUE(para->get_IsDeleteRevision());

        // The delete revision paragraph is removed once we accept changes.
        doc->AcceptAllRevisions();
        ASSERT_EQ(3, paragraphs->get_Count());
        ASSERT_EQ(0, para->get_Count());

        // Stopping the tracking of revisions makes this text appear as normal text.
        // Revisions are not counted when the document is changed.
        doc->StopTrackRevisions();

        // Save the document.
        doc->Save(ArtifactsDir + u"WorkingWithRevisions.AcceptRevisions.docx");
        //ExEnd:AcceptAllRevisions
    }

    void GetRevisionTypes()
    {
        //ExStart:GetRevisionTypes
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
        for (int i = 0; i < paragraphs->get_Count(); i++)
        {
            if (paragraphs->idx_get(i)->get_IsMoveFromRevision())
            {
                std::cout << "The paragraph " << i << " has been moved (deleted)." << std::endl;
            }
            if (paragraphs->idx_get(i)->get_IsMoveToRevision())
            {
                std::cout << "The paragraph " << i << " has been moved (inserted)." << std::endl;
            }
        }
        //ExEnd:GetRevisionTypes
    }

    void GetRevisionGroups()
    {
        //ExStart:GetRevisionGroups
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        for (auto group : doc->get_Revisions()->get_Groups())
        {
            std::cout << String::Format(u"{0}, {1}:", group->get_Author(), group->get_RevisionType()) << std::endl;
            std::cout << group->get_Text() << std::endl;
        }
        //ExEnd:GetRevisionGroups
    }

    void RemoveCommentsInPdf()
    {
        //ExStart:RemoveCommentsInPDF
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        // Do not render the comments in PDF.
        doc->get_LayoutOptions()->set_ShowComments(false);

        doc->Save(ArtifactsDir + u"WorkingWithRevisions.RemoveCommentsInPdf.pdf");
        //ExEnd:RemoveCommentsInPDF
    }

    void ShowRevisionsInBalloons()
    {
        //ExStart:ShowRevisionsInBalloons
        //ExStart:SetMeasurementUnit
        //ExStart:SetRevisionBarsPosition
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        // Renders insert and delete revisions inline, format revisions in balloons.
        doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowInBalloons(ShowInBalloons::Format);
        doc->get_LayoutOptions()->get_RevisionOptions()->set_MeasurementUnit(MeasurementUnits::Inches);

        // Renders revision bars on the right side of a page.
        doc->get_LayoutOptions()->get_RevisionOptions()->set_RevisionBarsPosition(HorizontalAlignment::Right);

        // Renders insert revisions inline, delete and format revisions in balloons.
        // doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.FormatAndDelete;

        doc->Save(ArtifactsDir + u"WorkingWithRevisions.ShowRevisionsInBalloons.pdf");
        //ExEnd:SetRevisionBarsPosition
        //ExEnd:SetMeasurementUnit
        //ExEnd:ShowRevisionsInBalloons
    }

    void GetRevisionGroupDetails()
    {
        //ExStart:GetRevisionGroupDetails
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        for (auto revision : System::IterateOver(doc->get_Revisions()))
        {
            String groupText =
                revision->get_Group() != nullptr ? String(u"Revision group text: ") + revision->get_Group()->get_Text() : u"Revision has no group";

            std::cout << (String(u"Type: ") + System::ObjectExt::ToString(revision->get_RevisionType())) << std::endl;
            std::cout << (String(u"Author: ") + revision->get_Author()) << std::endl;
            std::cout << (String(u"Date: ") + revision->get_DateTime()) << std::endl;
            std::cout << (String(u"Revision text: ") + revision->get_ParentNode()->ToString(SaveFormat::Text)) << std::endl;
            std::cout << groupText << std::endl;
        }
        //ExEnd:GetRevisionGroupDetails
    }

    void AccessRevisedVersion()
    {
        //ExStart:AccessRevisedVersion
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");
        doc->UpdateListLabels();

        // Switch to the revised version of the document.
        doc->set_RevisionsView(RevisionsView::Final);

        for (auto revision : System::IterateOver(doc->get_Revisions()))
        {
            if (revision->get_ParentNode()->get_NodeType() == NodeType::Paragraph)
            {
                auto paragraph = System::DynamicCast<Paragraph>(revision->get_ParentNode());
                if (paragraph->get_IsListItem())
                {
                    std::cout << paragraph->get_ListLabel()->get_LabelString() << std::endl;
                    std::cout << paragraph->get_ListFormat()->get_ListLevel() << std::endl;
                }
            }
        }
        //ExEnd:AccessRevisedVersion
    }

    void MoveNodeInTrackedDocument()
    {
        //ExStart:MoveNodeInTrackedDocument
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Paragraph 1");
        builder->Writeln(u"Paragraph 2");
        builder->Writeln(u"Paragraph 3");
        builder->Writeln(u"Paragraph 4");
        builder->Writeln(u"Paragraph 5");
        builder->Writeln(u"Paragraph 6");
        SharedPtr<Body> body = doc->get_FirstSection()->get_Body();
        std::cout << "Paragraph count: " << body->get_Paragraphs()->get_Count() << std::endl;

        // Start tracking revisions.
        doc->StartTrackRevisions(u"Author", System::DateTime(2020, 12, 23, 14, 0, 0));

        // Generate revisions when moving a node from one location to another.
        SharedPtr<Node> node = body->get_Paragraphs()->idx_get(3);
        SharedPtr<Node> endNode = body->get_Paragraphs()->idx_get(5)->get_NextSibling();
        SharedPtr<Node> referenceNode = body->get_Paragraphs()->idx_get(0);
        while (node != endNode)
        {
            SharedPtr<Node> nextNode = node->get_NextSibling();
            body->InsertBefore(node, referenceNode);
            node = nextNode;
        }

        // Stop the process of tracking revisions.
        doc->StopTrackRevisions();

        // There are 3 additional paragraphs in the move-from range.
        std::cout << "Paragraph count: " << body->get_Paragraphs()->get_Count() << std::endl;
        doc->Save(ArtifactsDir + u"WorkingWithRevisions.MoveNodeInTrackedDocument.docx");
        //ExEnd:MoveNodeInTrackedDocument
    }

    void ShapeRevision()
    {
        //ExStart:ShapeRevision
        auto doc = MakeObject<Document>();

        // Insert an inline shape without tracking revisions.
        ASSERT_FALSE(doc->get_TrackRevisions());
        auto shape = MakeObject<Shape>(doc, ShapeType::Cube);
        shape->set_WrapType(WrapType::Inline);
        shape->set_Width(100.0);
        shape->set_Height(100.0);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

        // Start tracking revisions and then insert another shape.
        doc->StartTrackRevisions(u"John Doe");
        shape = MakeObject<Shape>(doc, ShapeType::Sun);
        shape->set_WrapType(WrapType::Inline);
        shape->set_Width(100.0);
        shape->set_Height(100.0);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

        // Get the document's shape collection which includes just the two shapes we added.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes =
            doc->GetChildNodes(NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_ToList();
        ASSERT_EQ(2, shapes->get_Count());

        // Remove the first shape.
        shapes->idx_get(0)->Remove();

        // Because we removed that shape while changes were being tracked, the shape counts as a delete revision.
        ASSERT_EQ(ShapeType::Cube, shapes->idx_get(0)->get_ShapeType());
        ASSERT_TRUE(shapes->idx_get(0)->get_IsDeleteRevision());

        // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
        ASSERT_EQ(ShapeType::Sun, shapes->idx_get(1)->get_ShapeType());
        ASSERT_TRUE(shapes->idx_get(1)->get_IsInsertRevision());

        // The document has one shape that was moved, but shape move revisions will have two instances of that shape.
        // One will be the shape at its arrival destination and the other will be the shape at its original location.
        doc = MakeObject<Document>(MyDir + u"Revision shape.docx");

        shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_ToList();
        ASSERT_EQ(4, shapes->get_Count());

        // This is the move to revision, also the shape at its arrival destination.
        ASSERT_FALSE(shapes->idx_get(0)->get_IsMoveFromRevision());
        ASSERT_TRUE(shapes->idx_get(0)->get_IsMoveToRevision());

        // This is the move from revision, which is the shape at its original location.
        ASSERT_TRUE(shapes->idx_get(1)->get_IsMoveFromRevision());
        ASSERT_FALSE(shapes->idx_get(1)->get_IsMoveToRevision());
        //ExEnd:ShapeRevision
    }
};

}} // namespace DocsExamples::Programming_with_Documents
