#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace ApiExamples {

class ExMisc : public ApiExampleBase
{
public:
    void TrackingChanges()
    {
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
        doc->Save(ArtifactsDir + u"Document.Revisions.docx");
    }

    void MoveNodeWithinATrackedDocument()
    {
        // Generate document contents.
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
        doc->Save(ArtifactsDir + u"out.docx");
    }

    void ApplyDifferentPropertiesWithRevisions()
    {
        // Open a blank document.
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
        auto shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_ToList();
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
        auto nc = doc->GetChildNodes(NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_ToList();
        ASSERT_EQ(2, nc->get_Count());

        // This is the move to revision, also the shape at its arrival destination.
        ASSERT_FALSE(nc->idx_get(0)->get_IsMoveFromRevision());
        ASSERT_FALSE(nc->idx_get(0)->get_IsMoveToRevision());

        // This is the move from revision, which is the shape at its original location.
        ASSERT_FALSE(nc->idx_get(1)->get_IsMoveFromRevision());
        ASSERT_FALSE(nc->idx_get(1)->get_IsMoveToRevision());
    }
};

} // namespace ApiExamples
