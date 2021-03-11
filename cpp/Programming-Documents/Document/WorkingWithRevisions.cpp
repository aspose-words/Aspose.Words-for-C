#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/console.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/ShowInBalloons.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionType.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroupCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroup.h>
#include <Aspose.Words.Cpp/Model/Revisions/Revision.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Layout;

namespace
{
    void AcceptRevisions(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:AcceptAllRevisions
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
        // Start tracking and make some revisions.
        doc->StartTrackRevisions(u"Author");
        doc->get_FirstSection()->get_Body()->AppendParagraph(u"Hello world!");
        // Revisions will now show up as normal text in the output document.
        doc->AcceptAllRevisions();
        System::String outputPath = outputDataDir + u"WorkingWithRevisions.AcceptRevisions.doc";
        doc->Save(outputPath);
        // ExEnd:AcceptAllRevisions
        std::cout << "All revisions accepted." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void GetRevisionTypes(System::String const &inputDataDir)
    {
        // ExStart:GetRevisionTypes
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Revisions.docx");

        System::SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
        for (int32_t i = 0; i < paragraphs->get_Count(); i++)
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
        // ExEnd:GetRevisionTypes
    }

    void GetRevisionGroups(System::String const &inputDataDir)
    {
        // ExStart:GetRevisionGroups
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Revisions.docx");
        for (System::SharedPtr<RevisionGroup> group : System::IterateOver(doc->get_Revisions()->get_Groups()))
        {
            std::cout << group->get_Author().ToUtf8String() << ", :" << System::ObjectExt::Box<RevisionType>(group->get_RevisionType())->ToString().ToUtf8String() << std::endl;
            std::cout << group->get_Text().ToUtf8String() << std::endl;
        }
        // ExEnd:GetRevisionGroups
    }

    void SetShowCommentsInPDF(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetShowCommentsinPDF
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Revisions.docx");

        //Do not render the comments in PDF
        doc->get_LayoutOptions()->set_ShowComments(false);
        System::String outputPath = outputDataDir + u"WorkingWithRevisions.SetShowCommentsinPDF.pdf";
        doc->Save(outputPath);
        // ExEnd:SetShowCommentsinPDF
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void GetRevisionGroupDetails(System::String const &inputDataDir)
    {
        // ExStart:GetRevisionGroupDetails
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFormatDescription.docx");

        for (System::SharedPtr<Revision> revision : System::IterateOver(doc->get_Revisions()))
        {
            System::String groupText = revision->get_Group() != nullptr ? System::String(u"Revision group text: ") + revision->get_Group()->get_Text() : u"Revision has no group";

            std::cout << "Type: " << System::ObjectExt::ToString(revision->get_RevisionType()).ToUtf8String() << std::endl;
            std::cout << "Author: " << revision->get_Author().ToUtf8String() << std::endl;
            std::cout << "Date: " << revision->get_DateTime().ToString().ToUtf8String() << std::endl;
            std::cout << "Revision text: " << revision->get_ParentNode()->ToString(Aspose::Words::SaveFormat::Text).ToUtf8String() << std::endl;
            std::cout << groupText.ToUtf8String() << std::endl;
        }
        // ExEnd:GetRevisionGroupDetails
    }

    void AccessRevisedVersion(System::String const &inputDataDir)
    {
        // ExStart:AccessRevisedVersion
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test.docx");
        doc->UpdateListLabels();

        // Switch to the revised version of the document.
        doc->set_RevisionsView(RevisionsView::Final);

        for (System::SharedPtr<Revision> revision : System::IterateOver(doc->get_Revisions()))
        {
            if (revision->get_ParentNode()->get_NodeType() == NodeType::Paragraph)
            {
                System::SharedPtr<Paragraph> paragraph = System::DynamicCast<Paragraph>(revision->get_ParentNode());
                if (paragraph->get_IsListItem())
                {
                    // Print revised version of LabelString and ListLevel.
                    System::Console::WriteLine(paragraph->get_ListLabel()->get_LabelString());
                    System::Console::WriteLine(paragraph->get_ListFormat()->get_ListLevel());
                }
            }
        }
        // ExEnd:AccessRevisedVersion
    }

    void MoveNodeInTrackedDocument(const System::String& outputDataDir)
    {
        //ExStart:MoveNodeInTrackedDocument
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Paragraph 1");
        builder->Writeln(u"Paragraph 2");
        builder->Writeln(u"Paragraph 3");
        builder->Writeln(u"Paragraph 4");
        builder->Writeln(u"Paragraph 5");
        builder->Writeln(u"Paragraph 6");
        auto body = doc->get_FirstSection()->get_Body();
        std::cout << "Paragraph count: " << body->get_Paragraphs()->get_Count() << '\n';

        // Start tracking revisions.
        doc->StartTrackRevisions(u"Author", System::DateTime(2020, 12, 23, 14, 0, 0));

        // Generate revisions when moving a node from one location to another.
        auto node = System::StaticCast<Node>(body->get_Paragraphs()->idx_get(3));
        auto endNode = body->get_Paragraphs()->idx_get(5)->get_NextSibling();
        auto referenceNode = System::StaticCast<Node>(body->get_Paragraphs()->idx_get(0));
        while (node != endNode)
        {
            auto nextNode = node->get_NextSibling();
            body->InsertBefore(node, referenceNode);
            node = nextNode;
        }

        // Stop the process of tracking revisions.
        doc->StopTrackRevisions();

        // There are 3 additional paragraphs in the move-from range.
        std::cout << "Paragraph count: " << body->get_Paragraphs()->get_Count() << '\n';
        doc->Save(outputDataDir + u"WorkingWithRevisions.MoveNodeInTrackedDocument.docx");
        //ExEnd:MoveNodeInTrackedDocument
    }


    void ShapeRevision(const System::String& inputDataDir)
    {
        //ExStart:ShapeRevision
        auto doc = System::MakeObject<Document>();

        // Insert an inline shape without tracking revisions.
        std::cout << doc->get_TrackRevisions() << '\n';
        auto shape = System::MakeObject<Drawing::Shape>(doc, Drawing::ShapeType::Cube);
        shape->set_WrapType(Drawing::WrapType::Inline);
        shape->set_Width(100.0);
        shape->set_Height(100.0);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

        // Start tracking revisions and then insert another shape.
        doc->StartTrackRevisions(u"John Doe");
        shape = System::MakeObject<Drawing::Shape>(doc, Drawing::ShapeType::Sun);
        shape->set_WrapType(Drawing::WrapType::Inline);
        shape->set_Width(100.0);
        shape->set_Height(100.0);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

        // Get the document's shape collection which includes just the two shapes we added.
        std::vector<System::SharedPtr<Drawing::Shape>> shapes;
        for (auto node : System::IterateOver<Drawing::Shape>(doc->GetChildNodes(NodeType::Shape, true)))
        {
            shapes.push_back(node);
        }
        std::cout << shapes.size() << '\n';

        // Remove the first shape.
        shapes[0]->Remove();

        // Because we removed that shape while changes were being tracked, the shape counts as a delete revision.
        std::cout << System::ObjectExt::ToString(shapes[0]->get_ShapeType()) << '\n'
            << shapes[0]->get_IsDeleteRevision() << '\n';

        // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
        std::cout << System::ObjectExt::ToString(shapes[1]->get_ShapeType()) << '\n'
            << shapes[1]->get_IsInsertRevision() << '\n';

        // The document has one shape that was moved, but shape move revisions will have two instances of that shape.
        // One will be the shape at its arrival destination and the other will be the shape at its original location.
        doc = System::MakeObject<Document>(inputDataDir + u"Revision shape.docx");

        shapes.clear();
        for (auto node : System::IterateOver<Drawing::Shape>(doc->GetChildNodes(NodeType::Shape, true)))
        {
            shapes.push_back(node);
        }
        std::cout << shapes.size() << '\n';

        // This is the move to revision, also the shape at its arrival destination.
        std::cout << shapes[0]->get_IsMoveFromRevision() << '\n'
                  << shapes[0]->get_IsMoveToRevision() << '\n';

        // This is the move from revision, which is the shape at its original location.
        std::cout << shapes[1]->get_IsMoveFromRevision() << '\n'
                  << shapes[1]->get_IsMoveToRevision() << '\n';
        //ExEnd:ShapeRevision
    }
}

void WorkingWithRevisions()
{
    std::cout << "WorkingWithRevisions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    AcceptRevisions(inputDataDir, outputDataDir);
    GetRevisionTypes(inputDataDir);
    GetRevisionGroups(inputDataDir);
    SetShowCommentsInPDF(inputDataDir, outputDataDir);
    GetRevisionGroupDetails(inputDataDir);
    AccessRevisedVersion(inputDataDir);
    MoveNodeInTrackedDocument(outputDataDir);
    ShapeRevision(inputDataDir);
    std::cout << "WorkingWithRevisions example finished." << std::endl << std::endl;
}