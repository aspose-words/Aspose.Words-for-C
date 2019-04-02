#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/object_ext.h>
#include <Layout/Public/LayoutOptions.h>
#include <Layout/Public/RevisionOptions.h>
#include <Layout/Public/ShowInBalloons.h>
#include <Model/Document/Document.h>
#include <Model/Revisions/RevisionCollection.h>
#include <Model/Revisions/RevisionType.h>
#include <Model/Revisions/RevisionGroupCollection.h>
#include <Model/Revisions/RevisionGroup.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Text/ParagraphCollection.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;

namespace
{
    void AcceptRevisions(System::String const &dataDir)
    {
        // ExStart:AcceptAllRevisions
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
        // Start tracking and make some revisions.
        doc->StartTrackRevisions(u"Author");
        doc->get_FirstSection()->get_Body()->AppendParagraph(u"Hello world!");
        // Revisions will now show up as normal text in the output document.
        doc->AcceptAllRevisions();
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithRevisions.AcceptRevisions.doc");
        doc->Save(outputPath);
        // ExEnd:AcceptAllRevisions
        std::cout << "All revisions accepted." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void GetRevisionTypes(System::String const &dataDir)
    {
        // ExStart:GetRevisionTypes
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Revisions.docx");

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

    void GetRevisionGroups(System::String const &dataDir)
    {
        // ExStart:GetRevisionGroups
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Revisions.docx");
        for (System::SharedPtr<RevisionGroup> group : System::IterateOver(doc->get_Revisions()->get_Groups()))
        {
            std::cout << group->get_Author().ToUtf8String() << ", :" << System::ObjectExt::Box<RevisionType>(group->get_RevisionType())->ToString().ToUtf8String() << std::endl;
            std::cout << group->get_Text().ToUtf8String() << std::endl;
        }
        // ExEnd:GetRevisionGroups
    }

    void SetShowCommentsInPDF(System::String const &dataDir)
    {
        // ExStart:SetShowCommentsinPDF
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Revisions.docx");

        //Do not render the comments in PDF
        doc->get_LayoutOptions()->set_ShowComments(false);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithRevisions.SetShowCommentsinPDF.pdf");
        doc->Save(outputPath);
        // ExEnd:SetShowCommentsinPDF
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetShowInBalloons(System::String const &dataDir)
    {
        // ExStart:SetShowInBalloons
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Revisions.docx");

        // Renders insert and delete revisions inline, format revisions in balloons.
        doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowInBalloons(Aspose::Words::Layout::ShowInBalloons::Format);

        // Renders insert revisions inline, delete and format revisions in balloons.
        //doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowInBalloons(Aspose::Words::Layout::ShowInBalloons::FormatAndDelete);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithRevisions.SetShowInBalloons.pdf");
        doc->Save(outputPath);
        // ExEnd:SetShowInBalloons
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;

    }
}

void WorkingWithRevisions()
{
    std::cout << "WorkingWithRevisions example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    AcceptRevisions(dataDir);
    GetRevisionTypes(dataDir);
    GetRevisionGroups(dataDir);
    SetShowCommentsInPDF(dataDir);
    SetShowInBalloons(dataDir);
    std::cout << "WorkingWithRevisions example finished." << std::endl << std::endl;
}