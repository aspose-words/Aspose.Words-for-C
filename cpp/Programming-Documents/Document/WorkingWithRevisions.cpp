#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/object_ext.h>
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
}

void WorkingWithRevisions()
{
    std::cout << "WorkingWithRevisions example started." << std::endl;
    // ExStart:WorkingWithRevisions
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    AcceptRevisions(dataDir);
    GetRevisionTypes(dataDir);
    GetRevisionGroups(dataDir);
    // ExEnd:WorkingWithRevisions
    std::cout << "WorkingWithRevisions example finished." << std::endl << std::endl;
}