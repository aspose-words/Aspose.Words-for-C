#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <system/collections/list.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment {

class RemoveContent : public DocsExamplesBase
{
public:
    void RemovePageBreaks()
    {
        //ExStart:OpenFromFile
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        //ExEnd:OpenFromFile

        // In Aspose.Words section breaks are represented as separate Section nodes in the document.
        // To remove these separate sections, the sections are combined.
        RemovePageBreaks(doc);
        RemoveSectionBreaks(doc);

        doc->Save(ArtifactsDir + u"RemoveContent.RemovePageBreaks.docx");
    }

    void RemoveFooters()
    {
        //ExStart:RemoveFooters
        auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        for (const auto& section : System::IterateOver<Section>(doc))
        {
            // Up to three different footers are possible in a section (for first, even and odd pages)
            // we check and delete all of them.
            SharedPtr<HeaderFooter> footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterFirst);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            // Primary footer is the footer used for odd pages.
            footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterEven);
            if (footer != nullptr)
            {
                footer->Remove();
            }
        }

        doc->Save(ArtifactsDir + u"RemoveContent.RemoveFooters.docx");
        //ExEnd:RemoveFooters
    }

    void RemoveToc()
    {
        auto doc = MakeObject<Document>(MyDir + u"Table of contents.docx");

        // Remove the first table of contents from the document.
        RemoveTableOfContents(doc, 0);

        doc->Save(ArtifactsDir + u"RemoveContent.RemoveToc.doc");
    }

    /// <summary>
    /// Removes the specified table of contents field from the document.
    /// </summary>
    /// <param name="doc">The document to remove the field from.</param>
    /// <param name="index">The zero-based index of the TOC to remove.</param>
    void RemoveTableOfContents(SharedPtr<Document> doc, int index)
    {
        // Store the FieldStart nodes of TOC fields in the document for quick access.
        SharedPtr<System::Collections::Generic::List<SharedPtr<FieldStart>>> fieldStarts =
            MakeObject<System::Collections::Generic::List<SharedPtr<FieldStart>>>();
        // This is a list to store the nodes found inside the specified TOC. They will be removed at the end of this method.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> nodeList = MakeObject<System::Collections::Generic::List<SharedPtr<Node>>>();

        for (const auto& start : System::IterateOver<FieldStart>(doc->GetChildNodes(NodeType::FieldStart, true)))
        {
            if (start->get_FieldType() == FieldType::FieldTOC)
            {
                fieldStarts->Add(start);
            }
        }

        // Ensure the TOC specified by the passed index exists.
        if (index > fieldStarts->get_Count() - 1)
        {
            throw System::ArgumentOutOfRangeException(u"TOC index is out of range");
        }

        bool isRemoving = true;

        SharedPtr<Node> currentNode = fieldStarts->idx_get(index);
        while (isRemoving)
        {
            // It is safer to store these nodes and delete them all at once later.
            nodeList->Add(currentNode);
            currentNode = currentNode->NextPreOrder(doc);

            // Once we encounter a FieldEnd node of type FieldTOC,
            // we know we are at the end of the current TOC and stop here.
            if (currentNode->get_NodeType() == NodeType::FieldEnd)
            {
                auto fieldEnd = System::DynamicCast<FieldEnd>(currentNode);
                if (fieldEnd->get_FieldType() == FieldType::FieldTOC)
                {
                    isRemoving = false;
                }
            }
        }

        for (const auto& node : nodeList)
        {
            node->Remove();
        }
    }

protected:
    void RemovePageBreaks(SharedPtr<Document> doc)
    {
        SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);

        for (const auto& para : System::IterateOver<Paragraph>(paragraphs))
        {
            // If the paragraph has a page break before the set, then clear it.
            if (para->get_ParagraphFormat()->get_PageBreakBefore())
            {
                para->get_ParagraphFormat()->set_PageBreakBefore(false);
            }

            // Check all runs in the paragraph for page breaks and remove them.
            for (const auto& run : System::IterateOver<Aspose::Words::Run>(para->get_Runs()))
            {
                if (run->get_Text().Contains(ControlChar::PageBreak()))
                {
                    run->set_Text(run->get_Text().Replace(ControlChar::PageBreak(), String::Empty));
                }
            }
        }
    }

    void RemoveSectionBreaks(SharedPtr<Document> doc)
    {
        // Loop through all sections starting from the section that precedes the last one and moving to the first section.
        for (int i = doc->get_Sections()->get_Count() - 2; i >= 0; i--)
        {
            // Copy the content of the current section to the beginning of the last section.
            doc->get_LastSection()->PrependContent(doc->get_Sections()->idx_get(i));
            // Remove the copied section.
            doc->get_Sections()->idx_get(i)->Remove();
        }
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment
