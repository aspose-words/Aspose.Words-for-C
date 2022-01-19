#include "Programming with Documents/Split Documents/Page splitter.h"

#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/SectionStart.h>
#include <system/object_ext.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents {

System::SharedPtr<Aspose::Words::Document> DocumentPageSplitter::get_Document()
{
    return pageNumberFinder->get_Document();
}

DocumentPageSplitter::DocumentPageSplitter(System::SharedPtr<Aspose::Words::Document> source)
{
    pageNumberFinder = PageNumberFinderFactory::Create(source);
}

System::SharedPtr<Aspose::Words::Document> DocumentPageSplitter::GetDocumentOfPage(int32_t pageIndex)
{
    return GetDocumentOfPageRange(pageIndex, pageIndex);
}

System::SharedPtr<Aspose::Words::Document> DocumentPageSplitter::GetDocumentOfPageRange(int32_t startIndex, int32_t endIndex)
{
    auto result = System::DynamicCast<Aspose::Words::Document>(System::StaticCast<Aspose::Words::Node>(get_Document())->Clone(false));
    for (const auto& section : System::IterateOver(pageNumberFinder->RetrieveAllNodesOnPages(startIndex, endIndex, NodeType::Section)))
    {
        result->AppendChild(result->ImportNode(section, true));
    }

    return result;
}

namespace gtest_test {

class PageSplitter : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Split_Documents::PageSplitter> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Split_Documents::PageSplitter>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Split_Documents::PageSplitter> PageSplitter::s_instance;

TEST_F(PageSplitter, SplitDocuments)
{
    s_instance->SplitDocuments();
}

} // namespace gtest_test

SectionSplitter::SectionSplitter(System::SharedPtr<PageNumberFinder> pageNumberFinder)
{
    this->pageNumberFinder = pageNumberFinder;
}

VisitorAction SectionSplitter::VisitParagraphStart(System::SharedPtr<Paragraph> paragraph)
{
    return ContinueIfCompositeAcrossPageElseSkip(paragraph);
}

VisitorAction SectionSplitter::VisitTableStart(System::SharedPtr<Table> table)
{
    return ContinueIfCompositeAcrossPageElseSkip(table);
}

VisitorAction SectionSplitter::VisitRowStart(System::SharedPtr<Row> row)
{
    return ContinueIfCompositeAcrossPageElseSkip(row);
}

VisitorAction SectionSplitter::VisitCellStart(System::SharedPtr<Cell> cell)
{
    return ContinueIfCompositeAcrossPageElseSkip(cell);
}

VisitorAction SectionSplitter::VisitStructuredDocumentTagStart(System::SharedPtr<StructuredDocumentTag> sdt)
{
    return ContinueIfCompositeAcrossPageElseSkip(sdt);
}

VisitorAction SectionSplitter::VisitSmartTagStart(System::SharedPtr<SmartTag> smartTag)
{
    return ContinueIfCompositeAcrossPageElseSkip(smartTag);
}

VisitorAction SectionSplitter::VisitSectionStart(System::SharedPtr<Section> section)
{
    auto previousSection = System::DynamicCast<Section>(section->get_PreviousSibling());

    // If there is a previous section, attempt to copy any linked header footers.
    // Otherwise, they will not appear in an extracted document if the previous section is missing.
    if (previousSection != nullptr)
    {
        System::SharedPtr<HeaderFooterCollection> previousHeaderFooters = previousSection->get_HeadersFooters();
        if (!section->get_PageSetup()->get_RestartPageNumbering())
        {
            section->get_PageSetup()->set_RestartPageNumbering(true);
            section->get_PageSetup()->set_PageStartingNumber(previousSection->get_PageSetup()->get_PageStartingNumber() +
                                                             pageNumberFinder->PageSpan(previousSection));
        }

        for (const auto& previousHeaderFooter : System::IterateOver<HeaderFooter>(previousHeaderFooters))
        {
            if (section->get_HeadersFooters()->idx_get(previousHeaderFooter->get_HeaderFooterType()) == nullptr)
            {
                auto newHeaderFooter =
                    System::DynamicCast<HeaderFooter>(previousHeaderFooters->idx_get(previousHeaderFooter->get_HeaderFooterType())->Clone(true));
                section->get_HeadersFooters()->Add(newHeaderFooter);
            }
        }
    }

    return ContinueIfCompositeAcrossPageElseSkip(section);
}

VisitorAction SectionSplitter::VisitSmartTagEnd(System::SharedPtr<SmartTag> smartTag)
{
    SplitComposite(smartTag);
    return VisitorAction::Continue;
}

VisitorAction SectionSplitter::VisitStructuredDocumentTagEnd(System::SharedPtr<StructuredDocumentTag> sdt)
{
    SplitComposite(sdt);
    return VisitorAction::Continue;
}

VisitorAction SectionSplitter::VisitCellEnd(System::SharedPtr<Cell> cell)
{
    SplitComposite(cell);
    return VisitorAction::Continue;
}

VisitorAction SectionSplitter::VisitRowEnd(System::SharedPtr<Row> row)
{
    SplitComposite(row);
    return VisitorAction::Continue;
}

VisitorAction SectionSplitter::VisitTableEnd(System::SharedPtr<Table> table)
{
    SplitComposite(table);
    return VisitorAction::Continue;
}

VisitorAction SectionSplitter::VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph)
{
    // If the paragraph contains only section break, add fake run into.
    if (paragraph->get_IsEndOfSection() && paragraph->get_ChildNodes()->get_Count() == 1 && paragraph->get_ChildNodes()->idx_get(0)->GetText() == u"\f")
    {
        auto run = System::MakeObject<Run>(paragraph->get_Document());
        paragraph->AppendChild(run);
        int32_t currentEndPageNum = pageNumberFinder->GetPageEnd(paragraph);
        pageNumberFinder->AddPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
    }

    for (const auto& clonePara : System::IterateOver<Paragraph>(SplitComposite(paragraph)))
    {
        // Remove list numbering from the cloned paragraph but leave the indent the same
        // as the paragraph is supposed to be part of the item before.
        if (paragraph->get_IsListItem())
        {
            double textPosition = clonePara->get_ListFormat()->get_ListLevel()->get_TextPosition();
            clonePara->get_ListFormat()->RemoveNumbers();
            clonePara->get_ParagraphFormat()->set_LeftIndent(textPosition);
        }

        // Reset spacing of split paragraphs in tables as additional spacing may cause them to look different.
        if (paragraph->get_IsInCell())
        {
            clonePara->get_ParagraphFormat()->set_SpaceBefore(0);
            paragraph->get_ParagraphFormat()->set_SpaceAfter(0);
        }
    }

    return VisitorAction::Continue;
}

VisitorAction SectionSplitter::VisitSectionEnd(System::SharedPtr<Section> section)
{
    for (const auto& cloneSection : System::IterateOver<Section>(SplitComposite(section)))
    {
        cloneSection->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
        cloneSection->get_PageSetup()->set_RestartPageNumbering(true);
        cloneSection->get_PageSetup()->set_PageStartingNumber(section->get_PageSetup()->get_PageStartingNumber() +
                                                              (section->get_Document()->IndexOf(cloneSection) - section->get_Document()->IndexOf(section)));
        cloneSection->get_PageSetup()->set_DifferentFirstPageHeaderFooter(false);

        // Corrects page break at the end of the section.
        SplitPageBreakCorrector::ProcessSection(cloneSection);
    }

    SplitPageBreakCorrector::ProcessSection(section);

    // Add new page numbering for the body of the section as well.
    pageNumberFinder->AddPageNumbersForNode(section->get_Body(), pageNumberFinder->GetPage(section), pageNumberFinder->GetPageEnd(section));
    return VisitorAction::Continue;
}

VisitorAction SectionSplitter::ContinueIfCompositeAcrossPageElseSkip(System::SharedPtr<CompositeNode> composite)
{
    return pageNumberFinder->PageSpan(composite) > 1 ? VisitorAction::Continue : VisitorAction::SkipThisNode;
}

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> SectionSplitter::SplitComposite(System::SharedPtr<CompositeNode> composite)
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> splitNodes =
        System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Node>>>();
    for (const auto& splitNode : System::IterateOver(FindChildSplitPositions(composite)))
    {
        splitNodes->Add(SplitCompositeAtNode(composite, splitNode));
    }

    return splitNodes;
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Node>>> SectionSplitter::FindChildSplitPositions(
    System::SharedPtr<CompositeNode> node)
{
    // A node may span across multiple pages, so a list of split positions is returned.
    // The split node is the first node on the next page.
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> splitList =
        System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Node>>>();

    int32_t startingPage = pageNumberFinder->GetPage(node);

    System::ArrayPtr<System::SharedPtr<Node>> childNodes = node->get_NodeType() == NodeType::Section
                                                               ? (System::DynamicCast<Section>(node))->get_Body()->get_ChildNodes()->ToArray()
                                                               : node->get_ChildNodes()->ToArray();
    for (System::SharedPtr<Node> childNode : childNodes)
    {
        int32_t pageNum = pageNumberFinder->GetPage(childNode);

        if (System::ObjectExt::Is<Run>(childNode))
        {
            pageNum = pageNumberFinder->GetPageEnd(childNode);
        }

        // If the page of the child node has changed, then this is the split position.
        // Add this to the list.
        if (pageNum > startingPage)
        {
            splitList->Add(childNode);
            startingPage = pageNum;
        }

        if (pageNumberFinder->PageSpan(childNode) > 1)
        {
            pageNumberFinder->AddPageNumbersForNode(childNode, pageNum, pageNum);
        }
    }

    // Split composites backward, so the cloned nodes are inserted in the right order.
    splitList->Reverse();
    return splitList;
}

System::SharedPtr<CompositeNode> SectionSplitter::SplitCompositeAtNode(System::SharedPtr<CompositeNode> baseNode, System::SharedPtr<Node> targetNode)
{
    auto cloneNode = System::DynamicCast<CompositeNode>(baseNode->Clone(false));
    System::SharedPtr<Node> node = targetNode;
    int32_t currentPageNum = pageNumberFinder->GetPage(baseNode);

    // Move all nodes found on the next page into the copied node. Handle row nodes separately.
    if (baseNode->get_NodeType() != NodeType::Row)
    {
        System::SharedPtr<CompositeNode> composite = cloneNode;
        if (baseNode->get_NodeType() == NodeType::Section)
        {
            cloneNode = System::DynamicCast<CompositeNode>(baseNode->Clone(true));
            auto section = System::DynamicCast<Section>(cloneNode);
            section->get_Body()->RemoveAllChildren();
            composite = section->get_Body();
        }

        while (node != nullptr)
        {
            System::SharedPtr<Node> nextNode = node->get_NextSibling();
            composite->AppendChild(node);
            node = nextNode;
        }
    }
    else
    {
        // If we are dealing with a row, we need to add dummy cells for the cloned row.
        int32_t targetPageNum = pageNumberFinder->GetPage(targetNode);

        System::ArrayPtr<System::SharedPtr<Node>> childNodes = baseNode->get_ChildNodes()->ToArray();
        for (System::SharedPtr<Node> childNode : childNodes)
        {
            int32_t pageNum = pageNumberFinder->GetPage(childNode);
            if (pageNum == targetPageNum)
            {
                if (cloneNode->get_NodeType() == NodeType::Row)
                {
                    (System::DynamicCast<Row>(cloneNode))->EnsureMinimum();
                }

                if (cloneNode->get_NodeType() == NodeType::Cell)
                {
                    (System::DynamicCast<Cell>(cloneNode))->EnsureMinimum();
                }

                cloneNode->get_LastChild()->Remove();
                cloneNode->AppendChild(childNode);
            }
            else if (pageNum == currentPageNum)
            {
                cloneNode->AppendChild(childNode->Clone(false));
                if (cloneNode->get_LastChild()->get_NodeType() != NodeType::Cell)
                {
                    (System::DynamicCast<CompositeNode>(cloneNode->get_LastChild()))
                        ->AppendChild((System::DynamicCast<CompositeNode>(childNode))->get_FirstChild()->Clone(false));
                }
            }
        }
    }

    // Insert the split node after the original.
    baseNode->get_ParentNode()->InsertAfter(cloneNode, baseNode);

    // Update the new page numbers of the base node and the cloned node, including its descendants.
    // This will only be a single page as the cloned composite is split to be on one page.
    int32_t currentEndPageNum = pageNumberFinder->GetPageEnd(baseNode);
    pageNumberFinder->AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
    pageNumberFinder->AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);
    for (const auto& childNode : System::IterateOver(cloneNode->GetChildNodes(NodeType::Any, true)))
    {
        pageNumberFinder->AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);
    }

    return cloneNode;
}

const System::String SplitPageBreakCorrector::PageBreakStr = u"\f";
const char16_t SplitPageBreakCorrector::PageBreak = u'\f';

}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents
