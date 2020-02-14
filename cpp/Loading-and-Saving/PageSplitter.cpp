#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionStart.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Tables;

namespace
{
    using TNodePtr = System::SharedPtr<Node>;

    class DocumentPageSplitter;
    class PageNumberFinder;
    class SectionSplitter;

    bool IsHeaderFooterType(const TNodePtr& node)
    {
        return node->get_NodeType() == NodeType::HeaderFooter || node->GetAncestor(NodeType::HeaderFooter) != nullptr;
    }

    /// <summary>
    /// Splits text of the specified run into two runs.
    /// Inserts the new run just after the specified run.
    /// </summary>
    System::SharedPtr<Run> SplitRun(const System::SharedPtr<Run>& run, int32_t position)
    {
        // TODO (std_string) : fix using of overloaded members defined in the base classes
        auto afterRun = System::DynamicCast<Run>(System::StaticCast<Node>(run)->Clone(true));
        afterRun->set_Text(run->get_Text().Substring(position));
        run->set_Text(run->get_Text().Substring(0, position));
        run->get_ParentNode()->InsertAfter(afterRun, run);
        return afterRun;
    }

    /// <summary>
    /// Provides methods for extracting nodes of a document which are rendered on a specified pages.
    /// </summary>
    class PageNumberFinder
    {
    public:
        System::SharedPtr<Document> GetDocument() const;
        PageNumberFinder(const System::SharedPtr<Layout::LayoutCollector>& collector);
        int32_t GetPage(const TNodePtr& node) const;
        int32_t GetPageEnd(const TNodePtr& node) const;
        int32_t PageSpan(const TNodePtr& node) const;
        std::vector<TNodePtr> RetrieveAllNodesOnPages(int32_t startPage, int32_t endPage, NodeType nodeType);
        void SplitNodesAcrossPages();
        void AddPageNumbersForNode(const TNodePtr& node, int32_t startPage, int32_t endPage);

    private:
        System::SharedPtr<System::Collections::Generic::Dictionary<TNodePtr, int32_t>> nodeStartPageLookup =
            System::MakeObject<System::Collections::Generic::Dictionary<TNodePtr, int32_t>>();
        System::SharedPtr<System::Collections::Generic::Dictionary<TNodePtr, int32_t>> nodeEndPageLookup =
            System::MakeObject<System::Collections::Generic::Dictionary<TNodePtr, int32_t>>();
        System::SharedPtr<Layout::LayoutCollector> collector;
        std::unordered_multimap<int32_t, TNodePtr> reversePageLookup;

        void CheckPageListsPopulated();
        void SplitRunsByWords(const System::SharedPtr<Paragraph>& paragraph);
        void SplitRunByWords(const System::SharedPtr<Run>& run);
        void ClearCollector();
    };

    System::SharedPtr<Document> PageNumberFinder::GetDocument() const
    {
        return collector->get_Document();
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="PageNumberFinder"/> class.
    /// </summary>
    /// <param name="collector">A collector instance which has layout model records for the document.</param>
    PageNumberFinder::PageNumberFinder(const System::SharedPtr<Layout::LayoutCollector>& collector)
        : collector(collector) {}

    /// <summary>
    /// Retrieves 1-based index of a page that the node begins on.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    int32_t PageNumberFinder::GetPage(const TNodePtr& node) const
    {
        auto it = nodeStartPageLookup->data().find(node);
        return it == nodeStartPageLookup->data().cend() ? collector->GetStartPageIndex(node) : it->second;
    }

    /// <summary>
    /// Retrieves 1-based index of a page that the node ends on.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    int32_t PageNumberFinder::GetPageEnd(const TNodePtr& node) const
    {
        auto it = nodeEndPageLookup->data().find(node);
        return it == nodeEndPageLookup->data().cend() ? collector->GetEndPageIndex(node) : it->second;
    }

    /// <summary>
    /// Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    int32_t PageNumberFinder::PageSpan(const TNodePtr& node) const
    {
        return GetPageEnd(node) - GetPage(node) + 1;
    }

    /// <summary>
    /// Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
    /// </summary>
    /// <param name="startPage">
    /// The start Page.
    /// </param>
    /// <param name="endPage">
    /// The end Page.
    /// </param>
    /// <param name="nodeType">
    /// The node Type.
    /// </param>
    /// <returns>
    /// The <see cref="IList"/>.
    /// </returns>
    std::vector<TNodePtr> PageNumberFinder::RetrieveAllNodesOnPages(int32_t startPage, int32_t endPage, NodeType nodeType)
    {
        auto document = GetDocument();
        if (startPage < 1 || startPage > document->get_PageCount())
        {
            throw System::InvalidOperationException(u"'startPage' is out of range");
        }

        if (endPage < 1 || endPage > document->get_PageCount() || endPage < startPage)
        {
            throw System::InvalidOperationException(u"'endPage' is out of range");
        }

        CheckPageListsPopulated();
        std::vector<TNodePtr> pageNodes;
        for (int32_t page = startPage; page <= endPage; page++)
        {
            auto iterators = reversePageLookup.equal_range(page);
            // Some pages can be empty.
            if (iterators.first == reversePageLookup.end() && iterators.second == reversePageLookup.end())
            {
                continue;
            }

            for (auto iterator = iterators.first; iterator != iterators.second; ++iterator)
            {
                TNodePtr node = iterator->second;
                if (node->get_ParentNode() != nullptr &&
                    (nodeType == NodeType::Any || node->get_NodeType() == nodeType) &&
                    std::find(pageNodes.cbegin(), pageNodes.cend(), node) == pageNodes.cend())
                {
                    pageNodes.push_back(node);
                }
            }
        }

        return pageNodes;
    }

    /// <summary>
    /// Splits nodes which appear over two or more pages into separate nodes so that they still appear in the same way
    /// but no longer appear across a page.
    /// </summary>
    void PageNumberFinder::SplitNodesAcrossPages()
    {
        auto document = GetDocument();
        for (const System::SharedPtr<Paragraph>& paragraph : System::IterateOver<Paragraph>(document->GetChildNodes(NodeType::Paragraph, true)))
        {
            if (GetPage(paragraph) != GetPageEnd(paragraph))
            {
                SplitRunsByWords(paragraph);
            }
        }

        ClearCollector();

        // Visit any composites which are possibly split across pages and split them into separate nodes.
        document->Accept(System::DynamicCast<DocumentVisitor>(System::MakeObject<SectionSplitter>(*this)));
    }

    void PageNumberFinder::AddPageNumbersForNode(const TNodePtr& node, int32_t startPage, int32_t endPage)
    {
        if (startPage > 0)
        {
            nodeStartPageLookup->idx_set(node, startPage);
        }

        if (endPage > 0)
        {
            nodeEndPageLookup->idx_set(node, endPage);
        }
    }

    void PageNumberFinder::CheckPageListsPopulated()
    {
        if (!reversePageLookup.empty())
        {
            return;
        }

        // Add each node to a list which represent the nodes found on each page.
        for (const TNodePtr& node : System::IterateOver(GetDocument()->GetChildNodes(NodeType::Any, true)))
        {
            // Headers/Footers follow sections. They are not split by themselves.
            if (IsHeaderFooterType(node))
            {
                continue;
            }

            int32_t startPage = GetPage(node);
            int32_t endPage = GetPageEnd(node);
            for (int32_t page = startPage; page <= endPage; page++)
            {
                reversePageLookup.insert({ page, node });
            }
        }
    }

    void PageNumberFinder::SplitRunsByWords(const System::SharedPtr<Paragraph>& paragraph)
    {
        for (const System::SharedPtr<Run>& run : System::IterateOver<Run>(paragraph->get_Runs()))
        {
            if (GetPage(run) == GetPageEnd(run))
            {
                continue;
            }
            SplitRunByWords(run);
        }
    }

    void PageNumberFinder::SplitRunByWords(const System::SharedPtr<Run>& run)
    {
        auto words = run->get_Text().Split(u' ');
        const auto rbit = words->crbegin();
        const auto reit = words->crend();

        for (auto it = rbit; it != reit; ++it)
        {
            int32_t pos = run->get_Text().get_Length() - it->get_Length() - 1;
            if (pos > 1)
            {
                SplitRun(run, run->get_Text().get_Length() - it->get_Length() - 1);
            }
        }
    }

    void PageNumberFinder::ClearCollector()
    {
        collector->Clear();
        GetDocument()->UpdatePageLayout();

        nodeStartPageLookup->Clear();
        nodeEndPageLookup->Clear();
    }

    PageNumberFinder CreatePageNumberFinder(const System::SharedPtr<Document>& document)
    {
        auto layoutCollector = System::MakeObject<Layout::LayoutCollector>(document);
        document->UpdatePageLayout();
        PageNumberFinder pageNumberFinder(layoutCollector);
        pageNumberFinder.SplitNodesAcrossPages();
        return pageNumberFinder;
    }

    const System::String pageBreakStr = u"\f";
    const char16_t pageBreak = u'\f';

    void RemovePageBreak(const System::SharedPtr<Run>& run)
    {
        auto paragraph = run->get_ParentParagraph();
        if (run->get_Text().Equals(pageBreakStr))
        {
            paragraph->RemoveChild(run);
        }
        else if (run->get_Text().EndsWith(pageBreakStr))
        {
            run->set_Text(run->get_Text().TrimEnd(pageBreak));
        }

        if (paragraph->get_ChildNodes()->get_Count() == 0)
        {
            System::SharedPtr<CompositeNode> parent = paragraph->get_ParentNode();
            parent->RemoveChild(paragraph);
        }
    }

    void CorrectPageBreak(System::SharedPtr<Section> section)
    {
        if (section->get_ChildNodes()->get_Count() == 0)
        {
            return;
        }

        System::SharedPtr<NodeCollection> childBodies = section->GetChildNodes(NodeType::Body, false);
        System::SharedPtr<Body> lastBody = System::DynamicCast<Body>((childBodies->get_Count() == 0 ? nullptr : childBodies->idx_get(childBodies->get_Count() - 1)));
        if (lastBody == nullptr)
        {
            return;
        }

        System::SharedPtr<Run> firstRun;
        for (System::SharedPtr<Run> run : System::IterateOver<Run>(lastBody->GetChildNodes(NodeType::Run, true)))
        {
            if (run->get_Text().EndsWith(pageBreakStr))
            {
                firstRun = run;
                break;
            }
        }

        if (firstRun != nullptr)
        {
            RemovePageBreak(firstRun);
        }
    }

    /// <summary>
    /// Splits a document into multiple sections so that each page begins and ends at a section boundary.
    /// </summary>
    class SectionSplitter : public DocumentVisitor
    {
        typedef SectionSplitter ThisType;
        typedef DocumentVisitor BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);

    public:
        SectionSplitter(PageNumberFinder& pageNumberFinder);
        VisitorAction VisitParagraphStart(System::SharedPtr<Paragraph> paragraph) override;
        VisitorAction VisitTableStart(System::SharedPtr<Tables::Table> table) override;
        VisitorAction VisitRowStart(System::SharedPtr<Tables::Row> row) override;
        VisitorAction VisitCellStart(System::SharedPtr<Tables::Cell> cell) override;
        VisitorAction VisitStructuredDocumentTagStart(System::SharedPtr<StructuredDocumentTag> sdt) override;
        VisitorAction VisitSmartTagStart(System::SharedPtr<SmartTag> smartTag) override;
        VisitorAction VisitSectionStart(System::SharedPtr<Section> section) override;
        VisitorAction VisitSmartTagEnd(System::SharedPtr<SmartTag> smartTag) override;
        VisitorAction VisitStructuredDocumentTagEnd(System::SharedPtr<StructuredDocumentTag> sdt) override;
        VisitorAction VisitCellEnd(System::SharedPtr<Tables::Cell> cell) override;
        VisitorAction VisitRowEnd(System::SharedPtr<Tables::Row> row) override;
        VisitorAction VisitTableEnd(System::SharedPtr<Tables::Table> table) override;
        VisitorAction VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph) override;
        VisitorAction VisitSectionEnd(System::SharedPtr<Section> section) override;

    private:
        PageNumberFinder& pageNumberFinder;
        VisitorAction ContinueIfCompositeAcrossPageElseSkip(const System::SharedPtr<CompositeNode>& composite);
        std::vector<TNodePtr> SplitComposite(const System::SharedPtr<CompositeNode>& composite);
        std::vector<TNodePtr> FindChildSplitPositions(const System::SharedPtr<CompositeNode>& node);
        System::SharedPtr<CompositeNode> SplitCompositeAtNode(const System::SharedPtr<CompositeNode>& baseNode, const TNodePtr& targetNode);
    };

    SectionSplitter::SectionSplitter(PageNumberFinder& pageNumberFinder)
        : pageNumberFinder(pageNumberFinder) { }

    VisitorAction SectionSplitter::VisitParagraphStart(System::SharedPtr<Paragraph> paragraph)
    {
        return ContinueIfCompositeAcrossPageElseSkip(paragraph);
    }

    VisitorAction SectionSplitter::VisitTableStart(System::SharedPtr<Tables::Table> table)
    {
        return ContinueIfCompositeAcrossPageElseSkip(table);
    }

    VisitorAction SectionSplitter::VisitRowStart(System::SharedPtr<Tables::Row> row)
    {
        return ContinueIfCompositeAcrossPageElseSkip(row);
    }

    VisitorAction SectionSplitter::VisitCellStart(System::SharedPtr<Tables::Cell> cell)
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

        // If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an 
        // extracted document if the previous section is missing.
        if (previousSection != nullptr)
        {
            System::SharedPtr<HeaderFooterCollection> previousHeaderFooters = previousSection->get_HeadersFooters();
            if (!section->get_PageSetup()->get_RestartPageNumbering())
            {
                section->get_PageSetup()->set_RestartPageNumbering(true);
                section->get_PageSetup()->set_PageStartingNumber(
                    previousSection->get_PageSetup()->get_PageStartingNumber() + pageNumberFinder.PageSpan(previousSection));
            }

            for (const System::SharedPtr<HeaderFooter>& previousHeaderFooter : System::IterateOver<HeaderFooter>(previousHeaderFooters))
            {
                if (section->get_HeadersFooters()->idx_get(previousHeaderFooter->get_HeaderFooterType()) == nullptr)
                {
                    // TODO (std_string) : fix using of overloaded members defined in the base classes
                    System::SharedPtr<HeaderFooter> newHeaderFooter = System::DynamicCast<HeaderFooter>(System::StaticCast<Node>(previousHeaderFooters->idx_get(previousHeaderFooter->get_HeaderFooterType()))->Clone(true));
                    section->get_HeadersFooters()->Add(newHeaderFooter);
                }
            }
        }

        return this->ContinueIfCompositeAcrossPageElseSkip(section);
    }

    VisitorAction SectionSplitter::VisitSmartTagEnd(System::SharedPtr<SmartTag> smartTag)
    {
        this->SplitComposite(smartTag);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitStructuredDocumentTagEnd(System::SharedPtr<StructuredDocumentTag> sdt)
    {
        this->SplitComposite(sdt);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitCellEnd(System::SharedPtr<Tables::Cell> cell)
    {
        this->SplitComposite(cell);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitRowEnd(System::SharedPtr<Tables::Row> row)
    {
        this->SplitComposite(row);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitTableEnd(System::SharedPtr<Tables::Table> table)
    {
        this->SplitComposite(table);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph)
    {
        // If paragraph contains only section break, add fake run into 
        if (paragraph->get_IsEndOfSection() && paragraph->get_ChildNodes()->get_Count() == 1 && paragraph->get_ChildNodes()->idx_get(0)->GetText() == u"\f")
        {
            auto run = System::MakeObject<Run>(paragraph->get_Document());
            paragraph->AppendChild(run);
            int32_t currentEndPageNum = pageNumberFinder.GetPageEnd(paragraph);
            pageNumberFinder.AddPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
        }

        //for (System::SharedPtr<Paragraph> clonePara : System::IterateOver<Paragraph>(SplitComposite(paragraph)))
        for (const TNodePtr& cloneParaNode : SplitComposite(paragraph))
        {
            System::SharedPtr<Paragraph> clonePara = System::DynamicCast<Paragraph>(cloneParaNode);
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
        for (TNodePtr cloneSectionNode : SplitComposite(section))
        {
            System::SharedPtr<Section> cloneSection = System::DynamicCast<Section>(cloneSectionNode);
            cloneSection->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
            cloneSection->get_PageSetup()->set_RestartPageNumbering(true);
            cloneSection->get_PageSetup()->set_PageStartingNumber(section->get_PageSetup()->get_PageStartingNumber() + (section->get_Document()->IndexOf(cloneSection) - section->get_Document()->IndexOf(section)));
            cloneSection->get_PageSetup()->set_DifferentFirstPageHeaderFooter(false);
            // corrects page break on end of the section
            CorrectPageBreak(cloneSection);
        }

        // corrects page break on end of the section
        CorrectPageBreak(section);

        // Add new page numbering for the body of the section as well.
        pageNumberFinder.AddPageNumbersForNode(section->get_Body(), pageNumberFinder.GetPage(section), pageNumberFinder.GetPageEnd(section));
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::ContinueIfCompositeAcrossPageElseSkip(const System::SharedPtr<CompositeNode>& composite)
    {
        return (pageNumberFinder.PageSpan(composite) > 1) ? VisitorAction::Continue : VisitorAction::SkipThisNode;
    }

    std::vector<TNodePtr> SectionSplitter::SplitComposite(const System::SharedPtr<CompositeNode>& composite)
    {
        std::vector<TNodePtr> splitNodes;
        for (const TNodePtr& splitNode : FindChildSplitPositions(composite))
        {
            splitNodes.push_back(SplitCompositeAtNode(composite, splitNode));
        }

        return splitNodes;
    }

    std::vector<TNodePtr> SectionSplitter::FindChildSplitPositions(const System::SharedPtr<CompositeNode>& node)
    {
        // A node may span across multiple pages so a list of split positions is returned.
        // The split node is the first node on the next page.
        std::vector<TNodePtr> splitList;
        int32_t startingPage = pageNumberFinder.GetPage(node);
        System::ArrayPtr<TNodePtr> childNodes = node->get_NodeType() == NodeType::Section ?
            (System::DynamicCast<Section>(node))->get_Body()->get_ChildNodes()->ToArray() :
            node->get_ChildNodes()->ToArray();
        for (TNodePtr childNode : childNodes)
        {
            int32_t pageNum = pageNumberFinder.GetPage(childNode);

            if (System::ObjectExt::Is<Run>(childNode))
            {
                pageNum = pageNumberFinder.GetPageEnd(childNode);
            }

            // If the page of the child node has changed then this is the split position. Add
            // this to the list.
            if (pageNum > startingPage)
            {
                splitList.push_back(childNode);
                startingPage = pageNum;
            }

            if (pageNumberFinder.PageSpan(childNode) > 1)
            {
                pageNumberFinder.AddPageNumbersForNode(childNode, pageNum, pageNum);
            }
        }

        // Split composites backward so the cloned nodes are inserted in the right order.
        //splitList->Reverse();
        std::reverse(std::begin(splitList), std::end(splitList));
        return splitList;
    }

    System::SharedPtr<CompositeNode> SectionSplitter::SplitCompositeAtNode(const System::SharedPtr<CompositeNode>& baseNode, const TNodePtr& targetNode)
    {
        // TODO (std_string) : fix using of overloaded members defined in the base classes
        System::SharedPtr<CompositeNode> cloneNode = System::DynamicCast<CompositeNode>(System::StaticCast<Node>(baseNode)->Clone(false));
        TNodePtr node = targetNode;
        int32_t currentPageNum = pageNumberFinder.GetPage(baseNode);

        // Move all nodes found on the next page into the copied node. Handle row nodes separately.
        if (baseNode->get_NodeType() != NodeType::Row)
        {
            System::SharedPtr<CompositeNode> composite = cloneNode;
            if (baseNode->get_NodeType() == NodeType::Section)
            {
                // TODO (std_string) : fix using of overloaded members defined in the base classes
                cloneNode = System::DynamicCast<CompositeNode>((System::StaticCast<Node>(baseNode))->Clone(true));
                System::SharedPtr<Section> section = System::DynamicCast<Section>(cloneNode);
                section->get_Body()->RemoveAllChildren();
                composite = section->get_Body();
            }

            while (node != nullptr)
            {
                TNodePtr nextNode = node->get_NextSibling();
                composite->AppendChild(node);
                node = nextNode;
            }
        }
        else
        {
            // If we are dealing with a row then we need to add in dummy cells for the cloned row.
            int32_t targetPageNum = pageNumberFinder.GetPage(targetNode);
            for (const TNodePtr& childNode : baseNode->get_ChildNodes()->ToArray())
            {
                int32_t pageNum = pageNumberFinder.GetPage(childNode);
                if (pageNum == targetPageNum)
                {
                    if (cloneNode->get_NodeType() == NodeType::Row)
                        (System::DynamicCast<Tables::Row>(cloneNode))->EnsureMinimum();

                    if (cloneNode->get_NodeType() == NodeType::Cell)
                        (System::DynamicCast<Tables::Cell>(cloneNode))->EnsureMinimum();

                    cloneNode->get_LastChild()->Remove();
                    cloneNode->AppendChild(childNode);
                }
                else if (pageNum == currentPageNum)
                {
                    cloneNode->AppendChild(childNode->Clone(false));
                    if (cloneNode->get_LastChild()->get_NodeType() != NodeType::Cell)
                    {
                        (System::DynamicCast<CompositeNode>(cloneNode->get_LastChild()))->AppendChild((System::DynamicCast<CompositeNode>(childNode))->get_FirstChild()->Clone(false));
                    }
                }
            }
        }

        // Insert the split node after the original.
        baseNode->get_ParentNode()->InsertAfter(cloneNode, baseNode);

        // Update the new page numbers of the base node and the clone node including its descendents.
        // This will only be a single page as the cloned composite is split to be on one page.
        int32_t currentEndPageNum = pageNumberFinder.GetPageEnd(baseNode);
        pageNumberFinder.AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
        pageNumberFinder.AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);
        for (const TNodePtr& childNode : System::IterateOver(cloneNode->GetChildNodes(NodeType::Any, true)))
        {
            pageNumberFinder.AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);
        }

        return cloneNode;
    }

    /// <summary>
    /// Splits a document into multiple documents, one per page.
    /// </summary>
    class DocumentPageSplitter
    {
    public:
        DocumentPageSplitter(const System::SharedPtr<Document>& source);
        System::SharedPtr<Document> GetDocumentOfPage(int32_t pageIndex);
        System::SharedPtr<Document> GetDocumentOfPageRange(int32_t startIndex, int32_t endIndex);

    private:
        PageNumberFinder pageNumberFinder;
    };

    DocumentPageSplitter::DocumentPageSplitter(const System::SharedPtr<Document>& source)
        : pageNumberFinder(CreatePageNumberFinder(source)) {}

    System::SharedPtr<Document> DocumentPageSplitter::GetDocumentOfPage(int32_t pageIndex)
    {
        return GetDocumentOfPageRange(pageIndex, pageIndex);
    }

    System::SharedPtr<Document> DocumentPageSplitter::GetDocumentOfPageRange(int32_t startIndex, int32_t endIndex)
    {
        System::SharedPtr<Document> result = System::DynamicCast<Document>(System::StaticCast<Node>(pageNumberFinder.GetDocument())->Clone(false));
        for (const TNodePtr& section : pageNumberFinder.RetrieveAllNodesOnPages(startIndex, endIndex, NodeType::Section))
        {
            result->AppendChild(result->ImportNode(section, true));
        }
        return result;
    }


    void SplitDocumentToPages(System::String const &docName, System::String const &outputDataDir)
    {
        System::String fileName = System::IO::Path::GetFileNameWithoutExtension(docName);
        System::String extensionName = System::IO::Path::GetExtension(docName);

        std::cout << "Processing document: " << fileName.ToUtf8String() << extensionName.ToUtf8String() << std::endl;

        System::SharedPtr<Document> doc = System::MakeObject<Document>(docName);

        // Split nodes in the document into separate pages.
        DocumentPageSplitter splitter(doc);

        // Save each page to the disk as a separate document.
        for (int32_t page = 1; page <= doc->get_PageCount(); page++)
        {
            System::SharedPtr<Document> pageDoc = splitter.GetDocumentOfPage(page);
            pageDoc->Save(System::IO::Path::Combine(outputDataDir, System::String::Format(u"{0} - page{1} Out{2}", fileName, page, extensionName)));
        }
    }

    void SplitAllDocumentsToPages(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        System::ArrayPtr<System::String> fileNames = System::IO::Directory::GetFiles(inputDataDir, u"*.doc", System::IO::SearchOption::TopDirectoryOnly);
        for (const System::String& fileName : fileNames)
        {
            SplitDocumentToPages(fileName, outputDataDir);
        }
    }
}

void PageSplitter()
{
    std::cout << "PageSplitter example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving() + u"Split";
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving() + u"Split";
    SplitAllDocumentsToPages(inputDataDir, outputDataDir);
    std::cout << "Document split to pages successfully." << std::endl << "Files saved at " << outputDataDir.ToUtf8String() << std::endl;
    std::cout << "PageSplitter example finished." << std::endl << std::endl;
}