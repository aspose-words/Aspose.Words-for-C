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
    typedef System::SharedPtr<Node> TNodePtr;

    class DocumentPageSplitter;
    class PageNumberFinder;
    class SectionSplitter;

    bool IsHeaderFooterType(TNodePtr node)
    {
        return node->get_NodeType() == NodeType::HeaderFooter || node->GetAncestor(NodeType::HeaderFooter) != nullptr;
    }

    System::SharedPtr<Run> SplitRun(System::SharedPtr<Run> run, int32_t position)
    {
        // TODO (std_string) : fix using of overloaded members defined in the base classes
        System::SharedPtr<Run> afterRun = System::DynamicCast<Run>(System::StaticCast<Node>(run)->Clone(true));
        afterRun->set_Text(run->get_Text().Substring(position));
        run->set_Text(run->get_Text().Substring(0, position));
        run->get_ParentNode()->InsertAfter(afterRun, run);
        return afterRun;
    }

    /// <summary>
    /// Provides methods for extracting nodes of a document which are rendered on a specified pages.
    /// </summary>
    class PageNumberFinder : public System::Object
    {
        typedef PageNumberFinder ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<Document> GetDocument();
        PageNumberFinder(System::SharedPtr<LayoutCollector> collector);
        int32_t GetPage(TNodePtr node);
        int32_t GetPageEnd(TNodePtr node);
        int32_t PageSpan(TNodePtr node);
        std::vector<TNodePtr> RetrieveAllNodesOnPages(int32_t startPage, int32_t endPage, NodeType nodeType);
        void SplitNodesAcrossPages();
        void AddPageNumbersForNode(TNodePtr node, int32_t startPage, int32_t endPage);

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        std::unordered_map<TNodePtr, int32_t> nodeStartPageLookup;
        std::unordered_map<TNodePtr, int32_t> nodeEndPageLookup;
        System::SharedPtr<LayoutCollector> collector;
        std::unordered_multimap<int32_t, TNodePtr> reversePageLookup;

        void CheckPageListsPopulated();
        void SplitRunsByWords(System::SharedPtr<Paragraph> paragraph);
        void SplitRunByWords(System::SharedPtr<Run> run);
        void ClearCollector();
    };

    RTTI_INFO_IMPL_HASH(3293460369u, PageNumberFinder, ThisTypeBaseTypesInfo);

    System::SharedPtr<Document> PageNumberFinder::GetDocument()
    {
        return this->collector->get_Document();
    }

    PageNumberFinder::PageNumberFinder(System::SharedPtr<LayoutCollector> collector)
    {
        this->collector = collector;
    }

    int32_t PageNumberFinder::GetPage(TNodePtr node)
    {
        std::unordered_map<TNodePtr, int32_t>::const_iterator iterator = nodeStartPageLookup.find(node);
        return iterator == nodeStartPageLookup.cend() ? this->collector->GetStartPageIndex(node) : iterator->second;
    }

    int32_t PageNumberFinder::GetPageEnd(TNodePtr node)
    {
        std::unordered_map<TNodePtr, int32_t>::const_iterator iterator = nodeEndPageLookup.find(node);
        return iterator == nodeEndPageLookup.cend() ? this->collector->GetEndPageIndex(node) : iterator->second;
    }

    int32_t PageNumberFinder::PageSpan(TNodePtr node)
    {
        return this->GetPageEnd(node) - this->GetPage(node) + 1;
    }

    std::vector<TNodePtr> PageNumberFinder::RetrieveAllNodesOnPages(int32_t startPage, int32_t endPage, NodeType nodeType)
    {
        if (startPage < 1 || startPage > this->GetDocument()->get_PageCount())
        {
            throw System::InvalidOperationException(u"'startPage' is out of range");
        }

        if (endPage < 1 || endPage > this->GetDocument()->get_PageCount() || endPage < startPage)
        {
            throw System::InvalidOperationException(u"'endPage' is out of range");
        }

        this->CheckPageListsPopulated();
        std::vector<TNodePtr> pageNodes;
        for (int32_t page = startPage; page <= endPage; page++)
        {
            auto iterators = reversePageLookup.equal_range(page);
            // Some pages can be empty.
            if (iterators.first == reversePageLookup.end() && iterators.second == reversePageLookup.end())
            {
                continue;
            }

            for (auto iterator = iterators.first; iterator != iterators.second; ++ iterator)
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

    void PageNumberFinder::SplitNodesAcrossPages()
    {
        for (System::SharedPtr<Paragraph> paragraph : System::IterateOver<Paragraph>(this->GetDocument()->GetChildNodes(NodeType::Paragraph, true)))
        {
            if (this->GetPage(paragraph) != this->GetPageEnd(paragraph))
            {
                this->SplitRunsByWords(paragraph);
            }
        }

        this->ClearCollector();

        // Visit any composites which are possibly split across pages and split them into separate nodes.
        this->GetDocument()->Accept(System::DynamicCast<DocumentVisitor>(System::MakeObject<SectionSplitter>(System::MakeSharedPtr(this))));
    }

    void PageNumberFinder::AddPageNumbersForNode(TNodePtr node, int32_t startPage, int32_t endPage)
    {
        if (startPage > 0)
        {
            nodeStartPageLookup.insert({node, startPage});
        }

        if (endPage > 0)
        {
            nodeEndPageLookup.insert({node, endPage});
        }
    }

    void PageNumberFinder::CheckPageListsPopulated()
    {
        if (!reversePageLookup.empty())
        {
            return;
        }

        // Add each node to a list which represent the nodes found on each page.
        for (TNodePtr node : System::IterateOver(this->GetDocument()->GetChildNodes(NodeType::Any, true)))
        {
            // Headers/Footers follow sections. They are not split by themselves.
            if (IsHeaderFooterType(node))
            {
                continue;
            }

            int32_t startPage = this->GetPage(node);
            int32_t endPage = this->GetPageEnd(node);
            for (int32_t page = startPage; page <= endPage; page++)
            {
                reversePageLookup.insert({page, node});
            }
        }
    }

    void PageNumberFinder::SplitRunsByWords(System::SharedPtr<Paragraph> paragraph)
    {
        for (System::SharedPtr<Run> run : System::IterateOver<Run>(paragraph->get_Runs()))
        {
            if (this->GetPage(run) == this->GetPageEnd(run))
            {
                continue;
            }
            this->SplitRunByWords(run);
        }
    }

    void PageNumberFinder::SplitRunByWords(System::SharedPtr<Run> run)
    {
        System::ArrayPtr<System::String> splitResult = run->get_Text().Split(u' ');
        std::vector<System::String> words(splitResult.begin(), splitResult.end());
        std::reverse(std::begin(words), std::end(words));

        for (System::String const &word : words)
        {
            int32_t pos = run->get_Text().get_Length() - word.get_Length() - 1;
            if (pos > 1)
            {
                SplitRun(run, run->get_Text().get_Length() - word.get_Length() - 1);
            }
        }
    }

    void PageNumberFinder::ClearCollector()
    {
        this->collector->Clear();
        this->GetDocument()->UpdatePageLayout();

        nodeStartPageLookup.clear();
        nodeEndPageLookup.clear();
    }

    System::Object::shared_members_type PageNumberFinder::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("PageNumberFinder::collector", this->collector);
        return result;
    }

    System::SharedPtr<PageNumberFinder> CreatePageNumberFinder(System::SharedPtr<Document> document)
    {
        System::SharedPtr<LayoutCollector> layoutCollector = System::MakeObject<LayoutCollector>(document);
        document->UpdatePageLayout();
        System::SharedPtr<PageNumberFinder> pageNumberFinder = System::MakeObject<PageNumberFinder>(layoutCollector);
        pageNumberFinder->SplitNodesAcrossPages();
        return pageNumberFinder;
    }

    const System::String pageBreakStr = u"\f";
    const char16_t pageBreak = u'\f';

    void RemovePageBreak(System::SharedPtr<Run> run)
    {
        System::SharedPtr<Paragraph> paragraph = run->get_ParentParagraph();
        if (System::ObjectExt::Equals(run->get_Text(), pageBreakStr))
        {
            paragraph->RemoveChild(run);
        }
        else if (run->get_Text().EndsWith(pageBreakStr))
        {
            run->set_Text(run->get_Text().TrimEnd(System::MakeArray<char16_t>({pageBreak})));
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
        RTTI_INFO_DECL();

    public:
        SectionSplitter(System::SharedPtr<PageNumberFinder> pageNumberFinder);
        virtual VisitorAction VisitParagraphStart(System::SharedPtr<Paragraph> paragraph);
        virtual VisitorAction VisitTableStart(System::SharedPtr<Table> table);
        virtual VisitorAction VisitRowStart(System::SharedPtr<Row> row);
        virtual VisitorAction VisitCellStart(System::SharedPtr<Cell> cell);
        virtual VisitorAction VisitStructuredDocumentTagStart(System::SharedPtr<StructuredDocumentTag> sdt);
        virtual VisitorAction VisitSmartTagStart(System::SharedPtr<SmartTag> smartTag);
        virtual VisitorAction VisitSectionStart(System::SharedPtr<Section> section);
        virtual VisitorAction VisitSmartTagEnd(System::SharedPtr<SmartTag> smartTag);
        virtual VisitorAction VisitStructuredDocumentTagEnd(System::SharedPtr<StructuredDocumentTag> sdt);
        virtual VisitorAction VisitCellEnd(System::SharedPtr<Cell> cell);
        virtual VisitorAction VisitRowEnd(System::SharedPtr<Row> row);
        virtual VisitorAction VisitTableEnd(System::SharedPtr<Table> table);
        virtual VisitorAction VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph);
        virtual VisitorAction VisitSectionEnd(System::SharedPtr<Section> section);

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        System::SharedPtr<PageNumberFinder> pageNumberFinder;
        VisitorAction ContinueIfCompositeAcrossPageElseSkip(System::SharedPtr<CompositeNode> composite);
        std::vector<TNodePtr> SplitComposite(System::SharedPtr<CompositeNode> composite);
        std::vector<TNodePtr> FindChildSplitPositions(System::SharedPtr<CompositeNode> node);
        System::SharedPtr<CompositeNode> SplitCompositeAtNode(System::SharedPtr<CompositeNode> baseNode, TNodePtr targetNode);
    };

    RTTI_INFO_IMPL_HASH(728794745u, SectionSplitter, ThisTypeBaseTypesInfo);

    SectionSplitter::SectionSplitter(System::SharedPtr<PageNumberFinder> pageNumberFinder)
    {
        this->pageNumberFinder = pageNumberFinder;
    }

    VisitorAction SectionSplitter::VisitParagraphStart(System::SharedPtr<Paragraph> paragraph)
    {
        return this->ContinueIfCompositeAcrossPageElseSkip(paragraph);
    }

    VisitorAction SectionSplitter::VisitTableStart(System::SharedPtr<Table> table)
    {
        return this->ContinueIfCompositeAcrossPageElseSkip(table);
    }

    VisitorAction SectionSplitter::VisitRowStart(System::SharedPtr<Row> row)
    {
        return this->ContinueIfCompositeAcrossPageElseSkip(row);
    }

    VisitorAction SectionSplitter::VisitCellStart(System::SharedPtr<Cell> cell)
    {
        return this->ContinueIfCompositeAcrossPageElseSkip(cell);
    }

    VisitorAction SectionSplitter::VisitStructuredDocumentTagStart(System::SharedPtr<StructuredDocumentTag> sdt)
    {
        return this->ContinueIfCompositeAcrossPageElseSkip(sdt);
    }

    VisitorAction SectionSplitter::VisitSmartTagStart(System::SharedPtr<SmartTag> smartTag)
    {
        return this->ContinueIfCompositeAcrossPageElseSkip(smartTag);
    }

    VisitorAction SectionSplitter::VisitSectionStart(System::SharedPtr<Section> section)
    {
        System::SharedPtr<Section> previousSection = System::DynamicCast<Section>(section->get_PreviousSibling());

        // If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an 
        // extracted document if the previous section is missing.
        if (previousSection != nullptr)
        {
            System::SharedPtr<HeaderFooterCollection> previousHeaderFooters = previousSection->get_HeadersFooters();
            if (!section->get_PageSetup()->get_RestartPageNumbering())
            {
                section->get_PageSetup()->set_RestartPageNumbering(true);
                section->get_PageSetup()->set_PageStartingNumber(previousSection->get_PageSetup()->get_PageStartingNumber() + this->pageNumberFinder->PageSpan(previousSection));
            }

            for (System::SharedPtr<HeaderFooter> previousHeaderFooter : System::IterateOver<HeaderFooter>(previousHeaderFooters))
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

    VisitorAction SectionSplitter::VisitCellEnd(System::SharedPtr<Cell> cell)
    {
        this->SplitComposite(cell);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitRowEnd(System::SharedPtr<Row> row)
    {
        this->SplitComposite(row);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitTableEnd(System::SharedPtr<Table> table)
    {
        this->SplitComposite(table);
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph)
    {
        // If paragraph contains only section break, add fake run into 
        if (paragraph->get_IsEndOfSection() && paragraph->get_ChildNodes()->get_Count() == 1 && paragraph->get_ChildNodes()->idx_get(0)->GetText() == u"\f")
        {
            System::SharedPtr<Run> run = System::MakeObject<Run>(paragraph->get_Document());
            paragraph->AppendChild(run);
            int32_t currentEndPageNum = this->pageNumberFinder->GetPageEnd(paragraph);
            this->pageNumberFinder->AddPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
        }

        //for (System::SharedPtr<Paragraph> clonePara : System::IterateOver<Paragraph>(SplitComposite(paragraph)))
        for (TNodePtr cloneParaNode : SplitComposite(paragraph))
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
        this->pageNumberFinder->AddPageNumbersForNode(section->get_Body(), this->pageNumberFinder->GetPage(section), this->pageNumberFinder->GetPageEnd(section));
        return VisitorAction::Continue;
    }

    VisitorAction SectionSplitter::ContinueIfCompositeAcrossPageElseSkip(System::SharedPtr<CompositeNode> composite)
    {
        return (this->pageNumberFinder->PageSpan(composite) > 1) ? VisitorAction::Continue : VisitorAction::SkipThisNode;
    }

    std::vector<TNodePtr> SectionSplitter::SplitComposite(System::SharedPtr<CompositeNode> composite)
    {
        std::vector<TNodePtr> splitNodes;
        for (TNodePtr splitNode : FindChildSplitPositions(composite))
        {
            splitNodes.push_back(SplitCompositeAtNode(composite, splitNode));
        }

        return splitNodes;
    }

    std::vector<TNodePtr> SectionSplitter::FindChildSplitPositions(System::SharedPtr<CompositeNode> node)
    {
        // A node may span across multiple pages so a list of split positions is returned.
        // The split node is the first node on the next page.
        std::vector<TNodePtr> splitList;
        int32_t startingPage = this->pageNumberFinder->GetPage(node);
        System::ArrayPtr<TNodePtr> childNodes = node->get_NodeType() == NodeType::Section ?
                                                (System::DynamicCast<Section>(node))->get_Body()->get_ChildNodes()->ToArray() :
                                                node->get_ChildNodes()->ToArray();
        for (TNodePtr childNode : childNodes)
        {
            int32_t pageNum = this->pageNumberFinder->GetPage(childNode);

            if (System::ObjectExt::Is<Run>(childNode))
            {
                pageNum = this->pageNumberFinder->GetPageEnd(childNode);
            }

            // If the page of the child node has changed then this is the split position. Add
            // this to the list.
            if (pageNum > startingPage)
            {
                splitList.push_back(childNode);
                startingPage = pageNum;
            }

            if (this->pageNumberFinder->PageSpan(childNode) > 1)
            {
                this->pageNumberFinder->AddPageNumbersForNode(childNode, pageNum, pageNum);
            }
        }

        // Split composites backward so the cloned nodes are inserted in the right order.
        //splitList->Reverse();
        std::reverse(std::begin(splitList), std::end(splitList));
        return splitList;
    }

    System::SharedPtr<CompositeNode> SectionSplitter::SplitCompositeAtNode(System::SharedPtr<CompositeNode> baseNode, TNodePtr targetNode)
    {
        // TODO (std_string) : fix using of overloaded members defined in the base classes
        System::SharedPtr<CompositeNode> cloneNode = System::DynamicCast<CompositeNode>(System::StaticCast<Node>(baseNode)->Clone(false));
        TNodePtr node = targetNode;
        int32_t currentPageNum = this->pageNumberFinder->GetPage(baseNode);

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
            int32_t targetPageNum = this->pageNumberFinder->GetPage(targetNode);
            for (TNodePtr childNode : baseNode->get_ChildNodes()->ToArray())
            {
                int32_t pageNum = this->pageNumberFinder->GetPage(childNode);
                if (pageNum == targetPageNum)
                {
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
        int32_t currentEndPageNum = this->pageNumberFinder->GetPageEnd(baseNode);
        this->pageNumberFinder->AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
        this->pageNumberFinder->AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);
        for (TNodePtr childNode : System::IterateOver(cloneNode->GetChildNodes(NodeType::Any, true)))
        {
            this->pageNumberFinder->AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);
        }

        return cloneNode;
    }

    System::Object::shared_members_type SectionSplitter::GetSharedMembers()
    {
        auto result = DocumentVisitor::GetSharedMembers();
        result.Add("SectionSplitter::pageNumberFinder", this->pageNumberFinder);
        return result;
    }

    /// <summary>
    /// Splits a document into multiple documents, one per page.
    /// </summary>
    class DocumentPageSplitter : public System::Object
    {
        typedef DocumentPageSplitter ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        DocumentPageSplitter(System::SharedPtr<Document> source);
        System::SharedPtr<Document> GetDocumentOfPage(int32_t pageIndex);
        System::SharedPtr<Document> GetDocumentOfPageRange(int32_t startIndex, int32_t endIndex);

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        System::SharedPtr<PageNumberFinder> pageNumberFinder;
    };

    RTTI_INFO_IMPL_HASH(2297209060u, DocumentPageSplitter, ThisTypeBaseTypesInfo);

    DocumentPageSplitter::DocumentPageSplitter(System::SharedPtr<Document> source)
    {
        //Self Reference protector
        System::Details::ThisProtector selfRefProtector(this);
        this->pageNumberFinder = CreatePageNumberFinder(source);
    }

    System::SharedPtr<Document> DocumentPageSplitter::GetDocumentOfPage(int32_t pageIndex)
    {
        return this->GetDocumentOfPageRange(pageIndex, pageIndex);
    }

    System::SharedPtr<Document> DocumentPageSplitter::GetDocumentOfPageRange(int32_t startIndex, int32_t endIndex)
    {
        System::SharedPtr<Document> result = System::DynamicCast<Document>(System::StaticCast<Node>(this->pageNumberFinder->GetDocument())->Clone(false));
        for (TNodePtr section : pageNumberFinder->RetrieveAllNodesOnPages(startIndex, endIndex, NodeType::Section))
        {
            result->AppendChild(result->ImportNode(section, true));
        }
        return result;
    }

    System::Object::shared_members_type DocumentPageSplitter::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("DocumentPageSplitter::pageNumberFinder", this->pageNumberFinder);
        return result;
    }

    void SplitDocumentToPages(System::String const &docName, System::String const &outputDataDir)
    {
        System::String fileName = System::IO::Path::GetFileNameWithoutExtension(docName);
        System::String extensionName = System::IO::Path::GetExtension(docName);

        std::cout << "Processing document: " << fileName.ToUtf8String() << extensionName.ToUtf8String() << std::endl;

        System::SharedPtr<Document> doc = System::MakeObject<Document>(docName);

        // Split nodes in the document into separate pages.
        System::SharedPtr<DocumentPageSplitter> splitter = System::MakeObject<DocumentPageSplitter>(doc);

        // Save each page to the disk as a separate document.
        for (int32_t page = 1; page <= doc->get_PageCount(); page++)
        {
            System::SharedPtr<Document> pageDoc = splitter->GetDocumentOfPage(page);
            pageDoc->Save(System::IO::Path::Combine(outputDataDir, System::String::Format(u"PageSplitter.SplitDocumentToPages.Page{0}{1}", page, extensionName)));
        }
    }

    void SplitAllDocumentsToPages(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        System::ArrayPtr<System::String> fileNames = System::IO::Directory::GetFiles(inputDataDir, u"*.doc?", System::IO::SearchOption::TopDirectoryOnly);
        for (System::String fileName : fileNames)
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