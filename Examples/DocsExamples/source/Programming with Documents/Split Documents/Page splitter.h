#pragma once

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Markup/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <system/array.h>
#include <system/collections/dictionary.h>
#include <system/collections/idictionary.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ilist.h>
#include <system/collections/list.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/path.h>
#include <system/io/search_option.h>
#include <system/linq/enumerable.h>

#include "DocsExamplesBase.h"

namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents {
class PageNumberFinder;
}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents {

/// <summary>
/// Splits a document into multiple documents, one per page.
/// </summary>
class DocumentPageSplitter : public System::Object
{
public:
    /// <summary>
    /// Initializes a new instance of the <see cref="DocumentPageSplitter"></see> class.
    /// This method splits the document into sections so that each page begins and ends at a section boundary.
    /// It is recommended not to modify the document afterwards.
    /// </summary>
    /// <param name="source">Source document</param>
    DocumentPageSplitter(System::SharedPtr<Aspose::Words::Document> source);

    /// <summary>
    /// Gets the document of a page.
    /// </summary>
    /// <param name="pageIndex">
    /// 1-based index of a page.
    /// </param>
    /// <returns>
    /// The <see cref="Document"></see>.
    /// </returns>
    System::SharedPtr<Aspose::Words::Document> GetDocumentOfPage(int pageIndex);
    /// <summary>
    /// Gets the document of a page range.
    /// </summary>
    /// <param name="startIndex">
    /// 1-based index of the start page.
    /// </param>
    /// <param name="endIndex">
    /// 1-based index of the end page.
    /// </param>
    /// <returns>
    /// The <see cref="Document"></see>.
    /// </returns>
    System::SharedPtr<Aspose::Words::Document> GetDocumentOfPageRange(int startIndex, int endIndex);

private:
    System::SharedPtr<PageNumberFinder> pageNumberFinder;

    /// <summary>
    /// Gets the document this instance works with.
    /// </summary>
    System::SharedPtr<Aspose::Words::Document> get_Document();
};

class PageSplitter : public DocsExamplesBase
{
public:
    void SplitDocuments()
    {
        SplitAllDocumentsToPages(MyDir);
    }

    void SplitDocumentToPages(System::String docName)
    {
        System::String fileName = System::IO::Path::GetFileNameWithoutExtension(docName);
        System::String extensionName = System::IO::Path::GetExtension(docName);

        std::cout << (System::String(u"Processing document: ") + fileName + extensionName) << std::endl;

        auto doc = System::MakeObject<Document>(docName);

        // Split nodes in the document into separate pages.
        auto splitter = System::MakeObject<DocumentPageSplitter>(doc);

        // Save each page to the disk as a separate document.
        for (int page = 1; page <= doc->get_PageCount(); page++)
        {
            System::SharedPtr<Document> pageDoc = splitter->GetDocumentOfPage(page);
            pageDoc->Save(System::IO::Path::Combine(ArtifactsDir, System::String::Format(u"{0} - page{1} Out{2}", fileName, page, extensionName)));
        }
    }

    void SplitAllDocumentsToPages(System::String folderName)
    {
        System::SharedPtr<System::Collections::Generic::List<System::String>> fileNames =
            System::IO::Directory::GetFiles(folderName, u"*.doc", System::IO::SearchOption::TopDirectoryOnly)
                ->LINQ_Where(static_cast<System::Func<System::String, bool>>(
                    static_cast<std::function<bool(System::String item)>>([](System::String item) -> bool { return item.EndsWith(u".doc"); })))
                ->LINQ_ToList();

        for (const auto& fileName : fileNames)
        {
            SplitDocumentToPages(fileName);
        }
    }
};

/// <summary>
/// Splits a document into multiple sections so that each page begins and ends at a section boundary.
/// </summary>
class SectionSplitter : public DocumentVisitor
{
public:
    SectionSplitter(System::SharedPtr<PageNumberFinder> pageNumberFinder);

    VisitorAction VisitParagraphStart(System::SharedPtr<Paragraph> paragraph) override;
    VisitorAction VisitTableStart(System::SharedPtr<Table> table) override;
    VisitorAction VisitRowStart(System::SharedPtr<Row> row) override;
    VisitorAction VisitCellStart(System::SharedPtr<Cell> cell) override;
    VisitorAction VisitStructuredDocumentTagStart(System::SharedPtr<StructuredDocumentTag> sdt) override;
    VisitorAction VisitSmartTagStart(System::SharedPtr<SmartTag> smartTag) override;
    VisitorAction VisitSectionStart(System::SharedPtr<Section> section) override;
    VisitorAction VisitSmartTagEnd(System::SharedPtr<SmartTag> smartTag) override;
    VisitorAction VisitStructuredDocumentTagEnd(System::SharedPtr<StructuredDocumentTag> sdt) override;
    VisitorAction VisitCellEnd(System::SharedPtr<Cell> cell) override;
    VisitorAction VisitRowEnd(System::SharedPtr<Row> row) override;
    VisitorAction VisitTableEnd(System::SharedPtr<Table> table) override;
    VisitorAction VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph) override;
    VisitorAction VisitSectionEnd(System::SharedPtr<Section> section) override;

private:
    System::SharedPtr<PageNumberFinder> pageNumberFinder;

    VisitorAction ContinueIfCompositeAcrossPageElseSkip(System::SharedPtr<CompositeNode> composite);
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> SplitComposite(System::SharedPtr<CompositeNode> composite);
    System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Node>>> FindChildSplitPositions(System::SharedPtr<CompositeNode> node);
    System::SharedPtr<CompositeNode> SplitCompositeAtNode(System::SharedPtr<CompositeNode> baseNode, System::SharedPtr<Node> targetNode);
};

/// <summary>
/// Provides methods for extracting nodes of a document which are rendered on a specified pages.
/// </summary>
class PageNumberFinder : public System::Object
{
    friend class SectionSplitter;

public:
    // Maps node to a start/end page numbers.
    // This is used to override baseline page numbers provided by the collector when the document is split.
    System::SharedPtr<System::Collections::Generic::IDictionary<System::SharedPtr<Node>, int>> nodeStartPageLookup;
    System::SharedPtr<System::Collections::Generic::IDictionary<System::SharedPtr<Node>, int>> nodeEndPageLookup;
    System::SharedPtr<LayoutCollector> collector;

    // Maps page number to a list of nodes found on that page.
    System::SharedPtr<System::Collections::Generic::IDictionary<int, System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Node>>>>>
        reversePageLookup;

    /// <summary>
    /// Initializes a new instance of the <see cref="PageNumberFinder"></see> class.
    /// </summary>
    /// <param name="collector">A collector instance that has layout model records for the document.</param>
    PageNumberFinder(System::SharedPtr<LayoutCollector> collector)
        : nodeStartPageLookup(System::MakeObject<System::Collections::Generic::Dictionary<System::SharedPtr<Node>, int>>()),
          nodeEndPageLookup(System::MakeObject<System::Collections::Generic::Dictionary<System::SharedPtr<Node>, int>>())
    {
        this->collector = collector;
    }

    /// <summary>
    /// Gets the document this instance works with.
    /// </summary>
    System::SharedPtr<Aspose::Words::Document> get_Document()
    {
        return collector->get_Document();
    }

    /// <summary>
    /// Retrieves 1-based index of a page that the node begins on.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    int GetPage(System::SharedPtr<Node> node)
    {
        return nodeStartPageLookup->ContainsKey(node) ? nodeStartPageLookup->idx_get(node) : collector->GetStartPageIndex(node);
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
    int GetPageEnd(System::SharedPtr<Node> node)
    {
        return nodeEndPageLookup->ContainsKey(node) ? nodeEndPageLookup->idx_get(node) : collector->GetEndPageIndex(node);
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
    int PageSpan(System::SharedPtr<Node> node)
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
    /// The <see cref="IList{T}"></see>.
    /// </returns>
    System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Node>>> RetrieveAllNodesOnPages(int startPage, int endPage, NodeType nodeType)
    {
        if (startPage < 1 || startPage > get_Document()->get_PageCount())
        {
            throw System::InvalidOperationException(u"'startPage' is out of range");
        }

        if (endPage < 1 || endPage > get_Document()->get_PageCount() || endPage < startPage)
        {
            throw System::InvalidOperationException(u"'endPage' is out of range");
        }

        CheckPageListsPopulated();

        System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Node>>> pageNodes =
            System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Node>>>();
        for (int page = startPage; page <= endPage; page++)
        {
            // Some pages can be empty.
            if (!reversePageLookup->ContainsKey(page))
            {
                continue;
            }

            for (const auto& node : System::IterateOver(reversePageLookup->idx_get(page)))
            {
                if (node->get_ParentNode() != nullptr && (nodeType == NodeType::Any || node->get_NodeType() == nodeType) && !pageNodes->Contains(node))
                {
                    pageNodes->Add(node);
                }
            }
        }

        return pageNodes;
    }

    /// <summary>
    /// Splits nodes that appear over two or more pages into separate nodes so that they still appear in the same way
    /// but no longer appear across a page.
    /// </summary>
    void SplitNodesAcrossPages()
    {
        for (const auto& paragraph : System::IterateOver<Paragraph>(get_Document()->GetChildNodes(NodeType::Paragraph, true)))
        {
            if (GetPage(paragraph) != GetPageEnd(paragraph))
            {
                SplitRunsByWords(paragraph);
            }
        }

        ClearCollector();

        // Visit any composites which are possibly split across pages and split them into separate nodes.
        get_Document()->Accept(System::MakeObject<SectionSplitter>(System::MakeSharedPtr(this)));
    }

    /// <summary>
    /// This is called by <see cref="SectionSplitter"></see> to update page numbers of split nodes.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <param name="startPage">
    /// The start Page.
    /// </param>
    /// <param name="endPage">
    /// The end Page.
    /// </param>
    void AddPageNumbersForNode(System::SharedPtr<Node> node, int startPage, int endPage)
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

    bool IsHeaderFooterType(System::SharedPtr<Node> node)
    {
        return node->get_NodeType() == NodeType::HeaderFooter || node->GetAncestor(NodeType::HeaderFooter) != nullptr;
    }

    void CheckPageListsPopulated()
    {
        if (reversePageLookup != nullptr)
        {
            return;
        }

        reversePageLookup = System::MakeObject<
            System::Collections::Generic::Dictionary<int, System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Node>>>>>();

        // Add each node to a list that represent the nodes found on each page.
        for (const auto& node : System::IterateOver(get_Document()->GetChildNodes(NodeType::Any, true)))
        {
            // Headers/Footers follow sections and are not split by themselves.
            if (IsHeaderFooterType(node))
            {
                continue;
            }

            int startPage = GetPage(node);
            int endPage = GetPageEnd(node);
            for (int page = startPage; page <= endPage; page++)
            {
                if (!reversePageLookup->ContainsKey(page))
                {
                    reversePageLookup->Add(page, System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Node>>>());
                }

                reversePageLookup->idx_get(page)->Add(node);
            }
        }
    }

    void SplitRunsByWords(System::SharedPtr<Paragraph> paragraph)
    {
        for (const auto& run : System::IterateOver<Aspose::Words::Run>(paragraph->get_Runs()))
        {
            if (GetPage(run) == GetPageEnd(run))
            {
                continue;
            }

            SplitRunByWords(run);
        }
    }

    void SplitRunByWords(System::SharedPtr<Run> run)
    {
        System::SharedPtr<System::Collections::Generic::List<System::String>> words = run->get_Text().Split(System::MakeArray<char16_t>({u' '}))->LINQ_ToList();
        words->Reverse();

        for (const auto& word : words)
        {
            int pos = run->get_Text().get_Length() - word.get_Length() - 1;
            if (pos > 1)
            {
                SplitRun(run, run->get_Text().get_Length() - word.get_Length() - 1);
            }
        }
    }

    /// <summary>
    /// Splits text of the specified run into two runs.
    /// Inserts the new run just after the specified run.
    /// </summary>
    void SplitRun(System::SharedPtr<Run> run, int position)
    {
        auto afterRun = System::ExplicitCast<Aspose::Words::Run>(run->Clone(true));
        afterRun->set_Text(run->get_Text().Substring(position));
        run->set_Text(run->get_Text().Substring(0, position));
        run->get_ParentNode()->InsertAfter(afterRun, run);
    }

    void ClearCollector()
    {
        collector->Clear();
        get_Document()->UpdatePageLayout();

        nodeStartPageLookup->Clear();
        nodeEndPageLookup->Clear();
    }
    virtual ~PageNumberFinder()
    {
    }
};

class PageNumberFinderFactory : public System::Object
{
public:
    static System::SharedPtr<PageNumberFinder> Create(System::SharedPtr<Document> document)
    {
        auto layoutCollector = System::MakeObject<LayoutCollector>(document);
        document->UpdatePageLayout();
        auto pageNumberFinder = System::MakeObject<PageNumberFinder>(layoutCollector);
        pageNumberFinder->SplitNodesAcrossPages();
        return pageNumberFinder;
    }
};

class SplitPageBreakCorrector : public System::Object
{
public:
    static const System::String PageBreakStr;
    static const char16_t PageBreak;

    static void ProcessSection(System::SharedPtr<Section> section)
    {
        if (section->GetChildNodes(Aspose::Words::NodeType::Any, false)->get_Count() == 0)
        {
            return;
        }

        System::SharedPtr<Body> lastBody = section->GetChildNodes(Aspose::Words::NodeType::Any, false)->LINQ_OfType<System::SharedPtr<Body>>()->LINQ_LastOrDefault();

        System::SharedPtr<Run> run =
            lastBody != nullptr
                ? lastBody->GetChildNodes(NodeType::Run, true)
                      ->LINQ_OfType<System::SharedPtr<Run>>()
                      ->LINQ_FirstOrDefault(static_cast<System::Func<System::SharedPtr<Run>, bool>>(static_cast<std::function<bool(System::SharedPtr<Run> p)>>(
                          [](System::SharedPtr<Run> p) -> bool { return p->get_Text().EndsWith(PageBreakStr); })))
                : nullptr;

        if (run != nullptr)
        {
            RemovePageBreak(run);
        }
    }

    void RemovePageBreakFromParagraph(System::SharedPtr<Paragraph> paragraph)
    {
        auto run = System::ExplicitCast<Aspose::Words::Run>(paragraph->get_FirstChild());
        if (run->get_Text() == PageBreakStr)
        {
            paragraph->RemoveChild(run);
        }
    }

    void ProcessLastParagraph(System::SharedPtr<Paragraph> paragraph)
    {
        System::SharedPtr<Node> lastNode = paragraph->GetChildNodes(Aspose::Words::NodeType::Any, false)->idx_get(paragraph->GetChildNodes(Aspose::Words::NodeType::Any, false)->get_Count() - 1);
        if (lastNode->get_NodeType() != NodeType::Run)
        {
            return;
        }

        auto run = System::ExplicitCast<Aspose::Words::Run>(lastNode);
        RemovePageBreak(run);
    }

    static void RemovePageBreak(System::SharedPtr<Run> run)
    {
        System::SharedPtr<Paragraph> paragraph = run->get_ParentParagraph();

        if (run->get_Text() == PageBreakStr)
        {
            paragraph->RemoveChild(run);
        }
        else if (run->get_Text().EndsWith(PageBreakStr))
        {
            run->set_Text(run->get_Text().TrimEnd(System::MakeArray<char16_t>({PageBreak})));
        }

        if (paragraph->GetChildNodes(Aspose::Words::NodeType::Any, false)->get_Count() == 0)
        {
            System::SharedPtr<CompositeNode> parent = paragraph->get_ParentNode();
            parent->RemoveChild(paragraph);
        }
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents
