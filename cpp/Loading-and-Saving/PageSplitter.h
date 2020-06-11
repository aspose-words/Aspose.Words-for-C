#pragma once

#include <cstdint>
#include <vector>
#include <unordered_map>

#include <system/shared_ptr.h>
#include <system/collections/dictionary.h>

#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { namespace Layout { class LayoutCollector; } } }
namespace Aspose { namespace Words { class Node; } }
namespace Aspose { namespace Words { class Paragraph; } }
namespace Aspose { namespace Words { class Run; } }

/// <summary>
/// Provides methods for extracting nodes of a document which are rendered on a specified pages.
/// </summary>
class PageNumberFinder
{
public:
	System::SharedPtr<Aspose::Words::Document> GetDocument() const;
	PageNumberFinder(const System::SharedPtr<Aspose::Words::Layout::LayoutCollector>& collector);
	int32_t GetPage(const System::SharedPtr<Aspose::Words::Node>& node) const;
	int32_t GetPageEnd(const System::SharedPtr<Aspose::Words::Node>& node) const;
	int32_t PageSpan(const System::SharedPtr<Aspose::Words::Node>& node) const;
	std::vector<System::SharedPtr<Aspose::Words::Node>> RetrieveAllNodesOnPages(int32_t startPage, int32_t endPage, Aspose::Words::NodeType nodeType);
	void SplitNodesAcrossPages();
	void AddPageNumbersForNode(const System::SharedPtr<Aspose::Words::Node>& node, int32_t startPage, int32_t endPage);

private:
	System::SharedPtr<System::Collections::Generic::Dictionary<System::SharedPtr<Aspose::Words::Node>, int32_t>> nodeStartPageLookup =
		System::MakeObject<System::Collections::Generic::Dictionary<System::SharedPtr<Aspose::Words::Node>, int32_t>>();
	System::SharedPtr<System::Collections::Generic::Dictionary<System::SharedPtr<Aspose::Words::Node>, int32_t>> nodeEndPageLookup =
		System::MakeObject<System::Collections::Generic::Dictionary<System::SharedPtr<Aspose::Words::Node>, int32_t>>();
	System::SharedPtr<Aspose::Words::Layout::LayoutCollector> collector;
	std::unordered_multimap<int32_t, System::SharedPtr<Aspose::Words::Node>> reversePageLookup;

	void CheckPageListsPopulated();
	void SplitRunsByWords(const System::SharedPtr<Aspose::Words::Paragraph>& paragraph);
	void SplitRunByWords(const System::SharedPtr<Aspose::Words::Run>& run);
	void ClearCollector();
};


/// <summary>
/// Splits a document into multiple documents, one per page.
/// </summary>
class DocumentPageSplitter
{
public:
	DocumentPageSplitter(const System::SharedPtr<Aspose::Words::Document>& source);
	System::SharedPtr<Aspose::Words::Document> GetDocumentOfPage(int32_t pageIndex);
	System::SharedPtr<Aspose::Words::Document> GetDocumentOfPageRange(int32_t startIndex, int32_t endIndex);

private:
	PageNumberFinder pageNumberFinder;
};


