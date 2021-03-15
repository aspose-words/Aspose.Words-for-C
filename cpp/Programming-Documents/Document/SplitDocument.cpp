#include "stdafx.h"
#include "examples.h"

#include <system/convert.h>
#include <system/io/file_system_info.h>
#include <system/io/directory_info.h>
#include <system/io/directory.h>

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>


#include "Loading-and-Saving/PageSplitter.h"

using namespace Aspose::Words;

namespace 
{
void SplitDocumentBySections(const System::String& inputDataDir, const System::String& outputDataDir)
{
	
	//ExStart:SplitDocumentBySections
	// Open a Word document
	auto doc = System::MakeObject<Document>(inputDataDir + u"TestFile (Split).docx");

	for (int i = 0; i < doc->get_Sections()->get_Count(); i++)
	{
		// Split a document into smaller parts, in this instance split by section
		auto section = doc->get_Sections()->idx_get(i)->Clone();

		auto newDoc = System::MakeObject<Document>();
		newDoc->get_Sections()->Clear();

		auto newSection = System::StaticCast<Section>(newDoc->ImportNode(section, true));
		newDoc->get_Sections()->Add(newSection);

		// Save each section as a separate document
		newDoc->Save(outputDataDir + u"SplitDocumentBySectionsOut_" + System::Convert::ToString(i) + u".docx");
	}
	//ExEnd:SplitDocumentBySections
}

void MergeDocuments(const System::String& inputDataDir, const System::String& outputDataDir)
{
	//ExStart:MergeSplitDocuments
	// Create a new resulting document
	auto mergedDoc = System::MakeObject<Document>();
	auto mergedDocBuilder = System::MakeObject<DocumentBuilder>(mergedDoc);

	// Merge document parts one by one
	for (int idx = 1; idx <= 13; idx++)
	{
		System::String documentPath = inputDataDir + u"SplitDocumentPageByPageOut_" + System::Convert::ToString(idx) + u".docx";
		auto sourceDoc = System::MakeObject<Document>(documentPath);

		mergedDocBuilder->MoveToDocumentEnd();
		mergedDocBuilder->InsertDocument(sourceDoc, ImportFormatMode::KeepSourceFormatting);
	}

	mergedDoc->Save(outputDataDir + u"MergeDocuments_out.docx");
	//ExEnd:MergeSplitDocuments
}

void SplitDocumentPageByPage(const System::String& inputDataDir, const System::String& outputDataDir)
{
	//ExStart:SplitDocumentPageByPage
	// Open a Word document
	auto doc = System::MakeObject<Document>(inputDataDir + u"TestFile (Split).docx");

	// Split nodes in the document into separate pages
	DocumentPageSplitter splitter(doc);

	// Save each page as a separate document
	for (int page = 1; page <= doc->get_PageCount(); page++)
	{
		auto pageDoc = splitter.GetDocumentOfPage(page);
		pageDoc->Save(outputDataDir + u"SplitDocumentPageByPageOut_" + System::Convert::ToString(page) + u".docx");
	}
	//ExEnd:SplitDocumentPageByPage
}

void SplitDocumentByPageRange(const System::String& inputDataDir, const System::String& outputDataDir)
{
	//ExStart:SplitDocumentByPageRange
	// Open a Word document
	auto doc = System::MakeObject<Document>(inputDataDir + u"TestFile (Split).docx");

	// Split nodes in the document into separate pages
	DocumentPageSplitter splitter(doc);

	// Get part of the document
	auto pageDoc = splitter.GetDocumentOfPageRange(3, 6);
	pageDoc->Save(outputDataDir + u"SplitDocumentByPageRangeOut.docx");
	//ExEnd:SplitDocumentByPageRange
}

void SplitByHeadingsHtml(const System::String& inputDataDir, const System::String& outputDataDir)
{
    //ExStart:SplitDocumentByHeadingsHtml
    auto doc = System::MakeObject<Document>(inputDataDir + u"Rendering.docx");
	auto options = System::MakeObject<Saving::HtmlSaveOptions>();

    // Split a document into smaller parts, in this instance split by heading.
	options->set_DocumentSplitCriteria(Saving::DocumentSplitCriteria::HeadingParagraph);

    doc->Save(outputDataDir + u"SplitDocument.ByHeadingsHtml.html", options);
    //ExEnd:SplitDocumentByHeadingsHtml
}

void SplitBySectionsHtml(const System::String& inputDataDir, const System::String& outputDataDir)
{
    auto doc = System::MakeObject<Document>(inputDataDir + u"Rendering.docx");
	auto options = System::MakeObject<Saving::HtmlSaveOptions>();

    //ExStart:SplitDocumentBySectionsHtml
	options->set_DocumentSplitCriteria(Saving::DocumentSplitCriteria::SectionBreak);
    //ExEnd:SplitDocumentByHeadingsHtml

    doc->Save(outputDataDir + u"SplitDocument.BySectionsHtml.html", options);
}
}


void SplitDocument()
{

    std::cout << "SplitDocument example started.\n";
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

	SplitDocumentBySections(inputDataDir, outputDataDir);
	SplitDocumentPageByPage(inputDataDir, outputDataDir);
	MergeDocuments(inputDataDir, outputDataDir);
	SplitDocumentByPageRange(inputDataDir, outputDataDir);
	SplitByHeadingsHtml(inputDataDir, outputDataDir);
	SplitBySectionsHtml(inputDataDir, outputDataDir);

    std::cout << "SplitDocument example finished.\n\n";
}
