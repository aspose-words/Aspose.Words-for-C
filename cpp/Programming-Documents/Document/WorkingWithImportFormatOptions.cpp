#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>

using namespace Aspose::Words;

namespace 
{
	void SmartStyleBehavior(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:SmartStyleBehavior
		auto srcDoc = System::MakeObject<Document>(inputDataDir + u"source.docx");
		auto dstDoc = System::MakeObject<Document>(inputDataDir + u"destination.docx");

		auto builder = System::MakeObject<DocumentBuilder>(dstDoc);
		builder->MoveToDocumentEnd();
		builder->InsertBreak(BreakType::PageBreak);

		auto options = System::MakeObject<ImportFormatOptions>();
		options->set_SmartStyleBehavior(true);
		builder->InsertDocument(srcDoc, ImportFormatMode::UseDestinationStyles, options);
		// ExEnd:SmartStyleBehavior
	}

    void KeepSourceNumbering(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:KeepSourceNumbering
		auto srcDoc = System::MakeObject<Document>(inputDataDir + u"source.docx");
		auto dstDoc = System::MakeObject<Document>(inputDataDir + u"destination.docx");

		auto options = System::MakeObject<ImportFormatOptions>();
		// Keep source list formatting when importing numbered paragraphs.
		options->set_KeepSourceNumbering(true);

		auto importer = System::MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting, options);
		auto srcParas = srcDoc->get_FirstSection()->get_Body()->get_Paragraphs();
		for (const auto& srcPara : System::IterateOver<Paragraph>(srcParas))
		{
			auto importedNode = importer->ImportNode(srcPara, false);
			dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
		}

		dstDoc->Save(outputDataDir + u"WorkingWithImportFormatOptions.KeepSourceNumbering.docx");
		// ExEnd:KeepSourceNumbering
	}

    void IgnoreTextBoxes(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:IgnoreTextBoxes
		auto srcDoc = System::MakeObject<Document>(inputDataDir + u"source.docx");
		auto dstDoc = System::MakeObject<Document>(inputDataDir + u"destination.docx");
		
		auto options = System::MakeObject<ImportFormatOptions>();
		// Keep the source text boxes formatting when importing.
		options->set_IgnoreTextBoxes(false);
		auto importer = System::MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting, options);
		auto srcParas = srcDoc->get_FirstSection()->get_Body()->get_Paragraphs();

		for (const auto& srcPara : System::IterateOver<Paragraph>(srcParas))
		{
			auto importedNode = importer->ImportNode(srcPara, true);
			dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
		}

		dstDoc->Save(outputDataDir + u"WorkingWithImportFormatOptions.IgnoreTextBoxes.docx");
		// ExEnd:IgnoreTextBoxes
	}

    void IgnoreHeaderFooter(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:IgnoreHeaderFooter
		auto srcDoc = System::MakeObject<Document>(inputDataDir + u"source.docx");
		auto dstDoc = System::MakeObject<Document>(inputDataDir + u"destination.docx");
		
		auto options = System::MakeObject<ImportFormatOptions>();
		options->set_IgnoreHeaderFooter(false);

		dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting, options);
		dstDoc->Save(outputDataDir + u"WorkingWithImportFormatOptions.IgnoreHeaderFooter.docx");
	}

}

void WorkingWithImportFormatOptions()
{
    std::cout << "WorkingWithImportFormatOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

#if 0
	SmartStyleBehavior(inputDataDir,outputDataDir);
	KeepSourceNumbering(inputDataDir, outputDataDir);
	IgnoreTextBoxes(inputDataDir, outputDataDir);
#endif
	IgnoreHeaderFooter(inputDataDir, outputDataDir);

    std::cout << "WorkingWithImportFormatOptions example finished." << std::endl << std::endl;
}
