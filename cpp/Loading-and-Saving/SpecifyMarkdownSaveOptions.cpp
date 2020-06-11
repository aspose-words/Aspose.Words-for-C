#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace 
{
    void SaveAsMarkdown(const System::String& outputDataDir)
    {
		// ExStart:SaveAsMD
		auto builder = System::MakeObject<DocumentBuilder>();
		builder->Writeln(u"Some text!");

		// specify MarkDownSaveOptions
		auto saveOptions = SaveOptions::CreateSaveOptions(SaveFormat::Markdown);

		builder->get_Document()->Save(outputDataDir + u"TestDocument.md", saveOptions);
		// ExEnd:SaveAsMD
    }

	void ExportIntoMarkdownWithTableContentAlignment(const System::String& outputDataDir)
    {
		// ExStart:ExportIntoMarkdownWithTableContentAlignment
		auto builder = System::MakeObject<DocumentBuilder>();

		// Create a new table with two cells.
		builder->InsertCell();
		builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
		builder->Write(u"Cell1");
		builder->InsertCell();
		builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
		builder->Write(u"Cell2");

		auto saveOptions = System::MakeObject<MarkdownSaveOptions>();
		// Makes all paragraphs inside table to be aligned to Left. 
		saveOptions->set_TableContentAlignment(TableContentAlignment::Left);
		builder->get_Document()->Save(outputDataDir  + u"left.md", saveOptions);

		// Makes all paragraphs inside table to be aligned to Right. 
		saveOptions->set_TableContentAlignment(TableContentAlignment::Right);
		builder->get_Document()->Save(outputDataDir  + u"right.md", saveOptions);

		// Makes all paragraphs inside table to be aligned to Center. 
		saveOptions->set_TableContentAlignment(TableContentAlignment::Center);
		builder->get_Document()->Save(outputDataDir  + u"center.md", saveOptions);

		// Makes all paragraphs inside table to be aligned automatically.
		// The alignment in this case will be taken from the first paragraph in corresponding table column.
		saveOptions->set_TableContentAlignment(TableContentAlignment::Auto);
		builder->get_Document()->Save(outputDataDir  + u"auto.md", saveOptions);
		// ExEnd:ExportIntoMarkdownWithTableContentAlignment
    }

}

void SpecifyMarkdownSaveOptions()
{
    std::cout << "SpecifyMarkdownSaveOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    SaveAsMarkdown(outputDataDir);
	ExportIntoMarkdownWithTableContentAlignment(outputDataDir);
	
    std::cout << "SpecifyMarkdownSaveOptions example finished." << std::endl << std::endl;
}
