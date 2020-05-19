#include "stdafx.h"
#include "examples.h"

#include <drawing/color.h>
#include <drawing/image.h>

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/Watermark.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/TextWatermarkOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/ImageWatermarkOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/WatermarkLayout.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/WatermarkType.h>

using namespace Aspose::Words;

namespace 
{
	void AddTextWatermarkWithSpecificOptions(System::String const& inputDataDir, System::String const& outputDataDir)
	{
		// ExStart: AddTextWatermarkWithSpecificOptions
		auto doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

		auto options = System::MakeObject<TextWatermarkOptions>();
		options->set_FontFamily(u"Arial");
		options->set_FontSize(36);
		options->set_Color(System::Drawing::Color::get_Black());
		options->set_Layout(WatermarkLayout::Horizontal);
		options->set_IsSemitrasparent(false);

		doc->get_Watermark()->SetText(u"Test", options);

		auto outputPath = outputDataDir + u"AddTextWatermark.docx";

		doc->Save(outputPath);
		// ExEnd: AddTextWatermarkWithSpecificOptions

		std::cout << "\nDocument saved successfully.\nFile saved at " << outputPath.ToUtf8String() << '\n';
	}

	void AddImageWatermarkWithSpecificOptions(System::String const& inputDataDir, System::String const& outputDataDir)
	{
		// ExStart: AddImageWatermarkWithSpecificOptions
		auto doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

		auto options = System::MakeObject<ImageWatermarkOptions>();
		options->set_Scale(5);
		options->set_IsWashout(false);

		doc->get_Watermark()->SetImage(System::Drawing::Image::FromFile(inputDataDir + u"Watermark.png"), options);

		auto outputPath = outputDataDir + u"AddImageWatermark.docx";
		doc->Save(outputPath);
		// ExEnd: AddImageWatermarkWithSpecificOptions

		std::cout << "Document saved successfully.\nFile saved at " << outputPath.ToUtf8String() << '\n';
	}

	void RemoveWatermarkFromDocument(System::String const& inputDataDir, System::String const& outputDataDir)
	{
		// ExStart: RemoveWatermarkFromDocument
		auto doc = System::MakeObject<Document>(inputDataDir + u"TextWatermark.docx");
		if (doc->get_Watermark()->get_Type() == WatermarkType::Text)
		{
			doc->get_Watermark()->Remove();
		}

		auto outputPath = outputDataDir + u"RemoveWatermark.docx";
		doc->Save(outputPath);
		// ExEnd: RemoveWatermarkFromDocument
		std::cout << "Document saved successfully.\nFile saved at " << outputPath.ToUtf8String() << '\n';
	}

}

void WorkWithWatermark()
{
    std::cout << "WorkWithWatermark example started." << std::endl;

    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

	AddTextWatermarkWithSpecificOptions(inputDataDir, outputDataDir);
	AddImageWatermarkWithSpecificOptions(inputDataDir, outputDataDir);
	RemoveWatermarkFromDocument(inputDataDir, outputDataDir);

    std::cout << "WorkWithWatermark example finished." << std::endl << std::endl;
}
