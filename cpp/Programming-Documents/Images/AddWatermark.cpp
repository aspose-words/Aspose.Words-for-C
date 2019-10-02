#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextPath.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

// ExStart:AddWatermark
namespace
{
    void InsertWatermarkIntoHeader(const System::SharedPtr<Paragraph>& watermarkPara, const System::SharedPtr<Section>& sect, HeaderFooterType headerType)
    {
        System::SharedPtr<HeaderFooter> header = sect->get_HeadersFooters()->idx_get(headerType);

        if (header == nullptr)
        {
            // There is no header of the specified type in the current section, create it.
            header = System::MakeObject<HeaderFooter>(sect->get_Document(), headerType);
            sect->get_HeadersFooters()->Add(header);
        }

        // Insert a clone of the watermark into the header.
        header->AppendChild((System::StaticCast<Node>(watermarkPara))->Clone(true));
    }

    void InsertWatermarkText(const System::SharedPtr<Document>& doc, const System::String& watermarkText)
    {
        // Create a watermark shape. This will be a WordArt shape. 
        // You are free to try other shape types as watermarks.
        System::SharedPtr<Shape> watermark = System::MakeObject<Shape>(doc, ShapeType::TextPlainText);
        watermark->set_Name(u"WaterMark");
        // Set up the text of the watermark.
        watermark->get_TextPath()->set_Text(watermarkText);
        watermark->get_TextPath()->set_FontFamily(u"Arial");
        watermark->set_Width(500);
        watermark->set_Height(100);
        // Text will be directed from the bottom-left to the top-right corner.
        watermark->set_Rotation(-40);
        // Remove the following two lines if you need a solid black text.
        watermark->get_Fill()->set_Color(System::Drawing::Color::get_Gray());
        // Try LightGray to get more Word-style watermark
        watermark->set_StrokeColor(System::Drawing::Color::get_Gray());
        // Try LightGray to get more Word-style watermark

        // Place the watermark in the page center.
        watermark->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        watermark->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        watermark->set_WrapType(WrapType::None);
        watermark->set_VerticalAlignment(VerticalAlignment::Center);
        watermark->set_HorizontalAlignment(HorizontalAlignment::Center);

        // Create a new paragraph and append the watermark to this paragraph.
        System::SharedPtr<Paragraph> watermarkPara = System::MakeObject<Paragraph>(doc);
        watermarkPara->AppendChild(watermark);

        // Insert the watermark into all headers of each document section.
        for (System::SharedPtr<Node> sectionNode : System::IterateOver(doc->get_Sections()))
        {
            System::SharedPtr<Section> sect = System::DynamicCast<Section>(sectionNode);
            // There could be up to three different headers in each section, since we want
            // The watermark to appear on all pages, insert into all headers.
            InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType::HeaderPrimary);
            InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType::HeaderFirst);
            InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType::HeaderEven);
        }
    }
}

void AddWatermark()
{
    std::cout << "AddWatermark example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithImages();
    System::String outputDataDir = GetOutputDataDir_WorkingWithImages();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.Watermark.doc");
    InsertWatermarkText(doc, u"CONFIDENTIAL");
    System::String outputPath = outputDataDir + u"AddWatermark.doc";
    doc->Save(outputPath);
    std::cout << "Added watermark to the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AddWatermark example finished." << std::endl << std::endl;
}
// ExEnd:AddWatermark