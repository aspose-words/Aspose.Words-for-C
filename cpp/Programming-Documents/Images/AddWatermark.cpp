#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/ienumerator.h>
#include <Model/Text/Paragraph.h>
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/HeaderFooterType.h>
#include <Model/Sections/HeaderFooterCollection.h>
#include <Model/Sections/HeaderFooter.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Drawing/WrapType.h>
#include <Model/Drawing/VerticalAlignment.h>
#include <Model/Drawing/TextPath.h>
#include <Model/Drawing/ShapeType.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/RelativeVerticalPosition.h>
#include <Model/Drawing/RelativeHorizontalPosition.h>
#include <Model/Drawing/HorizontalAlignment.h>
#include <Model/Drawing/Fill.h>
#include <Model/Document/DocumentBase.h>
#include <Model/Document/Document.h>
#include <drawing/color.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

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
        header->AppendChild((System::StaticCast<Aspose::Words::Node>(watermarkPara))->Clone(true));
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
        auto sect_enumerator = doc->get_Sections()->GetEnumerator();
        System::SharedPtr<Section> sect;
        while (sect_enumerator->MoveNext() && (sect = System::DynamicCast<Section>(sect_enumerator->get_Current()), true))
        {
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
    // ExStart:AddWatermark
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithImages();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.Watermark.doc");
    InsertWatermarkText(doc, u"CONFIDENTIAL");
    System::String outputPath = dataDir + GetOutputFilePath(u"AddWatermark.doc");
    doc->Save(outputPath);
    // ExEnd:AddWatermark
    std::cout << "Added watermark to the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AddWatermark example finished." << std::endl << std::endl;
}
