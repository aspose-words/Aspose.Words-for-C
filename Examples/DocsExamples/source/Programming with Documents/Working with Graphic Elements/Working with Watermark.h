#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/TextPath.h>
#include <Aspose.Words.Cpp/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/ImageWatermarkOptions.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/TextWatermarkOptions.h>
#include <Aspose.Words.Cpp/Watermark.h>
#include <Aspose.Words.Cpp/WatermarkLayout.h>
#include <Aspose.Words.Cpp/WatermarkType.h>
#include <drawing/color.h>
#include <drawing/image.h>
#include <system/enumerator_adapter.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements {

class WorkWithWatermark : public DocsExamplesBase
{
public:
    void AddTextWatermarkWithSpecificOptions()
    {
        //ExStart:AddTextWatermarkWithSpecificOptions
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto options = MakeObject<TextWatermarkOptions>();
        options->set_FontFamily(u"Arial");
        options->set_FontSize(36);
        options->set_Color(System::Drawing::Color::get_Black());
        options->set_Layout(WatermarkLayout::Horizontal);
        options->set_IsSemitrasparent(false);

        doc->get_Watermark()->SetText(u"Test", options);

        doc->Save(ArtifactsDir + u"WorkWithWatermark.AddTextWatermarkWithSpecificOptions.docx");
        //ExEnd:AddTextWatermarkWithSpecificOptions
    }

    void AddImageWatermarkWithSpecificOptions()
    {
        //ExStart:AddImageWatermarkWithSpecificOptions
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto options = MakeObject<ImageWatermarkOptions>();
        options->set_Scale(5);
        options->set_IsWashout(false);

        doc->get_Watermark()->SetImage(System::Drawing::Image::FromFile(ImagesDir + u"Transparent background logo.png"), options);

        doc->Save(ArtifactsDir + u"WorkWithWatermark.AddImageWatermark.docx");
        //ExEnd:AddImageWatermarkWithSpecificOptions
    }

    void RemoveWatermarkFromDocument()
    {
        //ExStart:RemoveWatermarkFromDocument
        auto doc = MakeObject<Document>();

        // Add a plain text watermark.
        doc->get_Watermark()->SetText(u"Aspose Watermark");

        // If we wish to edit the text formatting using it as a watermark,
        // we can do so by passing a TextWatermarkOptions object when creating the watermark.
        auto textWatermarkOptions = MakeObject<TextWatermarkOptions>();
        textWatermarkOptions->set_FontFamily(u"Arial");
        textWatermarkOptions->set_FontSize(36.0f);
        textWatermarkOptions->set_Color(System::Drawing::Color::get_Black());
        textWatermarkOptions->set_Layout(WatermarkLayout::Diagonal);
        textWatermarkOptions->set_IsSemitrasparent(false);

        doc->get_Watermark()->SetText(u"Aspose Watermark", textWatermarkOptions);

        doc->Save(ArtifactsDir + u"Document.TextWatermark.docx");

        // We can remove a watermark from a document like this.
        if (doc->get_Watermark()->get_Type() == WatermarkType::Text)
        {
            doc->get_Watermark()->Remove();
        }

        doc->Save(ArtifactsDir + u"WorkWithWatermark.RemoveWatermarkFromDocument.docx");
        //ExEnd:RemoveWatermarkFromDocument
    }

    void AddAndRemoveWatermark()
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        InsertWatermarkText(doc, u"CONFIDENTIAL");
        doc->Save(ArtifactsDir + u"TestFile.Watermark.docx");

        RemoveWatermarkText(doc);
        doc->Save(ArtifactsDir + u"WorkWithWatermark.RemoveWatermark.docx");
    }

protected:
    /// <summary>
    /// Inserts a watermark into a document.
    /// </summary>
    /// <param name="doc">The input document.</param>
    /// <param name="watermarkText">Text of the watermark.</param>
    void InsertWatermarkText(SharedPtr<Document> doc, String watermarkText)
    {
        // Create a watermark shape, this will be a WordArt shape.
        auto watermark = MakeObject<Shape>(doc, ShapeType::TextPlainText);
        watermark->set_Name(u"Watermark");

        watermark->get_TextPath()->set_Text(watermarkText);
        watermark->get_TextPath()->set_FontFamily(u"Arial");
        watermark->set_Width(500);
        watermark->set_Height(100);

        // Text will be directed from the bottom-left to the top-right corner.
        watermark->set_Rotation(-40);

        // Remove the following two lines if you need a solid black text.
        watermark->get_Fill()->set_Color(System::Drawing::Color::get_Gray());
        watermark->set_StrokeColor(System::Drawing::Color::get_Gray());

        // Place the watermark in the page center.
        watermark->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        watermark->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        watermark->set_WrapType(WrapType::None);
        watermark->set_VerticalAlignment(VerticalAlignment::Center);
        watermark->set_HorizontalAlignment(HorizontalAlignment::Center);

        // Create a new paragraph and append the watermark to this paragraph.
        auto watermarkPara = MakeObject<Paragraph>(doc);
        watermarkPara->AppendChild(watermark);

        // Insert the watermark into all headers of each document section.
        for (auto sect : System::IterateOver<Section>(doc->get_Sections()))
        {
            // There could be up to three different headers in each section.
            // Since we want the watermark to appear on all pages, insert it into all headers.
            InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType::HeaderPrimary);
            InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType::HeaderFirst);
            InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType::HeaderEven);
        }
    }

    void InsertWatermarkIntoHeader(SharedPtr<Paragraph> watermarkPara, SharedPtr<Section> sect, HeaderFooterType headerType)
    {
        SharedPtr<HeaderFooter> header = sect->get_HeadersFooters()->idx_get(headerType);

        if (header == nullptr)
        {
            // There is no header of the specified type in the current section, so we need to create it.
            header = MakeObject<HeaderFooter>(sect->get_Document(), headerType);
            sect->get_HeadersFooters()->Add(header);
        }

        // Insert a clone of the watermark into the header.
        header->AppendChild(watermarkPara->Clone(true));
    }

    void RemoveWatermarkText(SharedPtr<Document> doc)
    {
        for (auto hf : System::IterateOver<HeaderFooter>(doc->GetChildNodes(NodeType::HeaderFooter, true)))
        {
            for (auto shape : System::IterateOver<Shape>(hf->GetChildNodes(NodeType::Shape, true)))
            {
                if (shape->get_Name().Contains(u"WaterMark"))
                {
                    shape->Remove();
                }
            }
        }
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements
