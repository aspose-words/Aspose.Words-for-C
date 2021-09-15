#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BorderType.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/CalendarType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/ConvertUtil.h>
#include <Aspose.Words.Cpp/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalRuleAlignment.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalRuleFormat.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/OleFormat.h>
#include <Aspose.Words.Cpp/Drawing/OlePackage.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Fields/DropDownItemCollection.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldToc.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Fields/GeneralFormat.h>
#include <Aspose.Words.Cpp/Fields/IFieldResultFormatter.h>
#include <Aspose.Words.Cpp/Fields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/HeightRule.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/LineStyle.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeImporter.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Notes/FootnoteType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/MarkdownSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/TableContentAlignment.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Shading.h>
#include <Aspose.Words.Cpp/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Story.h>
#include <Aspose.Words.Cpp/StoryType.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/CellVerticalAlignment.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidthType.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Tables/TableStyleOptions.h>
#include <Aspose.Words.Cpp/TextOrientation.h>
#include <Aspose.Words.Cpp/TextureIndex.h>
#include <Aspose.Words.Cpp/Underline.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/WarningSource.h>
#include <drawing/color.h>
#include <drawing/image.h>
#include <net/http_status_code.h>
#include <net/service_point_manager.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/guid.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/path.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/math.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/timespan.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::DigitalSignatures;

namespace ApiExamples {

class ExDocumentBuilder : public ApiExampleBase
{
public:
    void WriteAndFont()
    {
        //ExStart
        //ExFor:Font.Size
        //ExFor:Font.Bold
        //ExFor:Font.Name
        //ExFor:Font.Color
        //ExFor:Font.Underline
        //ExFor:DocumentBuilder.#ctor(Document)
        //ExSummary:Shows how to insert formatted text using DocumentBuilder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify font formatting, then add text.
        SharedPtr<Aspose::Words::Font> font = builder->get_Font();
        font->set_Size(16);
        font->set_Bold(true);
        font->set_Color(System::Drawing::Color::get_Blue());
        font->set_Name(u"Courier New");
        font->set_Underline(Underline::Dash);

        builder->Write(u"Hello world!");
        //ExEnd

        doc = DocumentHelper::SaveOpen(builder->get_Document());
        SharedPtr<Run> firstRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Hello world!", firstRun->GetText().Trim());
        ASPOSE_ASSERT_EQ(16.0, firstRun->get_Font()->get_Size());
        ASSERT_TRUE(firstRun->get_Font()->get_Bold());
        ASSERT_EQ(u"Courier New", firstRun->get_Font()->get_Name());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), firstRun->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(Underline::Dash, firstRun->get_Font()->get_Underline());
    }

    void HeadersAndFooters()
    {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.#ctor(Document)
        //ExFor:DocumentBuilder.MoveToHeaderFooter
        //ExFor:DocumentBuilder.MoveToSection
        //ExFor:DocumentBuilder.InsertBreak
        //ExFor:DocumentBuilder.Writeln(String)
        //ExFor:HeaderFooterType
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
        //ExFor:BreakType
        //ExSummary:Shows how to create headers and footers in a document using DocumentBuilder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify that we want different headers and footers for first, even and odd pages.
        builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(true);
        builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(true);

        // Create the headers, then add three pages to the document to display each header type.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->Write(u"Header for the first page");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderEven);
        builder->Write(u"Header for even pages");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"Header for all other pages");

        builder->MoveToSection(0);
        builder->Writeln(u"Page1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page3");

        doc->Save(ArtifactsDir + u"DocumentBuilder.HeadersAndFooters.docx");
        //ExEnd

        auto savedDoc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.HeadersAndFooters.docx");
        SharedPtr<HeaderFooterCollection> headersFooters = savedDoc->get_FirstSection()->get_HeadersFooters();

        ASSERT_EQ(3, headersFooters->get_Count());
        ASSERT_EQ(u"Header for the first page", headersFooters->idx_get(HeaderFooterType::HeaderFirst)->GetText().Trim());
        ASSERT_EQ(u"Header for even pages", headersFooters->idx_get(HeaderFooterType::HeaderEven)->GetText().Trim());
        ASSERT_EQ(u"Header for all other pages", headersFooters->idx_get(HeaderFooterType::HeaderPrimary)->GetText().Trim());
    }

    void MergeFields()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(String)
        //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
        //ExSummary:Shows how to insert fields, and move the document builder's cursor to them.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertField(u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder->InsertField(u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        // Move the cursor to the first MERGEFIELD.
        builder->MoveToMergeField(u"MyMergeField1", true, false);

        // Note that the cursor is placed immediately after the first MERGEFIELD, and before the second.
        ASPOSE_ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(1)->get_Start(), builder->get_CurrentNode());
        ASPOSE_ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(0)->get_End(), builder->get_CurrentNode()->get_PreviousSibling());

        // If we wish to edit the field's field code or contents using the builder,
        // its cursor would need to be inside a field.
        // To place it inside a field, we would need to call the document builder's MoveTo method
        // and pass the field's start or separator node as an argument.
        builder->Write(u" Text between our merge fields. ");

        doc->Save(ArtifactsDir + u"DocumentBuilder.MergeFields.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MergeFields.docx");

        ASSERT_EQ(String(u"\u0013MERGEFIELD MyMergeField1 \\* MERGEFORMAT\u0014«MyMergeField1»\u0015") + u" Text between our merge fields. " +
                      u"\u0013MERGEFIELD MyMergeField2 \\* MERGEFORMAT\u0014«MyMergeField2»\u0015",
                  doc->GetText().Trim());
        ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());
        TestUtil::VerifyField(FieldType::FieldMergeField, u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT", u"«MyMergeField1»",
                              doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldMergeField, u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT", u"«MyMergeField2»",
                              doc->get_Range()->get_Fields()->idx_get(1));
    }

    void InsertHorizontalRule()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHorizontalRule
        //ExFor:ShapeBase.IsHorizontalRule
        //ExFor:Shape.HorizontalRuleFormat
        //ExFor:HorizontalRuleFormat
        //ExFor:HorizontalRuleFormat.Alignment
        //ExFor:HorizontalRuleFormat.WidthPercent
        //ExFor:HorizontalRuleFormat.Height
        //ExFor:HorizontalRuleFormat.Color
        //ExFor:HorizontalRuleFormat.NoShade
        //ExSummary:Shows how to insert a horizontal rule shape, and customize its formatting.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<Shape> shape = builder->InsertHorizontalRule();

        SharedPtr<HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();
        horizontalRuleFormat->set_Alignment(HorizontalRuleAlignment::Center);
        horizontalRuleFormat->set_WidthPercent(70);
        horizontalRuleFormat->set_Height(3);
        horizontalRuleFormat->set_Color(System::Drawing::Color::get_Blue());
        horizontalRuleFormat->set_NoShade(true);

        ASSERT_TRUE(shape->get_IsHorizontalRule());
        ASSERT_TRUE(shape->get_HorizontalRuleFormat()->get_NoShade());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(HorizontalRuleAlignment::Center, shape->get_HorizontalRuleFormat()->get_Alignment());
        ASPOSE_ASSERT_EQ(70, shape->get_HorizontalRuleFormat()->get_WidthPercent());
        ASPOSE_ASSERT_EQ(3, shape->get_HorizontalRuleFormat()->get_Height());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), shape->get_HorizontalRuleFormat()->get_Color().ToArgb());
    }

    void HorizontalRuleFormatExceptions()
    {
        auto builder = MakeObject<DocumentBuilder>();
        SharedPtr<Shape> shape = builder->InsertHorizontalRule();

        SharedPtr<HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();
        horizontalRuleFormat->set_WidthPercent(1);
        horizontalRuleFormat->set_WidthPercent(100);
        ASSERT_THROW(horizontalRuleFormat->set_WidthPercent(0), System::ArgumentOutOfRangeException);
        ASSERT_THROW(horizontalRuleFormat->set_WidthPercent(101), System::ArgumentOutOfRangeException);

        horizontalRuleFormat->set_Height(0);
        horizontalRuleFormat->set_Height(1584);
        ASSERT_THROW(horizontalRuleFormat->set_Height(-1), System::ArgumentOutOfRangeException);
        ASSERT_THROW(horizontalRuleFormat->set_Height(1585), System::ArgumentOutOfRangeException);
    }

    void InsertHyperlink()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExFor:Font.ClearFormatting
        //ExFor:Font.Color
        //ExFor:Font.Underline
        //ExFor:Underline
        //ExSummary:Shows how to insert a hyperlink field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"For more information, please visit the ");

        // Insert a hyperlink and emphasize it with custom formatting.
        // The hyperlink will be a clickable piece of text which will take us to the location specified in the URL.
        builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        builder->get_Font()->set_Underline(Underline::Single);
        builder->InsertHyperlink(u"Google website", u"https://www.google.com", false);
        builder->get_Font()->ClearFormatting();
        builder->Writeln(u".");

        // Ctrl + left clicking the link in the text in Microsoft Word will take us to the URL via a new web browser window.
        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHyperlink.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertHyperlink.docx");

        auto hyperlink = System::DynamicCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, hyperlink->get_Address());

        auto fieldContents = System::DynamicCast<Run>(hyperlink->get_Start()->get_NextSibling());

        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), fieldContents->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(Underline::Single, fieldContents->get_Font()->get_Underline());
        ASSERT_EQ(u"HYPERLINK \"https://www.google.com\"", fieldContents->GetText().Trim());
    }

    void PushPopFont()
    {
        //ExStart
        //ExFor:DocumentBuilder.PushFont
        //ExFor:DocumentBuilder.PopFont
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExSummary:Shows how to use a document builder's formatting stack.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set up font formatting, then write the text that goes before the hyperlink.
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Size(24);
        builder->Write(u"To visit Google, hold Ctrl and click ");

        // Preserve our current formatting configuration on the stack.
        builder->PushFont();

        // Alter the builder's current formatting by applying a new style.
        builder->get_Font()->set_StyleIdentifier(StyleIdentifier::Hyperlink);
        builder->InsertHyperlink(u"here", u"http://www.google.com", false);

        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), builder->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(Underline::Single, builder->get_Font()->get_Underline());

        // Restore the font formatting that we saved earlier and remove the element from the stack.
        builder->PopFont();

        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), builder->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(Underline::None, builder->get_Font()->get_Underline());

        builder->Write(u". We hope you enjoyed the example.");

        doc->Save(ArtifactsDir + u"DocumentBuilder.PushPopFont.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.PushPopFont.docx");
        SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();

        ASSERT_EQ(4, runs->get_Count());

        ASSERT_EQ(u"To visit Google, hold Ctrl and click", runs->idx_get(0)->GetText().Trim());
        ASSERT_EQ(u". We hope you enjoyed the example.", runs->idx_get(3)->GetText().Trim());
        ASPOSE_ASSERT_EQ(runs->idx_get(0)->get_Font()->get_Color(), runs->idx_get(3)->get_Font()->get_Color());
        ASSERT_EQ(runs->idx_get(0)->get_Font()->get_Underline(), runs->idx_get(3)->get_Font()->get_Underline());

        ASSERT_EQ(u"here", runs->idx_get(2)->GetText().Trim());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), runs->idx_get(2)->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(Underline::Single, runs->idx_get(2)->get_Font()->get_Underline());
        ASPOSE_ASSERT_NE(runs->idx_get(0)->get_Font()->get_Color(), runs->idx_get(2)->get_Font()->get_Color());
        ASSERT_NE(runs->idx_get(0)->get_Font()->get_Underline(), runs->idx_get(2)->get_Font()->get_Underline());

        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK,
                                              (System::DynamicCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address());
    }

    void InsertWatermark()
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToHeaderFooter
        //ExFor:PageSetup.PageWidth
        //ExFor:PageSetup.PageHeight
        //ExFor:WrapType
        //ExFor:RelativeHorizontalPosition
        //ExFor:RelativeVerticalPosition
        //ExSummary:Shows how to insert an image, and use it as a watermark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert the image into the header so that it will be visible on every page.
        SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        SharedPtr<Shape> shape = builder->InsertImage(image);
        shape->set_WrapType(WrapType::None);
        shape->set_BehindText(true);

        // Place the image at the center of the page.
        shape->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        shape->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        shape->set_Left((builder->get_PageSetup()->get_PageWidth() - shape->get_Width()) / 2);
        shape->set_Top((builder->get_PageSetup()->get_PageHeight() - shape->get_Height()) / 2);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertWatermark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertWatermark.docx");
        shape = System::DynamicCast<Shape>(
            doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Png, shape);
        ASSERT_EQ(WrapType::None, shape->get_WrapType());
        ASSERT_TRUE(shape->get_BehindText());
        ASSERT_EQ(RelativeHorizontalPosition::Page, shape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Page, shape->get_RelativeVerticalPosition());
        ASPOSE_ASSERT_EQ((doc->get_FirstSection()->get_PageSetup()->get_PageWidth() - shape->get_Width()) / 2, shape->get_Left());
        ASPOSE_ASSERT_EQ((doc->get_FirstSection()->get_PageSetup()->get_PageHeight() - shape->get_Height()) / 2, shape->get_Top());
    }

    void InsertHtml()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String)
        //ExSummary:Shows how to use a document builder to insert html content into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        const String html = String(u"<p align='right'>Paragraph right</p>") + u"<b>Implicit paragraph left</b>" + u"<div align='center'>Div center</div>" +
                            u"<h1 align='left'>Heading 1 left.</h1>";

        builder->InsertHtml(html);

        // Inserting HTML code parses the formatting of each element into equivalent document text formatting.
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(u"Paragraph right", paragraphs->idx_get(0)->GetText().Trim());
        ASSERT_EQ(ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());

        ASSERT_EQ(u"Implicit paragraph left", paragraphs->idx_get(1)->GetText().Trim());
        ASSERT_EQ(ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
        ASSERT_TRUE(paragraphs->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Bold());

        ASSERT_EQ(u"Div center", paragraphs->idx_get(2)->GetText().Trim());
        ASSERT_EQ(ParagraphAlignment::Center, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());

        ASSERT_EQ(u"Heading 1 left.", paragraphs->idx_get(3)->GetText().Trim());
        ASSERT_EQ(u"Heading 1", paragraphs->idx_get(3)->get_ParagraphFormat()->get_Style()->get_Name());

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHtml.docx");
        //ExEnd
    }

    void InsertHtmlWithFormatting(bool useBuilderFormatting)
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
        //ExSummary:Shows how to apply a document builder's formatting while inserting HTML content.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set a text alignment for the builder, insert an HTML paragraph with a specified alignment, and one without.
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Distributed);
        builder->InsertHtml(String(u"<p align='right'>Paragraph 1.</p>") + u"<p>Paragraph 2.</p>", useBuilderFormatting);

        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        // The first paragraph has an alignment specified. When InsertHtml parses the HTML code,
        // the paragraph alignment value found in the HTML code always supersedes the document builder's value.
        ASSERT_EQ(u"Paragraph 1.", paragraphs->idx_get(0)->GetText().Trim());
        ASSERT_EQ(ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());

        // The second paragraph has no alignment specified. It can have its alignment value filled in
        // by the builder's value depending on the flag we passed to the InsertHtml method.
        ASSERT_EQ(u"Paragraph 2.", paragraphs->idx_get(1)->GetText().Trim());
        ASSERT_EQ(useBuilderFormatting ? ParagraphAlignment::Distributed : ParagraphAlignment::Left,
                  paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHtmlWithFormatting.docx");
        //ExEnd
    }

    void MathMl()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        const String mathMl = u"<math "
                              u"xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</"
                              u"mi><mrow><mn>1</mn></mrow></msub></mrow></math>";

        builder->InsertHtml(mathMl);

        doc->Save(ArtifactsDir + u"DocumentBuilder.MathML.docx");
        doc->Save(ArtifactsDir + u"DocumentBuilder.MathML.pdf");

        ASSERT_TRUE(DocumentHelper::CompareDocs(GoldsDir + u"DocumentBuilder.MathML Gold.docx", ArtifactsDir + u"DocumentBuilder.MathML.docx"));
    }

    void InsertTextAndBookmark()
    {
        //ExStart
        //ExFor:DocumentBuilder.StartBookmark
        //ExFor:DocumentBuilder.EndBookmark
        //ExSummary:Shows how create a bookmark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // A valid bookmark needs to have document body text enclosed by
        // BookmarkStart and BookmarkEnd nodes created with a matching bookmark name.
        builder->StartBookmark(u"MyBookmark");
        builder->Writeln(u"Hello world!");
        builder->EndBookmark(u"MyBookmark");

        ASSERT_EQ(1, doc->get_Range()->get_Bookmarks()->get_Count());
        ASSERT_EQ(u"MyBookmark", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Name());
        ASSERT_EQ(u"Hello world!", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Text().Trim());
        //ExEnd
    }

    void CreateColumnBookmark()
    {
        //ExStart
        //ExFor:DocumentBuilder.StartColumnBookmark
        //ExFor:DocumentBuilder.EndColumnBookmark
        //ExSummary:Shows how to create a column bookmark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();

        builder->InsertCell();
        // Cells 1,2,4,5 will be bookmarked.
        builder->StartColumnBookmark(u"MyBookmark_1");
        // Badly formed bookmarks or bookmarks with duplicate names will be ignored when the document is saved.
        builder->StartColumnBookmark(u"MyBookmark_1");
        builder->StartColumnBookmark(u"BadStartBookmark");
        builder->Write(u"Cell 1");

        builder->InsertCell();
        builder->Write(u"Cell 2");

        builder->InsertCell();
        builder->Write(u"Cell 3");

        builder->EndRow();

        builder->InsertCell();
        builder->Write(u"Cell 4");

        builder->InsertCell();
        builder->Write(u"Cell 5");
        builder->EndColumnBookmark(u"MyBookmark_1");
        builder->EndColumnBookmark(u"MyBookmark_1");

        ASSERT_THROW(builder->EndColumnBookmark(u"BadEndBookmark"), System::InvalidOperationException);
        //ExSkip

        builder->InsertCell();
        builder->Write(u"Cell 6");

        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"Bookmarks.CreateColumnBookmark.docx");
        //ExEnd
    }

    void CreateForm()
    {
        //ExStart
        //ExFor:TextFormFieldType
        //ExFor:DocumentBuilder.InsertTextInput
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Shows how to create form fields.
        auto builder = MakeObject<DocumentBuilder>();

        // Form fields are objects in the document that the user can interact with by being prompted to enter values.
        // We can create them using a document builder, and below are two ways of doing so.
        // 1 -  Basic text input:
        builder->InsertTextInput(u"My text input", TextFormFieldType::Regular, u"", u"Enter your name here", 30);

        // 2 -  Combo box with prompt text, and a range of possible values:
        ArrayPtr<String> items = MakeArray<String>({u"-- Select your favorite footwear --", u"Sneakers", u"Oxfords", u"Flip-flops", u"Other"});

        builder->InsertParagraph();
        builder->InsertComboBox(u"My combo box", items, 0);

        builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.CreateForm.docx");
        //ExEnd

        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.CreateForm.docx");
        SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_EQ(u"My text input", formField->get_Name());
        ASSERT_EQ(TextFormFieldType::Regular, formField->get_TextInputType());
        ASSERT_EQ(u"Enter your name here", formField->get_Result());

        formField = doc->get_Range()->get_FormFields()->idx_get(1);

        ASSERT_EQ(u"My combo box", formField->get_Name());
        ASSERT_EQ(TextFormFieldType::Regular, formField->get_TextInputType());
        ASSERT_EQ(u"-- Select your favorite footwear --", formField->get_Result());
        ASSERT_EQ(0, formField->get_DropDownSelectedIndex());
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"-- Select your favorite footwear --", u"Sneakers", u"Oxfords", u"Flip-flops", u"Other"}),
                         formField->get_DropDownItems()->LINQ_ToArray());
    }

    void InsertCheckBox()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertCheckBox(string, bool, bool, int)
        //ExFor:DocumentBuilder.InsertCheckBox(String, bool, int)
        //ExSummary:Shows how to insert checkboxes into the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert checkboxes of varying sizes and default checked statuses.
        builder->Write(u"Unchecked check box of a default size: ");
        builder->InsertCheckBox(String::Empty, false, false, 0);
        builder->InsertParagraph();

        builder->Write(u"Large checked check box: ");
        builder->InsertCheckBox(u"CheckBox_Default", true, true, 50);
        builder->InsertParagraph();

        // Form fields have a name length limit of 20 characters.
        builder->Write(u"Very large checked check box: ");
        builder->InsertCheckBox(u"CheckBox_OnlyCheckedValue", true, 100);

        ASSERT_EQ(u"CheckBox_OnlyChecked", doc->get_Range()->get_FormFields()->idx_get(2)->get_Name());

        // We can interact with these check boxes in Microsoft Word by double clicking them.
        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertCheckBox.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertCheckBox.docx");

        SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();

        ASSERT_EQ(String::Empty, formFields->idx_get(0)->get_Name());
        ASPOSE_ASSERT_EQ(false, formFields->idx_get(0)->get_Checked());
        ASPOSE_ASSERT_EQ(false, formFields->idx_get(0)->get_Default());
        ASPOSE_ASSERT_EQ(10, formFields->idx_get(0)->get_CheckBoxSize());

        ASSERT_EQ(u"CheckBox_Default", formFields->idx_get(1)->get_Name());
        ASPOSE_ASSERT_EQ(true, formFields->idx_get(1)->get_Checked());
        ASPOSE_ASSERT_EQ(true, formFields->idx_get(1)->get_Default());
        ASPOSE_ASSERT_EQ(50, formFields->idx_get(1)->get_CheckBoxSize());

        ASSERT_EQ(u"CheckBox_OnlyChecked", formFields->idx_get(2)->get_Name());
        ASPOSE_ASSERT_EQ(true, formFields->idx_get(2)->get_Checked());
        ASPOSE_ASSERT_EQ(true, formFields->idx_get(2)->get_Default());
        ASPOSE_ASSERT_EQ(100, formFields->idx_get(2)->get_CheckBoxSize());
    }

    void InsertCheckBoxEmptyName()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Checking that the checkbox insertion with an empty name working correctly
        builder->InsertCheckBox(u"", true, false, 1);
        builder->InsertCheckBox(String::Empty, false, 1);
    }

    void WorkingWithNodes()
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveTo(Node)
        //ExFor:DocumentBuilder.MoveToBookmark(String)
        //ExFor:DocumentBuilder.CurrentParagraph
        //ExFor:DocumentBuilder.CurrentNode
        //ExFor:DocumentBuilder.MoveToDocumentStart
        //ExFor:DocumentBuilder.MoveToDocumentEnd
        //ExFor:DocumentBuilder.IsAtEndOfParagraph
        //ExFor:DocumentBuilder.IsAtStartOfParagraph
        //ExSummary:Shows how to move a document builder's cursor to different nodes in a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a valid bookmark, an entity that consists of nodes enclosed by a bookmark start node,
        // and a bookmark end node.
        builder->StartBookmark(u"MyBookmark");
        builder->Write(u"Bookmark contents.");
        builder->EndBookmark(u"MyBookmark");

        SharedPtr<NodeCollection> firstParagraphNodes = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ChildNodes();

        ASSERT_EQ(NodeType::BookmarkStart, firstParagraphNodes->idx_get(0)->get_NodeType());
        ASSERT_EQ(NodeType::Run, firstParagraphNodes->idx_get(1)->get_NodeType());
        ASSERT_EQ(u"Bookmark contents.", firstParagraphNodes->idx_get(1)->GetText().Trim());
        ASSERT_EQ(NodeType::BookmarkEnd, firstParagraphNodes->idx_get(2)->get_NodeType());

        // The document builder's cursor is always ahead of the node that we last added with it.
        // If the builder's cursor is at the end of the document, its current node will be null.
        // The previous node is the bookmark end node that we last added.
        // Adding new nodes with the builder will append them to the last node.
        ASSERT_TRUE(builder->get_CurrentNode() == nullptr);

        // If we wish to edit a different part of the document with the builder,
        // we will need to bring its cursor to the node we wish to edit.
        builder->MoveToBookmark(u"MyBookmark");

        // Moving it to a bookmark will move it to the first node within the bookmark start and end nodes, the enclosed run.
        ASPOSE_ASSERT_EQ(firstParagraphNodes->idx_get(1), builder->get_CurrentNode());

        // We can also move the cursor to an individual node like this.
        builder->MoveTo(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(NodeType::Any, false)->idx_get(0));

        ASSERT_EQ(NodeType::BookmarkStart, builder->get_CurrentNode()->get_NodeType());
        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), builder->get_CurrentParagraph());
        ASSERT_TRUE(builder->get_IsAtStartOfParagraph());

        // We can use specific methods to move to the start/end of a document.
        builder->MoveToDocumentEnd();

        ASSERT_TRUE(builder->get_IsAtEndOfParagraph());

        builder->MoveToDocumentStart();

        ASSERT_TRUE(builder->get_IsAtStartOfParagraph());
        //ExEnd
    }

    void FillMergeFields()
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToMergeField(String)
        //ExFor:DocumentBuilder.Bold
        //ExFor:DocumentBuilder.Italic
        //ExSummary:Shows how to fill MERGEFIELDs with data with a document builder instead of a mail merge.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge,
        // and then fill them manually.
        builder->InsertField(u" MERGEFIELD Chairman ");
        builder->InsertField(u" MERGEFIELD ChiefFinancialOfficer ");
        builder->InsertField(u" MERGEFIELD ChiefTechnologyOfficer ");

        builder->MoveToMergeField(u"Chairman");
        builder->set_Bold(true);
        builder->Writeln(u"John Doe");

        builder->MoveToMergeField(u"ChiefFinancialOfficer");
        builder->set_Italic(true);
        builder->Writeln(u"Jane Doe");

        builder->MoveToMergeField(u"ChiefTechnologyOfficer");
        builder->set_Italic(true);
        builder->Writeln(u"John Bloggs");

        doc->Save(ArtifactsDir + u"DocumentBuilder.FillMergeFields.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.FillMergeFields.docx");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_TRUE(paragraphs->idx_get(0)->get_Runs()->idx_get(0)->get_Font()->get_Bold());
        ASSERT_EQ(u"John Doe", paragraphs->idx_get(0)->get_Runs()->idx_get(0)->GetText().Trim());

        ASSERT_TRUE(paragraphs->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Italic());
        ASSERT_EQ(u"Jane Doe", paragraphs->idx_get(1)->get_Runs()->idx_get(0)->GetText().Trim());

        ASSERT_TRUE(paragraphs->idx_get(2)->get_Runs()->idx_get(0)->get_Font()->get_Italic());
        ASSERT_EQ(u"John Bloggs", paragraphs->idx_get(2)->get_Runs()->idx_get(0)->GetText().Trim());
    }

    void InsertToc()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTableOfContents
        //ExFor:Document.UpdateFields
        //ExFor:DocumentBuilder.#ctor(Document)
        //ExFor:ParagraphFormat.StyleIdentifier
        //ExFor:DocumentBuilder.InsertBreak
        //ExFor:BreakType
        //ExSummary:Shows how to insert a Table of contents (TOC) into a document using heading styles as entries.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a table of contents for the first page of the document.
        // Configure the table to pick up paragraphs with headings of levels 1 to 3.
        // Also, set its entries to be hyperlinks that will take us
        // to the location of the heading when left-clicked in Microsoft Word.
        builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");
        builder->InsertBreak(BreakType::PageBreak);

        // Populate the table of contents by adding paragraphs with heading styles.
        // Each such heading with a level between 1 and 3 will create an entry in the table.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);
        builder->Writeln(u"Heading 1.1");
        builder->Writeln(u"Heading 1.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Writeln(u"Heading 2");
        builder->Writeln(u"Heading 3");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);
        builder->Writeln(u"Heading 3.1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading3);
        builder->Writeln(u"Heading 3.1.1");
        builder->Writeln(u"Heading 3.1.2");
        builder->Writeln(u"Heading 3.1.3");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading4);
        builder->Writeln(u"Heading 3.1.3.1");
        builder->Writeln(u"Heading 3.1.3.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);
        builder->Writeln(u"Heading 3.2");
        builder->Writeln(u"Heading 3.3");

        // A table of contents is a field of a type that needs to be updated to show an up-to-date result.
        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertToc.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertToc.docx");
        auto tableOfContents = System::DynamicCast<FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u"1-3", tableOfContents->get_HeadingLevelRange());
        ASSERT_TRUE(tableOfContents->get_InsertHyperlinks());
        ASSERT_TRUE(tableOfContents->get_HideInWebLayout());
        ASSERT_TRUE(tableOfContents->get_UseParagraphOutlineLevel());
    }

    void InsertTable()
    {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.Write
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.InsertCell
        //ExFor:DocumentBuilder.EndRow
        //ExFor:DocumentBuilder.EndTable
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:CellFormat
        //ExFor:CellFormat.FitText
        //ExFor:CellFormat.Width
        //ExFor:CellFormat.VerticalAlignment
        //ExFor:CellFormat.Shading
        //ExFor:CellFormat.Orientation
        //ExFor:CellFormat.WrapText
        //ExFor:RowFormat
        //ExFor:RowFormat.Borders
        //ExFor:RowFormat.ClearFormatting
        //ExFor:Shading.ClearFormatting
        //ExSummary:Shows how to build a table with custom borders.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();

        // Setting table formatting options for a document builder
        // will apply them to every row and cell that we add with it.
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        builder->get_CellFormat()->ClearFormatting();
        builder->get_CellFormat()->set_Width(150);
        builder->get_CellFormat()->set_VerticalAlignment(CellVerticalAlignment::Center);
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_GreenYellow());
        builder->get_CellFormat()->set_WrapText(false);
        builder->get_CellFormat()->set_FitText(true);

        builder->get_RowFormat()->ClearFormatting();
        builder->get_RowFormat()->set_HeightRule(HeightRule::Exactly);
        builder->get_RowFormat()->set_Height(50);
        builder->get_RowFormat()->get_Borders()->set_LineStyle(LineStyle::Engrave3D);
        builder->get_RowFormat()->get_Borders()->set_Color(System::Drawing::Color::get_Orange());

        builder->InsertCell();
        builder->Write(u"Row 1, Col 1");

        builder->InsertCell();
        builder->Write(u"Row 1, Col 2");
        builder->EndRow();

        // Changing the formatting will apply it to the current cell,
        // and any new cells that we create with the builder afterward.
        // This will not affect the cells that we have added previously.
        builder->get_CellFormat()->get_Shading()->ClearFormatting();

        builder->InsertCell();
        builder->Write(u"Row 2, Col 1");

        builder->InsertCell();
        builder->Write(u"Row 2, Col 2");

        builder->EndRow();

        // Increase row height to fit the vertical text.
        builder->InsertCell();
        builder->get_RowFormat()->set_Height(150);
        builder->get_CellFormat()->set_Orientation(TextOrientation::Upward);
        builder->Write(u"Row 3, Col 1");

        builder->InsertCell();
        builder->get_CellFormat()->set_Orientation(TextOrientation::Downward);
        builder->Write(u"Row 3, Col 2");

        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTable.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTable.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(u"Row 1, Col 1\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
        ASSERT_EQ(u"Row 1, Col 2\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());
        ASSERT_EQ(HeightRule::Exactly, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
        ASPOSE_ASSERT_EQ(50.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
        ASSERT_EQ(LineStyle::Engrave3D, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Borders()->get_LineStyle());
        ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), table->get_Rows()->idx_get(0)->get_RowFormat()->get_Borders()->get_Color().ToArgb());

        for (const auto& c : System::IterateOver<Cell>(table->get_Rows()->idx_get(0)->get_Cells()))
        {
            ASPOSE_ASSERT_EQ(150, c->get_CellFormat()->get_Width());
            ASSERT_EQ(CellVerticalAlignment::Center, c->get_CellFormat()->get_VerticalAlignment());
            ASSERT_EQ(System::Drawing::Color::get_GreenYellow().ToArgb(), c->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
            ASSERT_FALSE(c->get_CellFormat()->get_WrapText());
            ASSERT_TRUE(c->get_CellFormat()->get_FitText());

            ASSERT_EQ(ParagraphAlignment::Center, c->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
        }

        ASSERT_EQ(u"Row 2, Col 1\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
        ASSERT_EQ(u"Row 2, Col 2\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());

        for (const auto& c : System::IterateOver<Cell>(table->get_Rows()->idx_get(1)->get_Cells()))
        {
            ASPOSE_ASSERT_EQ(150, c->get_CellFormat()->get_Width());
            ASSERT_EQ(CellVerticalAlignment::Center, c->get_CellFormat()->get_VerticalAlignment());
            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
            ASSERT_FALSE(c->get_CellFormat()->get_WrapText());
            ASSERT_TRUE(c->get_CellFormat()->get_FitText());

            ASSERT_EQ(ParagraphAlignment::Center, c->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
        }

        ASPOSE_ASSERT_EQ(150, table->get_Rows()->idx_get(2)->get_RowFormat()->get_Height());

        ASSERT_EQ(u"Row 3, Col 1\a", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(0)->GetText().Trim());
        ASSERT_EQ(TextOrientation::Upward, table->get_Rows()->idx_get(2)->get_Cells()->idx_get(0)->get_CellFormat()->get_Orientation());
        ASSERT_EQ(ParagraphAlignment::Center,
                  table->get_Rows()->idx_get(2)->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());

        ASSERT_EQ(u"Row 3, Col 2\a", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(1)->GetText().Trim());
        ASSERT_EQ(TextOrientation::Downward, table->get_Rows()->idx_get(2)->get_Cells()->idx_get(1)->get_CellFormat()->get_Orientation());
        ASSERT_EQ(ParagraphAlignment::Center,
                  table->get_Rows()->idx_get(2)->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
    }

    void InsertTableWithStyle()
    {
        //ExStart
        //ExFor:Table.StyleIdentifier
        //ExFor:Table.StyleOptions
        //ExFor:TableStyleOptions
        //ExFor:Table.AutoFit
        //ExFor:AutoFitBehavior
        //ExSummary:Shows how to build a new table while applying a style.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<Table> table = builder->StartTable();

        // We must insert at least one row before setting any table formatting.
        builder->InsertCell();

        // Set the table style used based on the style identifier.
        // Note that not all table styles are available when saving to .doc format.
        table->set_StyleIdentifier(StyleIdentifier::MediumShading1Accent1);

        // Partially apply the style to features of the table based on predicates, then build the table.
        table->set_StyleOptions(TableStyleOptions::FirstColumn | TableStyleOptions::RowBands | TableStyleOptions::FirstRow);
        table->AutoFit(AutoFitBehavior::AutoFitToContents);

        builder->Writeln(u"Item");
        builder->get_CellFormat()->set_RightPadding(40);
        builder->InsertCell();
        builder->Writeln(u"Quantity (kg)");
        builder->EndRow();

        builder->InsertCell();
        builder->Writeln(u"Apples");
        builder->InsertCell();
        builder->Writeln(u"20");
        builder->EndRow();

        builder->InsertCell();
        builder->Writeln(u"Bananas");
        builder->InsertCell();
        builder->Writeln(u"40");
        builder->EndRow();

        builder->InsertCell();
        builder->Writeln(u"Carrots");
        builder->InsertCell();
        builder->Writeln(u"50");
        builder->EndRow();

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableWithStyle.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableWithStyle.docx");

        doc->ExpandTableStylesToDirectFormatting();

        ASSERT_EQ(u"Medium Shading 1 Accent 1", table->get_Style()->get_Name());
        ASSERT_EQ(TableStyleOptions::FirstColumn | TableStyleOptions::RowBands | TableStyleOptions::FirstRow, table->get_StyleOptions());
        ASSERT_EQ(189,
                  ASPOSECPP_CHECKED_CAST(int, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().get_B()));
        ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(),
                  table->get_FirstRow()->get_FirstCell()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Color().ToArgb());
        ASSERT_NE(System::Drawing::Color::get_LightBlue().ToArgb(),
                  ASPOSECPP_CHECKED_CAST(int, table->get_LastRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().get_B()));
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(),
                  table->get_LastRow()->get_FirstCell()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Color().ToArgb());
    }

    void InsertTableSetHeadingRow()
    {
        //ExStart
        //ExFor:RowFormat.HeadingFormat
        //ExSummary:Shows how to build a table with rows that repeat on every page.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();

        // Any rows inserted while the "HeadingFormat" flag is set to "true"
        // will show up at the top of the table on every page that it spans.
        builder->get_RowFormat()->set_HeadingFormat(true);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->get_CellFormat()->set_Width(100);
        builder->InsertCell();
        builder->Write(u"Heading row 1");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Heading row 2");
        builder->EndRow();

        builder->get_CellFormat()->set_Width(50);
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->get_RowFormat()->set_HeadingFormat(false);

        // Add enough rows for the table to span two pages.
        for (int i = 0; i < 50; i++)
        {
            builder->InsertCell();
            builder->Write(String::Format(u"Row {0}, column 1.", table->get_Rows()->get_Count()));
            builder->InsertCell();
            builder->Write(String::Format(u"Row {0}, column 2.", table->get_Rows()->get_Count()));
            builder->EndRow();
        }

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableSetHeadingRow.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableSetHeadingRow.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        for (int i = 0; i < table->get_Rows()->get_Count(); i++)
        {
            ASPOSE_ASSERT_EQ(i < 2, table->get_Rows()->idx_get(i)->get_RowFormat()->get_HeadingFormat());
        }
    }

    void InsertTableWithPreferredWidth()
    {
        //ExStart
        //ExFor:Table.PreferredWidth
        //ExFor:PreferredWidth.FromPercent
        //ExFor:PreferredWidth
        //ExSummary:Shows how to set a table to auto fit to 50% of the width of the page.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Cell #1");
        builder->InsertCell();
        builder->Write(u"Cell #2");
        builder->InsertCell();
        builder->Write(u"Cell #3");

        table->set_PreferredWidth(PreferredWidth::FromPercent(50));

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableWithPreferredWidth.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableWithPreferredWidth.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(PreferredWidthType::Percent, table->get_PreferredWidth()->get_Type());
        ASPOSE_ASSERT_EQ(50, table->get_PreferredWidth()->get_Value());
    }

    void InsertCellsWithPreferredWidths()
    {
        //ExStart
        //ExFor:CellFormat.PreferredWidth
        //ExFor:PreferredWidth
        //ExFor:PreferredWidth.Auto
        //ExFor:PreferredWidth.Equals(PreferredWidth)
        //ExFor:PreferredWidth.Equals(System.Object)
        //ExFor:PreferredWidth.FromPoints
        //ExFor:PreferredWidth.FromPercent
        //ExFor:PreferredWidth.GetHashCode
        //ExFor:PreferredWidth.ToString
        //ExSummary:Shows how to set a preferred width for table cells.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<Table> table = builder->StartTable();

        // There are two ways of applying the "PreferredWidth" class to table cells.
        // 1 -  Set an absolute preferred width based on points:
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(40));
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightYellow());
        builder->Writeln(String::Format(u"Cell with a width of {0}.", builder->get_CellFormat()->get_PreferredWidth()));

        // 2 -  Set a relative preferred width based on percent of the table's width:
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(20));
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
        builder->Writeln(String::Format(u"Cell with a width of {0}.", builder->get_CellFormat()->get_PreferredWidth()));

        builder->InsertCell();

        // A cell with no preferred width specified will take up the rest of the available space.
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::Auto());

        // Each configuration of the "PreferredWidth" property creates a new object.
        ASSERT_NE(System::ObjectExt::GetHashCode(table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()),
                  System::ObjectExt::GetHashCode(builder->get_CellFormat()->get_PreferredWidth()));

        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightGreen());
        builder->Writeln(u"Automatically sized cell.");

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertCellsWithPreferredWidths.docx");
        //ExEnd

        ASPOSE_ASSERT_EQ(100.0, PreferredWidth::FromPercent(100)->get_Value());
        ASPOSE_ASSERT_EQ(100.0, PreferredWidth::FromPoints(100)->get_Value());

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertCellsWithPreferredWidths.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(PreferredWidthType::Points, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_PreferredWidth()->get_Type());
        ASPOSE_ASSERT_EQ(40.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_PreferredWidth()->get_Value());
        ASSERT_EQ(u"Cell with a width of 800.\r\a", table->get_FirstRow()->get_Cells()->idx_get(0)->GetText().Trim());

        ASSERT_EQ(PreferredWidthType::Percent, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()->get_Type());
        ASPOSE_ASSERT_EQ(20.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()->get_Value());
        ASSERT_EQ(u"Cell with a width of 20%.\r\a", table->get_FirstRow()->get_Cells()->idx_get(1)->GetText().Trim());

        ASSERT_EQ(PreferredWidthType::Auto, table->get_FirstRow()->get_Cells()->idx_get(2)->get_CellFormat()->get_PreferredWidth()->get_Type());
        ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(2)->get_CellFormat()->get_PreferredWidth()->get_Value());
        ASSERT_EQ(u"Automatically sized cell.\r\a", table->get_FirstRow()->get_Cells()->idx_get(2)->GetText().Trim());
    }

    void InsertTableFromHtml()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
        // inserted from HTML.
        builder->InsertHtml(String(u"<table>") + u"<tr>" + u"<td>Row 1, Cell 1</td>" + u"<td>Row 1, Cell 2</td>" + u"</tr>" + u"<tr>" +
                            u"<td>Row 2, Cell 2</td>" + u"<td>Row 2, Cell 2</td>" + u"</tr>" + u"</table>");

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableFromHtml.docx");

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableFromHtml.docx");

        ASSERT_EQ(1, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Row, true)->get_Count());
        ASSERT_EQ(4, doc->GetChildNodes(NodeType::Cell, true)->get_Count());
    }

    void InsertNestedTable()
    {
        //ExStart
        //ExFor:Cell.FirstParagraph
        //ExSummary:Shows how to create a nested table using a document builder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Build the outer table.
        SharedPtr<Cell> cell = builder->InsertCell();
        builder->Writeln(u"Outer Table Cell 1");
        builder->InsertCell();
        builder->Writeln(u"Outer Table Cell 2");
        builder->EndTable();

        // Move to the first cell of the outer table, the build another table inside the cell.
        builder->MoveTo(cell->get_FirstParagraph());
        builder->InsertCell();
        builder->Writeln(u"Inner Table Cell 1");
        builder->InsertCell();
        builder->Writeln(u"Inner Table Cell 2");
        builder->EndTable();

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertNestedTable.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertNestedTable.docx");

        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        ASSERT_EQ(4, doc->GetChildNodes(NodeType::Cell, true)->get_Count());
        ASSERT_EQ(1, cell->get_Tables()->idx_get(0)->get_Count());
        ASSERT_EQ(2, cell->get_Tables()->idx_get(0)->get_FirstRow()->get_Cells()->get_Count());
    }

    void CreateTable()
    {
        //ExStart
        //ExFor:DocumentBuilder
        //ExFor:DocumentBuilder.Write
        //ExFor:DocumentBuilder.InsertCell
        //ExSummary:Shows how to use a document builder to create a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Start the table, then populate the first row with two cells.
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2.");

        // Call the builder's "EndRow" method to start a new row.
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 2.");
        builder->EndTable();

        doc->Save(ArtifactsDir + u"DocumentBuilder.CreateTable.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.CreateTable.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(4, table->GetChildNodes(NodeType::Cell, true)->get_Count());

        ASSERT_EQ(u"Row 1, Cell 1.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
        ASSERT_EQ(u"Row 1, Cell 2.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());
        ASSERT_EQ(u"Row 2, Cell 1.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
        ASSERT_EQ(u"Row 2, Cell 2.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
    }

    void BuildFormattedTable()
    {
        //ExStart
        //ExFor:RowFormat.Height
        //ExFor:RowFormat.HeightRule
        //ExFor:Table.LeftIndent
        //ExFor:DocumentBuilder.ParagraphFormat
        //ExFor:DocumentBuilder.Font
        //ExSummary:Shows how to create a formatted table using DocumentBuilder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        table->set_LeftIndent(20);

        // Set some formatting options for text and table appearance.
        builder->get_RowFormat()->set_Height(40);
        builder->get_RowFormat()->set_HeightRule(HeightRule::AtLeast);
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::FromArgb(198, 217, 241));

        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->get_Font()->set_Size(16);
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);

        // Configuring the formatting options in a document builder will apply them
        // to the current cell/row its cursor is in,
        // as well as any new cells and rows created using that builder.
        builder->Write(u"Header Row,\n Cell 1");
        builder->InsertCell();
        builder->Write(u"Header Row,\n Cell 2");
        builder->InsertCell();
        builder->Write(u"Header Row,\n Cell 3");
        builder->EndRow();

        // Reconfigure the builder's formatting objects for new rows and cells that we are about to make.
        // The builder will not apply these to the first row already created so that it will stand out as a header row.
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_White());
        builder->get_CellFormat()->set_VerticalAlignment(CellVerticalAlignment::Center);
        builder->get_RowFormat()->set_Height(30);
        builder->get_RowFormat()->set_HeightRule(HeightRule::Auto);
        builder->InsertCell();
        builder->get_Font()->set_Size(12);
        builder->get_Font()->set_Bold(false);

        builder->Write(u"Row 1, Cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2.");
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 3.");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 2.");
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 3.");
        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"DocumentBuilder.CreateFormattedTable.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.CreateFormattedTable.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASPOSE_ASSERT_EQ(20.0, table->get_LeftIndent());

        ASSERT_EQ(HeightRule::AtLeast, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
        ASPOSE_ASSERT_EQ(40.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());

        for (const auto& c : System::IterateOver<Cell>(doc->GetChildNodes(NodeType::Cell, true)))
        {
            ASSERT_EQ(ParagraphAlignment::Center, c->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());

            for (const auto& r : System::IterateOver<Run>(c->get_FirstParagraph()->get_Runs()))
            {
                ASSERT_EQ(u"Arial", r->get_Font()->get_Name());

                if (c->get_ParentRow() == table->get_FirstRow())
                {
                    ASPOSE_ASSERT_EQ(16, r->get_Font()->get_Size());
                    ASSERT_TRUE(r->get_Font()->get_Bold());
                }
                else
                {
                    ASPOSE_ASSERT_EQ(12, r->get_Font()->get_Size());
                    ASSERT_FALSE(r->get_Font()->get_Bold());
                }
            }
        }
    }

    void TableBordersAndShading()
    {
        //ExStart
        //ExFor:Shading
        //ExFor:Table.SetBorders
        //ExFor:BorderCollection.Left
        //ExFor:BorderCollection.Right
        //ExFor:BorderCollection.Top
        //ExFor:BorderCollection.Bottom
        //ExSummary:Shows how to apply border and shading color while building a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Start a table and set a default color/thickness for its borders.
        SharedPtr<Table> table = builder->StartTable();
        table->SetBorders(LineStyle::Single, 2.0, System::Drawing::Color::get_Black());

        // Create a row with two cells with different background colors.
        builder->InsertCell();
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightSkyBlue());
        builder->Writeln(u"Row 1, Cell 1.");
        builder->InsertCell();
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Orange());
        builder->Writeln(u"Row 1, Cell 2.");
        builder->EndRow();

        // Reset cell formatting to disable the background colors
        // set a custom border thickness for all new cells created by the builder,
        // then build a second row.
        builder->get_CellFormat()->ClearFormatting();
        builder->get_CellFormat()->get_Borders()->get_Left()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Right()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Top()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Bottom()->set_LineWidth(4.0);

        builder->InsertCell();
        builder->Writeln(u"Row 2, Cell 1.");
        builder->InsertCell();
        builder->Writeln(u"Row 2, Cell 2.");

        doc->Save(ArtifactsDir + u"DocumentBuilder.TableBordersAndShading.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.TableBordersAndShading.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        for (const auto& c : System::IterateOver<Cell>(table->get_FirstRow()))
        {
            ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Top()->get_LineWidth());
            ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Bottom()->get_LineWidth());
            ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Left()->get_LineWidth());
            ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Right()->get_LineWidth());

            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
            ASSERT_EQ(LineStyle::Single, c->get_CellFormat()->get_Borders()->get_Left()->get_LineStyle());
        }

        ASSERT_EQ(System::Drawing::Color::get_LightSkyBlue().ToArgb(),
                  table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(),
                  table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());

        for (const auto& c : System::IterateOver<Cell>(table->get_LastRow()))
        {
            ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Top()->get_LineWidth());
            ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Bottom()->get_LineWidth());
            ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Left()->get_LineWidth());
            ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Right()->get_LineWidth());

            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
            ASSERT_EQ(LineStyle::Single, c->get_CellFormat()->get_Borders()->get_Left()->get_LineStyle());
            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
        }
    }

    void SetPreferredTypeConvertUtil()
    {
        //ExStart
        //ExFor:PreferredWidth.FromPoints
        //ExSummary:Shows how to use unit conversion tools while specifying a preferred width for a cell.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(ConvertUtil::InchToPoint(3)));
        builder->InsertCell();

        ASPOSE_ASSERT_EQ(216.0, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_PreferredWidth()->get_Value());
        //ExEnd
    }

    void InsertHyperlinkToLocalBookmark()
    {
        //ExStart
        //ExFor:DocumentBuilder.StartBookmark
        //ExFor:DocumentBuilder.EndBookmark
        //ExFor:DocumentBuilder.InsertHyperlink
        //ExSummary:Shows how to insert a hyperlink which references a local bookmark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"Bookmark1");
        builder->Write(u"Bookmarked text. ");
        builder->EndBookmark(u"Bookmark1");
        builder->Writeln(u"Text outside of the bookmark.");

        // Insert a HYPERLINK field that links to the bookmark. We can pass field switches
        // to the "InsertHyperlink" method as part of the argument containing the referenced bookmark's name.
        builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        builder->get_Font()->set_Underline(Underline::Single);
        builder->InsertHyperlink(u"Link to Bookmark1", u"Bookmark1\" \\o \"Hyperlink Tip", true);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
        auto hyperlink = System::DynamicCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldHyperlink, u" HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ", u"Link to Bookmark1", hyperlink);
        ASSERT_EQ(u"Bookmark1", hyperlink->get_SubAddress());
        ASSERT_EQ(u"Hyperlink Tip", hyperlink->get_ScreenTip());
        ASSERT_TRUE(doc->get_Range()->get_Bookmarks()->LINQ_Any(static_cast<System::Func<SharedPtr<Bookmark>, bool>>(
            static_cast<std::function<bool(SharedPtr<Bookmark> b)>>([](SharedPtr<Bookmark> b) -> bool { return b->get_Name() == u"Bookmark1"; }))));
    }

    void CursorPosition()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Hello world!");

        // If the builder's cursor is at the end of the document,
        // there will be no nodes in front of it so that the current node will be null.
        ASSERT_TRUE(builder->get_CurrentNode() == nullptr);

        ASSERT_EQ(u"Hello world!", builder->get_CurrentParagraph()->GetText().Trim());

        // Move to the beginning of the document and place the cursor at an existing node.
        builder->MoveToDocumentStart();
        ASSERT_EQ(NodeType::Run, builder->get_CurrentNode()->get_NodeType());
    }

    void MoveTo()
    {
        //ExStart
        //ExFor:Story.LastParagraph
        //ExFor:DocumentBuilder.MoveTo(Node)
        //ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Run 1. ");

        // The document builder has a cursor, which acts as the part of the document
        // where the builder appends new nodes when we use its document construction methods.
        // This cursor functions in the same way as Microsoft Word's blinking cursor,
        // and it also always ends up immediately after any node that the builder just inserted.
        // To append content to a different part of the document,
        // we can move the cursor to a different node with the "MoveTo" method.
        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_LastParagraph(), builder->get_CurrentParagraph());
        //ExSkip
        builder->MoveTo(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0));
        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), builder->get_CurrentParagraph());
        //ExSkip

        // The cursor is now in front of the node that we moved it to.
        // Adding a second run will insert it in front of the first run.
        builder->Writeln(u"Run 2. ");

        ASSERT_EQ(u"Run 2. \rRun 1.", doc->GetText().Trim());

        // Move the cursor to the end of the document to continue appending text to the end as before.
        builder->MoveTo(doc->get_LastSection()->get_Body()->get_LastParagraph());
        builder->Writeln(u"Run 3. ");

        ASSERT_EQ(u"Run 2. \rRun 1. \rRun 3.", doc->GetText().Trim());
        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_LastParagraph(), builder->get_CurrentParagraph());
        //ExSkip
        //ExEnd
    }

    void MoveToParagraph()
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToParagraph
        //ExSummary:Shows how to move a builder's cursor position to a specified paragraph.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(22, paragraphs->get_Count());

        // Create document builder to edit the document. The builder's cursor,
        // which is the point where it will insert new nodes when we call its document construction methods,
        // is currently at the beginning of the document.
        auto builder = MakeObject<DocumentBuilder>(doc);

        ASSERT_EQ(0, paragraphs->IndexOf(builder->get_CurrentParagraph()));

        // Move that cursor to a different paragraph will place that cursor in front of that paragraph.
        builder->MoveToParagraph(2, 0);
        ASSERT_EQ(2, paragraphs->IndexOf(builder->get_CurrentParagraph()));
        //ExSkip

        // Any new content that we add will be inserted at that point.
        builder->Writeln(u"This is a new third paragraph. ");
        //ExEnd

        ASSERT_EQ(3, paragraphs->IndexOf(builder->get_CurrentParagraph()));

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_EQ(u"This is a new third paragraph.", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(2)->GetText().Trim());
    }

    void MoveToCell()
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToCell
        //ExSummary:Shows how to move a document builder's cursor to a cell in a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an empty 2x2 table.
        builder->StartTable();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndRow();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndTable();

        // Because we have ended the table with the EndTable method,
        // the document builder's cursor is currently outside the table.
        // This cursor has the same function as Microsoft Word's blinking text cursor.
        // It can also be moved to a different location in the document using the builder's MoveTo methods.
        // We can move the cursor back inside the table to a specific cell.
        builder->MoveToCell(0, 1, 1, 0);
        builder->Write(u"Column 2, cell 2.");

        doc->Save(ArtifactsDir + u"DocumentBuilder.MoveToCell.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MoveToCell.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(u"Column 2, cell 2.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
    }

    void MoveToBookmark()
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
        //ExSummary:Shows how to move a document builder's node insertion point cursor to a bookmark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // A valid bookmark consists of a BookmarkStart node, a BookmarkEnd node with a
        // matching bookmark name somewhere afterward, and contents enclosed by those nodes.
        builder->StartBookmark(u"MyBookmark");
        builder->Write(u"Hello world! ");
        builder->EndBookmark(u"MyBookmark");

        // There are 4 ways of moving a document builder's cursor to a bookmark.
        // If we are between the BookmarkStart and BookmarkEnd nodes, the cursor will be inside the bookmark.
        // This means that any text added by the builder will become a part of the bookmark.
        // 1 -  Outside of the bookmark, in front of the BookmarkStart node:
        ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", true, false));
        builder->Write(u"1. ");

        ASSERT_EQ(u"Hello world! ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
        ASSERT_EQ(u"1. Hello world!", doc->GetText().Trim());

        // 2 -  Inside the bookmark, right after the BookmarkStart node:
        ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", true, true));
        builder->Write(u"2. ");

        ASSERT_EQ(u"2. Hello world! ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
        ASSERT_EQ(u"1. 2. Hello world!", doc->GetText().Trim());

        // 2 -  Inside the bookmark, right in front of the BookmarkEnd node:
        ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", false, false));
        builder->Write(u"3. ");

        ASSERT_EQ(u"2. Hello world! 3. ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
        ASSERT_EQ(u"1. 2. Hello world! 3.", doc->GetText().Trim());

        // 4 -  Outside of the bookmark, after the BookmarkEnd node:
        ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", false, true));
        builder->Write(u"4.");

        ASSERT_EQ(u"2. Hello world! 3. ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
        ASSERT_EQ(u"1. 2. Hello world! 3. 4.", doc->GetText().Trim());
        //ExEnd
    }

    void BuildTable()
    {
        //ExStart
        //ExFor:Table
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.EndRow
        //ExFor:DocumentBuilder.EndTable
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:DocumentBuilder.Write(String)
        //ExFor:DocumentBuilder.Writeln(String)
        //ExFor:CellVerticalAlignment
        //ExFor:CellFormat.Orientation
        //ExFor:TextOrientation
        //ExFor:AutoFitBehavior
        //ExSummary:Shows how to build a formatted 2x2 table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->get_CellFormat()->set_VerticalAlignment(CellVerticalAlignment::Center);
        builder->Write(u"Row 1, cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 1, cell 2.");
        builder->EndRow();

        // While building the table, the document builder will apply its current RowFormat/CellFormat property values
        // to the current row/cell that its cursor is in and any new rows/cells as it creates them.
        ASSERT_EQ(CellVerticalAlignment::Center, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalAlignment());
        ASSERT_EQ(CellVerticalAlignment::Center, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->get_CellFormat()->get_VerticalAlignment());

        builder->InsertCell();
        builder->get_RowFormat()->set_Height(100);
        builder->get_RowFormat()->set_HeightRule(HeightRule::Exactly);
        builder->get_CellFormat()->set_Orientation(TextOrientation::Upward);
        builder->Write(u"Row 2, cell 1.");
        builder->InsertCell();
        builder->get_CellFormat()->set_Orientation(TextOrientation::Downward);
        builder->Write(u"Row 2, cell 2.");
        builder->EndRow();
        builder->EndTable();

        // Previously added rows and cells are not retroactively affected by changes to the builder's formatting.
        ASPOSE_ASSERT_EQ(0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
        ASPOSE_ASSERT_EQ(100, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());
        ASSERT_EQ(TextOrientation::Upward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->get_CellFormat()->get_Orientation());
        ASSERT_EQ(TextOrientation::Downward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->get_CellFormat()->get_Orientation());

        doc->Save(ArtifactsDir + u"DocumentBuilder.BuildTable.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.BuildTable.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(2, table->get_Rows()->get_Count());
        ASSERT_EQ(2, table->get_Rows()->idx_get(0)->get_Cells()->get_Count());
        ASSERT_EQ(2, table->get_Rows()->idx_get(1)->get_Cells()->get_Count());

        ASPOSE_ASSERT_EQ(0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
        ASPOSE_ASSERT_EQ(100, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());

        ASSERT_EQ(u"Row 1, cell 1.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
        ASSERT_EQ(CellVerticalAlignment::Center, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalAlignment());

        ASSERT_EQ(u"Row 1, cell 2.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());

        ASSERT_EQ(u"Row 2, cell 1.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
        ASSERT_EQ(TextOrientation::Upward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->get_CellFormat()->get_Orientation());

        ASSERT_EQ(u"Row 2, cell 2.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
        ASSERT_EQ(TextOrientation::Downward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->get_CellFormat()->get_Orientation());
    }

    void TableCellVerticalRotatedFarEastTextOrientation()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rotated cell text.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        SharedPtr<Cell> cell = table->get_FirstRow()->get_FirstCell();

        ASSERT_EQ(TextOrientation::VerticalRotatedFarEast, cell->get_CellFormat()->get_Orientation());

        doc = DocumentHelper::SaveOpen(doc);

        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        cell = table->get_FirstRow()->get_FirstCell();

        ASSERT_EQ(TextOrientation::VerticalRotatedFarEast, cell->get_CellFormat()->get_Orientation());
    }

    void InsertFloatingImage()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // There are two ways of using a document builder to source an image and then insert it as a floating shape.
        // 1 -  From a file in the local file system:
        builder->InsertImage(ImageDir + u"Transparent background logo.png", RelativeHorizontalPosition::Margin, 100.0, RelativeVerticalPosition::Margin, 0.0,
                             200.0, 200.0, WrapType::Square);

        // 2 -  From a URL:
        builder->InsertImage(AsposeLogoUrl, RelativeHorizontalPosition::Margin, 100.0, RelativeVerticalPosition::Margin, 250.0, 200.0, 200.0, WrapType::Square);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertFloatingImage.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertFloatingImage.docx");
        auto image = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Png, image);
        ASPOSE_ASSERT_EQ(100.0, image->get_Left());
        ASPOSE_ASSERT_EQ(0.0, image->get_Top());
        ASPOSE_ASSERT_EQ(200.0, image->get_Width());
        ASPOSE_ASSERT_EQ(200.0, image->get_Height());
        ASSERT_EQ(WrapType::Square, image->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, image->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, image->get_RelativeVerticalPosition());

        image = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyImageInShape(320, 320, ImageType::Png, image);
        ASPOSE_ASSERT_EQ(100.0, image->get_Left());
        ASPOSE_ASSERT_EQ(250.0, image->get_Top());
        ASPOSE_ASSERT_EQ(200.0, image->get_Width());
        ASPOSE_ASSERT_EQ(200.0, image->get_Height());
        ASSERT_EQ(WrapType::Square, image->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, image->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, image->get_RelativeVerticalPosition());
    }

    void InsertImageOriginalSize()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from the local file system into a document while preserving its dimensions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // The InsertImage method creates a floating shape with the passed image in its image data.
        // We can specify the dimensions of the shape can be passing them to this method.
        SharedPtr<Shape> imageShape = builder->InsertImage(ImageDir + u"Logo.jpg", RelativeHorizontalPosition::Margin, 0.0, RelativeVerticalPosition::Margin,
                                                           0.0, -1.0, -1.0, WrapType::Square);

        // Passing negative values as the intended dimensions will automatically define
        // the shape's dimensions based on the dimensions of its image.
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertImageOriginalSize.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertImageOriginalSize.docx");
        imageShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
        ASSERT_EQ(WrapType::Square, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());
    }

    void InsertTextInput()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert a text input form field into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a form that prompts the user to enter text.
        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Enter your text here", 0);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTextInput.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTextInput.docx");
        SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_TRUE(formField->get_Enabled());
        ASSERT_EQ(u"TextInput", formField->get_Name());
        ASSERT_EQ(0, formField->get_MaxLength());
        ASSERT_EQ(u"Enter your text here", formField->get_Result());
        ASSERT_EQ(FieldType::FieldFormTextInput, formField->get_Type());
        ASSERT_EQ(u"", formField->get_TextInputFormat());
        ASSERT_EQ(TextFormFieldType::Regular, formField->get_TextInputType());
    }

    void InsertComboBox()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertComboBox
        //ExSummary:Shows how to insert a combo box form field into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a form that prompts the user to pick one of the items from the menu.
        builder->Write(u"Pick a fruit: ");
        ArrayPtr<String> items = MakeArray<String>({u"Apple", u"Banana", u"Cherry"});
        builder->InsertComboBox(u"DropDown", items, 0);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertComboBox.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertComboBox.docx");
        SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_TRUE(formField->get_Enabled());
        ASSERT_EQ(u"DropDown", formField->get_Name());
        ASSERT_EQ(0, formField->get_DropDownSelectedIndex());
        ASPOSE_ASSERT_EQ(items, formField->get_DropDownItems());
        ASSERT_EQ(FieldType::FieldFormDropDown, formField->get_Type());
    }

    void SignatureLineProviderId()
    {
        //ExStart
        //ExFor:SignatureLine.IsSigned
        //ExFor:SignatureLine.IsValid
        //ExFor:SignatureLine.ProviderId
        //ExFor:SignatureLineOptions.ShowDate
        //ExFor:SignatureLineOptions.Email
        //ExFor:SignatureLineOptions.DefaultInstructions
        //ExFor:SignatureLineOptions.Instructions
        //ExFor:SignatureLineOptions.AllowComments
        //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
        //ExFor:SignOptions.ProviderId
        //ExSummary:Shows how to sign a document with a personal certificate and a signature line.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto signatureLineOptions = MakeObject<SignatureLineOptions>();
        signatureLineOptions->set_Signer(u"vderyushev");
        signatureLineOptions->set_SignerTitle(u"QA");
        signatureLineOptions->set_Email(u"vderyushev@aspose.com");
        signatureLineOptions->set_ShowDate(true);
        signatureLineOptions->set_DefaultInstructions(false);
        signatureLineOptions->set_Instructions(u"Please sign here.");
        signatureLineOptions->set_AllowComments(true);

        SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(signatureLineOptions)->get_SignatureLine();
        signatureLine->set_ProviderId(System::Guid::Parse(u"CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));

        ASSERT_FALSE(signatureLine->get_IsSigned());
        ASSERT_FALSE(signatureLine->get_IsValid());

        doc->Save(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.docx");

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        signOptions->set_ProviderId(signatureLine->get_ProviderId());
        signOptions->set_Comments(u"Document was signed by vderyushev");
        signOptions->set_SignTime(System::DateTime::get_Now());

        SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        DigitalSignatureUtil::Sign(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.docx",
                                   ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.Signed.docx", certHolder, signOptions);

        // Re-open our saved document, and verify that the "IsSigned" and "IsValid" properties both equal "true",
        // indicating that the signature line contains a signature.
        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.Signed.docx");
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        signatureLine = shape->get_SignatureLine();

        ASSERT_TRUE(signatureLine->get_IsSigned());
        ASSERT_TRUE(signatureLine->get_IsValid());
        //ExEnd

        ASSERT_EQ(u"vderyushev", signatureLine->get_Signer());
        ASSERT_EQ(u"QA", signatureLine->get_SignerTitle());
        ASSERT_EQ(u"vderyushev@aspose.com", signatureLine->get_Email());
        ASSERT_TRUE(signatureLine->get_ShowDate());
        ASSERT_FALSE(signatureLine->get_DefaultInstructions());
        ASSERT_EQ(u"Please sign here.", signatureLine->get_Instructions());
        ASSERT_TRUE(signatureLine->get_AllowComments());
        ASSERT_TRUE(signatureLine->get_IsSigned());
        ASSERT_TRUE(signatureLine->get_IsValid());

        SharedPtr<DigitalSignatureCollection> signatures =
            DigitalSignatureUtil::LoadSignatures(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.Signed.docx");

        ASSERT_EQ(1, signatures->get_Count());
        ASSERT_TRUE(signatures->idx_get(0)->get_IsValid());
        ASSERT_EQ(u"Document was signed by vderyushev", signatures->idx_get(0)->get_Comments());
        ASSERT_EQ(System::DateTime::get_Today(), signatures->idx_get(0)->get_SignTime().get_Date());
        ASSERT_EQ(u"CN=Morzal.Me", signatures->idx_get(0)->get_IssuerName());
        ASSERT_EQ(DigitalSignatureType::XmlDsig, signatures->idx_get(0)->get_SignatureType());
    }

    void SignatureLineInline()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
        //ExSummary:Shows how to insert an inline signature line into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto options = MakeObject<SignatureLineOptions>();
        options->set_Signer(u"John Doe");
        options->set_SignerTitle(u"Manager");
        options->set_Email(u"johndoe@aspose.com");
        options->set_ShowDate(true);
        options->set_DefaultInstructions(false);
        options->set_Instructions(u"Please sign here.");
        options->set_AllowComments(true);

        builder->InsertSignatureLine(options, RelativeHorizontalPosition::RightMargin, 2.0, RelativeVerticalPosition::Page, 3.0, WrapType::Inline);

        // The signature line can be signed in Microsoft Word by double clicking it.
        doc->Save(ArtifactsDir + u"DocumentBuilder.SignatureLineInline.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SignatureLineInline.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        SharedPtr<SignatureLine> signatureLine = shape->get_SignatureLine();

        ASSERT_EQ(u"John Doe", signatureLine->get_Signer());
        ASSERT_EQ(u"Manager", signatureLine->get_SignerTitle());
        ASSERT_EQ(u"johndoe@aspose.com", signatureLine->get_Email());
        ASSERT_TRUE(signatureLine->get_ShowDate());
        ASSERT_FALSE(signatureLine->get_DefaultInstructions());
        ASSERT_EQ(u"Please sign here.", signatureLine->get_Instructions());
        ASSERT_TRUE(signatureLine->get_AllowComments());
        ASSERT_FALSE(signatureLine->get_IsSigned());
        ASSERT_FALSE(signatureLine->get_IsValid());
    }

    void SetParagraphFormatting()
    {
        //ExStart
        //ExFor:ParagraphFormat.RightIndent
        //ExFor:ParagraphFormat.LeftIndent
        //ExSummary:Shows how to configure paragraph formatting to create off-center text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Center all text that the document builder writes, and set up indents.
        // The indent configuration below will create a body of text that will sit asymmetrically on the page.
        // The "center" that we align the text to will be the middle of the body of text, not the middle of the page.
        SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
        paragraphFormat->set_Alignment(ParagraphAlignment::Center);
        paragraphFormat->set_LeftIndent(100);
        paragraphFormat->set_RightIndent(50);
        paragraphFormat->set_SpaceAfter(25);

        builder->Writeln(u"This paragraph demonstrates how left and right indentation affects word wrapping.");
        builder->Writeln(u"The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.");

        doc->Save(ArtifactsDir + u"DocumentBuilder.SetParagraphFormatting.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SetParagraphFormatting.docx");

        for (const auto& paragraph : System::IterateOver<Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
        {
            ASSERT_EQ(ParagraphAlignment::Center, paragraph->get_ParagraphFormat()->get_Alignment());
            ASPOSE_ASSERT_EQ(100.0, paragraph->get_ParagraphFormat()->get_LeftIndent());
            ASPOSE_ASSERT_EQ(50.0, paragraph->get_ParagraphFormat()->get_RightIndent());
            ASPOSE_ASSERT_EQ(25.0, paragraph->get_ParagraphFormat()->get_SpaceAfter());
        }
    }

    void SetCellFormatting()
    {
        //ExStart
        //ExFor:DocumentBuilder.CellFormat
        //ExFor:CellFormat.Width
        //ExFor:CellFormat.LeftPadding
        //ExFor:CellFormat.RightPadding
        //ExFor:CellFormat.TopPadding
        //ExFor:CellFormat.BottomPadding
        //ExFor:DocumentBuilder.StartTable
        //ExFor:DocumentBuilder.EndTable
        //ExSummary:Shows how to format cells with a document builder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, cell 1.");

        // Insert a second cell, and then configure cell text padding options.
        // The builder will apply these settings at its current cell, and any new cells creates afterwards.
        builder->InsertCell();

        SharedPtr<CellFormat> cellFormat = builder->get_CellFormat();
        cellFormat->set_Width(250);
        cellFormat->set_LeftPadding(30);
        cellFormat->set_RightPadding(30);
        cellFormat->set_TopPadding(30);
        cellFormat->set_BottomPadding(30);

        builder->Write(u"Row 1, cell 2.");
        builder->EndRow();
        builder->EndTable();

        // The first cell was unaffected by the padding reconfiguration, and still holds the default values.
        ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width());
        ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_LeftPadding());
        ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_RightPadding());
        ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_TopPadding());
        ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_BottomPadding());

        ASPOSE_ASSERT_EQ(250.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_Width());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_LeftPadding());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_RightPadding());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_TopPadding());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_BottomPadding());

        // The first cell will still grow in the output document to match the size of its neighboring cell.
        doc->Save(ArtifactsDir + u"DocumentBuilder.SetCellFormatting.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SetCellFormatting.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASPOSE_ASSERT_EQ(159.3, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width());
        ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_LeftPadding());
        ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_RightPadding());
        ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_TopPadding());
        ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_BottomPadding());

        ASPOSE_ASSERT_EQ(310.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_Width());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_LeftPadding());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_RightPadding());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_TopPadding());
        ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_BottomPadding());
    }

    void SetRowFormatting()
    {
        //ExStart
        //ExFor:DocumentBuilder.RowFormat
        //ExFor:HeightRule
        //ExFor:RowFormat.Height
        //ExFor:RowFormat.HeightRule
        //ExSummary:Shows how to format rows with a document builder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, cell 1.");

        // Start a second row, and then configure its height. The builder will apply these settings to
        // its current row, as well as any new rows it creates afterwards.
        builder->EndRow();

        SharedPtr<RowFormat> rowFormat = builder->get_RowFormat();
        rowFormat->set_Height(100);
        rowFormat->set_HeightRule(HeightRule::Exactly);

        builder->InsertCell();
        builder->Write(u"Row 2, cell 1.");
        builder->EndTable();

        // The first row was unaffected by the padding reconfiguration and still holds the default values.
        ASPOSE_ASSERT_EQ(0.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());

        ASPOSE_ASSERT_EQ(100.0, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());

        doc->Save(ArtifactsDir + u"DocumentBuilder.SetRowFormatting.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SetRowFormatting.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASPOSE_ASSERT_EQ(0.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());

        ASPOSE_ASSERT_EQ(100.0, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
        ASSERT_EQ(HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());
    }

    void InsertFootnote()
    {
        //ExStart
        //ExFor:FootnoteType
        //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String)
        //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String,String)
        //ExSummary:Shows how to reference text with a footnote and an endnote.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert some text and mark it with a footnote with the IsAuto property set to "true" by default,
        // so the marker seen in the body text will be auto-numbered at "1",
        // and the footnote will appear at the bottom of the page.
        builder->Write(u"This text will be referenced by a footnote.");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote comment regarding referenced text.");

        // Insert more text and mark it with an endnote with a custom reference mark,
        // which will be used in place of the number "2" and set "IsAuto" to false.
        builder->Write(u"This text will be referenced by an endnote.");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote comment regarding referenced text.", u"CustomMark");

        // Footnotes always appear at the bottom of their referenced text,
        // so this page break will not affect the footnote.
        // On the other hand, endnotes are always at the end of the document
        // so that this page break will push the endnote down to the next page.
        builder->InsertBreak(BreakType::PageBreak);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertFootnote.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertFootnote.docx");

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote comment regarding referenced text.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, false, u"CustomMark", u"CustomMark Endnote comment regarding referenced text.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
    }

    void ApplyBordersAndShading()
    {
        //ExStart
        //ExFor:BorderCollection.Item(BorderType)
        //ExFor:Shading
        //ExFor:TextureIndex
        //ExFor:ParagraphFormat.Shading
        //ExFor:Shading.Texture
        //ExFor:Shading.BackgroundPatternColor
        //ExFor:Shading.ForegroundPatternColor
        //ExSummary:Shows how to decorate text with borders and shading.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
        borders->set_DistanceFromText(20);
        borders->idx_get(BorderType::Left)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Right)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Top)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Bottom)->set_LineStyle(LineStyle::Double);

        SharedPtr<Shading> shading = builder->get_ParagraphFormat()->get_Shading();
        shading->set_Texture(TextureIndex::TextureDiagonalCross);
        shading->set_BackgroundPatternColor(System::Drawing::Color::get_LightCoral());
        shading->set_ForegroundPatternColor(System::Drawing::Color::get_LightSalmon());

        builder->Write(u"This paragraph is formatted with a double border and shading.");
        doc->Save(ArtifactsDir + u"DocumentBuilder.ApplyBordersAndShading.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.ApplyBordersAndShading.docx");
        borders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();

        ASPOSE_ASSERT_EQ(20.0, borders->get_DistanceFromText());
        ASSERT_EQ(LineStyle::Double, borders->idx_get(BorderType::Left)->get_LineStyle());
        ASSERT_EQ(LineStyle::Double, borders->idx_get(BorderType::Right)->get_LineStyle());
        ASSERT_EQ(LineStyle::Double, borders->idx_get(BorderType::Top)->get_LineStyle());
        ASSERT_EQ(LineStyle::Double, borders->idx_get(BorderType::Bottom)->get_LineStyle());

        ASSERT_EQ(TextureIndex::TextureDiagonalCross, shading->get_Texture());
        ASSERT_EQ(System::Drawing::Color::get_LightCoral().ToArgb(), shading->get_BackgroundPatternColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_LightSalmon().ToArgb(), shading->get_ForegroundPatternColor().ToArgb());
    }

    void DeleteRow()
    {
        //ExStart
        //ExFor:DocumentBuilder.DeleteRow
        //ExSummary:Shows how to delete a row from a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 1, cell 2.");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Row 2, cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 2, cell 2.");
        builder->EndTable();

        ASSERT_EQ(2, table->get_Rows()->get_Count());

        // Delete the first row of the first table in the document.
        builder->DeleteRow(0, 0);

        ASSERT_EQ(1, table->get_Rows()->get_Count());
        ASSERT_EQ(u"Row 2, cell 1.\aRow 2, cell 2.\a\a", table->GetText().Trim());
        //ExEnd
    }

    void AppendDocumentAndResolveStyles(bool keepSourceNumbering)
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to manage list style clashes while appending a document.
        // Load a document with text in a custom style and clone it.
        auto srcDoc = MakeObject<Document>(MyDir + u"Custom list numbering.docx");
        SharedPtr<Document> dstDoc = srcDoc->Clone();

        // We now have two documents, each with an identical style named "CustomStyle".
        // Change the text color for one of the styles to set it apart from the other.
        dstDoc->get_Styles()->idx_get(u"CustomStyle")->get_Font()->set_Color(System::Drawing::Color::get_DarkRed());

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        auto options = MakeObject<ImportFormatOptions>();
        options->set_KeepSourceNumbering(keepSourceNumbering);

        // Joining two documents that have different styles that share the same name causes a style clash.
        // We can specify an import format mode while appending documents to resolve this clash.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepDifferentStyles, options);
        dstDoc->UpdateListLabels();

        dstDoc->Save(ArtifactsDir + u"DocumentBuilder.AppendDocumentAndResolveStyles.docx");
        //ExEnd
    }

    void InsertDocumentAndResolveStyles(bool keepSourceNumbering)
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to manage list style clashes while inserting a document.
        auto dstDoc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(dstDoc);
        builder->InsertBreak(BreakType::ParagraphBreak);

        dstDoc->get_Lists()->Add(ListTemplate::NumberDefault);
        SharedPtr<Aspose::Words::Lists::List> list = dstDoc->get_Lists()->idx_get(0);

        builder->get_ListFormat()->set_List(list);

        for (int i = 1; i <= 15; i++)
        {
            builder->Write(String::Format(u"List Item {0}\n", i));
        }

        auto attachDoc = System::DynamicCast<Document>(dstDoc->Clone(true));

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        auto importOptions = MakeObject<ImportFormatOptions>();
        importOptions->set_KeepSourceNumbering(keepSourceNumbering);

        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->InsertDocument(attachDoc, ImportFormatMode::KeepSourceFormatting, importOptions);

        dstDoc->Save(ArtifactsDir + u"DocumentBuilder.InsertDocumentAndResolveStyles.docx");
        //ExEnd
    }

    void LoadDocumentWithListNumbering(bool keepSourceNumbering)
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to manage list style clashes while appending a clone of a document to itself.
        auto srcDoc = MakeObject<Document>(MyDir + u"List item.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"List item.docx");

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        auto builder = MakeObject<DocumentBuilder>(dstDoc);
        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        auto options = MakeObject<ImportFormatOptions>();
        options->set_KeepSourceNumbering(keepSourceNumbering);
        builder->InsertDocument(srcDoc, ImportFormatMode::KeepSourceFormatting, options);

        dstDoc->UpdateListLabels();
        //ExEnd
    }

    void IgnoreTextBoxes(bool ignoreTextBoxes)
    {
        //ExStart
        //ExFor:ImportFormatOptions.IgnoreTextBoxes
        //ExSummary:Shows how to manage text box formatting while appending a document.
        // Create a document that will have nodes from another document inserted into it.
        auto dstDoc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(dstDoc);

        builder->Writeln(u"Hello world!");

        // Create another document with a text box, which we will import into the first document.
        auto srcDoc = MakeObject<Document>();
        builder = MakeObject<DocumentBuilder>(srcDoc);

        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 300, 100);
        builder->MoveTo(textBox->get_FirstParagraph());
        builder->get_ParagraphFormat()->get_Style()->get_Font()->set_Name(u"Courier New");
        builder->get_ParagraphFormat()->get_Style()->get_Font()->set_Size(24);
        builder->Write(u"Textbox contents");

        // Set a flag to specify whether to clear or preserve text box formatting
        // while importing them to other documents.
        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_IgnoreTextBoxes(ignoreTextBoxes);

        // Import the text box from the source document into the destination document,
        // and then verify whether we have preserved the styling of its text contents.
        auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting, importFormatOptions);
        auto importedTextBox = System::DynamicCast<Shape>(importer->ImportNode(textBox, true));
        dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->AppendChild(importedTextBox);

        if (ignoreTextBoxes)
        {
            ASPOSE_ASSERT_EQ(12.0, importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Size());
            ASSERT_EQ(u"Times New Roman", importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
        }
        else
        {
            ASPOSE_ASSERT_EQ(24.0, importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Size());
            ASSERT_EQ(u"Courier New", importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
        }

        dstDoc->Save(ArtifactsDir + u"DocumentBuilder.IgnoreTextBoxes.docx");
        //ExEnd
    }

    void MoveToField(bool moveCursorToAfterTheField)
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToField
        //ExSummary:Shows how to move a document builder's node insertion point cursor to a specific field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a field using the DocumentBuilder and add a run of text after it.
        SharedPtr<Field> field = builder->InsertField(u" AUTHOR \"John Doe\" ");

        // The builder's cursor is currently at end of the document.
        ASSERT_TRUE(builder->get_CurrentNode() == nullptr);

        // Move the cursor to the field while specifying whether to place that cursor before or after the field.
        builder->MoveToField(field, moveCursorToAfterTheField);

        // Note that the cursor is outside of the field in both cases.
        // This means that we cannot edit the field using the builder like this.
        // To edit a field, we can use the builder's MoveTo method on a field's FieldStart
        // or FieldSeparator node to place the cursor inside.
        if (moveCursorToAfterTheField)
        {
            ASSERT_TRUE(builder->get_CurrentNode() == nullptr);
            builder->Write(u" Text immediately after the field.");

            ASSERT_EQ(u"\u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015 Text immediately after the field.", doc->GetText().Trim());
        }
        else
        {
            ASPOSE_ASSERT_EQ(field->get_Start(), builder->get_CurrentNode());
            builder->Write(u"Text immediately before the field. ");

            ASSERT_EQ(u"Text immediately before the field. \u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015", doc->GetText().Trim());
        }
        //ExEnd
    }

    void InsertOleObjectException()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        ASSERT_THROW(static_cast<std::function<SharedPtr<Shape>()>>(
                         [&builder]() -> SharedPtr<Shape> { return builder->InsertOleObject(u"", u"checkbox", false, true, nullptr); })(),
                     System::ArgumentException);
    }

    void InsertPieChart()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
        //ExSummary:Shows how to insert a pie chart into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Chart> chart = builder->InsertChart(ChartType::Pie, ConvertUtil::PixelToPoint(300), ConvertUtil::PixelToPoint(300))->get_Chart();
        ASPOSE_ASSERT_EQ(225.0, ConvertUtil::PixelToPoint(300));
        //ExSkip
        chart->get_Series()->Clear();
        chart->get_Series()->Add(u"My fruit", MakeArray<String>({u"Apples", u"Bananas", u"Cherries"}), MakeArray<double>({1.3, 2.2, 1.5}));

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertPieChart.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertPieChart.docx");
        auto chartShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(u"Chart Title", chartShape->get_Chart()->get_Title()->get_Text());
        ASPOSE_ASSERT_EQ(225.0, chartShape->get_Width());
        ASPOSE_ASSERT_EQ(225.0, chartShape->get_Height());
    }

    void InsertChartRelativePosition()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to specify position and wrapping while inserting a chart.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertChart(ChartType::Pie, RelativeHorizontalPosition::Margin, 100, RelativeVerticalPosition::Margin, 100, 200, 100, WrapType::Square);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertedChartRelativePosition.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertedChartRelativePosition.docx");
        auto chartShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(100.0, chartShape->get_Top());
        ASPOSE_ASSERT_EQ(100.0, chartShape->get_Left());
        ASPOSE_ASSERT_EQ(200.0, chartShape->get_Width());
        ASPOSE_ASSERT_EQ(100.0, chartShape->get_Height());
        ASSERT_EQ(WrapType::Square, chartShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, chartShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, chartShape->get_RelativeVerticalPosition());
    }

    void InsertField()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(String)
        //ExFor:Field
        //ExFor:Field.Result
        //ExFor:Field.GetFieldCode()
        //ExFor:Field.Type
        //ExFor:FieldType
        //ExSummary:Shows how to insert a field into a document using a field code.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Field> field = builder->InsertField(u"DATE \\@ \"dddd, MMMM dd, yyyy\"");

        ASSERT_EQ(FieldType::FieldDate, field->get_Type());
        ASSERT_EQ(u"DATE \\@ \"dddd, MMMM dd, yyyy\"", field->GetFieldCode());

        // This overload of the InsertField method automatically updates inserted fields.
        ASSERT_LE(System::Math::Abs((System::DateTime::Parse(field->get_Result()) - System::DateTime::get_Today()).get_Hours()), 24);
        //ExEnd
    }

    void InsertFieldAndUpdate(bool updateInsertedFieldsImmediately)
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
        //ExFor:Field.Update()
        //ExSummary:Shows how to insert a field into a document using FieldType.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert two fields while passing a flag which determines whether to update them as the builder inserts them.
        // In some cases, updating fields could be computationally expensive, and it may be a good idea to defer the update.
        doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
        builder->Write(u"This document was written by ");
        builder->InsertField(FieldType::FieldAuthor, updateInsertedFieldsImmediately);

        builder->InsertParagraph();
        builder->Write(u"\nThis is page ");
        builder->InsertField(FieldType::FieldPage, updateInsertedFieldsImmediately);

        ASSERT_EQ(u" AUTHOR ", doc->get_Range()->get_Fields()->idx_get(0)->GetFieldCode());
        ASSERT_EQ(u" PAGE ", doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode());

        if (updateInsertedFieldsImmediately)
        {
            ASSERT_EQ(u"John Doe", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
            ASSERT_EQ(u"1", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
        }
        else
        {
            ASSERT_EQ(String::Empty, doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
            ASSERT_EQ(String::Empty, doc->get_Range()->get_Fields()->idx_get(1)->get_Result());

            // We will need to update these fields using the update methods manually.
            doc->get_Range()->get_Fields()->idx_get(0)->Update();

            ASSERT_EQ(u"John Doe", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());

            doc->UpdateFields();

            ASSERT_EQ(u"1", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
        }
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_EQ(String(u"This document was written by \u0013 AUTHOR \u0014John Doe\u0015") + u"\r\rThis is page \u0013 PAGE \u00141\u0015",
                  doc->GetText().Trim());

        TestUtil::VerifyField(FieldType::FieldAuthor, u" AUTHOR ", u"John Doe", doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldPage, u" PAGE ", u"1", doc->get_Range()->get_Fields()->idx_get(1));
    }

    //ExStart
    //ExFor:IFieldResultFormatter
    //ExFor:IFieldResultFormatter.Format(Double, GeneralFormat)
    //ExFor:IFieldResultFormatter.Format(String, GeneralFormat)
    //ExFor:IFieldResultFormatter.FormatDateTime(DateTime, String, CalendarType)
    //ExFor:IFieldResultFormatter.FormatNumeric(Double, String)
    //ExFor:FieldOptions.ResultFormatter
    //ExFor:CalendarType
    //ExSummary:Shows how to automatically apply a custom format to field results as the fields are updated.
    void FieldResultFormatting_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto formatter = MakeObject<ExDocumentBuilder::FieldResultFormatter>(u"${0}", u"Date: {0}", u"Item # {0}:");
        doc->get_FieldOptions()->set_ResultFormatter(formatter);

        // Our field result formatter applies a custom format to newly created fields of three types of formats.
        // Field result formatters apply new formatting to fields as they are updated,
        // which happens as soon as we create them using this InsertField method overload.
        // 1 -  Numeric:
        builder->InsertField(u" = 2 + 3 \\# $###");

        ASSERT_EQ(u"$5", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
        ASSERT_EQ(1, formatter->CountFormatInvocations(ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::Numeric));

        // 2 -  Date/time:
        builder->InsertField(u"DATE \\@ \"d MMMM yyyy\"");

        ASSERT_TRUE(doc->get_Range()->get_Fields()->idx_get(1)->get_Result().StartsWith(u"Date: "));
        ASSERT_EQ(1, formatter->CountFormatInvocations(ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::DateTime));

        // 3 -  General:
        builder->InsertField(u"QUOTE \"2\" \\* Ordinal");

        ASSERT_EQ(u"Item # 2:", doc->get_Range()->get_Fields()->idx_get(2)->get_Result());
        ASSERT_EQ(1, formatter->CountFormatInvocations(ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::General));

        formatter->PrintFormatInvocations();
    }

    /// <summary>
    /// When fields with formatting are updated, this formatter will override their formatting
    /// with a custom format, while tracking every invocation.
    /// </summary>
    class FieldResultFormatter : public IFieldResultFormatter
    {
    public:
        enum class FormatInvocationType
        {
            Numeric,
            DateTime,
            General,
            All
        };

    private:
        class FormatInvocation : public System::Object
        {
        public:
            ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType get_FormatInvocationType()
            {
                return pr_FormatInvocationType;
            }

            SharedPtr<System::Object> get_Value()
            {
                return pr_Value;
            }

            String get_OriginalFormat()
            {
                return pr_OriginalFormat;
            }

            String get_NewValue()
            {
                return pr_NewValue;
            }

            FormatInvocation(ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType formatInvocationType, SharedPtr<System::Object> value,
                             String originalFormat, String newValue)
                : pr_FormatInvocationType(((ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType)0))
            {
                pr_Value = value;
                pr_FormatInvocationType = formatInvocationType;
                pr_OriginalFormat = originalFormat;
                pr_NewValue = newValue;
            }

        private:
            ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType pr_FormatInvocationType;
            SharedPtr<System::Object> pr_Value;
            String pr_OriginalFormat;
            String pr_NewValue;
        };

    public:
        FieldResultFormatter(String numberFormat, String dateFormat, String generalFormat)
            : mFormatInvocations(MakeObject<System::Collections::Generic::List<SharedPtr<ExDocumentBuilder::FieldResultFormatter::FormatInvocation>>>())
        {
            mNumberFormat = numberFormat;
            mDateFormat = dateFormat;
            mGeneralFormat = generalFormat;
        }

        String FormatNumeric(double value, String format) override
        {
            if (String::IsNullOrEmpty(mNumberFormat))
            {
                return nullptr;
            }

            String newValue = String::Format(mNumberFormat, value);
            get_FormatInvocations()->Add(MakeObject<ExDocumentBuilder::FieldResultFormatter::FormatInvocation>(
                ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::Numeric, System::ObjectExt::Box<double>(value), format, newValue));
            return newValue;
        }

        String FormatDateTime(System::DateTime value, String format, CalendarType calendarType) override
        {
            if (String::IsNullOrEmpty(mDateFormat))
            {
                return nullptr;
            }

            String newValue = String::Format(mDateFormat, value);
            get_FormatInvocations()->Add(MakeObject<ExDocumentBuilder::FieldResultFormatter::FormatInvocation>(
                ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::DateTime,
                System::ObjectExt::Box<String>(String::Format(u"{0} ({1})", value, calendarType)), format, newValue));
            return newValue;
        }

        String Format(String value, GeneralFormat format) override
        {
            return Format(System::ObjectExt::Box<String>(value), format);
        }

        String Format(double value, GeneralFormat format) override
        {
            return Format(System::ObjectExt::Box<double>(value), format);
        }

        int CountFormatInvocations(ExDocumentBuilder::FieldResultFormatter::FormatInvocationType formatInvocationType)
        {
            if (formatInvocationType == ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::All)
            {
                return get_FormatInvocations()->get_Count();
            }

            std::function<bool(SharedPtr<FormatInvocation> f)> hasValidFormatInvocationType = [&formatInvocationType](SharedPtr<FormatInvocation> f)
            {
                return f->get_FormatInvocationType() == formatInvocationType;
            };

            return get_FormatInvocations()->LINQ_Count(hasValidFormatInvocationType);
        }

        void PrintFormatInvocations()
        {
            for (const auto& f : get_FormatInvocations())
            {
                std::cout << (String::Format(u"Invocation type:\t{0}\n", f->get_FormatInvocationType()) +
                              String::Format(u"\tOriginal value:\t\t{0}\n", f->get_Value()) +
                              String::Format(u"\tOriginal format:\t{0}\n", f->get_OriginalFormat()) +
                              String::Format(u"\tNew value:\t\t\t{0}\n", f->get_NewValue()))
                          << std::endl;
            }
        }

    private:
        String mNumberFormat;
        String mDateFormat;
        String mGeneralFormat;
        SharedPtr<System::Collections::Generic::List<SharedPtr<ExDocumentBuilder::FieldResultFormatter::FormatInvocation>>> mFormatInvocations;

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExDocumentBuilder::FieldResultFormatter::FormatInvocation>>> get_FormatInvocations()
        {
            return mFormatInvocations;
        }

        String Format(SharedPtr<System::Object> value, GeneralFormat format)
        {
            if (String::IsNullOrEmpty(mGeneralFormat))
            {
                return nullptr;
            }

            String newValue = String::Format(mGeneralFormat, value);
            get_FormatInvocations()->Add(MakeObject<ExDocumentBuilder::FieldResultFormatter::FormatInvocation>(
                ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::General, value, System::ObjectExt::ToString(format), newValue));
            return newValue;
        }
    };
    //ExEnd

    void InsertVideoWithUrl()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, Double, Double)
        //ExSummary:Shows how to insert an online video into a document using a URL.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertOnlineVideo(u"https://youtu.be/t_1LYZ102RA", 360, 270);

        // We can watch the video from Microsoft Word by clicking on the shape.
        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertVideoWithUrl.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertVideoWithUrl.docx");
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(480, 360, ImageType::Jpeg, shape);
        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, shape->get_HRef());

        ASPOSE_ASSERT_EQ(360.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(270.0, shape->get_Height());
    }

    void InsertUnderline()
    {
        //ExStart
        //ExFor:DocumentBuilder.Underline
        //ExSummary:Shows how to format text inserted by a document builder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->set_Underline(Underline::Dash);
        builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        builder->get_Font()->set_Size(32);

        // The builder applies formatting to its current paragraph and any new text added by it afterward.
        builder->Writeln(u"Large, blue, and underlined text.");

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertUnderline.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertUnderline.docx");
        SharedPtr<Run> firstRun = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Large, blue, and underlined text.", firstRun->GetText().Trim());
        ASSERT_EQ(Underline::Dash, firstRun->get_Font()->get_Underline());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), firstRun->get_Font()->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(32.0, firstRun->get_Font()->get_Size());
    }

    void CurrentStory()
    {
        //ExStart
        //ExFor:DocumentBuilder.CurrentStory
        //ExSummary:Shows how to work with a document builder's current story.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // A Story is a type of node that has child Paragraph nodes, such as a Body.
        ASPOSE_ASSERT_EQ(builder->get_CurrentStory(), doc->get_FirstSection()->get_Body());
        ASPOSE_ASSERT_EQ(builder->get_CurrentStory(), builder->get_CurrentParagraph()->get_ParentNode());
        ASSERT_EQ(StoryType::MainText, builder->get_CurrentStory()->get_StoryType());

        builder->get_CurrentStory()->AppendParagraph(u"Text added to current Story.");

        // A Story can also contain tables.
        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, cell 1");
        builder->InsertCell();
        builder->Write(u"Row 1, cell 2");
        builder->EndTable();

        ASSERT_TRUE(builder->get_CurrentStory()->get_Tables()->Contains(table));
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Tables()->get_Count());
        ASSERT_EQ(u"Row 1, cell 1\aRow 1, cell 2\a\a\rText added to current Story.", doc->get_FirstSection()->get_Body()->GetText().Trim());
    }

    void InsertStyleSeparator()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertStyleSeparator
        //ExSummary:Shows how to work with style separators.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Each paragraph can only have one style.
        // The InsertStyleSeparator method allows us to work around this limitation.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Write(u"This text is in a Heading style. ");
        builder->InsertStyleSeparator();

        SharedPtr<Style> paraStyle = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"MyParaStyle");
        paraStyle->get_Font()->set_Bold(false);
        paraStyle->get_Font()->set_Size(8);
        paraStyle->get_Font()->set_Name(u"Arial");

        builder->get_ParagraphFormat()->set_StyleName(paraStyle->get_Name());
        builder->Write(u"This text is in a custom style. ");

        // Calling the InsertStyleSeparator method creates another paragraph,
        // which can have a different style to the previous. There will be no break between paragraphs.
        // The text in the output document will look like one paragraph with two styles.
        ASSERT_EQ(2, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
        ASSERT_EQ(u"Heading 1", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Style()->get_Name());
        ASSERT_EQ(u"MyParaStyle", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_Style()->get_Name());

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertStyleSeparator.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertStyleSeparator.docx");

        ASSERT_EQ(2, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
        ASSERT_EQ(u"This text is in a Heading style. \r This text is in a custom style.", doc->GetText().Trim());
        ASSERT_EQ(u"Heading 1", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Style()->get_Name());
        ASSERT_EQ(u"MyParaStyle", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_Style()->get_Name());
        ASSERT_EQ(u" ", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->GetText());
    }

    void InsertDocument()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode)
        //ExFor:ImportFormatMode
        //ExSummary:Shows how to insert a document into another document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::PageBreak);

        auto docToInsert = MakeObject<Document>(MyDir + u"Formatted elements.docx");

        builder->InsertDocument(docToInsert, ImportFormatMode::KeepSourceFormatting);
        builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.InsertDocument.docx");
        //ExEnd

        ASSERT_EQ(29, doc->get_Styles()->get_Count());
        ASSERT_TRUE(DocumentHelper::CompareDocs(ArtifactsDir + u"DocumentBuilder.InsertDocument.docx", GoldsDir + u"DocumentBuilder.InsertDocument Gold.docx"));
    }

    void SmartStyleBehavior()
    {
        //ExStart
        //ExFor:ImportFormatOptions
        //ExFor:ImportFormatOptions.SmartStyleBehavior
        //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve duplicate styles while inserting documents.
        auto dstDoc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(dstDoc);

        SharedPtr<Style> myStyle = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"MyStyle");
        myStyle->get_Font()->set_Size(14);
        myStyle->get_Font()->set_Name(u"Courier New");
        myStyle->get_Font()->set_Color(System::Drawing::Color::get_Blue());

        builder->get_ParagraphFormat()->set_StyleName(myStyle->get_Name());
        builder->Writeln(u"Hello world!");

        // Clone the document and edit the clone's "MyStyle" style, so it is a different color than that of the original.
        // If we insert the clone into the original document, the two styles with the same name will cause a clash.
        SharedPtr<Document> srcDoc = dstDoc->Clone();
        srcDoc->get_Styles()->idx_get(u"MyStyle")->get_Font()->set_Color(System::Drawing::Color::get_Red());

        // When we enable SmartStyleBehavior and use the KeepSourceFormatting import format mode,
        // Aspose.Words will resolve style clashes by converting source document styles.
        // with the same names as destination styles into direct paragraph attributes.
        auto options = MakeObject<ImportFormatOptions>();
        options->set_SmartStyleBehavior(true);

        builder->InsertDocument(srcDoc, ImportFormatMode::KeepSourceFormatting, options);

        dstDoc->Save(ArtifactsDir + u"DocumentBuilder.SmartStyleBehavior.docx");
        //ExEnd

        dstDoc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SmartStyleBehavior.docx");

        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), dstDoc->get_Styles()->idx_get(u"MyStyle")->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(u"MyStyle", dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Style()->get_Name());

        ASSERT_EQ(u"Normal", dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_Style()->get_Name());
        ASPOSE_ASSERT_EQ(14, dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Size());
        ASSERT_EQ(u"Courier New", dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Name());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(),
                  dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Color().ToArgb());
    }

    void EmphasesWarningSourceMarkdown()
    {
        auto doc = MakeObject<Document>(MyDir + u"Emphases markdown warning.docx");

        auto warnings = MakeObject<WarningInfoCollection>();
        doc->set_WarningCallback(warnings);
        doc->Save(ArtifactsDir + u"DocumentBuilder.EmphasesWarningSourceMarkdown.md");

        for (const auto& warningInfo : warnings)
        {
            if (warningInfo->get_Source() == WarningSource::Markdown)
            {
                ASSERT_EQ(u"The (*, 0:11) cannot be properly written into Markdown.", warningInfo->get_Description());
            }
        }
    }

    void DoNotIgnoreHeaderFooter()
    {
        //ExStart
        //ExFor:ImportFormatOptions.IgnoreHeaderFooter
        //ExSummary:Shows how to specifies ignoring or not source formatting of headers/footers content.
        auto dstDoc = MakeObject<Document>(MyDir + u"Document.docx");
        auto srcDoc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_IgnoreHeaderFooter(false);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting, importFormatOptions);

        dstDoc->Save(ArtifactsDir + u"DocumentBuilder.DoNotIgnoreHeaderFooter.docx");
        //ExEnd
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentEmphases()
    {
        auto builder = MakeObject<DocumentBuilder>();

        // Bold and Italic are represented as Font.Bold and Font.Italic.
        builder->get_Font()->set_Italic(true);
        builder->Writeln(u"This text will be italic");

        // Use clear formatting if we don't want to combine styles between paragraphs.
        builder->get_Font()->ClearFormatting();

        builder->get_Font()->set_Bold(true);
        builder->Writeln(u"This text will be bold");

        builder->get_Font()->ClearFormatting();

        builder->get_Font()->set_Italic(true);
        builder->Write(u"You ");
        builder->get_Font()->set_Bold(true);
        builder->Write(u"can");
        builder->get_Font()->set_Bold(false);
        builder->Writeln(u" combine them");

        builder->get_Font()->ClearFormatting();

        builder->get_Font()->set_StrikeThrough(true);
        builder->Writeln(u"This text will be strikethrough");

        // Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis.
        builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentInlineCode()
    {
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder->MoveToDocumentEnd();
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(u"\n");

        // Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`).
        // If number of backticks is missed, then one backtick will be used by default.
        SharedPtr<Style> inlineCode1BackTicks = doc->get_Styles()->Add(StyleType::Character, u"InlineCode");
        builder->get_Font()->set_Style(inlineCode1BackTicks);
        builder->Writeln(u"Text with InlineCode style with one backtick");

        // Use optional dot (.) and number of backticks (`).
        // There will be 3 backticks.
        SharedPtr<Style> inlineCode3BackTicks = doc->get_Styles()->Add(StyleType::Character, u"InlineCode.3");
        builder->get_Font()->set_Style(inlineCode3BackTicks);
        builder->Writeln(u"Text with InlineCode style with 3 backticks");

        builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentHeadings()
    {
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder->MoveToDocumentEnd();
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(u"\n");

        // By default, Heading styles in Word may have bold and italic formatting.
        // If we do not want text to be emphasized, set these properties explicitly to false.
        // Thus we can't use 'builder.Font.ClearFormatting()' because Bold/Italic will be set to true.
        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        // Create for one heading for each level.
        builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->get_Font()->set_Italic(true);
        builder->Writeln(u"This is an italic H1 tag");

        // Reset our styles from the previous paragraph to not combine styles between paragraphs.
        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        // Structure-enhanced text heading can be added through style inheritance.
        SharedPtr<Style> setextHeading1 = doc->get_Styles()->Add(StyleType::Paragraph, u"SetextHeading1");
        builder->get_ParagraphFormat()->set_Style(setextHeading1);
        doc->get_Styles()->idx_get(u"SetextHeading1")->set_BaseStyleName(u"Heading 1");
        builder->Writeln(u"SetextHeading 1");

        builder->get_ParagraphFormat()->set_StyleName(u"Heading 2");
        builder->Writeln(u"This is an H2 tag");

        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        SharedPtr<Style> setextHeading2 = doc->get_Styles()->Add(StyleType::Paragraph, u"SetextHeading2");
        builder->get_ParagraphFormat()->set_Style(setextHeading2);
        doc->get_Styles()->idx_get(u"SetextHeading2")->set_BaseStyleName(u"Heading 2");
        builder->Writeln(u"SetextHeading 2");

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 3"));
        builder->Writeln(u"This is an H3 tag");

        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 4"));
        builder->get_Font()->set_Bold(true);
        builder->Writeln(u"This is an bold H4 tag");

        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 5"));
        builder->get_Font()->set_Italic(true);
        builder->get_Font()->set_Bold(true);
        builder->Writeln(u"This is an italic and bold H5 tag");

        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 6"));
        builder->Writeln(u"This is an H6 tag");

        doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentBlockquotes()
    {
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder->MoveToDocumentEnd();
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(u"\n");

        // By default, the document stores blockquote style for the first level.
        builder->get_ParagraphFormat()->set_StyleName(u"Quote");
        builder->Writeln(u"Blockquote");

        // Create styles for nested levels through style inheritance.
        SharedPtr<Style> quoteLevel2 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote1");
        builder->get_ParagraphFormat()->set_Style(quoteLevel2);
        doc->get_Styles()->idx_get(u"Quote1")->set_BaseStyleName(u"Quote");
        builder->Writeln(u"1. Nested blockquote");

        SharedPtr<Style> quoteLevel3 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote2");
        builder->get_ParagraphFormat()->set_Style(quoteLevel3);
        doc->get_Styles()->idx_get(u"Quote2")->set_BaseStyleName(u"Quote1");
        builder->get_Font()->set_Italic(true);
        builder->Writeln(u"2. Nested italic blockquote");

        SharedPtr<Style> quoteLevel4 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote3");
        builder->get_ParagraphFormat()->set_Style(quoteLevel4);
        doc->get_Styles()->idx_get(u"Quote3")->set_BaseStyleName(u"Quote2");
        builder->get_Font()->set_Italic(false);
        builder->get_Font()->set_Bold(true);
        builder->Writeln(u"3. Nested bold blockquote");

        SharedPtr<Style> quoteLevel5 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote4");
        builder->get_ParagraphFormat()->set_Style(quoteLevel5);
        doc->get_Styles()->idx_get(u"Quote4")->set_BaseStyleName(u"Quote3");
        builder->get_Font()->set_Bold(false);
        builder->Writeln(u"4. Nested blockquote");

        SharedPtr<Style> quoteLevel6 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote5");
        builder->get_ParagraphFormat()->set_Style(quoteLevel6);
        doc->get_Styles()->idx_get(u"Quote5")->set_BaseStyleName(u"Quote4");
        builder->Writeln(u"5. Nested blockquote");

        SharedPtr<Style> quoteLevel7 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote6");
        builder->get_ParagraphFormat()->set_Style(quoteLevel7);
        doc->get_Styles()->idx_get(u"Quote6")->set_BaseStyleName(u"Quote5");
        builder->get_Font()->set_Italic(true);
        builder->get_Font()->set_Bold(true);
        builder->Writeln(u"6. Nested italic bold blockquote");

        doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentIndentedCode()
    {
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder->MoveToDocumentEnd();
        builder->Writeln(u"\n");
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(u"\n");

        SharedPtr<Style> indentedCode = doc->get_Styles()->Add(StyleType::Paragraph, u"IndentedCode");
        builder->get_ParagraphFormat()->set_Style(indentedCode);
        builder->Writeln(u"This is an indented code");

        doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentFencedCode()
    {
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder->MoveToDocumentEnd();
        builder->Writeln(u"\n");
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(u"\n");

        SharedPtr<Style> fencedCode = doc->get_Styles()->Add(StyleType::Paragraph, u"FencedCode");
        builder->get_ParagraphFormat()->set_Style(fencedCode);
        builder->Writeln(u"This is a fenced code");

        SharedPtr<Style> fencedCodeWithInfo = doc->get_Styles()->Add(StyleType::Paragraph, u"FencedCode.C#");
        builder->get_ParagraphFormat()->set_Style(fencedCodeWithInfo);
        builder->Writeln(u"This is a fenced code with info string");

        doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentHorizontalRule()
    {
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder->MoveToDocumentEnd();
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(u"\n");

        // Insert HorizontalRule that will be present in .md file as '-----'.
        builder->InsertHorizontalRule();

        builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void MarkdownDocumentBulletedList()
    {
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Prepare our created document for further work
        // and clear paragraph formatting not to use the previous styles.
        builder->MoveToDocumentEnd();
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(u"\n");

        // Bulleted lists are represented using paragraph numbering.
        builder->get_ListFormat()->ApplyBulletDefault();
        // There can be 3 types of bulleted lists.
        // The only diff in a numbering format of the very first level are ‘-’, ‘+’ or ‘*’ respectively.
        builder->get_ListFormat()->get_List()->get_ListLevels()->idx_get(0)->set_NumberFormat(u"-");

        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2a");
        builder->Writeln(u"Item 2b");

        builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    }

    /// <summary>
    /// All markdown tests work with the same file. That's why we need order for them.
    /// </summary>
    void LoadMarkdownDocumentAndAssertContent(String text, String styleName, bool isItalic, bool isBold)
    {
        // Load created document from previous tests.
        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        for (const auto& paragraph : System::IterateOver<Paragraph>(paragraphs))
        {
            if (paragraph->get_Runs()->get_Count() != 0)
            {
                // Check that all document text has the necessary styles.
                if (paragraph->get_Runs()->idx_get(0)->get_Text() == text && !text.Contains(u"InlineCode"))
                {
                    ASSERT_EQ(styleName, paragraph->get_ParagraphFormat()->get_Style()->get_Name());
                    ASPOSE_ASSERT_EQ(isItalic, paragraph->get_Runs()->idx_get(0)->get_Font()->get_Italic());
                    ASPOSE_ASSERT_EQ(isBold, paragraph->get_Runs()->idx_get(0)->get_Font()->get_Bold());
                }
                else if (paragraph->get_Runs()->idx_get(0)->get_Text() == text && text.Contains(u"InlineCode"))
                {
                    ASSERT_EQ(styleName, paragraph->get_Runs()->idx_get(0)->get_Font()->get_StyleName());
                }
            }

            // Check that document also has a HorizontalRule present as a shape.
            SharedPtr<NodeCollection> shapesCollection = doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Shape, true);
            auto horizontalRuleShape = System::DynamicCast<Shape>(shapesCollection->idx_get(0));

            ASSERT_TRUE(shapesCollection->get_Count() == 1);
            ASSERT_TRUE(horizontalRuleShape->get_IsHorizontalRule());
        }
    }

    void MarkdownDocumentTableContentAlignment(TableContentAlignment tableContentAlignment)
    {
        auto builder = MakeObject<DocumentBuilder>();

        builder->InsertCell();
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
        builder->Write(u"Cell1");
        builder->InsertCell();
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->Write(u"Cell2");

        auto saveOptions = MakeObject<MarkdownSaveOptions>();
        saveOptions->set_TableContentAlignment(tableContentAlignment);

        builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocumentTableContentAlignment.md", saveOptions);

        auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocumentTableContentAlignment.md");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        switch (tableContentAlignment)
        {
        case TableContentAlignment::Auto:
            ASSERT_EQ(ParagraphAlignment::Right, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(ParagraphAlignment::Center, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;

        case TableContentAlignment::Left:
            ASSERT_EQ(ParagraphAlignment::Left, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(ParagraphAlignment::Left, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;

        case TableContentAlignment::Center:
            ASSERT_EQ(ParagraphAlignment::Center, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(ParagraphAlignment::Center, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;

        case TableContentAlignment::Right:
            ASSERT_EQ(ParagraphAlignment::Right, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(ParagraphAlignment::Right, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;
        }
    }

    //ExStart
    //ExFor:MarkdownSaveOptions.ImageSavingCallback
    //ExFor:IImageSavingCallback
    //ExSummary:Shows how to rename the image name during saving into Markdown document.
    void RenameImages()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<MarkdownSaveOptions>();

        // If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
        // Each image will be in the form of a file in the local file system.
        // There is also a callback that can customize the name and file system location of each image.
        options->set_ImageSavingCallback(MakeObject<ExDocumentBuilder::SavedImageRename>(u"DocumentBuilder.HandleDocument.md"));

        // The ImageSaving() method of our callback will be run at this time.
        doc->Save(ArtifactsDir + u"DocumentBuilder.HandleDocument.md", options);

        ASSERT_EQ(1, System::IO::Directory::GetFiles(ArtifactsDir)
                         ->LINQ_Where([](String s) { return s.StartsWith(ArtifactsDir + u"DocumentBuilder.HandleDocument.md shape"); })
                         ->LINQ_Count([](String f) { return f.EndsWith(u".jpeg"); }));
        ASSERT_EQ(8, System::IO::Directory::GetFiles(ArtifactsDir)
                         ->LINQ_Where([](String s) { return s.StartsWith(ArtifactsDir + u"DocumentBuilder.HandleDocument.md shape"); })
                         ->LINQ_Count([](String f) { return f.EndsWith(u".png"); }));
    }

    /// <summary>
    /// Renames saved images that are produced when an Markdown document is saved.
    /// </summary>
    class SavedImageRename : public IImageSavingCallback
    {
    public:
        SavedImageRename(String outFileName) : mCount(0)
        {
            mOutFileName = outFileName;
        }

    private:
        int mCount;
        String mOutFileName;

        void ImageSaving(SharedPtr<ImageSavingArgs> args) override
        {
            String imageFileName = String::Format(u"{0} shape {1}, of type {2}{3}", mOutFileName, ++mCount, args->get_CurrentShape()->get_ShapeType(),
                                                  System::IO::Path::GetExtension(args->get_ImageFileName()));

            args->set_ImageFileName(imageFileName);
            args->set_ImageStream(MakeObject<System::IO::FileStream>(ArtifactsDir + imageFileName, System::IO::FileMode::Create));

            ASSERT_TRUE(args->get_ImageStream()->get_CanWrite());
            ASSERT_TRUE(args->get_IsImageAvailable());
            ASSERT_FALSE(args->get_KeepImageStreamOpen());
        }
    };
    //ExEnd

    void MarkdownDocumentTests()
    {
        MarkdownDocumentEmphases();
        MarkdownDocumentInlineCode();
        MarkdownDocumentHeadings();
        MarkdownDocumentBlockquotes();
        MarkdownDocumentIndentedCode();
        MarkdownDocumentFencedCode();
        MarkdownDocumentHorizontalRule();
        MarkdownDocumentBulletedList();
        LoadMarkdownDocumentAndAssertContent(u"Italic", u"Normal", true, false);
        LoadMarkdownDocumentAndAssertContent(u"Bold", u"Normal", false, true);
        LoadMarkdownDocumentAndAssertContent(u"ItalicBold", u"Normal", true, true);
        LoadMarkdownDocumentAndAssertContent(u"Text with InlineCode style with one backtick", u"InlineCode", false, false);
        LoadMarkdownDocumentAndAssertContent(u"Text with InlineCode style with 3 backticks", u"InlineCode.3", false, false);
        LoadMarkdownDocumentAndAssertContent(u"This is an italic H1 tag", u"Heading 1", true, false);
        LoadMarkdownDocumentAndAssertContent(u"SetextHeading 1", u"SetextHeading1", false, false);
        LoadMarkdownDocumentAndAssertContent(u"This is an H2 tag", u"Heading 2", false, false);
        LoadMarkdownDocumentAndAssertContent(u"SetextHeading 2", u"SetextHeading2", false, false);
        LoadMarkdownDocumentAndAssertContent(u"This is an H3 tag", u"Heading 3", false, false);
        LoadMarkdownDocumentAndAssertContent(u"This is an bold H4 tag", u"Heading 4", false, true);
        LoadMarkdownDocumentAndAssertContent(u"This is an italic and bold H5 tag", u"Heading 5", true, true);
        LoadMarkdownDocumentAndAssertContent(u"This is an H6 tag", u"Heading 6", false, false);
        LoadMarkdownDocumentAndAssertContent(u"Blockquote", u"Quote", false, false);
        LoadMarkdownDocumentAndAssertContent(u"1. Nested blockquote", u"Quote1", false, false);
        LoadMarkdownDocumentAndAssertContent(u"2. Nested italic blockquote", u"Quote2", true, false);
        LoadMarkdownDocumentAndAssertContent(u"3. Nested bold blockquote", u"Quote3", false, true);
        LoadMarkdownDocumentAndAssertContent(u"4. Nested blockquote", u"Quote4", false, false);
        LoadMarkdownDocumentAndAssertContent(u"5. Nested blockquote", u"Quote5", false, false);
        LoadMarkdownDocumentAndAssertContent(u"6. Nested italic bold blockquote", u"Quote6", true, true);
        LoadMarkdownDocumentAndAssertContent(u"This is an indented code", u"IndentedCode", false, false);
        LoadMarkdownDocumentAndAssertContent(u"This is a fenced code with info string", u"FencedCode.C#", false, false);
        LoadMarkdownDocumentAndAssertContent(u"Item 1", u"Normal", false, false);
        LoadMarkdownDocumentAndAssertContent(u"This is a fenced code", u"FencedCode", false, false);
    }

    void InsertOnlineVideo()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an online video into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        String videoUrl = u"https://vimeo.com/52477838";

        // Insert a shape that plays a video from the web when clicked in Microsoft Word.
        // This rectangular shape will contain an image based on the first frame of the linked video
        // and a "play button" visual prompt. The video has an aspect ratio of 16:9.
        // We will set the shape's size to that ratio, so the image does not appear stretched.
        builder->InsertOnlineVideo(videoUrl, RelativeHorizontalPosition::LeftMargin, 0, RelativeVerticalPosition::TopMargin, 0, 320, 180, WrapType::Square);

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertOnlineVideo.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertOnlineVideo.docx");
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(640, 360, ImageType::Jpeg, shape);

        ASPOSE_ASSERT_EQ(320.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(180.0, shape->get_Height());
        ASPOSE_ASSERT_EQ(0.0, shape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, shape->get_Top());
        ASSERT_EQ(WrapType::Square, shape->get_WrapType());
        ASSERT_EQ(RelativeVerticalPosition::TopMargin, shape->get_RelativeVerticalPosition());
        ASSERT_EQ(RelativeHorizontalPosition::LeftMargin, shape->get_RelativeHorizontalPosition());

        ASSERT_EQ(u"https://vimeo.com/52477838", shape->get_HRef());
        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, shape->get_HRef());
    }

    void InsertOnlineVideoCustomThumbnail()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], Double, Double)
        //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an online video into a document with a custom thumbnail.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        String videoUrl = u"https://vimeo.com/52477838";
        String videoEmbedCode = String(u"<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" ") +
                                u"title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

        ArrayPtr<uint8_t> thumbnailImageBytes = System::IO::File::ReadAllBytes(ImageDir + u"Logo.jpg");

        {
            auto stream = MakeObject<System::IO::MemoryStream>(thumbnailImageBytes);
            {
                SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);
                // Below are two ways of creating a shape with a custom thumbnail, which links to an online video
                // that will play when we click on the shape in Microsoft Word.
                // 1 -  Insert an inline shape at the builder's node insertion cursor:
                builder->InsertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, image->get_Width(), image->get_Height());

                builder->InsertBreak(BreakType::PageBreak);

                // 2 -  Insert a floating shape:
                double left = builder->get_PageSetup()->get_RightMargin() - image->get_Width();
                double top = builder->get_PageSetup()->get_BottomMargin() - image->get_Height();

                builder->InsertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, RelativeHorizontalPosition::RightMargin, left,
                                           RelativeVerticalPosition::BottomMargin, top, image->get_Width(), image->get_Height(), WrapType::Square);
            }
        }

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
        ASPOSE_ASSERT_EQ(400.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(400.0, shape->get_Height());
        ASPOSE_ASSERT_EQ(0.0, shape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, shape->get_Top());
        ASSERT_EQ(WrapType::Inline, shape->get_WrapType());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, shape->get_RelativeVerticalPosition());
        ASSERT_EQ(RelativeHorizontalPosition::Column, shape->get_RelativeHorizontalPosition());

        ASSERT_EQ(u"https://vimeo.com/52477838", shape->get_HRef());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
        ASPOSE_ASSERT_EQ(400.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(400.0, shape->get_Height());
        ASPOSE_ASSERT_EQ(-329.15, shape->get_Left());
        ASPOSE_ASSERT_EQ(-329.15, shape->get_Top());
        ASSERT_EQ(WrapType::Square, shape->get_WrapType());
        ASSERT_EQ(RelativeVerticalPosition::BottomMargin, shape->get_RelativeVerticalPosition());
        ASSERT_EQ(RelativeHorizontalPosition::RightMargin, shape->get_RelativeHorizontalPosition());

        ASSERT_EQ(u"https://vimeo.com/52477838", shape->get_HRef());

        System::Net::ServicePointManager::set_SecurityProtocol(System::Net::SecurityProtocolType::Tls12);
        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, shape->get_HRef());
    }

    void InsertOleObjectAsIcon()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertOleObjectAsIcon(String, String, Boolean, String, String)
        //ExFor:DocumentBuilder.InsertOleObjectAsIcon(Stream, String, String, String)
        //ExSummary:Shows how to insert an embedded or linked OLE object as icon into the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        // the icon according to 'progId' and uses the filename for the icon caption.
        builder->InsertOleObjectAsIcon(MyDir + u"Presentation.pptx", u"Package", false, ImageDir + u"Logo icon.ico", u"My embedded file");

        builder->InsertBreak(BreakType::LineBreak);

        {
            auto stream = MakeObject<System::IO::FileStream>(MyDir + u"Presentation.pptx", System::IO::FileMode::Open);
            // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
            // the icon according to the file extension and uses the filename for the icon caption.
            SharedPtr<Shape> shape = builder->InsertOleObjectAsIcon(stream, u"PowerPoint.Application", ImageDir + u"Logo icon.ico", u"My embedded file stream");

            SharedPtr<OlePackage> setOlePackage = shape->get_OleFormat()->get_OlePackage();
            setOlePackage->set_FileName(u"Presentation.pptx");
            setOlePackage->set_DisplayName(u"Presentation.pptx");
        }

        doc->Save(ArtifactsDir + u"DocumentBuilder.InsertOleObjectAsIcon.docx");
        //ExEnd
    }
};

} // namespace ApiExamples
