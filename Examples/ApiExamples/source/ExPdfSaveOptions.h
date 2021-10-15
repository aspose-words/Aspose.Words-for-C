#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/FileFormatInfo.h>
#include <Aspose.Words.Cpp/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Fonts/PhysicalFontInfo.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/ColorMode.h>
#include <Aspose.Words.Cpp/Saving/DmlEffectsRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/DmlRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/DownsampleOptions.h>
#include <Aspose.Words.Cpp/Saving/EmfPlusDualRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/HeaderFooterBookmarksExportMode.h>
#include <Aspose.Words.Cpp/Saving/MetafileRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Saving/NumeralFormat.h>
#include <Aspose.Words.Cpp/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Saving/PageSet.h>
#include <Aspose.Words.Cpp/Saving/PdfCompliance.h>
#include <Aspose.Words.Cpp/Saving/PdfCustomPropertiesExport.h>
#include <Aspose.Words.Cpp/Saving/PdfDigitalSignatureDetails.h>
#include <Aspose.Words.Cpp/Saving/PdfDigitalSignatureHashAlgorithm.h>
#include <Aspose.Words.Cpp/Saving/PdfDigitalSignatureTimestampSettings.h>
#include <Aspose.Words.Cpp/Saving/PdfEncryptionAlgorithm.h>
#include <Aspose.Words.Cpp/Saving/PdfEncryptionDetails.h>
#include <Aspose.Words.Cpp/Saving/PdfFontEmbeddingMode.h>
#include <Aspose.Words.Cpp/Saving/PdfImageColorSpaceExportMode.h>
#include <Aspose.Words.Cpp/Saving/PdfImageCompression.h>
#include <Aspose.Words.Cpp/Saving/PdfPageMode.h>
#include <Aspose.Words.Cpp/Saving/PdfPermissions.h>
#include <Aspose.Words.Cpp/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/PdfTextCompression.h>
#include <Aspose.Words.Cpp/Saving/PdfZoomBehavior.h>
#include <Aspose.Words.Cpp/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/WarningType.h>
#include <system/array.h>
#include <system/collections/ilist.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/timespan.h>
#include <gtest/gtest-spi.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::DigitalSignatures;

namespace ApiExamples {

class ExPdfSaveOptions : public ApiExampleBase
{
public:
    void OnePage()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageSet
        //ExFor:Document.Save(Stream, SaveOptions)
        //ExSummary:Shows how to convert only some of the pages in a document to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 3.");

        {
            SharedPtr<System::IO::Stream> stream = System::IO::File::Create(ArtifactsDir + u"PdfSaveOptions.OnePage.pdf");
            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            auto options = MakeObject<PdfSaveOptions>();

            // Set the "PageIndex" to "1" to render a portion of the document starting from the second page.
            options->set_PageSet(MakeObject<PageSet>(1));

            // This document will contain one page starting from page two, which will only contain the second page.
            doc->Save(stream, options);
        }
        //ExEnd
    }

    void HeadingsOutlineLevels()
    {
        //ExStart
        //ExFor:ParagraphFormat.IsHeading
        //ExFor:PdfSaveOptions.OutlineOptions
        //ExFor:PdfSaveOptions.SaveFormat
        //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved PDF document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsHeading());

        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 1.1");
        builder->Writeln(u"Heading 1.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading3);

        builder->Writeln(u"Heading 1.2.1");
        builder->Writeln(u"Heading 1.2.2");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Pdf);

        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
        // The last two headings we have inserted above will not appear.
        saveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(2);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.HeadingsOutlineLevels.pdf", saveOptions);
        //ExEnd
    }

    void CreateMissingOutlineLevels(bool createMissingOutlineLevels)
    {
        //ExStart
        //ExFor:OutlineOptions.CreateMissingOutlineLevels
        //ExFor:PdfSaveOptions.OutlineOptions
        //ExSummary:Shows how to work with outline levels that do not contain any corresponding headings when saving a PDF document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert headings that can serve as TOC entries of levels 1 and 5.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsHeading());

        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading5);

        builder->Writeln(u"Heading 1.1.1.1.1");
        builder->Writeln(u"Heading 1.1.1.1.2");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();

        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "5" to include all headings of levels 5 and below in the outline.
        saveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(5);

        // This document contains headings of levels 1 and 5, and no headings with levels of 2, 3, and 4.
        // The output PDF document will treat outline levels 2, 3, and 4 as "missing".
        // Set the "CreateMissingOutlineLevels" property to "true" to include all missing levels in the outline,
        // leaving blank outline entries since there are no usable headings.
        // Set the "CreateMissingOutlineLevels" property to "false" to ignore missing outline levels,
        // and treat the outline level 5 headings as level 2.
        saveOptions->get_OutlineOptions()->set_CreateMissingOutlineLevels(createMissingOutlineLevels);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.CreateMissingOutlineLevels.pdf", saveOptions);
        //ExEnd
    }

    void TableHeadingOutlines(bool createOutlinesForHeadingsInTables)
    {
        //ExStart
        //ExFor:OutlineOptions.CreateOutlinesForHeadingsInTables
        //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a table with three rows. The first row,
        // whose text we will format in a heading-type style, will serve as the column header.
        builder->StartTable();
        builder->InsertCell();
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Write(u"Customers");
        builder->EndRow();
        builder->InsertCell();
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Normal);
        builder->Write(u"John Doe");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Jane Doe");
        builder->EndTable();

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();

        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "1" to get the outline
        // to only register headings with heading levels that are no larger than 1.
        pdfSaveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(1);

        // Set the "CreateOutlinesForHeadingsInTables" property to "false" to exclude all headings within tables,
        // such as the one we have created above from the outline.
        // Set the "CreateOutlinesForHeadingsInTables" property to "true" to include all headings within tables
        // in the outline, provided that they have a heading level that is no larger than the value of the "HeadingsOutlineLevels" property.
        pdfSaveOptions->get_OutlineOptions()->set_CreateOutlinesForHeadingsInTables(createOutlinesForHeadingsInTables);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
        //ExEnd
    }

    void ExpandedOutlineLevels()
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:PdfSaveOptions
        //ExFor:OutlineOptions.HeadingsOutlineLevels
        //ExFor:OutlineOptions.ExpandedOutlineLevels
        //ExSummary:Shows how to convert a whole document to PDF with three levels in the document outline.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert headings of levels 1 to 5.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsHeading());

        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 1.1");
        builder->Writeln(u"Heading 1.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading3);

        builder->Writeln(u"Heading 1.2.1");
        builder->Writeln(u"Heading 1.2.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading4);

        builder->Writeln(u"Heading 1.2.2.1");
        builder->Writeln(u"Heading 1.2.2.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading5);

        builder->Writeln(u"Heading 1.2.2.2.1");
        builder->Writeln(u"Heading 1.2.2.2.2");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "4" to exclude all headings whose levels are above 4 from the outline.
        options->get_OutlineOptions()->set_HeadingsOutlineLevels(4);

        // If an outline entry has subsequent entries of a higher level inbetween itself and the next entry of the same or lower level,
        // an arrow will appear to the left of the entry. This entry is the "owner" of several such "sub-entries".
        // In our document, the outline entries from the 5th heading level are sub-entries of the second 4th level outline entry,
        // the 4th and 5th heading level entries are sub-entries of the second 3rd level entry, and so on.
        // In the outline, we can click on the arrow of the "owner" entry to collapse/expand all its sub-entries.
        // Set the "ExpandedOutlineLevels" property to "2" to automatically expand all heading level 2 and lower outline entries
        // and collapse all level and 3 and higher entries when we open the document.
        options->get_OutlineOptions()->set_ExpandedOutlineLevels(2);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ExpandedOutlineLevels.pdf", options);
        //ExEnd
    }

    void UpdateFields(bool updateFields)
    {
        //ExStart
        //ExFor:PdfSaveOptions.Clone
        //ExFor:SaveOptions.UpdateFields
        //ExSummary:Shows how to update all the fields in a document immediately before saving it to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert text with PAGE and NUMPAGES fields. These fields do not display the correct value in real time.
        // We will need to manually update them using updating methods such as "Field.Update()", and "Document.UpdateFields()"
        // each time we need them to display accurate values.
        builder->Write(u"Page ");
        builder->InsertField(u"PAGE", u"");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES", u"");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Hello World!");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "UpdateFields" property to "false" to not update all the fields in a document right before a save operation.
        // This is the preferable option if we know that all our fields will be up to date before saving.
        // Set the "UpdateFields" property to "true" to iterate through all the document
        // fields and update them before we save it as a PDF. This will make sure that all the fields will display
        // the most accurate values in the PDF.
        options->set_UpdateFields(updateFields);

        // We can clone PdfSaveOptions objects.
        ASPOSE_ASSERT_NS(options, options->Clone());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.UpdateFields.pdf", options);
        //ExEnd
    }

    void PreserveFormFields(bool preserveFormFields)
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreserveFormFields
        //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Please select a fruit: ");

        // Insert a combo box which will allow a user to choose an option from a collection of strings.
        builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"Apple", u"Banana", u"Cherry"}), 0);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto pdfOptions = MakeObject<PdfSaveOptions>();

        // Set the "PreserveFormFields" property to "true" to save form fields as interactive objects in the output PDF.
        // Set the "PreserveFormFields" property to "false" to freeze all form fields in the document at
        // their current values and display them as plain text in the output PDF.
        pdfOptions->set_PreserveFormFields(preserveFormFields);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.PreserveFormFields.pdf", pdfOptions);
        //ExEnd
    }

    void Compliance(PdfCompliance pdfCompliance)
    {
        //ExStart
        //ExFor:PdfSaveOptions.Compliance
        //ExFor:PdfCompliance
        //ExSummary:Shows how to set the PDF standards compliance level of saved PDF documents.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();

        // Set the "Compliance" property to "PdfCompliance.PdfA1b" to comply with the "PDF/A-1b" standard,
        // which aims to preserve the visual appearance of the document as Aspose.Words convert it to PDF.
        // Set the "Compliance" property to "PdfCompliance.Pdf17" to comply with the "1.7" standard.
        // Set the "Compliance" property to "PdfCompliance.PdfA1a" to comply with the "PDF/A-1a" standard,
        // which complies with "PDF/A-1b" as well as preserving the document structure of the original document.
        // This helps with making documents searchable but may significantly increase the size of already large documents.
        saveOptions->set_Compliance(pdfCompliance);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.Compliance.pdf", saveOptions);
        //ExEnd
    }

    void TextCompression(PdfTextCompression pdfTextCompression)
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:PdfSaveOptions.TextCompression
        //ExFor:PdfTextCompression
        //ExSummary:Shows how to apply text compression when saving a document to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        for (int i = 0; i < 100; i++)
        {
            builder->Writeln(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                             u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        }

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "TextCompression" property to "PdfTextCompression.None" to not apply any
        // compression to text when we save the document to PDF.
        // Set the "TextCompression" property to "PdfTextCompression.Flate" to apply ZIP compression
        // to text when we save the document to PDF. The larger the document, the bigger the impact that this will have.
        options->set_TextCompression(pdfTextCompression);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.TextCompression.pdf", options);
        //ExEnd

        switch (pdfTextCompression)
        {
        case PdfTextCompression::None:
            ASSERT_LT(60000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.TextCompression.pdf")->get_Length());
            TestUtil::FileContainsString(u"12 0 obj\r\n<</Length 13 0 R>>stream", ArtifactsDir + u"PdfSaveOptions.TextCompression.pdf");
            break;

        case PdfTextCompression::Flate:
            ASSERT_GE(30000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.TextCompression.pdf")->get_Length());
            TestUtil::FileContainsString(u"12 0 obj\r\n<</Length 13 0 R/Filter /FlateDecode>>stream", ArtifactsDir + u"PdfSaveOptions.TextCompression.pdf");
            break;
        }
    }

    void ImageCompression(PdfImageCompression pdfImageCompression)
    {
        //ExStart
        //ExFor:PdfSaveOptions.ImageCompression
        //ExFor:PdfSaveOptions.JpegQuality
        //ExFor:PdfImageCompression
        //ExSummary:Shows how to specify a compression type for all images in a document that we are converting to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Jpeg image:");
        builder->InsertImage(ImageDir + u"Logo.jpg");
        builder->InsertParagraph();
        builder->Writeln(u"Png image:");
        builder->InsertImage(ImageDir + u"Transparent background logo.png");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();

        // Set the "ImageCompression" property to "PdfImageCompression.Auto" to use the
        // "ImageCompression" property to control the quality of the Jpeg images that end up in the output PDF.
        // Set the "ImageCompression" property to "PdfImageCompression.Jpeg" to use the
        // "ImageCompression" property to control the quality of all images that end up in the output PDF.
        pdfSaveOptions->set_ImageCompression(pdfImageCompression);

        // Set the "JpegQuality" property to "10" to strengthen compression at the cost of image quality.
        pdfSaveOptions->set_JpegQuality(10);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ImageCompression.pdf", pdfSaveOptions);
        //ExEnd
    }

    void ImageColorSpaceExportMode(PdfImageColorSpaceExportMode pdfImageColorSpaceExportMode)
    {
        //ExStart
        //ExFor:PdfImageColorSpaceExportMode
        //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
        //ExSummary:Shows how to set a different color space for images in a document as we export it to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Jpeg image:");
        builder->InsertImage(ImageDir + u"Logo.jpg");
        builder->InsertParagraph();
        builder->Writeln(u"Png image:");
        builder->InsertImage(ImageDir + u"Transparent background logo.png");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();

        // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.Auto" to get Aspose.Words to
        // automatically select the color space for images in the document that it converts to PDF.
        // In most cases, the color space will be RGB.
        // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.SimpleCmyk"
        // to use the CMYK color space for all images in the saved PDF.
        // Aspose.Words will also apply Flate compression to all images and ignore the "ImageCompression" property's value.
        pdfSaveOptions->set_ImageColorSpaceExportMode(pdfImageColorSpaceExportMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ImageColorSpaceExportMode.pdf", pdfSaveOptions);
        //ExEnd
    }

    void DownsampleOptions_()
    {
        //ExStart
        //ExFor:DownsampleOptions
        //ExFor:DownsampleOptions.DownsampleImages
        //ExFor:DownsampleOptions.Resolution
        //ExFor:DownsampleOptions.ResolutionThreshold
        //ExFor:PdfSaveOptions.DownsampleOptions
        //ExSummary:Shows how to change the resolution of images in the PDF document.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // By default, Aspose.Words downsample all images in a document that we save to PDF to 220 ppi.
        ASSERT_TRUE(options->get_DownsampleOptions()->get_DownsampleImages());
        ASSERT_EQ(220, options->get_DownsampleOptions()->get_Resolution());
        ASSERT_EQ(0, options->get_DownsampleOptions()->get_ResolutionThreshold());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.DownsampleOptions.Default.pdf", options);

        // Set the "Resolution" property to "36" to downsample all images to 36 ppi.
        options->get_DownsampleOptions()->set_Resolution(36);

        // Set the "ResolutionThreshold" property to only apply the downsampling to
        // images with a resolution that is above 128 ppi.
        options->get_DownsampleOptions()->set_ResolutionThreshold(128);

        // Only the first two images from the document will be downsampled at this stage.
        doc->Save(ArtifactsDir + u"PdfSaveOptions.DownsampleOptions.LowerResolution.pdf", options);
        //ExEnd
    }

    void ColorRendering(ColorMode colorMode)
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:ColorMode
        //ExFor:FixedPageSaveOptions.ColorMode
        //ExSummary:Shows how to change image color with saving options property.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        // Set the "ColorMode" property to "Grayscale" to render all images from the document in black and white.
        // The size of the output document may be larger with this setting.
        // Set the "ColorMode" property to "Normal" to render all images in color.
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->set_ColorMode(colorMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
        //ExEnd
    }

    void DocTitle(bool displayDocTitle)
    {
        //ExStart
        //ExFor:PdfSaveOptions.DisplayDocTitle
        //ExSummary:Shows how to display the title of the document as the title bar.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        doc->get_BuiltInDocumentProperties()->set_Title(u"Windows bar pdf title");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        // Set the "DisplayDocTitle" to "true" to get some PDF readers, such as Adobe Acrobat Pro,
        // to display the value of the document's "Title" built-in property in the tab that belongs to this document.
        // Set the "DisplayDocTitle" to "false" to get such readers to display the document's filename.
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->set_DisplayDocTitle(displayDocTitle);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.DocTitle.pdf", pdfSaveOptions);
        //ExEnd
    }

    void MemoryOptimization(bool memoryOptimization)
    {
        //ExStart
        //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
        //ExFor:SaveOptions.MemoryOptimization
        //ExSummary:Shows an option to optimize memory consumption when rendering large documents to PDF.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        SharedPtr<SaveOptions> saveOptions = SaveOptions::CreateSaveOptions(SaveFormat::Pdf);

        // Set the "MemoryOptimization" property to "true" to lower the memory footprint of large documents' saving operations
        // at the cost of increasing the duration of the operation.
        // Set the "MemoryOptimization" property to "false" to save the document as a PDF normally.
        saveOptions->set_MemoryOptimization(memoryOptimization);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.MemoryOptimization.pdf", saveOptions);
        //ExEnd
    }

    void EscapeUri(String uri, String result)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertHyperlink(u"Testlink", uri, false);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.EscapedUri.pdf");
    }

    void OpenHyperlinksInNewWindow(bool openHyperlinksInNewWindow)
    {
        //ExStart
        //ExFor:PdfSaveOptions.OpenHyperlinksInNewWindow
        //ExSummary:Shows how to save hyperlinks in a document we convert to PDF so that they open new pages when we click on them.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertHyperlink(u"Testlink", u"https://www.google.com/search?q=%20aspose", false);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "OpenHyperlinksInNewWindow" property to "true" to save all hyperlinks using Javascript code
        // that forces readers to open these links in new windows/browser tabs.
        // Set the "OpenHyperlinksInNewWindow" property to "false" to save all hyperlinks normally.
        options->set_OpenHyperlinksInNewWindow(openHyperlinksInNewWindow);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.OpenHyperlinksInNewWindow.pdf", options);
        //ExEnd

        if (openHyperlinksInNewWindow)
        {
            TestUtil::FileContainsString(
                String(u"<</Type /Annot/Subtype /Link/Rect [70.84999847 707.35101318 110.17799377 721.15002441]/BS ") +
                    u"<</Type/Border/S/S/W 0>>/A<</Type /Action/S /JavaScript/JS(app.launchURL\\(\"https://www.google.com/search?q=%20aspose\", true\\);)>>>>",
                ArtifactsDir + u"PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");
        }
        else
        {
            TestUtil::FileContainsString(String(u"<</Type /Annot/Subtype /Link/Rect [70.84999847 707.35101318 110.17799377 721.15002441]/BS ") +
                                             u"<</Type/Border/S/S/W 0>>/A<</Type /Action/S /URI/URI(https://www.google.com/search?q=%20aspose)>>>>",
                                         ArtifactsDir + u"PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");
        }
    }

    //ExStart
    //ExFor:MetafileRenderingMode
    //ExFor:MetafileRenderingOptions
    //ExFor:MetafileRenderingOptions.EmulateRasterOperations
    //ExFor:MetafileRenderingOptions.RenderingMode
    //ExFor:IWarningCallback
    //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
    //ExSummary:Shows added a fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
    void HandleBinaryRasterWarnings()
    {
        auto doc = MakeObject<Document>(MyDir + u"WMF with image.docx");

        auto metafileRenderingOptions = MakeObject<MetafileRenderingOptions>();

        // Set the "EmulateRasterOperations" property to "false" to fall back to bitmap when
        // it encounters a metafile, which will require raster operations to render in the output PDF.
        metafileRenderingOptions->set_EmulateRasterOperations(false);

        // Set the "RenderingMode" property to "VectorWithFallback" to try to render every metafile using vector graphics.
        metafileRenderingOptions->set_RenderingMode(MetafileRenderingMode::VectorWithFallback);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF and applies the configuration
        // in our MetafileRenderingOptions object to the saving operation.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_MetafileRenderingOptions(metafileRenderingOptions);

        auto callback = MakeObject<ExPdfSaveOptions::HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

        ASSERT_EQ(1, callback->Warnings->get_Count());
        ASSERT_EQ(u"'R2_XORPEN' binary raster operation is partly supported.", callback->Warnings->idx_get(0)->get_Description());
    }

    /// <summary>
    /// Prints and collects formatting loss-related warnings that occur upon saving a document.
    /// </summary>
    class HandleDocumentWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> Warnings;

        void Warning(SharedPtr<WarningInfo> info) override
        {
            if (info->get_WarningType() == WarningType::MinorFormattingLoss)
            {
                std::cout << (String(u"Unsupported operation: ") + info->get_Description()) << std::endl;
                Warnings->Warning(info);
            }
        }

        HandleDocumentWarnings() : Warnings(MakeObject<WarningInfoCollection>())
        {
        }
    };
    //ExEnd

    void HeaderFooterBookmarksExportMode(Aspose::Words::Saving::HeaderFooterBookmarksExportMode headerFooterBookmarksExportMode)
    {
        //ExStart
        //ExFor:HeaderFooterBookmarksExportMode
        //ExFor:OutlineOptions
        //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
        //ExFor:PdfSaveOptions.HeaderFooterBookmarksExportMode
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfPageMode
        //ExSummary:Shows to process bookmarks in headers/footers in a document that we are rendering to PDF.
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks in headers and footers.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();

        // Set the "PageMode" property to "PdfPageMode.UseOutlines" to display the outline navigation pane in the output PDF.
        saveOptions->set_PageMode(PdfPageMode::UseOutlines);

        // Set the "DefaultBookmarksOutlineLevel" property to "1" to display all
        // bookmarks at the first level of the outline in the output PDF.
        saveOptions->get_OutlineOptions()->set_DefaultBookmarksOutlineLevel(1);

        // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.None" to
        // not export any bookmarks that are inside headers/footers.
        // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.First" to
        // only export bookmarks in the first section's header/footers.
        // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.All" to
        // export bookmarks that are in all headers/footers.
        saveOptions->set_HeaderFooterBookmarksExportMode(headerFooterBookmarksExportMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
        //ExEnd
    }

    void UnsupportedImageFormatWarning()
    {
        auto doc = MakeObject<Document>(MyDir + u"Corrupted image.docx");

        auto saveWarningCallback = MakeObject<ExPdfSaveOptions::SaveWarningCallback>();
        doc->set_WarningCallback(saveWarningCallback);

        doc->Save(ArtifactsDir + u"PdfSaveOption.UnsupportedImageFormatWarning.pdf", SaveFormat::Pdf);

        ASSERT_EQ(saveWarningCallback->SaveWarnings->idx_get(0)->get_Description(), u"Image can not be processed. Possibly unsupported image format.");
    }

    class SaveWarningCallback : public IWarningCallback
    {
        friend class ExPdfSaveOptions;

    public:
        void Warning(SharedPtr<WarningInfo> info) override
        {
            if (info->get_WarningType() == WarningType::MinorFormattingLoss)
            {
                std::cout << String::Format(u"{0}: {1}.", info->get_WarningType(), info->get_Description()) << std::endl;
                SaveWarnings->Warning(info);
            }
        }

        SaveWarningCallback() : SaveWarnings(MakeObject<WarningInfoCollection>())
        {
        }

    protected:
        SharedPtr<WarningInfoCollection> SaveWarnings;
    };

    void FontsScaledToMetafileSize(bool scaleWmfFonts)
    {
        //ExStart
        //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
        //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
        auto doc = MakeObject<Document>(MyDir + u"WMF with text.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();

        // Set the "ScaleWmfFontsToMetafileSize" property to "true" to scale fonts
        // that format text within WMF images according to the size of the metafile on the page.
        // Set the "ScaleWmfFontsToMetafileSize" property to "false" to
        // preserve the default scale of these fonts.
        saveOptions->get_MetafileRenderingOptions()->set_ScaleWmfFontsToMetafileSize(scaleWmfFonts);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
        //ExEnd
    }

    void EmbedFullFonts(bool embedFullFonts)
    {
        //ExStart
        //ExFor:PdfSaveOptions.#ctor
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExSummary:Shows how to enable or disable subsetting when embedding fonts while rendering a document to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Arvo");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        // Configure our font sources to ensure that we have access to both the fonts in this document.
        ArrayPtr<SharedPtr<FontSourceBase>> originalFontsSources = FontSettings::get_DefaultInstance()->GetFontsSources();
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, true);
        FontSettings::get_DefaultInstance()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({originalFontsSources[0], folderFontSource}));

        ArrayPtr<SharedPtr<FontSourceBase>> fontSources = FontSettings::get_DefaultInstance()->GetFontsSources();
        ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));
        ASSERT_TRUE(fontSources[1]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arvo"; }));

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Since our document contains a custom font, embedding in the output document may be desirable.
        // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
        // The document's size may become very large, but we will have full use of all fonts if we edit the PDF.
        // Set the "EmbedFullFonts" property to "false" to apply subsetting to fonts, saving only the glyphs
        // that the document is using. The file will be considerably smaller,
        // but we may need access to any custom fonts if we edit the document.
        options->set_EmbedFullFonts(embedFullFonts);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.EmbedFullFonts.pdf", options);

        if (embedFullFonts)
        {
            ASSERT_LT(500000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.EmbedFullFonts.pdf")->get_Length());
        }
        else
        {
            ASSERT_GE(25000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.EmbedFullFonts.pdf")->get_Length());
        }

        // Restore the original font sources.
        FontSettings::get_DefaultInstance()->SetFontsSources(originalFontsSources);
        //ExEnd
    }

    void EmbedWindowsFonts(PdfFontEmbeddingMode pdfFontEmbeddingMode)
    {
        //ExStart
        //ExFor:PdfSaveOptions.FontEmbeddingMode
        //ExFor:PdfFontEmbeddingMode
        //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // "Arial" is a standard font, and "Courier New" is a nonstandard font.
        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Courier New");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
        options->set_EmbedFullFonts(true);

        // Set the "FontEmbeddingMode" property to "EmbedAll" to embed all fonts in the output PDF.
        // Set the "FontEmbeddingMode" property to "EmbedNonstandard" to only allow nonstandard fonts' embedding in the output PDF.
        // Set the "FontEmbeddingMode" property to "EmbedNone" to not embed any fonts in the output PDF.
        options->set_FontEmbeddingMode(pdfFontEmbeddingMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.EmbedWindowsFonts.pdf", options);

        switch (pdfFontEmbeddingMode)
        {
        case PdfFontEmbeddingMode::EmbedAll:
            ASSERT_LT(1000000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.EmbedWindowsFonts.pdf")->get_Length());
            break;

        case PdfFontEmbeddingMode::EmbedNonstandard:
            ASSERT_LT(480000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.EmbedWindowsFonts.pdf")->get_Length());
            break;

        case PdfFontEmbeddingMode::EmbedNone:
            ASSERT_GE(4212, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.EmbedWindowsFonts.pdf")->get_Length());
            break;
        }
        //ExEnd
    }

    void EmbedCoreFonts(bool useCoreFonts)
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseCoreFonts
        //ExSummary:Shows how enable/disable PDF Type 1 font substitution.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Courier New");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "UseCoreFonts" property to "true" to replace some fonts,
        // including the two fonts in our document, with their PDF Type 1 equivalents.
        // Set the "UseCoreFonts" property to "false" to not apply PDF Type 1 fonts.
        options->set_UseCoreFonts(useCoreFonts);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.EmbedCoreFonts.pdf", options);

        if (useCoreFonts)
        {
            ASSERT_GE(3000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.EmbedCoreFonts.pdf")->get_Length());
        }
        else
        {
            ASSERT_LT(30000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"PdfSaveOptions.EmbedCoreFonts.pdf")->get_Length());
        }
        //ExEnd
    }

    void AdditionalTextPositioning(bool applyAdditionalTextPositioning)
    {
        //ExStart
        //ExFor:PdfSaveOptions.AdditionalTextPositioning
        //ExSummary:Show how to write additional text positioning operators.
        auto doc = MakeObject<Document>(MyDir + u"Text positioning operators.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_TextCompression(PdfTextCompression::None);
        saveOptions->set_AdditionalTextPositioning(applyAdditionalTextPositioning);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd
    }

    void SaveAsPdfBookFold(bool renderTextAsBookfold)
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
        // in the output PDF in a way that helps us use it to make a booklet.
        // Set the "UseBookFoldPrintingSettings" property to "false" to render the PDF normally.
        options->set_UseBookFoldPrintingSettings(renderTextAsBookfold);

        // If we are rendering the document as a booklet, we must set the "MultiplePages"
        // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
        if (renderTextAsBookfold)
        {
            for (const auto& s : System::IterateOver<Section>(doc->get_Sections()))
            {
                s->get_PageSetup()->set_MultiplePages(MultiplePagesType::BookFoldPrinting);
            }
        }

        // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
        // and the contents will line up in a way that creates a booklet.
        doc->Save(ArtifactsDir + u"PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
        //ExEnd
    }

    void ZoomBehaviour()
    {
        //ExStart
        //ExFor:PdfSaveOptions.ZoomBehavior
        //ExFor:PdfSaveOptions.ZoomFactor
        //ExFor:PdfZoomBehavior
        //ExSummary:Shows how to set the default zooming that a reader applies when opening a rendered PDF document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        // Set the "ZoomBehavior" property to "PdfZoomBehavior.ZoomFactor" to get a PDF reader to
        // apply a percentage-based zoom factor when we open the document with it.
        // Set the "ZoomFactor" property to "25" to give the zoom factor a value of 25%.
        auto options = MakeObject<PdfSaveOptions>();
        options->set_ZoomBehavior(PdfZoomBehavior::ZoomFactor);
        options->set_ZoomFactor(25);

        // When we open this document using a reader such as Adobe Acrobat, we will see the document scaled at 1/4 of its actual size.
        doc->Save(ArtifactsDir + u"PdfSaveOptions.ZoomBehaviour.pdf", options);
        //ExEnd
    }

    void PageMode(PdfPageMode pageMode)
    {
        //ExStart
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfPageMode
        //ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "PageMode" property to "PdfPageMode.FullScreen" to get the PDF reader to open the saved
        // document in full-screen mode, which takes over the monitor's display and has no controls visible.
        // Set the "PageMode" property to "PdfPageMode.UseThumbs" to get the PDF reader to display a separate panel
        // with a thumbnail for each page in the document.
        // Set the "PageMode" property to "PdfPageMode.UseOC" to get the PDF reader to display a separate panel
        // that allows us to work with any layers present in the document.
        // Set the "PageMode" property to "PdfPageMode.UseOutlines" to get the PDF reader
        // also to display the outline, if possible.
        // Set the "PageMode" property to "PdfPageMode.UseNone" to get the PDF reader to display just the document itself.
        options->set_PageMode(pageMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.PageMode.pdf", options);
        //ExEnd

        String docLocaleName = MakeObject<System::Globalization::CultureInfo>(doc->get_Styles()->get_DefaultFont()->get_LocaleId())->get_Name();

        switch (pageMode)
        {
        case PdfPageMode::FullScreen:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({0})/Metadata 4 0 R>>\r\n", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;

        case PdfPageMode::UseThumbs:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({0})/Metadata 4 0 R>>", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;

        case PdfPageMode::UseOC:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({0})/Metadata 4 0 R>>\r\n", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;

        case PdfPageMode::UseOutlines:
        case PdfPageMode::UseNone:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/Lang({0})/Metadata 4 0 R>>\r\n", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;
        }
    }

    void NoteHyperlinks(bool createNoteHyperlinks)
    {
        //ExStart
        //ExFor:PdfSaveOptions.CreateNoteHyperlinks
        //ExSummary:Shows how to make footnotes and endnotes function as hyperlinks.
        auto doc = MakeObject<Document>(MyDir + u"Footnotes and endnotes.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "CreateNoteHyperlinks" property to "true" to turn all footnote/endnote symbols
        // in the text act as links that, upon clicking, take us to their respective footnotes/endnotes.
        // Set the "CreateNoteHyperlinks" property to "false" not to have footnote/endnote symbols link to anything.
        options->set_CreateNoteHyperlinks(createNoteHyperlinks);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf", options);
        //ExEnd

        if (createNoteHyperlinks)
        {
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[5 0 R /XYZ 85 677 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[5 0 R /XYZ 85 79 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[5 0 R /XYZ 85 654 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[5 0 R /XYZ 85 68 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(
                u"<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19904327 88.66500092 79.69804382]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 202 733 0]>>",
                ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(
                u"<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70004272 88.66500092 68.19904327]/BS <</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 258 711 0]>>",
                ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[5 0 R /XYZ 157 733 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[5 0 R /XYZ 212 711 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
        }
        else
        {
            if (!IsRunningOnMono())
            {
                EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect", ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf"),
                                     "not found in the provided source");
            }
        }
    }

    void CustomPropertiesExport(PdfCustomPropertiesExport pdfCustomPropertiesExportMode)
    {
        //ExStart
        //ExFor:PdfCustomPropertiesExport
        //ExFor:PdfSaveOptions.CustomPropertiesExport
        //ExSummary:Shows how to export custom properties while converting a document to PDF.
        auto doc = MakeObject<Document>();

        doc->get_CustomDocumentProperties()->Add(u"Company", String(u"My value"));

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.None" to discard
        // custom document properties as we save the document to .PDF.
        // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Standard"
        // to preserve custom properties within the output PDF document.
        // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Metadata"
        // to preserve custom properties in an XMP packet.
        options->set_CustomPropertiesExport(pdfCustomPropertiesExportMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf", options);
        //ExEnd

        switch (pdfCustomPropertiesExportMode)
        {
        case PdfCustomPropertiesExport::None:
            if (!IsRunningOnMono())
            {
                EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"Company", ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf"),
                                     "not found in the provided source");

                EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                                                                  ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf"),
                                     "not found in the provided source");
            }
            break;

        case PdfCustomPropertiesExport::Standard:
            TestUtil::FileContainsString(
                u"<</Creator(þÿ\u0000A\u0000s\u0000p\u0000o\u0000s\u0000e\u0000.\u0000W\u0000o\u0000r\u0000d\u0000s)/"
                u"Producer(þÿ\u0000A\u0000s\u0000p\u0000o\u0000s\u0000e\u0000.\u0000W\u0000o\u0000r\u0000d\u0000s\u0000 \u0000f\u0000o\u0000r\u0000",
                ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf");
            TestUtil::FileContainsString(u"/Company (þÿ\u0000M\u0000y\u0000 \u0000v\u0000a\u0000l\u0000u\u0000e)>>",
                                         ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf");
            break;

        case PdfCustomPropertiesExport::Metadata:
            TestUtil::FileContainsString(u"<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                                         ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf");
            break;
        }
    }

    void DrawingMLEffects(DmlEffectsRenderingMode effectsRenderingMode)
    {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:DmlEffectsRenderingMode
        //ExFor:PdfSaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to configure the rendering quality of DrawingML effects in a document as we save it to PDF.
        auto doc = MakeObject<Document>(MyDir + u"DrawingML shape effects.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.None" to discard all DrawingML effects.
        // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Simplified"
        // to render a simplified version of DrawingML effects.
        // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Fine" to
        // render DrawingML effects with more accuracy and also with more processing cost.
        options->set_DmlEffectsRenderingMode(effectsRenderingMode);

        ASSERT_EQ(DmlRenderingMode::DrawingML, options->get_DmlRenderingMode());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.DrawingMLEffects.pdf", options);
        //ExEnd
    }

    void DrawingMLFallback(DmlRenderingMode dmlRenderingMode)
    {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to render fallback shapes when saving to PDF.
        auto doc = MakeObject<Document>(MyDir + u"DrawingML shape fallbacks.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "DmlRenderingMode" property to "DmlRenderingMode.Fallback"
        // to substitute DML shapes with their fallback shapes.
        // Set the "DmlRenderingMode" property to "DmlRenderingMode.DrawingML"
        // to render the DML shapes themselves.
        options->set_DmlRenderingMode(dmlRenderingMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.DrawingMLFallback.pdf", options);
        //ExEnd

        switch (dmlRenderingMode)
        {
        case DmlRenderingMode::DrawingML:
            TestUtil::FileContainsString(u"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABB 11 0 "
                                         u"R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                                         ArtifactsDir + u"PdfSaveOptions.DrawingMLFallback.pdf");
            break;

        case DmlRenderingMode::Fallback:
            TestUtil::FileContainsString(u"5 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABD "
                                         u"13 0 R>>/ExtGState<</GS1 10 0 R/GS2 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                                         ArtifactsDir + u"PdfSaveOptions.DrawingMLFallback.pdf");
            break;
        }
    }

    void ExportDocumentStructure(bool exportDocumentStructure)
    {
        //ExStart
        //ExFor:PdfSaveOptions.ExportDocumentStructure
        //ExSummary:Shows how to preserve document structure elements, which can assist in programmatically interpreting our document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Hello world!");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
        builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "ExportDocumentStructure" property to "true" to make the document structure, such tags, available via the
        // "Content" navigation pane of Adobe Acrobat at the cost of increased file size.
        // Set the "ExportDocumentStructure" property to "false" to not export the document structure.
        options->set_ExportDocumentStructure(exportDocumentStructure);

        // Suppose we export document structure while saving this document. In that case,
        // we can open it using Adobe Acrobat and find tags for elements such as the heading
        // and the next paragraph via "View" -> "Show/Hide" -> "Navigation panes" -> "Tags".
        doc->Save(ArtifactsDir + u"PdfSaveOptions.ExportDocumentStructure.pdf", options);
        //ExEnd

        if (exportDocumentStructure)
        {
            TestUtil::FileContainsString(
                String(u"5 0 obj\r\n") +
                    u"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 10 0 "
                    u"R/GS2 14 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                ArtifactsDir + u"PdfSaveOptions.ExportDocumentStructure.pdf");
        }
        else
        {
            TestUtil::FileContainsString(String(u"5 0 obj\r\n") + u"<</Type /Page/Parent 3 0 R/Contents 6 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAI "
                                                                  u"8 0 R/FAAABB 11 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                                         ArtifactsDir + u"PdfSaveOptions.ExportDocumentStructure.pdf");
        }
    }

    void PdfDigitalSignature()
    {
        //ExStart
        //ExFor:PdfDigitalSignatureDetails
        //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
        //ExFor:PdfDigitalSignatureDetails.HashAlgorithm
        //ExFor:PdfDigitalSignatureDetails.Location
        //ExFor:PdfDigitalSignatureDetails.Reason
        //ExFor:PdfDigitalSignatureDetails.SignatureDate
        //ExFor:PdfDigitalSignatureHashAlgorithm
        //ExFor:PdfSaveOptions.DigitalSignatureDetails
        //ExSummary:Shows how to sign a generated PDF document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Contents of signed PDF.");

        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Configure the "DigitalSignatureDetails" object of the "SaveOptions" object to
        // digitally sign the document as we render it with the "Save" method.
        System::DateTime signingTime = System::DateTime::get_Now();
        options->set_DigitalSignatureDetails(MakeObject<PdfDigitalSignatureDetails>(certificateHolder, u"Test Signing", u"My Office", signingTime));
        options->get_DigitalSignatureDetails()->set_HashAlgorithm(PdfDigitalSignatureHashAlgorithm::Sha256);

        ASSERT_EQ(u"Test Signing", options->get_DigitalSignatureDetails()->get_Reason());
        ASSERT_EQ(u"My Office", options->get_DigitalSignatureDetails()->get_Location());
        ASSERT_EQ(signingTime.ToUniversalTime(), options->get_DigitalSignatureDetails()->get_SignatureDate().ToUniversalTime());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignature.pdf", options);
        //ExEnd

        TestUtil::FileContainsString(String(u"7 0 obj\r\n") + u"<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/DR <<>>/F 132/V 8 0 R/P 5 0 "
                                                              u"R/"
                                                              u"T("
                                                              u"þÿ\u0000A\u0000s\u0000p\u0000o\u0000s\u0000e\u0000D\u0000i\u0000g\u0000i\u0000t\u0000a\u0000l"
                                                              u"\u0000S\u0000i\u0000g\u0000n\u0000a\u0000t\u0000u\u0000r\u0000e)/AP <</N 9 0 R>>>>",
                                     ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignature.pdf");

        ASSERT_FALSE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignature.pdf")->get_HasDigitalSignature());
    }

    void PdfDigitalSignatureTimestamp()
    {
        //ExStart
        //ExFor:PdfDigitalSignatureDetails.TimestampSettings
        //ExFor:PdfDigitalSignatureTimestampSettings
        //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String)
        //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String,TimeSpan)
        //ExFor:PdfDigitalSignatureTimestampSettings.Password
        //ExFor:PdfDigitalSignatureTimestampSettings.ServerUrl
        //ExFor:PdfDigitalSignatureTimestampSettings.Timeout
        //ExFor:PdfDigitalSignatureTimestampSettings.UserName
        //ExSummary:Shows how to sign a saved PDF document digitally and timestamp it.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Signed PDF contents.");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Create a digital signature and assign it to our SaveOptions object to sign the document when we save it to PDF.
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");
        options->set_DigitalSignatureDetails(
            MakeObject<PdfDigitalSignatureDetails>(certificateHolder, u"Test Signing", u"Aspose Office", System::DateTime::get_Now()));

        // Create a timestamp authority-verified timestamp.
        options->get_DigitalSignatureDetails()->set_TimestampSettings(
            MakeObject<PdfDigitalSignatureTimestampSettings>(u"https://freetsa.org/tsr", u"JohnDoe", u"MyPassword"));

        // The default lifespan of the timestamp is 100 seconds.
        ASPOSE_ASSERT_EQ(100.0, options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_Timeout().get_TotalSeconds());

        // We can set our timeout period via the constructor.
        options->get_DigitalSignatureDetails()->set_TimestampSettings(
            MakeObject<PdfDigitalSignatureTimestampSettings>(u"https://freetsa.org/tsr", u"JohnDoe", u"MyPassword", System::TimeSpan::FromMinutes(30)));

        ASPOSE_ASSERT_EQ(1800.0, options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_Timeout().get_TotalSeconds());
        ASSERT_EQ(u"https://freetsa.org/tsr", options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_ServerUrl());
        ASSERT_EQ(u"JohnDoe", options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_UserName());
        ASSERT_EQ(u"MyPassword", options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_Password());

        // The "Save" method will apply our signature to the output document at this time.
        doc->Save(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
        //ExEnd

        ASSERT_FALSE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf")->get_HasDigitalSignature());
        TestUtil::FileContainsString(String(u"7 0 obj\r\n") + u"<</Type /Annot/Subtype /Widget/Rect [0 0 0 0]/FT /Sig/DR <<>>/F 132/V 8 0 R/P 5 0 "
                                                              u"R/"
                                                              u"T("
                                                              u"þÿ\u0000A\u0000s\u0000p\u0000o\u0000s\u0000e\u0000D\u0000i\u0000g\u0000i\u0000t\u0000a\u0000l"
                                                              u"\u0000S\u0000i\u0000g\u0000n\u0000a\u0000t\u0000u\u0000r\u0000e)/AP <</N 9 0 R>>>>",
                                     ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");
    }

    void RenderMetafile(EmfPlusDualRenderingMode renderingMode)
    {
        //ExStart
        //ExFor:EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.UseEmfEmbeddedToWmf
        //ExSummary:Shows how to configure Enhanced Windows Metafile-related rendering options when saving to PDF.
        auto doc = MakeObject<Document>(MyDir + u"EMF.docx");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();

        // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.Emf"
        // to only render the EMF part of an EMF+ dual metafile.
        // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlus" to
        // to render the EMF+ part of an EMF+ dual metafile.
        // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlusWithFallback"
        // to render the EMF+ part of an EMF+ dual metafile if all of the EMF+ records are supported.
        // Otherwise, Aspose.Words will render the EMF part.
        saveOptions->get_MetafileRenderingOptions()->set_EmfPlusDualRenderingMode(renderingMode);

        // Set the "UseEmfEmbeddedToWmf" property to "true" to render embedded EMF data
        // for metafiles that we can render as vector graphics.
        saveOptions->get_MetafileRenderingOptions()->set_UseEmfEmbeddedToWmf(true);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.RenderMetafile.pdf", saveOptions);
        //ExEnd
    }

    void EncryptionPermissions()
    {
        //ExStart
        //ExFor:PdfEncryptionDetails.#ctor
        //ExFor:PdfSaveOptions.EncryptionDetails
        //ExFor:PdfEncryptionDetails.Permissions
        //ExFor:PdfEncryptionDetails.EncryptionAlgorithm
        //ExFor:PdfEncryptionDetails.OwnerPassword
        //ExFor:PdfEncryptionDetails.UserPassword
        //ExFor:PdfEncryptionAlgorithm
        //ExFor:PdfPermissions
        //ExFor:PdfEncryptionDetails
        //ExSummary:Shows how to set permissions on a saved PDF document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");

        auto encryptionDetails = MakeObject<PdfEncryptionDetails>(u"password", String::Empty, PdfEncryptionAlgorithm::RC4_128);

        // Start by disallowing all permissions.
        encryptionDetails->set_Permissions(PdfPermissions::DisallowAll);

        // Extend permissions to allow the editing of annotations.
        encryptionDetails->set_Permissions(PdfPermissions::ModifyAnnotations | PdfPermissions::DocumentAssembly);

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto saveOptions = MakeObject<PdfSaveOptions>();

        // Enable encryption via the "EncryptionDetails" property.
        saveOptions->set_EncryptionDetails(encryptionDetails);

        // When we open this document, we will need to provide the password before accessing its contents.
        doc->Save(ArtifactsDir + u"PdfSaveOptions.EncryptionPermissions.pdf", saveOptions);
        //ExEnd
    }

    void SetNumeralFormat(NumeralFormat numeralFormat)
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.NumeralFormat
        //ExFor:NumeralFormat
        //ExSummary:Shows how to set the numeral format used when saving to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"ar-AR")->get_LCID());
        builder->Writeln(u"1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100");

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "NumeralFormat" property to "NumeralFormat.ArabicIndic" to
        // use glyphs from the U+0660 to U+0669 range as numbers.
        // Set the "NumeralFormat" property to "NumeralFormat.Context" to
        // look up the locale to determine what number of glyphs to use.
        // Set the "NumeralFormat" property to "NumeralFormat.EasternArabicIndic" to
        // use glyphs from the U+06F0 to U+06F9 range as numbers.
        // Set the "NumeralFormat" property to "NumeralFormat.European" to use european numerals.
        // Set the "NumeralFormat" property to "NumeralFormat.System" to determine the symbol set from regional settings.
        options->set_NumeralFormat(numeralFormat);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.SetNumeralFormat.pdf", options);
        //ExEnd
    }

    void ExportPageSet()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageSet
        //ExSummary:Shows how to export Odd pages from the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        for (int i = 0; i < 5; i++)
        {
            builder->Writeln(String::Format(u"Page {0} ({1})", i + 1, (i % 2 == 0 ? String(u"odd") : String(u"even"))));
            if (i < 4)
            {
                builder->InsertBreak(BreakType::PageBreak);
            }
        }

        // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Below are three PageSet properties that we can use to filter out a set of pages from
        // our document to save in an output PDF document based on the parity of their page numbers.
        // 1 -  Save only the even-numbered pages:
        options->set_PageSet(PageSet::get_Even());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ExportPageSet.Even.pdf", options);

        // 2 -  Save only the odd-numbered pages:
        options->set_PageSet(PageSet::get_Odd());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ExportPageSet.Odd.pdf", options);

        // 3 -  Save every page:
        options->set_PageSet(PageSet::get_All());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ExportPageSet.All.pdf", options);
        //ExEnd
    }
};

} // namespace ApiExamples
