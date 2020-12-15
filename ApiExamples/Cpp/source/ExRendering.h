#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceType.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/TableSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/NumeralFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfEncryptionAlgorithm.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfEncryptionDetails.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfFontEmbeddingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfPermissions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfTextCompression.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/TiffCompression.h>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Rendering/PageInfo.h>
#include <drawing/bitmap.h>
#include <drawing/color.h>
#include <drawing/graphics.h>
#include <drawing/graphics_unit.h>
#include <drawing/image.h>
#include <drawing/pen.h>
#include <drawing/pens.h>
#include <drawing/size.h>
#include <drawing/size_f.h>
#include <drawing/solid_brush.h>
#include <drawing/text/text_rendering_hint.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/details/dispose_guard.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/seekorigin.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/math.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExRendering : public ApiExampleBase
{
public:
    void SaveToPdfWithOutline()
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:PdfSaveOptions
        //ExFor:OutlineOptions.HeadingsOutlineLevels
        //ExFor:OutlineOptions.ExpandedOutlineLevels
        //ExSummary:Converts a whole document to PDF with three levels in the document outline.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<PdfSaveOptions>();
        options->get_OutlineOptions()->set_HeadingsOutlineLevels(3);
        options->get_OutlineOptions()->set_ExpandedOutlineLevels(1);

        doc->Save(ArtifactsDir + u"Rendering.SaveToPdfWithOutline.pdf", options);
        //ExEnd
    }

    void SaveToPdfStreamOnePage()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageIndex
        //ExFor:FixedPageSaveOptions.PageCount
        //ExFor:Document.Save(Stream, SaveOptions)
        //ExSummary:Converts just one page (third page in this example) of the document to PDF.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        {
            SharedPtr<System::IO::Stream> stream = System::IO::File::Create(ArtifactsDir + u"Rendering.SaveToPdfStreamOnePage.pdf");
            auto options = MakeObject<PdfSaveOptions>();
            options->set_PageIndex(2);
            options->set_PageCount(1);
            doc->Save(stream, options);
        }
        //ExEnd
    }

    void SaveToPdfNoCompression()
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:PdfSaveOptions.TextCompression
        //ExFor:PdfTextCompression
        //ExSummary:Saves a document to PDF without compression.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<PdfSaveOptions>();
        options->set_TextCompression(PdfTextCompression::None);

        doc->Save(ArtifactsDir + u"Rendering.SaveToPdfNoCompression.pdf", options);
        //ExEnd
    }

    void PdfCustomOptions()
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreserveFormFields
        //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
        // Open the document
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Option 1: Save document to file in the PDF format with default options
        doc->Save(ArtifactsDir + u"Rendering.PdfDefaultOptions.pdf");

        // Option 2: Save the document to stream in the PDF format with default options
        auto stream = MakeObject<System::IO::MemoryStream>();
        doc->Save(stream, SaveFormat::Pdf);
        // Rewind the stream position back to the beginning, ready for use
        stream->Seek(static_cast<int64_t>(0), System::IO::SeekOrigin::Begin);

        // Option 3: Save document to the PDF format with specified options
        // Render the first page only and preserve form fields as usable controls and not as plain text
        auto pdfOptions = MakeObject<PdfSaveOptions>();
        pdfOptions->set_PageIndex(0);
        pdfOptions->set_PageCount(1);
        pdfOptions->set_PreserveFormFields(true);
        doc->Save(ArtifactsDir + u"Rendering.PdfCustomOptions.pdf", pdfOptions);
        //ExEnd
    }

    void SaveAsXps()
    {
        //ExStart
        //ExFor:XpsSaveOptions
        //ExFor:XpsSaveOptions.#ctor()
        //ExFor:XpsSaveOptions.OutlineOptions
        //ExFor:XpsSaveOptions.SaveFormat
        //ExSummary:Shows how to save a document to the XPS format in different ways.
        // Open the document
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Save document to file in the XPS format with default options
        doc->Save(ArtifactsDir + u"Rendering.SaveAsXps.DefaultOptions.xps");

        // Save document to stream in the XPS format with default options
        auto docStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"Rendering.SaveAsXps.FromStream.xps", System::IO::FileMode::Create);
        doc->Save(docStream, SaveFormat::Xps);
        docStream->Close();

        // Save document to file in the XPS format with specified options
        // Render 3 pages starting from page 2; pages 2, 3 and 4
        auto xpsOptions = MakeObject<XpsSaveOptions>();
        xpsOptions->set_SaveFormat(SaveFormat::Xps);
        xpsOptions->set_PageIndex(1);
        xpsOptions->set_PageCount(3);

        // All paragraphs in the "Heading 1" style will be included in the outline but "Heading 2" and onwards will not
        xpsOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(1);

        doc->Save(ArtifactsDir + u"Rendering.SaveAsXps.PartialDocument.xps", xpsOptions);
        //ExEnd
    }

    void SaveAsXpsBookFold()
    {
        //ExStart
        //ExFor:XpsSaveOptions.#ctor(SaveFormat)
        //ExFor:XpsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
        // Open a document with multiple paragraphs
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Configure both page setup and XpsSaveOptions to create a book fold
        for (auto s : System::IterateOver<Section>(doc->get_Sections()))
        {
            s->get_PageSetup()->set_MultiplePages(MultiplePagesType::BookFoldPrinting);
        }

        auto xpsOptions = MakeObject<XpsSaveOptions>(SaveFormat::Xps);
        xpsOptions->set_UseBookFoldPrintingSettings(true);

        // Once we print this document, we can turn it into a booklet by stacking the pages
        // in the order they come out of the printer and then folding down the middle
        doc->Save(ArtifactsDir + u"Rendering.SaveAsXpsBookFold.xps", xpsOptions);
        //ExEnd
    }

    void SaveAsImage()
    {
        //ExStart
        //ExFor:ImageSaveOptions.#ctor
        //ExFor:Document.Save(String)
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Shows how to save a document to the JPEG format using the Save method and the ImageSaveOptions class.
        // Open the document
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        // Save as a JPEG image file with default options
        doc->Save(ArtifactsDir + u"Rendering.SaveAsImage.DefaultJpgOptions.jpg");

        // Save document to stream as a JPEG with default options
        auto docStream = MakeObject<System::IO::MemoryStream>();
        doc->Save(docStream, SaveFormat::Jpeg);
        // Rewind the stream position back to the beginning, ready for use
        docStream->Seek(static_cast<int64_t>(0), System::IO::SeekOrigin::Begin);

        // Save document to a JPEG image with specified options
        // Render the third page only and set the JPEG quality to 80%
        // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor
        // to signal what type of image to save as
        auto imageOptions = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);
        imageOptions->set_PageIndex(2);
        imageOptions->set_PageCount(1);
        imageOptions->set_JpegQuality(80);
        doc->Save(ArtifactsDir + u"Rendering.SaveAsImage.CustomJpgOptions.jpg", imageOptions);
        //ExEnd
    }

    void SaveToTiffDefault()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->Save(ArtifactsDir + u"Rendering.SaveToTiffDefault.tiff");
    }

    void SaveToTiffCompression()
    {
        //ExStart
        //ExFor:TiffCompression
        //ExFor:ImageSaveOptions.TiffCompression
        //ExFor:ImageSaveOptions.PageIndex
        //ExFor:ImageSaveOptions.PageCount
        //ExSummary:Converts a page of a Word document into a TIFF image and uses the CCITT compression.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        options->set_TiffCompression(TiffCompression::Ccitt3);
        options->set_PageIndex(0);
        options->set_PageCount(1);

        doc->Save(ArtifactsDir + u"Rendering.SaveToTiffCompression.tiff", options);
        //ExEnd
    }

    void SaveToImageResolution()
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.Resolution
        //ExSummary:Renders a page of a Word document into a PNG image at a specific resolution.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        options->set_Resolution(300);
        options->set_PageCount(1);

        doc->Save(ArtifactsDir + u"Rendering.SaveToImageResolution.png", options);
        //ExEnd
    }

    void SaveToEmf()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Converts every page of a DOC file into a separate scalable EMF file.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Emf);
        options->set_PageCount(1);

        for (int i = 0; i < doc->get_PageCount(); i++)
        {
            options->set_PageIndex(i);
            doc->Save(ArtifactsDir + u"Rendering.SaveToEmf." + i + u".emf", options);
        }
        //ExEnd
    }

    void SaveToImageJpegQuality()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.JpegQuality
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.JpegQuality
        //ExSummary:Converts a page of a Word document into JPEG images of different qualities.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);

        // Try worst quality
        saveOptions->set_JpegQuality(0);
        doc->Save(ArtifactsDir + u"Rendering.SaveToImageJpegQuality.0.jpeg", saveOptions);

        // Try best quality
        saveOptions->set_JpegQuality(100);
        doc->Save(ArtifactsDir + u"Rendering.SaveToImageJpegQuality.100.jpeg", saveOptions);
        //ExEnd
    }

    void SaveToImagePaperColor()
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.PaperColor
        //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto imgOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);

        imgOptions->set_PaperColor(System::Drawing::Color::get_Transparent());
        doc->Save(ArtifactsDir + u"Rendering.SaveToImagePaperColor.Transparent.png", imgOptions);

        imgOptions->set_PaperColor(System::Drawing::Color::get_LightCoral());
        doc->Save(ArtifactsDir + u"Rendering.SaveToImagePaperColor.Coral.png", imgOptions);
        //ExEnd
    }

    void SaveToImageStream()
    {
        //ExStart
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExSummary:Saves a document page as a BMP image into a stream.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto stream = MakeObject<System::IO::MemoryStream>();
        doc->Save(stream, SaveFormat::Bmp);

        // Rewind the stream and create a .NET image from it
        stream->set_Position(0);

        // Read the stream back into an image
        {
            SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);
            // ...Do something
        }
        //ExEnd
    }

    void RenderToSize()
    {
        //ExStart
        //ExFor:Document.RenderToSize
        //ExSummary:Render to a bitmap at a specified location and size.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        {
            auto bmp = MakeObject<System::Drawing::Bitmap>(700, 700);
            // User has some sort of a Graphics object. In this case created from a bitmap
            {
                SharedPtr<System::Drawing::Graphics> gr = System::Drawing::Graphics::FromImage(bmp);
                // The user can specify any options on the Graphics object including
                // transform, anti-aliasing, page units, etc.
                gr->set_TextRenderingHint(System::Drawing::Text::TextRenderingHint::AntiAliasGridFit);

                // If we want to fit the page into a 3" x 3" square on the screen, we will need to set the measurement units to inches
                gr->set_PageUnit(System::Drawing::GraphicsUnit::Inch);

                // The output should be offset 0.5" from the edge and rotated
                gr->TranslateTransform(0.5f, 0.5f);
                gr->RotateTransform(10.0f);

                // This is our test rectangle
                gr->DrawRectangle(MakeObject<System::Drawing::Pen>(System::Drawing::Color::get_Black(), 3.f / 72.f), 0.f, 0.f, 3.f, 3.f);

                // User specifies (in world coordinates) where on the Graphics to render and what size
                float returnedScale = doc->RenderToSize(0, gr, 0.f, 0.f, 3.f, 3.f);

                // This is the calculated scale factor to fit 297mm into 3"
                std::cout << String::Format(u"The image was rendered at {0:P0} zoom.", returnedScale) << std::endl;

                // One more example, this time in millimeters
                gr->set_PageUnit(System::Drawing::GraphicsUnit::Millimeter);

                gr->ResetTransform();

                // Move the origin 10mm
                gr->TranslateTransform(10.0f, 10.0f);

                // Apply both scale transform and page scale for fun
                gr->ScaleTransform(0.5f, 0.5f);
                gr->set_PageScale(2.f);

                // This is our test rectangle
                gr->DrawRectangle(MakeObject<System::Drawing::Pen>(System::Drawing::Color::get_Black(), 1.0f), 90, 10, 50, 100);

                // User specifies (in world coordinates) where on the Graphics to render and what size
                doc->RenderToSize(1, gr, 90.0f, 10.0f, 50.0f, 100.0f);

                bmp->Save(ArtifactsDir + u"Rendering.RenderToSize.png");
            }
        }
        //ExEnd
    }

    void Thumbnails()
    {
        //ExStart
        //ExFor:Document.RenderToScale
        //ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages.
        // The user opens or builds a document
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // This defines the number of columns to display the thumbnails in
        const int thumbColumns = 2;

        // Calculate the required number of rows for thumbnails
        // We can now get the number of pages in the document
        int remainder;
        int thumbRows = System::Math::DivRem(doc->get_PageCount(), thumbColumns, remainder);
        if (remainder > 0)
        {
            thumbRows++;
        }

        // Define a zoom factor for the thumbnails
        const float scale = 0.25f;

        // We can use the size of the first page to calculate the size of the thumbnail,
        // assuming that all pages in the document are of the same size
        System::Drawing::Size thumbSize = doc->GetPageInfo(0)->GetSizeInPixels(scale, 96.0f);

        // Calculate the size of the image that will contain all the thumbnails
        int imgWidth = thumbSize.get_Width() * thumbColumns;
        int imgHeight = thumbSize.get_Height() * thumbRows;

        {
            auto img = MakeObject<System::Drawing::Bitmap>(imgWidth, imgHeight);
            // The Graphics object, which we will draw on, can be created from a bitmap, metafile, printer, or window
            {
                SharedPtr<System::Drawing::Graphics> gr = System::Drawing::Graphics::FromImage(img);
                gr->set_TextRenderingHint(System::Drawing::Text::TextRenderingHint::AntiAliasGridFit);

                // Fill the "paper" with white, otherwise it will be transparent
                gr->FillRectangle(MakeObject<System::Drawing::SolidBrush>(System::Drawing::Color::get_White()), 0, 0, imgWidth, imgHeight);

                for (int pageIndex = 0; pageIndex < doc->get_PageCount(); pageIndex++)
                {
                    int columnIdx;
                    int rowIdx = System::Math::DivRem(pageIndex, thumbColumns, columnIdx);

                    // Specify where we want the thumbnail to appear
                    float thumbLeft = static_cast<float>(columnIdx * thumbSize.get_Width());
                    float thumbTop = static_cast<float>(rowIdx * thumbSize.get_Height());

                    System::Drawing::SizeF size = doc->RenderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);

                    // Draw the page rectangle
                    gr->DrawRectangle(System::Drawing::Pens::get_Black(), thumbLeft, thumbTop, size.get_Width(), size.get_Height());
                }

                img->Save(ArtifactsDir + u"Rendering.Thumbnails.png");
            }
        }
        //ExEnd
    }

    void UpdatePageLayout()
    {
        //ExStart
        //ExFor:StyleCollection.Item(String)
        //ExFor:SectionCollection.Item(Int32)
        //ExFor:Document.UpdatePageLayout
        //ExSummary:Shows when to request page layout of the document to be recalculated.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Saving a document to PDF or to image or printing for the first time will automatically
        // layout document pages and this information will be cached inside the document
        doc->Save(ArtifactsDir + u"Rendering.UpdatePageLayout.1.pdf");

        // Modify the document in any way
        doc->get_Styles()->idx_get(u"Normal")->get_Font()->set_Size(6);
        doc->get_Sections()->idx_get(0)->get_PageSetup()->set_Orientation(Orientation::Landscape);

        // In the current version of Aspose.Words, modifying the document does not automatically rebuild
        // the cached page layout. If you want to save to PDF or render a modified document again,
        // you need to manually request page layout to be updated
        doc->UpdatePageLayout();

        doc->Save(ArtifactsDir + u"Rendering.UpdatePageLayout.2.pdf");
        //ExEnd
    }

    void SetTrueTypeFontsFolder()
    {
        // Store the font sources currently used so we can restore them later
        ArrayPtr<SharedPtr<FontSourceBase>> fontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolder(String, Boolean)
        //ExSummary:Demonstrates how to set the folder Aspose.Words uses to look for TrueType fonts during rendering or embedding of fonts.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Note that this setting will override any default font sources that are being searched by default
        // Now only these folders will be searched for fonts when rendering or embedding fonts
        // To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead
        FontSettings::get_DefaultInstance()->SetFontsFolder(u"C:\\MyFonts\\", false);

        doc->Save(ArtifactsDir + u"Rendering.SetTrueTypeFontsFolder.pdf");
        //ExEnd

        // Restore the original sources used to search for fonts
        FontSettings::get_DefaultInstance()->SetFontsSources(fontSources);
    }

    void SetFontsFoldersMultipleFolders()
    {
        // Store the font sources currently used so we can restore them later
        ArrayPtr<SharedPtr<FontSourceBase>> fontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
        //ExSummary:Demonstrates how to set Aspose.Words to look in multiple folders for TrueType fonts when rendering or embedding fonts.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Note that this setting will override any default font sources that are being searched by default
        // Now only these folders will be searched for fonts when rendering or embedding fonts
        // To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead
        FontSettings::get_DefaultInstance()->SetFontsFolders(MakeArray<String>({u"C:\\MyFonts\\", u"D:\\Misc\\Fonts\\"}), true);

        doc->Save(ArtifactsDir + u"Rendering.SetFontsFoldersMultipleFolders.pdf");
        //ExEnd

        // Restore the original sources used to search for fonts
        FontSettings::get_DefaultInstance()->SetFontsSources(fontSources);
    }

    void SetFontsFoldersSystemAndCustomFolder()
    {
        // Store the font sources currently used so we can restore them later
        ArrayPtr<SharedPtr<FontSourceBase>> origFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.GetFontsSources()
        //ExFor:FontSettings.SetFontsSources(FontSourceBase[])
        //ExSummary:Demonstrates how to set Aspose.Words to look for TrueType fonts in system folders as well as a custom defined folder when scanning for fonts.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Retrieve the array of environment-dependent font sources that are searched by default
        // For example, this will contain a "Windows\Fonts\" source on a Windows machines
        // We add this array to a new ArrayList to make adding or removing font entries much easier
        SharedPtr<System::Collections::Generic::List<SharedPtr<FontSourceBase>>> fontSources =
            MakeObject<System::Collections::Generic::List<SharedPtr<FontSourceBase>>>(FontSettings::get_DefaultInstance()->GetFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
        auto folderFontSource = MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true);

        // Add the custom folder which contains our fonts to the list of existing font sources
        fontSources->Add(folderFontSource);

        // Convert the ArrayList of source back into a primitive array of FontSource objects
        ArrayPtr<SharedPtr<FontSourceBase>> updatedFontSources = fontSources->ToArray();

        // Apply the new set of font sources to use
        FontSettings::get_DefaultInstance()->SetFontsSources(updatedFontSources);

        doc->Save(ArtifactsDir + u"Rendering.SetFontsFoldersSystemAndCustomFolder.pdf");
        //ExEnd

        // The first source should be a system font source
        ASPOSE_ASSERT_ISINSTANCEOF(System::ObjectExt::GetType<SystemFontSource>(), FontSettings::get_DefaultInstance()->GetFontsSources()->idx_get(0));
        // The second source should be our folder font source
        ASPOSE_ASSERT_ISINSTANCEOF(System::ObjectExt::GetType<FolderFontSource>(), FontSettings::get_DefaultInstance()->GetFontsSources()->idx_get(1));

        SharedPtr<FolderFontSource> folderSource =
            (System::DynamicCast<FolderFontSource>(FontSettings::get_DefaultInstance()->GetFontsSources()->idx_get(1)));
        ASSERT_EQ(u"C:\\MyFonts\\", folderSource->get_FolderPath());
        ASSERT_TRUE(folderSource->get_ScanSubfolders());

        // Restore the original sources used to search for fonts
        FontSettings::get_DefaultInstance()->SetFontsSources(origFontSources);
    }

    void SetSpecifyFontFolder()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->SetFontsFolder(FontsDir, false);

        // Using load options
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(fontSettings);

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);

        SharedPtr<FolderFontSource> folderSource = (System::DynamicCast<FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0)));

        ASSERT_EQ(FontsDir, folderSource->get_FolderPath());
        ASSERT_FALSE(folderSource->get_ScanSubfolders());
    }

    void SetFontSubstitutes()
    {
        //ExStart
        //ExFor:Document.FontSettings
        //ExFor:TableSubstitutionRule.SetSubstitutes(String, String[])
        //ExSummary:Shows how to define alternative fonts if original does not exist
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->SetSubstitutes(u"Times New Roman",
                                                                                          MakeArray<String>({u"Slab", u"Arvo"}));
        //ExEnd
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->set_FontSettings(fontSettings);

        // Check that font source are default
        ArrayPtr<SharedPtr<FontSourceBase>> fontSource = doc->get_FontSettings()->GetFontsSources();
        ASSERT_EQ(u"SystemFonts", System::ObjectExt::ToString(fontSource[0]->get_Type()));

        ASSERT_EQ(u"Times New Roman", doc->get_FontSettings()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->get_DefaultFontName());

        ArrayPtr<String> alternativeFonts =
            doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Times New Roman")->LINQ_ToArray();
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Slab", u"Arvo"}), alternativeFonts);
    }

    void SetSpecifyFontFolders()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->SetFontsFolders(MakeArray<String>({FontsDir, u"C:\\Windows\\Fonts\\"}), true);

        // Using load options
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(fontSettings);
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);

        SharedPtr<FolderFontSource> folderSource = (System::DynamicCast<FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0)));
        ASSERT_EQ(FontsDir, folderSource->get_FolderPath());
        ASSERT_TRUE(folderSource->get_ScanSubfolders());

        folderSource = (System::DynamicCast<FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(1)));
        ASSERT_EQ(u"C:\\Windows\\Fonts\\", folderSource->get_FolderPath());
        ASSERT_TRUE(folderSource->get_ScanSubfolders());
    }

    void AddFontSubstitutes()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->SetSubstitutes(u"Slab",
                                                                                          MakeArray<String>({u"Times New Roman", u"Arial"}));
        fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Arvo", MakeArray<String>({u"Open Sans", u"Arial"}));

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->set_FontSettings(fontSettings);

        ArrayPtr<String> alternativeFonts =
            doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Slab")->LINQ_ToArray();
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Times New Roman", u"Arial"}), alternativeFonts);

        alternativeFonts = doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Arvo")->LINQ_ToArray();
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Open Sans", u"Arial"}), alternativeFonts);
    }

    void SetDefaultFontName()
    {
        //ExStart
        //ExFor:DefaultFontSubstitutionRule.DefaultFontName
        //ExSummary:Demonstrates how to specify what font to substitute for a missing font during rendering.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead
        FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial Unicode MS");

        // Now the set default font is used in place of any missing fonts during any rendering calls
        doc->Save(ArtifactsDir + u"Rendering.SetDefaultFontName.pdf");
        doc->Save(ArtifactsDir + u"Rendering.SetDefaultFontName.xps");
        //ExEnd
    }

    void UpdatePageLayoutWarnings()
    {
        // Store the font sources currently used so we can restore them later
        ArrayPtr<SharedPtr<FontSourceBase>> origFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        // Load the document to render
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        auto callback = MakeObject<ExRendering::HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts
        FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default
        // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
        FontSettings::get_DefaultInstance()->SetFontsFolder(String::Empty, false);

        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
        // are stored until the document save and then sent to the appropriate WarningCallback
        doc->UpdatePageLayout();

        // Even though the document was rendered previously, any save warnings are notified to the user during document save
        doc->Save(ArtifactsDir + u"Rendering.UpdatePageLayoutWarnings.pdf");

        ASSERT_GT(callback->FontWarnings->get_Count(), 0);
        ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_WarningType() == WarningType::FontSubstitution);
        ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_Description().Contains(u"has not been found"));

        // Restore default fonts
        FontSettings::get_DefaultInstance()->SetFontsSources(origFontSources);
    }

    class HandleDocumentWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> FontWarnings;

        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // We are only interested in fonts being substituted
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                std::cout << (String(u"Font substitution: ") + info->get_Description()) << std::endl;
                FontWarnings->Warning(info);
                //ExSkip
            }
        }

        HandleDocumentWarnings() : FontWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };

    void EmbedFullFonts()
    {
        //ExStart
        //ExFor:PdfSaveOptions.#ctor
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExSummary:Demonstrates how to set Aspose.Words to embed full fonts in the output PDF document.
        // Load the document to render
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true
        // The property below can be changed each time a document is rendered
        auto options = MakeObject<PdfSaveOptions>();
        options->set_EmbedFullFonts(true);

        // The output PDF will be embedded with all fonts found in the document
        doc->Save(ArtifactsDir + u"Rendering.EmbedFullFonts.pdf", options);
        //ExEnd
    }

    void SubsetFonts()
    {
        //ExStart
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExSummary:Demonstrates how to set Aspose.Words to subset fonts in the output PDF.
        // Load the document to render
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
        auto options = MakeObject<PdfSaveOptions>();
        options->set_EmbedFullFonts(false);

        // The output PDF will contain subsets of the fonts in the document
        // Only the glyphs used in the document are included in the PDF fonts
        doc->Save(ArtifactsDir + u"Rendering.SubsetFonts.pdf", options);
        //ExEnd
    }

    void DisableEmbedWindowsFonts()
    {
        //ExStart
        //ExFor:PdfSaveOptions.FontEmbeddingMode
        //ExFor:PdfFontEmbeddingMode
        //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
        // Load the document to render
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false
        auto options = MakeObject<PdfSaveOptions>();
        options->set_FontEmbeddingMode(PdfFontEmbeddingMode::EmbedNone);

        // The output PDF will be saved without embedding standard windows fonts
        doc->Save(ArtifactsDir + u"Rendering.DisableEmbedWindowsFonts.pdf", options);
        //ExEnd
    }

    void DisableEmbedCoreFonts()
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseCoreFonts
        //ExSummary:Shows how to set Aspose.Words to avoid embedding core fonts and let the reader substitute PDF Type 1 fonts instead.
        // Load the document to render
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // To disable embedding of core fonts and substitute PDF type 1 fonts set UseCoreFonts to true
        auto options = MakeObject<PdfSaveOptions>();
        options->set_UseCoreFonts(true);

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        doc->Save(ArtifactsDir + u"Rendering.DisableEmbedCoreFonts.pdf", options);
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
        //ExSummary:Demonstrates how to set permissions on a PDF document generated by Aspose.Words.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();

        // Create encryption details and set owner password
        auto encryptionDetails = MakeObject<PdfEncryptionDetails>(u"password", String::Empty, PdfEncryptionAlgorithm::RC4_128);

        // Start by disallowing all permissions
        encryptionDetails->set_Permissions(PdfPermissions::DisallowAll);

        // Extend permissions to allow editing or modifying annotations
        encryptionDetails->set_Permissions(PdfPermissions::ModifyAnnotations | PdfPermissions::DocumentAssembly);
        saveOptions->set_EncryptionDetails(encryptionDetails);

        // Render the document to PDF format with the specified permissions
        doc->Save(ArtifactsDir + u"Rendering.EncryptionPermissions.pdf", saveOptions);
        //ExEnd
    }

    void SetNumeralFormat()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.NumeralFormat
        //ExFor:NumeralFormat
        //ExSummary:Demonstrates how to set the numeral format used when saving to PDF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100");

        auto options = MakeObject<PdfSaveOptions>();
        options->set_NumeralFormat(NumeralFormat::EasternArabicIndic);

        doc->Save(ArtifactsDir + u"Rendering.SetNumeralFormat.pdf", options);
        //ExEnd
    }
};

} // namespace ApiExamples
