#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/CssSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Saving/DocumentPartSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/ICssSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/IDocumentPartSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/IPageSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/PageSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/XamlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/XpsSaveOptions.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/enum_helpers.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/path.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExSavingCallback : public ApiExampleBase
{
public:
    void CheckThatAllMethodsArePresent()
    {
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());

        auto imageSaveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        imageSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());

        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());

        auto psSaveOptions = MakeObject<PsSaveOptions>();
        psSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());

        auto svgSaveOptions = MakeObject<SvgSaveOptions>();
        svgSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());

        auto xamlFixedSaveOptions = MakeObject<XamlFixedSaveOptions>();
        xamlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());

        auto xpsSaveOptions = MakeObject<XpsSaveOptions>();
        xpsSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());
    }

    //ExStart
    //ExFor:IPageSavingCallback
    //ExFor:IPageSavingCallback.PageSaving(PageSavingArgs)
    //ExFor:PageSavingArgs
    //ExFor:PageSavingArgs.PageFileName
    //ExFor:PageSavingArgs.KeepPageStreamOpen
    //ExFor:PageSavingArgs.PageIndex
    //ExFor:PageSavingArgs.PageStream
    //ExFor:FixedPageSaveOptions.PageSavingCallback
    //ExSummary:Shows how to use a callback to save a document to HTML page by page.
    void PageFileNames()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2.");
        builder->InsertImage(ImageDir + u"Logo.jpg");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 3.");

        // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we convert the document to HTML.
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();

        // We will save each page in this document to a separate HTML file in the local file system.
        // Set a callback that allows us to name each output HTML document.
        htmlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomFileNamePageSavingCallback>());

        doc->Save(ArtifactsDir + u"SavingCallback.PageFileNames.html", htmlFixedSaveOptions);

        ArrayPtr<String> filePaths = System::IO::Directory::GetFiles(ArtifactsDir)
                                         ->LINQ_Where([](String s) { return s.StartsWith(ArtifactsDir + u"SavingCallback.PageFileNames.Page_"); })
                                         ->LINQ_ToArray();

        ASSERT_EQ(3, filePaths->get_Length());
    }

    /// <summary>
    /// Saves all pages to a file and directory specified within.
    /// </summary>
    class CustomFileNamePageSavingCallback : public IPageSavingCallback
    {
    public:
        void PageSaving(SharedPtr<PageSavingArgs> args) override
        {
            String outFileName = String::Format(u"{0}SavingCallback.PageFileNames.Page_{1}.html", ArtifactsDir, args->get_PageIndex());

            // Below are two ways of specifying where Aspose.Words will save each page of the document.
            // 1 -  Set a filename for the output page file:
            args->set_PageFileName(outFileName);

            // 2 -  Create a custom stream for the output page file:
            args->set_PageStream(MakeObject<System::IO::FileStream>(outFileName, System::IO::FileMode::Create));

            ASSERT_FALSE(args->get_KeepPageStreamOpen());
        }
    };
    //ExEnd

    //ExStart
    //ExFor:DocumentPartSavingArgs
    //ExFor:DocumentPartSavingArgs.Document
    //ExFor:DocumentPartSavingArgs.DocumentPartFileName
    //ExFor:DocumentPartSavingArgs.DocumentPartStream
    //ExFor:DocumentPartSavingArgs.KeepDocumentPartStreamOpen
    //ExFor:IDocumentPartSavingCallback
    //ExFor:IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs)
    //ExFor:IImageSavingCallback
    //ExFor:IImageSavingCallback.ImageSaving
    //ExFor:ImageSavingArgs
    //ExFor:ImageSavingArgs.ImageFileName
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.DocumentPartSavingCallback
    //ExFor:HtmlSaveOptions.ImageSavingCallback
    //ExSummary:Shows how to split a document into parts and save them.
    void DocumentPartsFileNames()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        String outFileName = u"SavingCallback.DocumentPartsFileNames.html";

        // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we convert the document to HTML.
        auto options = MakeObject<HtmlSaveOptions>();

        // If we save the document normally, there will be one output HTML
        // document with all the source document's contents.
        // Set the "DocumentSplitCriteria" property to "DocumentSplitCriteria.SectionBreak" to
        // save our document to multiple HTML files: one for each section.
        options->set_DocumentSplitCriteria(DocumentSplitCriteria::SectionBreak);

        // Assign a custom callback to the "DocumentPartSavingCallback" property to alter the document part saving logic.
        options->set_DocumentPartSavingCallback(MakeObject<ExSavingCallback::SavedDocumentPartRename>(outFileName, options->get_DocumentSplitCriteria()));

        // If we convert a document that contains images into html, we will end up with one html file which links to several images.
        // Each image will be in the form of a file in the local file system.
        // There is also a callback that can customize the name and file system location of each image.
        options->set_ImageSavingCallback(MakeObject<ExSavingCallback::SavedImageRename>(outFileName));

        doc->Save(ArtifactsDir + outFileName, options);
    }

    /// <summary>
    /// Sets custom filenames for output documents that the saving operation splits a document into.
    /// </summary>
    class SavedDocumentPartRename : public IDocumentPartSavingCallback
    {
    public:
        SavedDocumentPartRename(String outFileName, DocumentSplitCriteria documentSplitCriteria)
            : mCount(0), mDocumentSplitCriteria(((Aspose::Words::Saving::DocumentSplitCriteria)0))
        {
            mOutFileName = outFileName;
            mDocumentSplitCriteria = documentSplitCriteria;
        }

        void DocumentPartSaving(SharedPtr<DocumentPartSavingArgs> args) override
        {
            // We can access the entire source document via the "Document" property.
            ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));

            String partType = String::Empty;

            switch (mDocumentSplitCriteria)
            {
            case DocumentSplitCriteria::PageBreak:
                partType = u"Page";
                break;

            case DocumentSplitCriteria::ColumnBreak:
                partType = u"Column";
                break;

            case DocumentSplitCriteria::SectionBreak:
                partType = u"Section";
                break;

            case DocumentSplitCriteria::HeadingParagraph:
                partType = u"Paragraph from heading";
                break;

            default:
                break;
            }

            String partFileName = String::Format(u"{0} part {1}, of type {2}{3}", mOutFileName, ++mCount, partType,
                                                 System::IO::Path::GetExtension(args->get_DocumentPartFileName()));

            // Below are two ways of specifying where Aspose.Words will save each part of the document.
            // 1 -  Set a filename for the output part file:
            args->set_DocumentPartFileName(partFileName);

            // 2 -  Create a custom stream for the output part file:
            args->set_DocumentPartStream(MakeObject<System::IO::FileStream>(ArtifactsDir + partFileName, System::IO::FileMode::Create));

            ASSERT_TRUE(args->get_DocumentPartStream()->get_CanWrite());
            ASSERT_FALSE(args->get_KeepDocumentPartStreamOpen());
        }

    private:
        int mCount;
        String mOutFileName;
        DocumentSplitCriteria mDocumentSplitCriteria;
    };

    /// <summary>
    /// Sets custom filenames for image files that an HTML conversion creates.
    /// </summary>
    class SavedImageRename : public IImageSavingCallback
    {
    public:
        SavedImageRename(String outFileName) : mCount(0)
        {
            mOutFileName = outFileName;
        }

        void ImageSaving(SharedPtr<ImageSavingArgs> args) override
        {
            String imageFileName = String::Format(u"{0} shape {1}, of type {2}{3}", mOutFileName, ++mCount, args->get_CurrentShape()->get_ShapeType(),
                                                  System::IO::Path::GetExtension(args->get_ImageFileName()));

            // Below are two ways of specifying where Aspose.Words will save each part of the document.
            // 1 -  Set a filename for the output image file:
            args->set_ImageFileName(imageFileName);

            // 2 -  Create a custom stream for the output image file:
            args->set_ImageStream(MakeObject<System::IO::FileStream>(ArtifactsDir + imageFileName, System::IO::FileMode::Create));

            ASSERT_TRUE(args->get_ImageStream()->get_CanWrite());
            ASSERT_TRUE(args->get_IsImageAvailable());
            ASSERT_FALSE(args->get_KeepImageStreamOpen());
        }

    private:
        int mCount;
        String mOutFileName;
    };
    //ExEnd

    //ExStart
    //ExFor:CssSavingArgs
    //ExFor:CssSavingArgs.CssStream
    //ExFor:CssSavingArgs.Document
    //ExFor:CssSavingArgs.IsExportNeeded
    //ExFor:CssSavingArgs.KeepCssStreamOpen
    //ExFor:CssStyleSheetType
    //ExFor:HtmlSaveOptions.CssSavingCallback
    //ExFor:HtmlSaveOptions.CssStyleSheetFileName
    //ExFor:HtmlSaveOptions.CssStyleSheetType
    //ExFor:ICssSavingCallback
    //ExFor:ICssSavingCallback.CssSaving(CssSavingArgs)
    //ExSummary:Shows how to work with CSS stylesheets that an HTML conversion creates.
    void ExternalCssFilenames()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we convert the document to HTML.
        auto options = MakeObject<HtmlSaveOptions>();

        // Set the "CssStylesheetType" property to "CssStyleSheetType.External" to
        // accompany a saved HTML document with an external CSS stylesheet file.
        options->set_CssStyleSheetType(CssStyleSheetType::External);

        // Below are two ways of specifying directories and filenames for output CSS stylesheets.
        // 1 -  Use the "CssStyleSheetFileName" property to assign a filename to our stylesheet:
        options->set_CssStyleSheetFileName(ArtifactsDir + u"SavingCallback.ExternalCssFilenames.css");

        // 2 -  Use a custom callback to name our stylesheet:
        options->set_CssSavingCallback(
            MakeObject<ExSavingCallback::CustomCssSavingCallback>(ArtifactsDir + u"SavingCallback.ExternalCssFilenames.css", true, false));

        doc->Save(ArtifactsDir + u"SavingCallback.ExternalCssFilenames.html", options);
    }

    /// <summary>
    /// Sets a custom filename, along with other parameters for an external CSS stylesheet.
    /// </summary>
    class CustomCssSavingCallback : public ICssSavingCallback
    {
    public:
        CustomCssSavingCallback(String cssDocFilename, bool isExportNeeded, bool keepCssStreamOpen) : mIsExportNeeded(false), mKeepCssStreamOpen(false)
        {
            mCssTextFileName = cssDocFilename;
            mIsExportNeeded = isExportNeeded;
            mKeepCssStreamOpen = keepCssStreamOpen;
        }

        void CssSaving(SharedPtr<CssSavingArgs> args) override
        {
            // We can access the entire source document via the "Document" property.
            ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));

            args->set_CssStream(MakeObject<System::IO::FileStream>(mCssTextFileName, System::IO::FileMode::Create));
            args->set_IsExportNeeded(mIsExportNeeded);
            args->set_KeepCssStreamOpen(mKeepCssStreamOpen);

            ASSERT_TRUE(args->get_CssStream()->get_CanWrite());
        }

    private:
        String mCssTextFileName;
        bool mIsExportNeeded;
        bool mKeepCssStreamOpen;
    };
    //ExEnd
};

} // namespace ApiExamples
