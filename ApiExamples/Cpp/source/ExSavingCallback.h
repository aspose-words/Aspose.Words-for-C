#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Saving/CssSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentPartSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/ICssSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IDocumentPartSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IPageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/XamlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExSavingCallback : public ApiExampleBase
{
public:
    void CheckThatAllMethodsArePresent()
    {
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

        auto imageSaveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        imageSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

        auto psSaveOptions = MakeObject<PsSaveOptions>();
        psSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

        auto svgSaveOptions = MakeObject<SvgSaveOptions>();
        svgSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

        auto xamlFixedSaveOptions = MakeObject<XamlFixedSaveOptions>();
        xamlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

        auto xpsSaveOptions = MakeObject<XpsSaveOptions>();
        xpsSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());
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
    //ExSummary:Shows how separate pages are saved when a document is exported to fixed page format.
    void PageFileName()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_PageIndex(0);
        htmlFixedSaveOptions->set_PageCount(doc->get_PageCount());
        htmlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

        doc->Save(String::Format(u"{0}SavingCallback.PageFileName.html", ArtifactsDir), htmlFixedSaveOptions);

        ArrayPtr<String> filePaths = System::IO::Directory::GetFiles(ArtifactsDir)
            ->LINQ_Where([](String s) { return s.StartsWith(ArtifactsDir + u"SavingCallback.PageFileName.Page_"); })
            ->LINQ_OrderBy<String>([](String s) { return s; })
            ->LINQ_ToArray();

        for (int i = 0; i < doc->get_PageCount(); i++)
        {
            String file = String::Format(u"{0}SavingCallback.PageFileName.Page_{1}.html", ArtifactsDir, i);
            ASSERT_EQ(file, filePaths[i]);
            //ExSkip
        }
    }

private:
    /// <summary>
    /// Custom PageFileName is specified.
    /// </summary>
    class CustomPageFileNamePageSavingCallback : public IPageSavingCallback
    {
    public:
        void PageSaving(SharedPtr<PageSavingArgs> args) override
        {
            String outFileName = String::Format(u"{0}SavingCallback.PageFileName.Page_{1}.html", ArtifactsDir, args->get_PageIndex());

            // Specify name of the output file for the current page either in this
            args->set_PageFileName(outFileName);

            // ..or by setting up a custom stream
            args->set_PageStream(MakeObject<System::IO::FileStream>(outFileName, System::IO::FileMode::Create));
            ASSERT_FALSE(args->get_KeepPageStreamOpen());
        }
    };
    //ExEnd

public:
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
    //ExSummary:Shows how split a document into parts and save them.
    void DocumentParts()
    {
        // Open a document to be converted to html
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        String outFileName = u"SavingCallback.DocumentParts.Rendering.html";

        // We can use an appropriate SaveOptions subclass to customize the conversion process
        auto options = MakeObject<HtmlSaveOptions>();

        // We can use it to split a document into smaller parts, in this instance split by section breaks
        // Each part will be saved into a separate file, creating many files during the conversion process instead of just one
        options->set_DocumentSplitCriteria(DocumentSplitCriteria::SectionBreak);

        // We can set a callback to name each document part file ourselves
        options->set_DocumentPartSavingCallback(
            MakeObject<ExSavingCallback::SavedDocumentPartRename>(outFileName, options->get_DocumentSplitCriteria()));

        // If we convert a document that contains images into html, we will end up with one html file which links to several images
        // Each image will be in the form of a file in the local file system
        // There is also a callback that can customize the name and file system location of each image
        options->set_ImageSavingCallback(MakeObject<ExSavingCallback::SavedImageRename>(outFileName));

        // The DocumentPartSaving() and ImageSaving() methods of our callbacks will be run at this time
        doc->Save(ArtifactsDir + outFileName, options);
    }

private:
    /// <summary>
    /// Renames saved document parts that are produced when an HTML document is saved while being split according to a DocumentSplitCriteria.
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
            ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));

            String partType = u"";

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

            // We can designate the filename and location of each output file either by filename
            args->set_DocumentPartFileName(partFileName);

            // Or we can make a new stream and choose the location of the file at construction
            args->set_DocumentPartStream(MakeObject<System::IO::FileStream>(ArtifactsDir + partFileName, System::IO::FileMode::Create));
            ASSERT_TRUE(args->get_DocumentPartStream()->get_CanWrite());
            ASSERT_FALSE(args->get_KeepDocumentPartStreamOpen());
        }

    private:
        int mCount;
        String mOutFileName;
        DocumentSplitCriteria mDocumentSplitCriteria;
    };

public:
    /// <summary>
    /// Renames saved images that are produced when an HTML document is saved.
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
            // Same filename and stream functions as above in IDocumentPartSavingCallback apply here
            String imageFileName =
                String::Format(u"{0} shape {1}, of type {2}{3}", mOutFileName, ++mCount, args->get_CurrentShape()->get_ShapeType(),
                                       System::IO::Path::GetExtension(args->get_ImageFileName()));

            args->set_ImageFileName(imageFileName);

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
    //ExSummary:Shows how to work with CSS stylesheets that may be created along with Html documents.
    void CssSavingCallback()
    {
        // Open a document to be converted to html
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // If our output document will produce a CSS stylesheet, we can use an HtmlSaveOptions to control where it is saved
        auto options = MakeObject<HtmlSaveOptions>();

        // By default, a CSS stylesheet is stored inside its HTML document, but we can have it saved to a separate file
        options->set_CssStyleSheetType(CssStyleSheetType::External);

        // We can designate a filename for our stylesheet like this
        options->set_CssStyleSheetFileName(ArtifactsDir + u"SavingCallback.CssSavingCallback.css");

        // A custom ICssSavingCallback implementation can also control where that stylesheet will be saved and linked to by the Html document
        // This callback will override the filename we specified above in options.CssStyleSheetFileName,
        // but will give us more control over the saving process
        options->set_CssSavingCallback(
            MakeObject<ExSavingCallback::CustomCssSavingCallback>(ArtifactsDir + u"SavingCallback.CssSavingCallback.css", true, false));

        // The CssSaving() method of our callback will be called at this stage
        doc->Save(ArtifactsDir + u"SavingCallback.CssSavingCallback.html", options);
    }

private:
    /// <summary>
    /// Designates a filename and other parameters for the saving of a CSS stylesheet
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
            // Set up the stream that will create the CSS document
            args->set_CssStream(MakeObject<System::IO::FileStream>(mCssTextFileName, System::IO::FileMode::Create));
            ASSERT_TRUE(args->get_CssStream()->get_CanWrite());
            args->set_IsExportNeeded(mIsExportNeeded);
            args->set_KeepCssStreamOpen(mKeepCssStreamOpen);

            // We can also access the original document here like this
            ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));
        }

    private:
        String mCssTextFileName;
        bool mIsExportNeeded;
        bool mKeepCssStreamOpen;
    };
    //ExEnd
};

} // namespace ApiExamples
