#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Saving/ExportFontFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedPageHorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/path.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExHtmlFixedSaveOptions : public ApiExampleBase
{
public:
    void UseEncoding()
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.Encoding
        //ExSummary:Shows how to set encoding while exporting to HTML.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello World!");

        // The default encoding is UTF-8
        // If we want to represent our document using a different encoding, we can set one explicitly using a SaveOptions object
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_Encoding(System::Text::Encoding::GetEncoding(u"ASCII"));

        ASSERT_EQ(u"US-ASCII", htmlFixedSaveOptions->get_Encoding()->get_EncodingName());

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.UseEncoding.html", htmlFixedSaveOptions);
        //ExEnd

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.UseEncoding.html"),
                                                                   u"content=\"text/html; charset=us-ascii\"")
                        ->get_Success());
    }

    void GetEncoding()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_Encoding(System::Text::Encoding::GetEncoding(u"utf-16"));

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.GetEncoding.html", htmlFixedSaveOptions);
    }

    void ExportEmbeddedCSS(bool doExportEmbeddedCss)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
        //ExSummary:Shows how to export embedded stylesheets into an HTML file.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedCss(doExportEmbeddedCss);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS.html");

        if (doExportEmbeddedCss)
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<style type=\"text/css\">")->get_Success());
            ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
        }
        else
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents,
                            u"<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions[.]ExportEmbeddedCSS/styles[.]css\" media=\"all\" />")
                            ->get_Success());
            ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
        }
        //ExEnd
    }

    void ExportEmbeddedFonts(bool doExportEmbeddedFonts)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
        //ExSummary:Shows how to export embedded fonts into an HTML file.
        auto doc = MakeObject<Document>(MyDir + u"Embedded font.docx");

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedFonts(doExportEmbeddedFonts);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css");

        if (doExportEmbeddedFonts)
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents,
                            u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; }")
                            ->get_Success());

            ASSERT_EQ(0, System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts")->LINQ_Count([](String f) { return f.EndsWith(u".woff"); }));
        }
        else
        {
            ASSERT_TRUE(
                System::Text::RegularExpressions::Regex::Match(outDocContents, u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; "
                                                                               u"src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; }")
                    ->get_Success());

            ASSERT_EQ(2, System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts")->LINQ_Count([](String f) { return f.EndsWith(u".woff"); }));
        }
        //ExEnd
    }

    void ExportEmbeddedImages(bool doExportImages)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
        //ExSummary:Shows how to export embedded images into an HTML file.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedImages(doExportImages);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages.html");

        if (doExportImages)
        {
            ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, u"<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" src=\".+\" />")
                            ->get_Success());
        }
        else
        {
            ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, String(u"<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" ") +
                                                u"src=\"HtmlFixedSaveOptions[.]ExportEmbeddedImages/image001[.]jpeg\" />")
                            ->get_Success());
        }
        //ExEnd
    }

    void ExportEmbeddedSvgs(bool doExportSvgs)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
        //ExSummary:Shows how to export embedded SVG objects into an HTML file.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedSvg(doExportSvgs);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs.html");

        if (doExportSvgs)
        {
            ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<image id=\"image004\" xlink:href=.+/>")->get_Success());
        }
        else
        {
            ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, u"<object type=\"image/svg[+]xml\" data=\"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001[.]svg\"></object>")
                            ->get_Success());
        }
        //ExEnd
    }

    void ExportFormFields(bool doExportFormFields)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportFormFields
        //ExSummary:Show how to exporting form fields from a document into HTML file.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertCheckBox(u"CheckBox", false, 15);

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportFormFields(doExportFormFields);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportFormFields.html");

        if (doExportFormFields)
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, String(u"<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>") +
                                                u"<input style=\"position:absolute; left:0pt; top:0pt;\" type=\"checkbox\" name=\"CheckBox\" />")
                            ->get_Success());
        }
        else
        {
            ASSERT_TRUE(
                System::Text::RegularExpressions::Regex::Match(
                    outDocContents, String(u"<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>") +
                                        u"<div class=\"awdiv\" style=\"left:0.8pt; top:0.8pt; width:14.25pt; height:14.25pt; border:solid 0.75pt #000000;\"")
                    ->get_Success());
        }
        //ExEnd
    }

    void AddCssClassNamesPrefix()
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.CssClassNamesPrefix
        //ExFor:HtmlFixedSaveOptions.SaveFontFaceCssSeparately
        //ExSummary:Shows how to add prefix to all class names in css file.
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_CssClassNamesPrefix(u"myprefix");
        htmlFixedSaveOptions->set_SaveFontFaceCssSeparately(true);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.AddCssClassNamesPrefix.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.AddCssClassNamesPrefix.html");

        ASSERT_TRUE(
            System::Text::RegularExpressions::Regex::Match(
                outDocContents, String(u"<div class=\"myprefixdiv myprefixpage\" style=\"width:595[.]3pt; height:841[.]9pt;\">") +
                                    u"<div class=\"myprefixdiv\" style=\"left:85[.]05pt; top:36pt; clip:rect[(]0pt,510[.]25pt,74[.]95pt,-85.05pt[)];\">" +
                                    u"<span class=\"myprefixspan myprefixtext001\" style=\"font-size:11pt; left:294[.]73pt; top:0[.]36pt;\">")
                ->get_Success());
        //ExEnd
    }

    void HorizontalAlignment(HtmlFixedPageHorizontalAlignment pageHorizontalAlignment)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
        //ExFor:HtmlFixedPageHorizontalAlignment
        //ExSummary:Shows how to set the horizontal alignment of pages in HTML file.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_PageHorizontalAlignment(pageHorizontalAlignment);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.HorizontalAlignment.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.HorizontalAlignment/styles.css");

        switch (pageHorizontalAlignment)
        {
        case HtmlFixedPageHorizontalAlignment::Center:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; }")
                            ->get_Success());
            break;

        case HtmlFixedPageHorizontalAlignment::Left:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; }")
                            ->get_Success());
            break;

        case HtmlFixedPageHorizontalAlignment::Right:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; }")
                            ->get_Success());
            break;
        }
        //ExEnd
    }

    void PageMargins()
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.PageMargins
        //ExSummary:Shows how to set the margins around pages in HTML file.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_PageMargins(15);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.PageMargins.html", saveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.PageMargins/styles.css");

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                        outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }")
                        ->get_Success());
        //ExEnd
    }

    void PageMarginsException()
    {
        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<void()> _local_func_2 = [&saveOptions]()
        {
            saveOptions->set_PageMargins(-1);
        };

        ASSERT_THROW(_local_func_2(), System::ArgumentException);
    }

    void OptimizeGraphicsOutput()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExFor:HtmlFixedSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to html.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_OptimizeOutput(false);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html", saveOptions);

        saveOptions->set_OptimizeOutput(true);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html", saveOptions);

        ASSERT_TRUE(MakeObject<System::IO::FileInfo>(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html")->get_Length() >
                    MakeObject<System::IO::FileInfo>(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html")->get_Length());
        //ExEnd
    }

    //ExStart
    //ExFor:ExportFontFormat
    //ExFor:HtmlFixedSaveOptions.FontFormat
    //ExFor:HtmlFixedSaveOptions.UseTargetMachineFonts
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs
    //ExFor:ResourceSavingArgs.Document
    //ExFor:ResourceSavingArgs.KeepResourceStreamOpen
    //ExFor:ResourceSavingArgs.ResourceFileName
    //ExFor:ResourceSavingArgs.ResourceFileUri
    //ExFor:ResourceSavingArgs.ResourceStream
    //ExSummary:Shows how use target machine fonts to display the document.
    void UsingMachineFonts()
    {
        auto doc = MakeObject<Document>(MyDir + u"Bullet points with alternative font.docx");

        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_ExportEmbeddedCss(true);
        saveOptions->set_UseTargetMachineFonts(true);
        saveOptions->set_FontFormat(ExportFontFormat::Ttf);
        saveOptions->set_ExportEmbeddedFonts(false);
        saveOptions->set_ResourceSavingCallback(MakeObject<ExHtmlFixedSaveOptions::ResourceSavingCallback>());

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.UsingMachineFonts.html");

        if (saveOptions->get_UseTargetMachineFonts())
        {
            ASSERT_FALSE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"@font-face")->get_Success());
        }
        else
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, String(u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], ") +
                                                u"url[(]'HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf'[)] format[(]'truetype'[)]; }")
                            ->get_Success());
        }
    }

private:
    class ResourceSavingCallback : public IResourceSavingCallback
    {
    public:
        /// <summary>
        /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
        /// </summary>
        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
            std::cout << "Original document URI:\t" << args->get_Document()->get_OriginalFileName() << std::endl;
            std::cout << "Resource being saved:\t" << args->get_ResourceFileName() << std::endl;
            std::cout << "Full uri after saving:\t" << args->get_ResourceFileUri() << std::endl;

            args->set_ResourceStream(MakeObject<System::IO::MemoryStream>());
            args->set_KeepResourceStreamOpen(true);

            String extension = System::IO::Path::GetExtension(args->get_ResourceFileName());
            const String& switch_value_0 = extension;
            if (switch_value_0 == u".ttf" || switch_value_0 == u".woff")
            {

                {
                    FAIL() << "'ResourceSavingCallback' is not fired for fonts when 'UseTargetMachineFonts' is true";
                    goto switch_break_0;
                }
            }
        switch_break_0:;
        }
    };
    //ExEnd

public:
    //ExStart
    //ExFor:HtmlFixedSaveOptions
    //ExFor:HtmlFixedSaveOptions.ResourceSavingCallback
    //ExFor:HtmlFixedSaveOptions.ResourcesFolder
    //ExFor:HtmlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:HtmlFixedSaveOptions.SaveFormat
    //ExFor:HtmlFixedSaveOptions.ShowPageBorder
    //ExSummary:Shows how to print the URIs of linked resources created during conversion of a document to fixed-form .html.
    void HtmlFixedResourceFolder()
    {
        // Open a document which contains images
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<HtmlFixedSaveOptions>();
        options->set_SaveFormat(SaveFormat::HtmlFixed);
        options->set_ExportEmbeddedImages(false);
        options->set_ResourcesFolder(ArtifactsDir + u"HtmlFixedResourceFolder");
        options->set_ResourcesFolderAlias(ArtifactsDir + u"HtmlFixedResourceFolderAlias");
        options->set_ShowPageBorder(false);
        options->set_ResourceSavingCallback(MakeObject<ExHtmlFixedSaveOptions::ResourceUriPrinter>());

        // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
        // We must ensure the folder exists before the streams can put their resources into it
        System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

        ArrayPtr<String> resourceFiles = System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedResourceFolderAlias");

        ASSERT_FALSE(System::IO::Directory::Exists(ArtifactsDir + u"HtmlFixedResourceFolder"));

        std::function<bool(String f)> isImageOrCss = [](String f)
        {
            return f.EndsWith(u".jpeg") || f.EndsWith(u".png") || f.EndsWith(u".css");
        };

        ASSERT_EQ(6, resourceFiles->LINQ_Count(isImageOrCss));
    }

private:
    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed .Html
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
    public:
        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
            // If we set a folder alias in the SaveOptions object, it will be printed here
            std::cout << "Resource #" << ++mSavedResourceCount << " \"" << args->get_ResourceFileName() << "\"" << std::endl;

            String extension = System::IO::Path::GetExtension(args->get_ResourceFileName());
            const String& switch_value_1 = extension;
            if (switch_value_1 == u".ttf" || switch_value_1 == u".woff")
            {

                {
                    // By default, 'ResourceFileUri' used system folder for fonts
                    // To avoid problems across platforms you must explicitly specify the path for the fonts
                    args->set_ResourceFileUri(ArtifactsDir + System::IO::Path::DirectorySeparatorChar + args->get_ResourceFileName());
                    goto switch_break_1;
                }
            }
        switch_break_1:;
            std::cout << (String(u"\t") + args->get_ResourceFileUri()) << std::endl;

            // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
            args->set_ResourceStream(MakeObject<System::IO::FileStream>(args->get_ResourceFileUri(), System::IO::FileMode::Create));
            args->set_KeepResourceStreamOpen(false);
        }

        ResourceUriPrinter() : mSavedResourceCount(0)
        {
        }

    private:
        int mSavedResourceCount;
    };
    //ExEnd
};

} // namespace ApiExamples
