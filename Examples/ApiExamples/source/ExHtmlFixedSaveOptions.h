#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
#include <system/io/path.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/match_collection.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

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
        //ExSummary:Shows how to set which encoding to use while exporting a document to HTML.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello World!");

        // The default encoding is UTF-8. If we want to represent our document using a different encoding,
        // we can use a SaveOptions object to set a specific encoding.
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

    void ExportEmbeddedCss(bool exportEmbeddedCss)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
        //ExSummary:Shows how to determine where to store CSS stylesheets when exporting a document to Html.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // When we export a document to html, Aspose.Words will also create a CSS stylesheet to format the document with.
        // Setting the "ExportEmbeddedCss" flag to "true" save the CSS stylesheet to a .css file,
        // and link to the file from the html document using a <link> element.
        // Setting the flag to "false" will embed the CSS stylesheet within the Html document,
        // which will create only one file instead of two.
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedCss(exportEmbeddedCss);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCss.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCss.html");

        if (exportEmbeddedCss)
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<style type=\"text/css\">")->get_Success());
            ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css"));
        }
        else
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents,
                            u"<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions[.]ExportEmbeddedCss/styles[.]css\" media=\"all\" />")
                            ->get_Success());
            ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css"));
        }
        //ExEnd
    }

    void ExportEmbeddedFonts(bool exportEmbeddedFonts)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
        //ExSummary:Shows how to determine where to store embedded fonts when exporting a document to Html.
        auto doc = MakeObject<Document>(MyDir + u"Embedded font.docx");

        // When we export a document with embedded fonts to .html,
        // Aspose.Words can place the fonts in two possible locations.
        // Setting the "ExportEmbeddedFonts" flag to "true" will store the raw data for embedded fonts within the CSS stylesheet,
        // in the "url" property of the "@font-face" rule. This may create a huge CSS stylesheet file
        // and reduce the number of external files that this HTML conversion will create.
        // Setting this flag to "false" will create a file for each font.
        // The CSS stylesheet will link to each font file using the "url" property of the "@font-face" rule.
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedFonts(exportEmbeddedFonts);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css");

        if (exportEmbeddedFonts)
        {
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents,
                            u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; }")
                            ->get_Success());
            ASSERT_EQ(0, System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts")
                             ->LINQ_Count([](String f) { return f.EndsWith(u".woff"); }));
        }
        else
        {
            ASSERT_TRUE(
                System::Text::RegularExpressions::Regex::Match(outDocContents, u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; "
                                                                               u"src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; }")
                    ->get_Success());
            ASSERT_EQ(2, System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts")
                             ->LINQ_Count([](String f) { return f.EndsWith(u".woff"); }));
        }
        //ExEnd
    }

    void ExportEmbeddedImages(bool exportImages)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
        //ExSummary:Shows how to determine where to store images when exporting a document to Html.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // When we export a document with embedded images to .html,
        // Aspose.Words can place the images in two possible locations.
        // Setting the "ExportEmbeddedImages" flag to "true" will store the raw data
        // for all images within the output HTML document, in the "src" attribute of <image> tags.
        // Setting this flag to "false" will create an image file in the local file system for every image,
        // and store all these files in a separate folder.
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedImages(exportImages);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages.html");

        if (exportImages)
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

    void ExportEmbeddedSvgs(bool exportSvgs)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
        //ExSummary:Shows how to determine where to store SVG objects when exporting a document to Html.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // When we export a document with SVG objects to .html,
        // Aspose.Words can place these objects in two possible locations.
        // Setting the "ExportEmbeddedSvg" flag to "true" will embed all SVG object raw data
        // within the output HTML, inside <image> tags.
        // Setting this flag to "false" will create a file in the local file system for each SVG object.
        // The HTML will link to each file using the "data" attribute of an <object> tag.
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportEmbeddedSvg(exportSvgs);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs.html");

        if (exportSvgs)
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

    void ExportFormFields(bool exportFormFields)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportFormFields
        //ExSummary:Shows how to export form fields to Html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertCheckBox(u"CheckBox", false, 15);

        // When we export a document with form fields to .html,
        // there are two ways in which Aspose.Words can export form fields.
        // Setting the "ExportFormFields" flag to "true" will export them as interactive objects.
        // Setting this flag to "false" will display form fields as plain text.
        // This will freeze them at their current value, and prevent the reader of our HTML document
        // from being able to interact with them.
        auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
        htmlFixedSaveOptions->set_ExportFormFields(exportFormFields);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportFormFields.html");

        if (exportFormFields)
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
        //ExSummary:Shows how to place CSS into a separate file and add a prefix to all of its CSS class names.
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

        outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.AddCssClassNamesPrefix/styles.css");

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents,
                                                                   String(u".myprefixdiv { position:absolute; } ") +
                                                                       u".myprefixspan { position:absolute; white-space:pre; color:#000000; font-size:12pt; }")
                        ->get_Success());
        //ExEnd
    }

    void HorizontalAlignment(HtmlFixedPageHorizontalAlignment pageHorizontalAlignment)
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
        //ExFor:HtmlFixedPageHorizontalAlignment
        //ExSummary:Shows how to set the horizontal alignment of pages when saving a document to HTML.
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
        //ExSummary:Shows how to adjust page margins when saving a document to HTML.
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
        ASSERT_THROW(saveOptions->set_PageMargins(-1), System::ArgumentException);
    }

    void OptimizeGraphicsOutput(bool optimizeOutput)
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExFor:HtmlFixedSaveOptions.OptimizeOutput
        //ExSummary:Shows how to simplify a document when saving it to HTML by removing various redundant objects.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_OptimizeOutput(optimizeOutput);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.html", saveOptions);

        // The size of the optimized version of the document is almost a third of the size of the unoptimized document.
        if (optimizeOutput)
        {
            ASSERT_NEAR(58000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.html")->get_Length(), 200);
        }
        else
        {
            ASSERT_NEAR(161100, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.html")->get_Length(), 200);
        }
        //ExEnd
    }

    void UsingMachineFonts(bool useTargetMachineFonts)
    {
        //ExStart
        //ExFor:ExportFontFormat
        //ExFor:HtmlFixedSaveOptions.FontFormat
        //ExFor:HtmlFixedSaveOptions.UseTargetMachineFonts
        //ExSummary:Shows how use fonts only from the target machine when saving a document to HTML.
        auto doc = MakeObject<Document>(MyDir + u"Bullet points with alternative font.docx");

        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_ExportEmbeddedCss(true);
        saveOptions->set_UseTargetMachineFonts(useTargetMachineFonts);
        saveOptions->set_FontFormat(ExportFontFormat::Ttf);
        saveOptions->set_ExportEmbeddedFonts(false);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.UsingMachineFonts.html");

        if (useTargetMachineFonts)
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
        //ExEnd
    }

    //ExStart
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs
    //ExFor:ResourceSavingArgs.Document
    //ExFor:ResourceSavingArgs.ResourceFileName
    //ExFor:ResourceSavingArgs.ResourceFileUri
    //ExSummary:Shows how to use a callback to track external resources created while converting a document to HTML.
    void ResourceSavingCallback()
    {
        auto doc = MakeObject<Document>(MyDir + u"Bullet points with alternative font.docx");

        auto callback = MakeObject<ExHtmlFixedSaveOptions::FontSavingCallback>();

        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_ResourceSavingCallback(callback);

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

        std::cout << callback->GetText() << std::endl;
        TestResourceSavingCallback(callback);
        //ExSkip
    }

    class FontSavingCallback : public IResourceSavingCallback
    {
    public:
        /// <summary>
        /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
        /// </summary>
        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
            mText->AppendLine(String::Format(u"Original document URI:\t{0}", args->get_Document()->get_OriginalFileName()));
            mText->AppendLine(String::Format(u"Resource being saved:\t{0}", args->get_ResourceFileName()));
            mText->AppendLine(String::Format(u"Full uri after saving:\t{0}\n", args->get_ResourceFileUri()));
        }

        String GetText()
        {
            return mText->ToString();
        }

        FontSavingCallback() : mText(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        SharedPtr<System::Text::StringBuilder> mText;
    };
    //ExEnd

    void TestResourceSavingCallback(SharedPtr<ExHtmlFixedSaveOptions::FontSavingCallback> callback)
    {
        ASSERT_TRUE(callback->GetText().Contains(u"font001.woff"));
        ASSERT_TRUE(callback->GetText().Contains(u"styles.css"));
    }

    //ExStart
    //ExFor:HtmlFixedSaveOptions
    //ExFor:HtmlFixedSaveOptions.ResourceSavingCallback
    //ExFor:HtmlFixedSaveOptions.ResourcesFolder
    //ExFor:HtmlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:HtmlFixedSaveOptions.SaveFormat
    //ExFor:HtmlFixedSaveOptions.ShowPageBorder
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs.KeepResourceStreamOpen
    //ExFor:ResourceSavingArgs.ResourceStream
    //ExSummary:Shows how to use a callback to print the URIs of external resources created while converting a document to HTML.
    void HtmlFixedResourceFolder()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto callback = MakeObject<ExHtmlFixedSaveOptions::ResourceUriPrinter>();

        auto options = MakeObject<HtmlFixedSaveOptions>();
        options->set_SaveFormat(SaveFormat::HtmlFixed);
        options->set_ExportEmbeddedImages(false);
        options->set_ResourcesFolder(ArtifactsDir + u"HtmlFixedResourceFolder");
        options->set_ResourcesFolderAlias(ArtifactsDir + u"HtmlFixedResourceFolderAlias");
        options->set_ShowPageBorder(false);
        options->set_ResourceSavingCallback(callback);

        // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder.
        // We must ensure the folder exists before the streams can put their resources into it.
        System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

        doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

        std::cout << callback->GetText() << std::endl;

        ArrayPtr<String> resourceFiles = System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedResourceFolderAlias");

        ASSERT_FALSE(System::IO::Directory::Exists(ArtifactsDir + u"HtmlFixedResourceFolder"));
        std::function<bool(String f)> isImageOrCss = [](String f)
        {
            return f.EndsWith(u".jpeg") || f.EndsWith(u".png") || f.EndsWith(u".css");
        };

        ASSERT_EQ(6, resourceFiles->LINQ_Count(isImageOrCss));
        TestHtmlFixedResourceFolder(callback);
        //ExSkip
    }

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed HTML.
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
    public:
        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
            // If we set a folder alias in the SaveOptions object, we will be able to print it from here.
            mText->AppendLine(String::Format(u"Resource #{0} \"{1}\"", ++mSavedResourceCount, args->get_ResourceFileName()));

            String extension = System::IO::Path::GetExtension(args->get_ResourceFileName());
            if (extension == u".ttf" || extension == u".woff")
            {
                // By default, 'ResourceFileUri' uses system folder for fonts.
                // To avoid problems in other platforms you must explicitly specify the path for the fonts.
                args->set_ResourceFileUri(ArtifactsDir + System::IO::Path::DirectorySeparatorChar + args->get_ResourceFileName());
            }

            mText->AppendLine(String(u"\t") + args->get_ResourceFileUri());

            // If we have specified a folder in the "ResourcesFolderAlias" property,
            // we will also need to redirect each stream to put its resource in that folder.
            args->set_ResourceStream(MakeObject<System::IO::FileStream>(args->get_ResourceFileUri(), System::IO::FileMode::Create));
            args->set_KeepResourceStreamOpen(false);
        }

        String GetText()
        {
            return mText->ToString();
        }

        ResourceUriPrinter() : mSavedResourceCount(0), mText(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        int mSavedResourceCount;
        SharedPtr<System::Text::StringBuilder> mText;
    };
    //ExEnd

    void TestHtmlFixedResourceFolder(SharedPtr<ExHtmlFixedSaveOptions::ResourceUriPrinter> callback)
    {
        ASSERT_EQ(16, System::Text::RegularExpressions::Regex::Matches(callback->GetText(), u"Resource #")->get_Count());
    }
};

} // namespace ApiExamples
