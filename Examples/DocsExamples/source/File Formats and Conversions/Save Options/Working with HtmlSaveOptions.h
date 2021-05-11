#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Saving/HtmlMetafileFormat.h>
#include <Aspose.Words.Cpp/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/io/directory.h>
#include <system/io/path.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithHtmlSaveOptions : public DocsExamplesBase
{
public:
    void ExportRoundtripInformation()
    {
        //ExStart:ExportRoundtripInformation
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_ExportRoundtripInformation(true);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ExportRoundtripInformation.html", saveOptions);
        //ExEnd:ExportRoundtripInformation
    }

    void ExportFontsAsBase64()
    {
        //ExStart:ExportFontsAsBase64
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_ExportFontsAsBase64(true);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ExportFontsAsBase64.html", saveOptions);
        //ExEnd:ExportFontsAsBase64
    }

    void ExportResources()
    {
        //ExStart:ExportResources
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_ResourceFolder(ArtifactsDir + u"Resources");
        saveOptions->set_ResourceFolderAlias(u"http://example.com/resources");

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ExportResources.html", saveOptions);
        //ExEnd:ExportResources
    }

    void ConvertMetafilesToEmfOrWmf()
    {
        //ExStart:ConvertMetafilesToEmfOrWmf
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Here is an image as is: ");
        builder->InsertHtml(
            u"<img src=\"data:image/png;base64,\r\n                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP\r\n                    "
            u"C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA\r\n                    "
            u"AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J\r\n                    "
            u"REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq\r\n                    "
            u"ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0\r\n                    vr4MkhoXe0rZigAAAABJRU5ErkJggg==\" alt=\"Red dot\" />");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_MetafileFormat(HtmlMetafileFormat::EmfOrWmf);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ConvertMetafilesToEmfOrWmf.html", saveOptions);
        //ExEnd:ConvertMetafilesToEmfOrWmf
    }

    void ConvertMetafilesToSvg()
    {
        //ExStart:ConvertMetafilesToSvg
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Here is an SVG image: ");
        builder->InsertHtml(u"<svg height='210' width='500'>\r\n                <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                    "
                            u"style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n            </svg> ");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_MetafileFormat(HtmlMetafileFormat::Svg);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ConvertMetafilesToSvg.html", saveOptions);
        //ExEnd:ConvertMetafilesToSvg
    }

    void AddCssClassNamePrefix()
    {
        //ExStart:AddCssClassNamePrefix
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_CssClassNamePrefix(u"pfx_");

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.AddCssClassNamePrefix.html", saveOptions);
        //ExEnd:AddCssClassNamePrefix
    }

    void ExportCidUrlsForMhtmlResources()
    {
        //ExStart:ExportCidUrlsForMhtmlResources
        auto doc = MakeObject<Document>(MyDir + u"Content-ID.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Mhtml);
        saveOptions->set_PrettyFormat(true);
        saveOptions->set_ExportCidUrlsForMhtmlResources(true);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ExportCidUrlsForMhtmlResources.mhtml", saveOptions);
        //ExEnd:ExportCidUrlsForMhtmlResources
    }

    void ResolveFontNames()
    {
        //ExStart:ResolveFontNames
        auto doc = MakeObject<Document>(MyDir + u"Missing font.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        saveOptions->set_PrettyFormat(true);
        saveOptions->set_ResolveFontNames(true);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ResolveFontNames.html", saveOptions);
        //ExEnd:ResolveFontNames
    }

    void ExportTextInputFormFieldAsText()
    {
        //ExStart:ExportTextInputFormFieldAsText
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        String imagesDir = System::IO::Path::Combine(ArtifactsDir, u"Images");

        // The folder specified needs to exist and should be empty.
        if (System::IO::Directory::Exists(imagesDir))
        {
            System::IO::Directory::Delete(imagesDir, true);
        }

        System::IO::Directory::CreateDirectory_(imagesDir);

        // Set an option to export form fields as plain text, not as HTML input elements.
        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        saveOptions->set_ExportTextInputFormFieldAsText(true);
        saveOptions->set_ImagesFolder(imagesDir);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlSaveOptions.ExportTextInputFormFieldAsText.html", saveOptions);
        //ExEnd:ExportTextInputFormFieldAsText
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
