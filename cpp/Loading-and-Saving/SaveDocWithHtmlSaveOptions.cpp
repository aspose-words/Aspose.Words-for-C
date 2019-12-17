#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void SaveHtmlWithMetafileFormat(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SaveHtmlWithMetafileFormat
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.docx");
        System::SharedPtr<HtmlSaveOptions> options = System::MakeObject<HtmlSaveOptions>();
        options->set_MetafileFormat(HtmlMetafileFormat::EmfOrWmf);

        System::String outputPath = outputDataDir + u"SaveDocWithHtmlSaveOptions.SaveHtmlWithMetafileFormat.html";
        doc->Save(outputPath, options);
        // ExEnd:SaveHtmlWithMetafileFormat
        std::cout << "Document saved with Metafile format." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ImportExportSVGinHTML(System::String const &outputDataDir)
    {
        // ExStart:ImportExportSVGinHTML
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Here is an SVG image: ");
        builder->InsertHtml(u"<svg height='210' width='500'><polygon points='100,10 40,198 190,78 10,78 160,198' style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' /></svg> ");

        System::SharedPtr<HtmlSaveOptions> options = System::MakeObject<HtmlSaveOptions>();
        options->set_MetafileFormat(HtmlMetafileFormat::Svg);

        System::String outputPath = outputDataDir + u"SaveDocWithHtmlSaveOptions.ImportExportSVGinHTML.html";
        doc->Save(outputPath, options);
        // ExEnd:ImportExportSVGinHTML
        std::cout << "Document saved with SVG Metafile format." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetCssClassNamePrefix(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetCssClassNamePrefix
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.docx");

        System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_CssClassNamePrefix(u"pfx_");

        System::String outputPath = outputDataDir + u"SaveDocWithHtmlSaveOptions.SetCssClassNamePrefix.html";
        doc->Save(outputPath, saveOptions);
        // ExEnd:SetCssClassNamePrefix
        std::cout << "Document saved with CSS prefix pfx_." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetExportCidUrlsForMhtmlResources(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetExportCidUrlsForMhtmlResources
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"CidUrls.docx");

        System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>(SaveFormat::Mhtml);
        saveOptions->set_PrettyFormat(true);
        saveOptions->set_ExportCidUrlsForMhtmlResources(true);
        saveOptions->set_SaveFormat(SaveFormat::Mhtml);

        System::String outputPath = outputDataDir + u"SaveDocWithHtmlSaveOptions.SetExportCidUrlsForMhtmlResources.mhtml";
        doc->Save(outputPath, saveOptions);
        // ExEnd:SetExportCidUrlsForMhtmlResources
        std::cout << "Document has saved with Content - Id URL scheme." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetResolveFontNames(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetResolveFontNames
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (docx).docx");

        System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        saveOptions->set_PrettyFormat(true);
        saveOptions->set_ResolveFontNames(true);

        System::String outputPath = outputDataDir + u"SaveDocWithHtmlSaveOptions.SetResolveFontNames.html";
        doc->Save(outputPath, saveOptions);
        // ExEnd:SetResolveFontNames
        std::cout << "FontSettings is used to resolve font family name." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

}

void SaveDocWithHtmlSaveOptions()
{
    std::cout << "SaveDocWithHtmlSaveOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    SaveHtmlWithMetafileFormat(inputDataDir, outputDataDir);
    ImportExportSVGinHTML(outputDataDir);
    SetCssClassNamePrefix(inputDataDir, outputDataDir);
    SetExportCidUrlsForMhtmlResources(inputDataDir, outputDataDir);
    SetResolveFontNames(inputDataDir, outputDataDir);
    std::cout << "SaveDocWithHtmlSaveOptions example finished." << std::endl << std::endl;
}