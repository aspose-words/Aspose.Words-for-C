#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void ExportResourcesUsingHtmlSaveOptions()
{
    std::cout << "ExportResourcesUsingHtmlSaveOptions example started." << std::endl;
    // ExStart:ExportResourcesUsingHtmlSaveOptions
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    System::String fileName = u"Document.doc";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + fileName);
    System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();
    saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_ResourceFolder(outputDataDir + u"\\ExportResourcesUsingHtmlSaveOptions.Resources");
    saveOptions->set_ResourceFolderAlias(u"http://example.com/resources");
    System::String outputPath = outputDataDir + u"ExportResourcesUsingHtmlSaveOptions.html";
    doc->Save(outputPath, saveOptions);
    // ExEnd:ExportResourcesUsingHtmlSaveOptions
    std::cout << "Save option specified successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExportResourcesUsingHtmlSaveOptions example finished." << std::endl << std::endl;
}