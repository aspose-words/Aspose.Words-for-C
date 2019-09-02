#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/HtmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void ExportFontsAsBase64()
{
    std::cout << "ExportFontsAsBase64 example started." << std::endl;
    // ExStart:ExportFontsAsBase64
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_ExportFontsAsBase64(true);
    System::String outputPath = outputDataDir + u"ExportFontsAsBase64.html";
    doc->Save(outputPath, saveOptions);
    // ExEnd:ExportFontsAsBase64
    std::cout << "Save option specified successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExportFontsAsBase64 example finished." << std::endl << std::endl;
}