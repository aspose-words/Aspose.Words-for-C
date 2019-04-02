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
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_ExportFontsAsBase64(true);
    System::String outputPath = dataDir + GetOutputFilePath(u"ExportFontsAsBase64.html");
    doc->Save(outputPath, saveOptions);
    // ExEnd:ExportFontsAsBase64
    std::cout << "Save option specified successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExportFontsAsBase64 example finished." << std::endl << std::endl;
}