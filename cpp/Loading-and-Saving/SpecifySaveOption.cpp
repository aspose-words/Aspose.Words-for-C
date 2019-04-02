#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/HtmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void SpecifySaveOption()
{
    std::cout << "SpecifySaveOption example started." << std::endl;
    // ExStart:SpecifySaveOption
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile RenderShape.docx");

    // This is the directory we want the exported images to be saved to.
    System::String imagesDir = System::IO::Path::Combine(dataDir, u"SpecifySaveOption.Images");

    // The folder specified needs to exist and should be empty.
    if (System::IO::Directory::Exists(imagesDir))
    {
        System::IO::Directory::Delete(imagesDir, true);
    }
    System::IO::Directory::CreateDirectory_(imagesDir);

    // Set an option to export form fields as plain text, not as HTML input elements.
    System::SharedPtr<HtmlSaveOptions> options = System::MakeObject<HtmlSaveOptions>(SaveFormat::Html);
    options->set_ExportTextInputFormFieldAsText(true);
    options->set_ImagesFolder(imagesDir);

    System::String outputPath = dataDir + GetOutputFilePath(u"SpecifySaveOption.html");
    doc->Save(outputPath, options);
    // ExEnd:SpecifySaveOption

    std::cout << "Save option specified successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SpecifySaveOption example finished." << std::endl << std::endl;
}