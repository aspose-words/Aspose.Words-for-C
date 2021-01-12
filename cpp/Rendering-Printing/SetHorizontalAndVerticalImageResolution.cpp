#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageRange.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageSet.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void SetHorizontalAndVerticalImageResolution()
{
    std::cout << "SetHorizontalAndVerticalImageResolution example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();
    // ExStart:SetHorizontalAndVerticalImageResolution
    // Load the documents 
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    //Renders a page of a Word document into a PNG image at a specific horizontal and vertical resolution.
    System::SharedPtr<ImageSaveOptions> options = System::MakeObject<ImageSaveOptions>(SaveFormat::Png);
    options->set_HorizontalResolution(300.0f);
    options->set_VerticalResolution(300.0f);

    auto pageRange = System::MakeObject<PageRange>(0, 1);
    options->set_PageSet(System::MakeObject<PageSet>(System::MakeArray<System::SharedPtr<PageRange>>({ pageRange })));

    doc->Save(outputDataDir + u"SetHorizontalAndVerticalImageResolution.png", options);
    // ExEnd:SetHorizontalAndVerticalImageResolution
    std::cout << "SetHorizontalAndVerticalImageResolution example finished." << std::endl << std::endl;
}