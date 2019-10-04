#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

void ExtractImagesToFiles()
{
    std::cout << "ExtractImagesToFiles example started." << std::endl;
    // ExStart:ExtractImagesToFiles
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithImages();
    System::String outputDataDir = GetOutputDataDir_WorkingWithImages();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Image.SampleImages.doc");

    System::SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
    int32_t imageIndex = 0;

    for (System::SharedPtr<Shape> shape : System::IterateOver<System::SharedPtr<Shape>>(shapes))
    {
        if (shape->get_HasImage())
        {
            System::String imageFileName = System::String::Format(u"Image.ExportImages.{0}.{1}", imageIndex, FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));
            System::String imagePath = outputDataDir + imageFileName;
            shape->get_ImageData()->Save(imagePath);
            std::cout << "Image saved at " << imagePath.ToUtf8String() << std::endl;
            imageIndex++;
        }
    }
    // ExEnd:ExtractImagesToFiles
    std::cout << "All images extracted from document." << std::endl;
    std::cout << "ExtractImagesToFiles example finished." << std::endl << std::endl;
}
