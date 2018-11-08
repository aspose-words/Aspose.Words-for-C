#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/collections/ienumerator.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/ImageType.h>
#include <Model/Drawing/ImageData.h>
#include <Model/Document/FileFormatUtil.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

void ExtractImagesToFiles()
{
    std::cout << "ExtractImagesToFiles example started." << std::endl;
    // ExStart:ExtractImagesToFiles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithImages();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Image.SampleImages.doc");

    System::SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
    int32_t imageIndex = 0;

    for (System::SharedPtr<Shape> shape : System::IterateOver(System::DynamicCastEnumerableTo<System::SharedPtr<Shape>>(shapes)))
    {
        if (shape->get_HasImage())
        {
            System::String imageFileName = System::String::Format(u"Image.ExportImages.{0}_out{1}", imageIndex, FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));
            System::String imagePath = dataDir + imageFileName;
            shape->get_ImageData()->Save(imagePath);
            std::cout << "Image saved at " << imagePath.ToUtf8String() << std::endl;
            imageIndex++;
        }
    }
    // ExEnd:ExtractImagesToFiles
    std::cout << "All images extracted from document." << std::endl;
    std::cout << "ExtractImagesToFiles example finished." << std::endl << std::endl;
}
