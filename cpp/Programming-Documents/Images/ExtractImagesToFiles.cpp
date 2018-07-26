#include "../../examples.h"

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
    // ExStart:ExtractImagesToFiles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithImages();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Image.SampleImages.doc");
    
    System::SharedPtr<NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    int32_t imageIndex = 0;
    
    auto shape_enumerator = shapes->GetEnumerator();
    System::SharedPtr<Shape> shape;
    while (shape_enumerator->MoveNext() && (shape = System::DynamicCast<Shape>(shape_enumerator->get_Current()), true))
    {
        if (shape->get_HasImage())
        {
            System::String imageFileName = System::String::Format(u"Image.ExportImages.{0}_out{1}", imageIndex, FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));
            shape->get_ImageData()->Save(dataDir + imageFileName);
            imageIndex++;
        }
    }
    // ExEnd:ExtractImagesToFiles
    std::cout << "\nAll images extracted from document.\n";
}
