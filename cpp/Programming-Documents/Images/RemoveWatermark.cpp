#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Sections/HeaderFooter.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Drawing/Shape.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

// ExStart:RemoveWatermark
namespace
{
    void RemoveWatermarkText(const System::SharedPtr<Document>& doc)
    {
        System::SharedPtr<NodeCollection> headerFooterNodes = doc->GetChildNodes(NodeType::HeaderFooter, true);
        for (System::SharedPtr<HeaderFooter> hf : System::IterateOver<System::SharedPtr<HeaderFooter>>(headerFooterNodes))
        {
            System::SharedPtr<NodeCollection> shapeNodes = hf->GetChildNodes(NodeType::Shape, true);
            for (System::SharedPtr<Shape> shape: System::IterateOver<System::SharedPtr<Shape>>(shapeNodes))
            {
                if (shape->get_Name().Contains(u"WaterMark"))
                {
                    shape->Remove();
                }
            }
        }
    }
}

void RemoveWatermark()
{
    std::cout << "RemoveWatermark example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithImages();
    System::String outputDataDir = GetOutputDataDir_WorkingWithImages();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"RemoveWatermark.docx");
    RemoveWatermarkText(doc);
    System::String outputPath = outputDataDir + u"RemoveWatermark.docx";
    doc->Save(outputPath);
    std:: cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveWatermark example finished." << std::endl << std::endl;
}
// ExEnd:RemoveWatermark