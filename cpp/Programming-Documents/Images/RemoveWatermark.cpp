#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/collections/ienumerator.h>
#include <Model/Sections/HeaderFooter.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Drawing/Shape.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{

// ExStart:RemoveWatermark
void RemoveWatermarkText(const System::SharedPtr<Document>& doc)
{
    auto hf_enumerator = doc->GetChildNodes(Aspose::Words::NodeType::HeaderFooter, true)->GetEnumerator();
    System::SharedPtr<HeaderFooter> hf;
    while (hf_enumerator->MoveNext() && (hf = System::DynamicCast<HeaderFooter>(hf_enumerator->get_Current()), true))
    {
        auto shape_enumerator = hf->GetChildNodes(Aspose::Words::NodeType::Shape, true)->GetEnumerator();
        System::SharedPtr<Shape> shape;
        while (shape_enumerator->MoveNext() && (shape = System::DynamicCast<Shape>(shape_enumerator->get_Current()), true))
        {
            if (shape->get_Name().Contains(u"WaterMark"))
            {
                shape->Remove();
            }
        }
    }
}
// ExEnd:RemoveWatermark

}

void RemoveWatermark()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithImages();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"RemoveWatermark.doc");
    RemoveWatermarkText(doc);
    dataDir = dataDir + GetOutputFilePath(u"RemoveWatermark.doc");
    doc->Save(dataDir);
}
