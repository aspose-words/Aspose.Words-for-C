#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Nodes/Node.h>
#include <Model/Drawing/ShapeType.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/GroupShape.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <drawing/size.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

void AddGroupShapeToDocument()
{
    std::cout << "AddGroupShapeToDocument example started." << std::endl;
    // ExStart:AddGroupShapeToDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    doc->EnsureMinimum();
    System::SharedPtr<GroupShape> gs = System::MakeObject<GroupShape>(doc);

    System::SharedPtr<Shape> shape = System::MakeObject<Shape>(doc, ShapeType::AccentBorderCallout1);
    shape->set_Width(100);
    shape->set_Height(100);
    gs->AppendChild(shape);

    System::SharedPtr<Shape> shape1 = System::MakeObject<Shape>(doc, ShapeType::ActionButtonBeginning);
    shape1->set_Left(100);
    shape1->set_Width(100);
    shape1->set_Height(200);
    gs->AppendChild(shape1);

    gs->set_Width(200);
    gs->set_Height(200);

    gs->set_CoordSize(System::Drawing::Size(200, 200));

    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->InsertNode(gs);

    System::String outputPath = dataDir + GetOutputFilePath(u"AddGroupShapeToDocument.doc");

    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:AddGroupShapeToDocument
    std::cout << "Group shape added successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AddGroupShapeToDocument example finished." << std::endl << std::endl;
}
