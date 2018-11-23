#include "stdafx.h"
#include "examples.h"

#include <drawing/color.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Fill.h>
#include <Model/Drawing/HorizontalAlignment.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/ShapeType.h>
#include <Model/Drawing/RelativeVerticalPosition.h>
#include <Model/Drawing/RelativeHorizontalPosition.h>
#include <Model/Drawing/TextPath.h>
#include <Model/Drawing/VerticalAlignment.h>
#include <Model/Drawing/WrapType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Settings/CompatibilityOptions.h>
#include <Model/Settings/MsWordVersion.h>
#include <Model/Text/Run.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Settings;

namespace
{
    void InsertShapeUsingDocumentBuilder(System::String const &dataDir)
    {
        // ExStart:InsertShapeUsingDocumentBuilder
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        //Free-floating shape insertion.
        System::SharedPtr<Shape> shape = builder->InsertShape(ShapeType::TextBox, RelativeHorizontalPosition::Page, 100, RelativeVerticalPosition::Page, 100, 50, 50, WrapType::None);
        shape->set_Rotation(30.0);

        builder->Writeln();

        //Inline shape insertion.
        shape = builder->InsertShape(ShapeType::TextBox, 50, 50);
        shape->set_Rotation(30.0);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithShapes.InsertShapeUsingDocumentBuilder.docx");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:InsertShapeUsingDocumentBuilder
        std::cout << "Insert Shape successfully using DocumentBuilder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetAspectRatioLocked(System::String const &dataDir)
    {
        // ExStart:SetAspectRatioLocked
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        auto shape = builder->InsertImage(dataDir + u"Test.png");
        shape->set_AspectRatioLocked(false);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithShapes.SetAspectRatioLocked.doc");

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetAspectRatioLocked
        std::cout << "Shape's AspectRatioLocked property is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetShapeLayoutInCell(System::String const &dataDir)
    {
        // ExStart:SetShapeLayoutInCell
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"LayoutInCell.docx");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Shape> watermark = System::MakeObject<Shape>(doc, ShapeType::TextPlainText);
        watermark->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        watermark->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        watermark->set_IsLayoutInCell(true);
        // Display the shape outside of table cell if it will be placed into a cell.

        watermark->set_Width(300);
        watermark->set_Height(70);
        watermark->set_HorizontalAlignment(HorizontalAlignment::Center);
        watermark->set_VerticalAlignment(VerticalAlignment::Center);

        watermark->set_Rotation(-40);
        watermark->get_Fill()->set_Color(System::Drawing::Color::get_Gray());
        watermark->set_StrokeColor(System::Drawing::Color::get_Gray());

        watermark->get_TextPath()->set_Text(u"watermarkText");
        watermark->get_TextPath()->set_FontFamily(u"Arial");

        watermark->set_Name(System::String::Format(u"WaterMark_{0}",System::Guid::NewGuid()));
        watermark->set_WrapType(WrapType::None);

        System::SharedPtr<Run> run = System::DynamicCast_noexcept<Run>(doc->GetChildNodes(NodeType::Run, true)->idx_get(doc->GetChildNodes(NodeType::Run, true)->get_Count() - 1));

        builder->MoveTo(run);
        builder->InsertNode(watermark);
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2010);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithShapes.SetShapeLayoutInCell.docx");

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetShapeLayoutInCell
        std::cout << "Shape's IsLayoutInCell property is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithShapes()
{
    std::cout << "WorkingWithShapes example started." << std::endl;
    // ExStart:WorkingWithShapes
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithShapes();
    SetShapeLayoutInCell(dataDir);
    SetAspectRatioLocked(dataDir);
    InsertShapeUsingDocumentBuilder(dataDir);
    // ExEnd:WorkingWithShapes
    std::cout << "WorkingWithShapes example finished." << std::endl << std::endl;
}