#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextPath.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Rendering/ShapeRenderer.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleFormat.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace
{
    void InsertShapeUsingDocumentBuilder(System::String const &outputDataDir)
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

        System::SharedPtr<OoxmlSaveOptions> so = System::MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        // "Strict" or "Transitional" compliance allows to save shape as DML.
        so->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        System::String outputPath = outputDataDir + u"WorkingWithShapes.InsertShapeUsingDocumentBuilder.docx";
        // Save the document to disk.
        doc->Save(outputPath, so);
        // ExEnd:InsertShapeUsingDocumentBuilder
        std::cout << "Insert Shape successfully using DocumentBuilder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetAspectRatioLocked(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetAspectRatioLocked
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<Shape> shape = builder->InsertImage(inputDataDir + u"Test.png");
        shape->set_AspectRatioLocked(false);

        System::String outputPath = outputDataDir + u"WorkingWithShapes.SetAspectRatioLocked.doc";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetAspectRatioLocked
        std::cout << "Shape's AspectRatioLocked property is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetShapeLayoutInCell(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetShapeLayoutInCell
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"LayoutInCell.docx");
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

        System::String outputPath = outputDataDir + u"WorkingWithShapes.SetShapeLayoutInCell.docx";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetShapeLayoutInCell
        std::cout << "Shape's IsLayoutInCell property is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void AddCornersSnipped(System::String const &outputDataDir)
    {
        // ExStart:AddCornersSnipped
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<Shape> shape = builder->InsertShape(ShapeType::TopCornersSnipped, 50, 50);

        System::SharedPtr<OoxmlSaveOptions> so = System::MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        so->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);
        System::String outputPath = outputDataDir + u"WorkingWithShapes.AddCornersSnipped.docx";

        //Save the document to disk.
        doc->Save(outputPath, so);
        // ExEnd:AddCornersSnipped
        std::cout << "Corner Snip shape is created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void GetActualShapeBoundsPoints(System::String const &inputDataDir)
    {
        // ExStart:GetActualShapeBoundsPoints
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<Shape> shape = builder->InsertImage(inputDataDir + u"Test.png");
        shape->set_AspectRatioLocked(false);
        std::cout << "Gets the actual bounds of the shape in points." << shape->GetShapeRenderer()->get_BoundsInPoints().ToString().ToUtf8String() << std::endl;
        // ExEnd:GetActualShapeBoundsPoints
    }

    void DocumentBuilderHorizontalRuleFormat(System::String const &outputDataDir)
    {
        std::cout << "DocumentBuilderHorizontalRuleFormat example started." << std::endl;

        // ExStart:DocumentBuilderInsertHorizontalRule
        // Initialize document.
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>();

        System::SharedPtr<Shape> shape = builder->InsertHorizontalRule();
        System::SharedPtr<HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();

        horizontalRuleFormat->set_Alignment(HorizontalRuleAlignment::Center);
        horizontalRuleFormat->set_WidthPercent(70);
        horizontalRuleFormat->set_Height(3);
        horizontalRuleFormat->set_Color(System::Drawing::Color::get_Blue());
        horizontalRuleFormat->set_NoShade(true);

        builder->get_Document()->Save(outputDataDir + u"HorizontalRuleFormat.docx");

        // ExEnd:DocumentBuilderInsertHorizontalRule
        std::cout << "DocumentBuilderHorizontalRuleFormat example finished." << std::endl << std::endl;
    }
}

void WorkingWithShapes()
{
    std::cout << "WorkingWithShapes example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithShapes();
    System::String outputDataDir = GetOutputDataDir_WorkingWithShapes();
    SetShapeLayoutInCell(inputDataDir, outputDataDir);
    SetAspectRatioLocked(inputDataDir, outputDataDir);
    InsertShapeUsingDocumentBuilder(outputDataDir);
    AddCornersSnipped(outputDataDir);
    GetActualShapeBoundsPoints(inputDataDir);
    DocumentBuilderHorizontalRuleFormat(outputDataDir);
    //SpecifyVerticalAnchor(dataDir); // Source document is missing
    //DetectSmartArtShape(dataDir); // Source document is missing
    std::cout << "WorkingWithShapes example finished." << std::endl << std::endl;
}