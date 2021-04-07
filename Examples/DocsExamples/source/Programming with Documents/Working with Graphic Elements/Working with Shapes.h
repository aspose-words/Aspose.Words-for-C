#pragma once

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBoxAnchor.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextPath.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Rendering/ShapeRenderer.h>
#include <drawing/color.h>
#include <drawing/rectangle_f.h>
#include <drawing/size.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/guid.h>
#include <system/linq/enumerable.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements {

class WorkingWithShapes : public DocsExamplesBase
{
public:
    void AddGroupShape()
    {
        //ExStart:AddGroupShape
        auto doc = MakeObject<Document>();
        doc->EnsureMinimum();

        auto groupShape = MakeObject<GroupShape>(doc);
        auto accentBorderShape = MakeObject<Shape>(doc, ShapeType::AccentBorderCallout1);
        accentBorderShape->set_Width(100);
        accentBorderShape->set_Height(100);
        groupShape->AppendChild(accentBorderShape);

        auto actionButtonShape = MakeObject<Shape>(doc, ShapeType::ActionButtonBeginning);
        actionButtonShape->set_Left(100);
        actionButtonShape->set_Width(100);
        actionButtonShape->set_Height(200);
        groupShape->AppendChild(actionButtonShape);

        groupShape->set_Width(200);
        groupShape->set_Height(200);
        groupShape->set_CoordSize(System::Drawing::Size(200, 200));

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertNode(groupShape);

        doc->Save(ArtifactsDir + u"WorkingWithShapes.AddGroupShape.docx");
        //ExEnd:AddGroupShape
    }

    void InsertShape()
    {
        //ExStart:InsertShape
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape =
            builder->InsertShape(ShapeType::TextBox, RelativeHorizontalPosition::Page, 100, RelativeVerticalPosition::Page, 100, 50, 50, WrapType::None);
        shape->set_Rotation(30.0);

        builder->Writeln();

        shape = builder->InsertShape(ShapeType::TextBox, 50, 50);
        shape->set_Rotation(30.0);

        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        doc->Save(ArtifactsDir + u"WorkingWithShapes.InsertShape.docx", saveOptions);
        //ExEnd:InsertShape
    }

    void AspectRatioLocked()
    {
        //ExStart:AspectRatioLocked
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertImage(ImagesDir + u"Transparent background logo.png");
        shape->set_AspectRatioLocked(false);

        doc->Save(ArtifactsDir + u"WorkingWithShapes.AspectRatioLocked.docx");
        //ExEnd:AspectRatioLocked
    }

    void LayoutInCell()
    {
        //ExStart:LayoutInCell
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->get_RowFormat()->set_Height(100);
        builder->get_RowFormat()->set_HeightRule(HeightRule::Exactly);

        for (int i = 0; i < 31; i++)
        {
            if (i != 0 && i % 7 == 0)
            {
                builder->EndRow();
            }
            builder->InsertCell();
            builder->Write(u"Cell contents");
        }

        builder->EndTable();

        auto watermark = MakeObject<Shape>(doc, ShapeType::TextPlainText);
        watermark->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        watermark->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        watermark->set_IsLayoutInCell(true);
        watermark->set_Width(300);
        watermark->set_Height(70);
        watermark->set_HorizontalAlignment(HorizontalAlignment::Center);
        watermark->set_VerticalAlignment(VerticalAlignment::Center);
        watermark->set_Rotation(-40);

        watermark->get_Fill()->set_Color(System::Drawing::Color::get_Gray());
        watermark->set_StrokeColor(System::Drawing::Color::get_Gray());

        watermark->get_TextPath()->set_Text(u"watermarkText");
        watermark->get_TextPath()->set_FontFamily(u"Arial");

        watermark->set_Name(String::Format(u"WaterMark_{0}", System::Guid::NewGuid()));
        watermark->set_WrapType(WrapType::None);

        auto run =
            System::DynamicCast_noexcept<Run>(doc->GetChildNodes(NodeType::Run, true)->idx_get(doc->GetChildNodes(NodeType::Run, true)->get_Count() - 1));

        builder->MoveTo(run);
        builder->InsertNode(watermark);
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2010);

        doc->Save(ArtifactsDir + u"WorkingWithShapes.LayoutInCell.docx");
        //ExEnd:LayoutInCell
    }

    void AddCornersSnipped()
    {
        //ExStart:AddCornersSnipped
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertShape(ShapeType::TopCornersSnipped, 50, 50);

        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        doc->Save(ArtifactsDir + u"WorkingWithShapes.AddCornersSnipped.docx", saveOptions);
        //ExEnd:AddCornersSnipped
    }

    void GetActualShapeBoundsPoints()
    {
        //ExStart:GetActualShapeBoundsPoints
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertImage(ImagesDir + u"Transparent background logo.png");
        shape->set_AspectRatioLocked(false);

        std::cout << "\nGets the actual bounds of the shape in points: ";
        std::cout << shape->GetShapeRenderer()->get_BoundsInPoints() << std::endl;
        //ExEnd:GetActualShapeBoundsPoints
    }

    void VerticalAnchor()
    {
        //ExStart:VerticalAnchor
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 200, 200);
        textBox->get_TextBox()->set_VerticalAnchor(TextBoxAnchor::Bottom);

        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"Textbox contents");

        doc->Save(ArtifactsDir + u"WorkingWithShapes.VerticalAnchor.docx");
        //ExEnd:VerticalAnchor
    }

    void DetectSmartArtShape()
    {
        //ExStart:DetectSmartArtShape
        auto doc = MakeObject<Document>(MyDir + u"SmartArt.docx");

        int count = doc->GetChildNodes(NodeType::Shape, true)
                        ->LINQ_Cast<SharedPtr<Shape>>()
                        ->LINQ_Count([](SharedPtr<Shape> shape) { return shape->get_HasSmartArt(); });

        std::cout << "The document has " << count << " shapes with SmartArt." << std::endl;
        //ExEnd:DetectSmartArtShape
    }

    void UpdateSmartArtDrawing()
    {
        auto doc = MakeObject<Document>(MyDir + u"SmartArt.docx");

        //ExStart:UpdateSmartArtDrawing
        for (auto shape : System::IterateOver<Shape>(doc->GetChildNodes(NodeType::Shape, true)))
        {
            if (shape->get_HasSmartArt())
            {
                shape->UpdateSmartArtDrawing();
            }
        }
        //ExEnd:UpdateSmartArtDrawing
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements
