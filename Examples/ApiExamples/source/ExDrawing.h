#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Drawing;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDrawing : public ApiExampleBase
{
    typedef ExDrawing ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Prints the contents of a visited shape group to the console.
    /// </summary>
    class ShapeGroupPrinter : public DocumentVisitor
    {
        typedef ShapeGroupPrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ShapeGroupPrinter();
        
        System::String GetText();
        Aspose::Words::VisitorAction VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        Aspose::Words::VisitorAction VisitGroupShapeEnd(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        Aspose::Words::VisitorAction VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        Aspose::Words::VisitorAction VisitShapeEnd(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    
public:

    void TypeOfImage();
    void FillSolid();
    void StrokePattern();
    //ExStart
    //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
    //ExFor:DocumentVisitor.VisitShapeStart(Shape)
    //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
    //ExFor:GroupShape
    //ExFor:GroupShape.#ctor(DocumentBase)
    //ExFor:GroupShape.Accept(DocumentVisitor)
    //ExFor:GroupShape.AcceptStart(DocumentVisitor)
    //ExFor:GroupShape.AcceptEnd(DocumentVisitor)
    //ExFor:ShapeBase.IsGroup
    //ExFor:ShapeBase.ShapeType
    //ExSummary:Shows how to create a group of shapes, and print its contents using a document visitor.
    void GroupOfShapes();
    void TextBox();
    void GetDataFromImage();
    void ImageData();
    void ImageSize();
    
protected:

    //ExEnd
    static void TestGroupShapes(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


