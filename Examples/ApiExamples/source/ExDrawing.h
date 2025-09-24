#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
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
    void GroupOfShapes();
    void TextBox();
    void GetDataFromImage();
    void ImageData();
    void ImageSize();
    
protected:

    static void TestGroupShapes(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


