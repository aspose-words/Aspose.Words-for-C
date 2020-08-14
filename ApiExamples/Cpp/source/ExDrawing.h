#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { namespace Drawing { class GroupShape; } } }
namespace Aspose { namespace Words { namespace Drawing { class Shape; } } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExDrawing : public ApiExampleBase
{
public:

    /// <summary>
    /// Visitor that prints shape group contents information to the console.
    /// </summary>
    class ShapeInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef ShapeInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ShapeInfoPrinter();
        
        System::String GetText();
        Aspose::Words::VisitorAction VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        Aspose::Words::VisitorAction VisitGroupShapeEnd(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        Aspose::Words::VisitorAction VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        Aspose::Words::VisitorAction VisitShapeEnd(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    
public:

    void VariousShapes();
    void TypeOfImage();
    void SaveAllImages();
    void ImportImage();
    void StrokePattern();
    void GroupOfShapes();
    void TextBox();
    void GetDataFromImage();
    void ImageData();
    void ImageSize();
    
protected:

    void TestGroupShapes(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


