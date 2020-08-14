#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <system/test_tools/method_argument_tuple.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Math { enum class MathObjectType; } } }
namespace Aspose { namespace Words { namespace Settings { enum class MsWordVersion; } } }
namespace Aspose { namespace Words { namespace Drawing { enum class ShapeMarkupLanguage : uint8_t; } } }
namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { namespace Drawing { class Shape; } } }
namespace Aspose { namespace Words { namespace Drawing { class GroupShape; } } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { namespace Drawing { enum class ShapeType; } } }

namespace ApiExamples {

/// <summary>
/// Examples using shapes in documents.
/// </summary>
class ExShape : public ApiExampleBase
{
private:

    /// <summary>
    /// DocumentVisitor implementation that collects information about visited shapes into a StringBuilder, to be printed to the console.
    /// </summary>
    class ShapeVisitor : public Aspose::Words::DocumentVisitor
    {
        typedef ShapeVisitor ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ShapeVisitor();
        
        /// <summary>
        /// Return all the text that the StringBuilder has accumulated.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when the start of a Shape node is visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        /// <summary>
        /// Called when the end of a Shape node is visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitShapeEnd(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        /// <summary>
        /// Called when the start of a GroupShape node is visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        /// <summary>
        /// Called when the end of a GroupShape node is visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitGroupShapeEnd(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        int32_t mShapesVisited;
        int32_t mTextIndentLevel;
        System::SharedPtr<System::Text::StringBuilder> mStringBuilder;
        
        /// <summary>
        /// Appends a line to the StringBuilder with one prepended tab character for each indent level.
        /// </summary>
        void AppendLine(System::String text);
        
    };
    
    
public:

    void Insert();
    void AspectRatioLockedDefaultValue();
    void Coordinates();
    void InsertGroupShape();
    void DeleteAllShapes();
    void CheckShapeInline();
    void LineFlipOrientation();
    void Fill();
    void Title();
    void ReplaceTextboxesWithImages();
    void CreateTextBox();
    void GetActiveXControlProperties();
    void GetOleObjectRawData();
    void OleControl();
    void OleLinks();
    void OleControlCollection();
    void SuggestedFileName();
    void ObjectDidNotHaveSuggestedFileName();
    void ResolutionDefaultValues();
    void SaveShapeObjectAsImage();
    void OfficeMathDisplayException();
    void OfficeMathDefaultValue();
    void OfficeMath();
    void CannotBeSetDisplayWithInlineJustification();
    void CannotBeSetInlineDisplayWithJustification();
    void OfficeMathDisplayNestedObjects();
    void WorkWithMathObjectType(int32_t index, Aspose::Words::Math::MathObjectType objectType);
    void AspectRatioLocked(bool isLocked);
    void MarkupLunguageByDefault();
    void MarkupLunguageForDifferentMsWordVersions(Aspose::Words::Settings::MsWordVersion msWordVersion, Aspose::Words::Drawing::ShapeMarkupLanguage shapeMarkupLanguage);
    void ChangeStrokeProperties();
    void InsertOleObjectAsHtmlFile();
    void InsertOlePackage();
    void GetAccessToOlePackage();
    void Resize();
    void LayoutInTableCell();
    void ShapeInsertion();
    void VisitShapes();
    void SignatureLine();
    void TextBox();
    void TextBoxShapeType();
    void CreateLinkBetweenTextBoxes();
    void GetTextBoxAndChangeTextAnchor();
    void InsertTextPaths();
    void ShapeRevision();
    void MoveRevisions();
    void AdjustWithEffects();
    void RenderAllShapes();
    void DocumentHasSmartArtObject();
    void OfficeMathRenderer();
    
protected:

    /// <summary>
    /// Insert a new paragraph with a WordArt shape inside it.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Drawing::Shape> AppendWordArt(System::SharedPtr<Aspose::Words::Document> doc, System::String text, System::String textFontFamily, double shapeWidth, double shapeHeight, System::Drawing::Color wordArtFill, System::Drawing::Color line, Aspose::Words::Drawing::ShapeType wordArtShapeType);
    void TestInsertTextPaths(System::String filename);
    
};

} // namespace ApiExamples


