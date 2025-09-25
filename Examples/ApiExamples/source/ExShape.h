#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Math/MathObjectType.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBoxWrapMode.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBoxAnchor.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Settings;

namespace Aspose {

namespace Words {

namespace ApiExamples {

/// <summary>
/// Examples using shapes in documents.
/// </summary>
class ExShape : public ApiExampleBase
{
    typedef ExShape ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Logs appearance-related information about visited shapes.
    /// </summary>
    class ShapeAppearancePrinter : public DocumentVisitor
    {
        typedef ShapeAppearancePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ShapeAppearancePrinter();
        
        /// <summary>
        /// Return all the text that the StringBuilder has accumulated.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when this visitor visits the start of a Shape node.
        /// </summary>
        Aspose::Words::VisitorAction VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        /// <summary>
        /// Called when this visitor visits the end of a Shape node.
        /// </summary>
        Aspose::Words::VisitorAction VisitShapeEnd(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        /// <summary>
        /// Called when this visitor visits the start of a GroupShape node.
        /// </summary>
        Aspose::Words::VisitorAction VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        /// <summary>
        /// Called when this visitor visits the end of a GroupShape node.
        /// </summary>
        Aspose::Words::VisitorAction VisitGroupShapeEnd(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        
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

    void AltText();
    void Font(bool hideShape);
    void Rotate();
    void Coordinates();
    void GroupShape();
    void IsTopLevel();
    void LocalToParent();
    void AnchorLocked(bool anchorLocked);
    void DeleteAllShapes();
    void IsInline();
    void Bounds();
    void FlipShapeOrientation();
    void Fill();
    void TextureFill();
    void GradientFill();
    void GradientStops();
    void FillPattern();
    void FillThemeColor();
    void FillTintAndShade();
    void Title();
    void ReplaceTextboxesWithImages();
    void CreateTextBox();
    void ZOrder();
    void GetActiveXControlProperties();
    void GetOleObjectRawData();
    void LinkedChartSourceFullName();
    void OleControl();
    void OleLinks();
    void OleControlCollection();
    void SuggestedFileName();
    void ObjectDidNotHaveSuggestedFileName();
    void RenderOfficeMath();
    void OfficeMathDisplayException();
    void OfficeMathDefaultValue();
    void OfficeMath();
    void CannotBeSetDisplayWithInlineJustification();
    void CannotBeSetInlineDisplayWithJustification();
    void OfficeMathDisplayNestedObjects();
    void WorkWithMathObjectType(int32_t index, Aspose::Words::Math::MathObjectType objectType);
    void AspectRatio(bool lockAspectRatio);
    void MarkupLanguageByDefault();
    void MarkupLanguageForDifferentMsWordVersions(Aspose::Words::Settings::MsWordVersion msWordVersion, Aspose::Words::Drawing::ShapeMarkupLanguage shapeMarkupLanguage);
    void Stroke();
    void InsertOleObjectAsHtmlFile();
    void InsertOlePackage();
    void GetAccessToOlePackage();
    void Resize();
    void Calendar();
    void IsLayoutInCell(bool isLayoutInCell);
    void ShapeInsertion();
    void VisitShapes();
    void SignatureLine();
    void TextBoxLayoutFlow(Aspose::Words::Drawing::LayoutFlow layoutFlow);
    void TextBoxFitShapeToText();
    void TextBoxMargins();
    void TextBoxContentsWrapMode(Aspose::Words::Drawing::TextBoxWrapMode textBoxWrapMode);
    void TextBoxShapeType();
    void CreateLinkBetweenTextBoxes();
    void VerticalAnchor(Aspose::Words::Drawing::TextBoxAnchor verticalAnchor);
    void InsertTextPaths();
    void ShapeRevision();
    void MoveRevisions();
    void AdjustWithEffects();
    void RenderAllShapes();
    void DocumentHasSmartArtObject();
    void OfficeMathRenderer();
    void ShapeTypes();
    void IsDecorative();
    void FillImage();
    void ShadowFormat();
    void NoTextRotation();
    void RelativeSizeAndPosition();
    void FillBaseColor();
    void FitImageToShape();
    void StrokeForeThemeColors();
    void StrokeBackThemeColors();
    void TextBoxOleControl();
    void Glow();
    void Reflection();
    void SoftEdge();
    void Adjustments();
    void ShadowFormatColor();
    void SetActiveXProperties();
    void SelectRadioControl();
    void CheckedCheckBox();
    void InsertGroupShape();
    void CombineGroupShape();
    void InsertCommandButton();
    void Hidden();
    void CommandButtonCaption();
    void ShadowFormatTransparency();
    
protected:

    /// <summary>
    /// Insert a new paragraph with a WordArt shape inside it.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Drawing::Shape> AppendWordArt(System::SharedPtr<Aspose::Words::Document> doc, System::String text, System::String textFontFamily, double shapeWidth, double shapeHeight, System::Drawing::Color wordArtFill, System::Drawing::Color line, Aspose::Words::Drawing::ShapeType wordArtShapeType);
    void TestInsertTextPaths(System::String filename);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


