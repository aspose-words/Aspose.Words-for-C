#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
#include <Aspose.Words.Cpp/Model/Text/SpecialChar.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/EmphasisMark.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExFont : public ApiExampleBase
{
    typedef ExFont ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Removes all visited nodes marked as "hidden content".
    /// </summary>
    class RemoveHiddenContentVisitor : public DocumentVisitor
    {
        typedef RemoveHiddenContentVisitor ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldStart(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart) override;
        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd) override;
        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldSeparator(System::SharedPtr<Aspose::Words::Fields::FieldSeparator> fieldSeparator) override;
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitParagraphStart(System::SharedPtr<Aspose::Words::Paragraph> paragraph) override;
        /// <summary>
        /// Called when a FormField is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFormField(System::SharedPtr<Aspose::Words::Fields::FormField> formField) override;
        /// <summary>
        /// Called when a GroupShape is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape) override;
        /// <summary>
        /// Called when a Shape is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape) override;
        /// <summary>
        /// Called when a Comment is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentStart(System::SharedPtr<Aspose::Words::Comment> comment) override;
        /// <summary>
        /// Called when a Footnote is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFootnoteStart(System::SharedPtr<Aspose::Words::Notes::Footnote> footnote) override;
        /// <summary>
        /// Called when a SpecialCharacter is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitSpecialChar(System::SharedPtr<Aspose::Words::SpecialChar> specialChar) override;
        /// <summary>
        /// Called when visiting of a Table node is ended in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitTableEnd(System::SharedPtr<Aspose::Words::Tables::Table> table) override;
        /// <summary>
        /// Called when visiting of a Cell node is ended in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCellEnd(System::SharedPtr<Aspose::Words::Tables::Cell> cell) override;
        /// <summary>
        /// Called when visiting of a Row node is ended in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRowEnd(System::SharedPtr<Aspose::Words::Tables::Row> row) override;
        
    };
    
    
public:

    void CreateFormattedRun();
    void Caps();
    void GetDocumentFonts();
    void DefaultValuesEmbeddedFontsParameters();
    void FontInfoCollection(bool embedAllFonts);
    void WorkWithEmbeddedFonts(bool embedTrueTypeFonts, bool embedSystemFonts, bool saveSubsetFonts);
    void StrikeThrough();
    void PositionSubscript();
    void ScalingSpacing();
    void Italic();
    void EngraveEmboss();
    void Shadow();
    void Outline();
    void Hidden();
    void Kerning();
    void NoProofing();
    void LocaleId();
    void Underlines();
    void ComplexScript();
    void SparklingText();
    void ForegroundAndBackground();
    void Shading();
    void Bidi();
    void FarEast();
    void NameAscii();
    void ChangeStyle();
    void BuiltIn();
    void Style();
    void GetAvailableFonts();
    void SetFontAutoColor();
    void RemoveHiddenContentFromDocument();
    void DefaultFonts();
    void ExtractEmbeddedFont();
    void GetFontInfoFromFile();
    void LineSpacing();
    void HasDmlEffect();
    void SetEmphasisMark(Aspose::Words::EmphasisMark emphasisMark);
    void ThemeFontsColors();
    void CreateThemedStyle();
    void FontInfoEmbeddingLicensingRights();
    void PhysicalFontInfoEmbeddingLicensingRights();
    void NumberSpacing();
    
protected:

    void TestRemoveHiddenContent(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


