#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <system/io/stream.h>
#include <Aspose.Words.Cpp/Model/Fonts/StreamFontSource.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class WarningInfoCollection; } }
namespace Aspose { namespace Words { class WarningInfo; } }
namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { namespace Fields { class FieldStart; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldEnd; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldSeparator; } } }
namespace Aspose { namespace Words { class Run; } }
namespace Aspose { namespace Words { class Paragraph; } }
namespace Aspose { namespace Words { namespace Fields { class FormField; } } }
namespace Aspose { namespace Words { namespace Drawing { class GroupShape; } } }
namespace Aspose { namespace Words { namespace Drawing { class Shape; } } }
namespace Aspose { namespace Words { class Comment; } }
namespace Aspose { namespace Words { class Footnote; } }
namespace Aspose { namespace Words { class SpecialChar; } }
namespace Aspose { namespace Words { namespace Tables { class Table; } } }
namespace Aspose { namespace Words { namespace Tables { class Cell; } } }
namespace Aspose { namespace Words { namespace Tables { class Row; } } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExFont : public ApiExampleBase
{
public:

    class HandleDocumentSubstitutionWarnings : public Aspose::Words::IWarningCallback
    {
        typedef HandleDocumentSubstitutionWarnings ThisType;
        typedef Aspose::Words::IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<Aspose::Words::WarningInfoCollection> FontWarnings;
        
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        
        HandleDocumentSubstitutionWarnings();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    };
    
    /// <summary>
    /// This class when executed will remove all hidden content from the Document. Implemented as a Visitor.
    /// </summary>
    class RemoveHiddenContentVisitor : public Aspose::Words::DocumentVisitor
    {
        typedef RemoveHiddenContentVisitor ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
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
        Aspose::Words::VisitorAction VisitFootnoteStart(System::SharedPtr<Aspose::Words::Footnote> footnote) override;
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
    
    
private:

    /// <summary>
    /// Load the font data only when it is required and not to store it in the memory for the "FontSettings" lifetime.
    /// </summary>
    class StreamFontSourceFile : public Aspose::Words::Fonts::StreamFontSource
    {
        typedef StreamFontSourceFile ThisType;
        typedef Aspose::Words::Fonts::StreamFontSource BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::IO::Stream> OpenFontDataStream() override;
        
    protected:
    
        virtual ~StreamFontSourceFile();
        
    };
    
    
public:

    void CreateFormattedRun();
    void Caps();
    void GetDocumentFonts();
    void DefaultValuesEmbeddedFontsParameters();
    void FontInfoCollection();
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
    void Shading();
    void Bidi();
    void FarEast();
    void Names();
    void ChangeStyle();
    void Style();
    void SubstitutionNotification();
    void GetAvailableFonts();
    void EnableFontSubstitution();
    void DisableFontSubstitution();
    void SubstitutionWarnings();
    void SubstitutionWarningsClosestMatch();
    void SetFontAutoColor();
    void RemoveHiddenContentFromDocument();
    void BlankDocumentFonts();
    void ExtractEmbeddedFont();
    void GetFontInfoFromFile();
    void FontSourceFile();
    void FontSourceFolder();
    void FontSourceMemory();
    void FontSourceSystem();
    void LoadFontFallbackSettingsFromFile();
    void LoadFontFallbackSettingsFromStream();
    void LoadNotoFontsFallbackSettings();
    void DefaultFontSubstitutionRule();
    void FontConfigSubstitution();
    void FallbackSettings();
    void FallbackSettingsCustom();
    void TableSubstitutionRule();
    void TableSubstitutionRuleCustom();
    void ResolveFontsBeforeLoadingDocument();
    void LineSpacing();
    void HasDmlEffect();
    void StreamFontSourceFileRendering();
    void CheckScanUserFontsFolder();
    
protected:

    void TestRemoveHiddenContent(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


