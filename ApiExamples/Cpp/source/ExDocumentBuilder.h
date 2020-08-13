#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldResultFormatter.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class CalendarType; } }
namespace Aspose { namespace Words { namespace Fields { enum class GeneralFormat; } } }

namespace ApiExamples {

class ExDocumentBuilder : public ApiExampleBase
{
private:

    /// <summary>
    /// Custom IFieldResult implementation that applies formats and tracks format invocations
    /// </summary>
    class FieldResultFormatter : public Aspose::Words::Fields::IFieldResultFormatter
    {
        typedef FieldResultFormatter ThisType;
        typedef Aspose::Words::Fields::IFieldResultFormatter BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FieldResultFormatter(System::String numberFormat, System::String dateFormat, System::String generalFormat);
        
        System::String FormatNumeric(double value, System::String format) override;
        System::String FormatDateTime(System::DateTime value, System::String format, Aspose::Words::CalendarType calendarType) override;
        System::String Format(System::String value, Aspose::Words::Fields::GeneralFormat format) override;
        System::String Format(double value, Aspose::Words::Fields::GeneralFormat format) override;
        void PrintInvocations();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::String mNumberFormat;
        System::String mDateFormat;
        System::String mGeneralFormat;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<System::Object>>> mNumberFormatInvocations;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<System::Object>>> mDateFormatInvocations;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<System::Object>>> mGeneralFormatInvocations;
        
        System::String Format(System::SharedPtr<System::Object> value, Aspose::Words::Fields::GeneralFormat format);
        
    };
    
    
public:

    void WriteAndFont();
    void HeadersAndFooters();
    void MergeFields();
    void InsertHorizontalRule();
    void HorizontalRuleFormatExceptions();
    void InsertHyperlink();
    void PushPopFont();
    void InsertWatermark();
    void InsertOleObject();
    void InsertHtml();
    void InsertHtmlWithFormatting();
    void MathML();
    void InsertTextAndBookmark();
    void CreateForm();
    void InsertCheckBox();
    void InsertCheckBoxEmptyName();
    void WorkingWithNodes();
    void FillMergeFields();
    void InsertToc();
    void InsertTable();
    void InsertTableWithStyle();
    void InsertTableSetHeadingRow();
    void InsertTableWithPreferredWidth();
    void InsertCellsWithPreferredWidths();
    void InsertTableFromHtml();
    void InsertNestedTable();
    void CreateSimpleTable();
    void BuildFormattedTable();
    void TableBordersAndShading();
    void SetPreferredTypeConvertUtil();
    void InsertHyperlinkToLocalBookmark();
    void DocumentBuilderCursorPosition();
    void DocumentBuilderMoveToNode();
    void DocumentBuilderMoveToDocumentStartEnd();
    void DocumentBuilderMoveToSection();
    void DocumentBuilderMoveToParagraph();
    void DocumentBuilderMoveToTableCell();
    void DocumentBuilderMoveToBookmarkEnd();
    void DocumentBuilderBuildTable();
    void TableCellVerticalRotatedFarEastTextOrientation();
    void DocumentBuilderInsertBreak();
    void DocumentBuilderInsertInlineImage();
    void DocumentBuilderInsertFloatingImage();
    void InsertImageFromUrl();
    void InsertImageOriginalSize();
    void DocumentBuilderInsertTextInputFormField();
    void DocumentBuilderInsertComboBoxFormField();
    void DocumentBuilderInsertToc();
    void SignatureLineProviderId();
    void InsertSignatureLineCurrentPosition();
    void DocumentBuilderSetFontFormatting();
    void DocumentBuilderSetParagraphFormatting();
    void DocumentBuilderSetCellFormatting();
    void DocumentBuilderSetRowFormatting();
    void DocumentBuilderSetListFormatting();
    void DocumentBuilderSetSectionFormatting();
    void InsertFootnote();
    void DocumentBuilderApplyParagraphStyle();
    void DocumentBuilderApplyBordersAndShading();
    void DeleteRow();
    void InsertDocument();
    void KeepSourceNumbering();
    void ResolveStyleBehaviorWhileAppendDocument();
    void IgnoreTextBoxes(bool isIgnoreTextBoxes);
    void MoveToField();
    void InsertOleObjectException();
    void InsertChartDouble();
    void InsertChartRelativePosition();
    void InsertField();
    void InsertFieldByType();
    void FieldResultFormatting();
    void InsertVideoWithUrl();
    void InsertUnderline();
    void CurrentStory();
    void InsertOlePowerpoint();
    void InsertStyleSeparator();
    void WithoutStyleSeparator();
    void SmartStyleBehavior();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentEmphases();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentInlineCode();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentHeadings();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentBlockquotes();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentIndentedCode();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentFencedCode();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentHorizontalRule();
    /// <summary>
    /// All markdown tests work with the same file
    /// That's why we need order for them 
    /// </summary>
    void MarkdownDocumentBulletedList();
    /// <summary>
    /// All markdown tests work with the same file.
    /// That's why we need order for them.
    /// </summary>
    void LoadMarkdownDocumentAndAssertContent(System::String text, System::String styleName, bool isItalic, bool isBold);
    void MarkdownDocumentTests();
    void InsertOnlineVideo();
    
};

} // namespace ApiExamples


