#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/date_time.h>
#include <system/collections/list.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Fields/IFieldResultFormatter.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/GeneralFormat.h>
#include <Aspose.Words.Cpp/Model/DateTimeFormatting/CalendarType.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fields;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocumentBuilder : public ApiExampleBase
{
    typedef ExDocumentBuilder ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// When fields with formatting are updated, this formatter will override their formatting
    /// with a custom format, while tracking every invocation.
    /// </summary>
    class FieldResultFormatter : public IFieldResultFormatter
    {
        typedef FieldResultFormatter ThisType;
        typedef IFieldResultFormatter BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        enum class FormatInvocationType
        {
            Numeric,
            DateTime,
            General,
            All
        };
        
        
    private:
    
        class FormatInvocation : public System::Object
        {
            typedef FormatInvocation ThisType;
            typedef System::Object BaseType;
            
            typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
            RTTI_INFO_DECL();
            
        public:
        
            Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType get_FormatInvocationType() const;
            System::SharedPtr<System::Object> get_Value() const;
            System::String get_OriginalFormat() const;
            System::String get_NewValue() const;
            
            FormatInvocation(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType formatInvocationType, System::SharedPtr<System::Object> value, System::String originalFormat, System::String newValue);
            
        private:
        
            Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType pr_FormatInvocationType;
            System::SharedPtr<System::Object> pr_Value;
            System::String pr_OriginalFormat;
            System::String pr_NewValue;
            
        };
        
        
    public:
    
        FieldResultFormatter(System::String numberFormat, System::String dateFormat, System::String generalFormat);
        
        System::String FormatNumeric(double value, System::String format) override;
        System::String FormatDateTime(System::DateTime value, System::String format, Aspose::Words::CalendarType calendarType) override;
        System::String Format(System::String value, Aspose::Words::Fields::GeneralFormat format) override;
        System::String Format(double value, Aspose::Words::Fields::GeneralFormat format) override;
        int32_t CountFormatInvocations(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType formatInvocationType);
        void PrintFormatInvocations();
        
    private:
    
        System::String mNumberFormat;
        System::String mDateFormat;
        System::String mGeneralFormat;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>>> mFormatInvocations;
        
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>>> get_FormatInvocations() const;
        
        System::String Format(System::SharedPtr<System::Object> value, Aspose::Words::Fields::GeneralFormat format);
        
    };
    
    
public:

    void WriteAndFont();
    void MergeFields();
    void InsertHorizontalRule();
    void HorizontalRuleFormatExceptions();
    void InsertHyperlink();
    void PushPopFont();
    void InsertWatermark();
    void InsertOleObject();
    void InsertHtml();
    void InsertHtmlWithFormatting(bool useBuilderFormatting);
    void MathMl();
    void InsertTextAndBookmark();
    void CreateColumnBookmark();
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
    void CreateTable();
    void BuildFormattedTable();
    void TableBordersAndShading();
    void SetPreferredTypeConvertUtil();
    void InsertHyperlinkToLocalBookmark();
    void CursorPosition();
    void MoveTo();
    void MoveToParagraph();
    void MoveToCell();
    void MoveToBookmark();
    void BuildTable();
    void TableCellVerticalRotatedFarEastTextOrientation();
    void InsertFloatingImage();
    void InsertImageOriginalSize();
    void InsertTextInput();
    void InsertComboBox();
    void SignatureLineProviderId();
    void SignatureLineInline();
    void SetParagraphFormatting();
    void SetCellFormatting();
    void SetRowFormatting();
    void InsertFootnote();
    void ApplyBordersAndShading();
    void DeleteRow();
    void AppendDocumentAndResolveStyles(bool keepSourceNumbering);
    void InsertDocumentAndResolveStyles(bool keepSourceNumbering);
    void LoadDocumentWithListNumbering(bool keepSourceNumbering);
    void IgnoreTextBoxes(bool ignoreTextBoxes);
    void MoveToField(bool moveCursorToAfterTheField);
    void InsertOleObjectException();
    void InsertPieChart();
    void InsertChartRelativePosition();
    void InsertField();
    void InsertFieldAndUpdate(bool updateInsertedFieldsImmediately);
    void FieldResultFormatting();
    void InsertVideoWithUrl();
    void InsertUnderline();
    void CurrentStory();
    void InsertOleObjects();
    void InsertDocument();
    void SmartStyleBehavior();
    void EmphasesWarningSourceMarkdown();
    void DoNotIgnoreHeaderFooter();
    void MarkdownDocumentEmphases();
    void MarkdownDocumentInlineCode();
    void MarkdownDocumentHeadings();
    void MarkdownDocumentBlockquotes();
    void MarkdownDocumentIndentedCode();
    void MarkdownDocumentFencedCode();
    void MarkdownDocumentHorizontalRule();
    void MarkdownDocumentBulletedList();
    void LoadMarkdownDocumentAndAssertContent(System::String text, System::String styleName, bool isItalic, bool isBold);
    void InsertOnlineVideoCustomThumbnail();
    void InsertOleObjectAsIcon();
    void PreserveBlocks();
    void PhoneticGuide();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


