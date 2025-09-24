// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBuilder.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file.h>
#include <system/guid.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <iostream>
#include <functional>
#include <drawing/image.h>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/TextOrientation.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/PhoneticGuide.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TableStyleOptions.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidthType.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/CellVerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Sections/StoryType.h>
#include <Aspose.Words.Cpp/Model/Sections/Story.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/DropDownItemCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToc.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OlePackage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Model/Document/HtmlInsertOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/Borders/TextureIndex.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1818158218u, ::Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation, ThisTypeBaseTypesInfo);

Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType ExDocumentBuilder::FieldResultFormatter::FormatInvocation::get_FormatInvocationType() const
{
    return pr_FormatInvocationType;
}

System::SharedPtr<System::Object> ExDocumentBuilder::FieldResultFormatter::FormatInvocation::get_Value() const
{
    return pr_Value;
}

System::String ExDocumentBuilder::FieldResultFormatter::FormatInvocation::get_OriginalFormat() const
{
    return pr_OriginalFormat;
}

System::String ExDocumentBuilder::FieldResultFormatter::FormatInvocation::get_NewValue() const
{
    return pr_NewValue;
}

ExDocumentBuilder::FieldResultFormatter::FormatInvocation::FormatInvocation(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType formatInvocationType, System::SharedPtr<System::Object> value, System::String originalFormat, System::String newValue)
    : pr_FormatInvocationType(((Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType)0))
{
    pr_Value = value;
    pr_FormatInvocationType = formatInvocationType;
    pr_OriginalFormat = originalFormat;
    pr_NewValue = newValue;
}


RTTI_INFO_IMPL_HASH(3868185539u, ::Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter, ThisTypeBaseTypesInfo);

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>>> ExDocumentBuilder::FieldResultFormatter::get_FormatInvocations() const
{
    return mFormatInvocations;
}

ExDocumentBuilder::FieldResultFormatter::FieldResultFormatter(System::String numberFormat, System::String dateFormat, System::String generalFormat)
    : mFormatInvocations(System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>>>())
{
    mNumberFormat = numberFormat;
    mDateFormat = dateFormat;
    mGeneralFormat = generalFormat;
}

System::String ExDocumentBuilder::FieldResultFormatter::FormatNumeric(double value, System::String format)
{
    if (System::String::IsNullOrEmpty(mNumberFormat))
    {
        return nullptr;
    }
    
    System::String newValue = System::String::Format(mNumberFormat, value);
    get_FormatInvocations()->Add(System::MakeObject<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::Numeric, System::ExplicitCast<System::Object>(value), format, newValue));
    return newValue;
}

System::String ExDocumentBuilder::FieldResultFormatter::FormatDateTime(System::DateTime value, System::String format, Aspose::Words::CalendarType calendarType)
{
    if (System::String::IsNullOrEmpty(mDateFormat))
    {
        return nullptr;
    }
    
    System::String newValue = System::String::Format(mDateFormat, value);
    get_FormatInvocations()->Add(System::MakeObject<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::DateTime, System::ExplicitCast<System::Object>(System::String::Format(u"{0} ({1})", value, calendarType)), format, newValue));
    return newValue;
}

System::String ExDocumentBuilder::FieldResultFormatter::Format(System::String value, Aspose::Words::Fields::GeneralFormat format)
{
    return Format(System::ExplicitCast<System::Object>(value), format);
}

System::String ExDocumentBuilder::FieldResultFormatter::Format(double value, Aspose::Words::Fields::GeneralFormat format)
{
    return Format(System::ExplicitCast<System::Object>(value), format);
}

System::String ExDocumentBuilder::FieldResultFormatter::Format(System::SharedPtr<System::Object> value, Aspose::Words::Fields::GeneralFormat format)
{
    if (System::String::IsNullOrEmpty(mGeneralFormat))
    {
        return nullptr;
    }
    
    System::String newValue = System::String::Format(mGeneralFormat, value);
    get_FormatInvocations()->Add(System::MakeObject<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::General, value, System::ObjectExt::ToString(format), newValue));
    return newValue;
}

int32_t ExDocumentBuilder::FieldResultFormatter::CountFormatInvocations(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType formatInvocationType)
{
    if (formatInvocationType == Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::All)
    {
        return get_FormatInvocations()->get_Count();
    }
    return get_FormatInvocations()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation> f)>>([&formatInvocationType](System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocation> f) -> bool
    {
        return f->get_FormatInvocationType() == formatInvocationType;
    })));
}

void ExDocumentBuilder::FieldResultFormatter::PrintFormatInvocations()
{
    for (auto&& f : get_FormatInvocations())
    {
        std::cout << (System::String::Format(u"Invocation type:\t{0}\n", f->get_FormatInvocationType()) + System::String::Format(u"\tOriginal value:\t\t{0}\n", f->get_Value()) + System::String::Format(u"\tOriginal format:\t{0}\n", f->get_OriginalFormat()) + System::String::Format(u"\tNew value:\t\t\t{0}\n", f->get_NewValue())) << std::endl;
    }
}


RTTI_INFO_IMPL_HASH(2565804880u, ::Aspose::Words::ApiExamples::ExDocumentBuilder, ThisTypeBaseTypesInfo);

void ExDocumentBuilder::MarkdownDocumentEmphases()
{
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    // Bold and Italic are represented as Font.Bold and Font.Italic.
    builder->get_Font()->set_Italic(true);
    builder->Writeln(u"This text will be italic");
    
    // Use clear formatting if we don't want to combine styles between paragraphs.
    builder->get_Font()->ClearFormatting();
    
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"This text will be bold");
    
    builder->get_Font()->ClearFormatting();
    
    builder->get_Font()->set_Italic(true);
    builder->Write(u"You ");
    builder->get_Font()->set_Bold(true);
    builder->Write(u"can");
    builder->get_Font()->set_Bold(false);
    builder->Writeln(u" combine them");
    
    builder->get_Font()->ClearFormatting();
    
    builder->get_Font()->set_StrikeThrough(true);
    builder->Writeln(u"This text will be strikethrough");
    
    // Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis.
    builder->get_Document()->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentInlineCode()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Prepare our created document for further work
    // and clear paragraph formatting not to use the previous styles.
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");
    
    // Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`).
    // If number of backticks is missed, then one backtick will be used by default.
    System::SharedPtr<Aspose::Words::Style> inlineCode1BackTicks = doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"InlineCode");
    builder->get_Font()->set_Style(inlineCode1BackTicks);
    builder->Writeln(u"Text with InlineCode style with one backtick");
    
    // Use optional dot (.) and number of backticks (`).
    // There will be 3 backticks.
    System::SharedPtr<Aspose::Words::Style> inlineCode3BackTicks = doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"InlineCode.3");
    builder->get_Font()->set_Style(inlineCode3BackTicks);
    builder->Writeln(u"Text with InlineCode style with 3 backticks");
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentHeadings()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Prepare our created document for further work
    // and clear paragraph formatting not to use the previous styles.
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");
    
    // By default, Heading styles in Word may have bold and italic formatting.
    // If we do not want text to be emphasized, set these properties explicitly to false.
    // Thus we can't use 'builder.Font.ClearFormatting()' because Bold/Italic will be set to true.
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);
    
    // Create for one heading for each level.
    builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
    builder->get_Font()->set_Italic(true);
    builder->Writeln(u"This is an italic H1 tag");
    
    // Reset our styles from the previous paragraph to not combine styles between paragraphs.
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);
    
    // Structure-enhanced text heading can be added through style inheritance.
    System::SharedPtr<Aspose::Words::Style> setextHeading1 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"SetextHeading1");
    builder->get_ParagraphFormat()->set_Style(setextHeading1);
    doc->get_Styles()->idx_get(u"SetextHeading1")->set_BaseStyleName(u"Heading 1");
    builder->Writeln(u"SetextHeading 1");
    
    builder->get_ParagraphFormat()->set_StyleName(u"Heading 2");
    builder->Writeln(u"This is an H2 tag");
    
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);
    
    System::SharedPtr<Aspose::Words::Style> setextHeading2 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"SetextHeading2");
    builder->get_ParagraphFormat()->set_Style(setextHeading2);
    doc->get_Styles()->idx_get(u"SetextHeading2")->set_BaseStyleName(u"Heading 2");
    builder->Writeln(u"SetextHeading 2");
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 3"));
    builder->Writeln(u"This is an H3 tag");
    
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 4"));
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"This is an bold H4 tag");
    
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 5"));
    builder->get_Font()->set_Italic(true);
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"This is an italic and bold H5 tag");
    
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 6"));
    builder->Writeln(u"This is an H6 tag");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentBlockquotes()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Prepare our created document for further work
    // and clear paragraph formatting not to use the previous styles.
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");
    
    // By default, the document stores blockquote style for the first level.
    builder->get_ParagraphFormat()->set_StyleName(u"Quote");
    builder->Writeln(u"Blockquote");
    
    // Create styles for nested levels through style inheritance.
    System::SharedPtr<Aspose::Words::Style> quoteLevel2 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote1");
    builder->get_ParagraphFormat()->set_Style(quoteLevel2);
    doc->get_Styles()->idx_get(u"Quote1")->set_BaseStyleName(u"Quote");
    builder->Writeln(u"1. Nested blockquote");
    
    System::SharedPtr<Aspose::Words::Style> quoteLevel3 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote2");
    builder->get_ParagraphFormat()->set_Style(quoteLevel3);
    doc->get_Styles()->idx_get(u"Quote2")->set_BaseStyleName(u"Quote1");
    builder->get_Font()->set_Italic(true);
    builder->Writeln(u"2. Nested italic blockquote");
    
    System::SharedPtr<Aspose::Words::Style> quoteLevel4 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote3");
    builder->get_ParagraphFormat()->set_Style(quoteLevel4);
    doc->get_Styles()->idx_get(u"Quote3")->set_BaseStyleName(u"Quote2");
    builder->get_Font()->set_Italic(false);
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"3. Nested bold blockquote");
    
    System::SharedPtr<Aspose::Words::Style> quoteLevel5 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote4");
    builder->get_ParagraphFormat()->set_Style(quoteLevel5);
    doc->get_Styles()->idx_get(u"Quote4")->set_BaseStyleName(u"Quote3");
    builder->get_Font()->set_Bold(false);
    builder->Writeln(u"4. Nested blockquote");
    
    System::SharedPtr<Aspose::Words::Style> quoteLevel6 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote5");
    builder->get_ParagraphFormat()->set_Style(quoteLevel6);
    doc->get_Styles()->idx_get(u"Quote5")->set_BaseStyleName(u"Quote4");
    builder->Writeln(u"5. Nested blockquote");
    
    System::SharedPtr<Aspose::Words::Style> quoteLevel7 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote6");
    builder->get_ParagraphFormat()->set_Style(quoteLevel7);
    doc->get_Styles()->idx_get(u"Quote6")->set_BaseStyleName(u"Quote5");
    builder->get_Font()->set_Italic(true);
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"6. Nested italic bold blockquote");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentIndentedCode()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Prepare our created document for further work
    // and clear paragraph formatting not to use the previous styles.
    builder->MoveToDocumentEnd();
    builder->Writeln(u"\n");
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");
    
    System::SharedPtr<Aspose::Words::Style> indentedCode = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"IndentedCode");
    builder->get_ParagraphFormat()->set_Style(indentedCode);
    builder->Writeln(u"This is an indented code");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentFencedCode()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Prepare our created document for further work
    // and clear paragraph formatting not to use the previous styles.
    builder->MoveToDocumentEnd();
    builder->Writeln(u"\n");
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");
    
    System::SharedPtr<Aspose::Words::Style> fencedCode = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"FencedCode");
    builder->get_ParagraphFormat()->set_Style(fencedCode);
    builder->Writeln(u"This is a fenced code");
    
    System::SharedPtr<Aspose::Words::Style> fencedCodeWithInfo = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"FencedCode.C#");
    builder->get_ParagraphFormat()->set_Style(fencedCodeWithInfo);
    builder->Writeln(u"This is a fenced code with info string");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentHorizontalRule()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Prepare our created document for further work
    // and clear paragraph formatting not to use the previous styles.
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");
    
    // Insert HorizontalRule that will be present in .md file as '-----'.
    builder->InsertHorizontalRule();
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentBulletedList()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Prepare our created document for further work
    // and clear paragraph formatting not to use the previous styles.
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");
    
    // Bulleted lists are represented using paragraph numbering.
    builder->get_ListFormat()->ApplyBulletDefault();
    // There can be 3 types of bulleted lists.
    // The only diff in a numbering format of the very first level are ‘-’, ‘+’ or ‘*’ respectively.
    builder->get_ListFormat()->get_List()->get_ListLevels()->idx_get(0)->set_NumberFormat(u"-");
    
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Item 2a");
    builder->Writeln(u"Item 2b");
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
}


namespace gtest_test
{

class ExDocumentBuilder : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentBuilder> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDocumentBuilder>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentBuilder> ExDocumentBuilder::s_instance;

} // namespace gtest_test

void ExDocumentBuilder::WriteAndFont()
{
    //ExStart
    //ExFor:Font.Size
    //ExFor:Font.Bold
    //ExFor:Font.Name
    //ExFor:Font.Color
    //ExFor:Font.Underline
    //ExFor:DocumentBuilder.#ctor
    //ExSummary:Shows how to insert formatted text using DocumentBuilder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Specify font formatting, then add text.
    System::SharedPtr<Aspose::Words::Font> font = builder->get_Font();
    font->set_Size(16);
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_Blue());
    font->set_Name(u"Courier New");
    font->set_Underline(Aspose::Words::Underline::Dash);
    
    builder->Write(u"Hello world!");
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(builder->get_Document());
    System::SharedPtr<Aspose::Words::Run> firstRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Hello world!", firstRun->GetText().Trim());
    ASPOSE_ASSERT_EQ(16.0, firstRun->get_Font()->get_Size());
    ASSERT_TRUE(firstRun->get_Font()->get_Bold());
    ASSERT_EQ(u"Courier New", firstRun->get_Font()->get_Name());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), firstRun->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Underline::Dash, firstRun->get_Font()->get_Underline());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, WriteAndFont)
{
    s_instance->WriteAndFont();
}

} // namespace gtest_test

void ExDocumentBuilder::MergeFields()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertField(String)
    //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
    //ExSummary:Shows how to insert fields, and move the document builder's cursor to them.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertField(u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
    builder->InsertField(u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT");
    
    // Move the cursor to the first MERGEFIELD.
    builder->MoveToMergeField(u"MyMergeField1", true, false);
    
    // Note that the cursor is placed immediately after the first MERGEFIELD, and before the second.
    ASPOSE_ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(1)->get_Start(), builder->get_CurrentNode());
    ASPOSE_ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(0)->get_End(), builder->get_CurrentNode()->get_PreviousSibling());
    
    // If we wish to edit the field's field code or contents using the builder,
    // its cursor would need to be inside a field.
    // To place it inside a field, we would need to call the document builder's MoveTo method
    // and pass the field's start or separator node as an argument.
    builder->Write(u" Text between our merge fields. ");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MergeFields.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MergeFields.docx");
    
    ASSERT_EQ(System::String(u"\u0013MERGEFIELD MyMergeField1 \\* MERGEFORMAT\u0014«MyMergeField1»\u0015") + u" Text between our merge fields. " + u"\u0013MERGEFIELD MyMergeField2 \\* MERGEFORMAT\u0014«MyMergeField2»\u0015", doc->GetText().Trim());
    ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldMergeField, u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT", u"«MyMergeField1»", doc->get_Range()->get_Fields()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldMergeField, u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT", u"«MyMergeField2»", doc->get_Range()->get_Fields()->idx_get(1));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MergeFields)
{
    s_instance->MergeFields();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertHorizontalRule()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertHorizontalRule
    //ExFor:ShapeBase.IsHorizontalRule
    //ExFor:Shape.HorizontalRuleFormat
    //ExFor:HorizontalRuleAlignment
    //ExFor:HorizontalRuleFormat
    //ExFor:HorizontalRuleFormat.Alignment
    //ExFor:HorizontalRuleFormat.WidthPercent
    //ExFor:HorizontalRuleFormat.Height
    //ExFor:HorizontalRuleFormat.Color
    //ExFor:HorizontalRuleFormat.NoShade
    //ExSummary:Shows how to insert a horizontal rule shape, and customize its formatting.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertHorizontalRule();
    
    System::SharedPtr<Aspose::Words::Drawing::HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();
    horizontalRuleFormat->set_Alignment(Aspose::Words::Drawing::HorizontalRuleAlignment::Center);
    horizontalRuleFormat->set_WidthPercent(70);
    horizontalRuleFormat->set_Height(3);
    horizontalRuleFormat->set_Color(System::Drawing::Color::get_Blue());
    horizontalRuleFormat->set_NoShade(true);
    
    ASSERT_TRUE(shape->get_IsHorizontalRule());
    ASSERT_TRUE(shape->get_HorizontalRuleFormat()->get_NoShade());
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::HorizontalRuleAlignment::Center, shape->get_HorizontalRuleFormat()->get_Alignment());
    ASPOSE_ASSERT_EQ(70, shape->get_HorizontalRuleFormat()->get_WidthPercent());
    ASPOSE_ASSERT_EQ(3, shape->get_HorizontalRuleFormat()->get_Height());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), shape->get_HorizontalRuleFormat()->get_Color().ToArgb());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertHorizontalRule)
{
    s_instance->InsertHorizontalRule();
}

} // namespace gtest_test

void ExDocumentBuilder::HorizontalRuleFormatExceptions()
{
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertHorizontalRule();
    
    System::SharedPtr<Aspose::Words::Drawing::HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();
    horizontalRuleFormat->set_WidthPercent(1);
    horizontalRuleFormat->set_WidthPercent(100);
    ASSERT_THROW(static_cast<std::function<void()>>([&horizontalRuleFormat]() -> void
    {
        horizontalRuleFormat->set_WidthPercent(0);
    })(), System::ArgumentOutOfRangeException);
    ASSERT_THROW(static_cast<std::function<void()>>([&horizontalRuleFormat]() -> void
    {
        horizontalRuleFormat->set_WidthPercent(101);
    })(), System::ArgumentOutOfRangeException);
    
    horizontalRuleFormat->set_Height(0);
    horizontalRuleFormat->set_Height(1584);
    ASSERT_THROW(static_cast<std::function<void()>>([&horizontalRuleFormat]() -> void
    {
        horizontalRuleFormat->set_Height(-1);
    })(), System::ArgumentOutOfRangeException);
    ASSERT_THROW(static_cast<std::function<void()>>([&horizontalRuleFormat]() -> void
    {
        horizontalRuleFormat->set_Height(1585);
    })(), System::ArgumentOutOfRangeException);
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, HorizontalRuleFormatExceptions)
{
    RecordProperty("Description", "Checking the boundary conditions of WidthPercent and Height properties");
    
    s_instance->HorizontalRuleFormatExceptions();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertHyperlink()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertHyperlink
    //ExFor:Font.ClearFormatting
    //ExFor:Font.Color
    //ExFor:Font.Underline
    //ExFor:Underline
    //ExSummary:Shows how to insert a hyperlink field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"For more information, please visit the ");
    
    // Insert a hyperlink and emphasize it with custom formatting.
    // The hyperlink will be a clickable piece of text which will take us to the location specified in the URL.
    builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    builder->get_Font()->set_Underline(Aspose::Words::Underline::Single);
    builder->InsertHyperlink(u"Google website", u"https://www.google.com", false);
    builder->get_Font()->ClearFormatting();
    builder->Writeln(u".");
    
    // Ctrl + left clicking the link in the text in Microsoft Word will take us to the URL via a new web browser window.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertHyperlink.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertHyperlink.docx");
    
    auto hyperlink = System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));
    ASSERT_EQ(u"https://www.google.com", hyperlink->get_Address());
    
    // This field is written as w:hyperlink element therefore field code cannot have formatting.
    auto fieldCode = System::ExplicitCast<Aspose::Words::Run>(hyperlink->get_Start()->get_NextSibling());
    ASSERT_EQ(u"HYPERLINK \"https://www.google.com\"", fieldCode->GetText().Trim());
    
    auto fieldResult = System::ExplicitCast<Aspose::Words::Run>(hyperlink->get_Separator()->get_NextSibling());
    
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), fieldResult->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Underline::Single, fieldResult->get_Font()->get_Underline());
    ASSERT_EQ(u"Google website", fieldResult->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertHyperlink)
{
    s_instance->InsertHyperlink();
}

} // namespace gtest_test

void ExDocumentBuilder::PushPopFont()
{
    //ExStart
    //ExFor:DocumentBuilder.PushFont
    //ExFor:DocumentBuilder.PopFont
    //ExFor:DocumentBuilder.InsertHyperlink
    //ExSummary:Shows how to use a document builder's formatting stack.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set up font formatting, then write the text that goes before the hyperlink.
    builder->get_Font()->set_Name(u"Arial");
    builder->get_Font()->set_Size(24);
    builder->Write(u"To visit Google, hold Ctrl and click ");
    
    // Preserve our current formatting configuration on the stack.
    builder->PushFont();
    
    // Alter the builder's current formatting by applying a new style.
    builder->get_Font()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Hyperlink);
    builder->InsertHyperlink(u"here", u"http://www.google.com", false);
    
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), builder->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Underline::Single, builder->get_Font()->get_Underline());
    
    // Restore the font formatting that we saved earlier and remove the element from the stack.
    builder->PopFont();
    
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), builder->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Underline::None, builder->get_Font()->get_Underline());
    
    builder->Write(u". We hope you enjoyed the example.");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.PushPopFont.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.PushPopFont.docx");
    System::SharedPtr<Aspose::Words::RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();
    
    ASSERT_EQ(4, runs->get_Count());
    
    ASSERT_EQ(u"To visit Google, hold Ctrl and click", runs->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u". We hope you enjoyed the example.", runs->idx_get(3)->GetText().Trim());
    ASPOSE_ASSERT_EQ(runs->idx_get(0)->get_Font()->get_Color(), runs->idx_get(3)->get_Font()->get_Color());
    ASSERT_EQ(runs->idx_get(0)->get_Font()->get_Underline(), runs->idx_get(3)->get_Font()->get_Underline());
    
    ASSERT_EQ(u"here", runs->idx_get(2)->GetText().Trim());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), runs->idx_get(2)->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Underline::Single, runs->idx_get(2)->get_Font()->get_Underline());
    ASPOSE_ASSERT_NE(runs->idx_get(0)->get_Font()->get_Color(), runs->idx_get(2)->get_Font()->get_Color());
    ASSERT_NE(runs->idx_get(0)->get_Font()->get_Underline(), runs->idx_get(2)->get_Font()->get_Underline());
    
    ASSERT_EQ(u"http://www.google.com", (System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, PushPopFont)
{
    s_instance->PushPopFont();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertWatermark()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToHeaderFooter
    //ExFor:PageSetup.PageWidth
    //ExFor:PageSetup.PageHeight
    //ExFor:WrapType
    //ExFor:RelativeHorizontalPosition
    //ExFor:RelativeVerticalPosition
    //ExSummary:Shows how to insert an image, and use it as a watermark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert the image into the header so that it will be visible on every page.
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Transparent background logo.png");
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    shape->set_BehindText(true);
    
    // Place the image at the center of the page.
    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
    shape->set_Left((builder->get_PageSetup()->get_PageWidth() - shape->get_Width()) / 2);
    shape->set_Top((builder->get_PageSetup()->get_PageHeight() - shape->get_Height()) / 2);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertWatermark.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertWatermark.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, shape->get_WrapType());
    ASSERT_TRUE(shape->get_BehindText());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Page, shape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Page, shape->get_RelativeVerticalPosition());
    ASPOSE_ASSERT_EQ((doc->get_FirstSection()->get_PageSetup()->get_PageWidth() - shape->get_Width()) / 2, shape->get_Left());
    ASPOSE_ASSERT_EQ((doc->get_FirstSection()->get_PageSetup()->get_PageHeight() - shape->get_Height()) / 2, shape->get_Top());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertWatermark)
{
    s_instance->InsertWatermark();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertOleObject()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertOleObject(String, Boolean, Boolean, Stream)
    //ExFor:DocumentBuilder.InsertOleObject(String, String, Boolean, Boolean, Stream)
    //ExFor:DocumentBuilder.InsertOleObjectAsIcon(String, Boolean, String, String)
    //ExSummary:Shows how to insert an OLE object into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // OLE objects are links to files in our local file system that can be opened by other installed applications.
    // Double clicking these shapes will launch the application, and then use it to open the linked object.
    // There are three ways of using the InsertOleObject method to insert these shapes and configure their appearance.
    // 1 -  Image taken from the local file system:
    {
        auto imageStream = System::MakeObject<System::IO::FileStream>(get_ImageDir() + u"Logo.jpg", System::IO::FileMode::Open);
        // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
        // the icon according to the file extension and uses the filename for the icon caption.
        builder->InsertOleObject(get_MyDir() + u"Spreadsheet.xlsx", false, false, imageStream);
    }
    
    // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
    // the icon according to 'progId' and uses the filename for the icon caption.
    // 2 -  Icon based on the application that will open the object:
    builder->InsertOleObject(get_MyDir() + u"Spreadsheet.xlsx", u"Excel.Sheet", false, true, nullptr);
    
    // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
    // the icon according to 'progId' and uses the predefined icon caption.
    // 3 -  Image icon that's 32 x 32 pixels or smaller from the local file system, with a custom caption:
    builder->InsertOleObjectAsIcon(get_MyDir() + u"Presentation.pptx", false, get_ImageDir() + u"Logo icon.ico", u"Double click to view presentation!");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertOleObject.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertOleObject.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::OleObject, shape->get_ShapeType());
    ASSERT_EQ(u"Excel.Sheet.12", shape->get_OleFormat()->get_ProgId());
    ASSERT_EQ(u".xlsx", shape->get_OleFormat()->get_SuggestedExtension());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::OleObject, shape->get_ShapeType());
    ASSERT_EQ(u"Package", shape->get_OleFormat()->get_ProgId());
    ASSERT_EQ(u".xlsx", shape->get_OleFormat()->get_SuggestedExtension());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::OleObject, shape->get_ShapeType());
    ASSERT_EQ(u"PowerPoint.Show.12", shape->get_OleFormat()->get_ProgId());
    ASSERT_EQ(u".pptx", shape->get_OleFormat()->get_SuggestedExtension());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOleObject)
{
    s_instance->InsertOleObject();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertHtml()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertHtml(String)
    //ExSummary:Shows how to use a document builder to insert html content into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    const System::String html = System::String(u"<p align='right'>Paragraph right</p>") + u"<b>Implicit paragraph left</b>" + u"<div align='center'>Div center</div>" + u"<h1 align='left'>Heading 1 left.</h1>";
    
    builder->InsertHtml(html);
    
    // Inserting HTML code parses the formatting of each element into equivalent document text formatting.
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    ASSERT_EQ(u"Paragraph right", paragraphs->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());
    
    ASSERT_EQ(u"Implicit paragraph left", paragraphs->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
    ASSERT_TRUE(paragraphs->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Bold());
    
    ASSERT_EQ(u"Div center", paragraphs->idx_get(2)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());
    
    ASSERT_EQ(u"Heading 1 left.", paragraphs->idx_get(3)->GetText().Trim());
    ASSERT_EQ(u"Heading 1", paragraphs->idx_get(3)->get_ParagraphFormat()->get_Style()->get_Name());
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertHtml.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertHtml)
{
    s_instance->InsertHtml();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertHtmlWithFormatting(bool useBuilderFormatting)
{
    //ExStart
    //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
    //ExSummary:Shows how to apply a document builder's formatting while inserting HTML content.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set a text alignment for the builder, insert an HTML paragraph with a specified alignment, and one without.
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Distributed);
    builder->InsertHtml(System::String(u"<p align='right'>Paragraph 1.</p>") + u"<p>Paragraph 2.</p>", useBuilderFormatting);
    
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    // The first paragraph has an alignment specified. When InsertHtml parses the HTML code,
    // the paragraph alignment value found in the HTML code always supersedes the document builder's value.
    ASSERT_EQ(u"Paragraph 1.", paragraphs->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());
    
    // The second paragraph has no alignment specified. It can have its alignment value filled in
    // by the builder's value depending on the flag we passed to the InsertHtml method.
    ASSERT_EQ(u"Paragraph 2.", paragraphs->idx_get(1)->GetText().Trim());
    ASSERT_EQ(useBuilderFormatting ? Aspose::Words::ParagraphAlignment::Distributed : Aspose::Words::ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertHtmlWithFormatting.docx");
    //ExEnd
}

namespace gtest_test
{

using ExDocumentBuilder_InsertHtmlWithFormatting_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::InsertHtmlWithFormatting)>::type;

struct ExDocumentBuilder_InsertHtmlWithFormatting : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_InsertHtmlWithFormatting_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_InsertHtmlWithFormatting, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertHtmlWithFormatting(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_InsertHtmlWithFormatting, ::testing::ValuesIn(ExDocumentBuilder_InsertHtmlWithFormatting::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::MathMl()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    const System::String mathMl = u"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>";
    
    builder->InsertHtml(mathMl);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MathML.docx");
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MathML.pdf");
    
    ASSERT_TRUE(Aspose::Words::ApiExamples::DocumentHelper::CompareDocs(get_GoldsDir() + u"DocumentBuilder.MathML Gold.docx", get_ArtifactsDir() + u"DocumentBuilder.MathML.docx"));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MathMl)
{
    s_instance->MathMl();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTextAndBookmark()
{
    //ExStart
    //ExFor:DocumentBuilder.StartBookmark
    //ExFor:DocumentBuilder.EndBookmark
    //ExSummary:Shows how create a bookmark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A valid bookmark needs to have document body text enclosed by
    // BookmarkStart and BookmarkEnd nodes created with a matching bookmark name.
    builder->StartBookmark(u"MyBookmark");
    builder->Writeln(u"Hello world!");
    builder->EndBookmark(u"MyBookmark");
    
    ASSERT_EQ(1, doc->get_Range()->get_Bookmarks()->get_Count());
    ASSERT_EQ(u"MyBookmark", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Name());
    ASSERT_EQ(u"Hello world!", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Text().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTextAndBookmark)
{
    s_instance->InsertTextAndBookmark();
}

} // namespace gtest_test

void ExDocumentBuilder::CreateColumnBookmark()
{
    //ExStart
    //ExFor:DocumentBuilder.StartColumnBookmark
    //ExFor:DocumentBuilder.EndColumnBookmark
    //ExSummary:Shows how to create a column bookmark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartTable();
    
    builder->InsertCell();
    // Cells 1,2,4,5 will be bookmarked.
    builder->StartColumnBookmark(u"MyBookmark_1");
    // Badly formed bookmarks or bookmarks with duplicate names will be ignored when the document is saved.
    builder->StartColumnBookmark(u"MyBookmark_1");
    builder->StartColumnBookmark(u"BadStartBookmark");
    builder->Write(u"Cell 1");
    
    builder->InsertCell();
    builder->Write(u"Cell 2");
    
    builder->InsertCell();
    builder->Write(u"Cell 3");
    
    builder->EndRow();
    
    builder->InsertCell();
    builder->Write(u"Cell 4");
    
    builder->InsertCell();
    builder->Write(u"Cell 5");
    builder->EndColumnBookmark(u"MyBookmark_1");
    builder->EndColumnBookmark(u"MyBookmark_1");
    
    ASSERT_THROW(static_cast<std::function<void()>>([&builder]() -> void
    {
        builder->EndColumnBookmark(u"BadEndBookmark");
    })(), System::InvalidOperationException);
    //ExSkip
    
    builder->InsertCell();
    builder->Write(u"Cell 6");
    
    builder->EndRow();
    builder->EndTable();
    
    doc->Save(get_ArtifactsDir() + u"Bookmarks.CreateColumnBookmark.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, CreateColumnBookmark)
{
    s_instance->CreateColumnBookmark();
}

} // namespace gtest_test

void ExDocumentBuilder::CreateForm()
{
    //ExStart
    //ExFor:TextFormFieldType
    //ExFor:DocumentBuilder.InsertTextInput
    //ExFor:DocumentBuilder.InsertComboBox
    //ExSummary:Shows how to create form fields.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    // Form fields are objects in the document that the user can interact with by being prompted to enter values.
    // We can create them using a document builder, and below are two ways of doing so.
    // 1 -  Basic text input:
    builder->InsertTextInput(u"My text input", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Enter your name here", 30);
    
    // 2 -  Combo box with prompt text, and a range of possible values:
    System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"-- Select your favorite footwear --", u"Sneakers", u"Oxfords", u"Flip-flops", u"Other"});
    
    builder->InsertParagraph();
    builder->InsertComboBox(u"My combo box", items, 0);
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"DocumentBuilder.CreateForm.docx");
    //ExEnd
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.CreateForm.docx");
    System::SharedPtr<Aspose::Words::Fields::FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);
    
    ASSERT_EQ(u"My text input", formField->get_Name());
    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, formField->get_TextInputType());
    ASSERT_EQ(u"Enter your name here", formField->get_Result());
    
    formField = doc->get_Range()->get_FormFields()->idx_get(1);
    
    ASSERT_EQ(u"My combo box", formField->get_Name());
    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, formField->get_TextInputType());
    ASSERT_EQ(u"-- Select your favorite footwear --", formField->get_Result());
    ASSERT_EQ(0, formField->get_DropDownSelectedIndex());
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"-- Select your favorite footwear --", u"Sneakers", u"Oxfords", u"Flip-flops", u"Other"}), formField->get_DropDownItems()->LINQ_ToArray());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, CreateForm)
{
    s_instance->CreateForm();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertCheckBox()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertCheckBox(string, bool, bool, int)
    //ExFor:DocumentBuilder.InsertCheckBox(String, bool, int)
    //ExSummary:Shows how to insert checkboxes into the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert checkboxes of varying sizes and default checked statuses.
    builder->Write(u"Unchecked check box of a default size: ");
    builder->InsertCheckBox(System::String::Empty, false, false, 0);
    builder->InsertParagraph();
    
    builder->Write(u"Large checked check box: ");
    builder->InsertCheckBox(u"CheckBox_Default", true, true, 50);
    builder->InsertParagraph();
    
    // Form fields have a name length limit of 20 characters.
    builder->Write(u"Very large checked check box: ");
    builder->InsertCheckBox(u"CheckBox_OnlyCheckedValue", true, 100);
    
    ASSERT_EQ(u"CheckBox_OnlyChecked", doc->get_Range()->get_FormFields()->idx_get(2)->get_Name());
    
    // We can interact with these check boxes in Microsoft Word by double clicking them.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertCheckBox.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertCheckBox.docx");
    
    System::SharedPtr<Aspose::Words::Fields::FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
    
    ASSERT_EQ(System::String::Empty, formFields->idx_get(0)->get_Name());
    ASPOSE_ASSERT_EQ(false, formFields->idx_get(0)->get_Checked());
    ASPOSE_ASSERT_EQ(false, formFields->idx_get(0)->get_Default());
    ASPOSE_ASSERT_EQ(10, formFields->idx_get(0)->get_CheckBoxSize());
    
    ASSERT_EQ(u"CheckBox_Default", formFields->idx_get(1)->get_Name());
    ASPOSE_ASSERT_EQ(true, formFields->idx_get(1)->get_Checked());
    ASPOSE_ASSERT_EQ(true, formFields->idx_get(1)->get_Default());
    ASPOSE_ASSERT_EQ(50, formFields->idx_get(1)->get_CheckBoxSize());
    
    ASSERT_EQ(u"CheckBox_OnlyChecked", formFields->idx_get(2)->get_Name());
    ASPOSE_ASSERT_EQ(true, formFields->idx_get(2)->get_Checked());
    ASPOSE_ASSERT_EQ(true, formFields->idx_get(2)->get_Default());
    ASPOSE_ASSERT_EQ(100, formFields->idx_get(2)->get_CheckBoxSize());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertCheckBox)
{
    s_instance->InsertCheckBox();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertCheckBoxEmptyName()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Checking that the checkbox insertion with an empty name working correctly
    builder->InsertCheckBox(u"", true, false, 1);
    builder->InsertCheckBox(System::String::Empty, false, 1);
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertCheckBoxEmptyName)
{
    s_instance->InsertCheckBoxEmptyName();
}

} // namespace gtest_test

void ExDocumentBuilder::WorkingWithNodes()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveTo(Node)
    //ExFor:DocumentBuilder.MoveToBookmark(String)
    //ExFor:DocumentBuilder.CurrentParagraph
    //ExFor:DocumentBuilder.CurrentNode
    //ExFor:DocumentBuilder.MoveToDocumentStart
    //ExFor:DocumentBuilder.MoveToDocumentEnd
    //ExFor:DocumentBuilder.IsAtEndOfParagraph
    //ExFor:DocumentBuilder.IsAtStartOfParagraph
    //ExSummary:Shows how to move a document builder's cursor to different nodes in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a valid bookmark, an entity that consists of nodes enclosed by a bookmark start node,
    // and a bookmark end node.
    builder->StartBookmark(u"MyBookmark");
    builder->Write(u"Bookmark contents.");
    builder->EndBookmark(u"MyBookmark");
    
    System::SharedPtr<Aspose::Words::NodeCollection> firstParagraphNodes = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(Aspose::Words::NodeType::Any, false);
    
    ASSERT_EQ(Aspose::Words::NodeType::BookmarkStart, firstParagraphNodes->idx_get(0)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Run, firstParagraphNodes->idx_get(1)->get_NodeType());
    ASSERT_EQ(u"Bookmark contents.", firstParagraphNodes->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::NodeType::BookmarkEnd, firstParagraphNodes->idx_get(2)->get_NodeType());
    
    // The document builder's cursor is always ahead of the node that we last added with it.
    // If the builder's cursor is at the end of the document, its current node will be null.
    // The previous node is the bookmark end node that we last added.
    // Adding new nodes with the builder will append them to the last node.
    ASSERT_TRUE(System::TestTools::IsNull(builder->get_CurrentNode()));
    
    // If we wish to edit a different part of the document with the builder,
    // we will need to bring its cursor to the node we wish to edit.
    builder->MoveToBookmark(u"MyBookmark");
    
    // Moving it to a bookmark will move it to the first node within the bookmark start and end nodes, the enclosed run.
    ASPOSE_ASSERT_EQ(firstParagraphNodes->idx_get(1), builder->get_CurrentNode());
    
    // We can also move the cursor to an individual node like this.
    builder->MoveTo(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(Aspose::Words::NodeType::Any, false)->idx_get(0));
    
    ASSERT_EQ(Aspose::Words::NodeType::BookmarkStart, builder->get_CurrentNode()->get_NodeType());
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), builder->get_CurrentParagraph());
    ASSERT_TRUE(builder->get_IsAtStartOfParagraph());
    
    // We can use specific methods to move to the start/end of a document.
    builder->MoveToDocumentEnd();
    
    ASSERT_TRUE(builder->get_IsAtEndOfParagraph());
    
    builder->MoveToDocumentStart();
    
    ASSERT_TRUE(builder->get_IsAtStartOfParagraph());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, WorkingWithNodes)
{
    s_instance->WorkingWithNodes();
}

} // namespace gtest_test

void ExDocumentBuilder::FillMergeFields()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToMergeField(String)
    //ExFor:DocumentBuilder.Bold
    //ExFor:DocumentBuilder.Italic
    //ExSummary:Shows how to fill MERGEFIELDs with data with a document builder instead of a mail merge.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge,
    // and then fill them manually.
    builder->InsertField(u" MERGEFIELD Chairman ");
    builder->InsertField(u" MERGEFIELD ChiefFinancialOfficer ");
    builder->InsertField(u" MERGEFIELD ChiefTechnologyOfficer ");
    
    builder->MoveToMergeField(u"Chairman");
    builder->set_Bold(true);
    builder->Writeln(u"John Doe");
    
    builder->MoveToMergeField(u"ChiefFinancialOfficer");
    builder->set_Italic(true);
    builder->Writeln(u"Jane Doe");
    
    builder->MoveToMergeField(u"ChiefTechnologyOfficer");
    builder->set_Italic(true);
    builder->Writeln(u"John Bloggs");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.FillMergeFields.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.FillMergeFields.docx");
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    ASSERT_TRUE(paragraphs->idx_get(0)->get_Runs()->idx_get(0)->get_Font()->get_Bold());
    ASSERT_EQ(u"John Doe", paragraphs->idx_get(0)->get_Runs()->idx_get(0)->GetText().Trim());
    
    ASSERT_TRUE(paragraphs->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Italic());
    ASSERT_EQ(u"Jane Doe", paragraphs->idx_get(1)->get_Runs()->idx_get(0)->GetText().Trim());
    
    ASSERT_TRUE(paragraphs->idx_get(2)->get_Runs()->idx_get(0)->get_Font()->get_Italic());
    ASSERT_EQ(u"John Bloggs", paragraphs->idx_get(2)->get_Runs()->idx_get(0)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, FillMergeFields)
{
    s_instance->FillMergeFields();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertToc()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertTableOfContents
    //ExFor:Document.UpdateFields
    //ExFor:DocumentBuilder.#ctor(Document)
    //ExFor:ParagraphFormat.StyleIdentifier
    //ExFor:DocumentBuilder.InsertBreak
    //ExFor:BreakType
    //ExSummary:Shows how to insert a Table of contents (TOC) into a document using heading styles as entries.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a table of contents for the first page of the document.
    // Configure the table to pick up paragraphs with headings of levels 1 to 3.
    // Also, set its entries to be hyperlinks that will take us
    // to the location of the heading when left-clicked in Microsoft Word.
    builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // Populate the table of contents by adding paragraphs with heading styles.
    // Each such heading with a level between 1 and 3 will create an entry in the table.
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Writeln(u"Heading 1");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading2);
    builder->Writeln(u"Heading 1.1");
    builder->Writeln(u"Heading 1.2");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Writeln(u"Heading 2");
    builder->Writeln(u"Heading 3");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading2);
    builder->Writeln(u"Heading 3.1");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading3);
    builder->Writeln(u"Heading 3.1.1");
    builder->Writeln(u"Heading 3.1.2");
    builder->Writeln(u"Heading 3.1.3");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading4);
    builder->Writeln(u"Heading 3.1.3.1");
    builder->Writeln(u"Heading 3.1.3.2");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading2);
    builder->Writeln(u"Heading 3.2");
    builder->Writeln(u"Heading 3.3");
    
    // A table of contents is a field of a type that needs to be updated to show an up-to-date result.
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertToc.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertToc.docx");
    auto tableOfContents = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(u"1-3", tableOfContents->get_HeadingLevelRange());
    ASSERT_TRUE(tableOfContents->get_InsertHyperlinks());
    ASSERT_TRUE(tableOfContents->get_HideInWebLayout());
    ASSERT_TRUE(tableOfContents->get_UseParagraphOutlineLevel());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertToc)
{
    s_instance->InsertToc();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTable()
{
    //ExStart
    //ExFor:DocumentBuilder
    //ExFor:DocumentBuilder.Write(String)
    //ExFor:DocumentBuilder.StartTable
    //ExFor:DocumentBuilder.InsertCell
    //ExFor:DocumentBuilder.EndRow
    //ExFor:DocumentBuilder.EndTable
    //ExFor:DocumentBuilder.CellFormat
    //ExFor:DocumentBuilder.RowFormat
    //ExFor:CellFormat
    //ExFor:CellFormat.FitText
    //ExFor:CellFormat.Width
    //ExFor:CellFormat.VerticalAlignment
    //ExFor:CellFormat.Shading
    //ExFor:CellFormat.Orientation
    //ExFor:CellFormat.WrapText
    //ExFor:RowFormat
    //ExFor:RowFormat.Borders
    //ExFor:RowFormat.ClearFormatting
    //ExFor:Shading.ClearFormatting
    //ExSummary:Shows how to build a table with custom borders.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartTable();
    
    // Setting table formatting options for a document builder
    // will apply them to every row and cell that we add with it.
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    
    builder->get_CellFormat()->ClearFormatting();
    builder->get_CellFormat()->set_Width(150);
    builder->get_CellFormat()->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_GreenYellow());
    builder->get_CellFormat()->set_WrapText(false);
    builder->get_CellFormat()->set_FitText(true);
    
    builder->get_RowFormat()->ClearFormatting();
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    builder->get_RowFormat()->set_Height(50);
    builder->get_RowFormat()->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Engrave3D);
    builder->get_RowFormat()->get_Borders()->set_Color(System::Drawing::Color::get_Orange());
    
    builder->InsertCell();
    builder->Write(u"Row 1, Col 1");
    
    builder->InsertCell();
    builder->Write(u"Row 1, Col 2");
    builder->EndRow();
    
    // Changing the formatting will apply it to the current cell,
    // and any new cells that we create with the builder afterward.
    // This will not affect the cells that we have added previously.
    builder->get_CellFormat()->get_Shading()->ClearFormatting();
    
    builder->InsertCell();
    builder->Write(u"Row 2, Col 1");
    
    builder->InsertCell();
    builder->Write(u"Row 2, Col 2");
    
    builder->EndRow();
    
    // Increase row height to fit the vertical text.
    builder->InsertCell();
    builder->get_RowFormat()->set_Height(150);
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Upward);
    builder->Write(u"Row 3, Col 1");
    
    builder->InsertCell();
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Downward);
    builder->Write(u"Row 3, Col 2");
    
    builder->EndRow();
    builder->EndTable();
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertTable.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertTable.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(u"Row 1, Col 1\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u"Row 1, Col 2\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    ASPOSE_ASSERT_EQ(50.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::LineStyle::Engrave3D, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Borders()->get_LineStyle());
    ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), table->get_Rows()->idx_get(0)->get_RowFormat()->get_Borders()->get_Color().ToArgb());
    
    for (auto&& c : System::IterateOver<Aspose::Words::Tables::Cell>(table->get_Rows()->idx_get(0)->get_Cells()))
    {
        ASPOSE_ASSERT_EQ(150, c->get_CellFormat()->get_Width());
        ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, c->get_CellFormat()->get_VerticalAlignment());
        ASSERT_EQ(System::Drawing::Color::get_GreenYellow().ToArgb(), c->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_FALSE(c->get_CellFormat()->get_WrapText());
        ASSERT_TRUE(c->get_CellFormat()->get_FitText());
        
        ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, c->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
    }
    
    ASSERT_EQ(u"Row 2, Col 1\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u"Row 2, Col 2\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
    
    
    for (auto&& c : System::IterateOver<Aspose::Words::Tables::Cell>(table->get_Rows()->idx_get(1)->get_Cells()))
    {
        ASPOSE_ASSERT_EQ(150, c->get_CellFormat()->get_Width());
        ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, c->get_CellFormat()->get_VerticalAlignment());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_FALSE(c->get_CellFormat()->get_WrapText());
        ASSERT_TRUE(c->get_CellFormat()->get_FitText());
        
        ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, c->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
    }
    
    ASPOSE_ASSERT_EQ(150, table->get_Rows()->idx_get(2)->get_RowFormat()->get_Height());
    
    ASSERT_EQ(u"Row 3, Col 1\a", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::TextOrientation::Upward, table->get_Rows()->idx_get(2)->get_Cells()->idx_get(0)->get_CellFormat()->get_Orientation());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, table->get_Rows()->idx_get(2)->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
    
    ASSERT_EQ(u"Row 3, Col 2\a", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::TextOrientation::Downward, table->get_Rows()->idx_get(2)->get_Cells()->idx_get(1)->get_CellFormat()->get_Orientation());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, table->get_Rows()->idx_get(2)->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTable)
{
    s_instance->InsertTable();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTableWithStyle()
{
    //ExStart
    //ExFor:Table.StyleIdentifier
    //ExFor:Table.StyleOptions
    //ExFor:TableStyleOptions
    //ExFor:Table.AutoFit
    //ExFor:AutoFitBehavior
    //ExSummary:Shows how to build a new table while applying a style.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    
    // We must insert at least one row before setting any table formatting.
    builder->InsertCell();
    
    // Set the table style used based on the style identifier.
    // Note that not all table styles are available when saving to .doc format.
    table->set_StyleIdentifier(Aspose::Words::StyleIdentifier::MediumShading1Accent1);
    
    // Partially apply the style to features of the table based on predicates, then build the table.
    table->set_StyleOptions(Aspose::Words::Tables::TableStyleOptions::FirstColumn | Aspose::Words::Tables::TableStyleOptions::RowBands | Aspose::Words::Tables::TableStyleOptions::FirstRow);
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::AutoFitToContents);
    
    builder->Writeln(u"Item");
    builder->get_CellFormat()->set_RightPadding(40);
    builder->InsertCell();
    builder->Writeln(u"Quantity (kg)");
    builder->EndRow();
    
    builder->InsertCell();
    builder->Writeln(u"Apples");
    builder->InsertCell();
    builder->Writeln(u"20");
    builder->EndRow();
    
    builder->InsertCell();
    builder->Writeln(u"Bananas");
    builder->InsertCell();
    builder->Writeln(u"40");
    builder->EndRow();
    
    builder->InsertCell();
    builder->Writeln(u"Carrots");
    builder->InsertCell();
    builder->Writeln(u"50");
    builder->EndRow();
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertTableWithStyle.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertTableWithStyle.docx");
    
    doc->ExpandTableStylesToDirectFormatting();
    
    ASSERT_EQ(u"Medium Shading 1 Accent 1", table->get_Style()->get_Name());
    ASSERT_EQ(Aspose::Words::Tables::TableStyleOptions::FirstColumn | Aspose::Words::Tables::TableStyleOptions::RowBands | Aspose::Words::Tables::TableStyleOptions::FirstRow, table->get_StyleOptions());
    ASSERT_EQ(189, ASPOSECPP_CHECKED_CAST(int32_t, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().get_B()));
    ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Color().ToArgb());
    ASSERT_NE(System::Drawing::Color::get_LightBlue().ToArgb(), ASPOSECPP_CHECKED_CAST(int32_t, table->get_LastRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().get_B()));
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), table->get_LastRow()->get_FirstCell()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Color().ToArgb());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTableWithStyle)
{
    s_instance->InsertTableWithStyle();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTableSetHeadingRow()
{
    //ExStart
    //ExFor:RowFormat.HeadingFormat
    //ExSummary:Shows how to build a table with rows that repeat on every page.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    
    // Any rows inserted while the "HeadingFormat" flag is set to "true"
    // will show up at the top of the table on every page that it spans.
    builder->get_RowFormat()->set_HeadingFormat(true);
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->get_CellFormat()->set_Width(100);
    builder->InsertCell();
    builder->Write(u"Heading row 1");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Heading row 2");
    builder->EndRow();
    
    builder->get_CellFormat()->set_Width(50);
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->get_RowFormat()->set_HeadingFormat(false);
    
    // Add enough rows for the table to span two pages.
    for (int32_t i = 0; i < 50; i++)
    {
        builder->InsertCell();
        builder->Write(System::String::Format(u"Row {0}, column 1.", table->get_Rows()->get_Count()));
        builder->InsertCell();
        builder->Write(System::String::Format(u"Row {0}, column 2.", table->get_Rows()->get_Count()));
        builder->EndRow();
    }
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertTableSetHeadingRow.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertTableSetHeadingRow.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    for (int32_t i = 0; i < table->get_Rows()->get_Count(); i++)
    {
        ASPOSE_ASSERT_EQ(i < 2, table->get_Rows()->idx_get(i)->get_RowFormat()->get_HeadingFormat());
    }
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTableSetHeadingRow)
{
    s_instance->InsertTableSetHeadingRow();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTableWithPreferredWidth()
{
    //ExStart
    //ExFor:Table.PreferredWidth
    //ExFor:PreferredWidth.FromPercent
    //ExFor:PreferredWidth
    //ExSummary:Shows how to set a table to auto fit to 50% of the width of the page.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Cell #1");
    builder->InsertCell();
    builder->Write(u"Cell #2");
    builder->InsertCell();
    builder->Write(u"Cell #3");
    
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPercent(50));
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertTableWithPreferredWidth.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertTableWithPreferredWidth.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Percent, table->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(50, table->get_PreferredWidth()->get_Value());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTableWithPreferredWidth)
{
    s_instance->InsertTableWithPreferredWidth();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertCellsWithPreferredWidths()
{
    //ExStart
    //ExFor:CellFormat.PreferredWidth
    //ExFor:PreferredWidth
    //ExFor:PreferredWidth.Auto
    //ExFor:PreferredWidth.Equals(PreferredWidth)
    //ExFor:PreferredWidth.Equals(Object)
    //ExFor:PreferredWidth.FromPoints
    //ExFor:PreferredWidth.FromPercent
    //ExFor:PreferredWidth.GetHashCode
    //ExFor:PreferredWidth.ToString
    //ExSummary:Shows how to set a preferred width for table cells.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    
    // There are two ways of applying the "PreferredWidth" class to table cells.
    // 1 -  Set an absolute preferred width based on points:
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(40));
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightYellow());
    builder->Writeln(System::String::Format(u"Cell with a width of {0}.", builder->get_CellFormat()->get_PreferredWidth()));
    
    // 2 -  Set a relative preferred width based on percent of the table's width:
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPercent(20));
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
    builder->Writeln(System::String::Format(u"Cell with a width of {0}.", builder->get_CellFormat()->get_PreferredWidth()));
    
    builder->InsertCell();
    
    // A cell with no preferred width specified will take up the rest of the available space.
    builder->get_CellFormat()->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::Auto());
    
    // Each configuration of the "PreferredWidth" property creates a new object.
    ASSERT_NE(System::ObjectExt::GetHashCode(table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()), System::ObjectExt::GetHashCode(builder->get_CellFormat()->get_PreferredWidth()));
    
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightGreen());
    builder->Writeln(u"Automatically sized cell.");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertCellsWithPreferredWidths.docx");
    //ExEnd
    
    ASPOSE_ASSERT_EQ(100.0, Aspose::Words::Tables::PreferredWidth::FromPercent(100)->get_Value());
    ASPOSE_ASSERT_EQ(100.0, Aspose::Words::Tables::PreferredWidth::FromPoints(100)->get_Value());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertCellsWithPreferredWidths.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Points, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(40.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_PreferredWidth()->get_Value());
    ASSERT_EQ(u"Cell with a width of 800.\r\a", table->get_FirstRow()->get_Cells()->idx_get(0)->GetText().Trim());
    
    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Percent, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(20.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()->get_Value());
    ASSERT_EQ(u"Cell with a width of 20%.\r\a", table->get_FirstRow()->get_Cells()->idx_get(1)->GetText().Trim());
    
    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Auto, table->get_FirstRow()->get_Cells()->idx_get(2)->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(2)->get_CellFormat()->get_PreferredWidth()->get_Value());
    ASSERT_EQ(u"Automatically sized cell.\r\a", table->get_FirstRow()->get_Cells()->idx_get(2)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertCellsWithPreferredWidths)
{
    s_instance->InsertCellsWithPreferredWidths();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTableFromHtml()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
    // inserted from HTML.
    builder->InsertHtml(System::String(u"<table>") + u"<tr>" + u"<td>Row 1, Cell 1</td>" + u"<td>Row 1, Cell 2</td>" + u"</tr>" + u"<tr>" + u"<td>Row 2, Cell 2</td>" + u"<td>Row 2, Cell 2</td>" + u"</tr>" + u"</table>");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertTableFromHtml.docx");
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertTableFromHtml.docx");
    
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Row, true)->get_Count());
    ASSERT_EQ(4, doc->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTableFromHtml)
{
    s_instance->InsertTableFromHtml();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertNestedTable()
{
    //ExStart
    //ExFor:Cell.FirstParagraph
    //ExSummary:Shows how to create a nested table using a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Build the outer table.
    System::SharedPtr<Aspose::Words::Tables::Cell> cell = builder->InsertCell();
    builder->Writeln(u"Outer Table Cell 1");
    builder->InsertCell();
    builder->Writeln(u"Outer Table Cell 2");
    builder->EndTable();
    
    // Move to the first cell of the outer table, the build another table inside the cell.
    builder->MoveTo(cell->get_FirstParagraph());
    builder->InsertCell();
    builder->Writeln(u"Inner Table Cell 1");
    builder->InsertCell();
    builder->Writeln(u"Inner Table Cell 2");
    builder->EndTable();
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertNestedTable.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertNestedTable.docx");
    
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(4, doc->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    ASSERT_EQ(1, cell->get_Tables()->idx_get(0)->get_Count());
    ASSERT_EQ(2, cell->get_Tables()->idx_get(0)->get_FirstRow()->get_Cells()->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertNestedTable)
{
    s_instance->InsertNestedTable();
}

} // namespace gtest_test

void ExDocumentBuilder::CreateTable()
{
    //ExStart
    //ExFor:DocumentBuilder
    //ExFor:DocumentBuilder.Write(String)
    //ExFor:DocumentBuilder.InsertCell
    //ExSummary:Shows how to use a document builder to create a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Start the table, then populate the first row with two cells.
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 2.");
    
    // Call the builder's "EndRow" method to start a new row.
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Row 2, Cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 2, Cell 2.");
    builder->EndTable();
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.CreateTable.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.CreateTable.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(4, table->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    
    ASSERT_EQ(u"Row 1, Cell 1.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u"Row 1, Cell 2.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());
    ASSERT_EQ(u"Row 2, Cell 1.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u"Row 2, Cell 2.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, CreateTable)
{
    s_instance->CreateTable();
}

} // namespace gtest_test

void ExDocumentBuilder::BuildFormattedTable()
{
    //ExStart
    //ExFor:RowFormat.Height
    //ExFor:RowFormat.HeightRule
    //ExFor:Table.LeftIndent
    //ExFor:DocumentBuilder.ParagraphFormat
    //ExFor:DocumentBuilder.Font
    //ExSummary:Shows how to create a formatted table using DocumentBuilder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    table->set_LeftIndent(20);
    
    // Set some formatting options for text and table appearance.
    builder->get_RowFormat()->set_Height(40);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::AtLeast);
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::FromArgb(198, 217, 241));
    
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->get_Font()->set_Size(16);
    builder->get_Font()->set_Name(u"Arial");
    builder->get_Font()->set_Bold(true);
    
    // Configuring the formatting options in a document builder will apply them
    // to the current cell/row its cursor is in,
    // as well as any new cells and rows created using that builder.
    builder->Write(u"Header Row,\n Cell 1");
    builder->InsertCell();
    builder->Write(u"Header Row,\n Cell 2");
    builder->InsertCell();
    builder->Write(u"Header Row,\n Cell 3");
    builder->EndRow();
    
    // Reconfigure the builder's formatting objects for new rows and cells that we are about to make.
    // The builder will not apply these to the first row already created so that it will stand out as a header row.
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_White());
    builder->get_CellFormat()->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);
    builder->get_RowFormat()->set_Height(30);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Auto);
    builder->InsertCell();
    builder->get_Font()->set_Size(12);
    builder->get_Font()->set_Bold(false);
    
    builder->Write(u"Row 1, Cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 2.");
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 3.");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Row 2, Cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 2, Cell 2.");
    builder->InsertCell();
    builder->Write(u"Row 2, Cell 3.");
    builder->EndRow();
    builder->EndTable();
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.CreateFormattedTable.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.CreateFormattedTable.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASPOSE_ASSERT_EQ(20.0, table->get_LeftIndent());
    
    ASSERT_EQ(Aspose::Words::HeightRule::AtLeast, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    ASPOSE_ASSERT_EQ(40.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    
    for (auto&& c : System::IterateOver<Aspose::Words::Tables::Cell>(doc->GetChildNodes(Aspose::Words::NodeType::Cell, true)))
    {
        ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, c->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
        
        for (auto&& r : System::IterateOver<Aspose::Words::Run>(c->get_FirstParagraph()->get_Runs()))
        {
            ASSERT_EQ(u"Arial", r->get_Font()->get_Name());
            
            if (c->get_ParentRow() == table->get_FirstRow())
            {
                ASPOSE_ASSERT_EQ(16, r->get_Font()->get_Size());
                ASSERT_TRUE(r->get_Font()->get_Bold());
            }
            else
            {
                ASPOSE_ASSERT_EQ(12, r->get_Font()->get_Size());
                ASSERT_FALSE(r->get_Font()->get_Bold());
            }
        }
    }
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, BuildFormattedTable)
{
    s_instance->BuildFormattedTable();
}

} // namespace gtest_test

void ExDocumentBuilder::TableBordersAndShading()
{
    //ExStart
    //ExFor:Shading
    //ExFor:Table.SetBorders
    //ExFor:BorderCollection.Left
    //ExFor:BorderCollection.Right
    //ExFor:BorderCollection.Top
    //ExFor:BorderCollection.Bottom
    //ExSummary:Shows how to apply border and shading color while building a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Start a table and set a default color/thickness for its borders.
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    table->SetBorders(Aspose::Words::LineStyle::Single, 2.0, System::Drawing::Color::get_Black());
    
    // Create a row with two cells with different background colors.
    builder->InsertCell();
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightSkyBlue());
    builder->Writeln(u"Row 1, Cell 1.");
    builder->InsertCell();
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Orange());
    builder->Writeln(u"Row 1, Cell 2.");
    builder->EndRow();
    
    // Reset cell formatting to disable the background colors
    // set a custom border thickness for all new cells created by the builder,
    // then build a second row.
    builder->get_CellFormat()->ClearFormatting();
    builder->get_CellFormat()->get_Borders()->get_Left()->set_LineWidth(4.0);
    builder->get_CellFormat()->get_Borders()->get_Right()->set_LineWidth(4.0);
    builder->get_CellFormat()->get_Borders()->get_Top()->set_LineWidth(4.0);
    builder->get_CellFormat()->get_Borders()->get_Bottom()->set_LineWidth(4.0);
    
    builder->InsertCell();
    builder->Writeln(u"Row 2, Cell 1.");
    builder->InsertCell();
    builder->Writeln(u"Row 2, Cell 2.");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.TableBordersAndShading.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.TableBordersAndShading.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    for (auto&& c : System::IterateOver<Aspose::Words::Tables::Cell>(table->get_FirstRow()))
    {
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Top()->get_LineWidth());
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Bottom()->get_LineWidth());
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Left()->get_LineWidth());
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Right()->get_LineWidth());
        
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::Single, c->get_CellFormat()->get_Borders()->get_Left()->get_LineStyle());
    }
    
    ASSERT_EQ(System::Drawing::Color::get_LightSkyBlue().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
    
    for (auto&& c : System::IterateOver<Aspose::Words::Tables::Cell>(table->get_LastRow()))
    {
        ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Top()->get_LineWidth());
        ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Bottom()->get_LineWidth());
        ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Left()->get_LineWidth());
        ASPOSE_ASSERT_EQ(4.0, c->get_CellFormat()->get_Borders()->get_Right()->get_LineWidth());
        
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::Single, c->get_CellFormat()->get_Borders()->get_Left()->get_LineStyle());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
    }
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, TableBordersAndShading)
{
    s_instance->TableBordersAndShading();
}

} // namespace gtest_test

void ExDocumentBuilder::SetPreferredTypeConvertUtil()
{
    //ExStart
    //ExFor:PreferredWidth.FromPoints
    //ExSummary:Shows how to use unit conversion tools while specifying a preferred width for a cell.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->get_CellFormat()->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(Aspose::Words::ConvertUtil::InchToPoint(3)));
    builder->InsertCell();
    
    ASPOSE_ASSERT_EQ(216.0, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_PreferredWidth()->get_Value());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SetPreferredTypeConvertUtil)
{
    s_instance->SetPreferredTypeConvertUtil();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertHyperlinkToLocalBookmark()
{
    //ExStart
    //ExFor:DocumentBuilder.StartBookmark
    //ExFor:DocumentBuilder.EndBookmark
    //ExFor:DocumentBuilder.InsertHyperlink
    //ExSummary:Shows how to insert a hyperlink which references a local bookmark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartBookmark(u"Bookmark1");
    builder->Write(u"Bookmarked text. ");
    builder->EndBookmark(u"Bookmark1");
    builder->Writeln(u"Text outside of the bookmark.");
    
    // Insert a HYPERLINK field that links to the bookmark. We can pass field switches
    // to the "InsertHyperlink" method as part of the argument containing the referenced bookmark's name.
    builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    builder->get_Font()->set_Underline(Aspose::Words::Underline::Single);
    auto hyperlink = System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(builder->InsertHyperlink(u"Link to Bookmark1", u"Bookmark1", true));
    hyperlink->set_ScreenTip(u"Hyperlink Tip");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
    hyperlink = System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldHyperlink, u" HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ", u"Link to Bookmark1", hyperlink);
    ASSERT_EQ(u"Bookmark1", hyperlink->get_SubAddress());
    ASSERT_EQ(u"Hyperlink Tip", hyperlink->get_ScreenTip());
    ASSERT_TRUE(doc->get_Range()->get_Bookmarks()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Bookmark>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Bookmark> b)>>([](System::SharedPtr<Aspose::Words::Bookmark> b) -> bool
    {
        return b->get_Name() == u"Bookmark1";
    }))));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertHyperlinkToLocalBookmark)
{
    s_instance->InsertHyperlinkToLocalBookmark();
}

} // namespace gtest_test

void ExDocumentBuilder::CursorPosition()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"Hello world!");
    
    // If the builder's cursor is at the end of the document,
    // there will be no nodes in front of it so that the current node will be null.
    ASSERT_TRUE(System::TestTools::IsNull(builder->get_CurrentNode()));
    
    ASSERT_EQ(u"Hello world!", builder->get_CurrentParagraph()->GetText().Trim());
    
    // Move to the beginning of the document and place the cursor at an existing node.
    builder->MoveToDocumentStart();
    ASSERT_EQ(Aspose::Words::NodeType::Run, builder->get_CurrentNode()->get_NodeType());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, CursorPosition)
{
    s_instance->CursorPosition();
}

} // namespace gtest_test

void ExDocumentBuilder::MoveTo()
{
    //ExStart
    //ExFor:Story.LastParagraph
    //ExFor:DocumentBuilder.MoveTo(Node)
    //ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Run 1. ");
    
    // The document builder has a cursor, which acts as the part of the document
    // where the builder appends new nodes when we use its document construction methods.
    // This cursor functions in the same way as Microsoft Word's blinking cursor,
    // and it also always ends up immediately after any node that the builder just inserted.
    // To append content to a different part of the document,
    // we can move the cursor to a different node with the "MoveTo" method.
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_LastParagraph(), builder->get_CurrentParagraph());
    //ExSkip
    builder->MoveTo(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0));
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), builder->get_CurrentParagraph());
    //ExSkip
    
    // The cursor is now in front of the node that we moved it to.
    // Adding a second run will insert it in front of the first run.
    builder->Writeln(u"Run 2. ");
    
    ASSERT_EQ(u"Run 2. \rRun 1.", doc->GetText().Trim());
    
    // Move the cursor to the end of the document to continue appending text to the end as before.
    builder->MoveTo(doc->get_LastSection()->get_Body()->get_LastParagraph());
    builder->Writeln(u"Run 3. ");
    
    ASSERT_EQ(u"Run 2. \rRun 1. \rRun 3.", doc->GetText().Trim());
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_LastParagraph(), builder->get_CurrentParagraph());
    //ExSkip
    
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MoveTo)
{
    s_instance->MoveTo();
}

} // namespace gtest_test

void ExDocumentBuilder::MoveToParagraph()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToParagraph
    //ExSummary:Shows how to move a builder's cursor position to a specified paragraph.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    ASSERT_EQ(22, paragraphs->get_Count());
    
    // Create document builder to edit the document. The builder's cursor,
    // which is the point where it will insert new nodes when we call its document construction methods,
    // is currently at the beginning of the document.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    ASSERT_EQ(0, paragraphs->IndexOf(builder->get_CurrentParagraph()));
    
    // Move that cursor to a different paragraph will place that cursor in front of that paragraph.
    builder->MoveToParagraph(2, 0);
    ASSERT_EQ(2, paragraphs->IndexOf(builder->get_CurrentParagraph()));
    //ExSkip
    
    // Any new content that we add will be inserted at that point.
    builder->Writeln(u"This is a new third paragraph. ");
    //ExEnd
    
    ASSERT_EQ(3, paragraphs->IndexOf(builder->get_CurrentParagraph()));
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(u"This is a new third paragraph.", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(2)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MoveToParagraph)
{
    s_instance->MoveToParagraph();
}

} // namespace gtest_test

void ExDocumentBuilder::MoveToCell()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToCell
    //ExSummary:Shows how to move a document builder's cursor to a cell in a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an empty 2x2 table.
    builder->StartTable();
    builder->InsertCell();
    builder->InsertCell();
    builder->EndRow();
    builder->InsertCell();
    builder->InsertCell();
    builder->EndTable();
    
    // Because we have ended the table with the EndTable method,
    // the document builder's cursor is currently outside the table.
    // This cursor has the same function as Microsoft Word's blinking text cursor.
    // It can also be moved to a different location in the document using the builder's MoveTo methods.
    // We can move the cursor back inside the table to a specific cell.
    builder->MoveToCell(0, 1, 1, 0);
    builder->Write(u"Column 2, cell 2.");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.MoveToCell.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MoveToCell.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(u"Column 2, cell 2.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MoveToCell)
{
    s_instance->MoveToCell();
}

} // namespace gtest_test

void ExDocumentBuilder::MoveToBookmark()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
    //ExSummary:Shows how to move a document builder's node insertion point cursor to a bookmark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A valid bookmark consists of a BookmarkStart node, a BookmarkEnd node with a
    // matching bookmark name somewhere afterward, and contents enclosed by those nodes.
    builder->StartBookmark(u"MyBookmark");
    builder->Write(u"Hello world! ");
    builder->EndBookmark(u"MyBookmark");
    
    // There are 4 ways of moving a document builder's cursor to a bookmark.
    // If we are between the BookmarkStart and BookmarkEnd nodes, the cursor will be inside the bookmark.
    // This means that any text added by the builder will become a part of the bookmark.
    // 1 -  Outside of the bookmark, in front of the BookmarkStart node:
    ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", true, false));
    builder->Write(u"1. ");
    
    ASSERT_EQ(u"Hello world! ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
    ASSERT_EQ(u"1. Hello world!", doc->GetText().Trim());
    
    // 2 -  Inside the bookmark, right after the BookmarkStart node:
    ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", true, true));
    builder->Write(u"2. ");
    
    ASSERT_EQ(u"2. Hello world! ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
    ASSERT_EQ(u"1. 2. Hello world!", doc->GetText().Trim());
    
    // 2 -  Inside the bookmark, right in front of the BookmarkEnd node:
    ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", false, false));
    builder->Write(u"3. ");
    
    ASSERT_EQ(u"2. Hello world! 3. ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
    ASSERT_EQ(u"1. 2. Hello world! 3.", doc->GetText().Trim());
    
    // 4 -  Outside of the bookmark, after the BookmarkEnd node:
    ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark", false, true));
    builder->Write(u"4.");
    
    ASSERT_EQ(u"2. Hello world! 3. ", doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark")->get_Text());
    ASSERT_EQ(u"1. 2. Hello world! 3. 4.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MoveToBookmark)
{
    s_instance->MoveToBookmark();
}

} // namespace gtest_test

void ExDocumentBuilder::BuildTable()
{
    //ExStart
    //ExFor:Table
    //ExFor:DocumentBuilder.StartTable
    //ExFor:DocumentBuilder.EndRow
    //ExFor:DocumentBuilder.EndTable
    //ExFor:DocumentBuilder.CellFormat
    //ExFor:DocumentBuilder.RowFormat
    //ExFor:DocumentBuilder.Write(String)
    //ExFor:DocumentBuilder.Writeln(String)
    //ExFor:CellVerticalAlignment
    //ExFor:CellFormat.Orientation
    //ExFor:TextOrientation
    //ExFor:AutoFitBehavior
    //ExSummary:Shows how to build a formatted 2x2 table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);
    builder->Write(u"Row 1, cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 1, cell 2.");
    builder->EndRow();
    
    // While building the table, the document builder will apply its current RowFormat/CellFormat property values
    // to the current row/cell that its cursor is in and any new rows/cells as it creates them.
    ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalAlignment());
    ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->get_CellFormat()->get_VerticalAlignment());
    
    builder->InsertCell();
    builder->get_RowFormat()->set_Height(100);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Upward);
    builder->Write(u"Row 2, cell 1.");
    builder->InsertCell();
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Downward);
    builder->Write(u"Row 2, cell 2.");
    builder->EndRow();
    builder->EndTable();
    
    // Previously added rows and cells are not retroactively affected by changes to the builder's formatting.
    ASPOSE_ASSERT_EQ(0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    ASPOSE_ASSERT_EQ(100, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());
    ASSERT_EQ(Aspose::Words::TextOrientation::Upward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->get_CellFormat()->get_Orientation());
    ASSERT_EQ(Aspose::Words::TextOrientation::Downward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->get_CellFormat()->get_Orientation());
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.BuildTable.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.BuildTable.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(2, table->get_Rows()->get_Count());
    ASSERT_EQ(2, table->get_Rows()->idx_get(0)->get_Cells()->get_Count());
    ASSERT_EQ(2, table->get_Rows()->idx_get(1)->get_Cells()->get_Count());
    
    ASPOSE_ASSERT_EQ(0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    ASPOSE_ASSERT_EQ(100, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());
    
    ASSERT_EQ(u"Row 1, cell 1.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalAlignment());
    
    ASSERT_EQ(u"Row 1, cell 2.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());
    
    ASSERT_EQ(u"Row 2, cell 1.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::TextOrientation::Upward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->get_CellFormat()->get_Orientation());
    
    ASSERT_EQ(u"Row 2, cell 2.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::TextOrientation::Downward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->get_CellFormat()->get_Orientation());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, BuildTable)
{
    s_instance->BuildTable();
}

} // namespace gtest_test

void ExDocumentBuilder::TableCellVerticalRotatedFarEastTextOrientation()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rotated cell text.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::SharedPtr<Aspose::Words::Tables::Cell> cell = table->get_FirstRow()->get_FirstCell();
    
    ASSERT_EQ(Aspose::Words::TextOrientation::VerticalRotatedFarEast, cell->get_CellFormat()->get_Orientation());
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    cell = table->get_FirstRow()->get_FirstCell();
    
    ASSERT_EQ(Aspose::Words::TextOrientation::VerticalRotatedFarEast, cell->get_CellFormat()->get_Orientation());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, TableCellVerticalRotatedFarEastTextOrientation)
{
    s_instance->TableCellVerticalRotatedFarEastTextOrientation();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertFloatingImage()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // There are two ways of using a document builder to source an image and then insert it as a floating shape.
    // 1 -  From a file in the local file system:
    builder->InsertImage(get_ImageDir() + u"Transparent background logo.png", Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 0.0, 200.0, 200.0, Aspose::Words::Drawing::WrapType::Square);
    
    // 2 -  From a URL:
    builder->InsertImage(get_ImageUrl(), Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 250.0, 200.0, 200.0, Aspose::Words::Drawing::WrapType::Square);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertFloatingImage.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertFloatingImage.docx");
    auto image = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, image);
    ASPOSE_ASSERT_EQ(100.0, image->get_Left());
    ASPOSE_ASSERT_EQ(0.0, image->get_Top());
    ASPOSE_ASSERT_EQ(200.0, image->get_Width());
    ASPOSE_ASSERT_EQ(200.0, image->get_Height());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, image->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, image->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, image->get_RelativeVerticalPosition());
    
    image = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(272, 92, Aspose::Words::Drawing::ImageType::Png, image);
    ASPOSE_ASSERT_EQ(100.0, image->get_Left());
    ASPOSE_ASSERT_EQ(250.0, image->get_Top());
    ASPOSE_ASSERT_EQ(200.0, image->get_Width());
    ASPOSE_ASSERT_EQ(200.0, image->get_Height());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, image->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, image->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, image->get_RelativeVerticalPosition());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertFloatingImage)
{
    s_instance->InsertFloatingImage();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertImageOriginalSize()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from the local file system into a document while preserving its dimensions.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // The InsertImage method creates a floating shape with the passed image in its image data.
    // We can specify the dimensions of the shape can be passing them to this method.
    System::SharedPtr<Aspose::Words::Drawing::Shape> imageShape = builder->InsertImage(get_ImageDir() + u"Logo.jpg", Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 0.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 0.0, -1.0, -1.0, Aspose::Words::Drawing::WrapType::Square);
    
    // Passing negative values as the intended dimensions will automatically define
    // the shape's dimensions based on the dimensions of its image.
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertImageOriginalSize.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertImageOriginalSize.docx");
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertImageOriginalSize)
{
    s_instance->InsertImageOriginalSize();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTextInput()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertTextInput
    //ExSummary:Shows how to insert a text input form field into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a form that prompts the user to enter text.
    builder->InsertTextInput(u"TextInput", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Enter your text here", 0);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertTextInput.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertTextInput.docx");
    System::SharedPtr<Aspose::Words::Fields::FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);
    
    ASSERT_TRUE(formField->get_Enabled());
    ASSERT_EQ(u"TextInput", formField->get_Name());
    ASSERT_EQ(0, formField->get_MaxLength());
    ASSERT_EQ(u"Enter your text here", formField->get_Result());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormTextInput, formField->get_Type());
    ASSERT_EQ(u"", formField->get_TextInputFormat());
    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, formField->get_TextInputType());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTextInput)
{
    s_instance->InsertTextInput();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertComboBox()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertComboBox
    //ExSummary:Shows how to insert a combo box form field into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a form that prompts the user to pick one of the items from the menu.
    builder->Write(u"Pick a fruit: ");
    System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"Apple", u"Banana", u"Cherry"});
    builder->InsertComboBox(u"DropDown", items, 0);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertComboBox.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertComboBox.docx");
    System::SharedPtr<Aspose::Words::Fields::FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);
    
    ASSERT_TRUE(formField->get_Enabled());
    ASSERT_EQ(u"DropDown", formField->get_Name());
    ASSERT_EQ(0, formField->get_DropDownSelectedIndex());
    ASPOSE_ASSERT_EQ(items, formField->get_DropDownItems());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormDropDown, formField->get_Type());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertComboBox)
{
    s_instance->InsertComboBox();
}

} // namespace gtest_test

void ExDocumentBuilder::SignatureLineProviderId()
{
    //ExStart
    //ExFor:SignatureLine.IsSigned
    //ExFor:SignatureLine.IsValid
    //ExFor:SignatureLine.ProviderId
    //ExFor:SignatureLineOptions
    //ExFor:SignatureLineOptions.ShowDate
    //ExFor:SignatureLineOptions.Email
    //ExFor:SignatureLineOptions.DefaultInstructions
    //ExFor:SignatureLineOptions.Instructions
    //ExFor:SignatureLineOptions.AllowComments
    //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
    //ExFor:SignOptions.ProviderId
    //ExSummary:Shows how to sign a document with a personal certificate and a signature line.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto signatureLineOptions = System::MakeObject<Aspose::Words::SignatureLineOptions>();
    signatureLineOptions->set_Signer(u"vderyushev");
    signatureLineOptions->set_SignerTitle(u"QA");
    signatureLineOptions->set_Email(u"vderyushev@aspose.com");
    signatureLineOptions->set_ShowDate(true);
    signatureLineOptions->set_DefaultInstructions(false);
    signatureLineOptions->set_Instructions(u"Please sign here.");
    signatureLineOptions->set_AllowComments(true);
    
    System::SharedPtr<Aspose::Words::Drawing::SignatureLine> signatureLine = builder->InsertSignatureLine(signatureLineOptions)->get_SignatureLine();
    signatureLine->set_ProviderId(System::Guid::Parse(u"CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));
    
    ASSERT_FALSE(signatureLine->get_IsSigned());
    ASSERT_FALSE(signatureLine->get_IsValid());
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.SignatureLineProviderId.docx");
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_SignatureLineId(signatureLine->get_Id());
    signOptions->set_ProviderId(signatureLine->get_ProviderId());
    signOptions->set_Comments(u"Document was signed by vderyushev");
    signOptions->set_SignTime(System::DateTime::get_Now());
    
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(get_ArtifactsDir() + u"DocumentBuilder.SignatureLineProviderId.docx", get_ArtifactsDir() + u"DocumentBuilder.SignatureLineProviderId.Signed.docx", certHolder, signOptions);
    
    // Re-open our saved document, and verify that the "IsSigned" and "IsValid" properties both equal "true",
    // indicating that the signature line contains a signature.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.SignatureLineProviderId.Signed.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    signatureLine = shape->get_SignatureLine();
    
    ASSERT_TRUE(signatureLine->get_IsSigned());
    ASSERT_TRUE(signatureLine->get_IsValid());
    //ExEnd
    
    ASSERT_EQ(u"vderyushev", signatureLine->get_Signer());
    ASSERT_EQ(u"QA", signatureLine->get_SignerTitle());
    ASSERT_EQ(u"vderyushev@aspose.com", signatureLine->get_Email());
    ASSERT_TRUE(signatureLine->get_ShowDate());
    ASSERT_FALSE(signatureLine->get_DefaultInstructions());
    ASSERT_EQ(u"Please sign here.", signatureLine->get_Instructions());
    ASSERT_TRUE(signatureLine->get_AllowComments());
    ASSERT_TRUE(signatureLine->get_IsSigned());
    ASSERT_TRUE(signatureLine->get_IsValid());
    
    System::SharedPtr<Aspose::Words::DigitalSignatures::DigitalSignatureCollection> signatures = Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(get_ArtifactsDir() + u"DocumentBuilder.SignatureLineProviderId.Signed.docx");
    
    ASSERT_EQ(1, signatures->get_Count());
    ASSERT_TRUE(signatures->idx_get(0)->get_IsValid());
    ASSERT_EQ(u"Document was signed by vderyushev", signatures->idx_get(0)->get_Comments());
    ASSERT_EQ(System::DateTime::get_Today(), signatures->idx_get(0)->get_SignTime().get_Date());
    ASSERT_EQ(u"CN=Morzal.Me", signatures->idx_get(0)->get_IssuerName());
    ASSERT_EQ(Aspose::Words::DigitalSignatures::DigitalSignatureType::XmlDsig, signatures->idx_get(0)->get_SignatureType());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SignatureLineProviderId)
{
    s_instance->SignatureLineProviderId();
}

} // namespace gtest_test

void ExDocumentBuilder::SignatureLineInline()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
    //ExSummary:Shows how to insert an inline signature line into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto options = System::MakeObject<Aspose::Words::SignatureLineOptions>();
    options->set_Signer(u"John Doe");
    options->set_SignerTitle(u"Manager");
    options->set_Email(u"johndoe@aspose.com");
    options->set_ShowDate(true);
    options->set_DefaultInstructions(false);
    options->set_Instructions(u"Please sign here.");
    options->set_AllowComments(true);
    
    builder->InsertSignatureLine(options, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, 2.0, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 3.0, Aspose::Words::Drawing::WrapType::Inline);
    
    // The signature line can be signed in Microsoft Word by double clicking it.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.SignatureLineInline.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.SignatureLineInline.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::SignatureLine> signatureLine = shape->get_SignatureLine();
    
    ASSERT_EQ(u"John Doe", signatureLine->get_Signer());
    ASSERT_EQ(u"Manager", signatureLine->get_SignerTitle());
    ASSERT_EQ(u"johndoe@aspose.com", signatureLine->get_Email());
    ASSERT_TRUE(signatureLine->get_ShowDate());
    ASSERT_FALSE(signatureLine->get_DefaultInstructions());
    ASSERT_EQ(u"Please sign here.", signatureLine->get_Instructions());
    ASSERT_TRUE(signatureLine->get_AllowComments());
    ASSERT_FALSE(signatureLine->get_IsSigned());
    ASSERT_FALSE(signatureLine->get_IsValid());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SignatureLineInline)
{
    s_instance->SignatureLineInline();
}

} // namespace gtest_test

void ExDocumentBuilder::SetParagraphFormatting()
{
    //ExStart
    //ExFor:ParagraphFormat.RightIndent
    //ExFor:ParagraphFormat.LeftIndent
    //ExSummary:Shows how to configure paragraph formatting to create off-center text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Center all text that the document builder writes, and set up indents.
    // The indent configuration below will create a body of text that will sit asymmetrically on the page.
    // The "center" that we align the text to will be the middle of the body of text, not the middle of the page.
    System::SharedPtr<Aspose::Words::ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
    paragraphFormat->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    paragraphFormat->set_LeftIndent(100);
    paragraphFormat->set_RightIndent(50);
    paragraphFormat->set_SpaceAfter(25);
    
    builder->Writeln(u"This paragraph demonstrates how left and right indentation affects word wrapping.");
    builder->Writeln(u"The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.SetParagraphFormatting.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.SetParagraphFormatting.docx");
    
    for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
    {
        ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, paragraph->get_ParagraphFormat()->get_Alignment());
        ASPOSE_ASSERT_EQ(100.0, paragraph->get_ParagraphFormat()->get_LeftIndent());
        ASPOSE_ASSERT_EQ(50.0, paragraph->get_ParagraphFormat()->get_RightIndent());
        ASPOSE_ASSERT_EQ(25.0, paragraph->get_ParagraphFormat()->get_SpaceAfter());
    }
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SetParagraphFormatting)
{
    s_instance->SetParagraphFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::SetCellFormatting()
{
    //ExStart
    //ExFor:DocumentBuilder.CellFormat
    //ExFor:CellFormat.Width
    //ExFor:CellFormat.LeftPadding
    //ExFor:CellFormat.RightPadding
    //ExFor:CellFormat.TopPadding
    //ExFor:CellFormat.BottomPadding
    //ExFor:DocumentBuilder.StartTable
    //ExFor:DocumentBuilder.EndTable
    //ExSummary:Shows how to format cells with a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, cell 1.");
    
    // Insert a second cell, and then configure cell text padding options.
    // The builder will apply these settings at its current cell, and any new cells creates afterwards.
    builder->InsertCell();
    
    System::SharedPtr<Aspose::Words::Tables::CellFormat> cellFormat = builder->get_CellFormat();
    cellFormat->set_Width(250);
    cellFormat->set_LeftPadding(30);
    cellFormat->set_RightPadding(30);
    cellFormat->set_TopPadding(30);
    cellFormat->set_BottomPadding(30);
    
    builder->Write(u"Row 1, cell 2.");
    builder->EndRow();
    builder->EndTable();
    
    // The first cell was unaffected by the padding reconfiguration, and still holds the default values.
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width());
    ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_LeftPadding());
    ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_RightPadding());
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_TopPadding());
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_BottomPadding());
    
    ASPOSE_ASSERT_EQ(250.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_Width());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_LeftPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_RightPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_TopPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_BottomPadding());
    
    // The first cell will still grow in the output document to match the size of its neighboring cell.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.SetCellFormatting.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.SetCellFormatting.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASPOSE_ASSERT_EQ(159.3, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width());
    ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_LeftPadding());
    ASPOSE_ASSERT_EQ(5.4, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_RightPadding());
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_TopPadding());
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_BottomPadding());
    
    ASPOSE_ASSERT_EQ(310.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_Width());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_LeftPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_RightPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_TopPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_BottomPadding());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SetCellFormatting)
{
    s_instance->SetCellFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::SetRowFormatting()
{
    //ExStart
    //ExFor:DocumentBuilder.RowFormat
    //ExFor:HeightRule
    //ExFor:RowFormat.Height
    //ExFor:RowFormat.HeightRule
    //ExSummary:Shows how to format rows with a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, cell 1.");
    
    // Start a second row, and then configure its height. The builder will apply these settings to
    // its current row, as well as any new rows it creates afterwards.
    builder->EndRow();
    
    System::SharedPtr<Aspose::Words::Tables::RowFormat> rowFormat = builder->get_RowFormat();
    rowFormat->set_Height(100);
    rowFormat->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    
    builder->InsertCell();
    builder->Write(u"Row 2, cell 1.");
    builder->EndTable();
    
    // The first row was unaffected by the padding reconfiguration and still holds the default values.
    ASPOSE_ASSERT_EQ(0.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    
    ASPOSE_ASSERT_EQ(100.0, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.SetRowFormatting.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.SetRowFormatting.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASPOSE_ASSERT_EQ(0.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    
    ASPOSE_ASSERT_EQ(100.0, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SetRowFormatting)
{
    s_instance->SetRowFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertFootnote()
{
    //ExStart
    //ExFor:FootnoteType
    //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String)
    //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String,String)
    //ExSummary:Shows how to reference text with a footnote and an endnote.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert some text and mark it with a footnote with the IsAuto property set to "true" by default,
    // so the marker seen in the body text will be auto-numbered at "1",
    // and the footnote will appear at the bottom of the page.
    builder->Write(u"This text will be referenced by a footnote.");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote comment regarding referenced text.");
    
    // Insert more text and mark it with an endnote with a custom reference mark,
    // which will be used in place of the number "2" and set "IsAuto" to false.
    builder->Write(u"This text will be referenced by an endnote.");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote comment regarding referenced text.", u"CustomMark");
    
    // Footnotes always appear at the bottom of their referenced text,
    // so this page break will not affect the footnote.
    // On the other hand, endnotes are always at the end of the document
    // so that this page break will push the endnote down to the next page.
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertFootnote.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertFootnote.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote comment regarding referenced text.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, false, u"CustomMark", u"CustomMark Endnote comment regarding referenced text.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertFootnote)
{
    s_instance->InsertFootnote();
}

} // namespace gtest_test

void ExDocumentBuilder::ApplyBordersAndShading()
{
    //ExStart
    //ExFor:BorderCollection.Item(BorderType)
    //ExFor:Shading
    //ExFor:TextureIndex
    //ExFor:ParagraphFormat.Shading
    //ExFor:Shading.Texture
    //ExFor:Shading.BackgroundPatternColor
    //ExFor:Shading.ForegroundPatternColor
    //ExSummary:Shows how to decorate text with borders and shading.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
    borders->set_DistanceFromText(20);
    borders->idx_get(Aspose::Words::BorderType::Left)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Right)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Top)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Bottom)->set_LineStyle(Aspose::Words::LineStyle::Double);
    
    System::SharedPtr<Aspose::Words::Shading> shading = builder->get_ParagraphFormat()->get_Shading();
    shading->set_Texture(Aspose::Words::TextureIndex::TextureDiagonalCross);
    shading->set_BackgroundPatternColor(System::Drawing::Color::get_LightCoral());
    shading->set_ForegroundPatternColor(System::Drawing::Color::get_LightSalmon());
    
    builder->Write(u"This paragraph is formatted with a double border and shading.");
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.ApplyBordersAndShading.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.ApplyBordersAndShading.docx");
    borders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();
    
    ASPOSE_ASSERT_EQ(20.0, borders->get_DistanceFromText());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, borders->idx_get(Aspose::Words::BorderType::Left)->get_LineStyle());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, borders->idx_get(Aspose::Words::BorderType::Right)->get_LineStyle());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, borders->idx_get(Aspose::Words::BorderType::Top)->get_LineStyle());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, borders->idx_get(Aspose::Words::BorderType::Bottom)->get_LineStyle());
    
    ASSERT_EQ(Aspose::Words::TextureIndex::TextureDiagonalCross, shading->get_Texture());
    ASSERT_EQ(System::Drawing::Color::get_LightCoral().ToArgb(), shading->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightSalmon().ToArgb(), shading->get_ForegroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, ApplyBordersAndShading)
{
    s_instance->ApplyBordersAndShading();
}

} // namespace gtest_test

void ExDocumentBuilder::DeleteRow()
{
    //ExStart
    //ExFor:DocumentBuilder.DeleteRow
    //ExSummary:Shows how to delete a row from a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 1, cell 2.");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Row 2, cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 2, cell 2.");
    builder->EndTable();
    
    ASSERT_EQ(2, table->get_Rows()->get_Count());
    
    // Delete the first row of the first table in the document.
    builder->DeleteRow(0, 0);
    
    ASSERT_EQ(1, table->get_Rows()->get_Count());
    ASSERT_EQ(u"Row 2, cell 1.\aRow 2, cell 2.\a\a", table->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DeleteRow)
{
    s_instance->DeleteRow();
}

} // namespace gtest_test

void ExDocumentBuilder::AppendDocumentAndResolveStyles(bool keepSourceNumbering)
{
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to manage list style clashes while appending a document.
    // Load a document with text in a custom style and clone it.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Custom list numbering.docx");
    System::SharedPtr<Aspose::Words::Document> dstDoc = srcDoc->Clone();
    
    // We now have two documents, each with an identical style named "CustomStyle".
    // Change the text color for one of the styles to set it apart from the other.
    dstDoc->get_Styles()->idx_get(u"CustomStyle")->get_Font()->set_Color(System::Drawing::Color::get_DarkRed());
    
    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    auto options = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    options->set_KeepSourceNumbering(keepSourceNumbering);
    
    // Joining two documents that have different styles that share the same name causes a style clash.
    // We can specify an import format mode while appending documents to resolve this clash.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepDifferentStyles, options);
    dstDoc->UpdateListLabels();
    
    dstDoc->Save(get_ArtifactsDir() + u"DocumentBuilder.AppendDocumentAndResolveStyles.docx");
    //ExEnd
}

namespace gtest_test
{

using ExDocumentBuilder_AppendDocumentAndResolveStyles_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::AppendDocumentAndResolveStyles)>::type;

struct ExDocumentBuilder_AppendDocumentAndResolveStyles : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_AppendDocumentAndResolveStyles_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_AppendDocumentAndResolveStyles, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AppendDocumentAndResolveStyles(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_AppendDocumentAndResolveStyles, ::testing::ValuesIn(ExDocumentBuilder_AppendDocumentAndResolveStyles::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::InsertDocumentAndResolveStyles(bool keepSourceNumbering)
{
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to manage list style clashes while inserting a document.
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    builder->InsertBreak(Aspose::Words::BreakType::ParagraphBreak);
    
    dstDoc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);
    System::SharedPtr<Aspose::Words::Lists::List> list = dstDoc->get_Lists()->idx_get(0);
    
    builder->get_ListFormat()->set_List(list);
    
    for (int32_t i = 1; i <= 15; i++)
    {
        builder->Write(System::String::Format(u"List Item {0}\n", i));
    }
    
    auto attachDoc = System::ExplicitCast<Aspose::Words::Document>(System::ExplicitCast<Aspose::Words::Node>(dstDoc)->Clone(true));
    
    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    auto importOptions = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    importOptions->set_KeepSourceNumbering(keepSourceNumbering);
    
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->InsertDocument(attachDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, importOptions);
    
    dstDoc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertDocumentAndResolveStyles.docx");
    //ExEnd
}

namespace gtest_test
{

using ExDocumentBuilder_InsertDocumentAndResolveStyles_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::InsertDocumentAndResolveStyles)>::type;

struct ExDocumentBuilder_InsertDocumentAndResolveStyles : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_InsertDocumentAndResolveStyles_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_InsertDocumentAndResolveStyles, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertDocumentAndResolveStyles(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_InsertDocumentAndResolveStyles, ::testing::ValuesIn(ExDocumentBuilder_InsertDocumentAndResolveStyles::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::LoadDocumentWithListNumbering(bool keepSourceNumbering)
{
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to manage list style clashes while appending a clone of a document to itself.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List item.docx");
    auto dstDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List item.docx");
    
    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    builder->MoveToDocumentEnd();
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    
    auto options = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    options->set_KeepSourceNumbering(keepSourceNumbering);
    builder->InsertDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, options);
    
    dstDoc->UpdateListLabels();
    //ExEnd
}

namespace gtest_test
{

using ExDocumentBuilder_LoadDocumentWithListNumbering_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::LoadDocumentWithListNumbering)>::type;

struct ExDocumentBuilder_LoadDocumentWithListNumbering : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_LoadDocumentWithListNumbering_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_LoadDocumentWithListNumbering, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LoadDocumentWithListNumbering(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_LoadDocumentWithListNumbering, ::testing::ValuesIn(ExDocumentBuilder_LoadDocumentWithListNumbering::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::IgnoreTextBoxes(bool ignoreTextBoxes)
{
    //ExStart
    //ExFor:ImportFormatOptions.IgnoreTextBoxes
    //ExSummary:Shows how to manage text box formatting while appending a document.
    // Create a document that will have nodes from another document inserted into it.
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    
    builder->Writeln(u"Hello world!");
    
    // Create another document with a text box, which we will import into the first document.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>();
    builder = System::MakeObject<Aspose::Words::DocumentBuilder>(srcDoc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 300, 100);
    builder->MoveTo(textBox->get_FirstParagraph());
    builder->get_ParagraphFormat()->get_Style()->get_Font()->set_Name(u"Courier New");
    builder->get_ParagraphFormat()->get_Style()->get_Font()->set_Size(24);
    builder->Write(u"Textbox contents");
    
    // Set a flag to specify whether to clear or preserve text box formatting
    // while importing them to other documents.
    auto importFormatOptions = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    importFormatOptions->set_IgnoreTextBoxes(ignoreTextBoxes);
    
    // Import the text box from the source document into the destination document,
    // and then verify whether we have preserved the styling of its text contents.
    auto importer = System::MakeObject<Aspose::Words::NodeImporter>(srcDoc, dstDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, importFormatOptions);
    auto importedTextBox = System::ExplicitCast<Aspose::Words::Drawing::Shape>(importer->ImportNode(textBox, true));
    dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(importedTextBox);
    
    if (ignoreTextBoxes)
    {
        ASPOSE_ASSERT_EQ(12.0, importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Size());
        ASSERT_EQ(u"Times New Roman", importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
    }
    else
    {
        ASPOSE_ASSERT_EQ(24.0, importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Size());
        ASSERT_EQ(u"Courier New", importedTextBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
    }
    
    dstDoc->Save(get_ArtifactsDir() + u"DocumentBuilder.IgnoreTextBoxes.docx");
    //ExEnd
}

namespace gtest_test
{

using ExDocumentBuilder_IgnoreTextBoxes_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::IgnoreTextBoxes)>::type;

struct ExDocumentBuilder_IgnoreTextBoxes : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_IgnoreTextBoxes_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocumentBuilder_IgnoreTextBoxes, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreTextBoxes(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_IgnoreTextBoxes, ::testing::ValuesIn(ExDocumentBuilder_IgnoreTextBoxes::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::MoveToField(bool moveCursorToAfterTheField)
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToField
    //ExSummary:Shows how to move a document builder's node insertion point cursor to a specific field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a field using the DocumentBuilder and add a run of text after it.
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->InsertField(u" AUTHOR \"John Doe\" ");
    
    // The builder's cursor is currently at end of the document.
    ASSERT_TRUE(System::TestTools::IsNull(builder->get_CurrentNode()));
    
    // Move the cursor to the field while specifying whether to place that cursor before or after the field.
    builder->MoveToField(field, moveCursorToAfterTheField);
    
    // Note that the cursor is outside of the field in both cases.
    // This means that we cannot edit the field using the builder like this.
    // To edit a field, we can use the builder's MoveTo method on a field's FieldStart
    // or FieldSeparator node to place the cursor inside.
    if (moveCursorToAfterTheField)
    {
        ASSERT_TRUE(System::TestTools::IsNull(builder->get_CurrentNode()));
        builder->Write(u" Text immediately after the field.");
        
        ASSERT_EQ(u"\u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015 Text immediately after the field.", doc->GetText().Trim());
    }
    else
    {
        ASPOSE_ASSERT_EQ(field->get_Start(), builder->get_CurrentNode());
        builder->Write(u"Text immediately before the field. ");
        
        ASSERT_EQ(u"Text immediately before the field. \u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015", doc->GetText().Trim());
    }
    //ExEnd
}

namespace gtest_test
{

using ExDocumentBuilder_MoveToField_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::MoveToField)>::type;

struct ExDocumentBuilder_MoveToField : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_MoveToField_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_MoveToField, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MoveToField(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_MoveToField, ::testing::ValuesIn(ExDocumentBuilder_MoveToField::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::InsertOleObjectException()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    ASSERT_THROW(static_cast<std::function<void()>>([&builder]() -> void
    {
        builder->InsertOleObject(u"", u"checkbox", false, true, nullptr);
    })(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOleObjectException)
{
    s_instance->InsertOleObjectException();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertPieChart()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
    //ExSummary:Shows how to insert a pie chart into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, Aspose::Words::ConvertUtil::PixelToPoint(300), Aspose::Words::ConvertUtil::PixelToPoint(300))->get_Chart();
    ASPOSE_ASSERT_EQ(225.0, Aspose::Words::ConvertUtil::PixelToPoint(300));
    //ExSkip
    chart->get_Series()->Clear();
    chart->get_Series()->Add(u"My fruit", System::MakeArray<System::String>({u"Apples", u"Bananas", u"Cherries"}), System::MakeArray<double>({1.3, 2.2, 1.5}));
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertPieChart.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertPieChart.docx");
    auto chartShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(u"Chart Title", chartShape->get_Chart()->get_Title()->get_Text());
    ASPOSE_ASSERT_EQ(225.0, chartShape->get_Width());
    ASPOSE_ASSERT_EQ(225.0, chartShape->get_Height());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertPieChart)
{
    s_instance->InsertPieChart();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertChartRelativePosition()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to specify position and wrapping while inserting a chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100, 200, 100, Aspose::Words::Drawing::WrapType::Square);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertedChartRelativePosition.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertedChartRelativePosition.docx");
    auto chartShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(100.0, chartShape->get_Top());
    ASPOSE_ASSERT_EQ(100.0, chartShape->get_Left());
    ASPOSE_ASSERT_EQ(200.0, chartShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, chartShape->get_Height());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, chartShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, chartShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, chartShape->get_RelativeVerticalPosition());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertChartRelativePosition)
{
    s_instance->InsertChartRelativePosition();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertField()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertField(String)
    //ExFor:Field
    //ExFor:Field.Result
    //ExFor:Field.GetFieldCode
    //ExFor:Field.Type
    //ExFor:FieldType
    //ExSummary:Shows how to insert a field into a document using a field code.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->InsertField(u"DATE \\@ \"dddd, MMMM dd, yyyy\"");
    
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Type());
    ASSERT_EQ(u"DATE \\@ \"dddd, MMMM dd, yyyy\"", field->GetFieldCode());
    
    // This overload of the InsertField method automatically updates inserted fields.
    ASSERT_TRUE((System::DateTime::get_Today() - System::DateTime::Parse(field->get_Result())).get_Days() <= 1);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertField)
{
    s_instance->InsertField();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertFieldAndUpdate(bool updateInsertedFieldsImmediately)
{
    //ExStart
    //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
    //ExFor:Field.Update
    //ExSummary:Shows how to insert a field into a document using FieldType.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert two fields while passing a flag which determines whether to update them as the builder inserts them.
    // In some cases, updating fields could be computationally expensive, and it may be a good idea to defer the update.
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    builder->Write(u"This document was written by ");
    builder->InsertField(Aspose::Words::Fields::FieldType::FieldAuthor, updateInsertedFieldsImmediately);
    
    builder->InsertParagraph();
    builder->Write(u"\nThis is page ");
    builder->InsertField(Aspose::Words::Fields::FieldType::FieldPage, updateInsertedFieldsImmediately);
    
    ASSERT_EQ(u" AUTHOR ", doc->get_Range()->get_Fields()->idx_get(0)->GetFieldCode());
    ASSERT_EQ(u" PAGE ", doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode());
    
    if (updateInsertedFieldsImmediately)
    {
        ASSERT_EQ(u"John Doe", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
        ASSERT_EQ(u"1", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
    }
    else
    {
        ASSERT_EQ(System::String::Empty, doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
        ASSERT_EQ(System::String::Empty, doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
        
        // We will need to update these fields using the update methods manually.
        doc->get_Range()->get_Fields()->idx_get(0)->Update();
        
        ASSERT_EQ(u"John Doe", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
        
        doc->UpdateFields();
        
        ASSERT_EQ(u"1", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
    }
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(System::String(u"This document was written by \u0013 AUTHOR \u0014John Doe\u0015") + u"\r\rThis is page \u0013 PAGE \u00141\u0015", doc->GetText().Trim());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldAuthor, u" AUTHOR ", u"John Doe", doc->get_Range()->get_Fields()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldPage, u" PAGE ", u"1", doc->get_Range()->get_Fields()->idx_get(1));
}

namespace gtest_test
{

using ExDocumentBuilder_InsertFieldAndUpdate_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::InsertFieldAndUpdate)>::type;

struct ExDocumentBuilder_InsertFieldAndUpdate : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_InsertFieldAndUpdate_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_InsertFieldAndUpdate, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertFieldAndUpdate(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_InsertFieldAndUpdate, ::testing::ValuesIn(ExDocumentBuilder_InsertFieldAndUpdate::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::FieldResultFormatting()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto formatter = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter>(u"${0}", u"Date: {0}", u"Item # {0}:");
    doc->get_FieldOptions()->set_ResultFormatter(formatter);
    
    // Our field result formatter applies a custom format to newly created fields of three types of formats.
    // Field result formatters apply new formatting to fields as they are updated,
    // which happens as soon as we create them using this InsertField method overload.
    // 1 -  Numeric:
    builder->InsertField(u" = 2 + 3 \\# $###");
    
    ASSERT_EQ(u"$5", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
    ASSERT_EQ(1, formatter->CountFormatInvocations(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::Numeric));
    
    // 2 -  Date/time:
    builder->InsertField(u"DATE \\@ \"d MMMM yyyy\"");
    
    ASSERT_TRUE(doc->get_Range()->get_Fields()->idx_get(1)->get_Result().StartsWith(u"Date: "));
    ASSERT_EQ(1, formatter->CountFormatInvocations(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::DateTime));
    
    // 3 -  General:
    builder->InsertField(u"QUOTE \"2\" \\* Ordinal");
    
    ASSERT_EQ(u"Item # 2:", doc->get_Range()->get_Fields()->idx_get(2)->get_Result());
    ASSERT_EQ(1, formatter->CountFormatInvocations(Aspose::Words::ApiExamples::ExDocumentBuilder::FieldResultFormatter::FormatInvocationType::General));
    
    formatter->PrintFormatInvocations();
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, FieldResultFormatting)
{
    s_instance->FieldResultFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertVideoWithUrl()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertOnlineVideo(String, Double, Double)
    //ExSummary:Shows how to insert an online video into a document using a URL.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertOnlineVideo(u"https://youtu.be/g1N9ke8Prmk", 360, 270);
    
    // We can watch the video from Microsoft Word by clicking on the shape.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertVideoWithUrl.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertVideoWithUrl.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(480, 360, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    ASSERT_EQ(u"https://youtu.be/t_1LYZ102RA", shape->get_HRef());
    
    ASPOSE_ASSERT_EQ(360.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(270.0, shape->get_Height());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DISABLED_InsertVideoWithUrl)
{
    s_instance->InsertVideoWithUrl();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertUnderline()
{
    //ExStart
    //ExFor:DocumentBuilder.Underline
    //ExSummary:Shows how to format text inserted by a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->set_Underline(Aspose::Words::Underline::Dash);
    builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    builder->get_Font()->set_Size(32);
    
    // The builder applies formatting to its current paragraph and any new text added by it afterward.
    builder->Writeln(u"Large, blue, and underlined text.");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertUnderline.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertUnderline.docx");
    System::SharedPtr<Aspose::Words::Run> firstRun = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Large, blue, and underlined text.", firstRun->GetText().Trim());
    ASSERT_EQ(Aspose::Words::Underline::Dash, firstRun->get_Font()->get_Underline());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), firstRun->get_Font()->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(32.0, firstRun->get_Font()->get_Size());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertUnderline)
{
    s_instance->InsertUnderline();
}

} // namespace gtest_test

void ExDocumentBuilder::CurrentStory()
{
    //ExStart
    //ExFor:DocumentBuilder.CurrentStory
    //ExSummary:Shows how to work with a document builder's current story.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A Story is a type of node that has child Paragraph nodes, such as a Body.
    ASPOSE_ASSERT_EQ(builder->get_CurrentStory(), doc->get_FirstSection()->get_Body());
    ASPOSE_ASSERT_EQ(builder->get_CurrentStory(), builder->get_CurrentParagraph()->get_ParentNode());
    ASSERT_EQ(Aspose::Words::StoryType::MainText, builder->get_CurrentStory()->get_StoryType());
    
    builder->get_CurrentStory()->AppendParagraph(u"Text added to current Story.");
    
    // A Story can also contain tables.
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, cell 1");
    builder->InsertCell();
    builder->Write(u"Row 1, cell 2");
    builder->EndTable();
    
    ASSERT_TRUE(builder->get_CurrentStory()->get_Tables()->Contains(table));
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Tables()->get_Count());
    ASSERT_EQ(u"Row 1, cell 1\aRow 1, cell 2\a\a\rText added to current Story.", doc->get_FirstSection()->get_Body()->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, CurrentStory)
{
    s_instance->CurrentStory();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertOleObjects()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertOleObject(Stream, String, Boolean, Stream)
    //ExSummary:Shows how to use document builder to embed OLE objects in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Microsoft Excel spreadsheet from the local file system
    // into the document while keeping its default appearance.
    {
        System::SharedPtr<System::IO::Stream> spreadsheetStream = System::IO::File::Open(get_MyDir() + u"Spreadsheet.xlsx", System::IO::FileMode::Open);
        builder->Writeln(u"Spreadsheet Ole object:");
        // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
        // the icon according to 'progId' and uses the predefined icon caption.
        builder->InsertOleObject(spreadsheetStream, u"OleObject.xlsx", false, nullptr);
    }
    
    // Insert a Microsoft Powerpoint presentation as an OLE object.
    // This time, it will have an image downloaded from the web for an icon.
    {
        System::SharedPtr<System::IO::Stream> powerpointStream = System::IO::File::Open(get_MyDir() + u"Presentation.pptx", System::IO::FileMode::Open);
        System::ArrayPtr<uint8_t> imgBytes = System::IO::File::ReadAllBytes(get_ImageDir() + u"Logo.jpg");
        
        {
            auto imageStream = System::MakeObject<System::IO::MemoryStream>(imgBytes);
            builder->InsertParagraph();
            builder->Writeln(u"Powerpoint Ole object:");
            builder->InsertOleObject(powerpointStream, u"OleObject.pptx", true, imageStream);
        }
    }
    
    // Double-click these objects in Microsoft Word to open
    // the linked files using their respective applications.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertOleObjects.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertOleObjects.docx");
    
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASSERT_EQ(u"", shape->get_OleFormat()->get_IconCaption());
    ASSERT_FALSE(shape->get_OleFormat()->get_OleIcon());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    ASSERT_EQ(u"Unknown", shape->get_OleFormat()->get_IconCaption());
    ASSERT_TRUE(shape->get_OleFormat()->get_OleIcon());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOleObjects)
{
    s_instance->InsertOleObjects();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertDocument()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode)
    //ExFor:ImportFormatMode
    //ExSummary:Shows how to insert a document into another document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    auto docToInsert = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Formatted elements.docx");
    
    builder->InsertDocument(docToInsert, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    builder->get_Document()->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertDocument.docx");
    //ExEnd
    
    ASSERT_EQ(29, doc->get_Styles()->get_Count());
    ASSERT_TRUE(Aspose::Words::ApiExamples::DocumentHelper::CompareDocs(get_ArtifactsDir() + u"DocumentBuilder.InsertDocument.docx", get_GoldsDir() + u"DocumentBuilder.InsertDocument Gold.docx"));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DISABLED_InsertDocument)
{
    s_instance->InsertDocument();
}

} // namespace gtest_test

void ExDocumentBuilder::SmartStyleBehavior()
{
    //ExStart
    //ExFor:ImportFormatOptions
    //ExFor:ImportFormatOptions.SmartStyleBehavior
    //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to resolve duplicate styles while inserting documents.
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    
    System::SharedPtr<Aspose::Words::Style> myStyle = builder->get_Document()->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyStyle");
    myStyle->get_Font()->set_Size(14);
    myStyle->get_Font()->set_Name(u"Courier New");
    myStyle->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    
    builder->get_ParagraphFormat()->set_StyleName(myStyle->get_Name());
    builder->Writeln(u"Hello world!");
    
    // Clone the document and edit the clone's "MyStyle" style, so it is a different color than that of the original.
    // If we insert the clone into the original document, the two styles with the same name will cause a clash.
    System::SharedPtr<Aspose::Words::Document> srcDoc = dstDoc->Clone();
    srcDoc->get_Styles()->idx_get(u"MyStyle")->get_Font()->set_Color(System::Drawing::Color::get_Red());
    
    // When we enable SmartStyleBehavior and use the KeepSourceFormatting import format mode,
    // Aspose.Words will resolve style clashes by converting source document styles.
    // with the same names as destination styles into direct paragraph attributes.
    auto options = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    options->set_SmartStyleBehavior(true);
    
    builder->InsertDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, options);
    
    dstDoc->Save(get_ArtifactsDir() + u"DocumentBuilder.SmartStyleBehavior.docx");
    //ExEnd
    
    dstDoc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.SmartStyleBehavior.docx");
    
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), dstDoc->get_Styles()->idx_get(u"MyStyle")->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(u"MyStyle", dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Style()->get_Name());
    
    ASSERT_EQ(u"Normal", dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_Style()->get_Name());
    ASPOSE_ASSERT_EQ(14, dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Size());
    ASSERT_EQ(u"Courier New", dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Name());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), dstDoc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Color().ToArgb());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SmartStyleBehavior)
{
    s_instance->SmartStyleBehavior();
}

} // namespace gtest_test

void ExDocumentBuilder::EmphasesWarningSourceMarkdown()
{
    //ExStart
    //ExFor:WarningInfo.Source
    //ExFor:WarningSource
    //ExSummary:Shows how to work with the warning source.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Emphases markdown warning.docx");
    
    auto warnings = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    doc->set_WarningCallback(warnings);
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.EmphasesWarningSourceMarkdown.md");
    
    for (auto&& warningInfo : warnings)
    {
        if (warningInfo->get_Source() == Aspose::Words::WarningSource::Markdown)
        {
            ASSERT_EQ(u"The (*, 0:11) cannot be properly written into Markdown.", warningInfo->get_Description());
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, EmphasesWarningSourceMarkdown)
{
    s_instance->EmphasesWarningSourceMarkdown();
}

} // namespace gtest_test

void ExDocumentBuilder::DoNotIgnoreHeaderFooter()
{
    //ExStart
    //ExFor:ImportFormatOptions.IgnoreHeaderFooter
    //ExSummary:Shows how to specifies ignoring or not source formatting of headers/footers content.
    auto dstDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Header and footer types.docx");
    
    // If 'IgnoreHeaderFooter' is false then the original formatting for header/footer content
    // from "Header and footer types.docx" will be used.
    // If 'IgnoreHeaderFooter' is true then the formatting for header/footer content
    // from "Document.docx" will be used.
    auto importFormatOptions = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    importFormatOptions->set_IgnoreHeaderFooter(false);
    
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, importFormatOptions);
    
    dstDoc->Save(get_ArtifactsDir() + u"DocumentBuilder.DoNotIgnoreHeaderFooter.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DoNotIgnoreHeaderFooter)
{
    s_instance->DoNotIgnoreHeaderFooter();
}

} // namespace gtest_test

void ExDocumentBuilder::LoadMarkdownDocumentAndAssertContent(System::String text, System::String styleName, bool isItalic, bool isBold)
{
    // Prepeare document to test.
    MarkdownDocumentEmphases();
    MarkdownDocumentInlineCode();
    MarkdownDocumentHeadings();
    MarkdownDocumentBlockquotes();
    MarkdownDocumentIndentedCode();
    MarkdownDocumentFencedCode();
    MarkdownDocumentHorizontalRule();
    MarkdownDocumentBulletedList();
    
    // Load created document from previous tests.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.MarkdownDocument.md");
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(paragraphs))
    {
        if (paragraph->get_Runs()->get_Count() != 0)
        {
            // Check that all document text has the necessary styles.
            if (paragraph->get_Runs()->idx_get(0)->get_Text() == text && !text.Contains(u"InlineCode"))
            {
                ASSERT_EQ(styleName, paragraph->get_ParagraphFormat()->get_Style()->get_Name());
                ASPOSE_ASSERT_EQ(isItalic, paragraph->get_Runs()->idx_get(0)->get_Font()->get_Italic());
                ASPOSE_ASSERT_EQ(isBold, paragraph->get_Runs()->idx_get(0)->get_Font()->get_Bold());
            }
            else if (paragraph->get_Runs()->idx_get(0)->get_Text() == text && text.Contains(u"InlineCode"))
            {
                ASSERT_EQ(styleName, paragraph->get_Runs()->idx_get(0)->get_Font()->get_StyleName());
            }
        }
        
        // Check that document also has a HorizontalRule present as a shape.
        System::SharedPtr<Aspose::Words::NodeCollection> shapesCollection = doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Shape, true);
        auto horizontalRuleShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapesCollection->idx_get(0));
        
        ASSERT_TRUE(shapesCollection->get_Count() == 1);
        ASSERT_TRUE(horizontalRuleShape->get_IsHorizontalRule());
    }
}

namespace gtest_test
{

using ExDocumentBuilder_LoadMarkdownDocumentAndAssertContent_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocumentBuilder::LoadMarkdownDocumentAndAssertContent)>::type;

struct ExDocumentBuilder_LoadMarkdownDocumentAndAssertContent : public ExDocumentBuilder, public Aspose::Words::ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<ExDocumentBuilder_LoadMarkdownDocumentAndAssertContent_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(u"Italic", u"Normal", true, false),
            std::make_tuple(u"Bold", u"Normal", false, true),
            std::make_tuple(u"ItalicBold", u"Normal", true, true),
            std::make_tuple(u"Text with InlineCode style with one backtick", u"InlineCode", false, false),
            std::make_tuple(u"Text with InlineCode style with 3 backticks", u"InlineCode.3", false, false),
            std::make_tuple(u"This is an italic H1 tag", u"Heading 1", true, false),
            std::make_tuple(u"SetextHeading 1", u"SetextHeading1", false, false),
            std::make_tuple(u"This is an H2 tag", u"Heading 2", false, false),
            std::make_tuple(u"SetextHeading 2", u"SetextHeading2", false, false),
            std::make_tuple(u"This is an H3 tag", u"Heading 3", false, false),
            std::make_tuple(u"This is an bold H4 tag", u"Heading 4", false, true),
            std::make_tuple(u"This is an italic and bold H5 tag", u"Heading 5", true, true),
            std::make_tuple(u"This is an H6 tag", u"Heading 6", false, false),
            std::make_tuple(u"Blockquote", u"Quote", false, false),
            std::make_tuple(u"1. Nested blockquote", u"Quote1", false, false),
            std::make_tuple(u"2. Nested italic blockquote", u"Quote2", true, false),
            std::make_tuple(u"3. Nested bold blockquote", u"Quote3", false, true),
            std::make_tuple(u"4. Nested blockquote", u"Quote4", false, false),
            std::make_tuple(u"5. Nested blockquote", u"Quote5", false, false),
            std::make_tuple(u"6. Nested italic bold blockquote", u"Quote6", true, true),
            std::make_tuple(u"This is an indented code", u"IndentedCode", false, false),
            std::make_tuple(u"This is a fenced code", u"FencedCode", false, false),
            std::make_tuple(u"This is a fenced code with info string", u"FencedCode.C#", false, false),
            std::make_tuple(u"Item 1", u"Normal", false, false),
        };
    }
};

TEST_P(ExDocumentBuilder_LoadMarkdownDocumentAndAssertContent, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LoadMarkdownDocumentAndAssertContent(std::get<0>(params), std::get<1>(params), std::get<2>(params), std::get<3>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_LoadMarkdownDocumentAndAssertContent, ::testing::ValuesIn(ExDocumentBuilder_LoadMarkdownDocumentAndAssertContent::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::InsertOnlineVideoCustomThumbnail()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], Double, Double)
    //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an online video into a document with a custom thumbnail.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::String videoUrl = u"https://vimeo.com/52477838";
    System::String videoEmbedCode = System::String(u"<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" ") + u"title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";
    
    System::ArrayPtr<uint8_t> thumbnailImageBytes = System::IO::File::ReadAllBytes(get_ImageDir() + u"Logo.jpg");
    
    {
        auto stream = System::MakeObject<System::IO::MemoryStream>(thumbnailImageBytes);
        {
            System::SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);
            // Below are two ways of creating a shape with a custom thumbnail, which links to an online video
            // that will play when we click on the shape in Microsoft Word.
            // 1 -  Insert an inline shape at the builder's node insertion cursor:
            builder->InsertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, image->get_Width(), image->get_Height());
            
            builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
            
            // 2 -  Insert a floating shape:
            double left = builder->get_PageSetup()->get_RightMargin() - image->get_Width();
            double top = builder->get_PageSetup()->get_BottomMargin() - image->get_Height();
            
            builder->InsertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, left, Aspose::Words::Drawing::RelativeVerticalPosition::BottomMargin, top, image->get_Width(), image->get_Height(), Aspose::Words::Drawing::WrapType::Square);
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    ASPOSE_ASSERT_EQ(400.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(400.0, shape->get_Height());
    ASPOSE_ASSERT_EQ(0.0, shape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, shape->get_Top());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, shape->get_RelativeVerticalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, shape->get_RelativeHorizontalPosition());
    
    ASSERT_EQ(u"https://vimeo.com/52477838", shape->get_HRef());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    ASPOSE_ASSERT_EQ(400.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(400.0, shape->get_Height());
    ASPOSE_ASSERT_EQ(-329.15, shape->get_Left());
    ASPOSE_ASSERT_EQ(-329.15, shape->get_Top());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, shape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::BottomMargin, shape->get_RelativeVerticalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, shape->get_RelativeHorizontalPosition());
    
    ASSERT_EQ(u"https://vimeo.com/52477838", shape->get_HRef());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOnlineVideoCustomThumbnail)
{
    s_instance->InsertOnlineVideoCustomThumbnail();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertOleObjectAsIcon()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertOleObjectAsIcon(String, String, Boolean, String, String)
    //ExFor:DocumentBuilder.InsertOleObjectAsIcon(Stream, String, String, String)
    //ExSummary:Shows how to insert an embedded or linked OLE object as icon into the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
    // the icon according to 'progId' and uses the filename for the icon caption.
    builder->InsertOleObjectAsIcon(get_MyDir() + u"Presentation.pptx", u"Package", false, get_ImageDir() + u"Logo icon.ico", u"My embedded file");
    
    builder->InsertBreak(Aspose::Words::BreakType::LineBreak);
    
    {
        auto stream = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"Presentation.pptx", System::IO::FileMode::Open);
        // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
        // the icon according to the file extension and uses the filename for the icon caption.
        System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertOleObjectAsIcon(stream, u"PowerPoint.Application", get_ImageDir() + u"Logo icon.ico", u"My embedded file stream");
        
        System::SharedPtr<Aspose::Words::Drawing::OlePackage> setOlePackage = shape->get_OleFormat()->get_OlePackage();
        setOlePackage->set_FileName(u"Presentation.pptx");
        setOlePackage->set_DisplayName(u"Presentation.pptx");
    }
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.InsertOleObjectAsIcon.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOleObjectAsIcon)
{
    s_instance->InsertOleObjectAsIcon();
}

} // namespace gtest_test

void ExDocumentBuilder::PreserveBlocks()
{
    //ExStart
    //ExFor:HtmlInsertOptions
    //ExSummary:Shows how to allows better preserve borders and margins seen.
    const System::String html = u"\r\n                <html>\r\n                    <div style='border:dotted'>\r\n                    <div style='border:solid'>\r\n                        <p>paragraph 1</p>\r\n                        <p>paragraph 2</p>\r\n                    </div>\r\n                    </div>\r\n                </html>";
    
    // Set the new mode of import HTML block-level elements.
    Aspose::Words::HtmlInsertOptions insertOptions = Aspose::Words::HtmlInsertOptions::PreserveBlocks;
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    builder->InsertHtml(html, insertOptions);
    builder->get_Document()->Save(get_ArtifactsDir() + u"DocumentBuilder.PreserveBlocks.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, PreserveBlocks)
{
    s_instance->PreserveBlocks();
}

} // namespace gtest_test

void ExDocumentBuilder::PhoneticGuide()
{
    //ExStart
    //ExFor:Run.IsPhoneticGuide
    //ExFor:Run.PhoneticGuide
    //ExFor:PhoneticGuide.BaseText
    //ExFor:PhoneticGuide.RubyText
    //ExSummary:Shows how to get properties of the phonetic guide.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Phonetic guide.docx");
    
    System::SharedPtr<Aspose::Words::RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();
    // Use phonetic guide in the Asian text.
    ASPOSE_ASSERT_EQ(true, runs->idx_get(0)->get_IsPhoneticGuide());
    ASSERT_EQ(u"base", runs->idx_get(0)->get_PhoneticGuide()->get_BaseText());
    ASSERT_EQ(u"ruby", runs->idx_get(0)->get_PhoneticGuide()->get_RubyText());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, PhoneticGuide)
{
    s_instance->PhoneticGuide();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
