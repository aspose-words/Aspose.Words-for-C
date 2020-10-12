// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBuilder.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
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
#include <system/date_time.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <net/web_client.h>
#include <net/service_point_manager.h>
#include <net/http_status_code.h>
#include <functional>
#include <drawing/image.h>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/TextOrientation.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TableStyleOptions.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
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
#include <Aspose.Words.Cpp/Model/Sections/PaperSize.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
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
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldResultFormatter.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/DropDownItemCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/GeneralFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToc.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Document/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Document/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/Document/DateTimeFormatting/CalendarType.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Tables;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(835637885u, ::ApiExamples::ExDocumentBuilder::FieldResultFormatter, ThisTypeBaseTypesInfo);

ExDocumentBuilder::FieldResultFormatter::FieldResultFormatter(String numberFormat, String dateFormat, String generalFormat)
     : mNumberFormatInvocations(MakeObject<System::Collections::Generic::List<SharedPtr<System::Object>>>())
    , mDateFormatInvocations(MakeObject<System::Collections::Generic::List<SharedPtr<System::Object>>>())
    , mGeneralFormatInvocations(MakeObject<System::Collections::Generic::List<SharedPtr<System::Object>>>())
{
    mNumberFormat = numberFormat;
    mDateFormat = dateFormat;
    mGeneralFormat = generalFormat;
}

String ExDocumentBuilder::FieldResultFormatter::FormatNumeric(double value, String format)
{
    mNumberFormatInvocations->Add(MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<double>(value), System::ObjectExt::Box<String>(format)}));

    return String::IsNullOrEmpty(mNumberFormat) ? nullptr : String::Format(mNumberFormat,value);
}

String ExDocumentBuilder::FieldResultFormatter::FormatDateTime(System::DateTime value, String format, Aspose::Words::CalendarType calendarType)
{
    mDateFormatInvocations->Add(MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(value), System::ObjectExt::Box<String>(format), System::ObjectExt::Box<CalendarType>(calendarType)}));

    return String::IsNullOrEmpty(mDateFormat) ? nullptr : String::Format(mDateFormat,value);
}

String ExDocumentBuilder::FieldResultFormatter::Format(String value, Aspose::Words::Fields::GeneralFormat format)
{
    return Format(System::ObjectExt::Box<String>(value), format);
}

String ExDocumentBuilder::FieldResultFormatter::Format(double value, Aspose::Words::Fields::GeneralFormat format)
{
    return Format(System::ObjectExt::Box<double>(value), format);
}

String ExDocumentBuilder::FieldResultFormatter::Format(SharedPtr<System::Object> value, Aspose::Words::Fields::GeneralFormat format)
{
    mGeneralFormatInvocations->Add(MakeArray<SharedPtr<System::Object>>({value, System::ObjectExt::Box<GeneralFormat>(format)}));

    return String::IsNullOrEmpty(mGeneralFormat) ? nullptr : String::Format(mGeneralFormat,value);
}

void ExDocumentBuilder::FieldResultFormatter::PrintInvocations()
{
    System::Console::WriteLine(u"Number format invocations ({0}):", System::ObjectExt::Box<int>(mNumberFormatInvocations->get_Count()));
    for (auto s : System::IterateOver<System::Array<SharedPtr<System::Object>>>(mNumberFormatInvocations))
    {
        System::Console::WriteLine(String(u"\tValue: ") + s[0] + u", original format: " + s[1]);
    }

    System::Console::WriteLine(u"Date format invocations ({0}):", System::ObjectExt::Box<int>(mDateFormatInvocations->get_Count()));
    for (auto s : System::IterateOver<System::Array<SharedPtr<System::Object>>>(mDateFormatInvocations))
    {
        System::Console::WriteLine(String(u"\tValue: ") + s[0] + u", original format: " + s[1] + u", calendar type: " + s[2]);
    }

    System::Console::WriteLine(u"General format invocations ({0}):", System::ObjectExt::Box<int>(mGeneralFormatInvocations->get_Count()));
    for (auto s : System::IterateOver<System::Array<SharedPtr<System::Object>>>(mGeneralFormatInvocations))
    {
        System::Console::WriteLine(String(u"\tValue: ") + s[0] + u", original format: " + s[1]);
    }
}

System::Object::shared_members_type ApiExamples::ExDocumentBuilder::FieldResultFormatter::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentBuilder::FieldResultFormatter::mNumberFormatInvocations", this->mNumberFormatInvocations);
    result.Add("ApiExamples::ExDocumentBuilder::FieldResultFormatter::mDateFormatInvocations", this->mDateFormatInvocations);
    result.Add("ApiExamples::ExDocumentBuilder::FieldResultFormatter::mGeneralFormatInvocations", this->mGeneralFormatInvocations);

    return result;
}

void ExDocumentBuilder::MarkdownDocumentEmphases()
{
    auto builder = MakeObject<DocumentBuilder>();

    // Bold and Italic are represented as Font.Bold and Font.Italic
    builder->get_Font()->set_Italic(true);
    builder->Writeln(u"This text will be italic");

    // Use clear formatting if don't want to combine styles between paragraphs
    builder->get_Font()->ClearFormatting();

    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"This text will be bold");

    builder->get_Font()->ClearFormatting();

    // You can also write create BoldItalic text
    builder->get_Font()->set_Italic(true);
    builder->Write(u"You ");
    builder->get_Font()->set_Bold(true);
    builder->Write(u"can");
    builder->get_Font()->set_Bold(false);
    builder->Writeln(u" combine them");

    builder->get_Font()->ClearFormatting();

    builder->get_Font()->set_StrikeThrough(true);
    builder->Writeln(u"This text will be strikethrough");

    // Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis
    builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentInlineCode()
{
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Prepare our created document for further work
    // And clear paragraph formatting not to use the previous styles
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");

    // Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`)
    // If number of backticks is missed, then one backtick will be used by default
    SharedPtr<Style> inlineCode1BackTicks = doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"InlineCode");
    builder->get_Font()->set_Style(inlineCode1BackTicks);
    builder->Writeln(u"Text with InlineCode style with one backtick");

    // Use optional dot (.) and number of backticks (`)
    // There will be 3 backticks
    SharedPtr<Style> inlineCode3BackTicks = doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"InlineCode.3");
    builder->get_Font()->set_Style(inlineCode3BackTicks);
    builder->Writeln(u"Text with InlineCode style with 3 backticks");

    builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentHeadings()
{
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Prepare our created document for further work
    // And clear paragraph formatting not to use the previous styles
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");

    // By default Heading styles in Word may have bold and italic formatting
    // If we do not want text to be emphasized, set these properties explicitly to false
    // Thus we can't use 'builder.Font.ClearFormatting()' because Bold/Italic will be set to true
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);

    // Create for one heading for each level
    builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
    builder->get_Font()->set_Italic(true);
    builder->Writeln(u"This is an italic H1 tag");

    // Reset our styles from the previous paragraph to not combine styles between paragraphs
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);

    // Structure-enhanced text heading can be added through style inheritance
    SharedPtr<Style> setextHeading1 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"SetextHeading1");
    builder->get_ParagraphFormat()->set_Style(setextHeading1);
    doc->get_Styles()->idx_get(u"SetextHeading1")->set_BaseStyleName(u"Heading 1");
    builder->Writeln(u"SetextHeading 1");

    builder->get_ParagraphFormat()->set_StyleName(u"Heading 2");
    builder->Writeln(u"This is an H2 tag");

    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_Italic(false);

    SharedPtr<Style> setextHeading2 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"SetextHeading2");
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

    doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentBlockquotes()
{
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Prepare our created document for further work
    // And clear paragraph formatting not to use the previous styles
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");

    // By default document stores blockquote style for the first level
    builder->get_ParagraphFormat()->set_StyleName(u"Quote");
    builder->Writeln(u"Blockquote");

    // Create styles for nested levels through style inheritance
    SharedPtr<Style> quoteLevel2 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote1");
    builder->get_ParagraphFormat()->set_Style(quoteLevel2);
    doc->get_Styles()->idx_get(u"Quote1")->set_BaseStyleName(u"Quote");
    builder->Writeln(u"1. Nested blockquote");

    SharedPtr<Style> quoteLevel3 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote2");
    builder->get_ParagraphFormat()->set_Style(quoteLevel3);
    doc->get_Styles()->idx_get(u"Quote2")->set_BaseStyleName(u"Quote1");
    builder->get_Font()->set_Italic(true);
    builder->Writeln(u"2. Nested italic blockquote");

    SharedPtr<Style> quoteLevel4 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote3");
    builder->get_ParagraphFormat()->set_Style(quoteLevel4);
    doc->get_Styles()->idx_get(u"Quote3")->set_BaseStyleName(u"Quote2");
    builder->get_Font()->set_Italic(false);
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"3. Nested bold blockquote");

    SharedPtr<Style> quoteLevel5 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote4");
    builder->get_ParagraphFormat()->set_Style(quoteLevel5);
    doc->get_Styles()->idx_get(u"Quote4")->set_BaseStyleName(u"Quote3");
    builder->get_Font()->set_Bold(false);
    builder->Writeln(u"4. Nested blockquote");

    SharedPtr<Style> quoteLevel6 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote5");
    builder->get_ParagraphFormat()->set_Style(quoteLevel6);
    doc->get_Styles()->idx_get(u"Quote5")->set_BaseStyleName(u"Quote4");
    builder->Writeln(u"5. Nested blockquote");

    SharedPtr<Style> quoteLevel7 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote6");
    builder->get_ParagraphFormat()->set_Style(quoteLevel7);
    doc->get_Styles()->idx_get(u"Quote6")->set_BaseStyleName(u"Quote5");
    builder->get_Font()->set_Italic(true);
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"6. Nested italic bold blockquote");

    doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentIndentedCode()
{
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Prepare our created document for further work
    // And clear paragraph formatting not to use the previous styles
    builder->MoveToDocumentEnd();
    builder->Writeln(u"\n");
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");

    SharedPtr<Style> indentedCode = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"IndentedCode");
    builder->get_ParagraphFormat()->set_Style(indentedCode);
    builder->Writeln(u"This is an indented code");

    doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentFencedCode()
{
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Prepare our created document for further work
    // And clear paragraph formatting not to use the previous styles
    builder->MoveToDocumentEnd();
    builder->Writeln(u"\n");
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");

    SharedPtr<Style> fencedCode = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"FencedCode");
    builder->get_ParagraphFormat()->set_Style(fencedCode);
    builder->Writeln(u"This is a fenced code");

    SharedPtr<Style> fencedCodeWithInfo = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"FencedCode.C#");
    builder->get_ParagraphFormat()->set_Style(fencedCodeWithInfo);
    builder->Writeln(u"This is a fenced code with info string");

    doc->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentHorizontalRule()
{
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Prepare our created document for further work
    // And clear paragraph formatting not to use the previous styles
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");

    // Insert HorizontalRule that will be present in .md file as '-----'
    builder->InsertHorizontalRule();

    builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::MarkdownDocumentBulletedList()
{
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Prepare our created document for further work
    // And clear paragraph formatting not to use the previous styles
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(u"\n");

    // Bulleted lists are represented using paragraph numbering
    builder->get_ListFormat()->ApplyBulletDefault();
    // There can be 3 types of bulleted lists
    // The only diff in a numbering format of the very first level are: ‘-’, ‘+’ or ‘*’ respectively
    builder->get_ListFormat()->get_List()->get_ListLevels()->idx_get(0)->set_NumberFormat(u"-");

    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Item 2a");
    builder->Writeln(u"Item 2b");

    builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
}

void ExDocumentBuilder::LoadMarkdownDocumentAndAssertContent(String text, String styleName, bool isItalic, bool isBold)
{
    // Load created document from previous tests
    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MarkdownDocument.md");
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    for (auto paragraph : System::IterateOver<Paragraph>(paragraphs))
    {
        if (paragraph->get_Runs()->get_Count() != 0)
        {
            // Check that all document text has the necessary styles
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

        // Check that document also has a HorizontalRule present as a shape
        SharedPtr<NodeCollection> shapesCollection = doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Shape, true);
        auto horizontalRuleShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(shapesCollection->idx_get(0));

        ASSERT_TRUE(shapesCollection->get_Count() == 1);
        ASSERT_TRUE(horizontalRuleShape->get_IsHorizontalRule());
    }
}

namespace gtest_test
{

class ExDocumentBuilder : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExDocumentBuilder> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDocumentBuilder>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDocumentBuilder> ExDocumentBuilder::s_instance;

} // namespace gtest_test

void ExDocumentBuilder::WriteAndFont()
{
    //ExStart
    //ExFor:Font.Size
    //ExFor:Font.Bold
    //ExFor:Font.Name
    //ExFor:Font.Color
    //ExFor:Font.Underline
    //ExFor:DocumentBuilder.#ctor(Document)
    //ExSummary:Inserts formatted text using DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Specify font formatting before adding text
    SharedPtr<Aspose::Words::Font> font = builder->get_Font();
    font->set_Size(16);
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_Blue());
    font->set_Name(u"Courier New");
    font->set_Underline(Aspose::Words::Underline::Dash);

    builder->Write(u"Hello world!");
    //ExEnd

    doc = DocumentHelper::SaveOpen(builder->get_Document());
    SharedPtr<Run> firstRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

    ASSERT_EQ(u"Hello world!", firstRun->GetText().Trim());
    ASPOSE_ASSERT_EQ(16, firstRun->get_Font()->get_Size());
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

void ExDocumentBuilder::HeadersAndFooters()
{
    //ExStart
    //ExFor:DocumentBuilder
    //ExFor:DocumentBuilder.#ctor(Document)
    //ExFor:DocumentBuilder.MoveToHeaderFooter
    //ExFor:DocumentBuilder.MoveToSection
    //ExFor:DocumentBuilder.InsertBreak
    //ExFor:DocumentBuilder.Writeln(String)
    //ExFor:HeaderFooterType
    //ExFor:PageSetup.DifferentFirstPageHeaderFooter
    //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
    //ExFor:BreakType
    //ExSummary:Shows how to create headers and footers in a document using DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Specify that we want headers and footers different for first, even and odd pages
    builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(true);
    builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(true);

    // Create the headers
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderFirst);
    builder->Write(u"Header for the first page");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderEven);
    builder->Write(u"Header for even pages");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"Header for all other pages");

    // Create three pages in the document
    builder->MoveToSection(0);
    builder->Writeln(u"Page1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page3");

    doc->Save(ArtifactsDir + u"DocumentBuilder.HeadersAndFooters.docx");
    //ExEnd

    SharedPtr<HeaderFooterCollection> headersFooters = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.HeadersAndFooters.docx")->get_FirstSection()->get_HeadersFooters();

    ASSERT_EQ(3, headersFooters->get_Count());
    ASSERT_EQ(u"Header for the first page", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderFirst)->GetText().Trim());
    ASSERT_EQ(u"Header for even pages", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderEven)->GetText().Trim());
    ASSERT_EQ(u"Header for all other pages", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetText().Trim());

}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, HeadersAndFooters)
{
    s_instance->HeadersAndFooters();
}

} // namespace gtest_test

void ExDocumentBuilder::MergeFields()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertField(String)
    //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
    //ExSummary:Shows how to insert merge fields and move between them.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertField(u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
    builder->InsertField(u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

    // The second merge field starts immediately after the end of the first
    // We'll move the builder's cursor to the end of the first so we can split them by text
    builder->MoveToMergeField(u"MyMergeField1", true, false);
    ASPOSE_ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(1)->get_Start(), builder->get_CurrentNode());

    builder->Write(u" Text between our two merge fields. ");

    doc->Save(ArtifactsDir + u"DocumentBuilder.MergeFields.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.MergeFields.docx");

    ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());

    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldMergeField, u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT", u"«MyMergeField1»", doc->get_Range()->get_Fields()->idx_get(0));
    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldMergeField, u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT", u"«MyMergeField2»", doc->get_Range()->get_Fields()->idx_get(1));
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
    //ExFor:HorizontalRuleFormat
    //ExFor:HorizontalRuleFormat.Alignment
    //ExFor:HorizontalRuleFormat.WidthPercent
    //ExFor:HorizontalRuleFormat.Height
    //ExFor:HorizontalRuleFormat.Color
    //ExFor:HorizontalRuleFormat.NoShade
    //ExSummary:Shows how to insert horizontal rule shape in a document and customize the formatting.
    // Use a document builder to insert a horizontal rule
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    SharedPtr<Shape> shape = builder->InsertHorizontalRule();

    SharedPtr<HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();
    horizontalRuleFormat->set_Alignment(Aspose::Words::Drawing::HorizontalRuleAlignment::Center);
    horizontalRuleFormat->set_WidthPercent(70);
    horizontalRuleFormat->set_Height(3);
    horizontalRuleFormat->set_Color(System::Drawing::Color::get_Blue());
    horizontalRuleFormat->set_NoShade(true);

    ASSERT_TRUE(shape->get_IsHorizontalRule());
    ASSERT_TRUE(shape->get_HorizontalRuleFormat()->get_NoShade());
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

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
    auto builder = MakeObject<DocumentBuilder>();
    SharedPtr<Shape> shape = builder->InsertHorizontalRule();

    SharedPtr<HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();
    horizontalRuleFormat->set_WidthPercent(1);
    horizontalRuleFormat->set_WidthPercent(100);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&horizontalRuleFormat]()
    {
        horizontalRuleFormat->set_WidthPercent(0);
    };

    ASSERT_THROW(_local_func_0(), System::ArgumentOutOfRangeException);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_1 = [&horizontalRuleFormat]()
    {
        horizontalRuleFormat->set_WidthPercent(101);
    };

    ASSERT_THROW(_local_func_1(), System::ArgumentOutOfRangeException);

    horizontalRuleFormat->set_Height(0);
    horizontalRuleFormat->set_Height(1584);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_2 = [&horizontalRuleFormat]()
    {
        horizontalRuleFormat->set_Height(-1);
    };

    ASSERT_THROW(_local_func_2(), System::ArgumentOutOfRangeException);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_3 = [&horizontalRuleFormat]()
    {
        horizontalRuleFormat->set_Height(1585);
    };

    ASSERT_THROW(_local_func_3(), System::ArgumentOutOfRangeException);
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
    //ExSummary:Shows how to insert a hyperlink into a document using DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Please make sure to visit ");

    // Specify font formatting for the hyperlink
    builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    builder->get_Font()->set_Underline(Aspose::Words::Underline::Single);

    // Insert the link
    builder->InsertHyperlink(u"Aspose Website", u"https://www.aspose.com", false);

    // Revert to default formatting
    builder->get_Font()->ClearFormatting();
    builder->Write(u" for more information.");

    // Holding Ctrl and left clicking on the field in Microsoft Word will take you to the link's address in a web browser
    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHyperlink.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertHyperlink.docx");

    auto hyperlink = System::DynamicCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));
    TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, hyperlink->get_Address());

    auto fieldContents = System::DynamicCast<Aspose::Words::Run>(hyperlink->get_Start()->get_NextSibling());

    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), fieldContents->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Underline::Single, fieldContents->get_Font()->get_Underline());
    ASSERT_EQ(u"HYPERLINK \"https://www.aspose.com\"", fieldContents->GetText().Trim());
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
    //ExSummary:Shows how to use temporarily save and restore character formatting when building a document with DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set up font formatting and write text that goes before the hyperlink
    builder->get_Font()->set_Name(u"Arial");
    builder->get_Font()->set_Size(24);
    builder->get_Font()->set_Bold(true);
    builder->Write(u"To visit Google, hold Ctrl and click ");

    // Save the font formatting so we use different formatting for hyperlink and restore old formatting later
    builder->PushFont();

    // Set new font formatting for the hyperlink and insert the hyperlink
    // The "Hyperlink" style is a Microsoft Word built-in style so we don't have to worry to
    // create it, it will be created automatically if it does not yet exist in the document
    builder->get_Font()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Hyperlink);
    builder->InsertHyperlink(u"here", u"http://www.google.com", false);

    // Restore the formatting that was before the hyperlink
    builder->PopFont();

    builder->Write(u". We hope you enjoyed the example.");

    doc->Save(ArtifactsDir + u"DocumentBuilder.PushPopFont.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.PushPopFont.docx");
    SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();

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
    //ExFor:DocumentBuilder.InsertImage(Image)
    //ExFor:WrapType
    //ExFor:RelativeHorizontalPosition
    //ExFor:RelativeVerticalPosition
    //ExSummary:Shows how to a watermark image into a document using DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // The best place for the watermark image is in the header or footer so it is shown on every page
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);

    SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");

    // Insert a floating picture
    SharedPtr<Shape> shape = builder->InsertImage(image);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    shape->set_BehindText(true);

    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);

    // Calculate image left and top position so it appears in the center of the page
    shape->set_Left((builder->get_PageSetup()->get_PageWidth() - shape->get_Width()) / 2);
    shape->set_Top((builder->get_PageSetup()->get_PageHeight() - shape->get_Height()) / 2);

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertWatermark.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertWatermark.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, shape);
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert ole object
    auto imageStream = System::MakeObject<System::IO::FileStream>(ImageDir + u"Logo.jpg", System::IO::FileMode::Open);
    builder->InsertOleObject(MyDir + u"Spreadsheet.xlsx", false, false, imageStream);

    // Insert ole object with ProgId
    builder->InsertOleObject(MyDir + u"Spreadsheet.xlsx", u"Excel.Sheet", false, true, nullptr);

    // Insert ole object as Icon
    // There is one limitation for now: the maximum size of the icon must be 32x32 for the correct display
    builder->InsertOleObjectAsIcon(MyDir + u"Presentation.pptx", false, ImageDir + u"Logo icon.ico", u"Caption (can not be null)");

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertOleObject.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertOleObject.docx");
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::OleObject, shape->get_ShapeType());
    ASSERT_EQ(u"Excel.Sheet.12", shape->get_OleFormat()->get_ProgId());
    ASSERT_EQ(u".xlsx", shape->get_OleFormat()->get_SuggestedExtension());

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::OleObject, shape->get_ShapeType());
    ASSERT_EQ(u"Package", shape->get_OleFormat()->get_ProgId());
    ASSERT_EQ(u".xlsx", shape->get_OleFormat()->get_SuggestedExtension());

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

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
    //ExSummary:Shows how to insert Html content into a document using a builder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    const String html = String(u"<P align='right'>Paragraph right</P>") + u"<b>Implicit paragraph left</b>" + u"<div align='center'>Div center</div>" + u"<h1 align='left'>Heading 1 left.</h1>";

    builder->InsertHtml(html);

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHtml.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertHtml.docx");
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    ASSERT_EQ(u"Paragraph right", paragraphs->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());

    ASSERT_EQ(u"Implicit paragraph left", paragraphs->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
    ASSERT_TRUE(paragraphs->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Bold());

    ASSERT_EQ(u"Div center", paragraphs->idx_get(2)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());

    ASSERT_EQ(u"Heading 1 left.", paragraphs->idx_get(3)->GetText().Trim());
    ASSERT_EQ(u"Heading 1", paragraphs->idx_get(3)->get_ParagraphFormat()->get_Style()->get_Name());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertHtml)
{
    s_instance->InsertHtml();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertHtmlWithFormatting()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
    //ExSummary:Shows how to insert Html content into a document using a builder while applying the builder's formatting.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set the builder's text alignment
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Distributed);

    // If we insert text while setting useBuilderFormatting to true, any formatting applied to the builder will be applied to inserted .html content
    // However, if the html text has formatting coded into it, that formatting takes precedence over the builder's formatting
    // In this case, elements with "align" attributes do not get affected by the ParagraphAlignment we specified above
    builder->InsertHtml(String(u"<P align='right'>Paragraph right</P>") + u"<b>Implicit paragraph left</b>" + u"<div align='center'>Div center</div>" + u"<h1 align='left'>Heading 1 left.</h1>", true);

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHtmlWithFormatting.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertHtmlWithFormatting.docx");
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    ASSERT_EQ(u"Paragraph right", paragraphs->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());

    ASSERT_EQ(u"Implicit paragraph left", paragraphs->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Distributed, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
    ASSERT_TRUE(paragraphs->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_Bold());

    ASSERT_EQ(u"Div center", paragraphs->idx_get(2)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());

    ASSERT_EQ(u"Heading 1 left.", paragraphs->idx_get(3)->GetText().Trim());
    ASSERT_EQ(u"Heading 1", paragraphs->idx_get(3)->get_ParagraphFormat()->get_Style()->get_Name());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertHtmlWithFormatting)
{
    s_instance->InsertHtmlWithFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::MathML()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    const String mathMl = u"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>";

    builder->InsertHtml(mathMl);

    doc->Save(ArtifactsDir + u"DocumentBuilder.MathML.docx");
    doc->Save(ArtifactsDir + u"DocumentBuilder.MathML.pdf");

    ASSERT_TRUE(DocumentHelper::CompareDocs(GoldsDir + u"DocumentBuilder.MathML Gold.docx", ArtifactsDir + u"DocumentBuilder.MathML.docx"));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MathML)
{
    s_instance->MathML();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertTextAndBookmark()
{
    //ExStart
    //ExFor:DocumentBuilder.StartBookmark
    //ExFor:DocumentBuilder.EndBookmark
    //ExSummary:Shows how to add some text into the document and encloses the text in a bookmark using DocumentBuilder.
    auto builder = MakeObject<DocumentBuilder>();

    builder->StartBookmark(u"MyBookmark");
    builder->Writeln(u"Text inside a bookmark.");
    builder->EndBookmark(u"MyBookmark");
    //ExEnd

    SharedPtr<Document> doc = DocumentHelper::SaveOpen(builder->get_Document());

    ASSERT_EQ(1, doc->get_Range()->get_Bookmarks()->get_Count());
    ASSERT_EQ(u"MyBookmark", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Name());
    ASSERT_EQ(u"Text inside a bookmark.", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Text().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertTextAndBookmark)
{
    s_instance->InsertTextAndBookmark();
}

} // namespace gtest_test

void ExDocumentBuilder::CreateForm()
{
    //ExStart
    //ExFor:TextFormFieldType
    //ExFor:DocumentBuilder.InsertTextInput
    //ExFor:DocumentBuilder.InsertComboBox
    //ExSummary:Shows how to build a form field.
    auto builder = MakeObject<DocumentBuilder>();

    // Insert a text form field for input a name
    builder->InsertTextInput(u"", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Enter your name here", 30);

    // Insert two blank lines
    builder->Writeln(u"");
    builder->Writeln(u"");

    ArrayPtr<String> items = MakeArray<String>({u"-- Select your favorite footwear --", u"Sneakers", u"Oxfords", u"Flip-flops", u"Other", u"I prefer to be barefoot"});

    // Insert a combo box to select a footwear type
    builder->InsertComboBox(u"", items, 0);

    // Insert 2 blank lines
    builder->Writeln(u"");
    builder->Writeln(u"");

    builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.CreateForm.docx");
    //ExEnd

    auto doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.CreateForm.docx");
    SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);

    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, formField->get_TextInputType());
    ASSERT_EQ(u"Enter your name here", formField->get_Result());

    formField = doc->get_Range()->get_FormFields()->idx_get(1);

    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, formField->get_TextInputType());
    ASSERT_EQ(u"-- Select your favorite footwear --", formField->get_Result());
    ASSERT_EQ(0, formField->get_DropDownSelectedIndex());
    ASPOSE_ASSERT_EQ(MakeArray<String>({u"-- Select your favorite footwear --", u"Sneakers", u"Oxfords", u"Flip-flops", u"Other", u"I prefer to be barefoot"}), formField->get_DropDownItems()->LINQ_ToArray());
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
    //ExSummary:Shows how to insert checkboxes to the document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertCheckBox(String::Empty, false, false, 0);
    builder->InsertCheckBox(u"CheckBox_Default", true, true, 50);
    builder->InsertCheckBox(u"CheckBox_OnlyCheckedValue", true, 100);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);

    // Get checkboxes from the document
    SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();

    // Check that is the right checkbox
    ASSERT_EQ(String::Empty, formFields->idx_get(0)->get_Name());

    //Assert that parameters sets correctly
    ASPOSE_ASSERT_EQ(false, formFields->idx_get(0)->get_Checked());
    ASPOSE_ASSERT_EQ(false, formFields->idx_get(0)->get_Default());
    ASPOSE_ASSERT_EQ(10, formFields->idx_get(0)->get_CheckBoxSize());

    // Check that is the right checkbox
    // Please pay attention that MS Word allows strings with at most 20 characters
    ASSERT_EQ(u"CheckBox_Default", formFields->idx_get(1)->get_Name());

    //Assert that parameters sets correctly
    ASPOSE_ASSERT_EQ(true, formFields->idx_get(1)->get_Checked());
    ASPOSE_ASSERT_EQ(true, formFields->idx_get(1)->get_Default());
    ASPOSE_ASSERT_EQ(50, formFields->idx_get(1)->get_CheckBoxSize());

    // Check that is the right checkbox
    // Please pay attention that MS Word allows strings with at most 20 characters
    ASSERT_EQ(u"CheckBox_OnlyChecked", formFields->idx_get(2)->get_Name());

    // Assert that parameters sets correctly
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Checking that the checkbox insertion with an empty name working correctly
    builder->InsertCheckBox(u"", true, false, 1);
    builder->InsertCheckBox(String::Empty, false, 1);
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
    //ExSummary:Shows how to move a DocumentBuilder to different nodes in a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Start a bookmark and add content to it using a DocumentBuilder
    builder->StartBookmark(u"MyBookmark");
    builder->Writeln(u"Bookmark contents.");
    builder->EndBookmark(u"MyBookmark");

    // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark
    ASPOSE_ASSERT_EQ(doc->get_Range()->get_Bookmarks()->idx_get(0)->get_BookmarkEnd(), builder->get_CurrentParagraph()->get_FirstChild());

    // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this
    builder->MoveToBookmark(u"MyBookmark");

    // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it
    ASPOSE_ASSERT_EQ(doc->get_Range()->get_Bookmarks()->idx_get(0)->get_BookmarkStart(), builder->get_CurrentParagraph()->get_FirstChild());

    // We can move the builder to an individual node,
    // which in this case will be the first node of the first paragraph, like this
    builder->MoveTo(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(Aspose::Words::NodeType::Any, false)->idx_get(0));

    ASSERT_EQ(Aspose::Words::NodeType::BookmarkStart, builder->get_CurrentNode()->get_NodeType());
    ASSERT_TRUE(builder->get_IsAtStartOfParagraph());

    // A shorter way of moving the very start/end of a document is with these methods
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
    //ExSummary:Shows how to fill MERGEFIELDs with data with a DocumentBuilder and without a mail merge.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge
    builder->InsertField(u" MERGEFIELD Chairman ");
    builder->InsertField(u" MERGEFIELD ChiefFinancialOfficer ");
    builder->InsertField(u" MERGEFIELD ChiefTechnologyOfficer ");

    // They can also be filled in manually like this
    builder->MoveToMergeField(u"Chairman");
    builder->set_Bold(true);
    builder->Writeln(u"John Doe");

    builder->MoveToMergeField(u"ChiefFinancialOfficer");
    builder->set_Italic(true);
    builder->Writeln(u"Jane Doe");

    builder->MoveToMergeField(u"ChiefTechnologyOfficer");
    builder->set_Italic(true);
    builder->Writeln(u"John Bloggs");

    doc->Save(ArtifactsDir + u"DocumentBuilder.FillMergeFields.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.FillMergeFields.docx");
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table of contents at the beginning of the document,
    // and set it to pick up paragraphs with headings of levels 1 to 3 and entries to act like hyperlinks
    builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");

    // Start the actual document content on the second page
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    // Build a document with complex structure by applying different heading styles thus creating TOC entries
    // The heading levels we use below will affect the list levels in which these items will appear in the TOC,
    // and only levels 1-3 will be picked up by our TOC due to its settings
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

    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading2);
    builder->Writeln(u"Heading 3.2");
    builder->Writeln(u"Heading 3.3");

    // Call the method below to update the TOC and save
    doc->UpdateFields();
    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertToc.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertToc.docx");
    auto tableOfContents = System::DynamicCast<Aspose::Words::Fields::FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));

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
    //ExFor:DocumentBuilder.Write
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
    //ExSummary:Shows how to build a nice bordered table.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Start building a table
    builder->StartTable();

    // Set the appropriate paragraph, cell, and row formatting. The formatting properties are preserved
    // until they are explicitly modified so there's no need to set them for each row or cell
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

    // Remove the shading (clear background)
    builder->get_CellFormat()->get_Shading()->ClearFormatting();

    builder->InsertCell();
    builder->Write(u"Row 2, Col 1");

    builder->InsertCell();
    builder->Write(u"Row 2, Col 2");

    builder->EndRow();

    builder->InsertCell();

    // Make the row height bigger so that a vertically oriented text could fit into cells
    builder->get_RowFormat()->set_Height(150);
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Upward);
    builder->Write(u"Row 3, Col 1");

    builder->InsertCell();
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Downward);
    builder->Write(u"Row 3, Col 2");

    builder->EndRow();

    builder->EndTable();

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTable.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTable.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(u"Row 1, Col 1\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u"Row 1, Col 2\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    ASPOSE_ASSERT_EQ(50.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::LineStyle::Engrave3D, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Borders()->get_LineStyle());
    ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), table->get_Rows()->idx_get(0)->get_RowFormat()->get_Borders()->get_Color().ToArgb());

    for (auto c : System::IterateOver<Cell>(table->get_Rows()->idx_get(0)->get_Cells()))
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

    for (auto c : System::IterateOver<Cell>(table->get_Rows()->idx_get(1)->get_Cells()))
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
    //ExSummary:Shows how to build a new table with a table style applied.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    SharedPtr<Table> table = builder->StartTable();

    // We must insert at least one row first before setting any table formatting
    builder->InsertCell();

    // Set the table style used based of the unique style identifier
    // Note that not all table styles are available when saving as .doc format
    table->set_StyleIdentifier(Aspose::Words::StyleIdentifier::MediumShading1Accent1);

    // Apply which features should be formatted by the style
    table->set_StyleOptions(Aspose::Words::Tables::TableStyleOptions::FirstColumn | Aspose::Words::Tables::TableStyleOptions::RowBands | Aspose::Words::Tables::TableStyleOptions::FirstRow);
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::AutoFitToContents);

    // Continue with building the table as normal
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

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableWithStyle.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableWithStyle.docx");

    // Verify that the style was set by expanding to direct formatting
    doc->ExpandTableStylesToDirectFormatting();

    ASSERT_EQ(u"Medium Shading 1 Accent 1", table->get_Style()->get_Name());
    ASSERT_EQ(Aspose::Words::Tables::TableStyleOptions::FirstColumn | Aspose::Words::Tables::TableStyleOptions::RowBands | Aspose::Words::Tables::TableStyleOptions::FirstRow, table->get_StyleOptions());
    ASSERT_EQ(189, ASPOSECPP_CHECKED_CAST(int, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().get_B()));
    ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Color().ToArgb());
    ASSERT_NE(System::Drawing::Color::get_LightBlue().ToArgb(), ASPOSECPP_CHECKED_CAST(int, table->get_LastRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().get_B()));
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
    //ExSummary:Shows how to build a table which include heading rows that repeat on subsequent pages.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->StartTable();
    builder->get_RowFormat()->set_HeadingFormat(true);
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->get_CellFormat()->set_Width(100);
    builder->InsertCell();
    builder->Writeln(u"Heading row 1");
    builder->EndRow();
    builder->InsertCell();
    builder->Writeln(u"Heading row 2");
    builder->EndRow();

    builder->get_CellFormat()->set_Width(50);
    builder->get_ParagraphFormat()->ClearFormatting();

    // Insert some content so the table is long enough to continue onto the next page
    for (int i = 0; i < 50; i++)
    {
        builder->InsertCell();
        builder->get_RowFormat()->set_HeadingFormat(false);
        builder->Write(u"Column 1 Text");
        builder->InsertCell();
        builder->Write(u"Column 2 Text");
        builder->EndRow();
    }

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableSetHeadingRow.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableSetHeadingRow.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_TRUE(table->get_FirstRow()->get_RowFormat()->get_HeadingFormat());
    ASSERT_TRUE(table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeadingFormat());
    ASSERT_FALSE(table->get_Rows()->idx_get(2)->get_RowFormat()->get_HeadingFormat());
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
    //ExSummary:Shows how to set a table to auto fit to 50% of the page width.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table with a width that takes up half the page width
    SharedPtr<Table> table = builder->StartTable();

    // Insert a few cells
    builder->InsertCell();
    table->set_PreferredWidth(PreferredWidth::FromPercent(50));
    builder->Writeln(u"Cell #1");

    builder->InsertCell();
    builder->Writeln(u"Cell #2");

    builder->InsertCell();
    builder->Writeln(u"Cell #3");

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableWithPreferredWidth.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableWithPreferredWidth.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

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
    //ExFor:PreferredWidth.Equals(System.Object)
    //ExFor:PreferredWidth.FromPoints
    //ExFor:PreferredWidth.FromPercent
    //ExFor:PreferredWidth.GetHashCode
    //ExFor:PreferredWidth.ToString
    //ExSummary:Shows how to set the different preferred width settings.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table row made up of three cells which have different preferred widths
    SharedPtr<Table> table = builder->StartTable();

    // Insert an absolute sized cell
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(40));
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightYellow());
    builder->Writeln(u"Cell at 40 points width");

    SharedPtr<PreferredWidth> width = builder->get_CellFormat()->get_PreferredWidth();
    System::Console::WriteLine(String::Format(u"Width \"{0}\": {1}",System::ObjectExt::GetHashCode(width),System::ObjectExt::ToString(width)));

    // Insert a relative (percent) sized cell
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(20));
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
    builder->Writeln(u"Cell at 20% width");

    // Each cell had its own PreferredWidth
    ASSERT_FALSE(System::ObjectExt::Equals(builder->get_CellFormat()->get_PreferredWidth(), width));

    width = builder->get_CellFormat()->get_PreferredWidth();
    System::Console::WriteLine(String::Format(u"Width \"{0}\": {1}",System::ObjectExt::GetHashCode(width),System::ObjectExt::ToString(width)));

    // Insert a auto sized cell
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::Auto());
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightGreen());
    builder->Writeln(u"Cell automatically sized. The size of this cell is calculated from the table preferred width.");
    builder->Writeln(u"In this case the cell will fill up the rest of the available space.");

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertCellsWithPreferredWidths.docx");
    //ExEnd

    ASPOSE_ASSERT_EQ(100.0, PreferredWidth::FromPercent(100)->get_Value());
    ASPOSE_ASSERT_EQ(100.0, PreferredWidth::FromPoints(100)->get_Value());

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertCellsWithPreferredWidths.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Points, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(40.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_PreferredWidth()->get_Value());

    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Percent, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(20.0, table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_PreferredWidth()->get_Value());

    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Auto, table->get_FirstRow()->get_Cells()->idx_get(2)->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(2)->get_CellFormat()->get_PreferredWidth()->get_Value());
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
    // inserted from HTML.
    builder->InsertHtml(String(u"<table>") + u"<tr>" + u"<td>Row 1, Cell 1</td>" + u"<td>Row 1, Cell 2</td>" + u"</tr>" + u"<tr>" + u"<td>Row 2, Cell 2</td>" + u"<td>Row 2, Cell 2</td>" + u"</tr>" + u"</table>");

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertTableFromHtml.docx");

    // Verify the table was constructed properly
    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertTableFromHtml.docx");

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
    //ExSummary:Shows how to insert a nested table using DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Build the outer table
    SharedPtr<Cell> cell = builder->InsertCell();
    builder->Writeln(u"Outer Table Cell 1");

    builder->InsertCell();
    builder->Writeln(u"Outer Table Cell 2");

    // This call is important in order to create a nested table within the first table
    // Without this call the cells inserted below will be appended to the outer table
    builder->EndTable();

    // Move to the first cell of the outer table
    builder->MoveTo(cell->get_FirstParagraph());

    // Build the inner table
    builder->InsertCell();
    builder->Writeln(u"Inner Table Cell 1");
    builder->InsertCell();
    builder->Writeln(u"Inner Table Cell 2");

    builder->EndTable();

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertNestedTable.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertNestedTable.docx");

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

void ExDocumentBuilder::CreateSimpleTable()
{
    //ExStart
    //ExFor:DocumentBuilder
    //ExFor:DocumentBuilder.Write
    //ExFor:DocumentBuilder.InsertCell
    //ExSummary:Shows how to create a simple table using DocumentBuilder with default formatting.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // We call this method to start building the table
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 1 Content.");

    // Build the second cell
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 2 Content.");
    // Call the following method to end the row and start a new row
    builder->EndRow();

    // Build the first cell of the second row
    builder->InsertCell();
    builder->Write(u"Row 2, Cell 1 Content.");

    // Build the second cell.
    builder->InsertCell();
    builder->Write(u"Row 2, Cell 2 Content.");
    builder->EndRow();

    // Signal that we have finished building the table
    builder->EndTable();

    // Save the document to disk
    doc->Save(ArtifactsDir + u"DocumentBuilder.CreateSimpleTable.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.CreateSimpleTable.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(4, table->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());

    ASSERT_EQ(u"Row 1, Cell 1 Content.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u"Row 1, Cell 2 Content.\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());
    ASSERT_EQ(u"Row 2, Cell 1 Content.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(u"Row 2, Cell 2 Content.\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());

}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, CreateSimpleTable)
{
    s_instance->CreateSimpleTable();
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Table> table = builder->StartTable();

    // Make the header row
    builder->InsertCell();

    // Set the left indent for the table. Table wide formatting must be applied after
    // at least one row is present in the table
    table->set_LeftIndent(20.0);

    // Set height and define the height rule for the header row
    builder->get_RowFormat()->set_Height(40.0);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::AtLeast);

    // Some special features for the header row
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::FromArgb(198, 217, 241));
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->get_Font()->set_Size(16);
    builder->get_Font()->set_Name(u"Arial");
    builder->get_Font()->set_Bold(true);

    builder->get_CellFormat()->set_Width(100.0);
    builder->Write(u"Header Row,\n Cell 1");

    // We don't need to specify the width of this cell because it's inherited from the previous cell
    builder->InsertCell();
    builder->Write(u"Header Row,\n Cell 2");

    builder->InsertCell();
    builder->get_CellFormat()->set_Width(200.0);
    builder->Write(u"Header Row,\n Cell 3");
    builder->EndRow();

    // Set features for the other rows and cells
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_White());
    builder->get_CellFormat()->set_Width(100.0);
    builder->get_CellFormat()->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);

    // Reset height and define a different height rule for table body
    builder->get_RowFormat()->set_Height(30.0);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Auto);
    builder->InsertCell();
    // Reset font formatting
    builder->get_Font()->set_Size(12);
    builder->get_Font()->set_Bold(false);

    // Build the other cells
    builder->Write(u"Row 1, Cell 1 Content");
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 2 Content");

    builder->InsertCell();
    builder->get_CellFormat()->set_Width(200.0);
    builder->Write(u"Row 1, Cell 3 Content");
    builder->EndRow();

    builder->InsertCell();
    builder->get_CellFormat()->set_Width(100.0);
    builder->Write(u"Row 2, Cell 1 Content");

    builder->InsertCell();
    builder->Write(u"Row 2, Cell 2 Content");

    builder->InsertCell();
    builder->get_CellFormat()->set_Width(200.0);
    builder->Write(u"Row 2, Cell 3 Content.");
    builder->EndRow();
    builder->EndTable();

    doc->Save(ArtifactsDir + u"DocumentBuilder.CreateFormattedTable.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.CreateFormattedTable.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASPOSE_ASSERT_EQ(20.0, table->get_LeftIndent());

    ASSERT_EQ(Aspose::Words::HeightRule::AtLeast, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    ASPOSE_ASSERT_EQ(40.0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());

    for (auto c : System::IterateOver<Cell>(doc->GetChildNodes(Aspose::Words::NodeType::Cell, true)))
    {
        ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, c->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());

        for (auto r : System::IterateOver<Run>(c->get_FirstParagraph()->get_Runs()))
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
    //ExSummary:Shows how to format table and cell with different borders and shadings.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Start a table and set a default color/thickness for its borders
    SharedPtr<Table> table = builder->StartTable();
    table->SetBorders(Aspose::Words::LineStyle::Single, 2.0, System::Drawing::Color::get_Black());

    // Set the cell shading for this cell
    builder->InsertCell();
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Red());
    builder->Writeln(u"Cell #1");

    // Specify a different cell shading for the second cell
    builder->InsertCell();
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Green());
    builder->Writeln(u"Cell #2");

    // End this row
    builder->EndRow();

    // Clear the cell formatting from previous operations
    builder->get_CellFormat()->ClearFormatting();

    // Create the second row
    builder->InsertCell();
    builder->Writeln(u"Cell #3");

    // Clear the cell formatting from the previous cell
    builder->get_CellFormat()->ClearFormatting();

    builder->get_CellFormat()->get_Borders()->get_Left()->set_LineWidth(4.0);
    builder->get_CellFormat()->get_Borders()->get_Right()->set_LineWidth(4.0);
    builder->get_CellFormat()->get_Borders()->get_Top()->set_LineWidth(4.0);
    builder->get_CellFormat()->get_Borders()->get_Bottom()->set_LineWidth(4.0);

    builder->InsertCell();
    builder->Writeln(u"Cell #4");

    doc->Save(ArtifactsDir + u"DocumentBuilder.TableBordersAndShading.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.TableBordersAndShading.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    for (auto c : System::IterateOver<Cell>(table->get_FirstRow()))
    {
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Top()->get_LineWidth());
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Bottom()->get_LineWidth());
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Left()->get_LineWidth());
        ASPOSE_ASSERT_EQ(0.5, c->get_CellFormat()->get_Borders()->get_Right()->get_LineWidth());

        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), c->get_CellFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::Single, c->get_CellFormat()->get_Borders()->get_Left()->get_LineStyle());
    }

    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_Cells()->idx_get(1)->get_CellFormat()->get_Shading()->get_BackgroundPatternColor().ToArgb());

    for (auto c : System::IterateOver<Cell>(table->get_LastRow()))
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
    //ExSummary:Shows how to specify a cell preferred width by converting inches to points.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Table> table = builder->StartTable();
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(ConvertUtil::InchToPoint(3)));
    builder->InsertCell();
    //ExEnd

    ASPOSE_ASSERT_EQ(216.0, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_PreferredWidth()->get_Value());
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
    //ExSummary:Shows how to insert a hyperlink referencing a local bookmark.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->StartBookmark(u"Bookmark1");
    builder->Write(u"Bookmarked text.");
    builder->EndBookmark(u"Bookmark1");

    builder->Writeln(u"Some other text");

    // Specify font formatting for the hyperlink
    builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    builder->get_Font()->set_Underline(Aspose::Words::Underline::Single);

    // Insert hyperlink
    // Switch \o is used to provide hyperlink tip text
    builder->InsertHyperlink(u"Hyperlink Text", u"Bookmark1\" \\o \"Hyperlink Tip", true);

    // Clear hyperlink formatting
    builder->get_Font()->ClearFormatting();

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
    auto hyperlink = System::DynamicCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));

    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldHyperlink, u" HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ", u"Hyperlink Text", hyperlink);
    ASSERT_EQ(u"Bookmark1", hyperlink->get_SubAddress());

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Bookmark> b)> _local_func_4 = [](SharedPtr<Bookmark> b)
    {
        return b->get_Name() == u"Bookmark1";
    };

    ASSERT_TRUE(doc->get_Range()->get_Bookmarks()->LINQ_Any(static_cast<System::Func<SharedPtr<Bookmark>, bool>>(_local_func_4)));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertHyperlinkToLocalBookmark)
{
    s_instance->InsertHyperlinkToLocalBookmark();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderCursorPosition()
{
    // Write some text in a blank Document using a DocumentBuilder
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Write(u"Hello world!");

    // If the builder's cursor is at the end of the document, there will be no nodes in front of it so the current node will be null
    ASSERT_TRUE(builder->get_CurrentNode() == nullptr);

    // However, the current paragraph the cursor is in will be valid
    ASSERT_EQ(u"Hello world!", builder->get_CurrentParagraph()->GetText().Trim());

    // Move to the beginning of the document and place the cursor at an existing node
    builder->MoveToDocumentStart();
    ASSERT_EQ(Aspose::Words::NodeType::Run, builder->get_CurrentNode()->get_NodeType());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderCursorPosition)
{
    s_instance->DocumentBuilderCursorPosition();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderMoveToNode()
{
    //ExStart
    //ExFor:Story.LastParagraph
    //ExFor:DocumentBuilder.MoveTo(Node)
    //ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Write a paragraph with the DocumentBuilder
    builder->Writeln(u"Text 1. ");

    // Move the DocumentBuilder to the first paragraph of the document and add another paragraph
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_LastParagraph(), builder->get_CurrentParagraph());
    //ExSkip
    builder->MoveTo(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0));
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), builder->get_CurrentParagraph());
    //ExSkip
    builder->Writeln(u"Text 2. ");

    // Since we moved to a node before the first paragraph before we added a second paragraph,
    // the second paragraph will appear in front of the first paragraph
    ASSERT_EQ(u"Text 2. \rText 1.", doc->GetText().Trim());

    // We can move the DocumentBuilder back to the end of the document like this
    // and carry on adding text to the end of the document
    builder->MoveTo(doc->get_FirstSection()->get_Body()->get_LastParagraph());
    builder->Writeln(u"Text 3. ");

    ASSERT_EQ(u"Text 2. \rText 1. \rText 3.", doc->GetText().Trim());
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_LastParagraph(), builder->get_CurrentParagraph());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderMoveToNode)
{
    s_instance->DocumentBuilderMoveToNode();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderMoveToDocumentStartEnd()
{
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->MoveToDocumentEnd();
    builder->Writeln(u"This is the end of the document.");

    builder->MoveToDocumentStart();
    builder->Writeln(u"This is the beginning of the document.");
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderMoveToDocumentStartEnd)
{
    s_instance->DocumentBuilderMoveToDocumentStartEnd();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderMoveToSection()
{
    // Create a blank document and append a section to it, giving it two sections
    auto doc = MakeObject<Document>();
    doc->AppendChild(MakeObject<Section>(doc));

    // Move a DocumentBuilder to the second section and add text
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->MoveToSection(1);
    builder->Writeln(u"Text added to the 2nd section.");
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderMoveToSection)
{
    s_instance->DocumentBuilderMoveToSection();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderMoveToParagraph()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToParagraph
    //ExSummary:Shows how to move a cursor position to the specified paragraph.
    // Open a document with a lot of paragraphs
    auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    ASSERT_EQ(22, paragraphs->get_Count());

    // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
    // and any content added by the DocumentBuilder will just be prepended to the document
    auto builder = MakeObject<DocumentBuilder>(doc);

    ASSERT_EQ(0, paragraphs->IndexOf(builder->get_CurrentParagraph()));

    // We can manually move the DocumentBuilder to any paragraph in the document via a 0-based index like this
    builder->MoveToParagraph(2, 0);
    ASSERT_EQ(2, paragraphs->IndexOf(builder->get_CurrentParagraph()));
    //ExSkip
    builder->Writeln(u"This is a new third paragraph. ");
    //ExEnd

    ASSERT_EQ(3, paragraphs->IndexOf(builder->get_CurrentParagraph()));

    doc = DocumentHelper::SaveOpen(doc);

    ASSERT_EQ(u"This is a new third paragraph.", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(2)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderMoveToParagraph)
{
    s_instance->DocumentBuilderMoveToParagraph();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderMoveToTableCell()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToCell
    //ExSummary:Shows how to move a cursor position to the specified table cell.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Move the builder to row 3, cell 4 of the first table
    builder->MoveToCell(0, 2, 3, 0);
    builder->Write(u"\nCell contents added by DocumentBuilder");
    //ExEnd

    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASPOSE_ASSERT_EQ(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(3), builder->get_CurrentNode()->get_ParentNode()->get_ParentNode());
    ASSERT_EQ(u"Cell contents added by DocumentBuilderCell 3 contents\a", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(3)->GetText().Trim());

}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderMoveToTableCell)
{
    s_instance->DocumentBuilderMoveToTableCell();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderMoveToBookmarkEnd()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
    //ExSummary:Shows how to move a cursor position to just after the bookmark end.
    auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Move to after the end of the first bookmark
    ASSERT_TRUE(builder->MoveToBookmark(u"MyBookmark1", false, true));
    builder->Write(u" Text appended via DocumentBuilder.");
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);

    ASSERT_FALSE(doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark1")->get_Text().Contains(u" Text appended via DocumentBuilder."));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderMoveToBookmarkEnd)
{
    s_instance->DocumentBuilderMoveToBookmarkEnd();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderBuildTable()
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
    //ExFor:Table.AutoFit
    //ExFor:AutoFitBehavior
    //ExSummary:Shows how to build a formatted table that contains 2 rows and 2 columns.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Table> table = builder->StartTable();

    // Insert a cell
    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);
    builder->Write(u"This is row 1 cell 1");

    // Use fixed column widths
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::FixedColumnWidths);

    // Insert a cell
    builder->InsertCell();
    builder->Write(u"This is row 1 cell 2");
    builder->EndRow();

    // Insert a cell
    builder->InsertCell();

    // Apply new row formatting
    builder->get_RowFormat()->set_Height(100);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);

    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Upward);
    builder->Write(u"This is row 2 cell 1");

    // Insert a cell
    builder->InsertCell();
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Downward);
    builder->Write(u"This is row 2 cell 2");

    builder->EndRow();
    builder->EndTable();
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(2, table->get_Rows()->get_Count());
    ASSERT_EQ(2, table->get_Rows()->idx_get(0)->get_Cells()->get_Count());
    ASSERT_EQ(2, table->get_Rows()->idx_get(1)->get_Cells()->get_Count());
    ASSERT_FALSE(table->get_AllowAutoFit());

    ASPOSE_ASSERT_EQ(0, table->get_Rows()->idx_get(0)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Auto, table->get_Rows()->idx_get(0)->get_RowFormat()->get_HeightRule());
    ASPOSE_ASSERT_EQ(100, table->get_Rows()->idx_get(1)->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_Rows()->idx_get(1)->get_RowFormat()->get_HeightRule());

    ASSERT_EQ(u"This is row 1 cell 1\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalAlignment());

    ASSERT_EQ(u"This is row 1 cell 2\a", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(1)->GetText().Trim());

    ASSERT_EQ(u"This is row 2 cell 1\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::TextOrientation::Upward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->get_CellFormat()->get_Orientation());

    ASSERT_EQ(u"This is row 2 cell 2\a", table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::TextOrientation::Downward, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1)->get_CellFormat()->get_Orientation());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderBuildTable)
{
    s_instance->DocumentBuilderBuildTable();
}

} // namespace gtest_test

void ExDocumentBuilder::TableCellVerticalRotatedFarEastTextOrientation()
{
    auto doc = MakeObject<Document>(MyDir + u"Rotated cell text.docx");

    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    SharedPtr<Cell> cell = table->get_FirstRow()->get_FirstCell();

    ASSERT_EQ(Aspose::Words::TextOrientation::VerticalRotatedFarEast, cell->get_CellFormat()->get_Orientation());

    doc = DocumentHelper::SaveOpen(doc);

    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
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

void ExDocumentBuilder::DocumentBuilderInsertBreak()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"This is page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    builder->Writeln(u"This is page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    builder->Writeln(u"This is page 3.");
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderInsertBreak)
{
    s_instance->DocumentBuilderInsertBreak();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderInsertInlineImage()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertImage(ImageDir + u"Transparent background logo.png");
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderInsertInlineImage)
{
    s_instance->DocumentBuilderInsertInlineImage();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderInsertFloatingImage()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert a floating image from a file or URL.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertImage(ImageDir + u"Transparent background logo.png", Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    auto image = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, image);
    ASPOSE_ASSERT_EQ(100.0, image->get_Left());
    ASPOSE_ASSERT_EQ(100.0, image->get_Top());
    ASPOSE_ASSERT_EQ(200.0, image->get_Width());
    ASPOSE_ASSERT_EQ(100.0, image->get_Height());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, image->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, image->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, image->get_RelativeVerticalPosition());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderInsertFloatingImage)
{
    s_instance->DocumentBuilderInsertFloatingImage();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertImageFromUrl()
{
    // Insert an image from a URL
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertImage(AsposeLogoUrl);

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertImageFromUrl.doc");

    // Verify that the image was inserted into the document
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASSERT_FALSE(shape == nullptr);
    ASSERT_TRUE(shape->get_HasImage());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertImageFromUrl)
{
    s_instance->InsertImageFromUrl();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertImageOriginalSize()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert a floating image from a file or URL and retain the original image size in the document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Pass a negative value to the width and height values to specify using the size of the source image
    builder->InsertImage(ImageDir + u"Logo.jpg", Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 200.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, -1.0, -1.0, Aspose::Words::Drawing::WrapType::Square);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    auto image = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, image);
    ASPOSE_ASSERT_EQ(200.0, image->get_Left());
    ASPOSE_ASSERT_EQ(100.0, image->get_Top());
    ASPOSE_ASSERT_EQ(270.3, image->get_Width());
    ASPOSE_ASSERT_EQ(270.3, image->get_Height());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, image->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, image->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, image->get_RelativeVerticalPosition());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertImageOriginalSize)
{
    s_instance->InsertImageOriginalSize();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderInsertTextInputFormField()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertTextInput
    //ExSummary:Shows how to insert a text input form field into a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertTextInput(u"TextInput", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Hello", 0);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);

    ASSERT_TRUE(formField->get_Enabled());
    ASSERT_EQ(u"TextInput", formField->get_Name());
    ASSERT_EQ(0, formField->get_MaxLength());
    ASSERT_EQ(u"Hello", formField->get_Result());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormTextInput, formField->get_Type());
    ASSERT_EQ(u"", formField->get_TextInputFormat());
    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, formField->get_TextInputType());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderInsertTextInputFormField)
{
    s_instance->DocumentBuilderInsertTextInputFormField();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderInsertComboBoxFormField()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertComboBox
    //ExSummary:Shows how to insert a combobox form field into a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    ArrayPtr<String> items = MakeArray<String>({u"One", u"Two", u"Three"});
    builder->InsertComboBox(u"DropDown", items, 0);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);

    ASSERT_TRUE(formField->get_Enabled());
    ASSERT_EQ(u"DropDown", formField->get_Name());
    ASSERT_EQ(0, formField->get_DropDownSelectedIndex());
    ASPOSE_ASSERT_EQ(MakeArray<String>({u"One", u"Two", u"Three"}), formField->get_DropDownItems());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormDropDown, formField->get_Type());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderInsertComboBoxFormField)
{
    s_instance->DocumentBuilderInsertComboBoxFormField();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderInsertToc()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table of contents at the beginning of the document
    builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");

    // The newly inserted table of contents will be initially empty
    // It needs to be populated by updating the fields in the document
    doc->UpdateFields();
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderInsertToc)
{
    s_instance->DocumentBuilderInsertToc();
}

} // namespace gtest_test

void ExDocumentBuilder::SignatureLineProviderId()
{
    //ExStart
    //ExFor:SignatureLine.ProviderId
    //ExFor:SignatureLineOptions.ShowDate
    //ExFor:SignatureLineOptions.Email
    //ExFor:SignatureLineOptions.DefaultInstructions
    //ExFor:SignatureLineOptions.Instructions
    //ExFor:SignatureLineOptions.AllowComments
    //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
    //ExFor:SignOptions.ProviderId
    //ExSummary:Shows how to sign document with personal certificate and specific signature line.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    auto signatureLineOptions = MakeObject<SignatureLineOptions>();
    signatureLineOptions->set_Signer(u"vderyushev");
    signatureLineOptions->set_SignerTitle(u"QA");
    signatureLineOptions->set_Email(u"vderyushev@aspose.com");
    signatureLineOptions->set_ShowDate(true);
    signatureLineOptions->set_DefaultInstructions(false);
    signatureLineOptions->set_Instructions(u"You need more info about signature line");
    signatureLineOptions->set_AllowComments(true);

    SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(signatureLineOptions)->get_SignatureLine();
    signatureLine->set_ProviderId(System::Guid::Parse(u"CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));

    doc->Save(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.docx");

    auto signOptions = MakeObject<SignOptions>();
    signOptions->set_SignatureLineId(signatureLine->get_Id());
    signOptions->set_ProviderId(signatureLine->get_ProviderId());
    signOptions->set_Comments(u"Document was signed by vderyushev");
    signOptions->set_SignTime(System::DateTime::get_Now());

    SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

    DigitalSignatureUtil::Sign(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.docx", ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.Signed.docx", certHolder, signOptions);
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.Signed.docx");
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    signatureLine = shape->get_SignatureLine();

    ASSERT_EQ(u"vderyushev", signatureLine->get_Signer());
    ASSERT_EQ(u"QA", signatureLine->get_SignerTitle());
    ASSERT_EQ(u"vderyushev@aspose.com", signatureLine->get_Email());
    ASSERT_TRUE(signatureLine->get_ShowDate());
    ASSERT_FALSE(signatureLine->get_DefaultInstructions());
    ASSERT_EQ(u"You need more info about signature line", signatureLine->get_Instructions());
    ASSERT_TRUE(signatureLine->get_AllowComments());
    ASSERT_TRUE(signatureLine->get_IsSigned());
    ASSERT_TRUE(signatureLine->get_IsValid());

    SharedPtr<DigitalSignatureCollection> signatures = DigitalSignatureUtil::LoadSignatures(ArtifactsDir + u"DocumentBuilder.SignatureLineProviderId.Signed.docx");

    ASSERT_EQ(1, signatures->get_Count());
    ASSERT_TRUE(signatures->idx_get(0)->get_IsValid());
    ASSERT_EQ(u"Document was signed by vderyushev", signatures->idx_get(0)->get_Comments());
    ASSERT_EQ(System::DateTime::get_Today(), signatures->idx_get(0)->get_SignTime().get_Date());
    ASSERT_EQ(u"CN=Morzal.Me", signatures->idx_get(0)->get_IssuerName());
    ASSERT_EQ(Aspose::Words::DigitalSignatureType::XmlDsig, signatures->idx_get(0)->get_SignatureType());

}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, SignatureLineProviderId)
{
    s_instance->SignatureLineProviderId();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertSignatureLineCurrentPosition()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
    //ExSummary:Shows how to insert signature line at the specified position.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    auto options = MakeObject<SignatureLineOptions>();
    options->set_Signer(u"John Doe");
    options->set_SignerTitle(u"Manager");
    options->set_Email(u"johndoe@aspose.com");
    options->set_ShowDate(true);
    options->set_DefaultInstructions(false);
    options->set_Instructions(u"You need more info about signature line");
    options->set_AllowComments(true);

    builder->InsertSignatureLine(options, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, 2.0, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 3.0, Aspose::Words::Drawing::WrapType::Inline);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);

    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    SharedPtr<SignatureLine> signatureLine = shape->get_SignatureLine();

    ASSERT_EQ(u"John Doe", signatureLine->get_Signer());
    ASSERT_EQ(u"Manager", signatureLine->get_SignerTitle());
    ASSERT_EQ(u"johndoe@aspose.com", signatureLine->get_Email());
    ASSERT_TRUE(signatureLine->get_ShowDate());
    ASSERT_FALSE(signatureLine->get_DefaultInstructions());
    ASSERT_EQ(u"You need more info about signature line", signatureLine->get_Instructions());
    ASSERT_TRUE(signatureLine->get_AllowComments());
    ASSERT_FALSE(signatureLine->get_IsSigned());
    ASSERT_FALSE(signatureLine->get_IsValid());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertSignatureLineCurrentPosition)
{
    s_instance->InsertSignatureLineCurrentPosition();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderSetFontFormatting()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set font formatting properties
    SharedPtr<Aspose::Words::Font> font = builder->get_Font();
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_DarkBlue());
    font->set_Italic(true);
    font->set_Name(u"Arial");
    font->set_Size(24);
    font->set_Spacing(5);
    font->set_Underline(Aspose::Words::Underline::Double);

    // Output formatted text
    builder->Writeln(u"I'm a very nice formatted String.");
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderSetFontFormatting)
{
    s_instance->DocumentBuilderSetFontFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderSetParagraphFormatting()
{
    //ExStart
    //ExFor:ParagraphFormat.RightIndent
    //ExFor:ParagraphFormat.LeftIndent
    //ExSummary:Shows how to set paragraph formatting.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set paragraph formatting properties
    SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
    paragraphFormat->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    paragraphFormat->set_LeftIndent(50);
    paragraphFormat->set_RightIndent(50);
    paragraphFormat->set_SpaceAfter(25);

    // Output text
    builder->Writeln(u"This paragraph demonstrates how the left and right indents affect word wrapping.");
    builder->Writeln(u"The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.");

    doc->Save(ArtifactsDir + u"DocumentBuilder.DocumentBuilderSetParagraphFormatting.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.DocumentBuilderSetParagraphFormatting.docx");

    for (auto paragraph : System::IterateOver<Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
    {
        ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, paragraph->get_ParagraphFormat()->get_Alignment());
        ASPOSE_ASSERT_EQ(50.0, paragraph->get_ParagraphFormat()->get_LeftIndent());
        ASPOSE_ASSERT_EQ(50.0, paragraph->get_ParagraphFormat()->get_RightIndent());
        ASPOSE_ASSERT_EQ(25.0, paragraph->get_ParagraphFormat()->get_SpaceAfter());

    }
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderSetParagraphFormatting)
{
    s_instance->DocumentBuilderSetParagraphFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderSetCellFormatting()
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
    //ExSummary:Shows how to create a table that contains a single formatted cell.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->StartTable();
    builder->InsertCell();

    // Set the cell formatting
    SharedPtr<CellFormat> cellFormat = builder->get_CellFormat();
    cellFormat->set_Width(250);
    cellFormat->set_LeftPadding(30);
    cellFormat->set_RightPadding(30);
    cellFormat->set_TopPadding(30);
    cellFormat->set_BottomPadding(30);

    builder->Write(u"Formatted cell");
    builder->EndRow();
    builder->EndTable();

    doc->Save(ArtifactsDir + u"DocumentBuilder.DocumentBuilderSetCellFormatting.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.DocumentBuilderSetCellFormatting.docx");
    SharedPtr<Cell> firstCell = (System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true)))->get_FirstRow()->get_FirstCell();

    ASSERT_EQ(u"Formatted cell\a", firstCell->GetText().Trim());

    ASPOSE_ASSERT_EQ(250.0, firstCell->get_CellFormat()->get_Width());
    ASPOSE_ASSERT_EQ(30.0, firstCell->get_CellFormat()->get_LeftPadding());
    ASPOSE_ASSERT_EQ(30.0, firstCell->get_CellFormat()->get_RightPadding());
    ASPOSE_ASSERT_EQ(30.0, firstCell->get_CellFormat()->get_TopPadding());
    ASPOSE_ASSERT_EQ(30.0, firstCell->get_CellFormat()->get_BottomPadding());

}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderSetCellFormatting)
{
    s_instance->DocumentBuilderSetCellFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderSetRowFormatting()
{
    //ExStart
    //ExFor:DocumentBuilder.RowFormat
    //ExFor:HeightRule
    //ExFor:RowFormat.Height
    //ExFor:RowFormat.HeightRule
    //ExFor:Table.LeftPadding
    //ExFor:Table.RightPadding
    //ExFor:Table.TopPadding
    //ExFor:Table.BottomPadding
    //ExSummary:Shows how to create a table that contains a single cell and apply row formatting.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();

    // Set the row formatting
    SharedPtr<RowFormat> rowFormat = builder->get_RowFormat();
    rowFormat->set_Height(100);
    rowFormat->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    // These formatting properties are set on the table and are applied to all rows in the table
    table->set_LeftPadding(30);
    table->set_RightPadding(30);
    table->set_TopPadding(30);
    table->set_BottomPadding(30);

    builder->Writeln(u"Contents of formatted row.");

    builder->EndRow();
    builder->EndTable();

    doc->Save(ArtifactsDir + u"DocumentBuilder.DocumentBuilderSetRowFormatting.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.DocumentBuilderSetRowFormatting.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASPOSE_ASSERT_EQ(30.0, table->get_LeftPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_RightPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_TopPadding());
    ASPOSE_ASSERT_EQ(30.0, table->get_BottomPadding());

    ASPOSE_ASSERT_EQ(100.0, table->get_FirstRow()->get_RowFormat()->get_Height());
    ASSERT_EQ(Aspose::Words::HeightRule::Exactly, table->get_FirstRow()->get_RowFormat()->get_HeightRule());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderSetRowFormatting)
{
    s_instance->DocumentBuilderSetRowFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderSetListFormatting()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->get_ListFormat()->ApplyNumberDefault();

    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");

    builder->get_ListFormat()->ListIndent();

    builder->Writeln(u"Item 2.1");
    builder->Writeln(u"Item 2.2");

    builder->get_ListFormat()->ListIndent();

    builder->Writeln(u"Item 2.2.1");
    builder->Writeln(u"Item 2.2.2");

    builder->get_ListFormat()->ListOutdent();

    builder->Writeln(u"Item 2.3");

    builder->get_ListFormat()->ListOutdent();

    builder->Writeln(u"Item 3");

    builder->get_ListFormat()->RemoveNumbers();
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderSetListFormatting)
{
    s_instance->DocumentBuilderSetListFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderSetSectionFormatting()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set page properties
    builder->get_PageSetup()->set_Orientation(Aspose::Words::Orientation::Landscape);
    builder->get_PageSetup()->set_LeftMargin(50);
    builder->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Paper10x14);
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderSetSectionFormatting)
{
    s_instance->DocumentBuilderSetSectionFormatting();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertFootnote()
{
    //ExStart
    //ExFor:FootnoteType
    //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String)
    //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String,String)
    //ExSummary:Shows how to reference text with a footnote and an endnote.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert some text and mark it with a footnote with the IsAuto attribute set to "true" by default,
    // so the marker seen in the body text will be auto-numbered at "1", and the footnote will appear at the bottom of the page
    builder->Write(u"This text will be referenced by a footnote.");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Footnote, u"Footnote comment regarding referenced text.");

    // Insert more text and mark it with an endnote with a custom reference mark,
    // which will be used in place of the number "2" and will set "IsAuto" to false
    builder->Write(u"This text will be referenced by an endnote.");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Endnote, u"Endnote comment regarding referenced text.", u"CustomMark");

    // Footnotes always appear at the bottom of the page of their referenced text, so this page break will not affect the footnote
    // On the other hand, endnotes are always at the end of the document, so this page break will push the endnote down to the next page
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertFootnote.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertFootnote.docx");

    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Footnote, true, String::Empty, u"Footnote comment regarding referenced text.", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Endnote, false, u"CustomMark", u"CustomMark Endnote comment regarding referenced text.", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertFootnote)
{
    s_instance->InsertFootnote();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderApplyParagraphStyle()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set paragraph style
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Title);

    builder->Write(u"Hello");
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DocumentBuilderApplyParagraphStyle)
{
    s_instance->DocumentBuilderApplyParagraphStyle();
}

} // namespace gtest_test

void ExDocumentBuilder::DocumentBuilderApplyBordersAndShading()
{
    //ExStart
    //ExFor:BorderCollection.Item(BorderType)
    //ExFor:Shading
    //ExFor:TextureIndex
    //ExFor:ParagraphFormat.Shading
    //ExFor:Shading.Texture
    //ExFor:Shading.BackgroundPatternColor
    //ExFor:Shading.ForegroundPatternColor
    //ExSummary:Shows how to apply borders and shading to a paragraph.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set paragraph borders
    SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
    borders->set_DistanceFromText(20);
    borders->idx_get(Aspose::Words::BorderType::Left)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Right)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Top)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Bottom)->set_LineStyle(Aspose::Words::LineStyle::Double);

    // Set paragraph shading
    SharedPtr<Shading> shading = builder->get_ParagraphFormat()->get_Shading();
    shading->set_Texture(Aspose::Words::TextureIndex::TextureDiagonalCross);
    shading->set_BackgroundPatternColor(System::Drawing::Color::get_LightCoral());
    shading->set_ForegroundPatternColor(System::Drawing::Color::get_LightSalmon());

    builder->Write(u"This paragraph is formatted with a double border and shading.");
    doc->Save(ArtifactsDir + u"DocumentBuilder.DocumentBuilderApplyBordersAndShading.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.DocumentBuilderApplyBordersAndShading.docx");
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

TEST_F(ExDocumentBuilder, DocumentBuilderApplyBordersAndShading)
{
    s_instance->DocumentBuilderApplyBordersAndShading();
}

} // namespace gtest_test

void ExDocumentBuilder::DeleteRow()
{
    //ExStart
    //ExFor:DocumentBuilder.DeleteRow
    //ExSummary:Shows how to delete a row from a table.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table with 2 rows
    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Cell 1");
    builder->InsertCell();
    builder->Write(u"Cell 2");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Cell 3");
    builder->InsertCell();
    builder->Write(u"Cell 4");
    builder->EndTable();

    ASSERT_EQ(2, table->get_Rows()->get_Count());

    // Delete the first row of the first table in the document
    builder->DeleteRow(0, 0);

    ASSERT_EQ(1, table->get_Rows()->get_Count());
    //ExEnd

    ASSERT_EQ(u"Cell 3\aCell 4\a\a", table->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DeleteRow)
{
    s_instance->DeleteRow();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertDocument()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode)
    //ExFor:ImportFormatMode
    //ExSummary:Shows how to insert a document content into another document keep formatting of inserted document.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    auto docToInsert = MakeObject<Document>(MyDir + u"Formatted elements.docx");

    builder->InsertDocument(docToInsert, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.InsertDocument.docx");
    //ExEnd

    ASSERT_EQ(29, doc->get_Styles()->get_Count());
    ASSERT_TRUE(DocumentHelper::CompareDocs(ArtifactsDir + u"DocumentBuilder.InsertDocument.docx", GoldsDir + u"DocumentBuilder.InsertDocument Gold.docx"));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, DISABLED_InsertDocument)
{
    s_instance->InsertDocument();
}

} // namespace gtest_test

void ExDocumentBuilder::KeepSourceNumbering()
{
    //ExStart
    //ExFor:ImportFormatOptions.KeepSourceNumbering
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how the numbering will be imported when it clashes in source and destination documents.
    // Open a document with a custom list numbering scheme and clone it
    // Since both have the same numbering format, the formats will clash if we import one document into the other
    auto srcDoc = MakeObject<Document>(MyDir + u"Custom list numbering.docx");
    SharedPtr<Document> dstDoc = srcDoc->Clone();

    // Both documents have the same numbering in their lists, but if we set this flag to false and then import one document into the other
    // the numbering of the imported source document will continue from where it ends in the destination document
    auto importFormatOptions = MakeObject<ImportFormatOptions>();
    importFormatOptions->set_KeepSourceNumbering(false);

    auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, Aspose::Words::ImportFormatMode::KeepDifferentStyles, importFormatOptions);
    for (auto paragraph : System::IterateOver<Paragraph>(srcDoc->get_FirstSection()->get_Body()->get_Paragraphs()))
    {
        SharedPtr<Node> importedNode = importer->ImportNode(paragraph, true);
        dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
    }

    dstDoc->UpdateListLabels();
    dstDoc->Save(ArtifactsDir + u"DocumentBuilder.KeepSourceNumbering.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, KeepSourceNumbering)
{
    s_instance->KeepSourceNumbering();
}

} // namespace gtest_test

void ExDocumentBuilder::ResolveStyleBehaviorWhileAppendDocument()
{
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to resolve styles behavior while append document.
    // Open a document with text in a custom style and clone it
    auto srcDoc = MakeObject<Document>(MyDir + u"Custom list numbering.docx");
    SharedPtr<Document> dstDoc = srcDoc->Clone();

    // We now have two documents, each with an identical style named "CustomStyle"
    // We can change the text color of one of the styles
    dstDoc->get_Styles()->idx_get(u"CustomStyle")->get_Font()->set_Color(System::Drawing::Color::get_DarkRed());

    auto options = MakeObject<ImportFormatOptions>();
    // Specify that if numbering clashes in source and destination documents
    // then a numbering from the source document will be used
    options->set_KeepSourceNumbering(true);

    // If we join two documents which have different styles that share the same name,
    // we can resolve the style clash with an ImportFormatMode
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepDifferentStyles, options);
    dstDoc->UpdateListLabels();

    dstDoc->Save(ArtifactsDir + u"DocumentBuilder.ResolveStyleBehaviorWhileAppendDocument.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, ResolveStyleBehaviorWhileAppendDocument)
{
    s_instance->ResolveStyleBehaviorWhileAppendDocument();
}

} // namespace gtest_test

void ExDocumentBuilder::IgnoreTextBoxes(bool isIgnoreTextBoxes)
{
    //ExStart
    //ExFor:ImportFormatOptions.IgnoreTextBoxes
    //ExSummary:Shows how to manage formatting in the text boxes of the source destination during the import.
    // Create a document and add text
    auto dstDoc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(dstDoc);

    builder->Writeln(u"Hello world! Text box to follow.");

    // Create another document with a textbox, and insert some formatted text into it
    auto srcDoc = MakeObject<Document>();
    builder = MakeObject<DocumentBuilder>(srcDoc);

    SharedPtr<Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 300, 100);
    builder->MoveTo(textBox->get_FirstParagraph());
    builder->get_ParagraphFormat()->get_Style()->get_Font()->set_Name(u"Courier New");
    builder->get_ParagraphFormat()->get_Style()->get_Font()->set_Size(24.0);
    builder->Write(u"Textbox contents");

    // When we import the document with the textbox as a node into the first document, by default the text inside the text box will keep its formatting
    // Setting the IgnoreTextBoxes flag will clear the formatting during importing of the node
    auto importFormatOptions = MakeObject<ImportFormatOptions>();
    importFormatOptions->set_IgnoreTextBoxes(isIgnoreTextBoxes);

    auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, importFormatOptions);

    for (auto paragraph : System::IterateOver<Paragraph>(srcDoc->get_FirstSection()->get_Body()->get_Paragraphs()))
    {
        SharedPtr<Node> importedNode = importer->ImportNode(paragraph, true);
        dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
    }

    dstDoc->Save(ArtifactsDir + u"DocumentBuilder.IgnoreTextBoxes.docx");
    //ExEnd

    dstDoc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.IgnoreTextBoxes.docx");
    textBox = System::DynamicCast<Aspose::Words::Drawing::Shape>(dstDoc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASSERT_EQ(u"Textbox contents", textBox->GetText().Trim());

    if (isIgnoreTextBoxes)
    {
        ASPOSE_ASSERT_EQ(12.0, textBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Size());
        ASSERT_EQ(u"Times New Roman", textBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
    }
    else
    {
        ASPOSE_ASSERT_EQ(24.0, textBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Size());
        ASSERT_EQ(u"Courier New", textBox->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
    }
}

namespace gtest_test
{

using IgnoreTextBoxes_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::IgnoreTextBoxes)>::type;

struct IgnoreTextBoxesVP : public ExDocumentBuilder, public ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<IgnoreTextBoxes_Args>
{
    static std::vector<IgnoreTextBoxes_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(IgnoreTextBoxesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreTextBoxes(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocumentBuilder, IgnoreTextBoxesVP, ::testing::ValuesIn(IgnoreTextBoxesVP::TestCases()));

} // namespace gtest_test

void ExDocumentBuilder::MoveToField()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToField
    //ExSummary:Shows how to move document builder's cursor to a specific field.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a field using the DocumentBuilder and add a run of text after it
    SharedPtr<Field> field = builder->InsertField(u"MERGEFIELD field");
    builder->Write(u" Text after the field.");

    // The builder's cursor is currently at end of the document
    ASSERT_TRUE(builder->get_CurrentNode() == nullptr);

    // We can move the builder to a field like this, placing the cursor at immediately after the field
    builder->MoveToField(field, true);

    // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field
    // If we wish to move the DocumentBuilder to inside a field,
    // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method
    ASPOSE_ASSERT_EQ(field->get_End(), builder->get_CurrentNode()->get_PreviousSibling());

    builder->Write(u" Text immediately after the field.");
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);

    ASSERT_EQ(u"\u0013MERGEFIELD field\u0014«field»\u0015 Text immediately after the field. Text after the field.", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MoveToField)
{
    s_instance->MoveToField();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertOleObjectException()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_5 = [&builder]()
    {
        builder->InsertOleObject(u"", u"checkbox", false, true, nullptr);
    };

    ASSERT_THROW(_local_func_5(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOleObjectException)
{
    s_instance->InsertOleObjectException();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertChartDouble()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
    //ExSummary:Shows how to insert a chart into a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, ConvertUtil::PixelToPoint(300), ConvertUtil::PixelToPoint(300));
    ASPOSE_ASSERT_EQ(225.0, ConvertUtil::PixelToPoint(300));
    //ExSkip

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertedChartDouble.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertedChartDouble.docx");
    auto chartShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASSERT_EQ(u"Chart Title", chartShape->get_Chart()->get_Title()->get_Text());
    ASPOSE_ASSERT_EQ(225.0, chartShape->get_Width());
    ASPOSE_ASSERT_EQ(225.0, chartShape->get_Height());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertChartDouble)
{
    s_instance->InsertChartDouble();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertChartRelativePosition()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert a chart into a document and specify position and size.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100, 200, 100, Aspose::Words::Drawing::WrapType::Square);

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertedChartRelativePosition.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertedChartRelativePosition.docx");
    auto chartShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

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
    //ExFor:Field.Update()
    //ExFor:Field.Result
    //ExFor:Field.GetFieldCode()
    //ExFor:Field.Type
    //ExFor:Field.Remove
    //ExFor:FieldType
    //ExSummary:Shows how to insert a field into a document by FieldCode.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a simple Date field into the document
    // When we insert a field through the DocumentBuilder class we can get the
    // special Field object which contains information about the field
    SharedPtr<Field> dateField = builder->InsertField(u"DATE \\* MERGEFORMAT");

    // Update this particular field in the document so we can get the FieldResult
    dateField->Update();

    // Display some information from this field
    // The field result is where the last evaluated value is stored. This is what is displayed in the document
    // When field codes are not showing
    ASSERT_EQ(System::DateTime::get_Today(), System::DateTime::Parse(dateField->get_Result()));

    // Display the field code which defines the behavior of the field. This can been seen in Microsoft Word by pressing ALT+F9
    ASSERT_EQ(u"DATE \\* MERGEFORMAT", dateField->GetFieldCode());

    // The field type defines what type of field in the Document this is. In this case the type is "FieldDate"
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, dateField->get_Type());

    // Finally let's completely remove the field from the document. This can easily be done by invoking the Remove method on the object
    dateField->Remove();
    //ExEnd

    ASSERT_EQ(0, doc->get_Range()->get_Fields()->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertField)
{
    s_instance->InsertField();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertFieldByType()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
    //ExSummary:Shows how to insert a field into a document using FieldType.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert an AUTHOR field using a DocumentBuilder
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    builder->Write(u"This document was written by ");
    builder->InsertField(Aspose::Words::Fields::FieldType::FieldAuthor, true);
    ASSERT_EQ(u" AUTHOR ", doc->get_Range()->get_Fields()->idx_get(0)->GetFieldCode());
    //ExSkip
    ASSERT_EQ(u"John Doe", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
    //ExSkip

    // Insert a PAGE field using a DocumentBuilder, but do not immediately update it
    builder->Write(u"\nThis is page ");
    builder->InsertField(Aspose::Words::Fields::FieldType::FieldPage, false);
    ASSERT_EQ(u" PAGE ", doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode());
    //ExSkip
    ASSERT_EQ(u"", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
    //ExSkip

    // Some fields types, such as ones that display document word/page counts may not keep track of their results in real time,
    // and will only display an accurate result during a field update
    // We can defer the updating of those fields until right before we need to see an accurate result
    // This method will manually update all the fields in a document
    doc->UpdateFields();

    ASSERT_EQ(u"1", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);

    ASSERT_EQ(String(u"This document was written by \u0013 AUTHOR \u0014John Doe\u0015") + u"\rThis is page \u0013 PAGE \u00141\u0015", doc->GetText().Trim());

    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldAuthor, u" AUTHOR ", u"John Doe", doc->get_Range()->get_Fields()->idx_get(0));
    TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldPage, u" PAGE ", u"1", doc->get_Range()->get_Fields()->idx_get(1));
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertFieldByType)
{
    s_instance->InsertFieldByType();
}

} // namespace gtest_test

void ExDocumentBuilder::FieldResultFormatting()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    doc->get_FieldOptions()->set_ResultFormatter(MakeObject<ExDocumentBuilder::FieldResultFormatter>(u"${0}", u"Date: {0}", u"Item # {0}:"));

    // Insert a field with a numeric format
    builder->InsertField(u" = 2 + 3 \\# $###", nullptr);

    // Insert a field with a date/time format
    builder->InsertField(u"DATE \\@ \"d MMMM yyyy\"", nullptr);

    // Insert a field with a general format
    builder->InsertField(u"QUOTE \"2\" \\* Ordinal", nullptr);

    // Formats will be applied and recorded by the formatter during the field update
    doc->UpdateFields();
    (System::StaticCast<ApiExamples::ExDocumentBuilder::FieldResultFormatter>(doc->get_FieldOptions()->get_ResultFormatter()))->PrintInvocations();

    // Our formatter has also overridden the formats that were originally applied in the fields
    ASSERT_EQ(u"$5", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
    ASSERT_TRUE(doc->get_Range()->get_Fields()->idx_get(1)->get_Result().StartsWith(u"Date: "));
    ASSERT_EQ(u"Item # 2:", doc->get_Range()->get_Fields()->idx_get(2)->get_Result());
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
    //ExSummary:Shows how to insert online video into a document using video url
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a video from Youtube
    builder->InsertOnlineVideo(u"https://youtu.be/t_1LYZ102RA", 360, 270);

    // Click on the shape in the output document to watch the video from Microsoft Word
    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertVideoWithUrl.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertVideoWithUrl.docx");
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(480, 360, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, shape->get_HRef());

    ASPOSE_ASSERT_EQ(360.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(270.0, shape->get_Height());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertVideoWithUrl)
{
    s_instance->InsertVideoWithUrl();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertUnderline()
{
    //ExStart
    //ExFor:DocumentBuilder.Underline
    //ExSummary:Shows how to set and edit a document builder's underline.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set a new style for our underline
    builder->set_Underline(Aspose::Words::Underline::Dash);

    // Same object as DocumentBuilder.Font.Underline
    ASSERT_EQ(builder->get_Underline(), builder->get_Font()->get_Underline());
    ASSERT_EQ(Aspose::Words::Underline::Dash, builder->get_Font()->get_Underline());

    // These properties will be applied to the underline as well
    builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    builder->get_Font()->set_Size(32);

    builder->Writeln(u"Underlined text.");

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertUnderline.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertUnderline.docx");
    SharedPtr<Run> firstRun = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

    ASSERT_EQ(u"Underlined text.", firstRun->GetText().Trim());
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // A Story is a type of node that have child Paragraph nodes, such as a Body,
    // which would usually be a parent node to a DocumentBuilder's current paragraph
    ASPOSE_ASSERT_EQ(builder->get_CurrentStory(), doc->get_FirstSection()->get_Body());
    ASPOSE_ASSERT_EQ(builder->get_CurrentStory(), builder->get_CurrentParagraph()->get_ParentNode());
    ASSERT_EQ(Aspose::Words::StoryType::MainText, builder->get_CurrentStory()->get_StoryType());

    builder->get_CurrentStory()->AppendParagraph(u"Text added to current Story.");

    // A Story can contain tables too
    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1 cell 1");
    builder->InsertCell();
    builder->Write(u"Row 1 cell 2");
    builder->EndTable();

    // The table we just made is automatically placed in the story
    ASSERT_TRUE(builder->get_CurrentStory()->get_Tables()->Contains(table));
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Tables()->get_Count());
    ASSERT_EQ(u"Row 1 cell 1\aRow 1 cell 2\a\a\rText added to current Story.", doc->get_FirstSection()->get_Body()->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, CurrentStory)
{
    s_instance->CurrentStory();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertOlePowerpoint()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertOleObject(Stream, String, Boolean, Image)
    //ExSummary:Shows how to use document builder to embed Ole objects in a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Let's take a spreadsheet from our system and insert it into the document
    {
        SharedPtr<System::IO::Stream> spreadsheetStream = System::IO::File::Open(MyDir + u"Spreadsheet.xlsx", System::IO::FileMode::Open);
        // The spreadsheet can be activated by double clicking the panel that you'll see in the document immediately under the text we will add
        // We did not set the area to double click as an icon nor did we change its appearance so it looks like a simple panel
        builder->Writeln(u"Spreadsheet Ole object:");
        builder->InsertOleObject(spreadsheetStream, u"OleObject.xlsx", false, nullptr);

        // A powerpoint presentation is another type of object we can embed in our document
        // This time we'll also exercise some control over how it looks
        {
            SharedPtr<System::IO::Stream> powerpointStream = System::IO::File::Open(MyDir + u"Presentation.pptx", System::IO::FileMode::Open);
            // If we insert the Ole object as an icon, we are still provided with a default icon
            // If that is not suitable, we can make the icon to look like any image
            {
                auto webClient = MakeObject<System::Net::WebClient>();
                ArrayPtr<uint8_t> imgBytes = webClient->DownloadData(AsposeLogoUrl);

                {
                    auto stream = MakeObject<System::IO::MemoryStream>(imgBytes);
                    // If we double click the image, the powerpoint presentation will open
                    builder->InsertParagraph();
                    builder->Writeln(u"Powerpoint Ole object:");
                    builder->InsertOleObject(powerpointStream, u"OleObject.pptx", true, stream);
                }

            }
        }
    }

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertOlePowerpoint.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertOlePowerpoint.docx");

    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASSERT_EQ(u"", shape->get_OleFormat()->get_IconCaption());
    ASSERT_FALSE(shape->get_OleFormat()->get_OleIcon());

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    ASSERT_EQ(u"Unknown", shape->get_OleFormat()->get_IconCaption());
    ASSERT_TRUE(shape->get_OleFormat()->get_OleIcon());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOlePowerpoint)
{
    s_instance->InsertOlePowerpoint();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertStyleSeparator()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertStyleSeparator
    //ExSummary:Shows how to separate styles from two different paragraphs used in one logical printed paragraph.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Append text in the "Heading 1" style
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Write(u"This text is in a Heading style. ");

    // Insert a style separator
    builder->InsertStyleSeparator();

    // The style separator appears in the form of a paragraph break that doesn't start a new line
    // So, while this looks like one continuous paragraph with two styles in the output document,
    // it is actually two paragraphs with different styles, but no line break between the first and second paragraph
    ASSERT_EQ(2, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());

    // Append text with another style
    SharedPtr<Style> paraStyle = builder->get_Document()->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyParaStyle");
    paraStyle->get_Font()->set_Bold(false);
    paraStyle->get_Font()->set_Size(8);
    paraStyle->get_Font()->set_Name(u"Arial");

    // Set the style of the current paragraph to our custom style
    // This will apply to only the text after the style separator
    builder->get_ParagraphFormat()->set_StyleName(paraStyle->get_Name());
    builder->Write(u"This text is in a custom style. ");

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertStyleSeparator.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertStyleSeparator.docx");

    ASSERT_EQ(2, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
    ASSERT_EQ(u"This text is in a Heading style. \r This text is in a custom style.", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertStyleSeparator)
{
    s_instance->InsertStyleSeparator();
}

} // namespace gtest_test

void ExDocumentBuilder::WithoutStyleSeparator()
{
    auto builder = MakeObject<DocumentBuilder>(MakeObject<Document>());

    SharedPtr<Style> paraStyle = builder->get_Document()->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyParaStyle");
    paraStyle->get_Font()->set_Bold(false);
    paraStyle->get_Font()->set_Size(8);
    paraStyle->get_Font()->set_Name(u"Arial");

    // Append text with "Heading 1" style
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Write(u"This text is in a Heading style. ");

    // Append text with another style
    builder->get_ParagraphFormat()->set_StyleName(paraStyle->get_Name());
    builder->Write(u"This text is in a custom style. ");

    builder->get_Document()->Save(ArtifactsDir + u"DocumentBuilder.WithoutStyleSeparator.docx");
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, WithoutStyleSeparator)
{
    s_instance->WithoutStyleSeparator();
}

} // namespace gtest_test

void ExDocumentBuilder::SmartStyleBehavior()
{
    //ExStart
    //ExFor:ImportFormatOptions
    //ExFor:ImportFormatOptions.SmartStyleBehavior
    //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to resolve styles behavior while inserting documents.
    auto dstDoc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(dstDoc);

    SharedPtr<Style> myStyle = builder->get_Document()->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyStyle");
    myStyle->get_Font()->set_Size(14);
    myStyle->get_Font()->set_Name(u"Courier New");
    myStyle->get_Font()->set_Color(System::Drawing::Color::get_Blue());

    // Append text with custom style
    builder->get_ParagraphFormat()->set_StyleName(myStyle->get_Name());
    builder->Writeln(u"Hello world!");

    // Clone the document, and edit the clone's "MyStyle" style so it is a different color than that of the original
    // If we append this document to the original, the different styles will clash since they are the same name, and we will need to resolve it
    SharedPtr<Document> srcDoc = dstDoc->Clone();
    srcDoc->get_Styles()->idx_get(u"MyStyle")->get_Font()->set_Color(System::Drawing::Color::get_Red());

    // When SmartStyleBehavior is enabled,
    // a source style will be expanded into a direct attributes inside a destination document,
    // if KeepSourceFormatting importing mode is used
    auto options = MakeObject<ImportFormatOptions>();
    options->set_SmartStyleBehavior(true);

    builder->InsertDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, options);

    dstDoc->Save(ArtifactsDir + u"DocumentBuilder.SmartStyleBehavior.docx");
    //ExEnd

    dstDoc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SmartStyleBehavior.docx");

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

void ExDocumentBuilder::MarkdownDocumentTests()
{
    MarkdownDocumentEmphases();
    MarkdownDocumentInlineCode();
    MarkdownDocumentHeadings();
    MarkdownDocumentBlockquotes();
    MarkdownDocumentIndentedCode();
    MarkdownDocumentFencedCode();
    MarkdownDocumentHorizontalRule();
    MarkdownDocumentBulletedList();
    LoadMarkdownDocumentAndAssertContent(u"Italic", u"Normal", true, false);
    LoadMarkdownDocumentAndAssertContent(u"Bold", u"Normal", false, true);
    LoadMarkdownDocumentAndAssertContent(u"ItalicBold", u"Normal", true, true);
    LoadMarkdownDocumentAndAssertContent(u"Text with InlineCode style with one backtick", u"InlineCode", false, false);
    LoadMarkdownDocumentAndAssertContent(u"Text with InlineCode style with 3 backticks", u"InlineCode.3", false, false);
    LoadMarkdownDocumentAndAssertContent(u"This is an italic H1 tag", u"Heading 1", true, false);
    LoadMarkdownDocumentAndAssertContent(u"SetextHeading 1", u"SetextHeading1", false, false);
    LoadMarkdownDocumentAndAssertContent(u"This is an H2 tag", u"Heading 2", false, false);
    LoadMarkdownDocumentAndAssertContent(u"SetextHeading 2", u"SetextHeading2", false, false);
    LoadMarkdownDocumentAndAssertContent(u"This is an H3 tag", u"Heading 3", false, false);
    LoadMarkdownDocumentAndAssertContent(u"This is an bold H4 tag", u"Heading 4", false, true);
    LoadMarkdownDocumentAndAssertContent(u"This is an italic and bold H5 tag", u"Heading 5", true, true);
    LoadMarkdownDocumentAndAssertContent(u"This is an H6 tag", u"Heading 6", false, false);
    LoadMarkdownDocumentAndAssertContent(u"Blockquote", u"Quote", false, false);
    LoadMarkdownDocumentAndAssertContent(u"1. Nested blockquote", u"Quote1", false, false);
    LoadMarkdownDocumentAndAssertContent(u"2. Nested italic blockquote", u"Quote2", true, false);
    LoadMarkdownDocumentAndAssertContent(u"3. Nested bold blockquote", u"Quote3", false, true);
    LoadMarkdownDocumentAndAssertContent(u"4. Nested blockquote", u"Quote4", false, false);
    LoadMarkdownDocumentAndAssertContent(u"5. Nested blockquote", u"Quote5", false, false);
    LoadMarkdownDocumentAndAssertContent(u"6. Nested italic bold blockquote", u"Quote6", true, true);
    LoadMarkdownDocumentAndAssertContent(u"This is an indented code", u"IndentedCode", false, false);
    LoadMarkdownDocumentAndAssertContent(u"This is a fenced code with info string", u"FencedCode.C#", false, false);
    LoadMarkdownDocumentAndAssertContent(u"Item 1", u"Normal", false, false);
    LoadMarkdownDocumentAndAssertContent(u"This is a fenced code", u"FencedCode", false, false);
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, MarkdownDocumentTests)
{
    s_instance->MarkdownDocumentTests();
}

} // namespace gtest_test

void ExDocumentBuilder::InsertOnlineVideo()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], Double, Double)
    //ExFor:DocumentBuilder.InsertOnlineVideo(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert online video into a document using html code.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Visible url
    String vimeoVideoUrl = u"https://vimeo.com/52477838";

    // Embed Html code
    String vimeoEmbedCode = u"<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

    // This video will have an automatically generated thumbnail, and we are setting the size according to its 16:9 aspect ratio
    builder->Writeln(u"Video with an automatically generated thumbnail at the top left corner of the page:");
    builder->InsertOnlineVideo(vimeoVideoUrl, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 0, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 0, 320, 180, Aspose::Words::Drawing::WrapType::Square);
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    // We can get an image to use as a custom thumbnail
    {
        auto webClient = MakeObject<System::Net::WebClient>();
        ArrayPtr<uint8_t> imageBytes = webClient->DownloadData(AsposeLogoUrl);

        {
            auto stream = MakeObject<System::IO::MemoryStream>(imageBytes);
            {
                SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);
                // This puts the video where we are with our document builder, with a custom thumbnail and size depending on the size of the image
                builder->Writeln(u"Custom thumbnail at document builder's cursor:");
                builder->InsertOnlineVideo(vimeoVideoUrl, vimeoEmbedCode, imageBytes, image->get_Width(), image->get_Height());
                builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

                // We can put the video at the bottom right edge of the page too, but we'll have to take the page margins into account
                double left = builder->get_PageSetup()->get_RightMargin() - image->get_Width();
                double top = builder->get_PageSetup()->get_BottomMargin() - image->get_Height();

                // Here we use a custom thumbnail and relative positioning to put it and the bottom right of tha page
                builder->Writeln(u"Bottom right of page with custom thumbnail:");

                builder->InsertOnlineVideo(vimeoVideoUrl, vimeoEmbedCode, imageBytes, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, left, Aspose::Words::Drawing::RelativeVerticalPosition::BottomMargin, top, image->get_Width(), image->get_Height(), Aspose::Words::Drawing::WrapType::Square);
            }
        }
    }

    doc->Save(ArtifactsDir + u"DocumentBuilder.InsertOnlineVideo.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.InsertOnlineVideo.docx");
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(640, 360, Aspose::Words::Drawing::ImageType::Jpeg, shape);

    ASPOSE_ASSERT_EQ(320.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(180.0, shape->get_Height());
    ASPOSE_ASSERT_EQ(0.0, shape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, shape->get_Top());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, shape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, shape->get_RelativeVerticalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, shape->get_RelativeHorizontalPosition());

    ASSERT_EQ(u"https://vimeo.com/52477838", shape->get_HRef());

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    TestUtil::VerifyImageInShape(320, 320, Aspose::Words::Drawing::ImageType::Png, shape);
    ASPOSE_ASSERT_EQ(320.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(320.0, shape->get_Height());
    ASPOSE_ASSERT_EQ(0.0, shape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, shape->get_Top());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, shape->get_RelativeVerticalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, shape->get_RelativeHorizontalPosition());

    ASSERT_EQ(u"https://vimeo.com/52477838", shape->get_HRef());

    System::Net::ServicePointManager::set_SecurityProtocol(System::Net::SecurityProtocolType::Tls12);
    TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, shape->get_HRef());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilder, InsertOnlineVideo)
{
    s_instance->InsertOnlineVideo();
}

} // namespace gtest_test

} // namespace ApiExamples
