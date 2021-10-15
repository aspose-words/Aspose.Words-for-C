#pragma once

#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/WarningSource.h>
#include <system/enumerator_adapter.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithMarkdown : public DocsExamplesBase
{
public:
    void BoldText()
    {
        //ExStart:BoldText
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Make the text Bold.
        builder->get_Font()->set_Bold(true);
        builder->Writeln(u"This text will be Bold");
        //ExEnd:BoldText
    }

    void ItalicText()
    {
        //ExStart:ItalicText
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Make the text Italic.
        builder->get_Font()->set_Italic(true);
        builder->Writeln(u"This text will be Italic");
        //ExEnd:ItalicText
    }

    void Strikethrough()
    {
        //ExStart:Strikethrough
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Make the text Strikethrough.
        builder->get_Font()->set_StrikeThrough(true);
        builder->Writeln(u"This text will be StrikeThrough");
        //ExEnd:Strikethrough
    }

    void InlineCode()
    {
        //ExStart:InlineCode
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Number of backticks is missed, one backtick will be used by default.
        SharedPtr<Style> inlineCode1BackTicks = builder->get_Document()->get_Styles()->Add(StyleType::Character, u"InlineCode");
        builder->get_Font()->set_Style(inlineCode1BackTicks);
        builder->Writeln(u"Text with InlineCode style with 1 backtick");

        // There will be 3 backticks.
        SharedPtr<Style> inlineCode3BackTicks = builder->get_Document()->get_Styles()->Add(StyleType::Character, u"InlineCode.3");
        builder->get_Font()->set_Style(inlineCode3BackTicks);
        builder->Writeln(u"Text with InlineCode style with 3 backtick");
        //ExEnd:InlineCode
    }

    void Autolink()
    {
        //ExStart:Autolink
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Insert hyperlink.
        builder->InsertHyperlink(u"https://www.aspose.com", u"https://www.aspose.com", false);
        builder->InsertHyperlink(u"email@aspose.com", u"mailto:email@aspose.com", false);
        //ExEnd:Autolink
    }

    void Link()
    {
        //ExStart:Link
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Insert hyperlink.
        builder->InsertHyperlink(u"Aspose", u"https://www.aspose.com", false);
        //ExEnd:Link
    }

    void Image()
    {
        //ExStart:Image
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Insert image.
        auto shape = MakeObject<Shape>(builder->get_Document(), ShapeType::Image);
        shape->set_WrapType(WrapType::Inline);
        shape->get_ImageData()->set_SourceFullName(u"/attachment/1456/pic001.png");
        shape->get_ImageData()->set_Title(u"title");
        builder->InsertNode(shape);
        //ExEnd:Image
    }

    void HorizontalRule()
    {
        //ExStart:HorizontalRule
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Insert horizontal rule.
        builder->InsertHorizontalRule();
        //ExEnd:HorizontalRule
    }

    void Heading()
    {
        //ExStart:Heading
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // By default Heading styles in Word may have Bold and Italic formatting.
        // If we do not want to be emphasized, set these properties explicitly to false.
        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->Writeln(u"This is an H1 tag");
        //ExEnd:Heading
    }

    void SetextHeading()
    {
        //ExStart:SetextHeading
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->Writeln(u"This is an H1 tag");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        SharedPtr<Style> setexHeading1 = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"SetextHeading1");
        builder->get_ParagraphFormat()->set_Style(setexHeading1);
        builder->get_Document()->get_Styles()->idx_get(u"SetextHeading1")->set_BaseStyleName(u"Heading 1");
        builder->Writeln(u"Setext Heading level 1");

        builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 3"));
        builder->Writeln(u"This is an H3 tag");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        SharedPtr<Style> setexHeading2 = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"SetextHeading2");
        builder->get_ParagraphFormat()->set_Style(setexHeading2);
        builder->get_Document()->get_Styles()->idx_get(u"SetextHeading2")->set_BaseStyleName(u"Heading 3");

        // Setex heading level will be reset to 2 if the base paragraph has a Heading level greater than 2.
        builder->Writeln(u"Setext Heading level 2");
        //ExEnd:SetextHeading
    }

    void IndentedCode()
    {
        //ExStart:IndentedCode
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        SharedPtr<Style> indentedCode = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"IndentedCode");
        builder->get_ParagraphFormat()->set_Style(indentedCode);
        builder->Writeln(u"This is an indented code");
        //ExEnd:IndentedCode
    }

    void FencedCode()
    {
        //ExStart:FencedCode
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        SharedPtr<Style> fencedCode = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"FencedCode");
        builder->get_ParagraphFormat()->set_Style(fencedCode);
        builder->Writeln(u"This is an fenced code");

        SharedPtr<Style> fencedCodeWithInfo = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"FencedCode.C#");
        builder->get_ParagraphFormat()->set_Style(fencedCodeWithInfo);
        builder->Writeln(u"This is a fenced code with info string");
        //ExEnd:FencedCode
    }

    void Quote()
    {
        //ExStart:Quote
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // By default a document stores blockquote style for the first level.
        builder->get_ParagraphFormat()->set_StyleName(u"Quote");
        builder->Writeln(u"Blockquote");

        // Create styles for nested levels through style inheritance.
        SharedPtr<Style> quoteLevel2 = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"Quote1");
        builder->get_ParagraphFormat()->set_Style(quoteLevel2);
        builder->get_Document()->get_Styles()->idx_get(u"Quote1")->set_BaseStyleName(u"Quote");
        builder->Writeln(u"1. Nested blockquote");
        //ExEnd:Quote
    }

    void BulletedList()
    {
        //ExStart:BulletedList
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        builder->get_ListFormat()->ApplyBulletDefault();
        builder->get_ListFormat()->get_List()->get_ListLevels()->idx_get(0)->set_NumberFormat(u"-");

        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");

        builder->get_ListFormat()->ListIndent();

        builder->Writeln(u"Item 2a");
        builder->Writeln(u"Item 2b");
        //ExEnd:BulletedList
    }

    void OrderedList()
    {
        //ExStart:OrderedList
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ListFormat()->ApplyBulletDefault();
        builder->get_ListFormat()->get_List()->get_ListLevels()->idx_get(0)->set_NumberFormat(String::Format(u"{0}.", (char16_t)0));
        builder->get_ListFormat()->get_List()->get_ListLevels()->idx_get(1)->set_NumberFormat(String::Format(u"{0}.", (char16_t)1));

        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");

        builder->get_ListFormat()->ListIndent();

        builder->Writeln(u"Item 2a");
        builder->Writeln(u"Item 2b");
        //ExEnd:OrderedList
    }

    void Table()
    {
        //ExStart:Table
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Add the first row.
        builder->InsertCell();
        builder->Writeln(u"a");
        builder->InsertCell();
        builder->Writeln(u"b");

        // Add the second row.
        builder->InsertCell();
        builder->Writeln(u"c");
        builder->InsertCell();
        builder->Writeln(u"d");
        //ExEnd:Table
    }

    void ReadMarkdownDocument()
    {
        //ExStart:ReadMarkdownDocument
        auto doc = MakeObject<Document>(MyDir + u"Quotes.md");

        // Let's remove Heading formatting from a Quote in the very last paragraph.
        SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_LastParagraph();
        paragraph->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));

        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.ReadMarkdownDocument.md");
        //ExEnd:ReadMarkdownDocument
    }

    void Emphases()
    {
        //ExStart:Emphases
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Markdown treats asterisks (*) and underscores (_) as indicators of emphasis.");
        builder->Write(u"You can write ");

        builder->get_Font()->set_Bold(true);
        builder->Write(u"bold");

        builder->get_Font()->set_Bold(false);
        builder->Write(u" or ");

        builder->get_Font()->set_Italic(true);
        builder->Write(u"italic");

        builder->get_Font()->set_Italic(false);
        builder->Writeln(u" text. ");

        builder->Write(u"You can also write ");
        builder->get_Font()->set_Bold(true);

        builder->get_Font()->set_Italic(true);
        builder->Write(u"BoldItalic");

        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);
        builder->Write(u"text.");

        builder->get_Document()->Save(ArtifactsDir + u"WorkingWithMarkdown.Emphases.md");
        //ExEnd:Emphases
    }

    void UseWarningSource()
    {
        //ExStart:UseWarningSourceMarkdown
        auto doc = MakeObject<Document>(MyDir + u"Emphases markdown warning.docx");

        auto warnings = MakeObject<WarningInfoCollection>();
        doc->set_WarningCallback(warnings);

        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.UseWarningSource.md");

        for (const auto& warningInfo : warnings)
        {
            if (warningInfo->get_Source() == WarningSource::Markdown)
            {
                std::cout << warningInfo->get_Description() << std::endl;
            }
        }
        //ExEnd:UseWarningSourceMarkdown
    }
};

}} // namespace DocsExamples::Programming_with_Documents
