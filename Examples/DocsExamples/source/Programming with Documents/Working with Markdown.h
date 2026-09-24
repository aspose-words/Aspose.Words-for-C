#pragma once

#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
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
    void SupportedFeatures()
    {
        //ExStart:SupportedFeatures
        //GistId:1739a7dc53ee2cce1ac97f4ef9fa7310
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify the "Heading 1" style for the paragraph.
        builder->InsertParagraph();
        builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->Write(u"Heading 1");

        // Specify the Italic emphasis for the paragraph.
        builder->InsertParagraph();
        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder->get_ParagraphFormat()->set_StyleName(u"Normal");
        builder->get_Font()->set_Italic(true);
        builder->Write(u"Italic Text");
        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder->set_Italic(false);

        // Specify a Hyperlink for the desired text.
        builder->InsertParagraph();
        builder->InsertHyperlink(u"Aspose", u"https://www.aspose.com", false);
        builder->Write(u"Aspose");

        // Save your document as a Markdown file.
        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.SupportedFeatures.md");
        //ExEnd:SupportedFeatures
    }

    void BoldText()
    {
        //ExStart:BoldText
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:1739a7dc53ee2cce1ac97f4ef9fa7310
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Insert hyperlink.
        builder->InsertHyperlink(u"Aspose", u"https://www.aspose.com", false);
        //ExEnd:Link
    }

    void Image()
    {
        //ExStart:Image
        //GistId:2fdbe9a41a38579df7855f264be5edad
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Insert image.
        SharedPtr<Shape> shape = builder->InsertImage(ImagesDir + u"Logo.jpg");
        shape->get_ImageData()->set_Title(u"title");
        //ExEnd:Image
    }

    void HorizontalRule()
    {
        //ExStart:HorizontalRule
        //GistId:2fdbe9a41a38579df7855f264be5edad
        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // Insert horizontal rule.
        builder->InsertHorizontalRule();
        //ExEnd:HorizontalRule
    }

    void Heading()
    {
        //ExStart:Heading
        //GistId:2fdbe9a41a38579df7855f264be5edad
        // Use a document builder to add content to the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // By default Heading styles in Word may have Bold and Italic formatting.
        // If we do not want to be emphasized, set these properties explicitly to false.
        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_Italic(false);

        builder->Writeln(u"The following produces headings:");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Heading1");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 2"));
        builder->Writeln(u"Heading2");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 3"));
        builder->Writeln(u"Heading3");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 4"));
        builder->Writeln(u"Heading4");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 5"));
        builder->Writeln(u"Heading5");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 6"));
        builder->Writeln(u"Heading6");

        // Note, emphases are also allowed inside Headings:
        builder->get_Font()->set_Bold(true);
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Bold Heading1");

        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.Heading.md");
        //ExEnd:Heading
    }

    void SetextHeading()
    {
        //ExStart:SetextHeading
        //GistId:2fdbe9a41a38579df7855f264be5edad
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

        builder->get_Document()->Save(ArtifactsDir + u"Test.md");
    }

    void IndentedCode()
    {
        //ExStart:IndentedCode
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:2fdbe9a41a38579df7855f264be5edad
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ListFormat()->ApplyNumberDefault();

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
        //GistId:2fdbe9a41a38579df7855f264be5edad
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
        //GistId:7976171b2ed23e05b6148ea85f035327
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
        //GistId:7976171b2ed23e05b6148ea85f035327
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
