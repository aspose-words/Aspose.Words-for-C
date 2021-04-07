#pragma once

#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <system/enumerator_adapter.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithMarkdown : public DocsExamplesBase
{
public:
    void CreateMarkdownDocument()
    {
        //ExStart:CreateMarkdownDocument
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify the "Heading 1" style for the paragraph.
        builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->Writeln(u"Heading 1");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder->get_ParagraphFormat()->set_StyleName(u"Normal");

        // Insert horizontal rule.
        builder->InsertHorizontalRule();

        // Specify the ordered list.
        builder->InsertParagraph();
        builder->get_ListFormat()->ApplyNumberDefault();

        // Specify the Italic emphasis for the text.
        builder->get_Font()->set_Italic(true);
        builder->Writeln(u"Italic Text");
        builder->get_Font()->set_Italic(false);

        // Specify the Bold emphasis for the text.
        builder->get_Font()->set_Bold(true);
        builder->Writeln(u"Bold Text");
        builder->get_Font()->set_Bold(false);

        // Specify the StrikeThrough emphasis for the text.
        builder->get_Font()->set_StrikeThrough(true);
        builder->Writeln(u"StrikeThrough Text");
        builder->get_Font()->set_StrikeThrough(false);

        // Stop paragraphs numbering.
        builder->get_ListFormat()->RemoveNumbers();

        // Specify the "Quote" style for the paragraph.
        builder->get_ParagraphFormat()->set_StyleName(u"Quote");
        builder->Writeln(u"A Quote block");

        // Specify nesting Quote.
        SharedPtr<Style> nestedQuote = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote1");
        nestedQuote->set_BaseStyleName(u"Quote");
        builder->get_ParagraphFormat()->set_StyleName(u"Quote1");
        builder->Writeln(u"A nested Quote block");

        // Reset paragraph style to Normal to stop Quote blocks.
        builder->get_ParagraphFormat()->set_StyleName(u"Normal");

        // Specify a Hyperlink for the desired text.
        builder->get_Font()->set_Bold(true);
        // Note, the text of hyperlink can be emphasized.
        builder->InsertHyperlink(u"Aspose", u"https://www.aspose.com", false);
        builder->get_Font()->set_Bold(false);

        // Insert a simple table.
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Cell1");
        builder->InsertCell();
        builder->Write(u"Cell2");
        builder->EndTable();

        // Save your document as a Markdown file.
        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.CreateMarkdownDocument.md");
        //ExEnd:CreateMarkdownDocument
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

    void Headings()
    {
        //ExStart:Headings
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // By default Heading styles in Word may have bold and italic formatting.
        // If we do not want the text to be emphasized, set these properties explicitly to false.
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

        // Note that the emphases are also allowed inside Headings.
        builder->get_Font()->set_Bold(true);
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Bold Heading1");

        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.Headings.md");
        //ExEnd:Headings
    }

    void BlockQuotes()
    {
        //ExStart:BlockQuotes
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"We support blockquotes in Markdown:");

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));
        builder->Writeln(u"Lorem");
        builder->Writeln(u"ipsum");

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
        builder->Writeln(u"The quotes can be of any level and can be nested:");

        SharedPtr<Style> quoteLevel3 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote2");
        builder->get_ParagraphFormat()->set_Style(quoteLevel3);
        builder->Writeln(u"Quote level 3");

        SharedPtr<Style> quoteLevel4 = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote3");
        builder->get_ParagraphFormat()->set_Style(quoteLevel4);
        builder->Writeln(u"Nested quote level 4");

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));
        builder->Writeln();
        builder->Writeln(u"Back to first level");

        SharedPtr<Style> quoteLevel1WithHeading = doc->get_Styles()->Add(StyleType::Paragraph, u"Quote Heading 3");
        builder->get_ParagraphFormat()->set_Style(quoteLevel1WithHeading);
        builder->Write(u"Headings are allowed inside Quotes");

        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.BlockQuotes.md");
        //ExEnd:BlockQuotes
    }

    void HorizontalRule()
    {
        //ExStart:HorizontalRule
        auto builder = MakeObject<DocumentBuilder>(MakeObject<Document>());

        builder->Writeln(u"We support Horizontal rules (Thematic breaks) in Markdown:");
        builder->InsertHorizontalRule();

        builder->get_Document()->Save(ArtifactsDir + u"WorkingWithMarkdown.HorizontalRuleExample.md");
        //ExEnd:HorizontalRule
    }

    void UseWarningSource()
    {
        //ExStart:UseWarningSourceMarkdown
        auto doc = MakeObject<Document>(MyDir + u"Emphases markdown warning.docx");

        auto warnings = MakeObject<WarningInfoCollection>();
        doc->set_WarningCallback(warnings);

        doc->Save(ArtifactsDir + u"WorkingWithMarkdown.UseWarningSource.md");

        for (auto warningInfo : warnings)
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
