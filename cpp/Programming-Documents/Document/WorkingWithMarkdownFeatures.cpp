#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>


using namespace Aspose::Words;

namespace
{
    void MarkdownDocumentWithEmphases(System::String const &outputDataDir)
    {
        // ExStart:MarkdownDocumentWithEmphases
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
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

        System::String outputPath = outputDataDir + u"WorkingWithMarkdownFeatures.MarkdownDocumentWithEmphases.md";
        builder->get_Document()->Save(outputPath);
        // ExEnd:MarkdownDocumentWithEmphases
        std::cout << "Markdown Document With Emphases Produced!" << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void MarkdownDocumentWithHeadings(System::String const &outputDataDir)
    {
        //ExStart: MarkdownDocumentWithHeadings
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // By default Heading styles in Word may have bold and italic formatting.
        // If we do not want text to be emphasized, set these properties explicitly to false.
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

        System::String outputPath = outputDataDir + u"WorkingWithMarkdownFeatures.MarkdownDocumentWithHeadings.md";
        doc->Save(outputPath);
        // ExEnd:MarkdownDocumentWithHeadings
        std::cout << "Markdown Document With Headings Produced!" << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void MarkdownDocumentWithBlockQuotes(System::String const &outputDataDir)
    {
        // ExStart: MarkdownDocumentWithBlockQuotes
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"We support blockquotes in Markdown:");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));
        builder->Writeln(u"Lorem");
        builder->Writeln(u"ipsum");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
        builder->Writeln(u"The quotes can be of any level and can be nested:");
        System::SharedPtr<Style> quoteLevel3 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote2");
        builder->get_ParagraphFormat()->set_Style(quoteLevel3);
        builder->Writeln(u"Quote level 3");
        System::SharedPtr<Style> quoteLevel4 = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote3");
        builder->get_ParagraphFormat()->set_Style(quoteLevel4);
        builder->Writeln(u"Nested quote level 4");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));
        builder->Writeln();
        builder->Writeln(u"Back to first level");
        System::SharedPtr<Style> quoteLevel1WithHeading = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"Quote Heading 3");
        builder->get_ParagraphFormat()->set_Style(quoteLevel1WithHeading);
        builder->Write(u"Headings are allowed inside Quotes");

        System::String outputPath = outputDataDir + u"WorkingWithMarkdownFeatures.MarkdownDocumentWithBlockQuotes.md";
        doc->Save(outputPath);
        // ExEnd: MarkdownDocumentWithBlockQuotes
        std::cout << "Markdown Document With BlockQuotes Produced!" << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void MarkdownDocumentWithHorizontalRule(System::String const &outputDataDir)
    {
        // ExStart: MarkdownDocumentWithHorizontalRule
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(System::MakeObject<Document>());

        builder->Writeln(u"We support Horizontal rules (Thematic breaks) in Markdown:");
        builder->InsertHorizontalRule();

        System::String outputPath = outputDataDir + u"WorkingWithMarkdownFeatures.MarkdownDocumentWithHorizontalRule.md";
        builder->get_Document()->Save(outputPath);
        // ExEnd: MarkdownDocumentWithHorizontalRule
        std::cout << "Markdown Document With Horizontal Rule Produced!" << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ReadMarkdownDocument(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart: ReadMarkdownDocument
        // This is Markdown document that was produced in example of UC3.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"QuotesExample.md");

        // Let's remove Heading formatting from a Quote in the very last paragraph.
        System::SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_LastParagraph();
        paragraph->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));

        System::String outputPath = outputDataDir + u"WorkingWithMarkdownFeatures.ReadMarkdownDocument.md";
        doc->Save(outputPath);
        // ExEnd: ReadMarkdownDocument
        std::cout << "Read Markdown Document!" << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithMarkdownFeatures()
{
    std::cout << "WorkingWithMarkdownFeatures example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    MarkdownDocumentWithEmphases(outputDataDir);
    MarkdownDocumentWithHeadings(outputDataDir);
    MarkdownDocumentWithBlockQuotes(outputDataDir);
    MarkdownDocumentWithHorizontalRule(outputDataDir);
    ReadMarkdownDocument(inputDataDir, outputDataDir);
    std::cout << "WorkingWithMarkdownFeatures example finished." << std::endl << std::endl;
}