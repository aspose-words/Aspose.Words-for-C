#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BorderType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/LineStyle.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Shading.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/TextureIndex.h>
#include <drawing/color.h>
#include <system/enumerator_adapter.h>

#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/Layout/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <system/exceptions.h>
#include <iostream>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class DocumentFormatting : public DocsExamplesBase
{
public:
    //ExStart:GetParagraphLines
    //GistId:52fedeac237e3cd178ad8a516d496e8e
    void GetParagraphLines()
    {
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        auto collector = MakeObject<Layout::LayoutCollector>(doc);
        auto enumerator = MakeObject<Layout::LayoutEnumerator>(doc);
        for (const auto& paragraph : System::IterateOver<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
        {
            ProcessParagraph(paragraph, collector, enumerator);
        }
    }

    static void ProcessParagraph(SharedPtr<Paragraph> paragraph,
                                 SharedPtr<Layout::LayoutCollector> collector,
                                 SharedPtr<Layout::LayoutEnumerator> enumerator)
    {
        SharedPtr<System::Object> paragraphBreak = collector->GetEntity(paragraph);
        if (paragraphBreak == nullptr)
        {
            return;
        }

        SharedPtr<System::Object> stopEntity = GetStopEntity(paragraph, collector, enumerator);

        enumerator->set_Current(paragraphBreak);
        enumerator->MoveParent();

        int lineCount = CountLines(enumerator, stopEntity);

        String paragraphText = GetTruncatedText(paragraph->GetText());
        std::cout << String::Format(u"Paragraph '{0}' has {1} line(-s).", paragraphText, lineCount) << std::endl;
    }

    static SharedPtr<System::Object> GetStopEntity(SharedPtr<Paragraph> paragraph,
                                                   SharedPtr<Layout::LayoutCollector> collector,
                                                   SharedPtr<Layout::LayoutEnumerator> enumerator)
    {
        SharedPtr<Node> previousNode = paragraph->get_PreviousSibling();
        if (previousNode == nullptr)
        {
            return nullptr;
        }

        auto prevParagraph = System::AsCast<Paragraph>(previousNode);
        if (prevParagraph != nullptr)
        {
            enumerator->set_Current(collector->GetEntity(prevParagraph)); // Para break.
            enumerator->MoveParent();                                    // Last line.
            return enumerator->get_Current();
        }

        auto table = System::AsCast<Tables::Table>(previousNode);
        if (table != nullptr)
        {
            enumerator->set_Current(collector->GetEntity(table->get_LastRow()->get_LastCell()->get_LastParagraph())); // Cell break.
            enumerator->MoveParent();                                                                                // Cell.
            enumerator->MoveParent();                                                                                // Row.
            return enumerator->get_Current();
        }

        throw System::InvalidOperationException(u"Unsupported node type encountered.");
    }

    /// <summary>
    /// We move from line to line in a paragraph.
    /// When paragraph spans multiple pages the we will follow across them.
    /// </summary>
    static int CountLines(SharedPtr<Layout::LayoutEnumerator> enumerator, SharedPtr<System::Object> stopEntity)
    {
        int count = 1;
        while (enumerator->get_Current() != stopEntity)
        {
            if (!enumerator->MovePreviousLogical())
            {
                break;
            }
            count++;
        }
        return count;
    }

    static String GetTruncatedText(String text)
    {
        const int MaxChars = 16;
        return text.get_Length() > MaxChars ? String::Format(u"{0}...", text.Substring(0, MaxChars)) : text;
    }
    //ExEnd:GetParagraphLines

    void SpaceBetweenAsianAndLatinText()
    {
        //ExStart:SpaceBetweenAsianAndLatinText
        //GistId:b9230cb4ed425511afbb57c3223c7c7a
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
        paragraphFormat->set_AddSpaceBetweenFarEastAndAlpha(true);
        paragraphFormat->set_AddSpaceBetweenFarEastAndDigit(true);

        builder->Writeln(u"Automatically adjust space between Asian and Latin text");
        builder->Writeln(u"Automatically adjust space between Asian text and numbers");

        doc->Save(ArtifactsDir + u"DocumentFormatting.SpaceBetweenAsianAndLatinText.docx");
        //ExEnd:SpaceBetweenAsianAndLatinText
    }

    void AsianTypographyLineBreakGroup()
    {
        //ExStart:AsianTypographyLineBreakGroup
        //GistId:b9230cb4ed425511afbb57c3223c7c7a
        auto doc = MakeObject<Document>(MyDir + u"Asian typography.docx");

        SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat();
        format->set_FarEastLineBreakControl(false);
        format->set_WordWrap(true);
        format->set_HangingPunctuation(false);

        doc->Save(ArtifactsDir + u"DocumentFormatting.AsianTypographyLineBreakGroup.docx");
        //ExEnd:AsianTypographyLineBreakGroup
    }

    void ParagraphFormatting()
    {
        //ExStart:ParagraphFormatting
        //GistId:52fedeac237e3cd178ad8a516d496e8e
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
        paragraphFormat->set_Alignment(ParagraphAlignment::Center);
        paragraphFormat->set_LeftIndent(50);
        paragraphFormat->set_RightIndent(50);
        paragraphFormat->set_SpaceAfter(25);

        builder->Writeln(u"I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
        builder->Writeln(u"I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");

        doc->Save(ArtifactsDir + u"DocumentFormatting.ParagraphFormatting.docx");
        //ExEnd:ParagraphFormatting
    }

    void MultilevelListFormatting()
    {
        //ExStart:MultilevelListFormatting
        //GistId:3cea17411b8dd8066ca5ceddd7829bde
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

        doc->Save(ArtifactsDir + u"DocumentFormatting.MultilevelListFormatting.docx");
        //ExEnd:MultilevelListFormatting
    }

    void ApplyParagraphStyle()
    {
        //ExStart:ApplyParagraphStyle
        //GistId:52fedeac237e3cd178ad8a516d496e8e
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Title);
        builder->Write(u"Hello");

        doc->Save(ArtifactsDir + u"DocumentFormatting.ApplyParagraphStyle.docx");
        //ExEnd:ApplyParagraphStyle
    }

    void ApplyBordersAndShadingToParagraph()
    {
        //ExStart:ApplyBordersAndShadingToParagraph
        //GistId:52fedeac237e3cd178ad8a516d496e8e
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
        borders->set_DistanceFromText(20);
        borders->idx_get(BorderType::Left)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Right)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Top)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Bottom)->set_LineStyle(LineStyle::Double);

        SharedPtr<Shading> shading = builder->get_ParagraphFormat()->get_Shading();
        shading->set_Texture(TextureIndex::TextureDiagonalCross);
        shading->set_BackgroundPatternColor(System::Drawing::Color::get_LightCoral());
        shading->set_ForegroundPatternColor(System::Drawing::Color::get_LightSalmon());

        builder->Write(u"I'm a formatted paragraph with double border and nice shading.");

        doc->Save(ArtifactsDir + u"DocumentFormatting.ApplyBordersAndShadingToParagraph.doc");
        //ExEnd:ApplyBordersAndShadingToParagraph
    }

    void ChangeAsianParagraphSpacingAndIndents()
    {
        //ExStart:ChangeAsianParagraphSpacingAndIndents
        auto doc = MakeObject<Document>(MyDir + u"Asian typography.docx");

        SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();
        format->set_CharacterUnitLeftIndent(10);
        // ParagraphFormat.LeftIndent will be updated
        format->set_CharacterUnitRightIndent(10);
        // ParagraphFormat.RightIndent will be updated
        format->set_CharacterUnitFirstLineIndent(20);
        // ParagraphFormat.FirstLineIndent will be updated
        format->set_LineUnitBefore(5);
        // ParagraphFormat.SpaceBefore will be updated
        format->set_LineUnitAfter(10);
        // ParagraphFormat.SpaceAfter will be updated

        doc->Save(ArtifactsDir + u"DocumentFormatting.ChangeAsianParagraphSpacingAndIndents.doc");
        //ExEnd:ChangeAsianParagraphSpacingAndIndents
    }

    void SnapToGrid()
    {
        //ExStart:SnapToGrid
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Optimize the layout when typing in Asian characters.
        SharedPtr<Paragraph> par = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        par->get_ParagraphFormat()->set_SnapToGrid(true);

        builder->Writeln(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod ") +
                         u"tempor incididunt ut labore et dolore magna aliqua.");

        par->get_Runs()->idx_get(0)->get_Font()->set_SnapToGrid(true);

        doc->Save(ArtifactsDir + u"Paragraph.SnapToGrid.docx");
        //ExEnd:SnapToGrid
    }

    void GetParagraphStyleSeparator()
    {
        //ExStart:GetParagraphStyleSeparator
        //GistId:52fedeac237e3cd178ad8a516d496e8e
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        for (const auto& paragraph : System::IterateOver<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
        {
            if (paragraph->get_BreakIsStyleSeparator())
            {
                std::cout << "Separator Found!" << std::endl;
            }
        }
        //ExEnd:GetParagraphStyleSeparator
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
