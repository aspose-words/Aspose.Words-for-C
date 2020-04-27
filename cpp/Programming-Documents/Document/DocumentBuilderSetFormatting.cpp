#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PaperSize.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/TextureIndex.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void SetSpacebetweenAsianandLatintext(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderSetSpacebetweenAsianandLatintext
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Set paragraph formatting properties
        System::SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
        paragraphFormat->set_AddSpaceBetweenFarEastAndAlpha(true);
        paragraphFormat->set_AddSpaceBetweenFarEastAndDigit(true);

        builder->Writeln(u"Automatically adjust space between Asian and Latin text");
        builder->Writeln(u"Automatically adjust space between Asian text and numbers");

        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetSpacebetweenAsianandLatintext.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderSetSpacebetweenAsianandLatintext
        std::cout << "ParagraphFormat properties AddSpaceBetweenFarEastAndAlpha and AddSpaceBetweenFarEastAndDigit set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetAsianTypographyLinebreakGroupProp(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetAsianTypographyLinebreakGroupProp
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");

        System::SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat();
        format->set_FarEastLineBreakControl(false);
        format->set_WordWrap(true);
        format->set_HangingPunctuation(false);

        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetAsianTypographyLinebreakGroupProp.docx";
        doc->Save(outputPath);
        // ExEnd:SetAsianTypographyLinebreakGroupProp
        std::cout << "ParagraphFormat properties for Asian Typography line break group are set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }


    void SetFontFormatting(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderSetFontFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Set font formatting properties
        System::SharedPtr<Font> font = builder->get_Font();
        font->set_Bold(true);
        font->set_Color(System::Drawing::Color::get_DarkBlue());
        font->set_Italic(true);
        font->set_Name(u"Arial");
        font->set_Size(24);
        font->set_Spacing(5);
        font->set_Underline(Underline::Double);

        // Output formatted text
        builder->Writeln(u"I'm a very nice formatted string.");
        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetFontFormatting.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderSetFontFormatting
        std::cout << "Font formatting using DocumentBuilder set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetParagraphFormatting(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderSetParagraphFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Set paragraph formatting properties
        System::SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
        paragraphFormat->set_Alignment(ParagraphAlignment::Center);
        paragraphFormat->set_LeftIndent(50);
        paragraphFormat->set_RightIndent(50);
        paragraphFormat->set_SpaceAfter(25);

        // Output text
        builder->Writeln(u"I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
        builder->Writeln(u"I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");

        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetParagraphFormatting.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderSetParagraphFormatting
        std::cout << "Paragraph formatting using DocumentBuilder set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetTableCellFormatting(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderSetTableCellFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->InsertCell();

        // Set the cell formatting
        System::SharedPtr<CellFormat> cellFormat = builder->get_CellFormat();
        cellFormat->set_Width(250);
        cellFormat->set_LeftPadding(30);
        cellFormat->set_RightPadding(30);
        cellFormat->set_TopPadding(30);
        cellFormat->set_BottomPadding(30);

        builder->Writeln(u"I'm a wonderful formatted cell.");

        builder->EndRow();
        builder->EndTable();

        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetTableCellFormatting.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderSetTableCellFormatting
        std::cout << "Table cell formatting using DocumentBuilder set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetTableRowFormatting(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderSetTableRowFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();

        // Set the row formatting
        System::SharedPtr<RowFormat> rowFormat = builder->get_RowFormat();
        rowFormat->set_Height(100);
        rowFormat->set_HeightRule(HeightRule::Exactly);
        // These formatting properties are set on the table and are applied to all rows in the table.
        table->set_LeftPadding(30);
        table->set_RightPadding(30);
        table->set_TopPadding(30);
        table->set_BottomPadding(30);

        builder->Writeln(u"I'm a wonderful formatted row.");

        builder->EndRow();
        builder->EndTable();

        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetTableRowFormatting.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderSetTableRowFormatting
        std::cout << "Table row formatting using DocumentBuilder set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetMultilevelListFormatting(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderSetMultilevelListFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

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
        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetMultilevelListFormatting.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderSetMultilevelListFormatting
        std::cout << "Multilevel list formatting using DocumentBuilder set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetPageSetupAndSectionFormatting(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Set page properties
        builder->get_PageSetup()->set_Orientation(Orientation::Landscape);
        builder->get_PageSetup()->set_LeftMargin(50);
        builder->get_PageSetup()->set_PaperSize(PaperSize::Paper10x14);

        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.SetPageSetupAndSectionFormatting.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting
        std::cout << "Page setup and section formatting using DocumentBuilder set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ApplyParagraphStyle(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderApplyParagraphStyle
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Set paragraph style
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Title);

        builder->Write(u"Hello");
        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.ApplyParagraphStyle.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderApplyParagraphStyle
        std::cout << "Paragraph style using DocumentBuilder applied successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ApplyBordersAndShadingToParagraph(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderApplyBordersAndShadingToParagraph
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Set paragraph borders
        System::SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
        borders->set_DistanceFromText(20);
        borders->idx_get(BorderType::Left)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Right)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Top)->set_LineStyle(LineStyle::Double);
        borders->idx_get(BorderType::Bottom)->set_LineStyle(LineStyle::Double);

        // Set paragraph shading
        System::SharedPtr<Shading> shading = builder->get_ParagraphFormat()->get_Shading();
        shading->set_Texture(TextureIndex::TextureDiagonalCross);
        shading->set_BackgroundPatternColor(System::Drawing::Color::get_LightCoral());
        shading->set_ForegroundPatternColor(System::Drawing::Color::get_LightSalmon());

        builder->Write(u"I'm a formatted paragraph with double border and nice shading.");
        System::String outputPath = outputDataDir + u"DocumentBuilderSetFormatting.ApplyBordersAndShadingToParagraph.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderApplyBordersAndShadingToParagraph
        std::cout << "Borders and shading using DocumentBuilder applied successfully to paragraph." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ChangeAsianParagraphSpacingAndIndents(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ChangeAsianParagraphSpacingAndIndents
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");

        System::SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();
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

        const System::String outputPath = outputDataDir + u"ChangeAsianParagraphSpacingAndIndents_out.doc";
        doc->Save(outputPath);
        // ExEnd:ChangeAsianParagraphSpacingAndIndents
        std::cout << "Spacing and Indents applied successfully to paragraph." << std::endl << "File saved at " <<  outputPath.ToUtf8String();
    }

    void SetSnapToGrid(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetSnapToGrid
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");

        System::SharedPtr<Paragraph> par = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        par->get_ParagraphFormat()->set_SnapToGrid(true);
        par->get_Runs()->idx_get(0)->get_Font()->set_SnapToGrid(true);

        const System::String outputPath = outputDataDir + u"SetSnapToGrid_out.doc";
        doc->Save(outputPath);
        // ExEnd:SetSnapToGrid
        std::cout << "SetSnapToGrid successfully to paragraph." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void DocumentBuilderSetFormatting()
{
    std::cout << "DocumentBuilderSetFormatting example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    SetSpacebetweenAsianandLatintext(outputDataDir);
    SetFontFormatting(outputDataDir);
    SetParagraphFormatting(outputDataDir);
    SetTableCellFormatting(outputDataDir);
    SetTableRowFormatting(outputDataDir);
    SetMultilevelListFormatting(outputDataDir);
    SetPageSetupAndSectionFormatting(outputDataDir);
    ApplyParagraphStyle(outputDataDir);
    ApplyBordersAndShadingToParagraph(outputDataDir);
    SetAsianTypographyLinebreakGroupProp(inputDataDir, outputDataDir);
    ChangeAsianParagraphSpacingAndIndents(inputDataDir, outputDataDir);
    SetSnapToGrid(inputDataDir, outputDataDir);
    std::cout << "DocumentBuilderSetFormatting example finished." << std::endl << std::endl;
}
