#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Underline.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/ParagraphCollection.h>
#include <Model/Text/ParagraphAlignment.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ListFormat.h>
#include <Model/Text/HeightRule.h>
#include <Model/Text/Font.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/RowFormat.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Styles/StyleIdentifier.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PaperSize.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Sections/Orientation.h>
#include <Model/Sections/Body.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Borders/TextureIndex.h>
#include <Model/Borders/Shading.h>
#include <Model/Borders/LineStyle.h>
#include <Model/Borders/BorderType.h>
#include <Model/Borders/BorderCollection.h>
#include <Model/Borders/Border.h>
#include <drawing/color.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{

void SetSpacebetweenAsianandLatintext(System::String dataDir)
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
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetSpacebetweenAsianandLatintext.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderSetSpacebetweenAsianandLatintext
    std::cout << "\nParagraphFormat properties AddSpaceBetweenFarEastAndAlpha and AddSpaceBetweenFarEastAndDigit set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void SetAsianTypographyLinebreakGroupProp(System::String dataDir)
{
    // ExStart:SetAsianTypographyLinebreakGroupProp
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Input.doc");
    
    System::SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat();
    format->set_FarEastLineBreakControl(false);
    format->set_WordWrap(true);
    format->set_HangingPunctuation(false);
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetAsianTypographyLinebreakGroupProp.doc");
    doc->Save(dataDir);
    // ExEnd:SetAsianTypographyLinebreakGroupProp
    std::cout << "\nParagraphFormat properties for Asian Typography line break group are set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void SetFontFormatting(System::String dataDir)
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
    font->set_Underline(Aspose::Words::Underline::Double);
    
    // Output formatted text
    builder->Writeln(u"I'm a very nice formatted string.");
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetFontFormatting.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderSetFontFormatting
    std::cout << "\nFont formatting using DocumentBuilder set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void SetParagraphFormatting(System::String dataDir)
{
    // ExStart:DocumentBuilderSetParagraphFormatting
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Set paragraph formatting properties
    System::SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
    paragraphFormat->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    paragraphFormat->set_LeftIndent(50);
    paragraphFormat->set_RightIndent(50);
    paragraphFormat->set_SpaceAfter(25);
    
    // Output text
    builder->Writeln(u"I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
    builder->Writeln(u"I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetParagraphFormatting.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderSetParagraphFormatting
    std::cout << "\nParagraph formatting using DocumentBuilder set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void SetTableCellFormatting(System::String dataDir)
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
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetTableCellFormatting.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderSetTableCellFormatting
    std::cout << "\nTable cell formatting using DocumentBuilder set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void SetTableRowFormatting(System::String dataDir)
{
    // ExStart:DocumentBuilderSetTableRowFormatting
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    System::SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    
    // Set the row formatting
    System::SharedPtr<RowFormat> rowFormat = builder->get_RowFormat();
    rowFormat->set_Height(100);
    rowFormat->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    // These formatting properties are set on the table and are applied to all rows in the table.
    table->set_LeftPadding(30);
    table->set_RightPadding(30);
    table->set_TopPadding(30);
    table->set_BottomPadding(30);
    
    builder->Writeln(u"I'm a wonderful formatted row.");
    
    builder->EndRow();
    builder->EndTable();
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetTableRowFormatting.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderSetTableRowFormatting
    std::cout << "\nTable row formatting using DocumentBuilder set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void SetMultilevelListFormatting(System::String dataDir)
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
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetMultilevelListFormatting.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderSetMultilevelListFormatting
    std::cout << "\nMultilevel list formatting using DocumentBuilder set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void SetPageSetupAndSectionFormatting(System::String dataDir)
{
    // ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Set page properties
    builder->get_PageSetup()->set_Orientation(Aspose::Words::Orientation::Landscape);
    builder->get_PageSetup()->set_LeftMargin(50);
    builder->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Paper10x14);
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.SetPageSetupAndSectionFormatting.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting
    std::cout << "\nPage setup and section formatting using DocumentBuilder set successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void ApplyParagraphStyle(System::String dataDir)
{
    // ExStart:DocumentBuilderApplyParagraphStyle
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Set paragraph style
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Title);
    
    builder->Write(u"Hello");
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.ApplyParagraphStyle.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderApplyParagraphStyle
    std::cout << "\nParagraph style using DocumentBuilder applied successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void ApplyBordersAndShadingToParagraph(System::String dataDir)
{
    // ExStart:DocumentBuilderApplyBordersAndShadingToParagraph
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Set paragraph borders
    System::SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
    borders->set_DistanceFromText(20);
    borders->idx_get(Aspose::Words::BorderType::Left)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Right)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Top)->set_LineStyle(Aspose::Words::LineStyle::Double);
    borders->idx_get(Aspose::Words::BorderType::Bottom)->set_LineStyle(Aspose::Words::LineStyle::Double);
    
    // Set paragraph shading
    System::SharedPtr<Shading> shading = builder->get_ParagraphFormat()->get_Shading();
    shading->set_Texture(Aspose::Words::TextureIndex::TextureDiagonalCross);
    shading->set_BackgroundPatternColor(System::Drawing::Color::get_LightCoral());
    shading->set_ForegroundPatternColor(System::Drawing::Color::get_LightSalmon());
    
    builder->Write(u"I'm a formatted paragraph with double border and nice shading.");
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderSetFormatting.ApplyBordersAndShadingToParagraph.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderApplyBordersAndShadingToParagraph
    std::cout << "\nBorders and shading using DocumentBuilder applied successfully to paragraph.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
}

void DocumentBuilderSetFormatting()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    SetFontFormatting(dataDir);
    SetParagraphFormatting(dataDir);
    SetTableCellFormatting(dataDir);
    SetMultilevelListFormatting(dataDir);
    SetPageSetupAndSectionFormatting(dataDir);
    ApplyParagraphStyle(dataDir);
    ApplyBordersAndShadingToParagraph(dataDir);
    SetAsianTypographyLinebreakGroupProp(dataDir);
}
