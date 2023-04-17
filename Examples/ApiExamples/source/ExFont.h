#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Comment.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/EmphasisMark.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldSeparator.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Fonts/EmbeddedFontFormat.h>
#include <Aspose.Words.Cpp/Fonts/EmbeddedFontStyle.h>
#include <Aspose.Words.Cpp/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Fonts/FontFamily.h>
#include <Aspose.Words.Cpp/Fonts/FontInfo.h>
#include <Aspose.Words.Cpp/Fonts/FontInfoCollection.h>
#include <Aspose.Words.Cpp/Fonts/FontPitch.h>
#include <Aspose.Words.Cpp/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Fonts/PhysicalFontInfo.h>
#include <Aspose.Words.Cpp/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Shading.h>
#include <Aspose.Words.Cpp/SpecialChar.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/TextDmlEffect.h>
#include <Aspose.Words.Cpp/TextEffect.h>
#include <Aspose.Words.Cpp/TextureIndex.h>
#include <Aspose.Words.Cpp/Themes/Theme.h>
#include <Aspose.Words.Cpp/Themes/ThemeColor.h>
#include <Aspose.Words.Cpp/Themes/ThemeFont.h>
#include <Aspose.Words.Cpp/Themes/ThemeFonts.h>
#include <Aspose.Words.Cpp/Underline.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <drawing/color.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ilist.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/io/search_option.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Themes;

namespace ApiExamples {

class ExFont : public ApiExampleBase
{
public:
    void CreateFormattedRun()
    {
        //ExStart
        //ExFor:Document.#ctor()
        //ExFor:Font
        //ExFor:Font.Name
        //ExFor:Font.Size
        //ExFor:Font.HighlightColor
        //ExFor:Run
        //ExFor:Run.#ctor(DocumentBase,String)
        //ExFor:Story.FirstParagraph
        //ExSummary:Shows how to format a run of text using its font property.
        auto doc = MakeObject<Document>();
        auto run = MakeObject<Run>(doc, u"Hello world!");

        SharedPtr<Aspose::Words::Font> font = run->get_Font();
        font->set_Name(u"Courier New");
        font->set_Size(36);
        font->set_HighlightColor(System::Drawing::Color::get_Yellow());

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(run);
        doc->Save(ArtifactsDir + u"Font.CreateFormattedRun.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.CreateFormattedRun.docx");
        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Hello world!", run->GetText().Trim());
        ASSERT_EQ(u"Courier New", run->get_Font()->get_Name());
        ASPOSE_ASSERT_EQ(36, run->get_Font()->get_Size());
        ASSERT_EQ(System::Drawing::Color::get_Yellow().ToArgb(), run->get_Font()->get_HighlightColor().ToArgb());
    }

    void Caps()
    {
        //ExStart
        //ExFor:Font.AllCaps
        //ExFor:Font.SmallCaps
        //ExSummary:Shows how to format a run to display its contents in capitals.
        auto doc = MakeObject<Document>();
        auto para = System::ExplicitCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        // There are two ways of getting a run to display its lowercase text in uppercase without changing the contents.
        // 1 -  Set the AllCaps flag to display all characters in regular capitals:
        auto run = MakeObject<Run>(doc, u"all capitals");
        run->get_Font()->set_AllCaps(true);
        para->AppendChild(run);

        para = System::ExplicitCast<Paragraph>(para->get_ParentNode()->AppendChild(MakeObject<Paragraph>(doc)));

        // 2 -  Set the SmallCaps flag to display all characters in small capitals:
        // If a character is lower case, it will appear in its upper case form
        // but will have the same height as the lower case (the font's x-height).
        // Characters that were in upper case originally will look the same.
        run = MakeObject<Run>(doc, u"Small Capitals");
        run->get_Font()->set_SmallCaps(true);
        para->AppendChild(run);

        doc->Save(ArtifactsDir + u"Font.Caps.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Caps.docx");
        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"all capitals", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_AllCaps());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Small Capitals", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_SmallCaps());
    }

    void GetDocumentFonts()
    {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfo
        //ExFor:FontInfo.Name
        //ExFor:FontInfo.IsTrueType
        //ExSummary:Shows how to print the details of what fonts are present in a document.
        auto doc = MakeObject<Document>(MyDir + u"Embedded font.docx");

        SharedPtr<Aspose::Words::Fonts::FontInfoCollection> allFonts = doc->get_FontInfos();
        ASSERT_EQ(5, allFonts->get_Count());
        //ExSkip

        // Print all the used and unused fonts in the document.
        for (int i = 0; i < allFonts->get_Count(); i++)
        {
            std::cout << "Font index #" << i << std::endl;
            std::cout << "\tName: " << allFonts->idx_get(i)->get_Name() << std::endl;
            std::cout << "\tIs " << (allFonts->idx_get(i)->get_IsTrueType() ? String(u"") : String(u"not ")) << "a trueType font" << std::endl;
        }
        //ExEnd
    }

    void DefaultValuesEmbeddedFontsParameters()
    {
        auto doc = MakeObject<Document>();

        ASSERT_FALSE(doc->get_FontInfos()->get_EmbedTrueTypeFonts());
        ASSERT_FALSE(doc->get_FontInfos()->get_EmbedSystemFonts());
        ASSERT_FALSE(doc->get_FontInfos()->get_SaveSubsetFonts());
    }

    void FontInfoCollection(bool embedAllFonts)
    {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfoCollection.EmbedTrueTypeFonts
        //ExFor:FontInfoCollection.EmbedSystemFonts
        //ExFor:FontInfoCollection.SaveSubsetFonts
        //ExSummary:Shows how to save a document with embedded TrueType fonts.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Aspose::Words::Fonts::FontInfoCollection> fontInfos = doc->get_FontInfos();
        fontInfos->set_EmbedTrueTypeFonts(embedAllFonts);
        fontInfos->set_EmbedSystemFonts(embedAllFonts);
        fontInfos->set_SaveSubsetFonts(embedAllFonts);

        doc->Save(ArtifactsDir + u"Font.FontInfoCollection.docx");

        if (embedAllFonts)
        {
            ASSERT_LT(25000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"Font.FontInfoCollection.docx")->get_Length());
        }
        else
        {
            ASSERT_GE(15000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"Font.FontInfoCollection.docx")->get_Length());
        }
        //ExEnd
    }

    void WorkWithEmbeddedFonts(bool embedTrueTypeFonts, bool embedSystemFonts, bool saveSubsetFonts)
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Aspose::Words::Fonts::FontInfoCollection> fontInfos = doc->get_FontInfos();
        fontInfos->set_EmbedTrueTypeFonts(embedTrueTypeFonts);
        fontInfos->set_EmbedSystemFonts(embedSystemFonts);
        fontInfos->set_SaveSubsetFonts(saveSubsetFonts);

        doc->Save(ArtifactsDir + u"Font.WorkWithEmbeddedFonts.docx");
    }

    void StrikeThrough()
    {
        //ExStart
        //ExFor:Font.StrikeThrough
        //ExFor:Font.DoubleStrikeThrough
        //ExSummary:Shows how to add a line strikethrough to text.
        auto doc = MakeObject<Document>();
        auto para = System::ExplicitCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        auto run = MakeObject<Run>(doc, u"Text with a single-line strikethrough.");
        run->get_Font()->set_StrikeThrough(true);
        para->AppendChild(run);

        para = System::ExplicitCast<Paragraph>(para->get_ParentNode()->AppendChild(MakeObject<Paragraph>(doc)));

        run = MakeObject<Run>(doc, u"Text with a double-line strikethrough.");
        run->get_Font()->set_DoubleStrikeThrough(true);
        para->AppendChild(run);

        doc->Save(ArtifactsDir + u"Font.StrikeThrough.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.StrikeThrough.docx");

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Text with a single-line strikethrough.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_StrikeThrough());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Text with a double-line strikethrough.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_DoubleStrikeThrough());
    }

    void PositionSubscript()
    {
        //ExStart
        //ExFor:Font.Position
        //ExFor:Font.Subscript
        //ExFor:Font.Superscript
        //ExSummary:Shows how to format text to offset its position.
        auto doc = MakeObject<Document>();
        auto para = System::ExplicitCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        // Raise this run of text 5 points above the baseline.
        auto run = MakeObject<Run>(doc, u"Raised text. ");
        run->get_Font()->set_Position(5);
        para->AppendChild(run);

        // Lower this run of text 10 points below the baseline.
        run = MakeObject<Run>(doc, u"Lowered text. ");
        run->get_Font()->set_Position(-10);
        para->AppendChild(run);

        // Add a run of normal text.
        run = MakeObject<Run>(doc, u"Text in its default position. ");
        para->AppendChild(run);

        // Add a run of text that appears as subscript.
        run = MakeObject<Run>(doc, u"Subscript. ");
        run->get_Font()->set_Subscript(true);
        para->AppendChild(run);

        // Add a run of text that appears as superscript.
        run = MakeObject<Run>(doc, u"Superscript.");
        run->get_Font()->set_Superscript(true);
        para->AppendChild(run);

        doc->Save(ArtifactsDir + u"Font.PositionSubscript.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.PositionSubscript.docx");
        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Raised text.", run->GetText().Trim());
        ASPOSE_ASSERT_EQ(5, run->get_Font()->get_Position());

        doc = MakeObject<Document>(ArtifactsDir + u"Font.PositionSubscript.docx");
        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(1);

        ASSERT_EQ(u"Lowered text.", run->GetText().Trim());
        ASPOSE_ASSERT_EQ(-10, run->get_Font()->get_Position());

        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(3);

        ASSERT_EQ(u"Subscript.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Subscript());

        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(4);

        ASSERT_EQ(u"Superscript.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Superscript());
    }

    void ScalingSpacing()
    {
        //ExStart
        //ExFor:Font.Scaling
        //ExFor:Font.Spacing
        //ExSummary:Shows how to set horizontal scaling and spacing for characters.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add run of text and increase character width to 150%.
        builder->get_Font()->set_Scaling(150);
        builder->Writeln(u"Wide characters");

        // Add run of text and add 1pt of extra horizontal spacing between each character.
        builder->get_Font()->set_Spacing(1);
        builder->Writeln(u"Expanded by 1pt");

        // Add run of text and bring characters closer together by 1pt.
        builder->get_Font()->set_Spacing(-1);
        builder->Writeln(u"Condensed by 1pt");

        doc->Save(ArtifactsDir + u"Font.ScalingSpacing.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.ScalingSpacing.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Wide characters", run->GetText().Trim());
        ASSERT_EQ(150, run->get_Font()->get_Scaling());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Expanded by 1pt", run->GetText().Trim());
        ASPOSE_ASSERT_EQ(1, run->get_Font()->get_Spacing());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(2)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Condensed by 1pt", run->GetText().Trim());
        ASPOSE_ASSERT_EQ(-1, run->get_Font()->get_Spacing());
    }

    void Italic()
    {
        //ExStart
        //ExFor:Font.Italic
        //ExSummary:Shows how to write italicized text using a document builder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Size(36);
        builder->get_Font()->set_Italic(true);
        builder->Writeln(u"Hello world!");

        doc->Save(ArtifactsDir + u"Font.Italic.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Italic.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Hello world!", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Italic());
    }

    void EngraveEmboss()
    {
        //ExStart
        //ExFor:Font.Emboss
        //ExFor:Font.Engrave
        //ExSummary:Shows how to apply engraving/embossing effects to text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Size(36);
        builder->get_Font()->set_Color(System::Drawing::Color::get_LightBlue());

        // Below are two ways of using shadows to apply a 3D-like effect to the text.
        // 1 -  Engrave text to make it look like the letters are sunken into the page:
        builder->get_Font()->set_Engrave(true);

        builder->Writeln(u"This text is engraved.");

        // 2 -  Emboss text to make it look like the letters pop out of the page:
        builder->get_Font()->set_Engrave(false);
        builder->get_Font()->set_Emboss(true);

        builder->Writeln(u"This text is embossed.");

        doc->Save(ArtifactsDir + u"Font.EngraveEmboss.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.EngraveEmboss.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"This text is engraved.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Engrave());
        ASSERT_FALSE(run->get_Font()->get_Emboss());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"This text is embossed.", run->GetText().Trim());
        ASSERT_FALSE(run->get_Font()->get_Engrave());
        ASSERT_TRUE(run->get_Font()->get_Emboss());
    }

    void Shadow()
    {
        //ExStart
        //ExFor:Font.Shadow
        //ExSummary:Shows how to create a run of text formatted with a shadow.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set the Shadow flag to apply an offset shadow effect,
        // making it look like the letters are floating above the page.
        builder->get_Font()->set_Shadow(true);
        builder->get_Font()->set_Size(36);

        builder->Writeln(u"This text has a shadow.");

        doc->Save(ArtifactsDir + u"Font.Shadow.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Shadow.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"This text has a shadow.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Shadow());
    }

    void Outline()
    {
        //ExStart
        //ExFor:Font.Outline
        //ExSummary:Shows how to create a run of text formatted as outline.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set the Outline flag to change the text's fill color to white and
        // leave a thin outline around each character in the original color of the text.
        builder->get_Font()->set_Outline(true);
        builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        builder->get_Font()->set_Size(36);

        builder->Writeln(u"This text has an outline.");

        doc->Save(ArtifactsDir + u"Font.Outline.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Outline.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"This text has an outline.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Outline());
    }

    void Hidden()
    {
        //ExStart
        //ExFor:Font.Hidden
        //ExSummary:Shows how to create a run of hidden text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // With the Hidden flag set to true, any text that we create using this Font object will be invisible in the document.
        // We will not see or highlight hidden text unless we enable the "Hidden text" option
        // found in Microsoft Word via "File" -> "Options" -> "Display". The text will still be there,
        // and we will be able to access this text programmatically.
        // It is not advised to use this method to hide sensitive information.
        builder->get_Font()->set_Hidden(true);
        builder->get_Font()->set_Size(36);

        builder->Writeln(u"This text will not be visible in the document.");

        doc->Save(ArtifactsDir + u"Font.Hidden.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Hidden.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"This text will not be visible in the document.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Hidden());
    }

    void Kerning()
    {
        //ExStart
        //ExFor:Font.Kerning
        //ExSummary:Shows how to specify the font size at which kerning begins to take effect.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Name(u"Arial Black");

        // Set the builder's font size, and minimum size at which kerning will take effect.
        // The font size falls below the kerning threshold, so the run bellow will not have kerning.
        builder->get_Font()->set_Size(18);
        builder->get_Font()->set_Kerning(24);

        builder->Writeln(u"TALLY. (Kerning not applied)");

        // Set the kerning threshold so that the builder's current font size is above it.
        // Any text we add from this point will have kerning applied. The spaces between characters
        // will be adjusted, normally resulting in a slightly more aesthetically pleasing text run.
        builder->get_Font()->set_Kerning(12);

        builder->Writeln(u"TALLY. (Kerning applied)");

        doc->Save(ArtifactsDir + u"Font.Kerning.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Kerning.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"TALLY. (Kerning not applied)", run->GetText().Trim());
        ASPOSE_ASSERT_EQ(24, run->get_Font()->get_Kerning());
        ASPOSE_ASSERT_EQ(18, run->get_Font()->get_Size());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"TALLY. (Kerning applied)", run->GetText().Trim());
        ASPOSE_ASSERT_EQ(12, run->get_Font()->get_Kerning());
        ASPOSE_ASSERT_EQ(18, run->get_Font()->get_Size());
    }

    void NoProofing()
    {
        //ExStart
        //ExFor:Font.NoProofing
        //ExSummary:Shows how to prevent text from being spell checked by Microsoft Word.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Normally, Microsoft Word emphasizes spelling errors with a jagged red underline.
        // We can un-set the "NoProofing" flag to create a portion of text that
        // bypasses the spell checker while completely disabling it.
        builder->get_Font()->set_NoProofing(true);

        builder->Writeln(u"Proofing has been disabled, so these spelking errrs will not display red lines underneath.");

        doc->Save(ArtifactsDir + u"Font.NoProofing.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.NoProofing.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Proofing has been disabled, so these spelking errrs will not display red lines underneath.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_NoProofing());
    }

    void LocaleId()
    {
        //ExStart
        //ExFor:Font.LocaleId
        //ExSummary:Shows how to set the locale of the text that we are adding with a document builder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If we set the font's locale to English and insert some Russian text,
        // the English locale spell checker will not recognize the text and detect it as a spelling error.
        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"en-US", false)->get_LCID());
        builder->Writeln(u"Привет!");

        // Set a matching locale for the text that we are about to add to apply the appropriate spell checker.
        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"ru-RU", false)->get_LCID());
        builder->Writeln(u"Привет!");

        doc->Save(ArtifactsDir + u"Font.LocaleId.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.LocaleId.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Привет!", run->GetText().Trim());
        ASSERT_EQ(1033, run->get_Font()->get_LocaleId());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Привет!", run->GetText().Trim());
        ASSERT_EQ(1049, run->get_Font()->get_LocaleId());
    }

    void Underlines()
    {
        //ExStart
        //ExFor:Font.Underline
        //ExFor:Font.UnderlineColor
        //ExSummary:Shows how to configure the style and color of a text underline.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Underline(Underline::Dotted);
        builder->get_Font()->set_UnderlineColor(System::Drawing::Color::get_Red());

        builder->Writeln(u"Underlined text.");

        doc->Save(ArtifactsDir + u"Font.Underlines.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Underlines.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Underlined text.", run->GetText().Trim());
        ASSERT_EQ(Underline::Dotted, run->get_Font()->get_Underline());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), run->get_Font()->get_UnderlineColor().ToArgb());
    }

    void ComplexScript()
    {
        //ExStart
        //ExFor:Font.ComplexScript
        //ExSummary:Shows how to add text that is always treated as complex script.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_ComplexScript(true);

        builder->Writeln(u"Text treated as complex script.");

        doc->Save(ArtifactsDir + u"Font.ComplexScript.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.ComplexScript.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Text treated as complex script.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_ComplexScript());
    }

    void SparklingText()
    {
        //ExStart
        //ExFor:Font.TextEffect
        //ExSummary:Shows how to apply a visual effect to a run.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Size(36);
        builder->get_Font()->set_TextEffect(TextEffect::SparkleText);

        builder->Writeln(u"Text with a sparkle effect.");

        // Older versions of Microsoft Word only support font animation effects.
        doc->Save(ArtifactsDir + u"Font.SparklingText.doc");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.SparklingText.doc");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Text with a sparkle effect.", run->GetText().Trim());
        ASSERT_EQ(TextEffect::SparkleText, run->get_Font()->get_TextEffect());
    }

    void Shading_()
    {
        //ExStart
        //ExFor:Font.Shading
        //ExSummary:Shows how to apply shading to text created by a document builder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Color(System::Drawing::Color::get_White());

        // One way to make the text created using our white font color visible
        // is to apply a background shading effect.
        SharedPtr<Shading> shading = builder->get_Font()->get_Shading();
        shading->set_Texture(TextureIndex::TextureDiagonalUp);
        shading->set_BackgroundPatternColor(System::Drawing::Color::get_OrangeRed());
        shading->set_ForegroundPatternColor(System::Drawing::Color::get_DarkBlue());

        builder->Writeln(u"White text on an orange background with a two-tone texture.");

        doc->Save(ArtifactsDir + u"Font.Shading.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Shading.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"White text on an orange background with a two-tone texture.", run->GetText().Trim());
        ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), run->get_Font()->get_Color().ToArgb());

        ASSERT_EQ(TextureIndex::TextureDiagonalUp, run->get_Font()->get_Shading()->get_Texture());
        ASSERT_EQ(System::Drawing::Color::get_OrangeRed().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_DarkBlue().ToArgb(), run->get_Font()->get_Shading()->get_ForegroundPatternColor().ToArgb());
    }

    void Bidi()
    {
        //ExStart
        //ExFor:Font.Bidi
        //ExFor:Font.NameBi
        //ExFor:Font.SizeBi
        //ExFor:Font.ItalicBi
        //ExFor:Font.BoldBi
        //ExFor:Font.LocaleIdBi
        //ExSummary:Shows how to define separate sets of font settings for right-to-left, and right-to-left text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Define a set of font settings for left-to-right text.
        builder->get_Font()->set_Name(u"Courier New");
        builder->get_Font()->set_Size(16);
        builder->get_Font()->set_Italic(false);
        builder->get_Font()->set_Bold(false);
        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"en-US", false)->get_LCID());

        // Define another set of font settings for right-to-left text.
        builder->get_Font()->set_NameBi(u"Andalus");
        builder->get_Font()->set_SizeBi(24);
        builder->get_Font()->set_ItalicBi(true);
        builder->get_Font()->set_BoldBi(true);
        builder->get_Font()->set_LocaleIdBi(MakeObject<System::Globalization::CultureInfo>(u"ar-AR", false)->get_LCID());

        // We can use the Bidi flag to indicate whether the text we are about to add
        // with the document builder is right-to-left. When we add text with this flag set to true,
        // it will be formatted using the right-to-left set of font settings.
        builder->get_Font()->set_Bidi(true);
        builder->Write(u"مرحبًا");

        // Set the flag to false, and then add left-to-right text.
        // The document builder will format these using the left-to-right set of font settings.
        builder->get_Font()->set_Bidi(false);
        builder->Write(u" Hello world!");

        doc->Save(ArtifactsDir + u"Font.Bidi.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Bidi.docx");

        for (const auto& run : System::IterateOver<Run>(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()))
        {
            switch (doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->IndexOf(run))
            {
            case 0:
                ASSERT_EQ(u"مرحبًا", run->GetText().Trim());
                ASSERT_TRUE(run->get_Font()->get_Bidi());
                break;

            case 1:
                ASSERT_EQ(u"Hello world!", run->GetText().Trim());
                ASSERT_FALSE(run->get_Font()->get_Bidi());
                break;
            }

            ASSERT_EQ(1033, run->get_Font()->get_LocaleId());
            ASPOSE_ASSERT_EQ(16, run->get_Font()->get_Size());
            ASSERT_EQ(u"Courier New", run->get_Font()->get_Name());
            ASSERT_FALSE(run->get_Font()->get_Italic());
            ASSERT_FALSE(run->get_Font()->get_Bold());
            ASSERT_EQ(1025, run->get_Font()->get_LocaleIdBi());
            ASPOSE_ASSERT_EQ(24, run->get_Font()->get_SizeBi());
            ASSERT_EQ(u"Andalus", run->get_Font()->get_NameBi());
            ASSERT_TRUE(run->get_Font()->get_ItalicBi());
            ASSERT_TRUE(run->get_Font()->get_BoldBi());
        }
    }

    void FarEast()
    {
        //ExStart
        //ExFor:Font.NameFarEast
        //ExFor:Font.LocaleIdFarEast
        //ExSummary:Shows how to insert and format text in a Far East language.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify font settings that the document builder will apply to any text that it inserts.
        builder->get_Font()->set_Name(u"Courier New");
        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"en-US", false)->get_LCID());

        // Name "FarEast" equivalents for our font and locale.
        // If the builder inserts Asian characters with this Font configuration, then each run that contains
        // these characters will display them using the "FarEast" font/locale instead of the default.
        // This could be useful when a western font does not have ideal representations for Asian characters.
        builder->get_Font()->set_NameFarEast(u"SimSun");
        builder->get_Font()->set_LocaleIdFarEast(MakeObject<System::Globalization::CultureInfo>(u"zh-CN", false)->get_LCID());

        // This text will be displayed in the default font/locale.
        builder->Writeln(u"Hello world!");

        // Since these are Asian characters, this run will apply our "FarEast" font/locale equivalents.
        builder->Writeln(u"你好世界");

        doc->Save(ArtifactsDir + u"Font.FarEast.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.FarEast.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Hello world!", run->GetText().Trim());
        ASSERT_EQ(1033, run->get_Font()->get_LocaleId());
        ASSERT_EQ(u"Courier New", run->get_Font()->get_Name());
        ASSERT_EQ(2052, run->get_Font()->get_LocaleIdFarEast());
        ASSERT_EQ(u"SimSun", run->get_Font()->get_NameFarEast());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"你好世界", run->GetText().Trim());
        ASSERT_EQ(1033, run->get_Font()->get_LocaleId());
        ASSERT_EQ(u"SimSun", run->get_Font()->get_Name());
        ASSERT_EQ(2052, run->get_Font()->get_LocaleIdFarEast());
        ASSERT_EQ(u"SimSun", run->get_Font()->get_NameFarEast());
    }

    void NameAscii()
    {
        //ExStart
        //ExFor:Font.NameAscii
        //ExFor:Font.NameOther
        //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Suppose a run that we use the builder to insert while using this font configuration
        // contains characters within the ASCII characters' range. In that case,
        // it will display those characters using this font.
        builder->get_Font()->set_NameAscii(u"Calibri");

        // With no other font specified, the builder will also apply this font to all characters that it inserts.
        ASSERT_EQ(u"Calibri", builder->get_Font()->get_Name());

        // Specify a font to use for all characters outside of the ASCII range.
        // Ideally, this font should have a glyph for each required non-ASCII character code.
        builder->get_Font()->set_NameOther(u"Courier New");

        // Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range.
        // Each character will be displayed using either of the fonts, depending on.
        builder->Writeln(u"Hello, Привет");

        doc->Save(ArtifactsDir + u"Font.NameAscii.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.NameAscii.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Hello, Привет", run->GetText().Trim());
        ASSERT_EQ(u"Calibri", run->get_Font()->get_Name());
        ASSERT_EQ(u"Calibri", run->get_Font()->get_NameAscii());
        ASSERT_EQ(u"Courier New", run->get_Font()->get_NameOther());
    }

    void ChangeStyle()
    {
        //ExStart
        //ExFor:Font.StyleName
        //ExFor:Font.StyleIdentifier
        //ExFor:StyleIdentifier
        //ExSummary:Shows how to change the style of existing text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two ways of referencing styles.
        // 1 -  Using the style name:
        builder->get_Font()->set_StyleName(u"Emphasis");
        builder->Writeln(u"Text originally in \"Emphasis\" style");

        // 2 -  Using a built-in style identifier:
        builder->get_Font()->set_StyleIdentifier(StyleIdentifier::IntenseEmphasis);
        builder->Writeln(u"Text originally in \"Intense Emphasis\" style");

        // Convert all uses of one style to another,
        // using the above methods to reference old and new styles.
        for (const auto& run : System::IterateOver(doc->GetChildNodes(NodeType::Run, true)->LINQ_OfType<SharedPtr<Run>>()))
        {
            if (run->get_Font()->get_StyleName() == u"Emphasis")
            {
                run->get_Font()->set_StyleName(u"Strong");
            }

            if (run->get_Font()->get_StyleIdentifier() == StyleIdentifier::IntenseEmphasis)
            {
                run->get_Font()->set_StyleIdentifier(StyleIdentifier::Strong);
            }
        }

        doc->Save(ArtifactsDir + u"Font.ChangeStyle.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.ChangeStyle.docx");
        SharedPtr<Run> docRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Text originally in \"Emphasis\" style", docRun->GetText().Trim());
        ASSERT_EQ(StyleIdentifier::Strong, docRun->get_Font()->get_StyleIdentifier());
        ASSERT_EQ(u"Strong", docRun->get_Font()->get_StyleName());

        docRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Text originally in \"Intense Emphasis\" style", docRun->GetText().Trim());
        ASSERT_EQ(StyleIdentifier::Strong, docRun->get_Font()->get_StyleIdentifier());
        ASSERT_EQ(u"Strong", docRun->get_Font()->get_StyleName());
    }

    void BuiltIn()
    {
        //ExStart
        //ExFor:Style.BuiltIn
        //ExSummary:Shows how to differentiate custom styles from built-in styles.
        auto doc = MakeObject<Document>();

        // When we create a document using Microsoft Word, or programmatically using Aspose.Words,
        // the document will come with a collection of styles to apply to its text to modify its appearance.
        // We can access these built-in styles via the document's "Styles" collection.
        // These styles will all have the "BuiltIn" flag set to "true".
        SharedPtr<Style> style = doc->get_Styles()->idx_get(u"Emphasis");

        ASSERT_TRUE(style->get_BuiltIn());

        // Create a custom style and add it to the collection.
        // Custom styles such as this will have the "BuiltIn" flag set to "false".
        style = doc->get_Styles()->Add(StyleType::Character, u"MyStyle");
        style->get_Font()->set_Color(System::Drawing::Color::get_Navy());
        style->get_Font()->set_Name(u"Courier New");

        ASSERT_FALSE(style->get_BuiltIn());
        //ExEnd
    }

    void Style_()
    {
        //ExStart
        //ExFor:Font.Style
        //ExSummary:Applies a double underline to all runs in a document that are formatted with custom character styles.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a custom style and apply it to text created using a document builder.
        SharedPtr<Style> style = doc->get_Styles()->Add(StyleType::Character, u"MyStyle");
        style->get_Font()->set_Color(System::Drawing::Color::get_Red());
        style->get_Font()->set_Name(u"Courier New");

        builder->get_Font()->set_StyleName(u"MyStyle");
        builder->Write(u"This text is in a custom style.");

        // Iterate over every run and add a double underline to every custom style.
        for (const auto& run : System::IterateOver(doc->GetChildNodes(NodeType::Run, true)->LINQ_OfType<SharedPtr<Run>>()))
        {
            SharedPtr<Style> charStyle = run->get_Font()->get_Style();

            if (!charStyle->get_BuiltIn())
            {
                run->get_Font()->set_Underline(Underline::Double);
            }
        }

        doc->Save(ArtifactsDir + u"Font.Style.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Style.docx");
        SharedPtr<Run> docRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"This text is in a custom style.", docRun->GetText().Trim());
        ASSERT_EQ(u"MyStyle", docRun->get_Font()->get_StyleName());
        ASSERT_FALSE(docRun->get_Font()->get_Style()->get_BuiltIn());
        ASSERT_EQ(Underline::Double, docRun->get_Font()->get_Underline());
    }

    void GetAvailableFonts()
    {
        //ExStart
        //ExFor:Fonts.PhysicalFontInfo
        //ExFor:FontSourceBase.GetAvailableFonts
        //ExFor:PhysicalFontInfo.FontFamilyName
        //ExFor:PhysicalFontInfo.FullFontName
        //ExFor:PhysicalFontInfo.Version
        //ExFor:PhysicalFontInfo.FilePath
        //ExSummary:Shows how to list available fonts.
        // Configure Aspose.Words to source fonts from a custom folder, and then print every available font.
        ArrayPtr<SharedPtr<FontSourceBase>> folderFontSource = MakeArray<SharedPtr<FontSourceBase>>({MakeObject<FolderFontSource>(FontsDir, true)});

        for (const auto& fontInfo : System::IterateOver(folderFontSource[0]->GetAvailableFonts()))
        {
            std::cout << "FontFamilyName : " << fontInfo->get_FontFamilyName() << std::endl;
            std::cout << "FullFontName  : " << fontInfo->get_FullFontName() << std::endl;
            std::cout << "Version  : " << fontInfo->get_Version() << std::endl;
            std::cout << "FilePath : " << fontInfo->get_FilePath() << "\n" << std::endl;
        }
        //ExEnd

        ASSERT_EQ(folderFontSource[0]->GetAvailableFonts()->get_Count(),
                  System::IO::Directory::EnumerateFiles(FontsDir, u"*.*", System::IO::SearchOption::AllDirectories)
                      ->LINQ_Count(static_cast<System::Func<String, bool>>(
                          static_cast<std::function<bool(String f)>>([](String f) -> bool { return f.EndsWith(u".ttf") || f.EndsWith(u".otf"); }))));
    }

    void SetFontAutoColor()
    {
        //ExStart
        //ExFor:Font.AutoColor
        //ExSummary:Shows how to improve readability by automatically selecting text color based on the brightness of its background.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If a run's Font object does not specify text color, it will automatically
        // select either black or white depending on the background color's color.
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), builder->get_Font()->get_Color().ToArgb());

        // The default color for text is black. If the color of the background is dark, black text will be difficult to see.
        // To solve this problem, the AutoColor property will display this text in white.
        builder->get_Font()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_DarkBlue());

        builder->Writeln(u"The text color automatically chosen for this run is white.");

        ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(),
                  doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->get_Font()->get_AutoColor().ToArgb());

        // If we change the background to a light color, black will be a more
        // suitable text color than white so that the auto color will display it in black.
        builder->get_Font()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());

        builder->Writeln(u"The text color automatically chosen for this run is black.");

        ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(),
                  doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_AutoColor().ToArgb());

        doc->Save(ArtifactsDir + u"Font.SetFontAutoColor.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.SetFontAutoColor.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"The text color automatically chosen for this run is white.", run->GetText().Trim());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), run->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_DarkBlue().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"The text color automatically chosen for this run is black.", run->GetText().Trim());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), run->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());
    }

    //ExStart
    //ExFor:Font.Hidden
    //ExFor:Paragraph.Accept
    //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
    //ExFor:DocumentVisitor.VisitFormField(FormField)
    //ExFor:DocumentVisitor.VisitTableEnd(Table)
    //ExFor:DocumentVisitor.VisitCellEnd(Cell)
    //ExFor:DocumentVisitor.VisitRowEnd(Row)
    //ExFor:DocumentVisitor.VisitSpecialChar(SpecialChar)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
    //ExFor:DocumentVisitor.VisitShapeStart(Shape)
    //ExFor:DocumentVisitor.VisitCommentStart(Comment)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
    //ExFor:SpecialChar
    //ExFor:Node.Accept
    //ExFor:Paragraph.ParagraphBreakFont
    //ExFor:Table.Accept
    //ExSummary:Shows how to use a DocumentVisitor implementation to remove all hidden content from a document.
    void RemoveHiddenContentFromDocument()
    {
        auto doc = MakeObject<Document>(MyDir + u"Hidden content.docx");
        ASSERT_EQ(26, doc->GetChildNodes(NodeType::Paragraph, true)->get_Count());
        //ExSkip
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        //ExSkip

        auto hiddenContentRemover = MakeObject<ExFont::RemoveHiddenContentVisitor>();

        // Below are three types of fields which can accept a document visitor,
        // which will allow it to visit the accepting node, and then traverse its child nodes in a depth-first manner.
        // 1 -  Paragraph node:
        auto para = System::ExplicitCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 4, true));
        para->Accept(hiddenContentRemover);

        // 2 -  Table node:
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        table->Accept(hiddenContentRemover);

        // 3 -  Document node:
        doc->Accept(hiddenContentRemover);

        doc->Save(ArtifactsDir + u"Font.RemoveHiddenContentFromDocument.docx");
        TestRemoveHiddenContent(MakeObject<Document>(ArtifactsDir + u"Font.RemoveHiddenContentFromDocument.docx"));
        //ExSkip
    }

    /// <summary>
    /// Removes all visited nodes marked as "hidden content".
    /// </summary>
    class RemoveHiddenContentVisitor : public DocumentVisitor
    {
    public:
        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldStart(SharedPtr<FieldStart> fieldStart) override
        {
            if (fieldStart->get_Font()->get_Hidden())
            {
                fieldStart->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldEnd(SharedPtr<FieldEnd> fieldEnd) override
        {
            if (fieldEnd->get_Font()->get_Hidden())
            {
                fieldEnd->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldSeparator(SharedPtr<FieldSeparator> fieldSeparator) override
        {
            if (fieldSeparator->get_Font()->get_Hidden())
            {
                fieldSeparator->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (run->get_Font()->get_Hidden())
            {
                run->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        VisitorAction VisitParagraphStart(SharedPtr<Paragraph> paragraph) override
        {
            if (paragraph->get_ParagraphBreakFont()->get_Hidden())
            {
                paragraph->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FormField is encountered in the document.
        /// </summary>
        VisitorAction VisitFormField(SharedPtr<FormField> formField) override
        {
            if (formField->get_Font()->get_Hidden())
            {
                formField->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a GroupShape is encountered in the document.
        /// </summary>
        VisitorAction VisitGroupShapeStart(SharedPtr<GroupShape> groupShape) override
        {
            if (groupShape->get_Font()->get_Hidden())
            {
                groupShape->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Shape is encountered in the document.
        /// </summary>
        VisitorAction VisitShapeStart(SharedPtr<Shape> shape) override
        {
            if (shape->get_Font()->get_Hidden())
            {
                shape->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Comment is encountered in the document.
        /// </summary>
        VisitorAction VisitCommentStart(SharedPtr<Comment> comment) override
        {
            if (comment->get_Font()->get_Hidden())
            {
                comment->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Footnote is encountered in the document.
        /// </summary>
        VisitorAction VisitFootnoteStart(SharedPtr<Footnote> footnote) override
        {
            if (footnote->get_Font()->get_Hidden())
            {
                footnote->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a SpecialCharacter is encountered in the document.
        /// </summary>
        VisitorAction VisitSpecialChar(SharedPtr<SpecialChar> specialChar) override
        {
            if (specialChar->get_Font()->get_Hidden())
            {
                specialChar->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when visiting of a Table node is ended in the document.
        /// </summary>
        VisitorAction VisitTableEnd(SharedPtr<Table> table) override
        {
            // The content inside table cells may have the hidden content flag, but the tables themselves cannot.
            // If this table had nothing but hidden content, this visitor would have removed all of it,
            // and there would be no child nodes left.
            // Thus, we can also treat the table itself as hidden content and remove it.
            // Tables which are empty but do not have hidden content will have cells with empty paragraphs inside,
            // which this visitor will not remove.
            if (!table->get_HasChildNodes())
            {
                table->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when visiting of a Cell node is ended in the document.
        /// </summary>
        VisitorAction VisitCellEnd(SharedPtr<Cell> cell) override
        {
            if (!cell->get_HasChildNodes() && cell->get_ParentNode() != nullptr)
            {
                cell->Remove();
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when visiting of a Row node is ended in the document.
        /// </summary>
        VisitorAction VisitRowEnd(SharedPtr<Row> row) override
        {
            if (!row->get_HasChildNodes() && row->get_ParentNode() != nullptr)
            {
                row->Remove();
            }

            return VisitorAction::Continue;
        }
    };
    //ExEnd

    void TestRemoveHiddenContent(SharedPtr<Document> doc)
    {
        ASSERT_EQ(20, doc->GetChildNodes(NodeType::Paragraph, true)->get_Count());
        ASSERT_EQ(1, doc->GetChildNodes(NodeType::Table, true)->get_Count());

        for (const auto& node : System::IterateOver(doc->GetChildNodes(NodeType::Any, true)))
        {
            if (System::ObjectExt::Is<FieldStart>(node))
            {
                auto fieldStart = System::ExplicitCast<FieldStart>(node);
                ASSERT_FALSE(fieldStart->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<FieldEnd>(node))
            {
                auto fieldEnd = System::ExplicitCast<FieldEnd>(node);
                ASSERT_FALSE(fieldEnd->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<FieldSeparator>(node))
            {
                auto fieldSeparator = System::ExplicitCast<FieldSeparator>(node);
                ASSERT_FALSE(fieldSeparator->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Run>(node))
            {
                auto run = System::ExplicitCast<Run>(node);
                ASSERT_FALSE(run->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Paragraph>(node))
            {
                auto paragraph = System::ExplicitCast<Paragraph>(node);
                ASSERT_FALSE(paragraph->get_ParagraphBreakFont()->get_Hidden());
            }
            else if (System::ObjectExt::Is<FormField>(node))
            {
                auto formField = System::ExplicitCast<FormField>(node);
                ASSERT_FALSE(formField->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<GroupShape>(node))
            {
                auto groupShape = System::ExplicitCast<GroupShape>(node);
                ASSERT_FALSE(groupShape->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Shape>(node))
            {
                auto shape = System::ExplicitCast<Shape>(node);
                ASSERT_FALSE(shape->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Comment>(node))
            {
                auto comment = System::ExplicitCast<Comment>(node);
                ASSERT_FALSE(comment->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Footnote>(node))
            {
                auto footnote = System::ExplicitCast<Footnote>(node);
                ASSERT_FALSE(footnote->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<SpecialChar>(node))
            {
                auto specialChar = System::ExplicitCast<SpecialChar>(node);
                ASSERT_FALSE(specialChar->get_Font()->get_Hidden());
            }
        }
    }

    void DefaultFonts()
    {
        //ExStart
        //ExFor:Fonts.FontInfoCollection.Contains(String)
        //ExFor:Fonts.FontInfoCollection.Count
        //ExSummary:Shows info about the fonts that are present in the blank document.
        auto doc = MakeObject<Document>();

        // A blank document contains 3 default fonts. Each font in the document
        // will have a corresponding FontInfo object which contains details about that font.
        ASSERT_EQ(3, doc->get_FontInfos()->get_Count());

        ASSERT_TRUE(doc->get_FontInfos()->Contains(u"Times New Roman"));
        ASSERT_EQ(204, doc->get_FontInfos()->idx_get(u"Times New Roman")->get_Charset());

        ASSERT_TRUE(doc->get_FontInfos()->Contains(u"Symbol"));
        ASSERT_TRUE(doc->get_FontInfos()->Contains(u"Arial"));
        //ExEnd
    }

    void ExtractEmbeddedFont()
    {
        //ExStart
        //ExFor:Fonts.EmbeddedFontFormat
        //ExFor:Fonts.EmbeddedFontStyle
        //ExFor:Fonts.FontInfo.GetEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
        //ExFor:Fonts.FontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle)
        //ExFor:Fonts.FontInfoCollection.Item(Int32)
        //ExFor:Fonts.FontInfoCollection.Item(String)
        //ExSummary:Shows how to extract an embedded font from a document, and save it to the local file system.
        auto doc = MakeObject<Document>(MyDir + u"Embedded font.docx");

        SharedPtr<FontInfo> embeddedFont = doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift");
        ArrayPtr<uint8_t> embeddedFontBytes = embeddedFont->GetEmbeddedFont(EmbeddedFontFormat::OpenType, EmbeddedFontStyle::Regular);
        ASSERT_FALSE(embeddedFontBytes == nullptr);
        //ExSkip

        System::IO::File::WriteAllBytes(ArtifactsDir + u"Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);

        // Embedded font formats may be different in other formats such as .doc.
        // We need to know the correct format before we can extract the font.
        doc = MakeObject<Document>(MyDir + u"Embedded font.doc");

        ASSERT_TRUE(doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFont(EmbeddedFontFormat::OpenType, EmbeddedFontStyle::Regular) ==
                    nullptr);
        ASSERT_FALSE(
            doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFont(EmbeddedFontFormat::EmbeddedOpenType, EmbeddedFontStyle::Regular) ==
            nullptr);

        // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType.
        embeddedFontBytes = doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFontAsOpenType(EmbeddedFontStyle::Regular);

        System::IO::File::WriteAllBytes(ArtifactsDir + u"Alte DIN 1451 Mittelschrift.otf", embeddedFontBytes);
        //ExEnd
    }

    void GetFontInfoFromFile()
    {
        //ExStart
        //ExFor:Fonts.FontFamily
        //ExFor:Fonts.FontPitch
        //ExFor:Fonts.FontInfo.AltName
        //ExFor:Fonts.FontInfo.Charset
        //ExFor:Fonts.FontInfo.Family
        //ExFor:Fonts.FontInfo.Panose
        //ExFor:Fonts.FontInfo.Pitch
        //ExFor:Fonts.FontInfoCollection.GetEnumerator
        //ExSummary:Shows how to access and print details of each font in a document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<FontInfo>>> fontCollectionEnumerator = doc->get_FontInfos()->GetEnumerator();
        while (fontCollectionEnumerator->MoveNext())
        {
            SharedPtr<FontInfo> fontInfo = fontCollectionEnumerator->get_Current();
            if (fontInfo != nullptr)
            {
                std::cout << (String(u"Font name: ") + fontInfo->get_Name()) << std::endl;

                // Alt names are usually blank.
                std::cout << (String(u"Alt name: ") + fontInfo->get_AltName()) << std::endl;
                std::cout << (String(u"\t- Family: ") + System::ObjectExt::ToString(fontInfo->get_Family())) << std::endl;
                std::cout << (String(u"\t- ") + (fontInfo->get_IsTrueType() ? String(u"Is TrueType") : String(u"Is not TrueType"))) << std::endl;
                std::cout << (String(u"\t- Pitch: ") + System::ObjectExt::ToString(fontInfo->get_Pitch())) << std::endl;
                std::cout << (String(u"\t- Charset: ") + fontInfo->get_Charset()) << std::endl;
                std::cout << "\t- Panose:" << std::endl;
                std::cout << (String(u"\t\tFamily Kind: ") + fontInfo->get_Panose()[0]) << std::endl;
                std::cout << (String(u"\t\tSerif Style: ") + fontInfo->get_Panose()[1]) << std::endl;
                std::cout << (String(u"\t\tWeight: ") + fontInfo->get_Panose()[2]) << std::endl;
                std::cout << (String(u"\t\tProportion: ") + fontInfo->get_Panose()[3]) << std::endl;
                std::cout << (String(u"\t\tContrast: ") + fontInfo->get_Panose()[4]) << std::endl;
                std::cout << (String(u"\t\tStroke Variation: ") + fontInfo->get_Panose()[5]) << std::endl;
                std::cout << (String(u"\t\tArm Style: ") + fontInfo->get_Panose()[6]) << std::endl;
                std::cout << (String(u"\t\tLetterform: ") + fontInfo->get_Panose()[7]) << std::endl;
                std::cout << (String(u"\t\tMidline: ") + fontInfo->get_Panose()[8]) << std::endl;
                std::cout << (String(u"\t\tX-Height: ") + fontInfo->get_Panose()[9]) << std::endl;
            }
        }
        //ExEnd

        ASPOSE_ASSERT_EQ(MakeArray<int>({2, 15, 5, 2, 2, 2, 4, 3, 2, 4}), doc->get_FontInfos()->idx_get(u"Calibri")->get_Panose());
        ASPOSE_ASSERT_EQ(MakeArray<int>({2, 15, 3, 2, 2, 2, 4, 3, 2, 4}), doc->get_FontInfos()->idx_get(u"Calibri Light")->get_Panose());
        ASPOSE_ASSERT_EQ(MakeArray<int>({2, 2, 6, 3, 5, 4, 5, 2, 3, 4}), doc->get_FontInfos()->idx_get(u"Times New Roman")->get_Panose());
    }

    void LineSpacing()
    {
        //ExStart
        //ExFor:Font.LineSpacing
        //ExSummary:Shows how to get a font's line spacing, in points.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set different fonts for the DocumentBuilder and verify their line spacing.
        builder->get_Font()->set_Name(u"Calibri");
        ASPOSE_ASSERT_EQ(14.6484375, builder->get_Font()->get_LineSpacing());

        builder->get_Font()->set_Name(u"Times New Roman");
        ASPOSE_ASSERT_EQ(13.798828125, builder->get_Font()->get_LineSpacing());
        //ExEnd
    }

    void HasDmlEffect()
    {
        //ExStart
        //ExFor:Font.HasDmlEffect(TextDmlEffect)
        //ExSummary:Shows how to check if a run displays a DrawingML text effect.
        auto doc = MakeObject<Document>(MyDir + u"DrawingML text effects.docx");

        SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();

        ASSERT_TRUE(runs->idx_get(0)->get_Font()->HasDmlEffect(TextDmlEffect::Shadow));
        ASSERT_TRUE(runs->idx_get(1)->get_Font()->HasDmlEffect(TextDmlEffect::Shadow));
        ASSERT_TRUE(runs->idx_get(2)->get_Font()->HasDmlEffect(TextDmlEffect::Reflection));
        ASSERT_TRUE(runs->idx_get(3)->get_Font()->HasDmlEffect(TextDmlEffect::Effect3D));
        ASSERT_TRUE(runs->idx_get(4)->get_Font()->HasDmlEffect(TextDmlEffect::Fill));
        //ExEnd
    }

    void SetEmphasisMark(EmphasisMark emphasisMark)
    {
        //ExStart
        //ExFor:EmphasisMark
        //ExFor:Font.EmphasisMark
        //ExSummary:Shows how to add additional character rendered above/below the glyph-character.
        auto builder = MakeObject<DocumentBuilder>();

        // Possible types of emphasis mark:
        // https://apireference.aspose.com/words/net/aspose.words/emphasismark
        builder->get_Font()->set_EmphasisMark(emphasisMark);

        builder->Write(u"Emphasis text");
        builder->Writeln();
        builder->get_Font()->ClearFormatting();
        builder->Write(u"Simple text");

        builder->get_Document()->Save(ArtifactsDir + u"Fonts.SetEmphasisMark.docx");
        //ExEnd
    }

    void ThemeFontsColors()
    {
        //ExStart
        //ExFor:Font.ThemeFont
        //ExFor:Font.ThemeFontAscii
        //ExFor:Font.ThemeFontBi
        //ExFor:Font.ThemeFontFarEast
        //ExFor:Font.ThemeFontOther
        //ExFor:Font.ThemeColor
        //ExFor:ThemeFont
        //ExFor:ThemeColor
        //ExSummary:Shows how to work with theme fonts and colors.
        auto doc = MakeObject<Document>();

        // Define fonts for languages uses by default.
        doc->get_Theme()->get_MinorFonts()->set_Latin(u"Algerian");
        doc->get_Theme()->get_MinorFonts()->set_EastAsian(u"Aharoni");
        doc->get_Theme()->get_MinorFonts()->set_ComplexScript(u"Andalus");

        SharedPtr<Font> font = doc->get_Styles()->idx_get(u"Normal")->get_Font();
        std::cout << String::Format(u"Originally the Normal style theme color is: {0} and RGB color is: {1}\n", font->get_ThemeColor(), font->get_Color())
                  << std::endl;

        // We can use theme font and color instead of default values.
        font->set_ThemeFont(ThemeFont::Minor);
        font->set_ThemeColor(ThemeColor::Accent2);

        ASSERT_EQ(ThemeFont::Minor, font->get_ThemeFont());
        ASSERT_EQ(u"Algerian", font->get_Name());

        ASSERT_EQ(ThemeFont::Minor, font->get_ThemeFontAscii());
        ASSERT_EQ(u"Algerian", font->get_NameAscii());

        ASSERT_EQ(ThemeFont::Minor, font->get_ThemeFontBi());
        ASSERT_EQ(u"Andalus", font->get_NameBi());

        ASSERT_EQ(ThemeFont::Minor, font->get_ThemeFontFarEast());
        ASSERT_EQ(u"Aharoni", font->get_NameFarEast());

        ASSERT_EQ(ThemeFont::Minor, font->get_ThemeFontOther());
        ASSERT_EQ(u"Algerian", font->get_NameOther());

        ASSERT_EQ(ThemeColor::Accent2, font->get_ThemeColor());
        ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, font->get_Color());

        // There are several ways of reset them font and color.
        // 1 -  By setting ThemeFont.None/ThemeColor.None:
        font->set_ThemeFont(ThemeFont::None);
        font->set_ThemeColor(ThemeColor::None);

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFont());
        ASSERT_EQ(u"Algerian", font->get_Name());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontAscii());
        ASSERT_EQ(u"Algerian", font->get_NameAscii());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontBi());
        ASSERT_EQ(u"Andalus", font->get_NameBi());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontFarEast());
        ASSERT_EQ(u"Aharoni", font->get_NameFarEast());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontOther());
        ASSERT_EQ(u"Algerian", font->get_NameOther());

        ASSERT_EQ(ThemeColor::None, font->get_ThemeColor());
        ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, font->get_Color());

        // 2 -  By setting non-theme font/color names:
        font->set_Name(u"Arial");
        font->set_Color(System::Drawing::Color::get_Blue());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFont());
        ASSERT_EQ(u"Arial", font->get_Name());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontAscii());
        ASSERT_EQ(u"Arial", font->get_NameAscii());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontBi());
        ASSERT_EQ(u"Arial", font->get_NameBi());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontFarEast());
        ASSERT_EQ(u"Arial", font->get_NameFarEast());

        ASSERT_EQ(ThemeFont::None, font->get_ThemeFontOther());
        ASSERT_EQ(u"Arial", font->get_NameOther());

        ASSERT_EQ(ThemeColor::None, font->get_ThemeColor());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), font->get_Color().ToArgb());
        //ExEnd
    }

    void CreateThemedStyle()
    {
        //ExStart
        //ExFor:Font.ThemeFont
        //ExFor:Font.ThemeColor
        //ExFor:Font.TintAndShade
        //ExFor:ThemeFont
        //ExFor:ThemeColor
        //ExSummary:Shows how to create and use themed style.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln();

        // Create some style with theme font properties.
        SharedPtr<Style> style = doc->get_Styles()->Add(StyleType::Paragraph, u"ThemedStyle");
        style->get_Font()->set_ThemeFont(ThemeFont::Major);
        style->get_Font()->set_ThemeColor(ThemeColor::Accent5);
        style->get_Font()->set_TintAndShade(0.3);

        builder->get_ParagraphFormat()->set_StyleName(u"ThemedStyle");
        builder->Writeln(u"Text with themed style");
        //ExEnd

        auto run = System::ExplicitCast<Run>((System::ExplicitCast<Paragraph>(builder->get_CurrentParagraph()->get_PreviousSibling()))->get_FirstChild());

        ASSERT_EQ(ThemeFont::Major, run->get_Font()->get_ThemeFont());
        ASSERT_EQ(u"Times New Roman", run->get_Font()->get_Name());

        ASSERT_EQ(ThemeFont::Major, run->get_Font()->get_ThemeFontAscii());
        ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameAscii());

        ASSERT_EQ(ThemeFont::Major, run->get_Font()->get_ThemeFontBi());
        ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameBi());

        ASSERT_EQ(ThemeFont::Major, run->get_Font()->get_ThemeFontFarEast());
        ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameFarEast());

        ASSERT_EQ(ThemeFont::Major, run->get_Font()->get_ThemeFontOther());
        ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameOther());

        ASSERT_EQ(ThemeColor::Accent5, run->get_Font()->get_ThemeColor());
        ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, run->get_Font()->get_Color());
    }
};

} // namespace ApiExamples
