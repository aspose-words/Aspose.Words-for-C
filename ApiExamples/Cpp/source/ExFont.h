#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Borders/TextureIndex.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/EmbeddedFontFormat.h>
#include <Aspose.Words.Cpp/Model/Fonts/EmbeddedFontStyle.h>
#include <Aspose.Words.Cpp/Model/Fonts/FileFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontConfigSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontFallbackSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontFamily.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontPitch.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceType.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/MemoryFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/PhysicalFontInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/StreamFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/TableSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Text/EmphasisMark.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/SpecialChar.h>
#include <Aspose.Words.Cpp/Model/Text/TextDmlEffect.h>
#include <Aspose.Words.Cpp/Model/Text/TextEffect.h>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ilist.h>
#include <system/convert.h>
#include <system/details/dispose_guard.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/environment.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/operating_system.h>
#include <system/platform_id.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <xml/xml_attribute.h>
#include <xml/xml_attribute_collection.h>
#include <xml/xml_document.h>
#include <xml/xml_name_table.h>
#include <xml/xml_namespace_manager.h>
#include <xml/xml_node.h>
#include <xml/xml_node_list.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Tables;

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
        //ExSummary:Shows how to add a formatted run of text to a document using the object model.
        auto doc = MakeObject<Document>();

        // Create a new run of text
        auto run = MakeObject<Run>(doc, u"Hello");

        // Specify character formatting for the run of text
        SharedPtr<Aspose::Words::Font> f = run->get_Font();
        f->set_Name(u"Courier New");
        f->set_Size(36);
        f->set_HighlightColor(System::Drawing::Color::get_Yellow());

        // Append the run of text to the end of the first paragraph
        // in the body of the first section of the document
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(run);
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Hello", run->GetText().Trim());
        ASSERT_EQ(u"Courier New", run->get_Font()->get_Name());
        ASPOSE_ASSERT_EQ(36, run->get_Font()->get_Size());
        ASSERT_EQ(System::Drawing::Color::get_Yellow().ToArgb(), run->get_Font()->get_HighlightColor().ToArgb());
    }

    void Caps()
    {
        //ExStart
        //ExFor:Font.AllCaps
        //ExFor:Font.SmallCaps
        //ExSummary:Shows how to use all capitals and small capitals character formatting properties.
        auto doc = MakeObject<Document>();
        auto para = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        auto run = MakeObject<Run>(doc, u"All capitals");
        run->get_Font()->set_AllCaps(true);
        para->AppendChild(run);

        run = MakeObject<Run>(doc, u"SMALL CAPITALS");
        run->get_Font()->set_SmallCaps(true);
        para->AppendChild(run);

        doc->Save(ArtifactsDir + u"Font.Caps.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Caps.docx");
        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"All capitals", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_AllCaps());

        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(1);

        ASSERT_EQ(u"SMALL CAPITALS", run->GetText().Trim());
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

        SharedPtr<FontInfoCollection> fonts = doc->get_FontInfos();
        ASSERT_EQ(5, fonts->get_Count());
        //ExSkip

        // The fonts info extracted from this document does not necessarily mean that the fonts themselves are
        // used in the document. If a font is present but not used then most likely they were referenced at some time
        // and then removed from the Document
        for (int i = 0; i < fonts->get_Count(); i++)
        {
            std::cout << "Font index #" << i << std::endl;
            std::cout << "\tName: " << fonts->idx_get(i)->get_Name() << std::endl;
            std::cout << "\tIs " << (fonts->idx_get(i)->get_IsTrueType() ? String(u"") : String(u"not ")) << "a trueType font" << std::endl;
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

    void FontInfoCollection_()
    {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfoCollection.EmbedTrueTypeFonts
        //ExFor:FontInfoCollection.EmbedSystemFonts
        //ExFor:FontInfoCollection.SaveSubsetFonts
        //ExSummary:Shows how to save a document with embedded TrueType fonts.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<FontInfoCollection> fontInfos = doc->get_FontInfos();
        fontInfos->set_EmbedTrueTypeFonts(true);
        fontInfos->set_EmbedSystemFonts(false);
        fontInfos->set_SaveSubsetFonts(false);

        doc->Save(ArtifactsDir + u"Font.FontInfoCollection.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.FontInfoCollection.docx");
        fontInfos = doc->get_FontInfos();

        ASSERT_TRUE(fontInfos->get_EmbedTrueTypeFonts());
        ASSERT_FALSE(fontInfos->get_EmbedSystemFonts());
        ASSERT_FALSE(fontInfos->get_SaveSubsetFonts());
    }

    void WorkWithEmbeddedFonts(bool embedTrueTypeFonts, bool embedSystemFonts, bool saveSubsetFonts)
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<FontInfoCollection> fontInfos = doc->get_FontInfos();
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
        //ExSummary:Shows how to use strike-through character formatting properties.
        auto doc = MakeObject<Document>();
        auto para = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        auto run = MakeObject<Run>(doc, u"Double strike through text");
        run->get_Font()->set_DoubleStrikeThrough(true);
        para->AppendChild(run);

        run = MakeObject<Run>(doc, u"Single strike through text");
        run->get_Font()->set_StrikeThrough(true);
        para->AppendChild(run);

        doc->Save(ArtifactsDir + u"Font.StrikeThrough.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.StrikeThrough.docx");
        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Double strike through text", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_DoubleStrikeThrough());

        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(1);

        ASSERT_EQ(u"Single strike through text", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_StrikeThrough());
    }

    void PositionSubscript()
    {
        //ExStart
        //ExFor:Font.Position
        //ExFor:Font.Subscript
        //ExFor:Font.Superscript
        //ExSummary:Shows how to use subscript, superscript, complex script, text effects, and baseline text position properties.
        auto doc = MakeObject<Document>();
        auto para = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        // Add a run of text that is raised 5 points above the baseline
        auto run = MakeObject<Run>(doc, u"Raised text");
        run->get_Font()->set_Position(5);
        para->AppendChild(run);

        // Add a run of normal text
        run = MakeObject<Run>(doc, u"Normal text");
        para->AppendChild(run);

        // Add a run of text that appears as subscript
        run = MakeObject<Run>(doc, u"Subscript");
        run->get_Font()->set_Subscript(true);
        para->AppendChild(run);

        // Add a run of text that appears as superscript
        run = MakeObject<Run>(doc, u"Superscript");
        run->get_Font()->set_Superscript(true);
        para->AppendChild(run);

        doc->Save(ArtifactsDir + u"Font.PositionSubscript.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.PositionSubscript.docx");
        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Raised text", run->GetText().Trim());
        ASPOSE_ASSERT_EQ(5, run->get_Font()->get_Position());

        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(2);

        ASSERT_EQ(u"Subscript", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Subscript());

        run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(3);

        ASSERT_EQ(u"Superscript", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Superscript());
    }

    void ScalingSpacing()
    {
        //ExStart
        //ExFor:Font.Scaling
        //ExFor:Font.Spacing
        //ExSummary:Shows how to use character scaling and spacing properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add a run of text with characters 150% width of normal characters
        builder->get_Font()->set_Scaling(150);
        builder->Writeln(u"Wide characters");

        // Add a run of text with extra 1pt space between characters
        builder->get_Font()->set_Spacing(1);
        builder->Writeln(u"Expanded by 1pt");

        // Add a run of text with space between characters reduced by 1pt
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
        //ExSummary:Shows how to italicize a run of text.
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
        //ExSummary:Shows the difference between embossing and engraving text via font formatting.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Size(36);
        builder->get_Font()->set_Color(System::Drawing::Color::get_White());
        builder->get_Font()->set_Engrave(true);

        builder->Writeln(u"This text is engraved.");

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
        builder->get_Font()->set_Size(36);
        builder->get_Font()->set_Shadow(true);

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
        builder->get_Font()->set_Size(36);
        builder->get_Font()->set_Outline(true);

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
        //ExSummary:Shows how to create a hidden run of text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Size(36);
        builder->get_Font()->set_Hidden(true);

        // With the Hidden flag set to true, we can add text that will be present but invisible in the document
        // It is not recommended to use this as a way of hiding sensitive information as the text is still easily reachable
        builder->Writeln(u"This text won't be visible in the document.");

        doc->Save(ArtifactsDir + u"Font.Hidden.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Hidden.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"This text won't be visible in the document.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_Hidden());
    }

    void Kerning()
    {
        //ExStart
        //ExFor:Font.Kerning
        //ExSummary:Shows how to specify the font size at which kerning starts.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Name(u"Arial Black");

        // Set the font's kerning size threshold and font size
        builder->get_Font()->set_Kerning(24);
        builder->get_Font()->set_Size(18);

        // The font size falls below the kerning threshold so kerning will not be applied
        builder->Writeln(u"TALLY. (Kerning not applied)");

        // If we add runs of text with a document builder's writing methods,
        // the Font attributes of any new runs will inherit the values from the Font attributes of the previous runs
        // The font size is still 18, and we will change the kerning threshold to a value below that
        builder->get_Font()->set_Kerning(12);

        // Kerning has now been applied to this run
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
        //ExSummary:Shows how to specify that the run of text is not to be spell checked by Microsoft Word.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_NoProofing(true);

        builder->Writeln(u"Proofing has been disabled for this run, so these spelking errrs will not display red lines underneath.");

        doc->Save(ArtifactsDir + u"Font.NoProofing.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.NoProofing.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Proofing has been disabled for this run, so these spelking errrs will not display red lines underneath.", run->GetText().Trim());
        ASSERT_TRUE(run->get_Font()->get_NoProofing());
    }

    void LocaleId()
    {
        //ExStart
        //ExFor:Font.LocaleId
        //ExSummary:Shows how to specify the language of a text run so Microsoft Word can use a proper spell checker.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify the locale so Microsoft Word recognizes this text as Russian
        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"ru-RU", false)->get_LCID());
        builder->Writeln(u"Привет!");

        doc->Save(ArtifactsDir + u"Font.LocaleId.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.LocaleId.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"Привет!", run->GetText().Trim());
        ASSERT_EQ(1049, run->get_Font()->get_LocaleId());
    }

    void Underlines()
    {
        //ExStart
        //ExFor:Font.Underline
        //ExFor:Font.UnderlineColor
        //ExSummary:Shows how use the underline character formatting properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set an underline color and style
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
        //ExSummary:Shows how to make a run that's always treated as complex script.
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

        // Font animation effects are only visible in older versions of Microsoft Word
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
        //ExSummary:Shows how to apply shading for a run of text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shading> shd = builder->get_Font()->get_Shading();
        shd->set_Texture(TextureIndex::TextureDiagonalUp);
        shd->set_BackgroundPatternColor(System::Drawing::Color::get_OrangeRed());
        shd->set_ForegroundPatternColor(System::Drawing::Color::get_DarkBlue());

        builder->get_Font()->set_Color(System::Drawing::Color::get_White());

        builder->Writeln(u"White text on an orange background with texture.");

        doc->Save(ArtifactsDir + u"Font.Shading.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Shading.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"White text on an orange background with texture.", run->GetText().Trim());
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
        //ExSummary:Shows how to insert and format right-to-left text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Signal to Microsoft Word that this run of text contains right-to-left text
        builder->get_Font()->set_Bidi(true);

        // Specify the font and font size to be used for the right-to-left text
        builder->get_Font()->set_NameBi(u"Andalus");
        builder->get_Font()->set_SizeBi(48);

        // Specify that the right-to-left text in this run is bold and italic
        builder->get_Font()->set_ItalicBi(true);
        builder->get_Font()->set_BoldBi(true);

        // Specify the locale so Microsoft Word recognizes this text as Arabic - Saudi Arabia
        builder->get_Font()->set_LocaleIdBi(MakeObject<System::Globalization::CultureInfo>(u"ar-AR", false)->get_LCID());

        // Insert some Arabic text
        builder->Writeln(u"مرحبًا");

        doc->Save(ArtifactsDir + u"Font.Bidi.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Bidi.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"مرحبًا", run->GetText().Trim());
        ASSERT_EQ(1033, run->get_Font()->get_LocaleId());
        ASSERT_TRUE(run->get_Font()->get_Bidi());
        ASPOSE_ASSERT_EQ(48, run->get_Font()->get_SizeBi());
        ASSERT_EQ(u"Andalus", run->get_Font()->get_NameBi());
        ASSERT_TRUE(run->get_Font()->get_ItalicBi());
        ASSERT_TRUE(run->get_Font()->get_BoldBi());
    }

    void FarEast()
    {
        //ExStart
        //ExFor:Font.NameFarEast
        //ExFor:Font.LocaleIdFarEast
        //ExSummary:Shows how to insert and format text in a Far East language.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify the font name
        builder->get_Font()->set_NameFarEast(u"SimSun");

        // Specify the locale so Microsoft Word recognizes this text as Chinese
        builder->get_Font()->set_LocaleIdFarEast(MakeObject<System::Globalization::CultureInfo>(u"zh-CN", false)->get_LCID());

        // Insert some Chinese text
        builder->Writeln(u"你好世界");

        doc->Save(ArtifactsDir + u"Font.FarEast.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.FarEast.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"你好世界", run->GetText().Trim());
        ASSERT_EQ(2052, run->get_Font()->get_LocaleIdFarEast());
        ASSERT_EQ(u"SimSun", run->get_Font()->get_NameFarEast());
    }

    void Names()
    {
        //ExStart
        //ExFor:Font.NameAscii
        //ExFor:Font.NameOther
        //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify a font to use for all characters that fall within the ASCII character set
        builder->get_Font()->set_NameAscii(u"Calibri");

        // Specify a font to use for all other characters
        // This font should have a glyph for every other required character code
        builder->get_Font()->set_NameOther(u"Courier New");

        // The builder's font is the ASCII font
        ASSERT_EQ(u"Calibri", builder->get_Font()->get_Name());

        // Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range
        // This will create a run with two fonts
        builder->Writeln(u"Hello, Привет");

        doc->Save(ArtifactsDir + u"Font.Names.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.Names.docx");
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
        //ExSummary:Shows how to use style name or identifier to find text formatted with a specific character style and apply different character style.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert text with two styles that will be replaced by another style
        builder->get_Font()->set_StyleIdentifier(StyleIdentifier::Emphasis);
        builder->Writeln(u"Text originally in \"Emphasis\" style");
        builder->get_Font()->set_StyleIdentifier(StyleIdentifier::IntenseEmphasis);
        builder->Writeln(u"Text originally in \"Intense Emphasis\" style");

        // Loop through every run node
        for (auto run : System::IterateOver(doc->GetChildNodes(NodeType::Run, true)->LINQ_OfType<SharedPtr<Run>>()))
        {
            // If the run's text is of the "Emphasis" style, referenced by name, change the style to "Strong"
            if (run->get_Font()->get_StyleName() == u"Emphasis")
            {
                run->get_Font()->set_StyleName(u"Strong");
            }

            // If the run's text style is "Intense Emphasis", change it to "Strong" also, but this time reference using a StyleIdentifier
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

    void Style_()
    {
        //ExStart
        //ExFor:Font.Style
        //ExFor:Style.BuiltIn
        //ExSummary:Applies double underline to all runs in a document that are formatted with custom character styles.
        // Document doc = new Document(MyDir + "Custom style.docx");
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a custom style
        SharedPtr<Style> style = doc->get_Styles()->Add(StyleType::Character, u"MyStyle");
        style->get_Font()->set_Color(System::Drawing::Color::get_Red());
        style->get_Font()->set_Name(u"Courier New");

        // Set the style of the current paragraph to our custom style
        // This will apply to only the text after the style separator
        builder->get_Font()->set_StyleName(u"MyStyle");
        builder->Write(u"This text is in a custom style.");

        // Iterate through every run node and apply underlines to the run if its style is not a built in style,
        // like the one we added
        for (auto node : System::IterateOver(doc->GetChildNodes(NodeType::Run, true)))
        {
            auto run = System::DynamicCast<Run>(node);
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

    void SubstitutionNotification()
    {
        // Store the font sources currently used so we can restore them later
        ArrayPtr<SharedPtr<FontSourceBase>> origFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        //ExStart
        //ExFor:IWarningCallback
        //ExFor:DocumentBase.WarningCallback
        //ExFor:Fonts.FontSettings.DefaultInstance
        //ExSummary:Demonstrates how to receive notifications of font substitutions by using IWarningCallback.
        // Load the document to render
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        auto callback = MakeObject<ExFont::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts
        FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default
        // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
        FontSettings::get_DefaultInstance()->SetFontsFolder(String::Empty, false);

        // Pass the save options along with the save path to the save method
        doc->Save(ArtifactsDir + u"Font.SubstitutionNotification.pdf");
        //ExEnd

        ASSERT_GT(callback->FontWarnings->get_Count(), 0);
        ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_WarningType() == WarningType::FontSubstitution);
        ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_Description() ==
                    u"Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font.");

        // Restore default fonts
        FontSettings::get_DefaultInstance()->SetFontsSources(origFontSources);
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
        //ExSummary:Shows how to get available fonts and information about them.
        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
        ArrayPtr<SharedPtr<FontSourceBase>> folderFontSource =
            MakeArray<SharedPtr<FontSourceBase>>({MakeObject<FolderFontSource>(FontsDir, true)});

        for (auto fontInfo : System::IterateOver(folderFontSource[0]->GetAvailableFonts()))
        {
            std::cout << "FontFamilyName : " << fontInfo->get_FontFamilyName() << std::endl;
            std::cout << "FullFontName  : " << fontInfo->get_FullFontName() << std::endl;
            std::cout << "Version  : " << fontInfo->get_Version() << std::endl;
            std::cout << "FilePath : " << fontInfo->get_FilePath() << "\n" << std::endl;
        }
        //ExEnd

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(String f)> _local_func_0 = [](String f)
        {
            return f.EndsWith(u".ttf");
        };

        ASSERT_EQ(folderFontSource[0]->GetAvailableFonts()->get_Count(),
                  System::IO::Directory::GetFiles(FontsDir)->LINQ_Count(static_cast<System::Func<String, bool>>(_local_func_0)));
    }

    //ExStart
    //ExFor:Fonts.FontInfoSubstitutionRule
    //ExFor:Fonts.FontSubstitutionSettings.FontInfoSubstitution
    //ExFor:IWarningCallback
    //ExFor:IWarningCallback.Warning(WarningInfo)
    //ExFor:WarningInfo
    //ExFor:WarningInfo.Description
    //ExFor:WarningInfo.WarningType
    //ExFor:WarningInfoCollection
    //ExFor:WarningInfoCollection.Warning(WarningInfo)
    //ExFor:WarningInfoCollection.GetEnumerator
    //ExFor:WarningInfoCollection.Clear
    //ExFor:WarningType
    //ExFor:DocumentBase.WarningCallback
    //ExSummary:Shows how to set the property for finding the closest match font among the available font sources instead missing font.
    void EnableFontSubstitution()
    {
        auto doc = MakeObject<Document>(MyDir + u"Missing font.docx");

        // Assign a custom warning callback
        auto substitutionWarningHandler = MakeObject<ExFont::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(substitutionWarningHandler);

        // Set a default font name and enable font substitution
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(true);

        // When saving the document with the missing font, we should get a warning
        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"Font.EnableFontSubstitution.pdf");

        // List all warnings using an enumerator
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<WarningInfo>>> warnings =
                substitutionWarningHandler->FontWarnings->GetEnumerator();
            while (warnings->MoveNext())
            {
                std::cout << warnings->get_Current()->get_Description() << std::endl;
            }
        }

        // Warnings are stored in this format
        ASSERT_EQ(WarningSource::Layout, substitutionWarningHandler->FontWarnings->idx_get(0)->get_Source());
        ASSERT_TRUE(MakeObject<System::Text::RegularExpressions::Regex>(
                        u"Font '28 Days Later' has not been found. Using (.*) font instead. Reason: alternative name from document.")
                        ->Match(substitutionWarningHandler->FontWarnings->idx_get(0)->get_Description())
                        ->get_Success());

        // The warning info collection can also be cleared like this
        substitutionWarningHandler->FontWarnings->Clear();

        ASSERT_EQ(0, substitutionWarningHandler->FontWarnings->get_Count());
    }

    class HandleDocumentSubstitutionWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> FontWarnings;

        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // We are only interested in fonts being substituted
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                FontWarnings->Warning(info);
            }
        }

        HandleDocumentSubstitutionWarnings() : FontWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };
    //ExEnd

    void DisableFontSubstitution()
    {
        auto doc = MakeObject<Document>(MyDir + u"Missing font.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        auto callback = MakeObject<ExFont::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(callback);

        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);

        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"Font.DisableFontSubstitution.pdf");

        auto reg = MakeObject<System::Text::RegularExpressions::Regex>(
            u"Font '28 Days Later' has not been found. Using (.*) font instead. Reason: default font setting.");

        for (auto fontWarning : System::IterateOver(callback->FontWarnings))
        {
            SharedPtr<System::Text::RegularExpressions::Match> match = reg->Match(fontWarning->get_Description());
            if (match->get_Success())
            {
                SUCCEED();
                return;
                break;
            }
        }
    }

    void SubstitutionWarnings()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        auto callback = MakeObject<ExFont::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(callback);

        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->SetFontsFolder(FontsDir, false);
        fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Arial", MakeArray<String>({u"Arvo", u"Slab"}));

        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"Font.SubstitutionWarnings.pdf");

        ASSERT_EQ(u"Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution.",
                  callback->FontWarnings->idx_get(0)->get_Description());
        ASSERT_EQ(u"Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution.",
                  callback->FontWarnings->idx_get(1)->get_Description());
    }

    void SubstitutionWarningsClosestMatch()
    {
        auto doc = MakeObject<Document>(MyDir + u"Bullet points with alternative font.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        auto callback = MakeObject<ExFont::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(callback);

        doc->Save(ArtifactsDir + u"Font.SubstitutionWarningsClosestMatch.pdf");

        ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_Description() ==
                    u"Font \'SymbolPS\' has not been found. Using \'Wingdings\' font instead. Reason: font info substitution.");
    }

    void SetFontAutoColor()
    {
        //ExStart
        //ExFor:Font.AutoColor
        //ExSummary:Shows how calculated color of the text (black or white) to be used for 'auto color'
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Remove direct color, so it can be calculated automatically with Font.AutoColor
        builder->get_Font()->set_Color(System::Drawing::Color::Empty);

        // When we set black color for background, autocolor for font must be white
        builder->get_Font()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Black());

        builder->Writeln(u"The text color automatically chosen for this run is white.");

        // When we set white color for background, autocolor for font must be black
        builder->get_Font()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_White());

        builder->Writeln(u"The text color automatically chosen for this run is black.");

        doc->Save(ArtifactsDir + u"Font.SetFontAutoColor.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Font.SetFontAutoColor.docx");
        SharedPtr<Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"The text color automatically chosen for this run is white.", run->GetText().Trim());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), run->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());

        run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);

        ASSERT_EQ(u"The text color automatically chosen for this run is black.", run->GetText().Trim());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), run->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());
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
    //ExSummary:Implements the Visitor Pattern to remove all content formatted as hidden from the document.
    void RemoveHiddenContentFromDocument()
    {
        // Open the document we want to remove hidden content from
        auto doc = MakeObject<Document>(MyDir + u"Hidden content.docx");
        ASSERT_EQ(26, doc->GetChildNodes(NodeType::Paragraph, true)->get_Count());
        //ExSkip
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        //ExSkip

        // Create an object that inherits from the DocumentVisitor class
        auto hiddenContentRemover = MakeObject<ExFont::RemoveHiddenContentVisitor>();

        // We can run it over the entire the document like so
        doc->Accept(hiddenContentRemover);

        // Or we can run it on only a specific node
        auto para = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 4, true));
        para->Accept(hiddenContentRemover);

        // Or over a different type of node like below
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        table->Accept(hiddenContentRemover);

        doc->Save(ArtifactsDir + u"Font.RemoveHiddenContentFromDocument.docx");
        TestRemoveHiddenContent(MakeObject<Document>(ArtifactsDir + u"Font.RemoveHiddenContentFromDocument.docx"));
        //ExSkip
    }

    /// <summary>
    /// This class when executed will remove all hidden content from the Document. Implemented as a Visitor.
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
            // Currently there is no way to tell if a particular Table/Row/Cell is hidden.
            // Instead, if the content of a table is hidden, then all inline child nodes of the table should be
            // hidden and thus removed by previous visits as well. This will result in the container being empty
            // If this is the case, we know to remove the table node.
            // Note that a table which is not hidden but simply has no content will not be affected by this algorithm,
            // as technically they are not completely empty (for example a properly formed Cell will have at least
            // an empty paragraph in it)
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

protected:
    void TestRemoveHiddenContent(SharedPtr<Document> doc)
    {
        ASSERT_EQ(20, doc->GetChildNodes(NodeType::Paragraph, true)->get_Count());
        //ExSkip
        ASSERT_EQ(1, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        //ExSkip

        for (auto node : System::IterateOver(doc->GetChildNodes(NodeType::Any, true)))
        {
            if (System::ObjectExt::Is<FieldStart>(node))
            {
                auto fieldStart = System::DynamicCast<FieldStart>(node);
                ASSERT_FALSE(fieldStart->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<FieldEnd>(node))
            {
                auto fieldEnd = System::DynamicCast<FieldEnd>(node);
                ASSERT_FALSE(fieldEnd->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<FieldSeparator>(node))
            {
                auto fieldSeparator = System::DynamicCast<FieldSeparator>(node);
                ASSERT_FALSE(fieldSeparator->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Run>(node))
            {
                auto run = System::DynamicCast<Run>(node);
                ASSERT_FALSE(run->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Paragraph>(node))
            {
                auto paragraph = System::DynamicCast<Paragraph>(node);
                ASSERT_FALSE(paragraph->get_ParagraphBreakFont()->get_Hidden());
            }
            else if (System::ObjectExt::Is<FormField>(node))
            {
                auto formField = System::DynamicCast<FormField>(node);
                ASSERT_FALSE(formField->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<GroupShape>(node))
            {
                auto groupShape = System::DynamicCast<GroupShape>(node);
                ASSERT_FALSE(groupShape->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Shape>(node))
            {
                auto shape = System::DynamicCast<Shape>(node);
                ASSERT_FALSE(shape->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Comment>(node))
            {
                auto comment = System::DynamicCast<Comment>(node);
                ASSERT_FALSE(comment->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<Footnote>(node))
            {
                auto footnote = System::DynamicCast<Footnote>(node);
                ASSERT_FALSE(footnote->get_Font()->get_Hidden());
            }
            else if (System::ObjectExt::Is<SpecialChar>(node))
            {
                auto specialChar = System::DynamicCast<SpecialChar>(node);
                ASSERT_FALSE(specialChar->get_Font()->get_Hidden());
            }
        }
    }

public:
    void BlankDocumentFonts()
    {
        //ExStart
        //ExFor:Fonts.FontInfoCollection.Contains(String)
        //ExFor:Fonts.FontInfoCollection.Count
        //ExSummary:Shows info about the fonts that are present in the blank document.
        auto doc = MakeObject<Document>();

        // A blank document comes with 3 fonts
        ASSERT_EQ(3, doc->get_FontInfos()->get_Count());

        // Their names can be looked up like this
        ASSERT_EQ(u"Times New Roman", doc->get_FontInfos()->idx_get(0)->get_Name());
        ASSERT_EQ(u"Symbol", doc->get_FontInfos()->idx_get(1)->get_Name());
        ASSERT_EQ(u"Arial", doc->get_FontInfos()->idx_get(2)->get_Name());
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
        //ExSummary:Shows how to extract embedded font from a document.
        auto doc = MakeObject<Document>(MyDir + u"Embedded font.docx");

        // Get the FontInfo for the embedded font
        SharedPtr<FontInfo> embeddedFont = doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift");

        // We can now extract this embedded font
        ArrayPtr<uint8_t> embeddedFontBytes = embeddedFont->GetEmbeddedFont(EmbeddedFontFormat::OpenType, EmbeddedFontStyle::Regular);
        ASSERT_FALSE(embeddedFontBytes == nullptr);

        // Then we can save the font to our directory
        System::IO::File::WriteAllBytes(ArtifactsDir + u"Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);

        // If we want to extract a font from a .doc as opposed to a .docx, we need to make sure to set the appropriate embedded font format
        doc = MakeObject<Document>(MyDir + u"Embedded font.doc");

        ASSERT_TRUE(System::TestTools::IsNull(
            doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFont(EmbeddedFontFormat::OpenType, EmbeddedFontStyle::Regular)));
        ASSERT_FALSE(System::TestTools::IsNull(
            doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFont(EmbeddedFontFormat::EmbeddedOpenType, EmbeddedFontStyle::Regular)));
        // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType
        ASSERT_FALSE(
            System::TestTools::IsNull(doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFontAsOpenType(EmbeddedFontStyle::Regular)));
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
        //ExSummary:Shows how to get information about each font in a document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // We can iterate over all the fonts with an enumerator
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<FontInfo>>> fontCollectionEnumerator =
            doc->get_FontInfos()->GetEnumerator();
        // Print detailed information about each font to the console
        while (fontCollectionEnumerator->MoveNext())
        {
            SharedPtr<FontInfo> fontInfo = fontCollectionEnumerator->get_Current();
            if (fontInfo != nullptr)
            {
                std::cout << (String(u"Font name: ") + fontInfo->get_Name()) << std::endl;
                // Alt names are usually blank
                std::cout << (String(u"Alt name: ") + fontInfo->get_AltName()) << std::endl;
                std::cout << (String(u"\t- Family: ") + System::ObjectExt::ToString(fontInfo->get_Family())) << std::endl;
                std::cout << (String(u"\t- ") + (fontInfo->get_IsTrueType() ? String(u"Is TrueType") : String(u"Is not TrueType")))
                          << std::endl;
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
        ASPOSE_ASSERT_EQ(MakeArray<int>({2, 2, 6, 3, 5, 4, 5, 2, 3, 4}), doc->get_FontInfos()->idx_get(u"Times New Roman")->get_Panose());
    }

    void FontSourceFile()
    {
        //ExStart
        //ExFor:Fonts.FileFontSource
        //ExFor:Fonts.FileFontSource.#ctor(String)
        //ExFor:Fonts.FileFontSource.#ctor(String, Int32)
        //ExFor:Fonts.FileFontSource.FilePath
        //ExFor:Fonts.FileFontSource.Type
        //ExFor:Fonts.FontSourceBase
        //ExFor:Fonts.FontSourceBase.Priority
        //ExFor:Fonts.FontSourceBase.Type
        //ExFor:Fonts.FontSourceType
        //ExSummary:Shows how to create a file font source.
        auto doc = MakeObject<Document>();

        // Create a font settings object for our document
        doc->set_FontSettings(MakeObject<FontSettings>());

        // Create a font source from a file in our system
        auto fileFontSource = MakeObject<FileFontSource>(MyDir + u"Alte DIN 1451 Mittelschrift.ttf", 0);

        // Import the font source into our document
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({fileFontSource}));

        ASSERT_EQ(MyDir + u"Alte DIN 1451 Mittelschrift.ttf", fileFontSource->get_FilePath());
        ASSERT_EQ(FontSourceType::FontFile, fileFontSource->get_Type());
        ASSERT_EQ(0, fileFontSource->get_Priority());
        //ExEnd
    }

    void FontSourceFolder()
    {
        //ExStart
        //ExFor:Fonts.FolderFontSource
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean)
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean, Int32)
        //ExFor:Fonts.FolderFontSource.FolderPath
        //ExFor:Fonts.FolderFontSource.ScanSubfolders
        //ExFor:Fonts.FolderFontSource.Type
        //ExSummary:Shows how to create a folder font source.
        auto doc = MakeObject<Document>();

        // Create a font settings object for our document
        doc->set_FontSettings(MakeObject<FontSettings>());

        // Create a font source from a folder that contains font files
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false, 1);

        // Add that source to our document
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({folderFontSource}));

        ASSERT_EQ(FontsDir, folderFontSource->get_FolderPath());
        ASPOSE_ASSERT_EQ(false, folderFontSource->get_ScanSubfolders());
        ASSERT_EQ(FontSourceType::FontsFolder, folderFontSource->get_Type());
        ASSERT_EQ(1, folderFontSource->get_Priority());
        //ExEnd
    }

    void FontSourceMemory()
    {
        //ExStart
        //ExFor:Fonts.MemoryFontSource
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[])
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[], Int32)
        //ExFor:Fonts.MemoryFontSource.FontData
        //ExFor:Fonts.MemoryFontSource.Type
        //ExSummary:Shows how to create a memory font source.
        auto doc = MakeObject<Document>();

        // Create a font settings object for our document
        doc->set_FontSettings(MakeObject<FontSettings>());

        // Import a font file, putting its contents into a byte array
        ArrayPtr<uint8_t> fontBytes = System::IO::File::ReadAllBytes(MyDir + u"Alte DIN 1451 Mittelschrift.ttf");

        // Create a memory font source from our array
        auto memoryFontSource = MakeObject<MemoryFontSource>(fontBytes, 0);

        // Add that font source to our document
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({memoryFontSource}));

        ASSERT_EQ(FontSourceType::MemoryFont, memoryFontSource->get_Type());
        ASSERT_EQ(0, memoryFontSource->get_Priority());
        //ExEnd
    }

    void FontSourceSystem()
    {
        //ExStart
        //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
        //ExFor:FontSubstitutionRule.Enabled
        //ExFor:TableSubstitutionRule.GetSubstitutes(String)
        //ExFor:Fonts.FontSettings.ResetFontSources
        //ExFor:Fonts.FontSettings.SubstitutionSettings
        //ExFor:Fonts.FontSubstitutionSettings
        //ExFor:Fonts.SystemFontSource
        //ExFor:Fonts.SystemFontSource.GetSystemFontFolders
        //ExFor:Fonts.SystemFontSource.Type
        //ExSummary:Shows how to access a document's system font source and set font substitutes.
        auto doc = MakeObject<Document>();

        // Create a font settings object for our document
        doc->set_FontSettings(MakeObject<FontSettings>());

        // By default, we always start with a system font source
        ASSERT_EQ(1, doc->get_FontSettings()->GetFontsSources()->get_Length());

        auto systemFontSource = System::DynamicCast<SystemFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0));
        ASSERT_EQ(FontSourceType::SystemFonts, systemFontSource->get_Type());
        ASSERT_EQ(0, systemFontSource->get_Priority());

        System::PlatformID pid = System::Environment::get_OSVersion().get_Platform();
        bool isWindows = (pid == System::PlatformID::Win32NT) || (pid == System::PlatformID::Win32S) || (pid == System::PlatformID::Win32Windows) ||
                         (pid == System::PlatformID::WinCE);
        if (isWindows)
        {
            const String fontsPath = u"C:\\WINDOWS\\Fonts";
            ASSERT_EQ(fontsPath.ToLower(), SystemFontSource::GetSystemFontFolders()->LINQ_FirstOrDefault().ToLower());
        }

        for (String systemFontFolder : SystemFontSource::GetSystemFontFolders())
        {
            std::cout << systemFontFolder << std::endl;
        }

        // Set a font that exists in the Windows Fonts directory as a substitute for one that doesn't
        doc->get_FontSettings()->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(true);
        doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Kreon-Regular",
                                                                                                     MakeArray<String>({u"Calibri"}));

        ASSERT_EQ(1, doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Kreon-Regular")->LINQ_Count());
        ASSERT_TRUE(doc->get_FontSettings()
                        ->get_SubstitutionSettings()
                        ->get_TableSubstitution()
                        ->GetSubstitutes(u"Kreon-Regular")
                        ->LINQ_ToArray()
                        ->Contains(u"Calibri"));

        // Alternatively, we could add a folder font source in which the corresponding folder contains the font
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false);
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({systemFontSource, folderFontSource}));
        ASSERT_EQ(2, doc->get_FontSettings()->GetFontsSources()->get_Length());

        // Resetting the font sources still leaves us with the system font source as well as our substitutes
        doc->get_FontSettings()->ResetFontSources();

        ASSERT_EQ(1, doc->get_FontSettings()->GetFontsSources()->get_Length());
        ASSERT_EQ(FontSourceType::SystemFonts, doc->get_FontSettings()->GetFontsSources()->idx_get(0)->get_Type());
        ASSERT_EQ(1, doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Kreon-Regular")->LINQ_Count());
        //ExEnd
    }

    void LoadFontFallbackSettingsFromFile()
    {
        //ExStart
        //ExFor:FontFallbackSettings.Load(String)
        //ExFor:FontFallbackSettings.Save(String)
        //ExSummary:Shows how to load and save font fallback settings from file.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // By default, fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->Load(MyDir + u"Font fallback rules.xml");

        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"Font.LoadFontFallbackSettingsFromFile.pdf");

        // Saves font fallback setting by string
        doc->get_FontSettings()->get_FallbackSettings()->Save(ArtifactsDir + u"FallbackSettings.xml");
        //ExEnd
    }

    void LoadFontFallbackSettingsFromStream()
    {
        //ExStart
        //ExFor:FontFallbackSettings.Load(Stream)
        //ExFor:FontFallbackSettings.Save(Stream)
        //ExSummary:Shows how to load and save font fallback settings from stream.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // By default, fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback
        {
            auto fontFallbackStream = MakeObject<System::IO::FileStream>(MyDir + u"Font fallback rules.xml", System::IO::FileMode::Open);
            auto fontSettings = MakeObject<FontSettings>();
            fontSettings->get_FallbackSettings()->Load(fontFallbackStream);

            doc->set_FontSettings(fontSettings);
        }

        doc->Save(ArtifactsDir + u"Font.LoadFontFallbackSettingsFromStream.pdf");

        // Saves font fallback setting by stream
        {
            auto fontFallbackStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"FallbackSettings.xml", System::IO::FileMode::Create);
            doc->get_FontSettings()->get_FallbackSettings()->Save(fontFallbackStream);
        }
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"FallbackSettings.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

        ASSERT_EQ(u"0B80-0BFF", rules->idx_get(0)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Vijaya", rules->idx_get(0)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"1F300-1F64F", rules->idx_get(1)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Segoe UI Emoji, Segoe UI Symbol", rules->idx_get(1)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"2000-206F, 2070-209F, 20B9", rules->idx_get(2)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Arial", rules->idx_get(2)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"3040-309F", rules->idx_get(3)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"MS Gothic", rules->idx_get(3)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
        ASSERT_EQ(u"Times New Roman", rules->idx_get(3)->get_Attributes()->idx_get(u"BaseFonts")->get_Value());

        ASSERT_EQ(u"3040-309F", rules->idx_get(4)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"MS Mincho", rules->idx_get(4)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"Arial Unicode MS", rules->idx_get(5)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    }

    void LoadNotoFontsFallbackSettings()
    {
        //ExStart
        //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
        auto fontSettings = MakeObject<FontSettings>();

        // These are free fonts licensed under SIL OFL
        // They can be downloaded from https://www.google.com/get/noto/#sans-lgc
        fontSettings->SetFontsFolder(FontsDir + u"Noto", false);

        // Note that only Sans style Noto fonts with regular weight are used in the predefined settings
        // Some of the Noto fonts uses advanced typography features
        // Advanced typography is currently not supported by AW and these fonts may be rendered inaccurately
        fontSettings->get_FallbackSettings()->LoadNotoFallbackSettings();
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Noto Sans");

        auto doc = MakeObject<Document>();
        doc->set_FontSettings(fontSettings);
        //ExEnd
    }

    void DefaultFontSubstitutionRule_()
    {
        //ExStart
        //ExFor:Fonts.DefaultFontSubstitutionRule
        //ExFor:Fonts.DefaultFontSubstitutionRule.DefaultFontName
        //ExFor:Fonts.FontSubstitutionSettings.DefaultFontSubstitution
        //ExSummary:Shows how to set the default font substitution rule.
        // Create a blank document and a new FontSettings property
        auto doc = MakeObject<Document>();
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);

        // Get the default substitution rule within FontSettings, which will be enabled by default and will substitute all missing fonts with "Times New Roman"
        SharedPtr<DefaultFontSubstitutionRule> defaultFontSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution();
        ASSERT_TRUE(defaultFontSubstitutionRule->get_Enabled());
        ASSERT_EQ(u"Times New Roman", defaultFontSubstitutionRule->get_DefaultFontName());

        // Set the default font substitute to "Courier New"
        defaultFontSubstitutionRule->set_DefaultFontName(u"Courier New");

        // Using a document builder, add some text in a font that we do not have to see the substitution take place,
        // and render the result in a PDF
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Missing Font");
        builder->Writeln(u"Line written in a missing font, which will be substituted with Courier New.");

        doc->Save(ArtifactsDir + u"Font.DefaultFontSubstitutionRule.pdf");
        //ExEnd

        ASSERT_EQ(u"Courier New", defaultFontSubstitutionRule->get_DefaultFontName());
    }

    void FontConfigSubstitution()
    {
        //ExStart
        //ExFor:Fonts.FontConfigSubstitutionRule
        //ExFor:Fonts.FontConfigSubstitutionRule.Enabled
        //ExFor:Fonts.FontConfigSubstitutionRule.IsFontConfigAvailable
        //ExFor:Fonts.FontConfigSubstitutionRule.ResetCache
        //ExFor:Fonts.FontSubstitutionRule
        //ExFor:Fonts.FontSubstitutionRule.Enabled
        //ExFor:Fonts.FontSubstitutionSettings.FontConfigSubstitution
        //ExSummary:Shows OS-dependent font config substitution.
        // Create a new FontSettings object and get its font config substitution rule
        auto fontSettings = MakeObject<FontSettings>();
        SharedPtr<FontConfigSubstitutionRule> fontConfigSubstitution = fontSettings->get_SubstitutionSettings()->get_FontConfigSubstitution();

        auto isOsPlatform = [](System::PlatformID p)
        {
            return System::Environment::get_OSVersion().get_Platform() == p;
        };

        bool isWindows = MakeArray<System::PlatformID>({System::PlatformID::Win32NT, System::PlatformID::Win32S, System::PlatformID::Win32Windows, System::PlatformID::WinCE})->LINQ_Any(isOsPlatform);

        // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms
        // On Windows, it is unavailable
        if (isWindows)
        {
            ASSERT_FALSE(fontConfigSubstitution->get_Enabled());
            ASSERT_FALSE(fontConfigSubstitution->IsFontConfigAvailable());
        }

        bool isLinuxOrMac = MakeArray<System::PlatformID>({System::PlatformID::Unix, System::PlatformID::MacOSX})->LINQ_Any(isOsPlatform);

        // On Linux/Mac, we will have access and will be able to perform operations
        if (isLinuxOrMac)
        {
            ASSERT_FALSE(fontConfigSubstitution->get_Enabled());
            ASSERT_TRUE(fontConfigSubstitution->IsFontConfigAvailable());

            fontConfigSubstitution->ResetCache();
        }
        //ExEnd
    }

    void FallbackSettings()
    {
        //ExStart
        //ExFor:Fonts.FontFallbackSettings.LoadMsOfficeFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to load pre-defined fallback font settings.
        auto doc = MakeObject<Document>();

        // Create a FontSettings object for our document and get its FallbackSettings attribute
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);
        SharedPtr<FontFallbackSettings> fontFallbackSettings = fontSettings->get_FallbackSettings();

        // Save the default fallback font scheme in an XML document
        // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts
        // This means that if the font we are using does not have symbols for the 0x0C00-0x0C7F Unicode block,
        // the symbols from the "Vani" font will be used as a substitute
        fontFallbackSettings->Save(ArtifactsDir + u"Font.FallbackSettings.Default.xml");

        // There are two pre-defined font fallback schemes we can choose from
        // 1: Use the default Microsoft Office scheme, which is the same one as the default
        fontFallbackSettings->LoadMsOfficeFallbackSettings();
        fontFallbackSettings->Save(ArtifactsDir + u"Font.FallbackSettings.LoadMsOfficeFallbackSettings.xml");

        // 2: Use the scheme built from Google Noto fonts
        fontFallbackSettings->LoadNotoFallbackSettings();
        fontFallbackSettings->Save(ArtifactsDir + u"Font.FallbackSettings.LoadNotoFallbackSettings.xml");
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"Font.FallbackSettings.Default.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

        ASSERT_EQ(u"0C00-0C7F", rules->idx_get(5)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Vani", rules->idx_get(5)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    }

    void FallbackSettingsCustom()
    {
        //ExStart
        //ExFor:Fonts.FontSettings.FallbackSettings
        //ExFor:Fonts.FontFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.BuildAutomatic
        //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
        auto doc = MakeObject<Document>();

        // Create a FontSettings object for our document and get its FallbackSettings attribute
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);
        SharedPtr<FontFallbackSettings> fontFallbackSettings = fontSettings->get_FallbackSettings();

        // Set our fonts to be sourced exclusively from the "MyFonts" folder
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false);
        fontSettings->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({folderFontSource}));

        // Calling BuildAutomatic() will generate a fallback scheme that distributes accessible fonts across as many Unicode character codes as possible
        // In our case, it only has access to the handful of fonts inside the "MyFonts" folder
        fontFallbackSettings->BuildAutomatic();
        fontFallbackSettings->Save(ArtifactsDir + u"Font.FallbackSettingsCustom.BuildAutomatic.xml");

        // We can also load a custom substitution scheme from a file like this
        // This scheme applies the "Arvo" font across the "0000-00ff" Unicode blocks, the "Squarish Sans CT" font across "0100-024f",
        // and the "M+ 2m" font in every place that none of the other fonts cover
        fontFallbackSettings->Load(MyDir + u"Custom font fallback settings.xml");

        // Create a document builder and set its font to one that does not exist in any of our sources
        // In doing that we will rely completely on our font fallback scheme to render text
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Name(u"Missing Font");

        // Type out every Unicode character from 0x0021 to 0x052F, with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme
        for (int i = 0x0021; i < 0x0530; i++)
        {
            switch (i)
            {
            case 0x0021:
                builder->Writeln(u"\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in \"Arvo\" font:");
                break;

            case 0x0100:
                builder->Writeln(u"\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"Squarish Sans CT\" font:");
                break;

            case 0x0250:
                builder->Writeln(u"\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in \"M+ 2m\" font:");
                break;
            }

            builder->Write(String::Format(u"{0}", System::Convert::ToChar(i)));
        }

        doc->Save(ArtifactsDir + u"Font.FallbackSettingsCustom.pdf");
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"Font.FallbackSettingsCustom.BuildAutomatic.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

        ASSERT_EQ(u"0000-007F", rules->idx_get(0)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Arvo", rules->idx_get(0)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"0180-024F", rules->idx_get(3)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"M+ 2m", rules->idx_get(3)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"0300-036F", rules->idx_get(6)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Noticia Text", rules->idx_get(6)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"0590-05FF", rules->idx_get(10)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Squarish Sans CT", rules->idx_get(10)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    }

    void TableSubstitutionRule_()
    {
        //ExStart
        //ExFor:Fonts.TableSubstitutionRule
        //ExFor:Fonts.TableSubstitutionRule.LoadLinuxSettings
        //ExFor:Fonts.TableSubstitutionRule.LoadWindowsSettings
        //ExFor:Fonts.TableSubstitutionRule.Save(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Save(System.String)
        //ExSummary:Shows how to access font substitution tables for Windows and Linux.
        // Create a blank document and a new FontSettings object
        auto doc = MakeObject<Document>();
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);

        // Create a new table substitution rule and load the default Windows font substitution table
        SharedPtr<TableSubstitutionRule> tableSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();
        tableSubstitutionRule->LoadWindowsSettings();

        // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman"
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Times New Roman"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman CE")->LINQ_ToArray());

        // We can save the table for viewing in the form of an XML document
        tableSubstitutionRule->Save(ArtifactsDir + u"Font.TableSubstitutionRule.Windows.xml");

        // Linux has its own substitution table
        // If "FreeSerif" is unavailable to substitute for "Times New Roman CE", we then look for "Liberation Serif", and so on
        tableSubstitutionRule->LoadLinuxSettings();
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"FreeSerif", u"Liberation Serif", u"DejaVu Serif"}),
                         tableSubstitutionRule->GetSubstitutes(u"Times New Roman CE")->LINQ_ToArray());

        // Save the Linux substitution table using a stream
        {
            auto fileStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"Font.TableSubstitutionRule.Linux.xml", System::IO::FileMode::Create);
            tableSubstitutionRule->Save(fileStream);
        }
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"Font.TableSubstitutionRule.Windows.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules =
            fallbackSettingsDoc->SelectNodes(u"//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

        ASSERT_EQ(u"Times New Roman CE", rules->idx_get(16)->get_Attributes()->idx_get(u"OriginalFont")->get_Value());
        ASSERT_EQ(u"Times New Roman", rules->idx_get(16)->get_Attributes()->idx_get(u"SubstituteFonts")->get_Value());

        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"Font.TableSubstitutionRule.Linux.xml"));
        rules = fallbackSettingsDoc->SelectNodes(u"//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

        ASSERT_EQ(u"Times New Roman CE", rules->idx_get(31)->get_Attributes()->idx_get(u"OriginalFont")->get_Value());
        ASSERT_EQ(u"FreeSerif, Liberation Serif, DejaVu Serif", rules->idx_get(31)->get_Attributes()->idx_get(u"SubstituteFonts")->get_Value());
    }

    void TableSubstitutionRuleCustom()
    {
        //ExStart
        //ExFor:Fonts.FontSubstitutionSettings.TableSubstitution
        //ExFor:Fonts.TableSubstitutionRule.AddSubstitutes(System.String,System.String[])
        //ExFor:Fonts.TableSubstitutionRule.GetSubstitutes(System.String)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.String)
        //ExFor:Fonts.TableSubstitutionRule.SetSubstitutes(System.String,System.String[])
        //ExSummary:Shows how to work with custom font substitution tables.
        // Create a blank document and a new FontSettings object
        auto doc = MakeObject<Document>();
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);

        // Create a new table substitution rule and load the default Windows font substitution table
        SharedPtr<TableSubstitutionRule> tableSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();

        // If we select fonts exclusively from our own folder, we will need a custom substitution table
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false);
        fontSettings->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({folderFontSource}));

        // There are two ways of loading a substitution table from a file in the local file system
        // 1: Loading from a stream
        {
            auto fileStream = MakeObject<System::IO::FileStream>(MyDir + u"Font substitution rules.xml", System::IO::FileMode::Open);
            tableSubstitutionRule->Load(fileStream);
        }

        // 2: Load directly from file
        tableSubstitutionRule->Load(MyDir + u"Font substitution rules.xml");

        // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font", which we don't have,
        // and then with "Kreon", found in the "MyFonts" folder
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Missing Font", u"Kreon"}), tableSubstitutionRule->GetSubstitutes(u"Arial")->LINQ_ToArray());

        // If we find this substitution table lacking, we can also expand it programmatically
        // In this case, we add an entry that substitutes "Times New Roman" with "Arvo"
        ASSERT_TRUE(tableSubstitutionRule->GetSubstitutes(u"Times New Roman") == nullptr);
        tableSubstitutionRule->AddSubstitutes(u"Times New Roman", MakeArray<String>({u"Arvo"}));
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Arvo"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());

        // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes()
        // In case "Arvo" is unavailable, our table will look for "M+ 2m"
        tableSubstitutionRule->AddSubstitutes(u"Times New Roman", MakeArray<String>({u"M+ 2m"}));
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Arvo", u"M+ 2m"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());

        // SetSubstitutes() can set a new list of substitute fonts for a font
        tableSubstitutionRule->SetSubstitutes(u"Times New Roman", MakeArray<String>({u"Squarish Sans CT", u"M+ 2m"}));
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Squarish Sans CT", u"M+ 2m"}),
                         tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());

        // TO demonstrate substitution, write text in fonts we have no access to and render the result in a PDF
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Text written in Arial, to be substituted by Kreon.");

        builder->get_Font()->set_Name(u"Times New Roman");
        builder->Writeln(u"Text written in Times New Roman, to be substituted by Squarish Sans CT.");

        doc->Save(ArtifactsDir + u"Font.TableSubstitutionRule.Custom.pdf");
        //ExEnd
    }

    void ResolveFontsBeforeLoadingDocument()
    {
        //ExStart
        //ExFor:LoadOptions.FontSettings
        //ExSummary:Shows how to designate font substitutes during loading.
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(MakeObject<FontSettings>());

        // Set a font substitution rule for a LoadOptions object that replaces a font that's not installed in our system with one that is
        SharedPtr<TableSubstitutionRule> substitutionRule = loadOptions->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution();
        substitutionRule->AddSubstitutes(u"MissingFont", MakeArray<String>({u"Comic Sans MS"}));

        // If we pass that object while loading a document, any text with the "MissingFont" font will change to "Comic Sans MS"
        auto doc = MakeObject<Document>(MyDir + u"Missing font.html", loadOptions);

        // At this point such text will still be in "MissingFont", and font substitution will be carried out once we save
        ASSERT_EQ(u"MissingFont", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());

        doc->Save(ArtifactsDir + u"Font.ResolveFontsBeforeLoadingDocument.pdf");
        //ExEnd
    }

    void LineSpacing()
    {
        //ExStart
        //ExFor:Font.LineSpacing
        //ExSummary:Shows how to get line spacing of current font (in points).
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set different fonts for the DocumentBuilder and verify their line spacing
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
        //ExSummary:Shows how to checks if particular Dml text effect is applied.
        auto doc = MakeObject<Document>(MyDir + u"DrawingML text effects.docx");

        SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();

        ASSERT_TRUE(runs->idx_get(0)->get_Font()->HasDmlEffect(TextDmlEffect::Shadow));
        ASSERT_TRUE(runs->idx_get(1)->get_Font()->HasDmlEffect(TextDmlEffect::Shadow));
        ASSERT_TRUE(runs->idx_get(2)->get_Font()->HasDmlEffect(TextDmlEffect::Reflection));
        ASSERT_TRUE(runs->idx_get(3)->get_Font()->HasDmlEffect(TextDmlEffect::Effect3D));
        ASSERT_TRUE(runs->idx_get(4)->get_Font()->HasDmlEffect(TextDmlEffect::Fill));
        //ExEnd
    }

    //ExStart
    //ExFor:StreamFontSource
    //ExFor:StreamFontSource.OpenFontDataStream
    //ExSummary:Shows how to allows to load fonts from stream.
    void StreamFontSourceFileRendering()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({MakeObject<ExFont::StreamFontSourceFile>()}));

        auto builder = MakeObject<DocumentBuilder>();
        builder->get_Document()->set_FontSettings(fontSettings);
        builder->get_Font()->set_Name(u"Kreon-Regular");
        builder->Writeln(u"Test aspose text when saving to PDF.");

        builder->get_Document()->Save(ArtifactsDir + u"Font.StreamFontSourceFileRendering.pdf");
    }

private:
    /// <summary>
    /// Load the font data only when it is required and not to store it in the memory for the "FontSettings" lifetime.
    /// </summary>
    class StreamFontSourceFile : public StreamFontSource
    {
    public:
        SharedPtr<System::IO::Stream> OpenFontDataStream() override
        {
            return System::IO::File::OpenRead(FontsDir + u"Kreon-Regular.ttf");
        }

    protected:
        virtual ~StreamFontSourceFile()
        {
        }
    };
    //ExEnd

public:
    void CheckScanUserFontsFolder()
    {
        // On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
        // or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user
        auto systemFontSource = MakeObject<SystemFontSource>();

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> x)> _local_func_3 =
            [](SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> x)
        {
            return x->get_FilePath().Contains(u"\\AppData\\Local\\Microsoft\\Windows\\Fonts");
        };

        ASSERT_FALSE(System::TestTools::IsNull(
            systemFontSource->GetAvailableFonts()->LINQ_FirstOrDefault(static_cast<System::Func<SharedPtr<PhysicalFontInfo>, bool>>(_local_func_3))))
            << "Fonts did not install to the user font folder";
    }

    void SetEmphasisMark(EmphasisMark emphasisMark)
    {
        //ExStart
        //ExFor:EmphasisMark
        //ExFor:Font.EmphasisMark
        //ExSummary:Shows how to add additional character rendered above/below the glyph-character.
        auto builder = MakeObject<DocumentBuilder>();

        // Possible types of emphasis mark:
        // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdemphasismark?view=word-pia
        builder->get_Font()->set_EmphasisMark(emphasisMark);

        builder->Write(u"Emphasis text");
        builder->Writeln();
        builder->get_Font()->ClearFormatting();
        builder->Write(u"Simple text");

        builder->get_Document()->Save(ArtifactsDir + u"Fonts.SetEmphasisMark.docx");
        //ExEnd
    }
};

} // namespace ApiExamples
