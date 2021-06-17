#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DropCapPosition.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Hyphenation.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/LineSpacingRule.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/OutlineLevel.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <system/enumerator_adapter.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;

namespace ApiExamples {

class ExParagraphFormat : public ApiExampleBase
{
public:
    void AsianTypographyProperties()
    {
        //ExStart
        //ExFor:ParagraphFormat.FarEastLineBreakControl
        //ExFor:ParagraphFormat.WordWrap
        //ExFor:ParagraphFormat.HangingPunctuation
        //ExSummary:Shows how to set special properties for Asian typography.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();
        format->set_FarEastLineBreakControl(true);
        format->set_WordWrap(false);
        format->set_HangingPunctuation(true);

        doc->Save(ArtifactsDir + u"ParagraphFormat.AsianTypographyProperties.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.AsianTypographyProperties.docx");
        format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();

        ASSERT_TRUE(format->get_FarEastLineBreakControl());
        ASSERT_FALSE(format->get_WordWrap());
        ASSERT_TRUE(format->get_HangingPunctuation());
    }

    void DropCap(DropCapPosition dropCapPosition)
    {
        //ExStart
        //ExFor:DropCapPosition
        //ExSummary:Shows how to create a drop cap.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert one paragraph with a large letter that the text in the second and third paragraphs begins with.
        builder->get_Font()->set_Size(54);
        builder->Writeln(u"L");

        builder->get_Font()->set_Size(18);
        builder->Writeln(String(u"orem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                         u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
        builder->Writeln(String(u"Ut enim ad minim veniam, quis nostrud exercitation ") + u"ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Currently, the second and third paragraphs will appear underneath the first.
        // We can convert the first paragraph as a drop cap for the other paragraphs via its "ParagraphFormat" object.
        // Set the "DropCapPosition" property to "DropCapPosition.Margin" to place the drop cap
        // outside the left-hand side page margin if our text is left-to-right.
        // Set the "DropCapPosition" property to "DropCapPosition.Normal" to place the drop cap within the page margins
        // and to wrap the rest of the text around it.
        // "DropCapPosition.None" is the default state for all paragraphs.
        SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();
        format->set_DropCapPosition(dropCapPosition);

        doc->Save(ArtifactsDir + u"ParagraphFormat.DropCap.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.DropCap.docx");

        ASSERT_EQ(dropCapPosition, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_DropCapPosition());
        ASSERT_EQ(DropCapPosition::None, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_DropCapPosition());
    }

    void LineSpacing()
    {
        //ExStart
        //ExFor:ParagraphFormat.LineSpacing
        //ExFor:ParagraphFormat.LineSpacingRule
        //ExSummary:Shows how to work with line spacing.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are three line spacing rules that we can define using the
        // paragraph's "LineSpacingRule" property to configure spacing between paragraphs.
        // 1 -  Set a minimum amount of spacing.
        // This will give vertical padding to lines of text of any size
        // that is too small to maintain the minimum line-height.
        builder->get_ParagraphFormat()->set_LineSpacingRule(LineSpacingRule::AtLeast);
        builder->get_ParagraphFormat()->set_LineSpacing(20);

        builder->Writeln(u"Minimum line spacing of 20.");
        builder->Writeln(u"Minimum line spacing of 20.");

        // 2 -  Set exact spacing.
        // Using font sizes that are too large for the spacing will truncate the text.
        builder->get_ParagraphFormat()->set_LineSpacingRule(LineSpacingRule::Exactly);
        builder->get_ParagraphFormat()->set_LineSpacing(5);

        builder->Writeln(u"Line spacing of exactly 5.");
        builder->Writeln(u"Line spacing of exactly 5.");

        // 3 -  Set spacing as a multiple of default line spacing, which is 12 points by default.
        // This kind of spacing will scale to different font sizes.
        builder->get_ParagraphFormat()->set_LineSpacingRule(LineSpacingRule::Multiple);
        builder->get_ParagraphFormat()->set_LineSpacing(18);

        builder->Writeln(u"Line spacing of 1.5 default lines.");
        builder->Writeln(u"Line spacing of 1.5 default lines.");

        doc->Save(ArtifactsDir + u"ParagraphFormat.LineSpacing.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.LineSpacing.docx");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(LineSpacingRule::AtLeast, paragraphs->idx_get(0)->get_ParagraphFormat()->get_LineSpacingRule());
        ASPOSE_ASSERT_EQ(20.0, paragraphs->idx_get(0)->get_ParagraphFormat()->get_LineSpacing());
        ASSERT_EQ(LineSpacingRule::AtLeast, paragraphs->idx_get(1)->get_ParagraphFormat()->get_LineSpacingRule());
        ASPOSE_ASSERT_EQ(20.0, paragraphs->idx_get(1)->get_ParagraphFormat()->get_LineSpacing());

        ASSERT_EQ(LineSpacingRule::Exactly, paragraphs->idx_get(2)->get_ParagraphFormat()->get_LineSpacingRule());
        ASPOSE_ASSERT_EQ(5.0, paragraphs->idx_get(2)->get_ParagraphFormat()->get_LineSpacing());
        ASSERT_EQ(LineSpacingRule::Exactly, paragraphs->idx_get(3)->get_ParagraphFormat()->get_LineSpacingRule());
        ASPOSE_ASSERT_EQ(5.0, paragraphs->idx_get(3)->get_ParagraphFormat()->get_LineSpacing());

        ASSERT_EQ(LineSpacingRule::Multiple, paragraphs->idx_get(4)->get_ParagraphFormat()->get_LineSpacingRule());
        ASPOSE_ASSERT_EQ(18.0, paragraphs->idx_get(4)->get_ParagraphFormat()->get_LineSpacing());
        ASSERT_EQ(LineSpacingRule::Multiple, paragraphs->idx_get(5)->get_ParagraphFormat()->get_LineSpacingRule());
        ASPOSE_ASSERT_EQ(18.0, paragraphs->idx_get(5)->get_ParagraphFormat()->get_LineSpacing());
    }

    void ParagraphSpacingAuto(bool autoSpacing)
    {
        //ExStart
        //ExFor:ParagraphFormat.SpaceAfter
        //ExFor:ParagraphFormat.SpaceAfterAuto
        //ExFor:ParagraphFormat.SpaceBefore
        //ExFor:ParagraphFormat.SpaceBeforeAuto
        //ExSummary:Shows how to set automatic paragraph spacing.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Apply a large amount of spacing before and after paragraphs that this builder will create.
        builder->get_ParagraphFormat()->set_SpaceBefore(24);
        builder->get_ParagraphFormat()->set_SpaceAfter(24);

        // Set these flags to "true" to apply automatic spacing,
        // effectively ignoring the spacing in the properties we set above.
        // Leave them as "false" will apply our custom paragraph spacing.
        builder->get_ParagraphFormat()->set_SpaceAfterAuto(autoSpacing);
        builder->get_ParagraphFormat()->set_SpaceBeforeAuto(autoSpacing);

        // Insert two paragraphs that will have spacing above and below them and save the document.
        builder->Writeln(u"Paragraph 1.");
        builder->Writeln(u"Paragraph 2.");

        doc->Save(ArtifactsDir + u"ParagraphFormat.ParagraphSpacingAuto.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.ParagraphSpacingAuto.docx");
        SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat();

        ASPOSE_ASSERT_EQ(24.0, format->get_SpaceBefore());
        ASPOSE_ASSERT_EQ(24.0, format->get_SpaceAfter());
        ASPOSE_ASSERT_EQ(autoSpacing, format->get_SpaceAfterAuto());
        ASPOSE_ASSERT_EQ(autoSpacing, format->get_SpaceBeforeAuto());

        format = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat();

        ASPOSE_ASSERT_EQ(24.0, format->get_SpaceBefore());
        ASPOSE_ASSERT_EQ(24.0, format->get_SpaceAfter());
        ASPOSE_ASSERT_EQ(autoSpacing, format->get_SpaceAfterAuto());
        ASPOSE_ASSERT_EQ(autoSpacing, format->get_SpaceBeforeAuto());
    }

    void ParagraphSpacingSameStyle(bool noSpaceBetweenParagraphsOfSameStyle)
    {
        //ExStart
        //ExFor:ParagraphFormat.SpaceAfter
        //ExFor:ParagraphFormat.SpaceBefore
        //ExFor:ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle
        //ExSummary:Shows how to apply no spacing between paragraphs with the same style.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Apply a large amount of spacing before and after paragraphs that this builder will create.
        builder->get_ParagraphFormat()->set_SpaceBefore(24);
        builder->get_ParagraphFormat()->set_SpaceAfter(24);

        // Set the "NoSpaceBetweenParagraphsOfSameStyle" flag to "true" to apply
        // no spacing between paragraphs with the same style, which will group similar paragraphs.
        // Leave the "NoSpaceBetweenParagraphsOfSameStyle" flag as "false"
        // to evenly apply spacing to every paragraph.
        builder->get_ParagraphFormat()->set_NoSpaceBetweenParagraphsOfSameStyle(noSpaceBetweenParagraphsOfSameStyle);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
        builder->Writeln(String::Format(u"Paragraph in the \"{0}\" style.", builder->get_ParagraphFormat()->get_Style()->get_Name()));
        builder->Writeln(String::Format(u"Paragraph in the \"{0}\" style.", builder->get_ParagraphFormat()->get_Style()->get_Name()));
        builder->Writeln(String::Format(u"Paragraph in the \"{0}\" style.", builder->get_ParagraphFormat()->get_Style()->get_Name()));
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));
        builder->Writeln(String::Format(u"Paragraph in the \"{0}\" style.", builder->get_ParagraphFormat()->get_Style()->get_Name()));
        builder->Writeln(String::Format(u"Paragraph in the \"{0}\" style.", builder->get_ParagraphFormat()->get_Style()->get_Name()));
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
        builder->Writeln(String::Format(u"Paragraph in the \"{0}\" style.", builder->get_ParagraphFormat()->get_Style()->get_Name()));
        builder->Writeln(String::Format(u"Paragraph in the \"{0}\" style.", builder->get_ParagraphFormat()->get_Style()->get_Name()));

        doc->Save(ArtifactsDir + u"ParagraphFormat.ParagraphSpacingSameStyle.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.ParagraphSpacingSameStyle.docx");

        for (const auto& paragraph : System::IterateOver<Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
        {
            SharedPtr<ParagraphFormat> format = paragraph->get_ParagraphFormat();

            ASPOSE_ASSERT_EQ(24.0, format->get_SpaceBefore());
            ASPOSE_ASSERT_EQ(24.0, format->get_SpaceAfter());
            ASPOSE_ASSERT_EQ(noSpaceBetweenParagraphsOfSameStyle, format->get_NoSpaceBetweenParagraphsOfSameStyle());
        }
    }

    void ParagraphOutlineLevel()
    {
        //ExStart
        //ExFor:ParagraphFormat.OutlineLevel
        //ExSummary:Shows how to configure paragraph outline levels to create collapsible text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Each paragraph has an OutlineLevel, which could be any number from 1 to 9, or at the default "BodyText" value.
        // Setting the property to one of the numbered values will show an arrow to the left
        // of the beginning of the paragraph.
        builder->get_ParagraphFormat()->set_OutlineLevel(OutlineLevel::Level1);
        builder->Writeln(u"Paragraph outline level 1.");

        // Level 1 is the topmost level. If there is a paragraph with a lower level below a paragraph with a higher level,
        // collapsing the higher-level paragraph will collapse the lower level paragraph.
        builder->get_ParagraphFormat()->set_OutlineLevel(OutlineLevel::Level2);
        builder->Writeln(u"Paragraph outline level 2.");

        // Two paragraphs of the same level will not collapse each other,
        // and the arrows do not collapse the paragraphs they point to.
        builder->get_ParagraphFormat()->set_OutlineLevel(OutlineLevel::Level3);
        builder->Writeln(u"Paragraph outline level 3.");
        builder->Writeln(u"Paragraph outline level 3.");

        // The default "BodyText" value is the lowest, which a paragraph of any level can collapse.
        builder->get_ParagraphFormat()->set_OutlineLevel(OutlineLevel::BodyText);
        builder->Writeln(u"Paragraph at main text level.");

        doc->Save(ArtifactsDir + u"ParagraphFormat.ParagraphOutlineLevel.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.ParagraphOutlineLevel.docx");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(OutlineLevel::Level1, paragraphs->idx_get(0)->get_ParagraphFormat()->get_OutlineLevel());
        ASSERT_EQ(OutlineLevel::Level2, paragraphs->idx_get(1)->get_ParagraphFormat()->get_OutlineLevel());
        ASSERT_EQ(OutlineLevel::Level3, paragraphs->idx_get(2)->get_ParagraphFormat()->get_OutlineLevel());
        ASSERT_EQ(OutlineLevel::Level3, paragraphs->idx_get(3)->get_ParagraphFormat()->get_OutlineLevel());
        ASSERT_EQ(OutlineLevel::BodyText, paragraphs->idx_get(4)->get_ParagraphFormat()->get_OutlineLevel());
    }

    void PageBreakBefore(bool pageBreakBefore)
    {
        //ExStart
        //ExFor:ParagraphFormat.PageBreakBefore
        //ExSummary:Shows how to create paragraphs with page breaks at the beginning.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set this flag to "true" to apply a page break to each paragraph's beginning
        // that the document builder will create under this ParagraphFormat configuration.
        // The first paragraph will not receive a page break.
        // Leave this flag as "false" to start each new paragraph on the same page
        // as the previous, provided there is sufficient space.
        builder->get_ParagraphFormat()->set_PageBreakBefore(pageBreakBefore);

        builder->Writeln(u"Paragraph 1.");
        builder->Writeln(u"Paragraph 2.");

        auto layoutCollector = MakeObject<LayoutCollector>(doc);
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        if (pageBreakBefore)
        {
            ASSERT_EQ(1, layoutCollector->GetStartPageIndex(paragraphs->idx_get(0)));
            ASSERT_EQ(2, layoutCollector->GetStartPageIndex(paragraphs->idx_get(1)));
        }
        else
        {
            ASSERT_EQ(1, layoutCollector->GetStartPageIndex(paragraphs->idx_get(0)));
            ASSERT_EQ(1, layoutCollector->GetStartPageIndex(paragraphs->idx_get(1)));
        }

        doc->Save(ArtifactsDir + u"ParagraphFormat.PageBreakBefore.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.PageBreakBefore.docx");
        paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASPOSE_ASSERT_EQ(pageBreakBefore, paragraphs->idx_get(0)->get_ParagraphFormat()->get_PageBreakBefore());
        ASPOSE_ASSERT_EQ(pageBreakBefore, paragraphs->idx_get(1)->get_ParagraphFormat()->get_PageBreakBefore());
    }

    void WidowControl(bool widowControl)
    {
        //ExStart
        //ExFor:ParagraphFormat.WidowControl
        //ExSummary:Shows how to enable widow/orphan control for a paragraph.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // When we write the text that does not fit onto one page, one line may spill over onto the next page.
        // The single line that ends up on the next page is called an "Orphan",
        // and the previous line where the orphan broke off is called a "Widow".
        // We can fix orphans and widows by rearranging text via font size, spacing, or page margins.
        // If we wish to preserve our document's dimensions, we can set this flag to "true"
        // to push widows onto the same page as their respective orphans.
        // Leave this flag as "false" will leave widow/orphan pairs in text.
        // Every paragraph has this setting accessible in Microsoft Word via Home -> Paragraph -> Paragraph Settings
        // (button on bottom right hand corner of "Paragraph" tab) -> "Widow/Orphan control".
        builder->get_ParagraphFormat()->set_WidowControl(widowControl);

        // Insert text that produces an orphan and a widow.
        builder->get_Font()->set_Size(68);
        builder->Write(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                       u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc->Save(ArtifactsDir + u"ParagraphFormat.WidowControl.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.WidowControl.docx");

        ASPOSE_ASSERT_EQ(widowControl, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_WidowControl());
    }

    void LinesToDrop()
    {
        //ExStart
        //ExFor:ParagraphFormat.LinesToDrop
        //ExSummary:Shows how to set the size of a drop cap.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Modify the "LinesToDrop" property to designate a paragraph as a drop cap,
        // which will turn it into a large capital letter that will decorate the next paragraph.
        // Give this property a value of 4 to give the drop cap the height of four text lines.
        builder->get_ParagraphFormat()->set_LinesToDrop(4);
        builder->Writeln(u"H");

        // Reset the "LinesToDrop" property to 0 to turn the next paragraph into an ordinary paragraph.
        // The text in this paragraph will wrap around the drop cap.
        builder->get_ParagraphFormat()->set_LinesToDrop(0);
        builder->Writeln(u"ello world!");

        doc->Save(ArtifactsDir + u"ParagraphFormat.LinesToDrop.odt");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ParagraphFormat.LinesToDrop.odt");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(4, paragraphs->idx_get(0)->get_ParagraphFormat()->get_LinesToDrop());
        ASSERT_EQ(0, paragraphs->idx_get(1)->get_ParagraphFormat()->get_LinesToDrop());
    }

    void SuppressHyphens(bool suppressAutoHyphens)
    {
        //ExStart
        //ExFor:ParagraphFormat.SuppressAutoHyphens
        //ExSummary:Shows how to suppress hyphenation for a paragraph.
        Hyphenation::RegisterDictionary(u"de-CH", MyDir + u"hyph_de_CH.dic");

        ASSERT_TRUE(Hyphenation::IsDictionaryRegistered(u"de-CH"));

        // Open a document containing text with a locale matching that of our dictionary.
        // When we save this document to a fixed page save format, its text will have hyphenation.
        auto doc = MakeObject<Document>(MyDir + u"German text.docx");

        // We can set the "SuppressAutoHyphens" property to "true" to disable hyphenation
        // for a specific paragraph while keeping it enabled for the rest of the document.
        // The default value for this property is "false",
        // which means every paragraph by default uses hyphenation if any is available.
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->set_SuppressAutoHyphens(suppressAutoHyphens);

        doc->Save(ArtifactsDir + u"ParagraphFormat.SuppressHyphens.pdf");
        //ExEnd
    }

    void ParagraphSpacingAndIndents()
    {
        //ExStart
        //ExFor:ParagraphFormat.CharacterUnitLeftIndent
        //ExFor:ParagraphFormat.CharacterUnitRightIndent
        //ExFor:ParagraphFormat.CharacterUnitFirstLineIndent
        //ExFor:ParagraphFormat.LineUnitBefore
        //ExFor:ParagraphFormat.LineUnitAfter
        //ExSummary:Shows how to change paragraph spacing and indents.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<ParagraphFormat> format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();

        // Below are five different spacing options, along with the properties that their configuration indirectly affects.
        // 1 -  Left indent:
        ASPOSE_ASSERT_EQ(format->get_LeftIndent(), 0.0);

        format->set_CharacterUnitLeftIndent(10.0);

        ASPOSE_ASSERT_EQ(format->get_LeftIndent(), 120.0);

        // 2 -  Right indent:
        ASPOSE_ASSERT_EQ(format->get_RightIndent(), 0.0);

        format->set_CharacterUnitRightIndent(-5.5);

        ASPOSE_ASSERT_EQ(format->get_RightIndent(), -66.0);

        // 3 -  Hanging indent:
        ASPOSE_ASSERT_EQ(format->get_FirstLineIndent(), 0.0);

        format->set_CharacterUnitFirstLineIndent(20.3);

        ASSERT_NEAR(format->get_FirstLineIndent(), 243.59, 0.1);

        // 4 -  Line spacing before paragraphs:
        ASPOSE_ASSERT_EQ(format->get_SpaceBefore(), 0.0);

        format->set_LineUnitBefore(5.1);

        ASSERT_NEAR(format->get_SpaceBefore(), 61.1, 0.1);

        // 5 -  Line spacing after paragraphs:
        ASPOSE_ASSERT_EQ(format->get_SpaceAfter(), 0.0);

        format->set_LineUnitAfter(10.9);

        ASSERT_NEAR(format->get_SpaceAfter(), 130.8, 0.1);

        builder->Writeln(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                         u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder->Write(String(u"测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试") +
                       u"文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档");
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        format = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat();

        ASPOSE_ASSERT_EQ(format->get_CharacterUnitLeftIndent(), 10.0);
        ASPOSE_ASSERT_EQ(format->get_LeftIndent(), 120.0);

        ASPOSE_ASSERT_EQ(format->get_CharacterUnitRightIndent(), -5.5);
        ASPOSE_ASSERT_EQ(format->get_RightIndent(), -66.0);

        ASPOSE_ASSERT_EQ(format->get_CharacterUnitFirstLineIndent(), 20.3);
        ASSERT_NEAR(format->get_FirstLineIndent(), 243.59, 0.1);

        ASSERT_NEAR(format->get_LineUnitBefore(), 5.1, 0.1);
        ASSERT_NEAR(format->get_SpaceBefore(), 61.1, 0.1);

        ASPOSE_ASSERT_EQ(format->get_LineUnitAfter(), 10.9);
        ASSERT_NEAR(format->get_SpaceAfter(), 130.8, 0.1);
    }
};

} // namespace ApiExamples
