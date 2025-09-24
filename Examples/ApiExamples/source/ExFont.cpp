// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFont.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/search_option.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/globalization/culture_info.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/enum.h>
#include <system/convert.h>
#include <system/collections/ilist.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Themes/ThemeFonts.h>
#include <Aspose.Words.Cpp/Model/Themes/ThemeFont.h>
#include <Aspose.Words.Cpp/Model/Themes/ThemeColor.h>
#include <Aspose.Words.Cpp/Model/Themes/Theme.h>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/TextEffect.h>
#include <Aspose.Words.Cpp/Model/Text/TextDmlEffect.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/NumSpacing.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Fonts/PhysicalFontInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontPitch.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontFamily.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontEmbeddingUsagePermissions.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontEmbeddingLicensingRights.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/EmbeddedFontStyle.h>
#include <Aspose.Words.Cpp/Model/Fonts/EmbeddedFontFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Borders/TextureIndex.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Themes;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1801053122u, ::Aspose::Words::ApiExamples::ExFont::RemoveHiddenContentVisitor, ThisTypeBaseTypesInfo);

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitFieldStart(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart)
{
    if (fieldStart->get_Font()->get_Hidden())
    {
        fieldStart->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd)
{
    if (fieldEnd->get_Font()->get_Hidden())
    {
        fieldEnd->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitFieldSeparator(System::SharedPtr<Aspose::Words::Fields::FieldSeparator> fieldSeparator)
{
    if (fieldSeparator->get_Font()->get_Hidden())
    {
        fieldSeparator->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (run->get_Font()->get_Hidden())
    {
        run->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitParagraphStart(System::SharedPtr<Aspose::Words::Paragraph> paragraph)
{
    if (paragraph->get_ParagraphBreakFont()->get_Hidden())
    {
        paragraph->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitFormField(System::SharedPtr<Aspose::Words::Fields::FormField> formField)
{
    if (formField->get_Font()->get_Hidden())
    {
        formField->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    if (groupShape->get_Font()->get_Hidden())
    {
        groupShape->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    if (shape->get_Font()->get_Hidden())
    {
        shape->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitCommentStart(System::SharedPtr<Aspose::Words::Comment> comment)
{
    if (comment->get_Font()->get_Hidden())
    {
        comment->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitFootnoteStart(System::SharedPtr<Aspose::Words::Notes::Footnote> footnote)
{
    if (footnote->get_Font()->get_Hidden())
    {
        footnote->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitSpecialChar(System::SharedPtr<Aspose::Words::SpecialChar> specialChar)
{
    std::cout << specialChar->GetText() << std::endl;
    
    if (specialChar->get_Font()->get_Hidden())
    {
        specialChar->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitTableEnd(System::SharedPtr<Aspose::Words::Tables::Table> table)
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
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitCellEnd(System::SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    if (!cell->get_HasChildNodes() && cell->get_ParentNode() != nullptr)
    {
        cell->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExFont::RemoveHiddenContentVisitor::VisitRowEnd(System::SharedPtr<Aspose::Words::Tables::Row> row)
{
    if (!row->get_HasChildNodes() && row->get_ParentNode() != nullptr)
    {
        row->Remove();
    }
    
    return Aspose::Words::VisitorAction::Continue;
}


RTTI_INFO_IMPL_HASH(2650332447u, ::Aspose::Words::ApiExamples::ExFont, ThisTypeBaseTypesInfo);

void ExFont::TestRemoveHiddenContent(System::SharedPtr<Aspose::Words::Document> doc)
{
    ASSERT_EQ(20, doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());
    //ExSkip
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    //ExSkip
    
    for (auto&& node : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Any, true)))
    {
        if (System::ObjectExt::Is<Aspose::Words::Fields::FieldStart>(node))
        {
            auto fieldStart = System::ExplicitCast<Aspose::Words::Fields::FieldStart>(node);
            ASSERT_FALSE(fieldStart->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Fields::FieldEnd>(node))
        {
            auto fieldEnd = System::ExplicitCast<Aspose::Words::Fields::FieldEnd>(node);
            ASSERT_FALSE(fieldEnd->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Fields::FieldSeparator>(node))
        {
            auto fieldSeparator = System::ExplicitCast<Aspose::Words::Fields::FieldSeparator>(node);
            ASSERT_FALSE(fieldSeparator->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Run>(node))
        {
            auto run = System::ExplicitCast<Aspose::Words::Run>(node);
            ASSERT_FALSE(run->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Paragraph>(node))
        {
            auto paragraph = System::ExplicitCast<Aspose::Words::Paragraph>(node);
            ASSERT_FALSE(paragraph->get_ParagraphBreakFont()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Fields::FormField>(node))
        {
            auto formField = System::ExplicitCast<Aspose::Words::Fields::FormField>(node);
            ASSERT_FALSE(formField->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Drawing::GroupShape>(node))
        {
            auto groupShape = System::ExplicitCast<Aspose::Words::Drawing::GroupShape>(node);
            ASSERT_FALSE(groupShape->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Drawing::Shape>(node))
        {
            auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(node);
            ASSERT_FALSE(shape->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Comment>(node))
        {
            auto comment = System::ExplicitCast<Aspose::Words::Comment>(node);
            ASSERT_FALSE(comment->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::Notes::Footnote>(node))
        {
            auto footnote = System::ExplicitCast<Aspose::Words::Notes::Footnote>(node);
            ASSERT_FALSE(footnote->get_Font()->get_Hidden());
        }
        else if (System::ObjectExt::Is<Aspose::Words::SpecialChar>(node))
        {
            auto specialChar = System::ExplicitCast<Aspose::Words::SpecialChar>(node);
            ASSERT_FALSE(specialChar->get_Font()->get_Hidden());
        }
    }
}


namespace gtest_test
{

class ExFont : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExFont> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExFont>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExFont> ExFont::s_instance;

} // namespace gtest_test

void ExFont::CreateFormattedRun()
{
    //ExStart
    //ExFor:Document.#ctor
    //ExFor:Font
    //ExFor:Font.Name
    //ExFor:Font.Size
    //ExFor:Font.HighlightColor
    //ExFor:Run
    //ExFor:Run.#ctor(DocumentBase,String)
    //ExFor:Story.FirstParagraph
    //ExSummary:Shows how to format a run of text using its font property.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto run = System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!");
    
    System::SharedPtr<Aspose::Words::Font> font = run->get_Font();
    font->set_Name(u"Courier New");
    font->set_Size(36);
    font->set_HighlightColor(System::Drawing::Color::get_Yellow());
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    doc->Save(get_ArtifactsDir() + u"Font.CreateFormattedRun.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.CreateFormattedRun.docx");
    run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Hello world!", run->GetText().Trim());
    ASSERT_EQ(u"Courier New", run->get_Font()->get_Name());
    ASPOSE_ASSERT_EQ(36, run->get_Font()->get_Size());
    ASSERT_EQ(System::Drawing::Color::get_Yellow().ToArgb(), run->get_Font()->get_HighlightColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExFont, CreateFormattedRun)
{
    s_instance->CreateFormattedRun();
}

} // namespace gtest_test

void ExFont::Caps()
{
    //ExStart
    //ExFor:Font.AllCaps
    //ExFor:Font.SmallCaps
    //ExSummary:Shows how to format a run to display its contents in capitals.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
    
    // There are two ways of getting a run to display its lowercase text in uppercase without changing the contents.
    // 1 -  Set the AllCaps flag to display all characters in regular capitals:
    auto run = System::MakeObject<Aspose::Words::Run>(doc, u"all capitals");
    run->get_Font()->set_AllCaps(true);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    para = System::ExplicitCast<Aspose::Words::Paragraph>(para->get_ParentNode()->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc)));
    
    // 2 -  Set the SmallCaps flag to display all characters in small capitals:
    // If a character is lower case, it will appear in its upper case form
    // but will have the same height as the lower case (the font's x-height).
    // Characters that were in upper case originally will look the same.
    run = System::MakeObject<Aspose::Words::Run>(doc, u"Small Capitals");
    run->get_Font()->set_SmallCaps(true);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    doc->Save(get_ArtifactsDir() + u"Font.Caps.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Caps.docx");
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"all capitals", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_AllCaps());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Small Capitals", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_SmallCaps());
}

namespace gtest_test
{

TEST_F(ExFont, Caps)
{
    s_instance->Caps();
}

} // namespace gtest_test

void ExFont::GetDocumentFonts()
{
    //ExStart
    //ExFor:FontInfoCollection
    //ExFor:DocumentBase.FontInfos
    //ExFor:FontInfo
    //ExFor:FontInfo.Name
    //ExFor:FontInfo.IsTrueType
    //ExSummary:Shows how to print the details of what fonts are present in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Embedded font.docx");
    
    System::SharedPtr<Aspose::Words::Fonts::FontInfoCollection> allFonts = doc->get_FontInfos();
    ASSERT_EQ(5, allFonts->get_Count());
    //ExSkip
    
    // Print all the used and unused fonts in the document.
    for (int32_t i = 0; i < allFonts->get_Count(); i++)
    {
        std::cout << System::String::Format(u"Font index #{0}", i) << std::endl;
        std::cout << System::String::Format(u"\tName: {0}", allFonts->idx_get(i)->get_Name()) << std::endl;
        std::cout << System::String::Format(u"\tIs {0}a trueType font", (allFonts->idx_get(i)->get_IsTrueType() ? System::String(u"") : System::String(u"not "))) << std::endl;
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFont, GetDocumentFonts)
{
    s_instance->GetDocumentFonts();
}

} // namespace gtest_test

void ExFont::DefaultValuesEmbeddedFontsParameters()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_FALSE(doc->get_FontInfos()->get_EmbedTrueTypeFonts());
    ASSERT_FALSE(doc->get_FontInfos()->get_EmbedSystemFonts());
    ASSERT_FALSE(doc->get_FontInfos()->get_SaveSubsetFonts());
}

namespace gtest_test
{

TEST_F(ExFont, DefaultValuesEmbeddedFontsParameters)
{
    s_instance->DefaultValuesEmbeddedFontsParameters();
}

} // namespace gtest_test

void ExFont::FontInfoCollection(bool embedAllFonts)
{
    //ExStart
    //ExFor:FontInfoCollection
    //ExFor:DocumentBase.FontInfos
    //ExFor:FontInfoCollection.EmbedTrueTypeFonts
    //ExFor:FontInfoCollection.EmbedSystemFonts
    //ExFor:FontInfoCollection.SaveSubsetFonts
    //ExSummary:Shows how to save a document with embedded TrueType fonts.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::SharedPtr<Aspose::Words::Fonts::FontInfoCollection> fontInfos = doc->get_FontInfos();
    fontInfos->set_EmbedTrueTypeFonts(embedAllFonts);
    fontInfos->set_EmbedSystemFonts(embedAllFonts);
    fontInfos->set_SaveSubsetFonts(embedAllFonts);
    
    doc->Save(get_ArtifactsDir() + u"Font.FontInfoCollection.docx");
    //ExEnd
    
    int64_t testedFileLength = System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"Font.FontInfoCollection.docx")->get_Length();
    
    if (embedAllFonts)
    {
        ASSERT_TRUE(testedFileLength < 28000);
    }
    else
    {
        ASSERT_TRUE(testedFileLength < 13000);
    }
}

namespace gtest_test
{

using ExFont_FontInfoCollection_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExFont::FontInfoCollection)>::type;

struct ExFont_FontInfoCollection : public ExFont, public Aspose::Words::ApiExamples::ExFont, public ::testing::WithParamInterface<ExFont_FontInfoCollection_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExFont_FontInfoCollection, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FontInfoCollection(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFont_FontInfoCollection, ::testing::ValuesIn(ExFont_FontInfoCollection::TestCases()));

} // namespace gtest_test

void ExFont::WorkWithEmbeddedFonts(bool embedTrueTypeFonts, bool embedSystemFonts, bool saveSubsetFonts)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::SharedPtr<Aspose::Words::Fonts::FontInfoCollection> fontInfos = doc->get_FontInfos();
    fontInfos->set_EmbedTrueTypeFonts(embedTrueTypeFonts);
    fontInfos->set_EmbedSystemFonts(embedSystemFonts);
    fontInfos->set_SaveSubsetFonts(saveSubsetFonts);
    
    doc->Save(get_ArtifactsDir() + u"Font.WorkWithEmbeddedFonts.docx");
}

namespace gtest_test
{

using ExFont_WorkWithEmbeddedFonts_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExFont::WorkWithEmbeddedFonts)>::type;

struct ExFont_WorkWithEmbeddedFonts : public ExFont, public Aspose::Words::ApiExamples::ExFont, public ::testing::WithParamInterface<ExFont_WorkWithEmbeddedFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true, false, false),
            std::make_tuple(true, true, false),
            std::make_tuple(true, true, true),
            std::make_tuple(true, false, true),
            std::make_tuple(false, false, false),
        };
    }
};

TEST_P(ExFont_WorkWithEmbeddedFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithEmbeddedFonts(std::get<0>(params), std::get<1>(params), std::get<2>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFont_WorkWithEmbeddedFonts, ::testing::ValuesIn(ExFont_WorkWithEmbeddedFonts::TestCases()));

} // namespace gtest_test

void ExFont::StrikeThrough()
{
    //ExStart
    //ExFor:Font.StrikeThrough
    //ExFor:Font.DoubleStrikeThrough
    //ExSummary:Shows how to add a line strikethrough to text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
    
    auto run = System::MakeObject<Aspose::Words::Run>(doc, u"Text with a single-line strikethrough.");
    run->get_Font()->set_StrikeThrough(true);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    para = System::ExplicitCast<Aspose::Words::Paragraph>(para->get_ParentNode()->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc)));
    
    run = System::MakeObject<Aspose::Words::Run>(doc, u"Text with a double-line strikethrough.");
    run->get_Font()->set_DoubleStrikeThrough(true);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    doc->Save(get_ArtifactsDir() + u"Font.StrikeThrough.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.StrikeThrough.docx");
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Text with a single-line strikethrough.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_StrikeThrough());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Text with a double-line strikethrough.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_DoubleStrikeThrough());
}

namespace gtest_test
{

TEST_F(ExFont, StrikeThrough)
{
    s_instance->StrikeThrough();
}

} // namespace gtest_test

void ExFont::PositionSubscript()
{
    //ExStart
    //ExFor:Font.Position
    //ExFor:Font.Subscript
    //ExFor:Font.Superscript
    //ExSummary:Shows how to format text to offset its position.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
    
    // Raise this run of text 5 points above the baseline.
    auto run = System::MakeObject<Aspose::Words::Run>(doc, u"Raised text. ");
    run->get_Font()->set_Position(5);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    // Lower this run of text 10 points below the baseline.
    run = System::MakeObject<Aspose::Words::Run>(doc, u"Lowered text. ");
    run->get_Font()->set_Position(-10);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    // Add a run of normal text.
    run = System::MakeObject<Aspose::Words::Run>(doc, u"Text in its default position. ");
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    // Add a run of text that appears as subscript.
    run = System::MakeObject<Aspose::Words::Run>(doc, u"Subscript. ");
    run->get_Font()->set_Subscript(true);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    // Add a run of text that appears as superscript.
    run = System::MakeObject<Aspose::Words::Run>(doc, u"Superscript.");
    run->get_Font()->set_Superscript(true);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    doc->Save(get_ArtifactsDir() + u"Font.PositionSubscript.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.PositionSubscript.docx");
    run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Raised text.", run->GetText().Trim());
    ASPOSE_ASSERT_EQ(5, run->get_Font()->get_Position());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.PositionSubscript.docx");
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

namespace gtest_test
{

TEST_F(ExFont, PositionSubscript)
{
    s_instance->PositionSubscript();
}

} // namespace gtest_test

void ExFont::ScalingSpacing()
{
    //ExStart
    //ExFor:Font.Scaling
    //ExFor:Font.Spacing
    //ExSummary:Shows how to set horizontal scaling and spacing for characters.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add run of text and increase character width to 150%.
    builder->get_Font()->set_Scaling(150);
    builder->Writeln(u"Wide characters");
    
    // Add run of text and add 1pt of extra horizontal spacing between each character.
    builder->get_Font()->set_Spacing(1);
    builder->Writeln(u"Expanded by 1pt");
    
    // Add run of text and bring characters closer together by 1pt.
    builder->get_Font()->set_Spacing(-1);
    builder->Writeln(u"Condensed by 1pt");
    
    doc->Save(get_ArtifactsDir() + u"Font.ScalingSpacing.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.ScalingSpacing.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Wide characters", run->GetText().Trim());
    ASSERT_EQ(150, run->get_Font()->get_Scaling());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Expanded by 1pt", run->GetText().Trim());
    ASPOSE_ASSERT_EQ(1, run->get_Font()->get_Spacing());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(2)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Condensed by 1pt", run->GetText().Trim());
    ASPOSE_ASSERT_EQ(-1, run->get_Font()->get_Spacing());
}

namespace gtest_test
{

TEST_F(ExFont, ScalingSpacing)
{
    s_instance->ScalingSpacing();
}

} // namespace gtest_test

void ExFont::Italic()
{
    //ExStart
    //ExFor:Font.Italic
    //ExSummary:Shows how to write italicized text using a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Size(36);
    builder->get_Font()->set_Italic(true);
    builder->Writeln(u"Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"Font.Italic.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Italic.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Hello world!", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_Italic());
}

namespace gtest_test
{

TEST_F(ExFont, Italic)
{
    s_instance->Italic();
}

} // namespace gtest_test

void ExFont::EngraveEmboss()
{
    //ExStart
    //ExFor:Font.Emboss
    //ExFor:Font.Engrave
    //ExSummary:Shows how to apply engraving/embossing effects to text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
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
    
    doc->Save(get_ArtifactsDir() + u"Font.EngraveEmboss.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.EngraveEmboss.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"This text is engraved.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_Engrave());
    ASSERT_FALSE(run->get_Font()->get_Emboss());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"This text is embossed.", run->GetText().Trim());
    ASSERT_FALSE(run->get_Font()->get_Engrave());
    ASSERT_TRUE(run->get_Font()->get_Emboss());
}

namespace gtest_test
{

TEST_F(ExFont, EngraveEmboss)
{
    s_instance->EngraveEmboss();
}

} // namespace gtest_test

void ExFont::Shadow()
{
    //ExStart
    //ExFor:Font.Shadow
    //ExSummary:Shows how to create a run of text formatted with a shadow.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set the Shadow flag to apply an offset shadow effect,
    // making it look like the letters are floating above the page.
    builder->get_Font()->set_Shadow(true);
    builder->get_Font()->set_Size(36);
    
    builder->Writeln(u"This text has a shadow.");
    
    doc->Save(get_ArtifactsDir() + u"Font.Shadow.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Shadow.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"This text has a shadow.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_Shadow());
}

namespace gtest_test
{

TEST_F(ExFont, Shadow)
{
    s_instance->Shadow();
}

} // namespace gtest_test

void ExFont::Outline()
{
    //ExStart
    //ExFor:Font.Outline
    //ExSummary:Shows how to create a run of text formatted as outline.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set the Outline flag to change the text's fill color to white and
    // leave a thin outline around each character in the original color of the text.
    builder->get_Font()->set_Outline(true);
    builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    builder->get_Font()->set_Size(36);
    
    builder->Writeln(u"This text has an outline.");
    
    doc->Save(get_ArtifactsDir() + u"Font.Outline.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Outline.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"This text has an outline.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_Outline());
}

namespace gtest_test
{

TEST_F(ExFont, Outline)
{
    s_instance->Outline();
}

} // namespace gtest_test

void ExFont::Hidden()
{
    //ExStart
    //ExFor:Font.Hidden
    //ExSummary:Shows how to create a run of hidden text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // With the Hidden flag set to true, any text that we create using this Font object will be invisible in the document.
    // We will not see or highlight hidden text unless we enable the "Hidden text" option
    // found in Microsoft Word via "File" -> "Options" -> "Display". The text will still be there,
    // and we will be able to access this text programmatically.
    // It is not advised to use this method to hide sensitive information.
    builder->get_Font()->set_Hidden(true);
    builder->get_Font()->set_Size(36);
    
    builder->Writeln(u"This text will not be visible in the document.");
    
    doc->Save(get_ArtifactsDir() + u"Font.Hidden.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Hidden.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"This text will not be visible in the document.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_Hidden());
}

namespace gtest_test
{

TEST_F(ExFont, Hidden)
{
    s_instance->Hidden();
}

} // namespace gtest_test

void ExFont::Kerning()
{
    //ExStart
    //ExFor:Font.Kerning
    //ExSummary:Shows how to specify the font size at which kerning begins to take effect.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
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
    
    doc->Save(get_ArtifactsDir() + u"Font.Kerning.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Kerning.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"TALLY. (Kerning not applied)", run->GetText().Trim());
    ASPOSE_ASSERT_EQ(24, run->get_Font()->get_Kerning());
    ASPOSE_ASSERT_EQ(18, run->get_Font()->get_Size());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"TALLY. (Kerning applied)", run->GetText().Trim());
    ASPOSE_ASSERT_EQ(12, run->get_Font()->get_Kerning());
    ASPOSE_ASSERT_EQ(18, run->get_Font()->get_Size());
}

namespace gtest_test
{

TEST_F(ExFont, Kerning)
{
    s_instance->Kerning();
}

} // namespace gtest_test

void ExFont::NoProofing()
{
    //ExStart
    //ExFor:Font.NoProofing
    //ExSummary:Shows how to prevent text from being spell checked by Microsoft Word.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Normally, Microsoft Word emphasizes spelling errors with a jagged red underline.
    // We can un-set the "NoProofing" flag to create a portion of text that
    // bypasses the spell checker while completely disabling it.
    builder->get_Font()->set_NoProofing(true);
    
    builder->Writeln(u"Proofing has been disabled, so these spelking errrs will not display red lines underneath.");
    
    doc->Save(get_ArtifactsDir() + u"Font.NoProofing.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.NoProofing.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Proofing has been disabled, so these spelking errrs will not display red lines underneath.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_NoProofing());
}

namespace gtest_test
{

TEST_F(ExFont, NoProofing)
{
    s_instance->NoProofing();
}

} // namespace gtest_test

void ExFont::LocaleId()
{
    //ExStart
    //ExFor:Font.LocaleId
    //ExSummary:Shows how to set the locale of the text that we are adding with a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // If we set the font's locale to English and insert some Russian text,
    // the English locale spell checker will not recognize the text and detect it as a spelling error.
    builder->get_Font()->set_LocaleId(System::MakeObject<System::Globalization::CultureInfo>(u"en-US", false)->get_LCID());
    builder->Writeln(u"Привет!");
    
    // Set a matching locale for the text that we are about to add to apply the appropriate spell checker.
    builder->get_Font()->set_LocaleId(System::MakeObject<System::Globalization::CultureInfo>(u"ru-RU", false)->get_LCID());
    builder->Writeln(u"Привет!");
    
    doc->Save(get_ArtifactsDir() + u"Font.LocaleId.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.LocaleId.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Привет!", run->GetText().Trim());
    ASSERT_EQ(1033, run->get_Font()->get_LocaleId());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Привет!", run->GetText().Trim());
    ASSERT_EQ(1049, run->get_Font()->get_LocaleId());
}

namespace gtest_test
{

TEST_F(ExFont, LocaleId)
{
    s_instance->LocaleId();
}

} // namespace gtest_test

void ExFont::Underlines()
{
    //ExStart
    //ExFor:Font.Underline
    //ExFor:Font.UnderlineColor
    //ExSummary:Shows how to configure the style and color of a text underline.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Underline(Aspose::Words::Underline::Dotted);
    builder->get_Font()->set_UnderlineColor(System::Drawing::Color::get_Red());
    
    builder->Writeln(u"Underlined text.");
    
    doc->Save(get_ArtifactsDir() + u"Font.Underlines.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Underlines.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Underlined text.", run->GetText().Trim());
    ASSERT_EQ(Aspose::Words::Underline::Dotted, run->get_Font()->get_Underline());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), run->get_Font()->get_UnderlineColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExFont, Underlines)
{
    s_instance->Underlines();
}

} // namespace gtest_test

void ExFont::ComplexScript()
{
    //ExStart
    //ExFor:Font.ComplexScript
    //ExSummary:Shows how to add text that is always treated as complex script.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_ComplexScript(true);
    
    builder->Writeln(u"Text treated as complex script.");
    
    doc->Save(get_ArtifactsDir() + u"Font.ComplexScript.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.ComplexScript.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Text treated as complex script.", run->GetText().Trim());
    ASSERT_TRUE(run->get_Font()->get_ComplexScript());
}

namespace gtest_test
{

TEST_F(ExFont, ComplexScript)
{
    s_instance->ComplexScript();
}

} // namespace gtest_test

void ExFont::SparklingText()
{
    //ExStart
    //ExFor:TextEffect
    //ExFor:Font.TextEffect
    //ExSummary:Shows how to apply a visual effect to a run.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Size(36);
    builder->get_Font()->set_TextEffect(Aspose::Words::TextEffect::SparkleText);
    
    builder->Writeln(u"Text with a sparkle effect.");
    
    // Older versions of Microsoft Word only support font animation effects.
    doc->Save(get_ArtifactsDir() + u"Font.SparklingText.doc");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.SparklingText.doc");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Text with a sparkle effect.", run->GetText().Trim());
    ASSERT_EQ(Aspose::Words::TextEffect::SparkleText, run->get_Font()->get_TextEffect());
}

namespace gtest_test
{

TEST_F(ExFont, SparklingText)
{
    s_instance->SparklingText();
}

} // namespace gtest_test

void ExFont::ForegroundAndBackground()
{
    //ExStart
    //ExFor:Shading.ForegroundPatternThemeColor
    //ExFor:Shading.BackgroundPatternThemeColor
    //ExFor:Shading.ForegroundTintAndShade
    //ExFor:Shading.BackgroundTintAndShade
    //ExSummary:Shows how to set foreground and background colors for shading texture.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Shading> shading = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Shading();
    shading->set_Texture(Aspose::Words::TextureIndex::Texture12Pt5Percent);
    shading->set_ForegroundPatternThemeColor(Aspose::Words::Themes::ThemeColor::Dark1);
    shading->set_BackgroundPatternThemeColor(Aspose::Words::Themes::ThemeColor::Dark2);
    
    shading->set_ForegroundTintAndShade(0.5);
    shading->set_BackgroundTintAndShade(-0.2);
    
    builder->get_Font()->get_Border()->set_Color(System::Drawing::Color::get_Green());
    builder->get_Font()->get_Border()->set_LineWidth(2.5);
    builder->get_Font()->get_Border()->set_LineStyle(Aspose::Words::LineStyle::DashDotStroker);
    
    builder->Writeln(u"Foreground and background pattern colors for shading texture.");
    
    doc->Save(get_ArtifactsDir() + u"Font.ForegroundAndBackground.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.ForegroundAndBackground.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Foreground and background pattern colors for shading texture.", run->GetText().Trim());
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::Dark1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Shading()->get_ForegroundPatternThemeColor());
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::Dark2, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Shading()->get_BackgroundPatternThemeColor());
    
    ASSERT_NEAR(0.5, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Shading()->get_ForegroundTintAndShade(), 0.1);
    ASSERT_NEAR(-0.2, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Shading()->get_BackgroundTintAndShade(), 0.1);
}

namespace gtest_test
{

TEST_F(ExFont, ForegroundAndBackground)
{
    s_instance->ForegroundAndBackground();
}

} // namespace gtest_test

void ExFont::Shading()
{
    //ExStart
    //ExFor:Font.Shading
    //ExSummary:Shows how to apply shading to text created by a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Color(System::Drawing::Color::get_White());
    
    // One way to make the text created using our white font color visible
    // is to apply a background shading effect.
    System::SharedPtr<Aspose::Words::Shading> shading = builder->get_Font()->get_Shading();
    shading->set_Texture(Aspose::Words::TextureIndex::TextureDiagonalUp);
    shading->set_BackgroundPatternColor(System::Drawing::Color::get_OrangeRed());
    shading->set_ForegroundPatternColor(System::Drawing::Color::get_DarkBlue());
    
    builder->Writeln(u"White text on an orange background with a two-tone texture.");
    
    doc->Save(get_ArtifactsDir() + u"Font.Shading.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Shading.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"White text on an orange background with a two-tone texture.", run->GetText().Trim());
    ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), run->get_Font()->get_Color().ToArgb());
    
    ASSERT_EQ(Aspose::Words::TextureIndex::TextureDiagonalUp, run->get_Font()->get_Shading()->get_Texture());
    ASSERT_EQ(System::Drawing::Color::get_OrangeRed().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_DarkBlue().ToArgb(), run->get_Font()->get_Shading()->get_ForegroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExFont, Shading)
{
    s_instance->Shading();
}

} // namespace gtest_test

void ExFont::Bidi()
{
    //ExStart
    //ExFor:Font.Bidi
    //ExFor:Font.NameBi
    //ExFor:Font.SizeBi
    //ExFor:Font.ItalicBi
    //ExFor:Font.BoldBi
    //ExFor:Font.LocaleIdBi
    //ExSummary:Shows how to define separate sets of font settings for right-to-left, and right-to-left text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Define a set of font settings for left-to-right text.
    builder->get_Font()->set_Name(u"Courier New");
    builder->get_Font()->set_Size(16);
    builder->get_Font()->set_Italic(false);
    builder->get_Font()->set_Bold(false);
    builder->get_Font()->set_LocaleId(System::MakeObject<System::Globalization::CultureInfo>(u"en-US", false)->get_LCID());
    
    // Define another set of font settings for right-to-left text.
    builder->get_Font()->set_NameBi(u"Andalus");
    builder->get_Font()->set_SizeBi(24);
    builder->get_Font()->set_ItalicBi(true);
    builder->get_Font()->set_BoldBi(true);
    builder->get_Font()->set_LocaleIdBi(System::MakeObject<System::Globalization::CultureInfo>(u"ar-AR", false)->get_LCID());
    
    // We can use the Bidi flag to indicate whether the text we are about to add
    // with the document builder is right-to-left. When we add text with this flag set to true,
    // it will be formatted using the right-to-left set of font settings.
    builder->get_Font()->set_Bidi(true);
    builder->Write(u"مرحبًا");
    
    // Set the flag to false, and then add left-to-right text.
    // The document builder will format these using the left-to-right set of font settings.
    builder->get_Font()->set_Bidi(false);
    builder->Write(u" Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"Font.Bidi.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Bidi.docx");
    
    for (auto&& run : System::IterateOver<Aspose::Words::Run>(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()))
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
        ASSERT_FALSE(run->get_Font()->get_Italic());
        ASSERT_FALSE(run->get_Font()->get_Bold());
        ASSERT_EQ(1025, run->get_Font()->get_LocaleIdBi());
        ASPOSE_ASSERT_EQ(24, run->get_Font()->get_SizeBi());
        ASSERT_EQ(u"Andalus", run->get_Font()->get_NameBi());
        ASSERT_TRUE(run->get_Font()->get_ItalicBi());
        ASSERT_TRUE(run->get_Font()->get_BoldBi());
    }
}

namespace gtest_test
{

TEST_F(ExFont, SkipMono_Bidi)
{
    RecordProperty("category", "SkipMono");
    
    s_instance->Bidi();
}

} // namespace gtest_test

void ExFont::FarEast()
{
    //ExStart
    //ExFor:Font.NameFarEast
    //ExFor:Font.LocaleIdFarEast
    //ExSummary:Shows how to insert and format text in a Far East language.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Specify font settings that the document builder will apply to any text that it inserts.
    builder->get_Font()->set_Name(u"Courier New");
    builder->get_Font()->set_LocaleId(System::MakeObject<System::Globalization::CultureInfo>(u"en-US", false)->get_LCID());
    
    // Name "FarEast" equivalents for our font and locale.
    // If the builder inserts Asian characters with this Font configuration, then each run that contains
    // these characters will display them using the "FarEast" font/locale instead of the default.
    // This could be useful when a western font does not have ideal representations for Asian characters.
    builder->get_Font()->set_NameFarEast(u"SimSun");
    builder->get_Font()->set_LocaleIdFarEast(System::MakeObject<System::Globalization::CultureInfo>(u"zh-CN", false)->get_LCID());
    
    // This text will be displayed in the default font/locale.
    builder->Writeln(u"Hello world!");
    
    // Since these are Asian characters, this run will apply our "FarEast" font/locale equivalents.
    builder->Writeln(u"你好世界");
    
    doc->Save(get_ArtifactsDir() + u"Font.FarEast.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.FarEast.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
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

namespace gtest_test
{

TEST_F(ExFont, FarEast)
{
    s_instance->FarEast();
}

} // namespace gtest_test

void ExFont::NameAscii()
{
    //ExStart
    //ExFor:Font.NameAscii
    //ExFor:Font.NameOther
    //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
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
    
    doc->Save(get_ArtifactsDir() + u"Font.NameAscii.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.NameAscii.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Hello, Привет", run->GetText().Trim());
    ASSERT_EQ(u"Calibri", run->get_Font()->get_Name());
    ASSERT_EQ(u"Calibri", run->get_Font()->get_NameAscii());
    ASSERT_EQ(u"Courier New", run->get_Font()->get_NameOther());
}

namespace gtest_test
{

TEST_F(ExFont, NameAscii)
{
    s_instance->NameAscii();
}

} // namespace gtest_test

void ExFont::ChangeStyle()
{
    //ExStart
    //ExFor:Font.StyleName
    //ExFor:Font.StyleIdentifier
    //ExFor:StyleIdentifier
    //ExSummary:Shows how to change the style of existing text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two ways of referencing styles.
    // 1 -  Using the style name:
    builder->get_Font()->set_StyleName(u"Emphasis");
    builder->Writeln(u"Text originally in \"Emphasis\" style");
    
    // 2 -  Using a built-in style identifier:
    builder->get_Font()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::IntenseEmphasis);
    builder->Writeln(u"Text originally in \"Intense Emphasis\" style");
    
    // Convert all uses of one style to another,
    // using the above methods to reference old and new styles.
    for (auto&& run : System::IterateOver<Aspose::Words::Run>(doc->GetChildNodes(Aspose::Words::NodeType::Run, true)))
    {
        if (run->get_Font()->get_StyleName() == u"Emphasis")
        {
            run->get_Font()->set_StyleName(u"Strong");
        }
        
        if (run->get_Font()->get_StyleIdentifier() == Aspose::Words::StyleIdentifier::IntenseEmphasis)
        {
            run->get_Font()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Strong);
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"Font.ChangeStyle.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.ChangeStyle.docx");
    System::SharedPtr<Aspose::Words::Run> docRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Text originally in \"Emphasis\" style", docRun->GetText().Trim());
    ASSERT_EQ(Aspose::Words::StyleIdentifier::Strong, docRun->get_Font()->get_StyleIdentifier());
    ASSERT_EQ(u"Strong", docRun->get_Font()->get_StyleName());
    
    docRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"Text originally in \"Intense Emphasis\" style", docRun->GetText().Trim());
    ASSERT_EQ(Aspose::Words::StyleIdentifier::Strong, docRun->get_Font()->get_StyleIdentifier());
    ASSERT_EQ(u"Strong", docRun->get_Font()->get_StyleName());
}

namespace gtest_test
{

TEST_F(ExFont, ChangeStyle)
{
    s_instance->ChangeStyle();
}

} // namespace gtest_test

void ExFont::BuiltIn()
{
    //ExStart
    //ExFor:Style.BuiltIn
    //ExSummary:Shows how to differentiate custom styles from built-in styles.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // When we create a document using Microsoft Word, or programmatically using Aspose.Words,
    // the document will come with a collection of styles to apply to its text to modify its appearance.
    // We can access these built-in styles via the document's "Styles" collection.
    // These styles will all have the "BuiltIn" flag set to "true".
    System::SharedPtr<Aspose::Words::Style> style = doc->get_Styles()->idx_get(u"Emphasis");
    
    ASSERT_TRUE(style->get_BuiltIn());
    
    // Create a custom style and add it to the collection.
    // Custom styles such as this will have the "BuiltIn" flag set to "false".
    style = doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"MyStyle");
    style->get_Font()->set_Color(System::Drawing::Color::get_Navy());
    style->get_Font()->set_Name(u"Courier New");
    
    ASSERT_FALSE(style->get_BuiltIn());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFont, BuiltIn)
{
    s_instance->BuiltIn();
}

} // namespace gtest_test

void ExFont::Style()
{
    //ExStart
    //ExFor:Font.Style
    //ExSummary:Applies a double underline to all runs in a document that are formatted with custom character styles.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a custom style and apply it to text created using a document builder.
    System::SharedPtr<Aspose::Words::Style> style = doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"MyStyle");
    style->get_Font()->set_Color(System::Drawing::Color::get_Red());
    style->get_Font()->set_Name(u"Courier New");
    
    builder->get_Font()->set_StyleName(u"MyStyle");
    builder->Write(u"This text is in a custom style.");
    
    // Iterate over every run and add a double underline to every custom style.
    for (auto&& run : System::IterateOver<Aspose::Words::Run>(doc->GetChildNodes(Aspose::Words::NodeType::Run, true)))
    {
        System::SharedPtr<Aspose::Words::Style> charStyle = run->get_Font()->get_Style();
        
        if (!charStyle->get_BuiltIn())
        {
            run->get_Font()->set_Underline(Aspose::Words::Underline::Double);
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"Font.Style.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.Style.docx");
    System::SharedPtr<Aspose::Words::Run> docRun = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"This text is in a custom style.", docRun->GetText().Trim());
    ASSERT_EQ(u"MyStyle", docRun->get_Font()->get_StyleName());
    ASSERT_FALSE(docRun->get_Font()->get_Style()->get_BuiltIn());
    ASSERT_EQ(Aspose::Words::Underline::Double, docRun->get_Font()->get_Underline());
}

namespace gtest_test
{

TEST_F(ExFont, Style)
{
    s_instance->Style();
}

} // namespace gtest_test

void ExFont::GetAvailableFonts()
{
    //ExStart
    //ExFor:PhysicalFontInfo
    //ExFor:FontSourceBase.GetAvailableFonts
    //ExFor:PhysicalFontInfo.FontFamilyName
    //ExFor:PhysicalFontInfo.FullFontName
    //ExFor:PhysicalFontInfo.Version
    //ExFor:PhysicalFontInfo.FilePath
    //ExSummary:Shows how to list available fonts.
    // Configure Aspose.Words to source fonts from a custom folder, and then print every available font.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> folderFontSource = System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({System::MakeObject<Aspose::Words::Fonts::FolderFontSource>(get_FontsDir(), true)});
    
    for (auto&& fontInfo : System::IterateOver(folderFontSource[0]->GetAvailableFonts()))
    {
        std::cout << "FontFamilyName : " << fontInfo->get_FontFamilyName() << std::endl;
        std::cout << "FullFontName  : " << fontInfo->get_FullFontName() << std::endl;
        std::cout << "Version  : " << fontInfo->get_Version() << std::endl;
        std::cout << "FilePath : " << fontInfo->get_FilePath() << "\n" << std::endl;
    }
    //ExEnd
    
    ASSERT_EQ(folderFontSource[0]->GetAvailableFonts()->get_Count(), System::IO::Directory::EnumerateFiles(get_FontsDir(), u"*.*", System::IO::SearchOption::AllDirectories)->LINQ_Count(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String f)>>([](System::String f) -> bool
    {
        return f.EndsWith(u".ttf") || f.EndsWith(u".otf");
    }))) + 5);
}

namespace gtest_test
{

TEST_F(ExFont, GetAvailableFonts)
{
    s_instance->GetAvailableFonts();
}

} // namespace gtest_test

void ExFont::SetFontAutoColor()
{
    //ExStart
    //ExFor:Font.AutoColor
    //ExSummary:Shows how to improve readability by automatically selecting text color based on the brightness of its background.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // If a run's Font object does not specify text color, it will automatically
    // select either black or white depending on the background color's color.
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), builder->get_Font()->get_Color().ToArgb());
    
    // The default color for text is black. If the color of the background is dark, black text will be difficult to see.
    // To solve this problem, the AutoColor property will display this text in white.
    builder->get_Font()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_DarkBlue());
    
    builder->Writeln(u"The text color automatically chosen for this run is white.");
    
    ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->get_Font()->get_AutoColor().ToArgb());
    
    // If we change the background to a light color, black will be a more
    // suitable text color than white so that the auto color will display it in black.
    builder->get_Font()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
    
    builder->Writeln(u"The text color automatically chosen for this run is black.");
    
    ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->get_Font()->get_AutoColor().ToArgb());
    
    doc->Save(get_ArtifactsDir() + u"Font.SetFontAutoColor.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.SetFontAutoColor.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"The text color automatically chosen for this run is white.", run->GetText().Trim());
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), run->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_DarkBlue().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());
    
    run = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0);
    
    ASSERT_EQ(u"The text color automatically chosen for this run is black.", run->GetText().Trim());
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), run->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), run->get_Font()->get_Shading()->get_BackgroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExFont, SetFontAutoColor)
{
    s_instance->SetFontAutoColor();
}

} // namespace gtest_test

void ExFont::RemoveHiddenContentFromDocument()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Hidden content.docx");
    ASSERT_EQ(26, doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());
    //ExSkip
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    //ExSkip
    
    auto hiddenContentRemover = System::MakeObject<Aspose::Words::ApiExamples::ExFont::RemoveHiddenContentVisitor>();
    
    // Below are three types of fields which can accept a document visitor,
    // which will allow it to visit the accepting node, and then traverse its child nodes in a depth-first manner.
    // 1 -  Paragraph node:
    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 4, true));
    para->Accept(hiddenContentRemover);
    
    // 2 -  Table node:
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    table->Accept(hiddenContentRemover);
    
    // 3 -  Document node:
    doc->Accept(hiddenContentRemover);
    
    doc->Save(get_ArtifactsDir() + u"Font.RemoveHiddenContentFromDocument.docx");
    TestRemoveHiddenContent(System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Font.RemoveHiddenContentFromDocument.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExFont, RemoveHiddenContentFromDocument)
{
    s_instance->RemoveHiddenContentFromDocument();
}

} // namespace gtest_test

void ExFont::DefaultFonts()
{
    //ExStart
    //ExFor:FontInfoCollection.Contains(String)
    //ExFor:FontInfoCollection.Count
    //ExSummary:Shows info about the fonts that are present in the blank document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A blank document contains 3 default fonts. Each font in the document
    // will have a corresponding FontInfo object which contains details about that font.
    ASSERT_EQ(3, doc->get_FontInfos()->get_Count());
    
    ASSERT_TRUE(doc->get_FontInfos()->Contains(u"Times New Roman"));
    ASSERT_EQ(204, doc->get_FontInfos()->idx_get(u"Times New Roman")->get_Charset());
    
    ASSERT_TRUE(doc->get_FontInfos()->Contains(u"Symbol"));
    ASSERT_TRUE(doc->get_FontInfos()->Contains(u"Arial"));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFont, DefaultFonts)
{
    s_instance->DefaultFonts();
}

} // namespace gtest_test

void ExFont::ExtractEmbeddedFont()
{
    //ExStart
    //ExFor:EmbeddedFontFormat
    //ExFor:EmbeddedFontStyle
    //ExFor:FontInfo.GetEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
    //ExFor:FontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle)
    //ExFor:FontInfoCollection.Item(Int32)
    //ExFor:FontInfoCollection.Item(String)
    //ExSummary:Shows how to extract an embedded font from a document, and save it to the local file system.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Embedded font.docx");
    
    System::SharedPtr<Aspose::Words::Fonts::FontInfo> embeddedFont = doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift");
    System::ArrayPtr<uint8_t> embeddedFontBytes = embeddedFont->GetEmbeddedFont(Aspose::Words::Fonts::EmbeddedFontFormat::OpenType, Aspose::Words::Fonts::EmbeddedFontStyle::Regular);
    ASSERT_FALSE(System::TestTools::IsNull(embeddedFontBytes));
    //ExSkip
    
    System::IO::File::WriteAllBytes(get_ArtifactsDir() + u"Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);
    
    // Embedded font formats may be different in other formats such as .doc.
    // We need to know the correct format before we can extract the font.
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Embedded font.doc");
    
    ASSERT_TRUE(System::TestTools::IsNull(doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFont(Aspose::Words::Fonts::EmbeddedFontFormat::OpenType, Aspose::Words::Fonts::EmbeddedFontStyle::Regular)));
    ASSERT_FALSE(System::TestTools::IsNull(doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFont(Aspose::Words::Fonts::EmbeddedFontFormat::EmbeddedOpenType, Aspose::Words::Fonts::EmbeddedFontStyle::Regular)));
    
    // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType.
    embeddedFontBytes = doc->get_FontInfos()->idx_get(u"Alte DIN 1451 Mittelschrift")->GetEmbeddedFontAsOpenType(Aspose::Words::Fonts::EmbeddedFontStyle::Regular);
    
    System::IO::File::WriteAllBytes(get_ArtifactsDir() + u"Alte DIN 1451 Mittelschrift.otf", embeddedFontBytes);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFont, ExtractEmbeddedFont)
{
    s_instance->ExtractEmbeddedFont();
}

} // namespace gtest_test

void ExFont::GetFontInfoFromFile()
{
    //ExStart
    //ExFor:FontFamily
    //ExFor:FontPitch
    //ExFor:FontInfo.AltName
    //ExFor:FontInfo.Charset
    //ExFor:FontInfo.Family
    //ExFor:FontInfo.Panose
    //ExFor:FontInfo.Pitch
    //ExFor:FontInfoCollection.GetEnumerator
    //ExSummary:Shows how to access and print details of each font in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Fonts::FontInfo>>> fontCollectionEnumerator = doc->get_FontInfos()->GetEnumerator();
    while (fontCollectionEnumerator->MoveNext())
    {
        System::SharedPtr<Aspose::Words::Fonts::FontInfo> fontInfo = fontCollectionEnumerator->get_Current();
        if (fontInfo != nullptr)
        {
            std::cout << (System::String(u"Font name: ") + fontInfo->get_Name()) << std::endl;
            
            // Alt names are usually blank.
            std::cout << (System::String(u"Alt name: ") + fontInfo->get_AltName()) << std::endl;
            std::cout << (System::String(u"\t- Family: ") + System::ObjectExt::ToString(fontInfo->get_Family())) << std::endl;
            std::cout << (System::String(u"\t- ") + (fontInfo->get_IsTrueType() ? System::String(u"Is TrueType") : System::String(u"Is not TrueType"))) << std::endl;
            std::cout << (System::String(u"\t- Pitch: ") + System::ObjectExt::ToString(fontInfo->get_Pitch())) << std::endl;
            std::cout << (System::String(u"\t- Charset: ") + fontInfo->get_Charset()) << std::endl;
            std::cout << "\t- Panose:" << std::endl;
            std::cout << (System::String(u"\t\tFamily Kind: ") + fontInfo->get_Panose()[0]) << std::endl;
            std::cout << (System::String(u"\t\tSerif Style: ") + fontInfo->get_Panose()[1]) << std::endl;
            std::cout << (System::String(u"\t\tWeight: ") + fontInfo->get_Panose()[2]) << std::endl;
            std::cout << (System::String(u"\t\tProportion: ") + fontInfo->get_Panose()[3]) << std::endl;
            std::cout << (System::String(u"\t\tContrast: ") + fontInfo->get_Panose()[4]) << std::endl;
            std::cout << (System::String(u"\t\tStroke Variation: ") + fontInfo->get_Panose()[5]) << std::endl;
            std::cout << (System::String(u"\t\tArm Style: ") + fontInfo->get_Panose()[6]) << std::endl;
            std::cout << (System::String(u"\t\tLetterform: ") + fontInfo->get_Panose()[7]) << std::endl;
            std::cout << (System::String(u"\t\tMidline: ") + fontInfo->get_Panose()[8]) << std::endl;
            std::cout << (System::String(u"\t\tX-Height: ") + fontInfo->get_Panose()[9]) << std::endl;
        }
    }
    //ExEnd
    
    ASPOSE_ASSERT_EQ(System::MakeArray<int32_t>({2, 15, 5, 2, 2, 2, 4, 3, 2, 4}), doc->get_FontInfos()->idx_get(u"Calibri")->get_Panose());
    ASPOSE_ASSERT_EQ(System::MakeArray<int32_t>({2, 15, 3, 2, 2, 2, 4, 3, 2, 4}), doc->get_FontInfos()->idx_get(u"Calibri Light")->get_Panose());
    ASPOSE_ASSERT_EQ(System::MakeArray<int32_t>({2, 2, 6, 3, 5, 4, 5, 2, 3, 4}), doc->get_FontInfos()->idx_get(u"Times New Roman")->get_Panose());
}

namespace gtest_test
{

TEST_F(ExFont, GetFontInfoFromFile)
{
    s_instance->GetFontInfoFromFile();
}

} // namespace gtest_test

void ExFont::LineSpacing()
{
    //ExStart
    //ExFor:Font.LineSpacing
    //ExSummary:Shows how to get a font's line spacing, in points.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set different fonts for the DocumentBuilder and verify their line spacing.
    builder->get_Font()->set_Name(u"Calibri");
    ASPOSE_ASSERT_EQ(14.6484375, builder->get_Font()->get_LineSpacing());
    
    builder->get_Font()->set_Name(u"Times New Roman");
    ASPOSE_ASSERT_EQ(13.798828125, builder->get_Font()->get_LineSpacing());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFont, LineSpacing)
{
    s_instance->LineSpacing();
}

} // namespace gtest_test

void ExFont::HasDmlEffect()
{
    //ExStart
    //ExFor:Font.HasDmlEffect(TextDmlEffect)
    //ExFor:TextDmlEffect
    //ExSummary:Shows how to check if a run displays a DrawingML text effect.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DrawingML text effects.docx");
    
    System::SharedPtr<Aspose::Words::RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();
    
    ASSERT_TRUE(runs->idx_get(0)->get_Font()->HasDmlEffect(Aspose::Words::TextDmlEffect::Shadow));
    ASSERT_TRUE(runs->idx_get(1)->get_Font()->HasDmlEffect(Aspose::Words::TextDmlEffect::Shadow));
    ASSERT_TRUE(runs->idx_get(2)->get_Font()->HasDmlEffect(Aspose::Words::TextDmlEffect::Reflection));
    ASSERT_TRUE(runs->idx_get(3)->get_Font()->HasDmlEffect(Aspose::Words::TextDmlEffect::Effect3D));
    ASSERT_TRUE(runs->idx_get(4)->get_Font()->HasDmlEffect(Aspose::Words::TextDmlEffect::Fill));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFont, HasDmlEffect)
{
    s_instance->HasDmlEffect();
}

} // namespace gtest_test

void ExFont::SetEmphasisMark(Aspose::Words::EmphasisMark emphasisMark)
{
    //ExStart
    //ExFor:EmphasisMark
    //ExFor:Font.EmphasisMark
    //ExSummary:Shows how to add additional character rendered above/below the glyph-character.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    // Possible types of emphasis mark:
    // https://apireference.aspose.com/words/net/aspose.words/emphasismark
    builder->get_Font()->set_EmphasisMark(emphasisMark);
    
    builder->Write(u"Emphasis text");
    builder->Writeln();
    builder->get_Font()->ClearFormatting();
    builder->Write(u"Simple text");
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"Fonts.SetEmphasisMark.docx");
    //ExEnd
}

namespace gtest_test
{

using ExFont_SetEmphasisMark_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExFont::SetEmphasisMark)>::type;

struct ExFont_SetEmphasisMark : public ExFont, public Aspose::Words::ApiExamples::ExFont, public ::testing::WithParamInterface<ExFont_SetEmphasisMark_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::EmphasisMark::None),
            std::make_tuple(Aspose::Words::EmphasisMark::OverComma),
            std::make_tuple(Aspose::Words::EmphasisMark::OverSolidCircle),
            std::make_tuple(Aspose::Words::EmphasisMark::OverWhiteCircle),
            std::make_tuple(Aspose::Words::EmphasisMark::UnderSolidCircle),
        };
    }
};

TEST_P(ExFont_SetEmphasisMark, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetEmphasisMark(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFont_SetEmphasisMark, ::testing::ValuesIn(ExFont_SetEmphasisMark::TestCases()));

} // namespace gtest_test

void ExFont::ThemeFontsColors()
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Define fonts for languages uses by default.
    doc->get_Theme()->get_MinorFonts()->set_Latin(u"Algerian");
    doc->get_Theme()->get_MinorFonts()->set_EastAsian(u"Aharoni");
    doc->get_Theme()->get_MinorFonts()->set_ComplexScript(u"Andalus");
    
    System::SharedPtr<Aspose::Words::Font> font = doc->get_Styles()->idx_get(u"Normal")->get_Font();
    std::cout << System::String::Format(u"Originally the Normal style theme color is: {0} and RGB color is: {1}\n", font->get_ThemeColor(), font->get_Color()) << std::endl;
    
    // We can use theme font and color instead of default values.
    font->set_ThemeFont(Aspose::Words::Themes::ThemeFont::Minor);
    font->set_ThemeColor(Aspose::Words::Themes::ThemeColor::Accent2);
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Minor, font->get_ThemeFont());
    ASSERT_EQ(u"Algerian", font->get_Name());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Minor, font->get_ThemeFontAscii());
    ASSERT_EQ(u"Algerian", font->get_NameAscii());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Minor, font->get_ThemeFontBi());
    ASSERT_EQ(u"Andalus", font->get_NameBi());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Minor, font->get_ThemeFontFarEast());
    ASSERT_EQ(u"Aharoni", font->get_NameFarEast());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Minor, font->get_ThemeFontOther());
    ASSERT_EQ(u"Algerian", font->get_NameOther());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::Accent2, font->get_ThemeColor());
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, font->get_Color());
    
    // There are several ways of reset them font and color.
    // 1 -  By setting ThemeFont.None/ThemeColor.None:
    font->set_ThemeFont(Aspose::Words::Themes::ThemeFont::None);
    font->set_ThemeColor(Aspose::Words::Themes::ThemeColor::None);
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFont());
    ASSERT_EQ(u"Algerian", font->get_Name());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontAscii());
    ASSERT_EQ(u"Algerian", font->get_NameAscii());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontBi());
    ASSERT_EQ(u"Andalus", font->get_NameBi());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontFarEast());
    ASSERT_EQ(u"Aharoni", font->get_NameFarEast());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontOther());
    ASSERT_EQ(u"Algerian", font->get_NameOther());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::None, font->get_ThemeColor());
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, font->get_Color());
    
    // 2 -  By setting non-theme font/color names:
    font->set_Name(u"Arial");
    font->set_Color(System::Drawing::Color::get_Blue());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFont());
    ASSERT_EQ(u"Arial", font->get_Name());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontAscii());
    ASSERT_EQ(u"Arial", font->get_NameAscii());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontBi());
    ASSERT_EQ(u"Arial", font->get_NameBi());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontFarEast());
    ASSERT_EQ(u"Arial", font->get_NameFarEast());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::None, font->get_ThemeFontOther());
    ASSERT_EQ(u"Arial", font->get_NameOther());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::None, font->get_ThemeColor());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), font->get_Color().ToArgb());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFont, ThemeFontsColors)
{
    s_instance->ThemeFontsColors();
}

} // namespace gtest_test

void ExFont::CreateThemedStyle()
{
    //ExStart
    //ExFor:Font.ThemeFont
    //ExFor:Font.ThemeColor
    //ExFor:Font.TintAndShade
    //ExFor:ThemeFont
    //ExFor:ThemeColor
    //ExSummary:Shows how to create and use themed style.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln();
    
    // Create some style with theme font properties.
    System::SharedPtr<Aspose::Words::Style> style = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"ThemedStyle");
    style->get_Font()->set_ThemeFont(Aspose::Words::Themes::ThemeFont::Major);
    style->get_Font()->set_ThemeColor(Aspose::Words::Themes::ThemeColor::Accent5);
    style->get_Font()->set_TintAndShade(0.3);
    
    builder->get_ParagraphFormat()->set_StyleName(u"ThemedStyle");
    builder->Writeln(u"Text with themed style");
    //ExEnd
    
    auto run = System::ExplicitCast<Aspose::Words::Run>((System::ExplicitCast<Aspose::Words::Paragraph>(builder->get_CurrentParagraph()->get_PreviousSibling()))->get_FirstChild());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Major, run->get_Font()->get_ThemeFont());
    ASSERT_EQ(u"Times New Roman", run->get_Font()->get_Name());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Major, run->get_Font()->get_ThemeFontAscii());
    ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameAscii());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Major, run->get_Font()->get_ThemeFontBi());
    ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameBi());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Major, run->get_Font()->get_ThemeFontFarEast());
    ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameFarEast());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeFont::Major, run->get_Font()->get_ThemeFontOther());
    ASSERT_EQ(u"Times New Roman", run->get_Font()->get_NameOther());
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::Accent5, run->get_Font()->get_ThemeColor());
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, run->get_Font()->get_Color());
}

namespace gtest_test
{

TEST_F(ExFont, CreateThemedStyle)
{
    s_instance->CreateThemedStyle();
}

} // namespace gtest_test

void ExFont::FontInfoEmbeddingLicensingRights()
{
    //ExStart:FontInfoEmbeddingLicensingRights
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:FontInfo.EmbeddingLicensingRights
    //ExFor:FontEmbeddingUsagePermissions
    //ExFor:FontEmbeddingLicensingRights
    //ExFor:FontEmbeddingLicensingRights.EmbeddingUsagePermissions
    //ExFor:FontEmbeddingLicensingRights.BitmapEmbeddingOnly
    //ExFor:FontEmbeddingLicensingRights.NoSubsetting
    //ExSummary:Shows how to get license rights information for embedded fonts (FontInfo).
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Embedded font rights.docx");
    
    // Get the list of document fonts.
    System::SharedPtr<Aspose::Words::Fonts::FontInfoCollection> fontInfos = doc->get_FontInfos();
    for (auto&& fontInfo : fontInfos)
    {
        if (fontInfo->get_EmbeddingLicensingRights() != nullptr)
        {
            std::cout << System::EnumGetName(fontInfo->get_EmbeddingLicensingRights()->get_EmbeddingUsagePermissions()) << std::endl;
            std::cout << System::Convert::ToString(fontInfo->get_EmbeddingLicensingRights()->get_BitmapEmbeddingOnly()) << std::endl;
            std::cout << System::Convert::ToString(fontInfo->get_EmbeddingLicensingRights()->get_NoSubsetting()) << std::endl;
        }
    }
    //ExEnd:FontInfoEmbeddingLicensingRights
}

namespace gtest_test
{

TEST_F(ExFont, FontInfoEmbeddingLicensingRights)
{
    s_instance->FontInfoEmbeddingLicensingRights();
}

} // namespace gtest_test

void ExFont::PhysicalFontInfoEmbeddingLicensingRights()
{
    //ExStart:PhysicalFontInfoEmbeddingLicensingRights
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:PhysicalFontInfo.EmbeddingLicensingRights
    //ExSummary:Shows how to get license rights information for embedded fonts (PhysicalFontInfo).
    System::SharedPtr<Aspose::Words::Fonts::FontSettings> settings = Aspose::Words::Fonts::FontSettings::get_DefaultInstance();
    System::SharedPtr<Aspose::Words::Fonts::FontSourceBase> source = settings->GetFontsSources()->idx_get(0);
    
    // Get the list of available fonts.
    System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>>> fontInfos = source->GetAvailableFonts();
    for (auto&& fontInfo : System::IterateOver(fontInfos))
    {
        if (fontInfo->get_EmbeddingLicensingRights() != nullptr)
        {
            std::cout << System::EnumGetName(fontInfo->get_EmbeddingLicensingRights()->get_EmbeddingUsagePermissions()) << std::endl;
            std::cout << System::Convert::ToString(fontInfo->get_EmbeddingLicensingRights()->get_BitmapEmbeddingOnly()) << std::endl;
            std::cout << System::Convert::ToString(fontInfo->get_EmbeddingLicensingRights()->get_NoSubsetting()) << std::endl;
        }
    }
    //ExEnd:PhysicalFontInfoEmbeddingLicensingRights
}

namespace gtest_test
{

TEST_F(ExFont, PhysicalFontInfoEmbeddingLicensingRights)
{
    s_instance->PhysicalFontInfoEmbeddingLicensingRights();
}

} // namespace gtest_test

void ExFont::NumberSpacing()
{
    //ExStart:NumberSpacing
    //GistId:95fdae949cefbf2ce485acc95cccc495
    //ExFor:Font.NumberSpacing
    //ExFor:NumSpacing
    //ExSummary:Shows how to set the spacing type of the numeral.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // This effect is only supported in newer versions of MS Word.
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2019);
    
    builder->Write(u"1 ");
    builder->Write(u"This is an example");
    
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    if (run->get_Font()->get_NumberSpacing() == Aspose::Words::NumSpacing::Default)
    {
        run->get_Font()->set_NumberSpacing(Aspose::Words::NumSpacing::Proportional);
    }
    
    doc->Save(get_ArtifactsDir() + u"Fonts.NumberSpacing.docx");
    //ExEnd:NumberSpacing
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Fonts.NumberSpacing.docx");
    
    run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    ASSERT_EQ(Aspose::Words::NumSpacing::Proportional, run->get_Font()->get_NumberSpacing());
}

namespace gtest_test
{

TEST_F(ExFont, NumberSpacing)
{
    s_instance->NumberSpacing();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
