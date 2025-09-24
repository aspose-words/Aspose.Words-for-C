// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExThemes.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Themes/ThemeFonts.h>
#include <Aspose.Words.Cpp/Model/Themes/ThemeColors.h>
#include <Aspose.Words.Cpp/Model/Themes/Theme.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Themes;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3117042714u, ::Aspose::Words::ApiExamples::ExThemes, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExThemes : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExThemes> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExThemes>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExThemes> ExThemes::s_instance;

} // namespace gtest_test

void ExThemes::CustomColorsAndFonts()
{
    //ExStart
    //ExFor:Document.Theme
    //ExFor:Theme
    //ExFor:Theme.Colors
    //ExFor:Theme.MajorFonts
    //ExFor:Theme.MinorFonts
    //ExFor:ThemeColors
    //ExFor:ThemeColors.Accent1
    //ExFor:ThemeColors.Accent2
    //ExFor:ThemeColors.Accent3
    //ExFor:ThemeColors.Accent4
    //ExFor:ThemeColors.Accent5
    //ExFor:ThemeColors.Accent6
    //ExFor:ThemeColors.Dark1
    //ExFor:ThemeColors.Dark2
    //ExFor:ThemeColors.FollowedHyperlink
    //ExFor:ThemeColors.Hyperlink
    //ExFor:ThemeColors.Light1
    //ExFor:ThemeColors.Light2
    //ExFor:ThemeFonts
    //ExFor:ThemeFonts.ComplexScript
    //ExFor:ThemeFonts.EastAsian
    //ExFor:ThemeFonts.Latin
    //ExSummary:Shows how to set custom colors and fonts for themes.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Theme colors.docx");
    
    // The "Theme" object gives us access to the document theme, a source of default fonts and colors.
    System::SharedPtr<Aspose::Words::Themes::Theme> theme = doc->get_Theme();
    
    // Some styles, such as "Heading 1" and "Subtitle", will inherit these fonts.
    theme->get_MajorFonts()->set_Latin(u"Courier New");
    theme->get_MinorFonts()->set_Latin(u"Agency FB");
    
    // Other languages may also have their custom fonts in this theme.
    ASSERT_EQ(System::String::Empty, theme->get_MajorFonts()->get_ComplexScript());
    ASSERT_EQ(System::String::Empty, theme->get_MajorFonts()->get_EastAsian());
    ASSERT_EQ(System::String::Empty, theme->get_MinorFonts()->get_ComplexScript());
    ASSERT_EQ(System::String::Empty, theme->get_MinorFonts()->get_EastAsian());
    
    // The "Colors" property contains the color palette from Microsoft Word,
    // which appears when changing shading or font color.
    // Apply custom colors to the color palette so we have easy access to them in Microsoft Word
    // when we, for example, change the font color via "Home" -> "Font" -> "Font Color",
    // or insert a shape, and then set a color for it via "Shape Format" -> "Shape Styles".
    System::SharedPtr<Aspose::Words::Themes::ThemeColors> colors = theme->get_Colors();
    colors->set_Dark1(System::Drawing::Color::get_MidnightBlue());
    colors->set_Light1(System::Drawing::Color::get_PaleGreen());
    colors->set_Dark2(System::Drawing::Color::get_Indigo());
    colors->set_Light2(System::Drawing::Color::get_Khaki());
    
    colors->set_Accent1(System::Drawing::Color::get_OrangeRed());
    colors->set_Accent2(System::Drawing::Color::get_LightSalmon());
    colors->set_Accent3(System::Drawing::Color::get_Yellow());
    colors->set_Accent4(System::Drawing::Color::get_Gold());
    colors->set_Accent5(System::Drawing::Color::get_BlueViolet());
    colors->set_Accent6(System::Drawing::Color::get_DarkViolet());
    
    // Apply custom colors to hyperlinks in their clicked and un-clicked states.
    colors->set_Hyperlink(System::Drawing::Color::get_Black());
    colors->set_FollowedHyperlink(System::Drawing::Color::get_Gray());
    
    doc->Save(get_ArtifactsDir() + u"Themes.CustomColorsAndFonts.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Themes.CustomColorsAndFonts.docx");
    
    ASSERT_EQ(System::Drawing::Color::get_OrangeRed().ToArgb(), doc->get_Theme()->get_Colors()->get_Accent1().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_MidnightBlue().ToArgb(), doc->get_Theme()->get_Colors()->get_Dark1().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Gray().ToArgb(), doc->get_Theme()->get_Colors()->get_FollowedHyperlink().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), doc->get_Theme()->get_Colors()->get_Hyperlink().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_PaleGreen().ToArgb(), doc->get_Theme()->get_Colors()->get_Light1().ToArgb());
    
    ASSERT_EQ(System::String::Empty, doc->get_Theme()->get_MajorFonts()->get_ComplexScript());
    ASSERT_EQ(System::String::Empty, doc->get_Theme()->get_MajorFonts()->get_EastAsian());
    ASSERT_EQ(u"Courier New", doc->get_Theme()->get_MajorFonts()->get_Latin());
    
    ASSERT_EQ(System::String::Empty, doc->get_Theme()->get_MinorFonts()->get_ComplexScript());
    ASSERT_EQ(System::String::Empty, doc->get_Theme()->get_MinorFonts()->get_EastAsian());
    ASSERT_EQ(u"Agency FB", doc->get_Theme()->get_MinorFonts()->get_Latin());
}

namespace gtest_test
{

TEST_F(ExThemes, CustomColorsAndFonts)
{
    s_instance->CustomColorsAndFonts();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
