#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Themes/Theme.h>
#include <Aspose.Words.Cpp/Model/Themes/ThemeColors.h>
#include <Aspose.Words.Cpp/Model/Themes/ThemeFonts.h>
#include <drawing/color.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Themes;

namespace ApiExamples {

class ExThemes : public ApiExampleBase
{
public:
    void CustomColorsAndFonts()
    {
        //ExStart
        //ExFor:Document.Theme
        //ExFor:Theme
        //ExFor:Theme.Colors
        //ExFor:Theme.MajorFonts
        //ExFor:Theme.MinorFonts
        //ExFor:Themes.ThemeColors
        //ExFor:Themes.ThemeColors.Accent1
        //ExFor:Themes.ThemeColors.Accent2
        //ExFor:Themes.ThemeColors.Accent3
        //ExFor:Themes.ThemeColors.Accent4
        //ExFor:Themes.ThemeColors.Accent5
        //ExFor:Themes.ThemeColors.Accent6
        //ExFor:Themes.ThemeColors.Dark1
        //ExFor:Themes.ThemeColors.Dark2
        //ExFor:Themes.ThemeColors.FollowedHyperlink
        //ExFor:Themes.ThemeColors.Hyperlink
        //ExFor:Themes.ThemeColors.Light1
        //ExFor:Themes.ThemeColors.Light2
        //ExFor:Themes.ThemeFonts
        //ExFor:Themes.ThemeFonts.ComplexScript
        //ExFor:Themes.ThemeFonts.EastAsian
        //ExFor:Themes.ThemeFonts.Latin
        //ExSummary:Shows how to set custom theme colors and fonts.
        auto doc = MakeObject<Document>(MyDir + u"Theme colors.docx");

        // This object gives us access to the document theme, which is a source of default fonts and colors
        SharedPtr<Theme> theme = doc->get_Theme();

        // These fonts will be inherited by some styles like "Heading 1" and "Subtitle"
        theme->get_MajorFonts()->set_Latin(u"Courier New");
        theme->get_MinorFonts()->set_Latin(u"Agency FB");

        ASSERT_EQ(String::Empty, theme->get_MajorFonts()->get_ComplexScript());
        ASSERT_EQ(String::Empty, theme->get_MajorFonts()->get_EastAsian());
        ASSERT_EQ(String::Empty, theme->get_MinorFonts()->get_ComplexScript());
        ASSERT_EQ(String::Empty, theme->get_MinorFonts()->get_EastAsian());

        // This collection of colors corresponds to the color palette from Microsoft Word which appears when changing shading or font color
        SharedPtr<ThemeColors> colors = theme->get_Colors();

        // We will set the color of each color palette column going from left to right like this
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

        // We can also set colors for hyperlinks like this
        colors->set_Hyperlink(System::Drawing::Color::get_Black());
        colors->set_FollowedHyperlink(System::Drawing::Color::get_Gray());

        doc->Save(ArtifactsDir + u"Themes.CustomColorsAndFonts.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Themes.CustomColorsAndFonts.docx");

        ASSERT_EQ(System::Drawing::Color::get_OrangeRed().ToArgb(), doc->get_Theme()->get_Colors()->get_Accent1().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_MidnightBlue().ToArgb(), doc->get_Theme()->get_Colors()->get_Dark1().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Gray().ToArgb(), doc->get_Theme()->get_Colors()->get_FollowedHyperlink().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), doc->get_Theme()->get_Colors()->get_Hyperlink().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_PaleGreen().ToArgb(), doc->get_Theme()->get_Colors()->get_Light1().ToArgb());

        ASSERT_EQ(String::Empty, doc->get_Theme()->get_MajorFonts()->get_ComplexScript());
        ASSERT_EQ(String::Empty, doc->get_Theme()->get_MajorFonts()->get_EastAsian());
        ASSERT_EQ(u"Courier New", doc->get_Theme()->get_MajorFonts()->get_Latin());

        ASSERT_EQ(String::Empty, doc->get_Theme()->get_MinorFonts()->get_ComplexScript());
        ASSERT_EQ(String::Empty, doc->get_Theme()->get_MinorFonts()->get_EastAsian());
        ASSERT_EQ(u"Agency FB", doc->get_Theme()->get_MinorFonts()->get_Latin());
    }
};

} // namespace ApiExamples
