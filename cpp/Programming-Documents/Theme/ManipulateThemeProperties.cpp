#include "stdafx.h"
#include "examples.h"

#include <drawing/color.h>
#include <system/object.h>
#include <system/object_ext.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/Themes/Theme.h>
#include <Model/Themes/ThemeColors.h>
#include <Model/Themes/ThemeFonts.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Themes;

namespace
{
    void GetThemeProperties(System::String const &dataDir)
    {
        // ExStart:GetThemeProperties
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir);
        System::SharedPtr<Theme> theme = doc->get_Theme();
        // Major (Headings) font for Latin characters.
        std::cout << theme->get_MajorFonts()->get_Latin().ToUtf8String() << std::endl;
        // Minor (Body) font for EastAsian characters.
        std::cout << theme->get_MinorFonts()->get_EastAsian().ToUtf8String() << std::endl;
        // Color for theme color Accent 1.
        std::cout << System::ObjectExt::Box<System::Drawing::Color>(theme->get_Colors()->get_Accent1())->ToString().ToUtf8String();
        // ExEnd:GetThemeProperties 
    }

    void SetThemeProperties(System::String const &dataDir)
    {
        // ExStart:SetThemeProperties
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir);
        System::SharedPtr<Theme> theme = doc->get_Theme();
        // Set Times New Roman font as Body theme font for Latin Character.
        theme->get_MinorFonts()->set_Latin(u"Times New Roman");
        // Set Color.Gold for theme color Hyperlink.
        theme->get_Colors()->set_Hyperlink(System::Drawing::Color::get_Gold());
        // ExEnd:SetThemeProperties 
    }
}

void ManipulateThemeProperties()
{
    std::cout << "ManipulateThemeProperties example started." << std::endl;
    // ExStart:ManipulateThemeProperties
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTheme();
    dataDir = dataDir + u"Document.doc";
    GetThemeProperties(dataDir);
    SetThemeProperties(dataDir);
    // ExEnd:ManipulateThemeProperties
    std::cout << "ManipulateThemeProperties example finished." << std::endl << std::endl;
}