#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/Themes/Theme.h>
#include <Aspose.Words.Cpp/Themes/ThemeColors.h>
#include <Aspose.Words.Cpp/Themes/ThemeFonts.h>
#include <drawing/color.h>
#include <system/enumerator_adapter.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management {

class WorkingWithStylesAndThemes : public DocsExamplesBase
{
public:
    void AccessStyles()
    {
        //ExStart:AccessStyles
        auto doc = MakeObject<Document>();

        String styleName = u"";

        // Get styles collection from the document.
        SharedPtr<StyleCollection> styles = doc->get_Styles();
        for (const auto& style : System::IterateOver(styles))
        {
            if (styleName == u"")
            {
                styleName = style->get_Name();
                std::cout << styleName << std::endl;
            }
            else
            {
                styleName = styleName + u", " + style->get_Name();
                std::cout << styleName << std::endl;
            }
        }
        //ExEnd:AccessStyles
    }

    void CopyStyles()
    {
        //ExStart:CopyStyles
        auto doc = MakeObject<Document>();
        auto target = MakeObject<Document>(MyDir + u"Rendering.docx");

        target->CopyStylesFromTemplate(doc);

        doc->Save(ArtifactsDir + u"WorkingWithStylesAndThemes.CopyStyles.docx");
        //ExEnd:CopyStyles
    }

    void GetThemeProperties()
    {
        //ExStart:GetThemeProperties
        auto doc = MakeObject<Document>();

        SharedPtr<Aspose::Words::Themes::Theme> theme = doc->get_Theme();

        std::cout << theme->get_MajorFonts()->get_Latin() << std::endl;
        std::cout << theme->get_MinorFonts()->get_EastAsian() << std::endl;
        std::cout << theme->get_Colors()->get_Accent1() << std::endl;
        //ExEnd:GetThemeProperties
    }

    void SetThemeProperties()
    {
        //ExStart:SetThemeProperties
        auto doc = MakeObject<Document>();

        SharedPtr<Aspose::Words::Themes::Theme> theme = doc->get_Theme();
        theme->get_MinorFonts()->set_Latin(u"Times New Roman");
        theme->get_Colors()->set_Hyperlink(System::Drawing::Color::get_Gold());
        //ExEnd:SetThemeProperties
    }

    void InsertStyleSeparator()
    {
        //ExStart:InsertStyleSeparator
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Style> paraStyle = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"MyParaStyle");
        paraStyle->get_Font()->set_Bold(false);
        paraStyle->get_Font()->set_Size(8);
        paraStyle->get_Font()->set_Name(u"Arial");

        // Append text with "Heading 1" style.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Write(u"Heading 1");
        builder->InsertStyleSeparator();

        // Append text with another style.
        builder->get_ParagraphFormat()->set_StyleName(paraStyle->get_Name());
        builder->Write(u"This is text with some other formatting ");

        doc->Save(ArtifactsDir + u"WorkingWithStylesAndThemes.InsertStyleSeparator.docx");
        //ExEnd:InsertStyleSeparator
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management
