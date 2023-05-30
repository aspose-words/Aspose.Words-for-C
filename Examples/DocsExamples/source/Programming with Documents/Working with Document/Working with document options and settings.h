#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BorderType.h>
#include <Aspose.Words.Cpp/CleanupOptions.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Loading/EditingLanguage.h>
#include <Aspose.Words.Cpp/Loading/LanguagePreferences.h>
#include <Aspose.Words.Cpp/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Orientation.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/PaperSize.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionLayoutMode.h>
#include <Aspose.Words.Cpp/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Settings/ViewOptions.h>
#include <Aspose.Words.Cpp/Settings/ViewType.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <drawing/color.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Settings;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class WorkingWithDocumentOptionsAndSettings : public DocsExamplesBase
{
public:
    void OptimizeForMsWord()
    {
        //ExStart:OptimizeForMsWord
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2016);

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.OptimizeForMsWord.docx");
        //ExEnd:OptimizeForMsWord
    }

    void ShowGrammaticalAndSpellingErrors()
    {
        //ExStart:ShowGrammaticalAndSpellingErrors
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->set_ShowGrammaticalErrors(true);
        doc->set_ShowSpellingErrors(true);

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.ShowGrammaticalAndSpellingErrors.docx");
        //ExEnd:ShowGrammaticalAndSpellingErrors
    }

    void CleanupUnusedStylesAndLists()
    {
        //ExStart:CleanupUnusedStylesandLists
        auto doc = MakeObject<Document>(MyDir + u"Unused styles.docx");

        // Combined with the built-in styles, the document now has eight styles.
        // A custom style is marked as "used" while there is any text within the document
        // formatted in that style. This means that the 4 styles we added are currently unused.
        std::cout << (String::Format(u"Count of styles before Cleanup: {0}\n", doc->get_Styles()->get_Count()) +
                      String::Format(u"Count of lists before Cleanup: {0}", doc->get_Lists()->get_Count()))
                  << std::endl;

        // Cleans unused styles and lists from the document depending on given CleanupOptions.
        auto cleanupOptions = MakeObject<CleanupOptions>();
        cleanupOptions->set_UnusedLists(false);
        cleanupOptions->set_UnusedStyles(true);
        doc->Cleanup(cleanupOptions);

        std::cout << (String::Format(u"Count of styles after Cleanup was decreased: {0}\n", doc->get_Styles()->get_Count()) +
                      String::Format(u"Count of lists after Cleanup is the same: {0}", doc->get_Lists()->get_Count()))
                  << std::endl;

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.CleanupUnusedStylesAndLists.docx");
        //ExEnd:CleanupUnusedStylesandLists
    }

    void CleanupDuplicateStyle()
    {
        //ExStart:CleanupDuplicateStyle
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Count of styles before Cleanup.
        std::cout << doc->get_Styles()->get_Count() << std::endl;

        // Cleans duplicate styles from the document.
        auto options = MakeObject<CleanupOptions>();
        options->set_DuplicateStyle(true);
        doc->Cleanup(options);

        // Count of styles after Cleanup was decreased.
        std::cout << doc->get_Styles()->get_Count() << std::endl;

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.CleanupDuplicateStyle.docx");
        //ExEnd:CleanupDuplicateStyle
    }

    void ViewOptions()
    {
        //ExStart:SetViewOption
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->get_ViewOptions()->set_ViewType(ViewType::PageLayout);
        doc->get_ViewOptions()->set_ZoomPercent(50);

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.ViewOptions.docx");
        //ExEnd:SetViewOption
    }

    void DocumentPageSetup()
    {
        //ExStart:DocumentPageSetup
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Set the layout mode for a section allowing to define the document grid behavior.
        // Note that the Document Grid tab becomes visible in the Page Setup dialog of MS Word
        // if any Asian language is defined as editing language.
        doc->get_FirstSection()->get_PageSetup()->set_LayoutMode(SectionLayoutMode::Grid);
        doc->get_FirstSection()->get_PageSetup()->set_CharactersPerLine(30);
        doc->get_FirstSection()->get_PageSetup()->set_LinesPerPage(10);

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.DocumentPageSetup.docx");
        //ExEnd:DocumentPageSetup
    }

    void AddJapaneseAsEditingLanguages()
    {
        //ExStart:AddJapaneseAsEditinglanguages
        auto loadOptions = MakeObject<LoadOptions>();

        // Set language preferences that will be used when document is loading.
        loadOptions->get_LanguagePreferences()->AddEditingLanguage(EditingLanguage::Japanese);
        //ExEnd:AddJapaneseAsEditinglanguages

        auto doc = MakeObject<Document>(MyDir + u"No default editing language.docx", loadOptions);

        int localeIdFarEast = doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast();
        std::cout << (localeIdFarEast == (int)EditingLanguage::Japanese
                          ? String(u"The document either has no any FarEast language set in defaults or it was set to Japanese originally.")
                          : String(u"The document default FarEast language was set to another than Japanese language originally, so it is not overridden."))
                  << std::endl;
    }

    void SetRussianAsDefaultEditingLanguage()
    {
        //ExStart:SetRussianAsDefaultEditingLanguage
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->get_LanguagePreferences()->set_DefaultEditingLanguage(EditingLanguage::Russian);

        auto doc = MakeObject<Document>(MyDir + u"No default editing language.docx", loadOptions);

        int localeId = doc->get_Styles()->get_DefaultFont()->get_LocaleId();
        std::cout << (localeId == (int)EditingLanguage::Russian
                          ? String(u"The document either has no any language set in defaults or it was set to Russian originally.")
                          : String(u"The document default language was set to another than Russian language originally, so it is not overridden."))
                  << std::endl;
        //ExEnd:SetRussianAsDefaultEditingLanguage
    }

    void PageSetupAndSectionFormatting()
    {
        //ExStart:PageSetupAndSectionFormatting
        //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_PageSetup()->set_Orientation(Orientation::Landscape);
        builder->get_PageSetup()->set_LeftMargin(50);
        builder->get_PageSetup()->set_PaperSize(PaperSize::Paper10x14);

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.PageSetupAndSectionFormatting.docx");
        //ExEnd:PageSetupAndSectionFormatting
    }

    void PageBorderProperties()
    {
        //ExStart:PageBorderProperties
        auto doc = MakeObject<Document>();

        auto pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->set_BorderAlwaysInFront(false);
        pageSetup->set_BorderDistanceFrom(PageBorderDistanceFrom::PageEdge);
        pageSetup->set_BorderAppliesTo(PageBorderAppliesTo::FirstPage);

        auto border = pageSetup->get_Borders()->idx_get(BorderType::Top);
        border->set_LineStyle(LineStyle::Single);
        border->set_LineWidth(30);
        border->set_Color(System::Drawing::Color::get_Blue());
        border->set_DistanceFromText(0);

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.PageBorderProperties.docx");
        //ExEnd:PageBorderProperties
    }

    void LineGridSectionLayoutMode()
    {
        //ExStart:LineGridSectionLayoutMode
        //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Enable pitching, and then use it to set the number of lines per page in this section.
        // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
        builder->get_PageSetup()->set_LayoutMode(SectionLayoutMode::LineGrid);
        builder->get_PageSetup()->set_LinesPerPage(15);

        builder->get_ParagraphFormat()->set_SnapToGrid(true);

        for (int i = 0; i < 30; i++)
            builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");

        doc->Save(ArtifactsDir + u"WorkingWithDocumentOptionsAndSettings.LinesPerPage.docx");
        //ExEnd:LineGridSectionLayoutMode
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
