#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/LanguagePreferences.h>
#include <Aspose.Words.Cpp/Model/Document/EditingLanguage.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{
    void AddJapaneseAsEditinglanguages(const System::String& inputDataDir)
    {
        // ExStart:AddJapaneseAsEditinglanguages
        // The path to the documents directory.
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();
        loadOptions->get_LanguagePreferences()->AddEditingLanguage(EditingLanguage::Japanese);

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"languagepreferences.docx", loadOptions);

        int32_t localeIdFarEast = doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast();
        if (localeIdFarEast == static_cast<int32_t>(EditingLanguage::Japanese))
        {
            std::cout << "The document either has no any FarEast language set in defaults or it was set to Japanese originally." << std::endl;
        }
        else
        {
            std::cout << "The document default FarEast language was set to another than Japanese language originally, so it is not overridden." << std::endl;
        }
        // ExEnd:AddJapaneseAsEditinglanguages
    }

    void SetRussianAsDefaultEditingLanguage(const System::String& inputDataDir)
    {
        // ExStart:SetRussianAsDefaultEditingLanguage
        // The path to the documents directory.
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();

        loadOptions->get_LanguagePreferences()->set_DefaultEditingLanguage(EditingLanguage::Russian);

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"languagepreferences.docx", loadOptions);

        int32_t localeId = doc->get_Styles()->get_DefaultFont()->get_LocaleId();
        if (localeId == static_cast<int32_t>(EditingLanguage::Russian))
        {
            std::cout << "The document either has no any language set in defaults or it was set to Russian originally." << std::endl;
        }
        else
        {
            std::cout << "The document default language was set to another than Russian language originally, so it is not overridden." << std::endl;
        }
        // ExEnd:SetRussianAsDefaultEditingLanguage
    }
}

void Setuplanguagepreferences()
{
    std::cout << "Setuplanguagepreferences example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    AddJapaneseAsEditinglanguages(inputDataDir);
    SetRussianAsDefaultEditingLanguage(inputDataDir);
    std::cout << "Setuplanguagepreferences example finished." << std::endl << std::endl;
}
