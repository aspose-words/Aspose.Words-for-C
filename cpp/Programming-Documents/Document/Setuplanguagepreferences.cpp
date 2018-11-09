#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Text/Font.h>
#include <Model/Styles/StyleCollection.h>
#include <Model/Document/LoadOptions.h>
#include <Model/Document/LanguagePreferences.h>
#include <Model/Document/EditingLanguage.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;

namespace
{
    void AddJapaneseAsEditinglanguages(const System::String& dataDir)
    {
        // ExStart:AddJapaneseAsEditinglanguages
        // The path to the documents directory.
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();
        loadOptions->get_LanguagePreferences()->AddEditingLanguage(EditingLanguage::Japanese);

        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"languagepreferences.doc", loadOptions);

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

    void SetRussianAsDefaultEditingLanguage(const System::String& dataDir)
    {
        // ExStart:SetRussianAsDefaultEditingLanguage
        // The path to the documents directory.
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();

        loadOptions->get_LanguagePreferences()->SetAsDefault(EditingLanguage::Russian);

        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"languagepreferences.doc", loadOptions);

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
    // ExStart:Setuplanguagepreferences
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    AddJapaneseAsEditinglanguages(dataDir);
    SetRussianAsDefaultEditingLanguage(dataDir);
    // ExEnd:Setuplanguagepreferences
    std::cout << "Setuplanguagepreferences example finished." << std::endl << std::endl;
}
