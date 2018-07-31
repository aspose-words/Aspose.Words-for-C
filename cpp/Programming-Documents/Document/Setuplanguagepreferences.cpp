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
    loadOptions->get_LanguagePreferences()->AddEditingLanguage(Aspose::Words::EditingLanguage::Japanese);
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"languagepreferences.doc", loadOptions);
    
    int32_t localeIdFarEast = doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast();
    if (localeIdFarEast == (int32_t)Aspose::Words::EditingLanguage::Japanese)
    {
        std::cout << "The document either has no any FarEast language set in defaults or it was set to Japanese originally.\n";
    }
    else
    {
        std::cout << "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.\n";
    }
    // ExEnd:AddJapaneseAsEditinglanguages
}

void SetRussianAsDefaultEditingLanguage(const System::String& dataDir)
{
    // ExStart:SetRussianAsDefaultEditingLanguage
    // The path to the documents directory.
    System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();
    
    loadOptions->get_LanguagePreferences()->SetAsDefault(Aspose::Words::EditingLanguage::Russian);
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"languagepreferences.doc", loadOptions);
    
    int32_t localeId = doc->get_Styles()->get_DefaultFont()->get_LocaleId();
    if (localeId == (int32_t)Aspose::Words::EditingLanguage::Russian)
    {
        std::cout << "The document either has no any language set in defaults or it was set to Russian originally.\n";
    }
    else
    {
        std::cout << "The document default language was set to another than Russian language originally, so it is not overridden.\n";
    }
    // ExEnd:SetRussianAsDefaultEditingLanguage
}

}

void Setuplanguagepreferences()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    AddJapaneseAsEditinglanguages(dataDir);
    SetRussianAsDefaultEditingLanguage(dataDir);
}
