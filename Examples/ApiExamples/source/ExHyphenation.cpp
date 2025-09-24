// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHyphenation.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/globalization/culture_info.h>
#include <system/func.h>
#include <system/enum_helpers.h>
#include <system/collections/ienumerable.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Layout/Hyphenation/Hyphenation.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1808460568u, ::Aspose::Words::ApiExamples::ExHyphenation::CustomHyphenationDictionaryRegister, ThisTypeBaseTypesInfo);

ExHyphenation::CustomHyphenationDictionaryRegister::CustomHyphenationDictionaryRegister()
{
    mHyphenationDictionaryFiles = [&]
    {
        auto dictionary_0 = System::MakeObject<System::Collections::Generic::Dictionary<System::String, System::String>>();
        dictionary_0->data()[u"en-US"] = get_MyDir() + u"hyph_en_US.dic";
        dictionary_0->data()[u"de-CH"] = get_MyDir() + u"hyph_de_CH.dic";
        return dictionary_0;
    }();
}

void ExHyphenation::CustomHyphenationDictionaryRegister::RequestDictionary(System::String language)
{
    std::cout << (System::String(u"Hyphenation dictionary requested: ") + language);
    
    if (Aspose::Words::Hyphenation::IsDictionaryRegistered(language))
    {
        std::cout << ", is already registered." << std::endl;
        return;
    }
    
    if (mHyphenationDictionaryFiles->ContainsKey(language))
    {
        Aspose::Words::Hyphenation::RegisterDictionary(language, mHyphenationDictionaryFiles->idx_get(language));
        std::cout << ", successfully registered." << std::endl;
        return;
    }
    
    std::cout << ", no respective dictionary file known by this Callback." << std::endl;
}


RTTI_INFO_IMPL_HASH(559853875u, ::Aspose::Words::ApiExamples::ExHyphenation, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExHyphenation : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExHyphenation> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExHyphenation>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExHyphenation> ExHyphenation::s_instance;

} // namespace gtest_test

void ExHyphenation::Dictionary()
{
    //ExStart
    //ExFor:Hyphenation.IsDictionaryRegistered(String)
    //ExFor:Hyphenation.RegisterDictionary(String, String)
    //ExFor:Hyphenation.UnregisterDictionary(String)
    //ExSummary:Shows how to register a hyphenation dictionary.
    // A hyphenation dictionary contains a list of strings that define hyphenation rules for the dictionary's language.
    // When a document contains lines of text in which a word could be split up and continued on the next line,
    // hyphenation will look through the dictionary's list of strings for that word's substrings.
    // If the dictionary contains a substring, then hyphenation will split the word across two lines
    // by the substring and add a hyphen to the first half.
    // Register a dictionary file from the local file system to the "de-CH" locale.
    Aspose::Words::Hyphenation::RegisterDictionary(u"de-CH", get_MyDir() + u"hyph_de_CH.dic");
    
    ASSERT_TRUE(Aspose::Words::Hyphenation::IsDictionaryRegistered(u"de-CH"));
    
    // Open a document containing text with a locale matching that of our dictionary,
    // and save it to a fixed-page save format. The text in that document will be hyphenated.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"German text.docx");
    
    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->LINQ_OfType<System::SharedPtr<Aspose::Words::Run> >()->LINQ_All(static_cast<System::Func<System::SharedPtr<Aspose::Words::Run>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Run> r)>>([](System::SharedPtr<Aspose::Words::Run> r) -> bool
    {
        return r->get_Font()->get_LocaleId() == System::MakeObject<System::Globalization::CultureInfo>(u"de-CH")->get_LCID();
    }))));
    
    doc->Save(get_ArtifactsDir() + u"Hyphenation.Dictionary.Registered.pdf");
    
    // Re-load the document after un-registering the dictionary,
    // and save it to another PDF, which will not have hyphenated text.
    Aspose::Words::Hyphenation::UnregisterDictionary(u"de-CH");
    
    ASSERT_FALSE(Aspose::Words::Hyphenation::IsDictionaryRegistered(u"de-CH"));
    
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"German text.docx");
    doc->Save(get_ArtifactsDir() + u"Hyphenation.Dictionary.Unregistered.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHyphenation, Dictionary)
{
    s_instance->Dictionary();
}

} // namespace gtest_test

void ExHyphenation::RegisterDictionary()
{
    // Set up a callback that tracks warnings that occur during hyphenation dictionary registration.
    auto warningInfoCollection = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    Aspose::Words::Hyphenation::set_WarningCallback(warningInfoCollection);
    
    // Register an English (US) hyphenation dictionary by stream.
    System::SharedPtr<System::IO::Stream> dictionaryStream = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"hyph_en_US.dic", System::IO::FileMode::Open);
    Aspose::Words::Hyphenation::RegisterDictionary(u"en-US", dictionaryStream);
    
    ASSERT_EQ(0, warningInfoCollection->get_Count());
    
    // Open a document with a locale that Microsoft Word may not hyphenate on an English machine, such as German.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"German text.docx");
    
    // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code.
    // This callback will handle the automatic request for that dictionary.
    Aspose::Words::Hyphenation::set_Callback(System::MakeObject<Aspose::Words::ApiExamples::ExHyphenation::CustomHyphenationDictionaryRegister>());
    
    // When we save the document, German hyphenation will take effect.
    doc->Save(get_ArtifactsDir() + u"Hyphenation.RegisterDictionary.pdf");
    
    // This dictionary contains two identical patterns, which will trigger a warning.
    ASSERT_EQ(1, warningInfoCollection->get_Count());
    ASSERT_EQ(Aspose::Words::WarningType::MinorFormattingLoss, warningInfoCollection->idx_get(0)->get_WarningType());
    ASSERT_EQ(Aspose::Words::WarningSource::Layout, warningInfoCollection->idx_get(0)->get_Source());
    ASSERT_EQ(System::String(u"Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. ") + u"Content can be wrapped differently.", warningInfoCollection->idx_get(0)->get_Description());
    
    Aspose::Words::Hyphenation::set_WarningCallback(nullptr);
    //ExSkip
    Aspose::Words::Hyphenation::UnregisterDictionary(u"en-US");
    //ExSkip
    Aspose::Words::Hyphenation::set_Callback(nullptr);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExHyphenation, RegisterDictionary)
{
    s_instance->RegisterDictionary();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
