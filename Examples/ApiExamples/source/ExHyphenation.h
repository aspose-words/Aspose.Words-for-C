#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Hyphenation.h>
#include <Aspose.Words.Cpp/IHyphenationCallback.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/WarningSource.h>
#include <Aspose.Words.Cpp/WarningType.h>
#include <system/collections/dictionary.h>
#include <system/collections/ienumerable.h>
#include <system/enum_helpers.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace ApiExamples {

class ExHyphenation : public ApiExampleBase
{
public:
    void Dictionary()
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
        Hyphenation::RegisterDictionary(u"de-CH", MyDir + u"hyph_de_CH.dic");

        ASSERT_TRUE(Hyphenation::IsDictionaryRegistered(u"de-CH"));

        // Open a document containing text with a locale matching that of our dictionary,
        // and save it to a fixed-page save format. The text in that document will be hyphenated.
        auto doc = MakeObject<Document>(MyDir + u"German text.docx");

        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->LINQ_OfType<SharedPtr<Run>>()->LINQ_All(
            [](SharedPtr<Run> r) { return r->get_Font()->get_LocaleId() == MakeObject<System::Globalization::CultureInfo>(u"de-CH")->get_LCID(); }));

        doc->Save(ArtifactsDir + u"Hyphenation.Dictionary.Registered.pdf");

        // Re-load the document after un-registering the dictionary,
        // and save it to another PDF, which will not have hyphenated text.
        Hyphenation::UnregisterDictionary(u"de-CH");

        ASSERT_FALSE(Hyphenation::IsDictionaryRegistered(u"de-CH"));

        doc = MakeObject<Document>(MyDir + u"German text.docx");
        doc->Save(ArtifactsDir + u"Hyphenation.Dictionary.Unregistered.pdf");
        //ExEnd
    }

    //ExStart
    //ExFor:Hyphenation
    //ExFor:Hyphenation.Callback
    //ExFor:Hyphenation.RegisterDictionary(String, Stream)
    //ExFor:Hyphenation.RegisterDictionary(String, String)
    //ExFor:Hyphenation.WarningCallback
    //ExFor:IHyphenationCallback
    //ExFor:IHyphenationCallback.RequestDictionary(System.String)
    //ExSummary:Shows how to open and register a dictionary from a file.
    void RegisterDictionary()
    {
        // Set up a callback that tracks warnings that occur during hyphenation dictionary registration.
        auto warningInfoCollection = MakeObject<WarningInfoCollection>();
        Hyphenation::set_WarningCallback(warningInfoCollection);

        // Register an English (US) hyphenation dictionary by stream.
        SharedPtr<System::IO::Stream> dictionaryStream = MakeObject<System::IO::FileStream>(MyDir + u"hyph_en_US.dic", System::IO::FileMode::Open);
        Hyphenation::RegisterDictionary(u"en-US", dictionaryStream);

        ASSERT_EQ(0, warningInfoCollection->get_Count());

        // Open a document with a locale that Microsoft Word may not hyphenate on an English machine, such as German.
        auto doc = MakeObject<Document>(MyDir + u"German text.docx");

        // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code.
        // This callback will handle the automatic request for that dictionary.
        Hyphenation::set_Callback(MakeObject<ExHyphenation::CustomHyphenationDictionaryRegister>());

        // When we save the document, German hyphenation will take effect.
        doc->Save(ArtifactsDir + u"Hyphenation.RegisterDictionary.pdf");

        // This dictionary contains two identical patterns, which will trigger a warning.
        ASSERT_EQ(1, warningInfoCollection->get_Count());
        ASSERT_EQ(WarningType::MinorFormattingLoss, warningInfoCollection->idx_get(0)->get_WarningType());
        ASSERT_EQ(WarningSource::Layout, warningInfoCollection->idx_get(0)->get_Source());
        ASSERT_EQ(String(u"Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. ") +
                      u"Content can be wrapped differently.",
                  warningInfoCollection->idx_get(0)->get_Description());
    }

    /// <summary>
    /// Associates ISO language codes with local system filenames for hyphenation dictionary files.
    /// </summary>
    class CustomHyphenationDictionaryRegister : public IHyphenationCallback
    {
    public:
        CustomHyphenationDictionaryRegister()
        {
            mHyphenationDictionaryFiles = MakeObject<System::Collections::Generic::Dictionary<String, String>>();
            mHyphenationDictionaryFiles->Add(u"en-US", MyDir + u"hyph_en_US.dic");
            mHyphenationDictionaryFiles->Add(u"de-CH", MyDir + u"hyph_de_CH.dic");
        }

        void RequestDictionary(String language) override
        {
            std::cout << (String(u"Hyphenation dictionary requested: ") + language);

            if (Hyphenation::IsDictionaryRegistered(language))
            {
                std::cout << ", is already registered." << std::endl;
                return;
            }

            if (mHyphenationDictionaryFiles->ContainsKey(language))
            {
                Hyphenation::RegisterDictionary(language, mHyphenationDictionaryFiles->idx_get(language));
                std::cout << ", successfully registered." << std::endl;
                return;
            }

            std::cout << ", no respective dictionary file known by this Callback." << std::endl;
        }

    private:
        SharedPtr<System::Collections::Generic::Dictionary<String, String>> mHyphenationDictionaryFiles;
    };
    //ExEnd
};

} // namespace ApiExamples
