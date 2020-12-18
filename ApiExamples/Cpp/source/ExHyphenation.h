#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Layout/Hyphenation/Hyphenation.h>
#include <Aspose.Words.Cpp/Layout/Hyphenation/IHyphenationCallback.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/collections/dictionary.h>
#include <system/enum_helpers.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
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
        //ExSummary:Shows how to perform and verify hyphenation dictionary registration.
        // Register a dictionary file from the local file system to the "de-CH" locale
        Hyphenation::RegisterDictionary(u"de-CH", MyDir + u"hyph_de_CH.dic");

        // This method can be used to verify that a language has a matching registered hyphenation dictionary
        ASSERT_TRUE(Hyphenation::IsDictionaryRegistered(u"de-CH"));

        // The dictionary file contains a long list of words in a specified language, which is German in this case
        // These words define a set of rules for hyphenating text (splitting words across lines)
        // If we open a document with text of a language matching that of a registered dictionary,
        // that dictionary's hyphenation rules will be applied and visible upon saving
        auto doc = MakeObject<Document>(MyDir + u"German text.docx");
        doc->Save(ArtifactsDir + u"Hyphenation.Dictionary.Registered.pdf");

        // We can also un-register a dictionary to disable these effects on any documents opened after the operation
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
        // Set up a callback that tracks warnings that occur during hyphenation dictionary registration
        auto warningInfoCollection = MakeObject<WarningInfoCollection>();
        Hyphenation::set_WarningCallback(warningInfoCollection);

        // Register an English (US) hyphenation dictionary by stream
        SharedPtr<System::IO::Stream> dictionaryStream =
            MakeObject<System::IO::FileStream>(MyDir + u"hyph_en_US.dic", System::IO::FileMode::Open);
        Hyphenation::RegisterDictionary(u"en-US", dictionaryStream);

        // No warnings detected
        ASSERT_EQ(0, warningInfoCollection->get_Count());

        // Open a document with a German locale that might not get automatically hyphenated by Microsoft Word on an English machine
        auto doc = MakeObject<Document>(MyDir + u"German text.docx");

        // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code
        // This callback will handle the automatic request for that dictionary
        Hyphenation::set_Callback(MakeObject<ExHyphenation::CustomHyphenationDictionaryRegister>());

        // When we save the document, it will be hyphenated according to rules defined by the dictionary known by our callback
        doc->Save(ArtifactsDir + u"Hyphenation.RegisterDictionary.pdf");

        // This dictionary contains two identical patterns, which will trigger a warning
        ASSERT_EQ(1, warningInfoCollection->get_Count());
        ASSERT_EQ(WarningType::MinorFormattingLoss, warningInfoCollection->idx_get(0)->get_WarningType());
        ASSERT_EQ(WarningSource::Layout, warningInfoCollection->idx_get(0)->get_Source());
        ASSERT_EQ(String(u"Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. ") +
                      u"Content can be wrapped differently.",
                  warningInfoCollection->idx_get(0)->get_Description());
    }

private:
    /// <summary>
    /// Associates ISO language codes with custom local system dictionary files for their respective languages
    /// </summary>
    class CustomHyphenationDictionaryRegister : public IHyphenationCallback
    {
    public:
        CustomHyphenationDictionaryRegister()
        {
            mHyphenationDictionaryFiles = [&]
            {
                auto dictonary_0 = MakeObject<System::Collections::Generic::Dictionary<String, String>>();
                dictonary_0->data()[u"en-US"] = MyDir + u"hyph_en_US.dic";
                dictonary_0->data()[u"de-CH"] = MyDir + u"hyph_de_CH.dic";
                return dictonary_0;
            }();
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
