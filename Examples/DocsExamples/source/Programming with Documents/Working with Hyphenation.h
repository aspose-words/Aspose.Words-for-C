#pragma once

#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Hyphenation.h>
#include <Aspose.Words.Cpp/IHyphenationCallback.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/path.h>
#include <system/io/stream.h>
#include <system/scope_guard.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithHyphenation : public DocsExamplesBase
{
public:
    class CustomHyphenationCallback : public IHyphenationCallback
    {
    public:
        void RequestDictionary(String language) override
        {
            String dictionaryFolder = MyDir;
            String dictionaryFullFileName;
            if (language == u"en-US")
            {
                dictionaryFullFileName = System::IO::Path::Combine(dictionaryFolder, u"hyph_en_US.dic");
            }
            else if (language == u"de-CH")
            {
                dictionaryFullFileName = System::IO::Path::Combine(dictionaryFolder, u"hyph_de_CH.dic");
            }
            else
            {
                throw System::Exception(String::Format(u"Missing hyphenation dictionary for {0}.", language));
            }
            // Register dictionary for requested language.
            Hyphenation::RegisterDictionary(language, dictionaryFullFileName);
        }
    };

public:
    void HyphenateWordsOfLanguages()
    {
        //ExStart:HyphenateWordsOfLanguages
        auto doc = MakeObject<Document>(MyDir + u"German text.docx");

        Hyphenation::RegisterDictionary(u"en-US", MyDir + u"hyph_en_US.dic");
        Hyphenation::RegisterDictionary(u"de-CH", MyDir + u"hyph_de_CH.dic");

        doc->Save(ArtifactsDir + u"WorkingWithHyphenation.HyphenateWordsOfLanguages.pdf");
        //ExEnd:HyphenateWordsOfLanguages
    }

    void LoadHyphenationDictionaryForLanguage()
    {
        //ExStart:LoadHyphenationDictionaryForLanguage
        auto doc = MakeObject<Document>(MyDir + u"German text.docx");

        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"hyph_de_CH.dic");
        Hyphenation::RegisterDictionary(u"de-CH", stream);

        doc->Save(ArtifactsDir + u"WorkingWithHyphenation.LoadHyphenationDictionaryForLanguage.pdf");
        //ExEnd:LoadHyphenationDictionaryForLanguage
    }

    void HyphenationCallback()
    {

        {
            auto __finally_guard_0 = ::System::MakeScopeGuard([]() { Hyphenation::set_Callback(nullptr); });

            try
            {
                // Register hyphenation callback.
                Hyphenation::set_Callback(MakeObject<WorkingWithHyphenation::CustomHyphenationCallback>());

                auto document = MakeObject<Document>(MyDir + u"German text.docx");
                document->Save(ArtifactsDir + u"WorkingWithHyphenation.HyphenationCallback.pdf");
            }
            catch (System::Exception& e)
            {
                ASSERT_TRUE(e->get_Message().StartsWith(u"Missing hyphenation dictionary"));
                std::cout << e->get_Message() << std::endl;
            }
        }
    }
};

}} // namespace DocsExamples::Programming_with_Documents
