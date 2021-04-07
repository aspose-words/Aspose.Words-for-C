#pragma once

#include <Aspose.Words.Cpp/Layout/Hyphenation/Hyphenation.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>

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
};

}} // namespace DocsExamples::Programming_with_Documents
