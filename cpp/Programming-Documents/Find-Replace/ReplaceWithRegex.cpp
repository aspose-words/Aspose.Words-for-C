#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>

using namespace System::Text::RegularExpressions;
using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    void RecognizeAndSubstitutionsWithinReplacementPatterns()
    {
        // ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
        // Create new document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Write some text.
        builder->Write(u"Jason give money to Paul.");

        System::SharedPtr<System::Text::RegularExpressions::Regex> regex = System::MakeObject<System::Text::RegularExpressions::Regex>(u"([A-z]+) give money to ([A-z]+)");

        // Replace text using substitutions.
        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
        options->set_UseSubstitutions(true);
        doc->get_Range()->Replace(regex, u"$2 take money from $1", options);
        // ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns

        std::cout << doc->GetText().ToUtf8String() << std::endl;
        // The output is: Paul take money from Jason.\f
    }

    void FindAndReplaceWithRegex(const System::String& inputDataDir, const System::String& outputDataDir)
    {
        // ExStart:ReplaceWithRegex
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();

        doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"[s|m]ad"), u"bad", options);

        const System::String outputPath = outputDataDir + u"FindAndReplaceWithRegex_out.doc";
        doc->Save(outputPath);
        // ExEnd:ReplaceWithRegex

        std::cout << "Text replaced with regex successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void ReplaceWithRegex()
{
    std::cout << "ReplaceWithRegex example started." << std::endl;

    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    RecognizeAndSubstitutionsWithinReplacementPatterns();
    FindAndReplaceWithRegex(inputDataDir, outputDataDir);

    std::cout << "ReplaceWithRegex example finished." << std::endl << std::endl;
}
