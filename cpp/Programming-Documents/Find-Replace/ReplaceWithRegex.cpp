#include "stdafx.h"
#include "examples.h"

#include <Model/Text/Range.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/Document/Document.h>

using namespace System::Text::RegularExpressions;
using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

void ReplaceWithRegex()
{
    std::cout << "ReplaceWithRegex example started." << std::endl;
    // ExStart:ReplaceWithRegex
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();

    doc->get_Range()->Replace(System::MakeObject<Regex>(u"[s|m]ad"), u"bad", options);

    System::String outputPath = outputDataDir + u"ReplaceWithRegex.doc";
    doc->Save(outputPath);
    // ExEnd:ReplaceWithRegex
    std::cout << "Text replaced with regex successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceWithRegex example finished." << std::endl << std::endl;
}
