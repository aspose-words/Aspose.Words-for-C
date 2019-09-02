#include "stdafx.h"
#include "examples.h"

#include <Model/Text/Range.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/FindReplace/FindReplaceDirection.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

void ReplaceWithString()
{
    std::cout << "ReplaceWithString example started." << std::endl;
    // ExStart:ReplaceWithString
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    doc->get_Range()->Replace(u"sad", u"bad", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

    System::String outputPath = outputDataDir + u"ReplaceWithString.doc";
    doc->Save(outputPath);
    // ExEnd:ReplaceWithString
    std::cout << "Text replaced with string successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceWithString example finished." << std::endl << std::endl;
}
