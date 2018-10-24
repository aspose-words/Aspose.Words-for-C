#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <Model/Text/Range.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/FindReplace/FindReplaceDirection.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

void ReplaceWithString()
{
    std::cout << "ReplaceWithString example started." << std::endl;
    // ExStart:ReplaceWithString
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    doc->get_Range()->Replace(u"sad", u"bad", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

    System::String outputPath = dataDir + GetOutputFilePath(u"ReplaceWithString.doc");
    doc->Save(outputPath);
    // ExEnd:ReplaceWithString
    std::cout << "Text replaced with string successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceWithString example finished." << std::endl << std::endl;
}
