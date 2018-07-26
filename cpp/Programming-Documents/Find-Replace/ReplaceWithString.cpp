#include "../../examples.h"

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
    // ExStart:ReplaceWithString
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    doc->get_Range()->Replace(u"sad", u"bad", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));
    
    dataDir = dataDir + GetOutputFilePath(u"ReplaceWithString.doc");
    doc->Save(dataDir);
    // ExEnd:ReplaceWithString
    std::cout << "\nText replaced with string successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
