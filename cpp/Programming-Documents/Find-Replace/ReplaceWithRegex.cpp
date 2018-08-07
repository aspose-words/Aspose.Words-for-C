#include "stdafx.h"
#include "examples.h"

#include <system/text/regularexpressions/regex.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Range.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/Document/Document.h>

using namespace System::Text::RegularExpressions;
using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

void ReplaceWithRegex()
{
    std::cout << "ReplaceWithRegex example started." << std::endl;
    // ExStart:ReplaceWithRegex
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();

    doc->get_Range()->Replace(System::MakeObject<Regex>(u"[s|m]ad"), u"bad", options);

    System::String outputPath = dataDir + GetOutputFilePath(u"ReplaceWithRegex.doc");
    doc->Save(outputPath);
    // ExEnd:ReplaceWithRegex
    std::cout << "Text replaced with regex successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceWithRegex example finished." << std::endl << std::endl;
}
