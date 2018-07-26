#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Settings/MsWordVersion.h>
#include <Model/Settings/CompatibilityOptions.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{

void OptimizeFor(System::String dataDir)
{
    // ExStart:OptimizeFor
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"OtherTestFile.doc");
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2016);
    
    dataDir = dataDir + GetOutputFilePath(u"SetCompatibilityOptions.doc");
    
    // Save the document to disk.
    doc->Save(dataDir);
    // ExEnd:OptimizeFor
    std::cout << "\nDocument is optimized for MS Word 2016 successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

}

void SetCompatibilityOptions()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    OptimizeFor(dataDir);
}
