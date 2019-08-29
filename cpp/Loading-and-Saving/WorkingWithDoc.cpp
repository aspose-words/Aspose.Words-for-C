#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/DocSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void AlwaysCompressMetafiles(System::String const &dataDir)
    {
        //ExStart:AlwaysCompressMetafiles
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
        System::SharedPtr<DocSaveOptions> saveOptions = System::MakeObject<DocSaveOptions>();
        saveOptions->set_AlwaysCompressMetafiles(false);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithDoc.AlwaysCompressMetafiles.doc");
        doc->Save(outputPath, saveOptions);
        //ExEnd:AlwaysCompressMetafiles
        std::cout << "The document is saved with AlwaysCompressMetafiles setting to false. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithDoc()
{
    std::cout << "WorkingWithDoc example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();

    AlwaysCompressMetafiles(dataDir);
    std::cout << "WorkingWithTxt example finished." << std::endl << std::endl;
}