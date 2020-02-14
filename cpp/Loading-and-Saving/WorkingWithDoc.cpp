#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/DocSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void EncryptDocumentWithPassword(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:EncryptDocumentWithPassword
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
        System::SharedPtr<DocSaveOptions> docSaveOptions = System::MakeObject<DocSaveOptions>();
        docSaveOptions->set_Password(u"password");
        System::String outputPath = outputDataDir + u"WorkingWithDoc.EncryptDocumentWithPassword.doc";
        doc->Save(outputPath, docSaveOptions);
        //ExEnd:EncryptDocumentWithPassword
        std::cout << "The password of document is set using RC4 encryption method. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void AlwaysCompressMetafiles(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:AlwaysCompressMetafiles
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
        System::SharedPtr<DocSaveOptions> saveOptions = System::MakeObject<DocSaveOptions>();
        saveOptions->set_AlwaysCompressMetafiles(false);
        System::String outputPath = outputDataDir + u"WorkingWithDoc.AlwaysCompressMetafiles.doc";
        doc->Save(outputPath, saveOptions);
        //ExEnd:AlwaysCompressMetafiles
        std::cout << "The document is saved with AlwaysCompressMetafiles setting to false. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithDoc()
{
    std::cout << "WorkingWithDoc example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    EncryptDocumentWithPassword(inputDataDir, outputDataDir);
    AlwaysCompressMetafiles(inputDataDir, outputDataDir);
    //SavePictureBullet(dataDir); // Source document is missing
    std::cout << "WorkingWithTxt example finished." << std::endl << std::endl;
}