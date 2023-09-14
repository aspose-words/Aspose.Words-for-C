#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Saving/DocSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithDocSaveOptions : public DocsExamplesBase
{
public:
    void EncryptDocumentWithPassword()
    {
        //ExStart:EncryptDocumentWithPassword
        //GistId:b4e8a7baa7d3c08127f9a043487de21b
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Hello world!");

        auto saveOptions = MakeObject<DocSaveOptions>();
        saveOptions->set_Password(u"password");

        doc->Save(ArtifactsDir + u"WorkingWithDocSaveOptions.EncryptDocumentWithPassword.docx", saveOptions);
        //ExEnd:EncryptDocumentWithPassword
    }

    void DoNotCompressSmallMetafiles()
    {
        //ExStart:DoNotCompressSmallMetafiles
        auto doc = MakeObject<Document>(MyDir + u"Microsoft equation object.docx");

        auto saveOptions = MakeObject<DocSaveOptions>();
        saveOptions->set_AlwaysCompressMetafiles(false);

        doc->Save(ArtifactsDir + u"WorkingWithDocSaveOptions.NotCompressSmallMetafiles.docx", saveOptions);
        //ExEnd:DoNotCompressSmallMetafiles
    }

    void DoNotSavePictureBullet()
    {
        //ExStart:DoNotSavePictureBullet
        auto doc = MakeObject<Document>(MyDir + u"Image bullet points.docx");

        auto saveOptions = MakeObject<DocSaveOptions>();
        saveOptions->set_SavePictureBullet(false);

        doc->Save(ArtifactsDir + u"WorkingWithDocSaveOptions.DoNotSavePictureBullet.docx", saveOptions);
        //ExEnd:DoNotSavePictureBullet
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
