#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Saving/DocSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/date_time.h>
#include <system/io/directory.h>
#include <system/io/file_info.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExDocSaveOptions : public ApiExampleBase
{
public:
    void SaveAsDoc()
    {
        //ExStart
        //ExFor:DocSaveOptions
        //ExFor:DocSaveOptions.#ctor()
        //ExFor:DocSaveOptions.#ctor(SaveFormat)
        //ExFor:DocSaveOptions.Password
        //ExFor:DocSaveOptions.SaveFormat
        //ExFor:DocSaveOptions.SaveRoutingSlip
        //ExSummary:Shows how to set save options for older Microsoft Word formats.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Hello world!");

        auto options = MakeObject<DocSaveOptions>(SaveFormat::Doc);

        // Set a password which will protect the loading of the document by Microsoft Word or Aspose.Words.
        // Note that this does not encrypt the contents of the document in any way.
        options->set_Password(u"MyPassword");

        // If the document contains a routing slip, we can preserve it while saving by setting this flag to true.
        options->set_SaveRoutingSlip(true);

        doc->Save(ArtifactsDir + u"DocSaveOptions.SaveAsDoc.doc", options);

        // To be able to load the document,
        // we will need to apply the password we specified in the DocSaveOptions object in a LoadOptions object.
        ASSERT_THROW(doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.SaveAsDoc.doc"), IncorrectPasswordException);

        auto loadOptions = MakeObject<LoadOptions>(u"MyPassword");
        doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.SaveAsDoc.doc", loadOptions);

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        //ExEnd
    }

    void TempFolder()
    {
        //ExStart
        //ExFor:SaveOptions.TempFolder
        //ExSummary:Shows how to use the hard drive instead of memory when saving a document.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // When we save a document, various elements are temporarily stored in memory as the save operation is taking place.
        // We can use this option to use a temporary folder in the local file system instead,
        // which will reduce our application's memory overhead.
        auto options = MakeObject<DocSaveOptions>();
        options->set_TempFolder(ArtifactsDir + u"TempFiles");

        // The specified temporary folder must exist in the local file system before the save operation.
        System::IO::Directory::CreateDirectory_(options->get_TempFolder());

        doc->Save(ArtifactsDir + u"DocSaveOptions.TempFolder.doc", options);

        // The folder will persist with no residual contents from the load operation.
        ASSERT_EQ(0, System::IO::Directory::GetFiles(options->get_TempFolder())->get_Length());
        //ExEnd
    }

    void PictureBullets()
    {
        //ExStart
        //ExFor:DocSaveOptions.SavePictureBullet
        //ExSummary:Shows how to omit PictureBullet data from the document when saving.
        auto doc = MakeObject<Document>(MyDir + u"Image bullet points.docx");
        ASSERT_FALSE(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData() == nullptr);
        //ExSkip

        // Some word processors, such as Microsoft Word 97, are incompatible with PictureBullet data.
        // By setting a flag in the SaveOptions object,
        // we can convert all image bullet points to ordinary bullet points while saving.
        auto saveOptions = MakeObject<DocSaveOptions>(SaveFormat::Doc);
        saveOptions->set_SavePictureBullet(false);

        doc->Save(ArtifactsDir + u"DocSaveOptions.PictureBullets.doc", saveOptions);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.PictureBullets.doc");

        ASSERT_TRUE(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData() == nullptr);
    }

    void UpdateLastPrintedProperty(bool isUpdateLastPrintedProperty)
    {
        //ExStart
        //ExFor:SaveOptions.UpdateLastPrintedProperty
        //ExSummary:Shows how to update a document's "Last printed" property when saving.
        auto doc = MakeObject<Document>();
        doc->get_BuiltInDocumentProperties()->set_LastPrinted(System::DateTime(2019, 12, 20));

        // This flag determines whether the last printed date, which is a built-in property, is updated.
        // If so, then the date of the document's most recent save operation
        // with this SaveOptions object passed as a parameter is used as the print date.
        auto saveOptions = MakeObject<DocSaveOptions>();
        saveOptions->set_UpdateLastPrintedProperty(isUpdateLastPrintedProperty);

        // In Microsoft Word 2003, this property can be found via File -> Properties -> Statistics -> Printed.
        // It can also be displayed in the document's body by using a PRINTDATE field.
        doc->Save(ArtifactsDir + u"DocSaveOptions.UpdateLastPrintedProperty.doc", saveOptions);

        // Open the saved document, then verify the value of the property.
        doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.UpdateLastPrintedProperty.doc");

        ASPOSE_ASSERT_NE(isUpdateLastPrintedProperty, System::DateTime(2019, 12, 20) == doc->get_BuiltInDocumentProperties()->get_LastPrinted());
        //ExEnd
    }

    void UpdateCreatedTimeProperty(bool isUpdateCreatedTimeProperty)
    {
        //ExStart
        //ExFor:SaveOptions.UpdateLastPrintedProperty
        //ExSummary:Shows how to update a document's "CreatedTime" property when saving.
        auto doc = MakeObject<Document>();
        doc->get_BuiltInDocumentProperties()->set_CreatedTime(System::DateTime(2019, 12, 20));

        // This flag determines whether the created time, which is a built-in property, is updated.
        // If so, then the date of the document's most recent save operation
        // with this SaveOptions object passed as a parameter is used as the created time.
        auto saveOptions = MakeObject<DocSaveOptions>();
        saveOptions->set_UpdateCreatedTimeProperty(isUpdateCreatedTimeProperty);

        doc->Save(ArtifactsDir + u"DocSaveOptions.UpdateCreatedTimeProperty.docx", saveOptions);

        // Open the saved document, then verify the value of the property.
        doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.UpdateCreatedTimeProperty.docx");

        ASPOSE_ASSERT_NE(isUpdateCreatedTimeProperty, System::DateTime(2019, 12, 20) == doc->get_BuiltInDocumentProperties()->get_CreatedTime());
        //ExEnd
    }

    void AlwaysCompressMetafiles(bool compressAllMetafiles)
    {
        //ExStart
        //ExFor:DocSaveOptions.AlwaysCompressMetafiles
        //ExSummary:Shows how to change metafiles compression in a document while saving.
        // Open a document that contains a Microsoft Equation 3.0 formula.
        auto doc = MakeObject<Document>(MyDir + u"Microsoft equation object.docx");

        // When we save a document, smaller metafiles are not compressed for performance reasons.
        // We can set a flag in a SaveOptions object to compress every metafile when saving.
        // Some editors such as LibreOffice cannot read uncompressed metafiles.
        auto saveOptions = MakeObject<DocSaveOptions>();
        saveOptions->set_AlwaysCompressMetafiles(compressAllMetafiles);

        doc->Save(ArtifactsDir + u"DocSaveOptions.AlwaysCompressMetafiles.docx", saveOptions);

        if (compressAllMetafiles)
        {
            ASSERT_LT(10000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"DocSaveOptions.AlwaysCompressMetafiles.docx")->get_Length());
        }
        else
        {
            ASSERT_GE(30000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"DocSaveOptions.AlwaysCompressMetafiles.docx")->get_Length());
        }
        //ExEnd
    }
};

} // namespace ApiExamples
