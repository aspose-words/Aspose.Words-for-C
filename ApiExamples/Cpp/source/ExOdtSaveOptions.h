#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveMeasureUnit.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExOdtSaveOptions : public ApiExampleBase
{
public:
    void MeasureUnit(bool doExportToOdt11Specs)
    {
        //ExStart
        //ExFor:OdtSaveOptions
        //ExFor:OdtSaveOptions.#ctor()
        //ExFor:OdtSaveOptions.IsStrictSchema11
        //ExFor:OdtSaveOptions.MeasureUnit
        //ExFor:OdtSaveMeasureUnit
        //ExSummary:Shows how to work with units of measure of document content.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Open Office uses centimeters, MS Office uses inches
        auto saveOptions = MakeObject<OdtSaveOptions>();
        saveOptions->set_MeasureUnit(OdtSaveMeasureUnit::Centimeters);
        saveOptions->set_IsStrictSchema11(doExportToOdt11Specs);

        doc->Save(ArtifactsDir + u"OdtSaveOptions.MeasureUnit.odt", saveOptions);
        //ExEnd
    }

    void Encrypt(SaveFormat saveFormat)
    {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(SaveFormat)
        //ExFor:OdtSaveOptions.Password
        //ExFor:OdtSaveOptions.SaveFormat
        //ExSummary:Shows how to encrypted your odt/ott documents with a password.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<OdtSaveOptions>(saveFormat);
        saveOptions->set_Password(u"@sposeEncrypted_1145");

        // Saving document using password property of OdtSaveOptions
        doc->Save(ArtifactsDir + u"OdtSaveOptions.Encrypt" + FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
        //ExEnd

        // Check that all documents are encrypted with a password
        SharedPtr<FileFormatInfo> docInfo =
            FileFormatUtil::DetectFileFormat(ArtifactsDir + u"OdtSaveOptions.Encrypt" + FileFormatUtil::SaveFormatToExtension(saveFormat));
        ASSERT_TRUE(docInfo->get_IsEncrypted());
    }

    void WorkWithEncryptedDocument(SaveFormat saveFormat)
    {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(String)
        //ExSummary:Shows how to load and change odt/ott encrypted document.
        auto doc = MakeObject<Document>(MyDir + u"Encrypted" + FileFormatUtil::SaveFormatToExtension(saveFormat),
                                                MakeObject<LoadOptions>(u"@sposeEncrypted_1145"));

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->Writeln(u"Encrypted document after changes.");

        // Saving document using new instance of OdtSaveOptions
        doc->Save(ArtifactsDir + u"OdtSaveOptions.WorkWithEncryptedDocument" + FileFormatUtil::SaveFormatToExtension(saveFormat),
                  MakeObject<OdtSaveOptions>(u"@sposeEncrypted_1145"));
        //ExEnd

        // Check that document is still encrypted with a password
        SharedPtr<FileFormatInfo> docInfo =
            FileFormatUtil::DetectFileFormat(ArtifactsDir + u"OdtSaveOptions.WorkWithEncryptedDocument" + FileFormatUtil::SaveFormatToExtension(saveFormat));
        ASSERT_TRUE(docInfo->get_IsEncrypted());
    }
};

} // namespace ApiExamples
