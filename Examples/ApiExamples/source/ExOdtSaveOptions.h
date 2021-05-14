#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/FileFormatInfo.h>
#include <Aspose.Words.Cpp/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Layout/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/RevisionOptions.h>
#include <Aspose.Words.Cpp/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/MeasurementUnits.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/OdtSaveMeasureUnit.h>
#include <Aspose.Words.Cpp/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
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
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExOdtSaveOptions : public ApiExampleBase
{
public:
    void Odt11Schema(bool exportToOdt11Specs)
    {
        //ExStart
        //ExFor:OdtSaveOptions
        //ExFor:OdtSaveOptions.#ctor()
        //ExFor:OdtSaveOptions.IsStrictSchema11
        //ExSummary:Shows how to make a saved document conform to an older ODT schema.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<OdtSaveOptions>();
        saveOptions->set_MeasureUnit(OdtSaveMeasureUnit::Centimeters);
        saveOptions->set_IsStrictSchema11(exportToOdt11Specs);

        doc->Save(ArtifactsDir + u"OdtSaveOptions.Odt11Schema.odt", saveOptions);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"OdtSaveOptions.Odt11Schema.odt");

        ASSERT_EQ(Aspose::Words::MeasurementUnits::Centimeters, doc->get_LayoutOptions()->get_RevisionOptions()->get_MeasurementUnit());

        if (exportToOdt11Specs)
        {
            ASSERT_EQ(2, doc->get_Range()->get_FormFields()->get_Count());
            ASSERT_EQ(FieldType::FieldFormTextInput, doc->get_Range()->get_FormFields()->idx_get(0)->get_Type());
            ASSERT_EQ(FieldType::FieldFormCheckBox, doc->get_Range()->get_FormFields()->idx_get(1)->get_Type());
        }
        else
        {
            ASSERT_EQ(3, doc->get_Range()->get_FormFields()->get_Count());
            ASSERT_EQ(FieldType::FieldFormTextInput, doc->get_Range()->get_FormFields()->idx_get(0)->get_Type());
            ASSERT_EQ(FieldType::FieldFormCheckBox, doc->get_Range()->get_FormFields()->idx_get(1)->get_Type());
            ASSERT_EQ(FieldType::FieldFormDropDown, doc->get_Range()->get_FormFields()->idx_get(2)->get_Type());
        }
    }

    void MeasurementUnits(OdtSaveMeasureUnit odtSaveMeasureUnit)
    {
        //ExStart
        //ExFor:OdtSaveOptions
        //ExFor:OdtSaveOptions.MeasureUnit
        //ExFor:OdtSaveMeasureUnit
        //ExSummary:Shows how to use different measurement units to define style parameters of a saved ODT document.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // When we export the document to .odt, we can use an OdtSaveOptions object to modify how we save the document.
        // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Centimeters"
        // to define content such as style parameters using the metric system, which Open Office uses.
        // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Inches"
        // to define content such as style parameters using the imperial system, which Microsoft Word uses.
        auto saveOptions = MakeObject<OdtSaveOptions>();
        saveOptions->set_MeasureUnit(odtSaveMeasureUnit);

        doc->Save(ArtifactsDir + u"OdtSaveOptions.Odt11Schema.odt", saveOptions);
        //ExEnd
    }

    void Encrypt(SaveFormat saveFormat)
    {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(SaveFormat)
        //ExFor:OdtSaveOptions.Password
        //ExFor:OdtSaveOptions.SaveFormat
        //ExSummary:Shows how to encrypt a saved ODT/OTT document with a password, and then load it using Aspose.Words.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Create a new OdtSaveOptions, and pass either "SaveFormat.Odt",
        // or "SaveFormat.Ott" as the format to save the document in.
        auto saveOptions = MakeObject<OdtSaveOptions>(saveFormat);
        saveOptions->set_Password(u"@sposeEncrypted_1145");

        String extensionString = FileFormatUtil::SaveFormatToExtension(saveFormat);

        // If we open this document with an appropriate editor,
        // it will prompt us for the password we specified in the SaveOptions object.
        doc->Save(ArtifactsDir + u"OdtSaveOptions.Encrypt" + extensionString, saveOptions);

        SharedPtr<FileFormatInfo> docInfo = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"OdtSaveOptions.Encrypt" + extensionString);

        ASSERT_TRUE(docInfo->get_IsEncrypted());

        // If we wish to open or edit this document again using Aspose.Words,
        // we will have to provide a LoadOptions object with the correct password to the loading constructor.
        doc = MakeObject<Document>(ArtifactsDir + u"OdtSaveOptions.Encrypt" + extensionString, MakeObject<LoadOptions>(u"@sposeEncrypted_1145"));

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        //ExEnd
    }
};

} // namespace ApiExamples
