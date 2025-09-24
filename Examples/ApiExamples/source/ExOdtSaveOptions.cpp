// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOdtSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Settings/MeasurementUnits.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveMeasureUnit.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>


using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2371625970u, ::Aspose::Words::ApiExamples::ExOdtSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExOdtSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExOdtSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExOdtSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExOdtSaveOptions> ExOdtSaveOptions::s_instance;

} // namespace gtest_test

void ExOdtSaveOptions::Odt11Schema(bool exportToOdt11Specs)
{
    //ExStart
    //ExFor:OdtSaveOptions
    //ExFor:OdtSaveOptions.#ctor
    //ExFor:OdtSaveOptions.IsStrictSchema11
    //ExFor:RevisionOptions.MeasurementUnit
    //ExFor:MeasurementUnits
    //ExSummary:Shows how to make a saved document conform to an older ODT schema.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OdtSaveOptions>();
    saveOptions->set_MeasureUnit(Aspose::Words::Saving::OdtSaveMeasureUnit::Centimeters);
    saveOptions->set_IsStrictSchema11(exportToOdt11Specs);
    
    doc->Save(get_ArtifactsDir() + u"OdtSaveOptions.Odt11Schema.odt", saveOptions);
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OdtSaveOptions.Odt11Schema.odt");
    
    ASSERT_EQ(Aspose::Words::MeasurementUnits::Centimeters, doc->get_LayoutOptions()->get_RevisionOptions()->get_MeasurementUnit());
    
    if (exportToOdt11Specs)
    {
        ASSERT_EQ(2, doc->get_Range()->get_FormFields()->get_Count());
        ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormTextInput, doc->get_Range()->get_FormFields()->idx_get(0)->get_Type());
        ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormCheckBox, doc->get_Range()->get_FormFields()->idx_get(1)->get_Type());
    }
    else
    {
        ASSERT_EQ(3, doc->get_Range()->get_FormFields()->get_Count());
        ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormTextInput, doc->get_Range()->get_FormFields()->idx_get(0)->get_Type());
        ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormCheckBox, doc->get_Range()->get_FormFields()->idx_get(1)->get_Type());
        ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormDropDown, doc->get_Range()->get_FormFields()->idx_get(2)->get_Type());
    }
}

namespace gtest_test
{

using ExOdtSaveOptions_Odt11Schema_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExOdtSaveOptions::Odt11Schema)>::type;

struct ExOdtSaveOptions_Odt11Schema : public ExOdtSaveOptions, public Aspose::Words::ApiExamples::ExOdtSaveOptions, public ::testing::WithParamInterface<ExOdtSaveOptions_Odt11Schema_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExOdtSaveOptions_Odt11Schema, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Odt11Schema(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOdtSaveOptions_Odt11Schema, ::testing::ValuesIn(ExOdtSaveOptions_Odt11Schema::TestCases()));

} // namespace gtest_test

void ExOdtSaveOptions::Encrypt(Aspose::Words::SaveFormat saveFormat)
{
    //ExStart
    //ExFor:OdtSaveOptions.#ctor(SaveFormat)
    //ExFor:OdtSaveOptions.Password
    //ExFor:OdtSaveOptions.SaveFormat
    //ExSummary:Shows how to encrypt a saved ODT/OTT document with a password, and then load it using Aspose.Words.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // Create a new OdtSaveOptions, and pass either "SaveFormat.Odt",
    // or "SaveFormat.Ott" as the format to save the document in. 
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OdtSaveOptions>(saveFormat);
    saveOptions->set_Password(u"@sposeEncrypted_1145");
    
    System::String extensionString = Aspose::Words::FileFormatUtil::SaveFormatToExtension(saveFormat);
    
    // If we open this document with an appropriate editor,
    // it will prompt us for the password we specified in the SaveOptions object.
    doc->Save(get_ArtifactsDir() + u"OdtSaveOptions.Encrypt" + extensionString, saveOptions);
    
    System::SharedPtr<Aspose::Words::FileFormatInfo> docInfo = Aspose::Words::FileFormatUtil::DetectFileFormat(get_ArtifactsDir() + u"OdtSaveOptions.Encrypt" + extensionString);
    
    ASSERT_TRUE(docInfo->get_IsEncrypted());
    
    // If we wish to open or edit this document again using Aspose.Words,
    // we will have to provide a LoadOptions object with the correct password to the loading constructor.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OdtSaveOptions.Encrypt" + extensionString, System::MakeObject<Aspose::Words::Loading::LoadOptions>(u"@sposeEncrypted_1145"));
    
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExOdtSaveOptions_Encrypt_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExOdtSaveOptions::Encrypt)>::type;

struct ExOdtSaveOptions_Encrypt : public ExOdtSaveOptions, public Aspose::Words::ApiExamples::ExOdtSaveOptions, public ::testing::WithParamInterface<ExOdtSaveOptions_Encrypt_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Odt),
            std::make_tuple(Aspose::Words::SaveFormat::Ott),
        };
    }
};

TEST_P(ExOdtSaveOptions_Encrypt, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Encrypt(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOdtSaveOptions_Encrypt, ::testing::ValuesIn(ExOdtSaveOptions_Encrypt::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
