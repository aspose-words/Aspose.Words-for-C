// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOdtSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveMeasureUnit.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExOdtSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExOdtSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExOdtSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExOdtSaveOptions> ExOdtSaveOptions::s_instance;

} // namespace gtest_test

void ExOdtSaveOptions::MeasureUnit(bool doExportToOdt11Specs)
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
    saveOptions->set_MeasureUnit(Aspose::Words::Saving::OdtSaveMeasureUnit::Centimeters);
    saveOptions->set_IsStrictSchema11(doExportToOdt11Specs);

    doc->Save(ArtifactsDir + u"OdtSaveOptions.MeasureUnit.odt", saveOptions);
    //ExEnd
}

namespace gtest_test
{

using MeasureUnit_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::MeasureUnit)>::type;

struct MeasureUnitVP : public ExOdtSaveOptions, public ApiExamples::ExOdtSaveOptions, public ::testing::WithParamInterface<MeasureUnit_Args>
{
    static std::vector<MeasureUnit_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(MeasureUnitVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MeasureUnit(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExOdtSaveOptions, MeasureUnitVP, ::testing::ValuesIn(MeasureUnitVP::TestCases()));

} // namespace gtest_test

void ExOdtSaveOptions::Encrypt(Aspose::Words::SaveFormat saveFormat)
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
    SharedPtr<FileFormatInfo> docInfo = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"OdtSaveOptions.Encrypt" + FileFormatUtil::SaveFormatToExtension(saveFormat));
    ASSERT_TRUE(docInfo->get_IsEncrypted());
}

namespace gtest_test
{

using Encrypt_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::Encrypt)>::type;

struct EncryptVP : public ExOdtSaveOptions, public ApiExamples::ExOdtSaveOptions, public ::testing::WithParamInterface<Encrypt_Args>
{
    static std::vector<Encrypt_Args> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Odt),
            std::make_tuple(Aspose::Words::SaveFormat::Ott),
        };
    }
};

TEST_P(EncryptVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Encrypt(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExOdtSaveOptions, EncryptVP, ::testing::ValuesIn(EncryptVP::TestCases()));

} // namespace gtest_test

void ExOdtSaveOptions::WorkWithEncryptedDocument(Aspose::Words::SaveFormat saveFormat)
{
    //ExStart
    //ExFor:OdtSaveOptions.#ctor(String)
    //ExSummary:Shows how to load and change odt/ott encrypted document.
    auto doc = MakeObject<Document>(MyDir + u"Encrypted" + FileFormatUtil::SaveFormatToExtension(saveFormat), MakeObject<LoadOptions>(u"@sposeEncrypted_1145"));

    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->Writeln(u"Encrypted document after changes.");

    // Saving document using new instance of OdtSaveOptions
    doc->Save(ArtifactsDir + u"OdtSaveOptions.WorkWithEncryptedDocument" + FileFormatUtil::SaveFormatToExtension(saveFormat), MakeObject<OdtSaveOptions>(u"@sposeEncrypted_1145"));
    //ExEnd

    // Check that document is still encrypted with a password
    SharedPtr<FileFormatInfo> docInfo = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"OdtSaveOptions.WorkWithEncryptedDocument" + FileFormatUtil::SaveFormatToExtension(saveFormat));
    ASSERT_TRUE(docInfo->get_IsEncrypted());
}

namespace gtest_test
{

using WorkWithEncryptedDocument_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::WorkWithEncryptedDocument)>::type;

struct WorkWithEncryptedDocumentVP : public ExOdtSaveOptions, public ApiExamples::ExOdtSaveOptions, public ::testing::WithParamInterface<WorkWithEncryptedDocument_Args>
{
    static std::vector<WorkWithEncryptedDocument_Args> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Odt),
            std::make_tuple(Aspose::Words::SaveFormat::Ott),
        };
    }
};

TEST_P(WorkWithEncryptedDocumentVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithEncryptedDocument(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExOdtSaveOptions, WorkWithEncryptedDocumentVP, ::testing::ValuesIn(WorkWithEncryptedDocumentVP::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
