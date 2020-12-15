// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOdtSaveOptions.h"

#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExOdtSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExOdtSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExOdtSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExOdtSaveOptions> ExOdtSaveOptions::s_instance;

using MeasureUnit_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::MeasureUnit)>::type;

struct MeasureUnitVP : public ExOdtSaveOptions, public ApiExamples::ExOdtSaveOptions, public ::testing::WithParamInterface<MeasureUnit_Args>
{
    static std::vector<MeasureUnit_Args> TestCases()
    {
        return {
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

using Encrypt_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::Encrypt)>::type;

struct EncryptVP : public ExOdtSaveOptions, public ApiExamples::ExOdtSaveOptions, public ::testing::WithParamInterface<Encrypt_Args>
{
    static std::vector<Encrypt_Args> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Odt),
            std::make_tuple(SaveFormat::Ott),
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

using WorkWithEncryptedDocument_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::WorkWithEncryptedDocument)>::type;

struct WorkWithEncryptedDocumentVP : public ExOdtSaveOptions,
                                     public ApiExamples::ExOdtSaveOptions,
                                     public ::testing::WithParamInterface<WorkWithEncryptedDocument_Args>
{
    static std::vector<WorkWithEncryptedDocument_Args> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Odt),
            std::make_tuple(SaveFormat::Ott),
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

}} // namespace ApiExamples::gtest_test
