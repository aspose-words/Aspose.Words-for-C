// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOdtSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
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

using ExOdtSaveOptions_Odt11Schema_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::Odt11Schema)>::type;

struct ExOdtSaveOptions_Odt11Schema : public ExOdtSaveOptions,
                                      public ApiExamples::ExOdtSaveOptions,
                                      public ::testing::WithParamInterface<ExOdtSaveOptions_Odt11Schema_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
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

using ExOdtSaveOptions_MeasurementUnits_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::MeasurementUnits)>::type;

struct ExOdtSaveOptions_MeasurementUnits : public ExOdtSaveOptions,
                                           public ApiExamples::ExOdtSaveOptions,
                                           public ::testing::WithParamInterface<ExOdtSaveOptions_MeasurementUnits_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(OdtSaveMeasureUnit::Centimeters),
            std::make_tuple(OdtSaveMeasureUnit::Inches),
        };
    }
};

TEST_P(ExOdtSaveOptions_MeasurementUnits, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MeasurementUnits(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOdtSaveOptions_MeasurementUnits, ::testing::ValuesIn(ExOdtSaveOptions_MeasurementUnits::TestCases()));

using ExOdtSaveOptions_Encrypt_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOdtSaveOptions::Encrypt)>::type;

struct ExOdtSaveOptions_Encrypt : public ExOdtSaveOptions,
                                  public ApiExamples::ExOdtSaveOptions,
                                  public ::testing::WithParamInterface<ExOdtSaveOptions_Encrypt_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Odt),
            std::make_tuple(SaveFormat::Ott),
        };
    }
};

TEST_P(ExOdtSaveOptions_Encrypt, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Encrypt(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOdtSaveOptions_Encrypt, ::testing::ValuesIn(ExOdtSaveOptions_Encrypt::TestCases()));

}} // namespace ApiExamples::gtest_test
