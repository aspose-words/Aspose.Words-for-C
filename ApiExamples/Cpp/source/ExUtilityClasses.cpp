// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExUtilityClasses.h"

namespace ApiExamples { namespace gtest_test {

class ExUtilityClasses : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExUtilityClasses> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExUtilityClasses>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExUtilityClasses> ExUtilityClasses::s_instance;

TEST_F(ExUtilityClasses, PointsAndInches)
{
    s_instance->PointsAndInches();
}

TEST_F(ExUtilityClasses, PointsAndMillimeters)
{
    s_instance->PointsAndMillimeters();
}

TEST_F(ExUtilityClasses, PointsAndPixels)
{
    s_instance->PointsAndPixels();
}

TEST_F(ExUtilityClasses, PointsAndPixelsDpi)
{
    s_instance->PointsAndPixelsDpi();
}

}} // namespace ApiExamples::gtest_test
