// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestBuilders/NumericTestBuilder.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;
namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

RTTI_INFO_IMPL_HASH(236373439u, ::Aspose::Words::ApiExamples::TestData::TestBuilders::NumericTestBuilder, ThisTypeBaseTypesInfo);

NumericTestBuilder::NumericTestBuilder() : mValue2(0), mValue3(0), mLogical(false)
{
    mValue1 = 1;
    mValue2 = 1.0;
    mValue3 = 1;
    mValue4 = 1;
    mLogical = false;
    mDate = System::DateTime(2018, 1, 1);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::NumericTestBuilder> NumericTestBuilder::WithValuesAndDate(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, System::DateTime dateTime)
{
    mValue1 = value1;
    mValue2 = value2;
    mValue3 = value3;
    mValue4 = value4;
    mDate = dateTime;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::NumericTestBuilder> NumericTestBuilder::WithValuesAndLogical(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, bool logical)
{
    mValue1 = value1;
    mValue2 = value2;
    mValue3 = value3;
    mValue4 = value4;
    mLogical = logical;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::NumericTestBuilder> NumericTestBuilder::WithValues(System::Nullable<int32_t> value1, double value2)
{
    mValue1 = value1;
    mValue2 = value2;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::NumericTestClass> NumericTestBuilder::Build()
{
    return System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::NumericTestClass>(mValue1, mValue2, mValue3, mValue4, mLogical, mDate);
}

NumericTestBuilder::~NumericTestBuilder()
{
}

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
