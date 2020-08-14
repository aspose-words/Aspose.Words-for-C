#include "TestData/TestBuilders/NumericTestBuilder.h"

#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/nullable.h>
#include <system/date_time.h>
#include <cstdint>

#include "TestData/TestClasses/NumericTestClass.h"

using namespace ApiExamples::TestData::TestClasses;
namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

RTTI_INFO_IMPL_HASH(577368709u, ::ApiExamples::TestData::TestBuilders::NumericTestBuilder, ThisTypeBaseTypesInfo);

NumericTestBuilder::NumericTestBuilder() : mValue2(0), mValue3(0), mLogical(false)
{
    mValue1 = 1;
    mValue2 = 1.0;
    mValue3 = 1;
    mValue4 = 1;
    mLogical = false;
    mDate = System::DateTime(2018, 1, 1);
}

System::SharedPtr<NumericTestBuilder> NumericTestBuilder::WithValuesAndDate(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, System::DateTime dateTime)
{
    mValue1 = value1;
    mValue2 = value2;
    mValue3 = value3;
    mValue4 = value4;
    mDate = dateTime;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<NumericTestBuilder> NumericTestBuilder::WithValuesAndLogical(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, bool logical)
{
    mValue1 = value1;
    mValue2 = value2;
    mValue3 = value3;
    mValue4 = value4;
    mLogical = logical;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ApiExamples::TestData::TestClasses::NumericTestClass> NumericTestBuilder::Build()
{
    return System::MakeObject<NumericTestClass>(mValue1, mValue2, mValue3, mValue4, mLogical, mDate);
}

NumericTestBuilder::~NumericTestBuilder()
{
}

System::Object::shared_members_type ApiExamples::TestData::TestBuilders::NumericTestBuilder::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::TestData::TestBuilders::NumericTestBuilder::mValue1", this->mValue1);
    result.Add("ApiExamples::TestData::TestBuilders::NumericTestBuilder::mValue4", this->mValue4);
    result.Add("ApiExamples::TestData::TestBuilders::NumericTestBuilder::mDate", this->mDate);

    return result;
}

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
