#include "TestData/TestClasses/NumericTestClass.h"

#include <system/scope_guard.h>
#include <system/nullable.h>
#include <system/date_time.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(2251799918u, ::ApiExamples::TestData::TestClasses::NumericTestClass, ThisTypeBaseTypesInfo);

System::Nullable<int32_t> NumericTestClass::get_Value1()
{
    return pr_Value1;
}

void NumericTestClass::set_Value1(System::Nullable<int32_t> value)
{
    pr_Value1 = value;
}

double NumericTestClass::get_Value2()
{
    return pr_Value2;
}

void NumericTestClass::set_Value2(double value)
{
    pr_Value2 = value;
}

int32_t NumericTestClass::get_Value3()
{
    return pr_Value3;
}

void NumericTestClass::set_Value3(int32_t value)
{
    pr_Value3 = value;
}

System::Nullable<int32_t> NumericTestClass::get_Value4()
{
    return pr_Value4;
}

void NumericTestClass::set_Value4(System::Nullable<int32_t> value)
{
    pr_Value4 = value;
}

bool NumericTestClass::get_Logical()
{
    return pr_Logical;
}

void NumericTestClass::set_Logical(bool value)
{
    pr_Logical = value;
}

System::DateTime NumericTestClass::get_Date()
{
    return pr_Date;
}

void NumericTestClass::set_Date(System::DateTime value)
{
    pr_Date = value;
}

NumericTestClass::NumericTestClass(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, bool logical, System::DateTime dateTime)
     : pr_Value2(0), pr_Value3(0), pr_Logical(false)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    set_Value1(value1);
    set_Value2(value2);
    set_Value3(value3);
    set_Value4(value4);
    set_Logical(logical);
    set_Date(dateTime);
}

int32_t NumericTestClass::Sum(int32_t value1, int32_t value2)
{
    int32_t result = value1 + value2;
    return result;
}

System::Object::shared_members_type ApiExamples::TestData::TestClasses::NumericTestClass::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::TestData::TestClasses::NumericTestClass::pr_Value1", this->pr_Value1);
    result.Add("ApiExamples::TestData::TestClasses::NumericTestClass::pr_Value4", this->pr_Value4);
    result.Add("ApiExamples::TestData::TestClasses::NumericTestClass::pr_Date", this->pr_Date);

    return result;
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
