#include "TestData/TestClasses/NumericTestClass.h"

namespace ApiExamples { namespace TestData { namespace TestClasses {

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

NumericTestClass::NumericTestClass(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, bool logical,
                                   System::DateTime dateTime)
    : pr_Value2(0), pr_Value3(0), pr_Logical(false)
{
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

}}} // namespace ApiExamples::TestData::TestClasses
