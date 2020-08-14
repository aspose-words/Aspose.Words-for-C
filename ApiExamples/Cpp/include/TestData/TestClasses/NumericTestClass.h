#pragma once

#include <system/object.h>
#include <system/nullable.h>
#include <system/date_time.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class NumericTestClass : public System::Object
{
    typedef NumericTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::Nullable<int32_t> get_Value1();
    void set_Value1(System::Nullable<int32_t> value);
    double get_Value2();
    void set_Value2(double value);
    int32_t get_Value3();
    void set_Value3(int32_t value);
    System::Nullable<int32_t> get_Value4();
    void set_Value4(System::Nullable<int32_t> value);
    bool get_Logical();
    void set_Logical(bool value);
    System::DateTime get_Date();
    void set_Date(System::DateTime value);
    
    NumericTestClass(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, bool logical, System::DateTime dateTime);
    
    int32_t Sum(int32_t value1, int32_t value2);
    
protected:

    System::Object::shared_members_type GetSharedMembers() override;
    
private:

    System::Nullable<int32_t> pr_Value1;
    double pr_Value2;
    int32_t pr_Value3;
    System::Nullable<int32_t> pr_Value4;
    bool pr_Logical;
    System::DateTime pr_Date;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples


