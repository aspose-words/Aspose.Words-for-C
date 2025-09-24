#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/object.h>
#include <system/nullable.h>
#include <system/date_time.h>
#include <cstdint>

namespace Aspose {

namespace Words {

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

    System::Nullable<int32_t> get_Value1() const;
    void set_Value1(System::Nullable<int32_t> value);
    double get_Value2() const;
    void set_Value2(double value);
    int32_t get_Value3() const;
    void set_Value3(int32_t value);
    System::Nullable<int32_t> get_Value4() const;
    void set_Value4(System::Nullable<int32_t> value);
    bool get_Logical() const;
    void set_Logical(bool value);
    System::DateTime get_Date() const;
    void set_Date(System::DateTime value);
    
    NumericTestClass(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, bool logical, System::DateTime dateTime);
    
    int32_t Sum(int32_t value1, int32_t value2);
    
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
} // namespace Words
} // namespace Aspose


