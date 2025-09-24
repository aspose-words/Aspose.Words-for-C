#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/object.h>
#include <system/nullable.h>
#include <system/date_time.h>
#include <cstdint>

#include "HelperClasses/TestClasses/NumericTestClass.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

class NumericTestBuilder : public System::Object
{
    typedef NumericTestBuilder ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    NumericTestBuilder();
    
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::NumericTestBuilder> WithValuesAndDate(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, System::DateTime dateTime);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::NumericTestBuilder> WithValuesAndLogical(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4, bool logical);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::NumericTestBuilder> WithValues(System::Nullable<int32_t> value1, double value2);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::NumericTestClass> Build();
    
protected:

    virtual ~NumericTestBuilder();
    
private:

    System::Nullable<int32_t> mValue1;
    double mValue2;
    int32_t mValue3;
    System::Nullable<int32_t> mValue4;
    bool mLogical;
    System::DateTime mDate;
    
};

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


