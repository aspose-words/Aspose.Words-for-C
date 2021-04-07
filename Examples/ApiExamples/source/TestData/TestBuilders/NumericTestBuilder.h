#pragma once

#include <cstdint>
#include <system/date_time.h>
#include <system/nullable.h>
#include <system/object.h>

#include "TestData/TestClasses/NumericTestClass.h"

using namespace ApiExamples::TestData::TestClasses;

namespace ApiExamples { namespace TestData { namespace TestBuilders {

class NumericTestBuilder : public System::Object
{
public:
    NumericTestBuilder();

    System::SharedPtr<NumericTestBuilder> WithValuesAndDate(System::Nullable<int32_t> value1, double value2, int32_t value3, System::Nullable<int32_t> value4,
                                                            System::DateTime dateTime);
    System::SharedPtr<NumericTestBuilder> WithValuesAndLogical(System::Nullable<int32_t> value1, double value2, int32_t value3,
                                                               System::Nullable<int32_t> value4, bool logical);
    System::SharedPtr<NumericTestClass> Build();

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

}}} // namespace ApiExamples::TestData::TestBuilders
