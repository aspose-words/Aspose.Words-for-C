// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestClasses/ShareQuoteTestClass.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace HelperClasses {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(1659023376u, ::Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass, ThisTypeBaseTypesInfo);

ShareQuoteTestClass::ShareQuoteTestClass(int32_t date, int32_t volume, double open, double high, double low, double close)
    : Date(0), Volume(0), Open(0), High(0), Low(0), Close(0)
{
    this->Date = date;
    this->Volume = volume;
    this->Open = open;
    this->High = high;
    this->Low = low;
    this->Close = close;
}

System::String ShareQuoteTestClass::Color()
{
    return (Open < Close) ? System::String(u"#1B9629") : System::String(u"#96002C");
}

} // namespace TestClasses
} // namespace HelperClasses
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
