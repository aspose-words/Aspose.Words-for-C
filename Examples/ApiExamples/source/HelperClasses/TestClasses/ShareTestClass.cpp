// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestClasses/ShareTestClass.h"

#include <system/primitive_types.h>
#include <system/math.h>
#include <system/convert.h>
#include <cstdint>

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace HelperClasses {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(3093842496u, ::Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass, ThisTypeBaseTypesInfo);

ShareTestClass::ShareTestClass(System::String sector, System::String industry, System::String ticker, double weight, double delta)
    : Weight(0), Delta(0)
{
    this->Sector = sector;
    this->Industry = industry;
    this->Ticker = ticker;
    this->Weight = weight;
    this->Delta = delta;
}

System::String ShareTestClass::Title()
{
    double percentValue = Delta * 100;
    return System::String::Format(u"{0}\r\n{1}%", Ticker, System::Convert::ToString(percentValue));
}

System::String ShareTestClass::Color()
{
    const double fullColorDelta = 0.016;
    const uint8_t unusedColorChannelValue = 80;
    
    uint8_t r = unusedColorChannelValue;
    uint8_t g = unusedColorChannelValue;
    uint8_t b = unusedColorChannelValue;
    
    int32_t value = unusedColorChannelValue + (int32_t)System::Math::Round(System::Math::Abs(Delta) / fullColorDelta * (std::numeric_limits<uint8_t>::max() - unusedColorChannelValue));
    
    if (value > std::numeric_limits<uint8_t>::max())
    {
        value = std::numeric_limits<uint8_t>::max();
    }
    
    if (Delta < 0)
    {
        r = (uint8_t)value;
    }
    else
    {
        g = (uint8_t)value;
    }
    
    return System::String::Format(u"#{0:X2}{1:X2}{2:X2}", r, g, b);
}

System::String ShareTestClass::IndustryColor()
{
    if (Industry == u"Consumer Electronics")
    {
        return u"#1B9629";
    }
    else if (Industry == u"Software - Infrastructure")
    {
        return u"#6029E3";
    }
    else if (Industry == u"Semiconductors")
    {
        return u"#E38529";
    }
    else if (Industry == u"Internet Content & Information")
    {
        return u"#964D05";
    }
    else if (Industry == u"Entertainment")
    {
        return u"#12E32B";
    }
    else if (Industry == u"Internet Retail")
    {
        return u"#96002C";
    }
    else if (Industry == u"Auto Manufactures")
    {
        return u"#1EE3A4";
    }
    else if (Industry == u"Credit Services")
    {
        return u"#D40B70";
    }
    else
    {
        return u"#888888";
    }
}

} // namespace TestClasses
} // namespace HelperClasses
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
