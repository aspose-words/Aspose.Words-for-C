// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestClasses/MessageTestClass.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(905123642u, ::Aspose::Words::ApiExamples::TestData::TestClasses::MessageTestClass, ThisTypeBaseTypesInfo);

System::String MessageTestClass::get_Name() const
{
    return pr_Name;
}

void MessageTestClass::set_Name(System::String value)
{
    pr_Name = value;
}

System::String MessageTestClass::get_Message() const
{
    return pr_Message;
}

void MessageTestClass::set_Message(System::String value)
{
    pr_Message = value;
}

MessageTestClass::MessageTestClass(System::String name, System::String message)
{
    set_Name(name);
    set_Message(message);
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
