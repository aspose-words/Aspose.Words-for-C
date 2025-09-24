#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// 18/02/2024 by Vyacheslav Deryushev

#include <system/string.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class ClientTestClass : public System::Object
{
    typedef ClientTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::String get_Name() const;
    void set_Name(System::String value);
    System::String get_Country() const;
    void set_Country(System::String value);
    System::String get_LocalAddress() const;
    void set_LocalAddress(System::String value);
    
private:

    System::String pr_Name;
    System::String pr_Country;
    System::String pr_LocalAddress;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


