#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <iostream>
#include <Aspose.Words.Cpp/BuildVersionInfo.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/regex.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace ApiExamples {

class ExBuildVersion : public ApiExampleBase
{
public:
    void PrintBuildVersionInfo()
    {
        //ExStart
        //ExFor:BuildVersionInfo
        //ExFor:BuildVersionInfo.Product
        //ExFor:BuildVersionInfo.Version
        //ExSummary:Shows how to display information about your installed version of Aspose.Words.
        std::cout << "I am currently using " << BuildVersionInfo::get_Product() << ", version number " << BuildVersionInfo::get_Version() << "!" << std::endl;
        //ExEnd

        ASSERT_EQ(u"Aspose.Words for C++", BuildVersionInfo::get_Product());
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::IsMatch(BuildVersionInfo::get_Version(), u"[0-9]{2}.[0-9]{1,2}"));
    }
};

} // namespace ApiExamples
