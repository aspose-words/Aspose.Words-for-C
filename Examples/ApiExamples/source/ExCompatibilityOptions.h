#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/ilist.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Settings;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExCompatibilityOptions : public ApiExampleBase
{
    typedef ExCompatibilityOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void OptimizeFor();
    void Tables();
    void Breaks();
    void Spacing();
    void WordPerfect();
    void Alignment();
    void Legacy();
    void List();
    void Misc();
    
protected:

    /// <summary>
    /// Groups all flags in a document's compatibility options object by state, then prints each group.
    /// </summary>
    static void PrintCompatibilityOptions(System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> options);
    static void AddOptionName(bool option, System::String optionName, System::SharedPtr<System::Collections::Generic::IList<System::String>> enabledOptions, System::SharedPtr<System::Collections::Generic::IList<System::String>> disabledOptions);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


