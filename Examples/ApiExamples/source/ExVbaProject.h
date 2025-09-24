#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReference.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Vba;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExVbaProject : public ApiExampleBase
{
    typedef ExVbaProject ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void CreateNewVbaProject();
    void CloneVbaProject();
    void RemoveVbaReference();
    void IsProtected();
    
protected:

    /// <summary>
    /// Returns string representing LibId path of a specified reference. 
    /// </summary>
    static System::String GetLibIdPath(System::SharedPtr<Aspose::Words::Vba::VbaReference> reference);
    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    static System::String GetLibIdReferencePath(System::String libIdReference);
    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    static System::String GetLibIdProjectPath(System::String libIdProject);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


