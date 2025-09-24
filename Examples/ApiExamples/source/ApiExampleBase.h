#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/string.h>
#include <system/reflection/assembly.h>
#include <mutex>
#include <memory>
#include <gtest/gtest.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

/// <summary>
/// Provides common infrastructure for all API examples that are implemented as unit tests.
/// </summary>
class ApiExampleBase : public System::Object
{
    typedef ApiExampleBase ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Gets the path to the currently running executable.
    /// </summary>
    static System::String get_AssemblyDir();
    /// <summary>
    /// Gets the path to the codebase directory.
    /// </summary>
    static System::String get_CodeBaseDir();
    /// <summary>
    /// Gets the path to the license used by the code examples.
    /// </summary>
    static System::String get_LicenseDir();
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String get_ArtifactsDir();
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String get_GoldsDir();
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String get_MyDir();
    /// <summary>
    /// Gets the path to the images used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String get_ImageDir();
    /// <summary>
    /// Gets the path of the demo database. Ends with a back slash.
    /// </summary>
    static System::String get_DatabaseDir();
    /// <summary>
    /// Gets the path of the free fonts. Ends with a back slash.
    /// </summary>
    static System::String get_FontsDir();
    /// <summary>
    /// Gets the URL of the test image.
    /// </summary>
    static System::String get_ImageUrl();
    
    void OneTimeSetUp();
    void SetUp();
    void OneTimeTearDown();
    static void SetUnlimitedLicense();
    /// <summary>
    /// Returns the code-base directory.
    /// </summary>
    static System::String GetCodeBaseDir(System::SharedPtr<System::Reflection::Assembly> assembly);
    /// <summary>
    /// Returns the assembly directory correctly even if the assembly is shadow-copied.
    /// </summary>
    static System::String GetAssemblyDir(System::SharedPtr<System::Reflection::Assembly> assembly);
    
    ApiExampleBase();
    
private:

    static System::String pr_AssemblyDir;
    static System::String pr_CodeBaseDir;
    static System::String pr_LicenseDir;
    static System::String pr_ArtifactsDir;
    static System::String pr_GoldsDir;
    static System::String pr_MyDir;
    static System::String pr_ImageDir;
    static System::String pr_DatabaseDir;
    static System::String pr_FontsDir;
    static System::String pr_ImageUrl;
    
    static void __StaticConstructor__();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


