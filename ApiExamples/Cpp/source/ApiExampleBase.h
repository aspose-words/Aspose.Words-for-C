#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/object.h>
#include <system/reflection/assembly.h>
#include <system/shared_ptr.h>
#include <system/string.h>
#include <gtest/gtest.h>

namespace ApiExamples {

/// <summary>
/// Provides common infrastructure for all API examples that are implemented as unit tests.
/// </summary>
class ApiExampleBase : public System::Object
{
public:
    void OneTimeSetUp();
    void SetUp();
    void OneTimeTearDown();
    /// <summary>
    /// Determine if runtime is Mono.
    /// Workaround for .netcore.
    /// </summary>
    /// <returns>True if being executed in Mono, false otherwise.</returns>
    static bool IsRunningOnMono();

protected:
    /// <summary>
    /// Gets the path to the currently running executable.
    /// </summary>
    static System::String AssemblyDir;
    /// <summary>
    /// Gets the path to the codebase directory.
    /// </summary>
    static System::String CodeBaseDir;
    /// <summary>
    /// Gets the path to the license used by the code examples.
    /// </summary>
    static System::String LicenseDir;
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String ArtifactsDir;
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String GoldsDir;
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String MyDir;
    /// <summary>
    /// Gets the path to the images used by the code examples. Ends with a back slash.
    /// </summary>
    static System::String ImageDir;
    /// <summary>
    /// Gets the path of the demo database. Ends with a back slash.
    /// </summary>
    static System::String DatabaseDir;
    /// <summary>
    /// Gets the path of the free fonts. Ends with a back slash.
    /// </summary>
    static System::String FontsDir;
    /// <summary>
    /// Gets the URL of the Aspose logo.
    /// </summary>
    static System::String AsposeLogoUrl;

    /// <summary>
    /// Checks when we need to ignore test on mono.
    /// </summary>
    static bool CheckForSkipMono();
    static void SetUnlimitedLicense();
    /// <summary>
    /// Returns the code-base directory.
    /// </summary>
    static System::String GetCodeBaseDir(System::SharedPtr<System::Reflection::Assembly> assembly);
    /// <summary>
    /// Returns the assembly directory correctly even if the assembly is shadow-copied.
    /// </summary>
    static System::String GetAssemblyDir(System::SharedPtr<System::Reflection::Assembly> assembly);

private:
    static struct __StaticConstructor__
    {
        __StaticConstructor__();
    } s_constructor__;
};

} // namespace ApiExamples
