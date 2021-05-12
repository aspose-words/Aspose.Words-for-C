#pragma once

#include <system/reflection/assembly.h>
#include <system/string.h>
#include <gtest/gtest.h>

namespace DocsExamples {

class DocsExamplesBase : public System::Object
{
public:
    static struct __StaticConstructor__
    {
        __StaticConstructor__();
    } s_constructor__;

    static void OneTimeSetUp();

    static void SetUp();

    static void OneTimeTearDown();

    static void SetUnlimitedLicense();

    /// <summary>
    /// Returns the code-base directory.
    /// </summary>
    static System::String GetCodeBaseDir(System::SharedPtr<System::Reflection::Assembly> assembly);

    /// <summary>
    /// Gets the path to the codebase directory.
    /// </summary>
    static System::String MainDataDir;

    /// <summary>
    /// Gets the path to the documents used by the code examples.
    /// </summary>
    static System::String MyDir;

    /// <summary>
    /// Gets the path to the images used by the code examples.
    /// </summary>
    static System::String ImagesDir;

    /// <summary>
    /// Gets the path of the demo database.
    /// </summary>
    static System::String DatabaseDir;

    /// <summary>
    /// Gets the path to the license used by the code examples.
    /// </summary>
    static System::String LicenseDir;

    /// <summary>
    /// Gets the path to the artifacts used by the code examples.
    /// </summary>
    static System::String ArtifactsDir;

    /// <summary>
    /// Gets the path of the free fonts. Ends with a back slash.
    /// </summary>
    static System::String FontsDir;
};

} // namespace DocsExamples
