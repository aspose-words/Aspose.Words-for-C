// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXamlFixedSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/enumerator_adapter.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/XamlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1375976225u, ::Aspose::Words::ApiExamples::ExXamlFixedSaveOptions::ResourceUriPrinter, ThisTypeBaseTypesInfo);

System::SharedPtr<System::Collections::Generic::List<System::String>> ExXamlFixedSaveOptions::ResourceUriPrinter::get_Resources() const
{
    return pr_Resources;
}

ExXamlFixedSaveOptions::ResourceUriPrinter::ResourceUriPrinter()
{
    pr_Resources = System::MakeObject<System::Collections::Generic::List<System::String>>();
}

void ExXamlFixedSaveOptions::ResourceUriPrinter::ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args)
{
    get_Resources()->Add(System::String::Format(u"Resource \"{0}\"\n\t{1}", args->get_ResourceFileName(), args->get_ResourceFileUri()));
    
    // If we specified a resource folder alias, we would also need
    // to redirect each stream to put its resource in the alias folder.
    args->set_ResourceStream(System::MakeObject<System::IO::FileStream>(args->get_ResourceFileUri(), System::IO::FileMode::Create));
    args->set_KeepResourceStreamOpen(false);
}


RTTI_INFO_IMPL_HASH(708711941u, ::Aspose::Words::ApiExamples::ExXamlFixedSaveOptions, ThisTypeBaseTypesInfo);

void ExXamlFixedSaveOptions::TestResourceFolder(System::SharedPtr<Aspose::Words::ApiExamples::ExXamlFixedSaveOptions::ResourceUriPrinter> callback)
{
    ASSERT_EQ(15, callback->get_Resources()->get_Count());
    for (auto&& resource : callback->get_Resources())
    {
        ASSERT_TRUE(System::IO::File::Exists(resource.Split(u'\t')->idx_get(1)));
    }
}


namespace gtest_test
{

class ExXamlFixedSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExXamlFixedSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExXamlFixedSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExXamlFixedSaveOptions> ExXamlFixedSaveOptions::s_instance;

} // namespace gtest_test

void ExXamlFixedSaveOptions::ResourceFolder()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExXamlFixedSaveOptions::ResourceUriPrinter>();
    
    // Create a "XamlFixedSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to the XAML save format.
    auto options = System::MakeObject<Aspose::Words::Saving::XamlFixedSaveOptions>();
    
    ASSERT_EQ(Aspose::Words::SaveFormat::XamlFixed, options->get_SaveFormat());
    
    // Use the "ResourcesFolder" property to assign a folder in the local file system into which
    // Aspose.Words will save all the document's linked resources, such as images and fonts.
    options->set_ResourcesFolder(get_ArtifactsDir() + u"XamlFixedResourceFolder");
    
    // Use the "ResourcesFolderAlias" property to use this folder
    // when constructing image URIs instead of the resources folder's name.
    options->set_ResourcesFolderAlias(get_ArtifactsDir() + u"XamlFixedFolderAlias");
    
    options->set_ResourceSavingCallback(callback);
    
    // A folder specified by "ResourcesFolderAlias" will need to contain the resources instead of "ResourcesFolder".
    // We must ensure the folder exists before the callback's streams can put their resources into it.
    System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());
    
    doc->Save(get_ArtifactsDir() + u"XamlFixedSaveOptions.ResourceFolder.xaml", options);
    
    for (auto&& resource : callback->get_Resources())
    {
        std::cout << resource << std::endl;
    }
    TestResourceFolder(callback);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExXamlFixedSaveOptions, ResourceFolder)
{
    s_instance->ResourceFolder();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
