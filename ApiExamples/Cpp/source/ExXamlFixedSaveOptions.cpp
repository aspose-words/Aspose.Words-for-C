// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXamlFixedSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/scope_guard.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/XamlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1716971495u, ::ApiExamples::ExXamlFixedSaveOptions::ResourceUriPrinter, ThisTypeBaseTypesInfo);

SharedPtr<System::Collections::Generic::List<String>> ExXamlFixedSaveOptions::ResourceUriPrinter::get_Resources()
{
    return pr_Resources;
}

ExXamlFixedSaveOptions::ResourceUriPrinter::ResourceUriPrinter()
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    pr_Resources = MakeObject<System::Collections::Generic::List<String>>();
}

void ExXamlFixedSaveOptions::ResourceUriPrinter::ResourceSaving(SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args)
{
    // If we set a folder alias in the SaveOptions object, it will be stored here
    get_Resources()->Add(String::Format(u"Resource \"{0}\"\n\t{1}",args->get_ResourceFileName(),args->get_ResourceFileUri()));

    // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
    args->set_ResourceStream(MakeObject<System::IO::FileStream>(args->get_ResourceFileUri(), System::IO::FileMode::Create));
    args->set_KeepResourceStreamOpen(false);
}

System::Object::shared_members_type ApiExamples::ExXamlFixedSaveOptions::ResourceUriPrinter::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExXamlFixedSaveOptions::ResourceUriPrinter::pr_Resources", this->pr_Resources);

    return result;
}

void ExXamlFixedSaveOptions::TestResourceFolder(SharedPtr<ExXamlFixedSaveOptions::ResourceUriPrinter> callback)
{
    ASSERT_EQ(15, callback->get_Resources()->get_Count());
    for (auto resource : callback->get_Resources())
    {
        ASSERT_TRUE(System::IO::File::Exists(resource.Split(MakeArray<char16_t>({u'\t'}))->idx_get(1)));
    }
}

namespace gtest_test
{

class ExXamlFixedSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExXamlFixedSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExXamlFixedSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExXamlFixedSaveOptions> ExXamlFixedSaveOptions::s_instance;

} // namespace gtest_test

void ExXamlFixedSaveOptions::ResourceFolder()
{
    // Open a document which contains resources
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto callback = MakeObject<ExXamlFixedSaveOptions::ResourceUriPrinter>();

    auto options = MakeObject<XamlFixedSaveOptions>();
    options->set_SaveFormat(Aspose::Words::SaveFormat::XamlFixed);
    options->set_ResourcesFolder(ArtifactsDir + u"XamlFixedResourceFolder");
    options->set_ResourcesFolderAlias(ArtifactsDir + u"XamlFixedFolderAlias");
    options->set_ResourceSavingCallback(callback);

    // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
    // We must ensure the folder exists before the streams can put their resources into it
    System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

    doc->Save(ArtifactsDir + u"XamlFixedSaveOptions.ResourceFolder.xaml", options);

    for (auto resource : callback->get_Resources())
    {
        System::Console::WriteLine(resource);
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
