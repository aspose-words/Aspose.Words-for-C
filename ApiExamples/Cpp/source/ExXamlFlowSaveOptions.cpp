// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXamlFlowSaveOptions.h"

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
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/XamlFlowSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
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

RTTI_INFO_IMPL_HASH(2751780628u, ::ApiExamples::ExXamlFlowSaveOptions::ImageUriPrinter, ThisTypeBaseTypesInfo);

String ExXamlFlowSaveOptions::ImageUriPrinter::get_ImagesFolderAlias()
{
    return pr_ImagesFolderAlias;
}

SharedPtr<System::Collections::Generic::List<String>> ExXamlFlowSaveOptions::ImageUriPrinter::get_Resources()
{
    return pr_Resources;
}

ExXamlFlowSaveOptions::ImageUriPrinter::ImageUriPrinter(String imagesFolderAlias)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    pr_ImagesFolderAlias = imagesFolderAlias;
    pr_Resources = MakeObject<System::Collections::Generic::List<String>>();
}

void ExXamlFlowSaveOptions::ImageUriPrinter::ImageSaving(SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args)
{
    get_Resources()->Add(args->get_ImageFileName());

    // If we specified a ImagesFolderAlias we will also need to redirect each stream to put its image in that folder
    args->set_ImageStream(MakeObject<System::IO::FileStream>(String::Format(u"{0}/{1}",get_ImagesFolderAlias(),args->get_ImageFileName()), System::IO::FileMode::Create));
    args->set_KeepImageStreamOpen(false);
}

System::Object::shared_members_type ApiExamples::ExXamlFlowSaveOptions::ImageUriPrinter::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExXamlFlowSaveOptions::ImageUriPrinter::pr_Resources", this->pr_Resources);

    return result;
}

void ExXamlFlowSaveOptions::TestImageFolder(SharedPtr<ExXamlFlowSaveOptions::ImageUriPrinter> callback)
{
    ASSERT_EQ(9, callback->get_Resources()->get_Count());
    for (auto resource : callback->get_Resources())
    {
        ASSERT_TRUE(System::IO::File::Exists(String::Format(u"{0}/{1}",callback->get_ImagesFolderAlias(),resource)));
    }
}

namespace gtest_test
{

class ExXamlFlowSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExXamlFlowSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExXamlFlowSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExXamlFlowSaveOptions> ExXamlFlowSaveOptions::s_instance;

} // namespace gtest_test

void ExXamlFlowSaveOptions::ImageFolder()
{
    // Open a document which contains images
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto callback = MakeObject<ExXamlFlowSaveOptions::ImageUriPrinter>(ArtifactsDir + u"XamlFlowImageFolderAlias");

    auto options = MakeObject<XamlFlowSaveOptions>();
    options->set_SaveFormat(Aspose::Words::SaveFormat::XamlFlow);
    options->set_ImagesFolder(ArtifactsDir + u"XamlFlowImageFolder");
    options->set_ImagesFolderAlias(ArtifactsDir + u"XamlFlowImageFolderAlias");
    options->set_ImageSavingCallback(callback);

    // A folder specified by ImagesFolderAlias will contain the images instead of ImagesFolder
    // We must ensure the folder exists before the streams can put their images into it
    System::IO::Directory::CreateDirectory_(options->get_ImagesFolderAlias());

    doc->Save(ArtifactsDir + u"XamlFlowSaveOptions.ImageFolder.xaml", options);

    for (auto resource : callback->get_Resources())
    {
        System::Console::WriteLine(String::Format(u"{0}/{1}",callback->get_ImagesFolderAlias(),resource));
    }
    TestImageFolder(callback);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExXamlFlowSaveOptions, ImageFolder)
{
    s_instance->ImageFolder();
}

} // namespace gtest_test

} // namespace ApiExamples
