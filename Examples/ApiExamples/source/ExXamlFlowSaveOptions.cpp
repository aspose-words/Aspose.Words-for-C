// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXamlFlowSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/nullable.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/default.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/XamlFlowSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(529605710u, ::Aspose::Words::ApiExamples::ExXamlFlowSaveOptions::ImageUriPrinter, ThisTypeBaseTypesInfo);

System::String ExXamlFlowSaveOptions::ImageUriPrinter::get_ImagesFolderAlias() const
{
    return pr_ImagesFolderAlias;
}

System::SharedPtr<System::Collections::Generic::List<System::String>> ExXamlFlowSaveOptions::ImageUriPrinter::get_Resources() const
{
    return pr_Resources;
}

ExXamlFlowSaveOptions::ImageUriPrinter::ImageUriPrinter(System::String imagesFolderAlias)
{
    pr_ImagesFolderAlias = imagesFolderAlias;
    pr_Resources = System::MakeObject<System::Collections::Generic::List<System::String>>();
}

void ExXamlFlowSaveOptions::ImageUriPrinter::ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args)
{
    get_Resources()->Add(args->get_ImageFileName());
    
    // If we specified an image folder alias, we would also need
    // to redirect each stream to put its image in the alias folder.
    args->set_ImageStream(System::MakeObject<System::IO::FileStream>(System::String::Format(u"{0}/{1}", get_ImagesFolderAlias(), args->get_ImageFileName()), System::IO::FileMode::Create));
    args->set_KeepImageStreamOpen(false);
}

RTTI_INFO_IMPL_HASH(4109639399u, ::Aspose::Words::ApiExamples::ExXamlFlowSaveOptions::SavingProgressCallback, ThisTypeBaseTypesInfo);

const double ExXamlFlowSaveOptions::SavingProgressCallback::MaxDuration = 0.01;

ExXamlFlowSaveOptions::SavingProgressCallback::SavingProgressCallback()
{
    mSavingStartedAt = System::DateTime::get_Now();
}

void ExXamlFlowSaveOptions::SavingProgressCallback::Notify(System::SharedPtr<Aspose::Words::Saving::DocumentSavingArgs> args)
{
    System::DateTime canceledAt = System::DateTime::get_Now();
    double ellapsedSeconds = (canceledAt - mSavingStartedAt).get_TotalSeconds();
    if (ellapsedSeconds > MaxDuration)
    {
        throw System::OperationCanceledException(System::String::Format(u"EstimatedProgress = {0}; CanceledAt = {1}", args->get_EstimatedProgress(), canceledAt));
    }
}


RTTI_INFO_IMPL_HASH(3072051451u, ::Aspose::Words::ApiExamples::ExXamlFlowSaveOptions, ThisTypeBaseTypesInfo);

void ExXamlFlowSaveOptions::TestImageFolder(System::SharedPtr<Aspose::Words::ApiExamples::ExXamlFlowSaveOptions::ImageUriPrinter> callback)
{
    ASSERT_EQ(9, callback->get_Resources()->get_Count());
    for (auto&& resource : callback->get_Resources())
    {
        ASSERT_TRUE(System::IO::File::Exists(System::String::Format(u"{0}/{1}", callback->get_ImagesFolderAlias(), resource)));
    }
}


namespace gtest_test
{

class ExXamlFlowSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExXamlFlowSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExXamlFlowSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExXamlFlowSaveOptions> ExXamlFlowSaveOptions::s_instance;

} // namespace gtest_test

void ExXamlFlowSaveOptions::ImageFolder()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExXamlFlowSaveOptions::ImageUriPrinter>(get_ArtifactsDir() + u"XamlFlowImageFolderAlias");
    
    // Create a "XamlFlowSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to the XAML save format.
    auto options = System::MakeObject<Aspose::Words::Saving::XamlFlowSaveOptions>();
    
    ASSERT_EQ(Aspose::Words::SaveFormat::XamlFlow, options->get_SaveFormat());
    
    // Use the "ImagesFolder" property to assign a folder in the local file system into which
    // Aspose.Words will save all the document's linked images.
    options->set_ImagesFolder(get_ArtifactsDir() + u"XamlFlowImageFolder");
    
    // Use the "ImagesFolderAlias" property to use this folder
    // when constructing image URIs instead of the images folder's name.
    options->set_ImagesFolderAlias(get_ArtifactsDir() + u"XamlFlowImageFolderAlias");
    
    options->set_ImageSavingCallback(callback);
    
    // A folder specified by "ImagesFolderAlias" will need to contain the resources instead of "ImagesFolder".
    // We must ensure the folder exists before the callback's streams can put their resources into it.
    System::IO::Directory::CreateDirectory_(options->get_ImagesFolderAlias());
    
    doc->Save(get_ArtifactsDir() + u"XamlFlowSaveOptions.ImageFolder.xaml", options);
    
    for (auto&& resource : callback->get_Resources())
    {
        std::cout << System::String::Format(u"{0}/{1}", callback->get_ImagesFolderAlias(), resource) << std::endl;
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

void ExXamlFlowSaveOptions::ProgressCallback(Aspose::Words::SaveFormat saveFormat, System::String ext)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    // Following formats are supported: XamlFlow, XamlFlowPack.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::XamlFlowSaveOptions>(saveFormat);
    saveOptions->set_ProgressCallback(System::MakeObject<Aspose::Words::ApiExamples::ExXamlFlowSaveOptions::SavingProgressCallback>());
    
    System::OperationCanceledException exception; ASPOSE_ASSERT_THROW(static_cast<std::function<void()>>([&doc, &ext, &saveOptions]() -> void
    {
        doc->Save(get_ArtifactsDir() + System::String::Format(u"XamlFlowSaveOptions.ProgressCallback.{0}", ext), saveOptions);
    })(), System::OperationCanceledException, &exception);
    System::Nullable<bool> actual = System::Default<System::Nullable<bool>>();
    System::OperationCanceledException condExpression = exception;
    if (condExpression != nullptr)
    {
        actual = condExpression->get_Message().Contains(u"EstimatedProgress");
    }
    ASSERT_TRUE(actual);
}

namespace gtest_test
{

using ExXamlFlowSaveOptions_ProgressCallback_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExXamlFlowSaveOptions::ProgressCallback)>::type;

struct ExXamlFlowSaveOptions_ProgressCallback : public ExXamlFlowSaveOptions, public Aspose::Words::ApiExamples::ExXamlFlowSaveOptions, public ::testing::WithParamInterface<ExXamlFlowSaveOptions_ProgressCallback_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::XamlFlow, u"xamlflow"),
            std::make_tuple(Aspose::Words::SaveFormat::XamlFlowPack, u"xamlflowpack"),
        };
    }
};

TEST_P(ExXamlFlowSaveOptions_ProgressCallback, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ProgressCallback(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExXamlFlowSaveOptions_ProgressCallback, ::testing::ValuesIn(ExXamlFlowSaveOptions_ProgressCallback::TestCases()));

} // namespace gtest_test

void ExXamlFlowSaveOptions::XamlReplaceBackslashWithYenSign()
{
    //ExStart:XamlReplaceBackslashWithYenSign
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:XamlFlowSaveOptions.ReplaceBackslashWithYenSign
    //ExSummary:Shows how to replace backslash characters with yen signs (Xaml).
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Korean backslash symbol.docx");
    
    // By default, Aspose.Words mimics MS Word's behavior and doesn't replace backslash characters with yen signs in
    // generated HTML documents. However, previous versions of Aspose.Words performed such replacements in certain
    // scenarios. This flag enables backward compatibility with previous versions of Aspose.Words.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::XamlFlowSaveOptions>();
    saveOptions->set_ReplaceBackslashWithYenSign(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ReplaceBackslashWithYenSign.xaml", saveOptions);
    //ExEnd:XamlReplaceBackslashWithYenSign
}

namespace gtest_test
{

TEST_F(ExXamlFlowSaveOptions, XamlReplaceBackslashWithYenSign)
{
    s_instance->XamlReplaceBackslashWithYenSign();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
