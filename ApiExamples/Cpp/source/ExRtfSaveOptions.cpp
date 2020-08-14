// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRtfSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/RtfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExRtfSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExRtfSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExRtfSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExRtfSaveOptions> ExRtfSaveOptions::s_instance;

} // namespace gtest_test

void ExRtfSaveOptions::ExportImages(bool doExportImagesForOldReaders)
{
    //ExStart
    //ExFor:RtfSaveOptions
    //ExFor:RtfSaveOptions.ExportCompactSize
    //ExFor:RtfSaveOptions.ExportImagesForOldReaders
    //ExFor:RtfSaveOptions.SaveFormat
    //ExSummary:Shows how to save a document to .rtf with custom options.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // Configure a RtfSaveOptions instance to make our output document more suitable for older devices
    auto options = MakeObject<RtfSaveOptions>();
    options->set_SaveFormat(Aspose::Words::SaveFormat::Rtf);
    options->set_ExportCompactSize(true);
    options->set_ExportImagesForOldReaders(doExportImagesForOldReaders);

    doc->Save(ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf", options);
    //ExEnd

    if (doExportImagesForOldReaders)
    {
        TestUtil::FileContainsString(u"nonshppict", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf");
        TestUtil::FileContainsString(u"shprslt", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf");
    }
    else
    {

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<void()> _local_func_0 = []()
        {
            TestUtil::FileContainsString(u"nonshppict", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf");
        };

        ASSERT_THROW(_local_func_0(), NUnit::Framework::AssertionException);

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<void()> _local_func_1 = []()
        {
            TestUtil::FileContainsString(u"shprslt", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf");
        };

        ASSERT_THROW(_local_func_1(), NUnit::Framework::AssertionException);
    }
}

namespace gtest_test
{

using ExportImages_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRtfSaveOptions::ExportImages)>::type;

struct ExportImagesVP : public ExRtfSaveOptions, public ApiExamples::ExRtfSaveOptions, public ::testing::WithParamInterface<ExportImages_Args>
{
    static std::vector<ExportImages_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExportImagesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportImages(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRtfSaveOptions, ExportImagesVP, ::testing::ValuesIn(ExportImagesVP::TestCases()));

} // namespace gtest_test

void ExRtfSaveOptions::SaveImagesAsWmf()
{
    //ExStart
    //ExFor:RtfSaveOptions.SaveImagesAsWmf
    //ExSummary:Shows how to save all images as Wmf when saving to the Rtf document.
    // Open a document that contains images in the jpeg format
    auto doc = MakeObject<Document>(MyDir + u"Images.docx");

    SharedPtr<NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    auto shapeWithJpg = System::DynamicCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(0));
    ASSERT_EQ(Aspose::Words::Drawing::ImageType::Jpeg, shapeWithJpg->get_ImageData()->get_ImageType());

    auto rtfSaveOptions = MakeObject<RtfSaveOptions>();
    rtfSaveOptions->set_SaveImagesAsWmf(true);
    doc->Save(ArtifactsDir + u"RtfSaveOptions.SaveImagesAsWmf.rtf", rtfSaveOptions);
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"RtfSaveOptions.SaveImagesAsWmf.rtf");

    shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    auto shapeWithWmf = System::DynamicCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(0));
    ASSERT_EQ(Aspose::Words::Drawing::ImageType::Wmf, shapeWithWmf->get_ImageData()->get_ImageType());
}

namespace gtest_test
{

TEST_F(ExRtfSaveOptions, SaveImagesAsWmf)
{
    s_instance->SaveImagesAsWmf();
}

} // namespace gtest_test

} // namespace ApiExamples
