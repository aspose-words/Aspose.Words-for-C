// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRtfSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/RtfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4124858189u, ::Aspose::Words::ApiExamples::ExRtfSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExRtfSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExRtfSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExRtfSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExRtfSaveOptions> ExRtfSaveOptions::s_instance;

} // namespace gtest_test

void ExRtfSaveOptions::ExportImages(bool exportImagesForOldReaders)
{
    //ExStart
    //ExFor:RtfSaveOptions
    //ExFor:RtfSaveOptions.ExportCompactSize
    //ExFor:RtfSaveOptions.ExportImagesForOldReaders
    //ExFor:RtfSaveOptions.SaveFormat
    //ExSummary:Shows how to save a document to .rtf with custom options.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
    auto options = System::MakeObject<Aspose::Words::Saving::RtfSaveOptions>();
    
    ASSERT_EQ(Aspose::Words::SaveFormat::Rtf, options->get_SaveFormat());
    
    // Set the "ExportCompactSize" property to "true" to
    // reduce the saved document's size at the cost of right-to-left text compatibility.
    options->set_ExportCompactSize(true);
    
    // Set the "ExportImagesFotOldReaders" property to "true" to use extra keywords to ensure that our document is
    // compatible with pre-Microsoft Word 97 readers and WordPad.
    // Set the "ExportImagesFotOldReaders" property to "false" to reduce the size of the document,
    // but prevent old readers from being able to read any non-metafile or BMP images that the document may contain.
    options->set_ExportImagesForOldReaders(exportImagesForOldReaders);
    
    doc->Save(get_ArtifactsDir() + u"RtfSaveOptions.ExportImages.rtf", options);
    //ExEnd
    
    if (exportImagesForOldReaders)
    {
        Aspose::Words::ApiExamples::TestUtil::FileContainsString(u"nonshppict", get_ArtifactsDir() + u"RtfSaveOptions.ExportImages.rtf");
        Aspose::Words::ApiExamples::TestUtil::FileContainsString(u"shprslt", get_ArtifactsDir() + u"RtfSaveOptions.ExportImages.rtf");
    }
    else
    {
    }
}

namespace gtest_test
{

using ExRtfSaveOptions_ExportImages_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRtfSaveOptions::ExportImages)>::type;

struct ExRtfSaveOptions_ExportImages : public ExRtfSaveOptions, public Aspose::Words::ApiExamples::ExRtfSaveOptions, public ::testing::WithParamInterface<ExRtfSaveOptions_ExportImages_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRtfSaveOptions_ExportImages, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportImages(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRtfSaveOptions_ExportImages, ::testing::ValuesIn(ExRtfSaveOptions_ExportImages::TestCases()));

} // namespace gtest_test

void ExRtfSaveOptions::SaveImagesAsWmf(bool saveImagesAsWmf)
{
    //ExStart
    //ExFor:RtfSaveOptions.SaveImagesAsWmf
    //ExSummary:Shows how to convert all images in a document to the Windows Metafile format as we save the document as an RTF.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Jpeg image:");
    System::SharedPtr<Aspose::Words::Drawing::Shape> imageShape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    ASSERT_EQ(Aspose::Words::Drawing::ImageType::Jpeg, imageShape->get_ImageData()->get_ImageType());
    
    builder->InsertParagraph();
    builder->Writeln(u"Png image:");
    imageShape = builder->InsertImage(get_ImageDir() + u"Transparent background logo.png");
    
    ASSERT_EQ(Aspose::Words::Drawing::ImageType::Png, imageShape->get_ImageData()->get_ImageType());
    
    // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
    auto rtfSaveOptions = System::MakeObject<Aspose::Words::Saving::RtfSaveOptions>();
    
    // Set the "SaveImagesAsWmf" property to "true" to convert all images in the document to WMF as we save it to RTF.
    // Doing so will help readers such as WordPad to read our document.
    // Set the "SaveImagesAsWmf" property to "false" to preserve the original format of all images in the document
    // as we save it to RTF. This will preserve the quality of the images at the cost of compatibility with older RTF readers.
    rtfSaveOptions->set_SaveImagesAsWmf(saveImagesAsWmf);
    
    doc->Save(get_ArtifactsDir() + u"RtfSaveOptions.SaveImagesAsWmf.rtf", rtfSaveOptions);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"RtfSaveOptions.SaveImagesAsWmf.rtf");
    
    System::SharedPtr<Aspose::Words::NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    
    if (saveImagesAsWmf)
    {
        ASSERT_EQ(Aspose::Words::Drawing::ImageType::Wmf, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(0)))->get_ImageData()->get_ImageType());
        ASSERT_EQ(Aspose::Words::Drawing::ImageType::Wmf, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(1)))->get_ImageData()->get_ImageType());
    }
    else
    {
        ASSERT_EQ(Aspose::Words::Drawing::ImageType::Jpeg, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(0)))->get_ImageData()->get_ImageType());
        ASSERT_EQ(Aspose::Words::Drawing::ImageType::Png, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(1)))->get_ImageData()->get_ImageType());
    }
    //ExEnd
}

namespace gtest_test
{

using ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRtfSaveOptions::SaveImagesAsWmf)>::type;

struct ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf : public ExRtfSaveOptions, public Aspose::Words::ApiExamples::ExRtfSaveOptions, public ::testing::WithParamInterface<ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf, Test)
{
    RecordProperty("category", "SkipMono");
    RecordProperty("category1", "SkipMono");
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SaveImagesAsWmf(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf, ::testing::ValuesIn(ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
