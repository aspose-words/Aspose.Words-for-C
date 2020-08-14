// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/directory.h>
#include <system/date_time.h>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/DocSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExDocSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExDocSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDocSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDocSaveOptions> ExDocSaveOptions::s_instance;

} // namespace gtest_test

void ExDocSaveOptions::SaveAsDoc()
{
    //ExStart
    //ExFor:DocSaveOptions
    //ExFor:DocSaveOptions.#ctor()
    //ExFor:DocSaveOptions.#ctor(SaveFormat)
    //ExFor:DocSaveOptions.Password
    //ExFor:DocSaveOptions.SaveFormat
    //ExFor:DocSaveOptions.SaveRoutingSlip
    //ExSummary:Shows how to set save options for classic Microsoft Word document versions.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Write(u"Hello world!");

    // DocSaveOptions only applies to Doc and Dot save formats
    auto options = MakeObject<DocSaveOptions>(Aspose::Words::SaveFormat::Doc);

    // Set a password with which the document will be encrypted, and which will be required to open it
    options->set_Password(u"MyPassword");

    // If the document contains a routing slip, we can preserve it while saving by setting this flag to true
    options->set_SaveRoutingSlip(true);

    doc->Save(ArtifactsDir + u"DocSaveOptions.SaveAsDoc.doc", options);
    //ExEnd

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&doc]()
    {
        doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.SaveAsDoc.doc");
    };

    ASSERT_THROW(_local_func_0(), IncorrectPasswordException);

    auto loadOptions = MakeObject<LoadOptions>(u"MyPassword");
    doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.SaveAsDoc.doc", loadOptions);

    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocSaveOptions, SaveAsDoc)
{
    s_instance->SaveAsDoc();
}

} // namespace gtest_test

void ExDocSaveOptions::TempFolder()
{
    //ExStart
    //ExFor:SaveOptions.TempFolder
    //ExSummary:Shows how to save a document using temporary files.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // We can use a SaveOptions object to set the saving method of a document from a MemoryStream to temporary files
    // While saving, the files will briefly pop up in the folder we set as the TempFolder attribute below
    // Doing this will free up space in the memory that the stream would usually occupy
    auto options = MakeObject<DocSaveOptions>();
    options->set_TempFolder(ArtifactsDir + u"TempFiles");

    // Ensure that the directory exists and save
    System::IO::Directory::CreateDirectory_(options->get_TempFolder());

    doc->Save(ArtifactsDir + u"DocSaveOptions.TempFolder.doc", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocSaveOptions, TempFolder)
{
    s_instance->TempFolder();
}

} // namespace gtest_test

void ExDocSaveOptions::PictureBullets()
{
    //ExStart
    //ExFor:DocSaveOptions.SavePictureBullet
    //ExSummary:Shows how to remove PictureBullet data from the document.
    auto doc = MakeObject<Document>(MyDir + u"Image bullet points.docx");
    ASSERT_FALSE(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData() == nullptr);
    //ExSkip

    // Word 97 cannot work correctly with PictureBullet data
    // To remove PictureBullet data, set the option to "false"
    auto saveOptions = MakeObject<DocSaveOptions>(Aspose::Words::SaveFormat::Doc);
    saveOptions->set_SavePictureBullet(false);

    doc->Save(ArtifactsDir + u"DocSaveOptions.PictureBullets.doc", saveOptions);
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.PictureBullets.doc");

    ASSERT_TRUE(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData() == nullptr);
}

namespace gtest_test
{

TEST_F(ExDocSaveOptions, PictureBullets)
{
    s_instance->PictureBullets();
}

} // namespace gtest_test

void ExDocSaveOptions::UpdateLastPrintedProperty(bool isUpdateLastPrintedProperty)
{
    //ExStart
    //ExFor:SaveOptions.UpdateLastPrintedProperty
    //ExSummary:Shows how to update BuiltInDocumentProperties.LastPrinted property before saving.
    auto doc = MakeObject<Document>();

    // Aspose.Words update BuiltInDocumentProperties.LastPrinted property by default
    auto saveOptions = MakeObject<DocSaveOptions>();
    saveOptions->set_UpdateLastPrintedProperty(isUpdateLastPrintedProperty);

    doc->Save(ArtifactsDir + u"DocSaveOptions.UpdateLastPrintedProperty.docx", saveOptions);
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocSaveOptions.UpdateLastPrintedProperty.docx");

    ASPOSE_ASSERT_NE(isUpdateLastPrintedProperty, System::DateTime::Parse(u"1/1/0001 00:00:00") == doc->get_BuiltInDocumentProperties()->get_LastPrinted().get_Date());
}

namespace gtest_test
{

using UpdateLastPrintedProperty_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocSaveOptions::UpdateLastPrintedProperty)>::type;

struct UpdateLastPrintedPropertyVP : public ExDocSaveOptions, public ApiExamples::ExDocSaveOptions, public ::testing::WithParamInterface<UpdateLastPrintedProperty_Args>
{
    static std::vector<UpdateLastPrintedProperty_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(UpdateLastPrintedPropertyVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateLastPrintedProperty(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocSaveOptions, UpdateLastPrintedPropertyVP, ::testing::ValuesIn(UpdateLastPrintedPropertyVP::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
