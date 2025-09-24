// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/io/file_info.h>
#include <system/io/directory.h>
#include <system/date_time.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/DocSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2363433113u, ::Aspose::Words::ApiExamples::ExDocSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExDocSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDocSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDocSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDocSaveOptions> ExDocSaveOptions::s_instance;

} // namespace gtest_test

void ExDocSaveOptions::SaveAsDoc()
{
    //ExStart
    //ExFor:DocSaveOptions
    //ExFor:DocSaveOptions.#ctor
    //ExFor:DocSaveOptions.#ctor(SaveFormat)
    //ExFor:DocSaveOptions.Password
    //ExFor:DocSaveOptions.SaveFormat
    //ExFor:DocSaveOptions.SaveRoutingSlip
    //ExFor:IncorrectPasswordException
    //ExSummary:Shows how to set save options for older Microsoft Word formats.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"Hello world!");
    
    auto options = System::MakeObject<Aspose::Words::Saving::DocSaveOptions>(Aspose::Words::SaveFormat::Doc);
    
    // Set a password which will protect the loading of the document by Microsoft Word or Aspose.Words.
    // Note that this does not encrypt the contents of the document in any way.
    options->set_Password(u"MyPassword");
    
    // If the document contains a routing slip, we can preserve it while saving by setting this flag to true.
    options->set_SaveRoutingSlip(true);
    
    doc->Save(get_ArtifactsDir() + u"DocSaveOptions.SaveAsDoc.doc", options);
    
    // To be able to load the document,
    // we will need to apply the password we specified in the DocSaveOptions object in a LoadOptions object.
    ASSERT_THROW(static_cast<std::function<void()>>([&doc]() -> void
    {
        doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocSaveOptions.SaveAsDoc.doc");
    })(), Aspose::Words::IncorrectPasswordException);
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>(u"MyPassword");
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocSaveOptions.SaveAsDoc.doc", loadOptions);
    
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    //ExEnd
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
    //ExSummary:Shows how to use the hard drive instead of memory when saving a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // When we save a document, various elements are temporarily stored in memory as the save operation is taking place.
    // We can use this option to use a temporary folder in the local file system instead,
    // which will reduce our application's memory overhead.
    auto options = System::MakeObject<Aspose::Words::Saving::DocSaveOptions>();
    options->set_TempFolder(get_ArtifactsDir() + u"TempFiles");
    
    // The specified temporary folder must exist in the local file system before the save operation.
    System::IO::Directory::CreateDirectory_(options->get_TempFolder());
    
    doc->Save(get_ArtifactsDir() + u"DocSaveOptions.TempFolder.doc", options);
    
    // The folder will persist with no residual contents from the load operation.
    ASSERT_EQ(0, System::IO::Directory::GetFiles(options->get_TempFolder())->get_Length());
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
    //ExSummary:Shows how to omit PictureBullet data from the document when saving.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Image bullet points.docx");
    ASSERT_FALSE(System::TestTools::IsNull(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData()));
    //ExSkip
    
    // Some word processors, such as Microsoft Word 97, are incompatible with PictureBullet data.
    // By setting a flag in the SaveOptions object,
    // we can convert all image bullet points to ordinary bullet points while saving.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::DocSaveOptions>(Aspose::Words::SaveFormat::Doc);
    saveOptions->set_SavePictureBullet(false);
    
    doc->Save(get_ArtifactsDir() + u"DocSaveOptions.PictureBullets.doc", saveOptions);
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocSaveOptions.PictureBullets.doc");
    
    ASSERT_TRUE(System::TestTools::IsNull(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData()));
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
    //ExSummary:Shows how to update a document's "Last printed" property when saving.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::DateTime lastPrinted(2019, 12, 20);
    doc->get_BuiltInDocumentProperties()->set_LastPrinted(lastPrinted);
    
    // This flag determines whether the last printed date, which is a built-in property, is updated.
    // If so, then the date of the document's most recent save operation
    // with this SaveOptions object passed as a parameter is used as the print date.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::DocSaveOptions>();
    saveOptions->set_UpdateLastPrintedProperty(isUpdateLastPrintedProperty);
    
    // In Microsoft Word 2003, this property can be found via File -> Properties -> Statistics -> Printed.
    // It can also be displayed in the document's body by using a PRINTDATE field.
    doc->Save(get_ArtifactsDir() + u"DocSaveOptions.UpdateLastPrintedProperty.doc", saveOptions);
    
    // Open the saved document, then verify the value of the property.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocSaveOptions.UpdateLastPrintedProperty.doc");
    
    if (isUpdateLastPrintedProperty)
    {
        ASSERT_NE(lastPrinted, doc->get_BuiltInDocumentProperties()->get_LastPrinted());
    }
    else
    {
        ASSERT_EQ(lastPrinted, doc->get_BuiltInDocumentProperties()->get_LastPrinted());
    }
    //ExEnd
}

namespace gtest_test
{

using ExDocSaveOptions_UpdateLastPrintedProperty_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocSaveOptions::UpdateLastPrintedProperty)>::type;

struct ExDocSaveOptions_UpdateLastPrintedProperty : public ExDocSaveOptions, public Aspose::Words::ApiExamples::ExDocSaveOptions, public ::testing::WithParamInterface<ExDocSaveOptions_UpdateLastPrintedProperty_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocSaveOptions_UpdateLastPrintedProperty, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateLastPrintedProperty(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocSaveOptions_UpdateLastPrintedProperty, ::testing::ValuesIn(ExDocSaveOptions_UpdateLastPrintedProperty::TestCases()));

} // namespace gtest_test

void ExDocSaveOptions::UpdateCreatedTimeProperty(bool isUpdateCreatedTimeProperty)
{
    //ExStart
    //ExFor:SaveOptions.UpdateCreatedTimeProperty
    //ExSummary:Shows how to update a document's "CreatedTime" property when saving.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::DateTime createdTime(2019, 12, 20);
    doc->get_BuiltInDocumentProperties()->set_CreatedTime(createdTime);
    
    // This flag determines whether the created time, which is a built-in property, is updated.
    // If so, then the date of the document's most recent save operation
    // with this SaveOptions object passed as a parameter is used as the created time.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::DocSaveOptions>();
    saveOptions->set_UpdateCreatedTimeProperty(isUpdateCreatedTimeProperty);
    
    doc->Save(get_ArtifactsDir() + u"DocSaveOptions.UpdateCreatedTimeProperty.docx", saveOptions);
    
    // Open the saved document, then verify the value of the property.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocSaveOptions.UpdateCreatedTimeProperty.docx");
    
    if (isUpdateCreatedTimeProperty)
    {
        ASSERT_NE(createdTime, doc->get_BuiltInDocumentProperties()->get_CreatedTime());
    }
    else
    {
        ASSERT_EQ(createdTime, doc->get_BuiltInDocumentProperties()->get_CreatedTime());
    }
    
    //ExEnd
}

namespace gtest_test
{

using ExDocSaveOptions_UpdateCreatedTimeProperty_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocSaveOptions::UpdateCreatedTimeProperty)>::type;

struct ExDocSaveOptions_UpdateCreatedTimeProperty : public ExDocSaveOptions, public Aspose::Words::ApiExamples::ExDocSaveOptions, public ::testing::WithParamInterface<ExDocSaveOptions_UpdateCreatedTimeProperty_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocSaveOptions_UpdateCreatedTimeProperty, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateCreatedTimeProperty(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocSaveOptions_UpdateCreatedTimeProperty, ::testing::ValuesIn(ExDocSaveOptions_UpdateCreatedTimeProperty::TestCases()));

} // namespace gtest_test

void ExDocSaveOptions::AlwaysCompressMetafiles(bool compressAllMetafiles)
{
    //ExStart
    //ExFor:DocSaveOptions.AlwaysCompressMetafiles
    //ExSummary:Shows how to change metafiles compression in a document while saving.
    // Open a document that contains a Microsoft Equation 3.0 formula.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Microsoft equation object.docx");
    
    // When we save a document, smaller metafiles are not compressed for performance reasons.
    // We can set a flag in a SaveOptions object to compress every metafile when saving.
    // Some editors such as LibreOffice cannot read uncompressed metafiles.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::DocSaveOptions>();
    saveOptions->set_AlwaysCompressMetafiles(compressAllMetafiles);
    
    doc->Save(get_ArtifactsDir() + u"DocSaveOptions.AlwaysCompressMetafiles.docx", saveOptions);
    //ExEnd
    
    int64_t testedFileLength = System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"DocSaveOptions.AlwaysCompressMetafiles.docx")->get_Length();
    
    if (compressAllMetafiles)
    {
        ASSERT_TRUE(testedFileLength < 14000);
    }
    else
    {
        ASSERT_TRUE(testedFileLength < 22000);
    }
}

namespace gtest_test
{

using ExDocSaveOptions_AlwaysCompressMetafiles_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocSaveOptions::AlwaysCompressMetafiles)>::type;

struct ExDocSaveOptions_AlwaysCompressMetafiles : public ExDocSaveOptions, public Aspose::Words::ApiExamples::ExDocSaveOptions, public ::testing::WithParamInterface<ExDocSaveOptions_AlwaysCompressMetafiles_Args>
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

TEST_P(ExDocSaveOptions_AlwaysCompressMetafiles, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AlwaysCompressMetafiles(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocSaveOptions_AlwaysCompressMetafiles, ::testing::ValuesIn(ExDocSaveOptions_AlwaysCompressMetafiles::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
