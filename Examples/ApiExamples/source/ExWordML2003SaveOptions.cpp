// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExWordML2003SaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/io/file.h>
#include <system/environment.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Saving/WordML2003SaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3800896039u, ::Aspose::Words::ApiExamples::ExWordML2003SaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExWordML2003SaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExWordML2003SaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExWordML2003SaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExWordML2003SaveOptions> ExWordML2003SaveOptions::s_instance;

} // namespace gtest_test

void ExWordML2003SaveOptions::PrettyFormat(bool prettyFormat)
{
    //ExStart
    //ExFor:WordML2003SaveOptions
    //ExFor:WordML2003SaveOptions.SaveFormat
    //ExSummary:Shows how to manage output document's raw content.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
    // to modify how we save the document to the WordML save format.
    auto options = System::MakeObject<Aspose::Words::Saving::WordML2003SaveOptions>();
    
    ASSERT_EQ(Aspose::Words::SaveFormat::WordML, options->get_SaveFormat());
    
    // Set the "PrettyFormat" property to "true" to apply tab character indentation and
    // newlines to make the output document's raw content easier to read.
    // Set the "PrettyFormat" property to "false" to save the document's raw content in one continuous body of the text.
    options->set_PrettyFormat(prettyFormat);
    
    doc->Save(get_ArtifactsDir() + u"WordML2003SaveOptions.PrettyFormat.xml", options);
    
    System::String fileContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"WordML2003SaveOptions.PrettyFormat.xml");
    System::String newLine = System::Environment::get_NewLine();
    if (prettyFormat)
    {
        ASSERT_TRUE(fileContents.Contains(System::String::Format(u"<o:DocumentProperties>{0}\t\t", newLine) + System::String::Format(u"<o:Revision>1</o:Revision>{0}\t\t", newLine) + System::String::Format(u"<o:TotalTime>0</o:TotalTime>{0}\t\t", newLine) + System::String::Format(u"<o:Pages>1</o:Pages>{0}\t\t", newLine) + System::String::Format(u"<o:Words>0</o:Words>{0}\t\t", newLine) + System::String::Format(u"<o:Characters>0</o:Characters>{0}\t\t", newLine) + System::String::Format(u"<o:Lines>1</o:Lines>{0}\t\t", newLine) + System::String::Format(u"<o:Paragraphs>1</o:Paragraphs>{0}\t\t", newLine) + System::String::Format(u"<o:CharactersWithSpaces>0</o:CharactersWithSpaces>{0}\t\t", newLine) + System::String::Format(u"<o:Version>11.5606</o:Version>{0}\t", newLine) + u"</o:DocumentProperties>"));
    }
    else
    {
        ASSERT_TRUE(fileContents.Contains(System::String(u"<o:DocumentProperties><o:Revision>1</o:Revision><o:TotalTime>0</o:TotalTime><o:Pages>1</o:Pages>") + u"<o:Words>0</o:Words><o:Characters>0</o:Characters><o:Lines>1</o:Lines><o:Paragraphs>1</o:Paragraphs>" + u"<o:CharactersWithSpaces>0</o:CharactersWithSpaces><o:Version>11.5606</o:Version></o:DocumentProperties>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExWordML2003SaveOptions_PrettyFormat_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExWordML2003SaveOptions::PrettyFormat)>::type;

struct ExWordML2003SaveOptions_PrettyFormat : public ExWordML2003SaveOptions, public Aspose::Words::ApiExamples::ExWordML2003SaveOptions, public ::testing::WithParamInterface<ExWordML2003SaveOptions_PrettyFormat_Args>
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

TEST_P(ExWordML2003SaveOptions_PrettyFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PrettyFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExWordML2003SaveOptions_PrettyFormat, ::testing::ValuesIn(ExWordML2003SaveOptions_PrettyFormat::TestCases()));

} // namespace gtest_test

void ExWordML2003SaveOptions::MemoryOptimization(bool memoryOptimization)
{
    //ExStart
    //ExFor:WordML2003SaveOptions
    //ExSummary:Shows how to manage memory optimization.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
    // to modify how we save the document to the WordML save format.
    auto options = System::MakeObject<Aspose::Words::Saving::WordML2003SaveOptions>();
    
    // Set the "MemoryOptimization" flag to "true" to decrease memory consumption
    // during the document's saving operation at the cost of a longer saving time.
    // Set the "MemoryOptimization" flag to "false" to save the document normally.
    options->set_MemoryOptimization(memoryOptimization);
    
    doc->Save(get_ArtifactsDir() + u"WordML2003SaveOptions.MemoryOptimization.xml", options);
    //ExEnd
}

namespace gtest_test
{

using ExWordML2003SaveOptions_MemoryOptimization_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExWordML2003SaveOptions::MemoryOptimization)>::type;

struct ExWordML2003SaveOptions_MemoryOptimization : public ExWordML2003SaveOptions, public Aspose::Words::ApiExamples::ExWordML2003SaveOptions, public ::testing::WithParamInterface<ExWordML2003SaveOptions_MemoryOptimization_Args>
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

TEST_P(ExWordML2003SaveOptions_MemoryOptimization, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MemoryOptimization(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExWordML2003SaveOptions_MemoryOptimization, ::testing::ValuesIn(ExWordML2003SaveOptions_MemoryOptimization::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
