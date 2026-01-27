// Copyright (c) 2001-2026 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExAI.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/environment.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/AI/SummaryLength.h>
#include <Aspose.Words.Cpp/Model/AI/SummarizeOptions.h>
#include <Aspose.Words.Cpp/Model/AI/OpenAi/Models/OpenAiModel.h>
#include <Aspose.Words.Cpp/Model/AI/Language.h>
#include <Aspose.Words.Cpp/Model/AI/Google/Models/GoogleAiModel.h>
#include <Aspose.Words.Cpp/Model/AI/CheckGrammarOptions.h>
#include <Aspose.Words.Cpp/Model/AI/AiModelType.h>
#include <Aspose.Words.Cpp/Model/AI/AiModel.h>


using namespace Aspose::Words::AI;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2792701112u, ::Aspose::Words::ApiExamples::ExAI, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExAI : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExAI> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExAI>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExAI> ExAI::s_instance;

} // namespace gtest_test

void ExAI::AiSummarize()
{
    //ExStart:AiSummarize
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:GoogleAiModel
    //ExFor:OpenAiModel
    //ExFor:OpenAiModel.WithOrganization(String)
    //ExFor:OpenAiModel.WithProject(String)
    //ExFor:AiModel
    //ExFor:AiModel.Summarize(Document, SummarizeOptions)
    //ExFor:AiModel.Summarize(Document[], SummarizeOptions)
    //ExFor:AiModel.Create(AiModelType)
    //ExFor:AiModel.WithApiKey(String)
    //ExFor:AiModelType
    //ExFor:SummarizeOptions
    //ExFor:SummarizeOptions.#ctor
    //ExFor:SummarizeOptions.SummaryLength
    //ExFor:SummaryLength
    //ExSummary:Shows how to summarize text using OpenAI and Google models.
    auto firstDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    auto secondDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
    // Use OpenAI or Google generative language models.
    System::SharedPtr<Aspose::Words::AI::AiModel> model = (System::ExplicitCast<Aspose::Words::AI::OpenAiModel>(Aspose::Words::AI::AiModel::Create(Aspose::Words::AI::AiModelType::Gpt4OMini)->WithApiKey(apiKey)))->WithOrganization(u"Organization")->WithProject(u"Project");
    
    auto options = System::MakeObject<Aspose::Words::AI::SummarizeOptions>();
    
    options->set_SummaryLength(Aspose::Words::AI::SummaryLength::Short);
    System::SharedPtr<Aspose::Words::Document> oneDocumentSummary = model->Summarize(firstDoc, options);
    oneDocumentSummary->Save(get_ArtifactsDir() + u"AI.AiSummarize.One.docx");
    
    options->set_SummaryLength(Aspose::Words::AI::SummaryLength::Long);
    System::SharedPtr<Aspose::Words::Document> multiDocumentSummary = model->Summarize(System::MakeArray<System::SharedPtr<Aspose::Words::Document>>({firstDoc, secondDoc}), options);
    multiDocumentSummary->Save(get_ArtifactsDir() + u"AI.AiSummarize.Multi.docx");
    //ExEnd:AiSummarize
}

namespace gtest_test
{

TEST_F(ExAI, DISABLED_AiSummarize)
{
    s_instance->AiSummarize();
}

} // namespace gtest_test

void ExAI::AiTranslate()
{
    //ExStart:AiTranslate
    //GistId:695136dbbe4f541a8a0a17b3d3468689
    //ExFor:AiModel.Translate(Document, AI.Language)
    //ExFor:AI.Language
    //ExSummary:Shows how to translate text using Google models.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
    // Use Google generative language models.
    System::SharedPtr<Aspose::Words::AI::AiModel> model = Aspose::Words::AI::AiModel::Create(Aspose::Words::AI::AiModelType::GeminiFlashLatest)->WithApiKey(apiKey);
    
    System::SharedPtr<Aspose::Words::Document> translatedDoc = model->Translate(doc, Aspose::Words::AI::Language::Arabic);
    translatedDoc->Save(get_ArtifactsDir() + u"AI.AiTranslate.docx");
    //ExEnd:AiTranslate
}

namespace gtest_test
{

TEST_F(ExAI, DISABLED_AiTranslate)
{
    s_instance->AiTranslate();
}

} // namespace gtest_test

void ExAI::AiGrammar()
{
    //ExStart:AiGrammar
    //GistId:f86d49dc0e6781b93e576539a01e6ca2
    //ExFor:AiModel.CheckGrammar(Document, CheckGrammarOptions)
    //ExFor:CheckGrammarOptions
    //ExSummary:Shows how to check the grammar of a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    System::String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
    // Use OpenAI generative language models.
    System::SharedPtr<Aspose::Words::AI::AiModel> model = Aspose::Words::AI::AiModel::Create(Aspose::Words::AI::AiModelType::Gpt4OMini)->WithApiKey(apiKey);
    
    auto grammarOptions = System::MakeObject<Aspose::Words::AI::CheckGrammarOptions>();
    grammarOptions->set_ImproveStylistics(true);
    
    System::SharedPtr<Aspose::Words::Document> proofedDoc = model->CheckGrammar(doc, grammarOptions);
    proofedDoc->Save(get_ArtifactsDir() + u"AI.AiGrammar.docx");
    //ExEnd:AiGrammar
}

namespace gtest_test
{

TEST_F(ExAI, DISABLED_AiGrammar)
{
    s_instance->AiGrammar();
}

} // namespace gtest_test

void ExAI::ChangeDefaultUrl()
{
    //ExStart:ChangeDefaultUrl
    //GistId:bd7947d9ad5eb092f532604cb15f593b
    //ExFor:AiModel.Url
    //ExSummary:Shows how to change model default url.
    System::String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
    System::SharedPtr<Aspose::Words::AI::AiModel> model = Aspose::Words::AI::AiModel::Create(Aspose::Words::AI::AiModelType::Gpt4OMini)->WithApiKey(apiKey);
    // Default value "https://api.openai.com/".
    model->set_Url(u"https://my.a.com/");
    //ExEnd:ChangeDefaultUrl
    
    ASSERT_EQ(u"https://my.a.com/", model->get_Url());
}

namespace gtest_test
{

TEST_F(ExAI, ChangeDefaultUrl)
{
    s_instance->ChangeDefaultUrl();
}

} // namespace gtest_test

void ExAI::ChangeDefaultTimeout()
{
    //ExStart:ChangeDefaultTimeout
    //GistId:bd7947d9ad5eb092f532604cb15f593b
    //ExFor:AiModel.Timeout
    //ExSummary:Shows how to change model default timeout.
    System::String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
    System::SharedPtr<Aspose::Words::AI::AiModel> model = Aspose::Words::AI::AiModel::Create(Aspose::Words::AI::AiModelType::Gpt4OMini)->WithApiKey(apiKey);
    // Default value 100000ms.
    model->set_Timeout(250000);
    //ExEnd:ChangeDefaultTimeout
    
    ASSERT_EQ(250000, model->get_Timeout());
}

namespace gtest_test
{

TEST_F(ExAI, ChangeDefaultTimeout)
{
    s_instance->ChangeDefaultTimeout();
}

} // namespace gtest_test

void ExAI::Gemini()
{
    //ExStart:Gemini
    //GistId:0da8468118377c4860b28603bc95ffe6
    //ExFor:GoogleAiModel
    //ExFor:GoogleAiModel.#ctor(String)
    //ExFor:GoogleAiModel.#ctor(String, String)
    //ExSummary:Shows how to use google AI model.
    System::String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
    auto model = System::MakeObject<Aspose::Words::AI::GoogleAiModel>(u"gemini-flash-latest", apiKey);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    auto summarizeOptions = System::MakeObject<Aspose::Words::AI::SummarizeOptions>();
    summarizeOptions->set_SummaryLength(Aspose::Words::AI::SummaryLength::VeryShort);
    System::SharedPtr<Aspose::Words::Document> summary = model->Summarize(doc, summarizeOptions);
    //ExEnd:Gemini
}

namespace gtest_test
{

TEST_F(ExAI, DISABLED_Gemini)
{
    s_instance->Gemini();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
