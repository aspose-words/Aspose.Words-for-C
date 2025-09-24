// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExAI.h"

#include <system/object_ext.h>
#include <system/environment.h>
#include <system/array.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/AI/SummaryLength.h>
#include <Aspose.Words.Cpp/Model/AI/SummarizeOptions.h>
#include <Aspose.Words.Cpp/Model/AI/OpenAi/Models/OpenAiModel.h>
#include <Aspose.Words.Cpp/Model/AI/Language.h>
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
    System::SharedPtr<Aspose::Words::AI::AiModel> model = Aspose::Words::AI::AiModel::Create(Aspose::Words::AI::AiModelType::Gemini15Flash)->WithApiKey(apiKey);
    
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

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
