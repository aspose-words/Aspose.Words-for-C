#pragma once

#include <Aspose.Words.Cpp/AI/AiModel.h>
#include <Aspose.Words.Cpp/AI/AiModelType.h>
#include <Aspose.Words.Cpp/AI/CheckGrammarOptions.h>
#include <Aspose.Words.Cpp/AI/GoogleAiModel.h>
#include <Aspose.Words.Cpp/AI/Language.h>
#include <Aspose.Words.Cpp/AI/OpenAiModel.h>
#include <Aspose.Words.Cpp/AI/SummarizeOptions.h>
#include <Aspose.Words.Cpp/AI/SummaryLength.h>
#include <Aspose.Words.Cpp/Document.h>
#include <system/array.h>
#include <system/environment.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::AI;

namespace DocsExamples { namespace AI_powered_Features {

class WorkingWithAI : public DocsExamplesBase
{
public:
    void AiSummarize()
    {
        //ExStart:AiSummarize
        //GistId:1e379bedb2b759c1be24c64aad54d13d
        auto firstDoc = MakeObject<Document>(MyDir + u"Big document.docx");
        auto secondDoc = MakeObject<Document>(MyDir + u"Document.docx");

        String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
        // Use OpenAI or Google generative language models.
        SharedPtr<AiModel> model = System::ExplicitCast<OpenAiModel>(AiModel::Create(AiModelType::Gpt4OMini)->WithApiKey(apiKey))
                                       ->WithOrganization(u"Organization")
                                       ->WithProject(u"Project");

        auto options = MakeObject<SummarizeOptions>();

        options->set_SummaryLength(SummaryLength::Short);
        SharedPtr<Document> oneDocumentSummary = model->Summarize(firstDoc, options);
        oneDocumentSummary->Save(ArtifactsDir + u"AI.AiSummarize.One.docx");

        options->set_SummaryLength(SummaryLength::Long);
        SharedPtr<Document> multiDocumentSummary = model->Summarize(MakeArray<SharedPtr<Document>>({firstDoc, secondDoc}), options);
        multiDocumentSummary->Save(ArtifactsDir + u"AI.AiSummarize.Multi.docx");
        //ExEnd:AiSummarize
    }

    void AiTranslate()
    {
        //ExStart:AiTranslate
        //GistId:ea14b3e44c0233eecd663f783a21c4f6
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
        // Use Google generative language models.
        SharedPtr<AiModel> model = System::ExplicitCast<GoogleAiModel>(AiModel::Create(AiModelType::GeminiFlashLatest)->WithApiKey(apiKey));

        SharedPtr<Document> translatedDoc = model->Translate(doc, Language::Arabic);
        translatedDoc->Save(ArtifactsDir + u"AI.AiTranslate.docx");
        //ExEnd:AiTranslate
    }

    void AiGrammar()
    {
        //ExStart:AiGrammar
        //GistId:98a646d19cd7708ed0cd3d97b993a053
        auto doc = MakeObject<Document>(MyDir + u"Big document.docx");

        String apiKey = System::Environment::GetEnvironmentVariable(u"API_KEY");
        // Use OpenAI generative language models.
        SharedPtr<AiModel> model = AiModel::Create(AiModelType::Gpt4OMini)->WithApiKey(apiKey);

        auto grammarOptions = MakeObject<CheckGrammarOptions>();
        grammarOptions->set_ImproveStylistics(true);

        SharedPtr<Document> proofedDoc = model->CheckGrammar(doc, grammarOptions);
        proofedDoc->Save(ArtifactsDir + u"AI.AiGrammar.docx");
        //ExEnd:AiGrammar
    }
};

}} // namespace DocsExamples::AI_powered_Features
