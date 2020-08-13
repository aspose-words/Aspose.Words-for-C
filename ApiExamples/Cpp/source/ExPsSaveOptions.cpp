// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPsSaveOptions.h"

#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples {

namespace gtest_test
{

class ExPsSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExPsSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExPsSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExPsSaveOptions> ExPsSaveOptions::s_instance;

} // namespace gtest_test

void ExPsSaveOptions::UseBookFoldPrintingSettings()
{
    //ExStart
    //ExFor:PsSaveOptions
    //ExFor:PsSaveOptions.SaveFormat
    //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
    //ExSummary:Shows how to create a bookfold in the PostScript format.
    auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

    // Configure both page setup and PsSaveOptions to create a book fold
    for (auto s : System::IterateOver<Section>(doc->get_Sections()))
    {
        s->get_PageSetup()->set_MultiplePages(Aspose::Words::Settings::MultiplePagesType::BookFoldPrinting);
    }

    auto saveOptions = MakeObject<PsSaveOptions>();
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Ps);
    saveOptions->set_UseBookFoldPrintingSettings(true);

    // In order to make a booklet, we will need to print this document, stack the pages
    // in the order they come out of the printer and then fold down the middle
    doc->Save(ArtifactsDir + u"PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPsSaveOptions, UseBookFoldPrintingSettings)
{
    s_instance->UseBookFoldPrintingSettings();
}

} // namespace gtest_test

} // namespace ApiExamples
