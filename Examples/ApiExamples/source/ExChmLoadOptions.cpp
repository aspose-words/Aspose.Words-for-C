// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExChmLoadOptions.h"

#include <system/io/memory_stream.h>
#include <system/io/file.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Loading/ChmLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Loading;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4129072864u, ::Aspose::Words::ApiExamples::ExChmLoadOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExChmLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExChmLoadOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExChmLoadOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExChmLoadOptions> ExChmLoadOptions::s_instance;

} // namespace gtest_test

void ExChmLoadOptions::OriginalFileName()
{
    //ExStart
    //ExFor:ChmLoadOptions
    //ExFor:ChmLoadOptions.#ctor
    //ExFor:ChmLoadOptions.OriginalFileName
    //ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
    // Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
    // so file links don't work after saving it to HTML.
    // We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::ChmLoadOptions>();
    loadOptions->set_OriginalFileName(u"amhelp.chm");
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::IO::File::ReadAllBytes(get_MyDir() + u"Document with ms-its links.chm")), loadOptions);
    
    doc->Save(get_ArtifactsDir() + u"ExChmLoadOptions.OriginalFileName.html");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExChmLoadOptions, OriginalFileName)
{
    s_instance->OriginalFileName();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
