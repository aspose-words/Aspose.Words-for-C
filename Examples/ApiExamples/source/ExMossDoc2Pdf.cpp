// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMossDoc2Pdf.h"

#include <system/globalization/culture_info.h>
#include <system/exceptions.h>
#include <system/environment.h>
#include <system/do_try_finally.h>
#include <system/date_time.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

System::String& ExMossDoc2Pdf::gInFileName()
{
    static System::String value;
    return value;
}

System::String& ExMossDoc2Pdf::gOutFileName()
{
    static System::String value;
    return value;
}

System::SharedPtr<System::IO::StreamWriter>& ExMossDoc2Pdf::gLog()
{
    static System::SharedPtr<System::IO::StreamWriter> value;
    return value;
}

void ExMossDoc2Pdf::MossDoc2Pdf(System::ArrayPtr<System::String> args)
{
    // Although SharePoint passes "-log <filename>" to us and we are
    // supposed to log there, we will use our hardcoded path to the log file for the sake of simplicity.
    // Make sure there are permissions to write into this folder.
    // The document converter will be called under the document 
    // conversion account (not sure what name), so for testing purposes, 
    // I would give the Users group write permissions into this folder.
    gLog() = System::MakeObject<System::IO::StreamWriter>(u"C:\\Aspose2Pdf\\log.txt", true);
    
    System::DoTryFinally([&] /* try-catch block */ 
    {
        try
        {
            gLog()->WriteLine(System::DateTime::get_Now().ToString(System::Globalization::CultureInfo::get_InvariantCulture()) + u" Started");
            gLog()->WriteLine(System::Environment::get_CommandLine());
            
            ParseCommandLine(args);
            
            // Uncomment the code below when you have purchased a license for Aspose.Words.
            // You need to deploy the license in the same folder as your 
            // executable, alternatively you can add the license file as an 
            // embedded resource to your project.
            // Set license for Aspose.Words.
            // Aspose.Words.License wordsLicense = new Aspose.Words.License();
            // wordsLicense.SetLicense("Aspose.Total.lic");
            
            ConvertDoc2Pdf(gInFileName(), gOutFileName());
        }
        catch (System::Exception& e)
        {
            gLog()->WriteLine(e->get_Message());
            System::Environment::set_ExitCode(100);
        }
        
    }
    , [&] /* finally block */ 
    {
        gLog()->Close();
    });
    
}

void ExMossDoc2Pdf::ParseCommandLine(System::ArrayPtr<System::String> args)
{
    int32_t i = 0;
    while (i < args->get_Length())
    {
        System::String s = args[i];
        const System::String& switch_value_0 = s.ToLower();
        if (switch_value_0 == u"-in")
        {
            i++;
            gInFileName() = args[i];
        }
        else if (switch_value_0 == u"-out")
        {
            i++;
            gOutFileName() = args[i];
        }
        else if (switch_value_0 == u"-config")
        {
            i++;
        }
        else if (switch_value_0 == u"-log")
        {
            i++;
        }
        else 
        {
            throw System::Exception(System::String(u"Unknown command line argument: ") + s);
        }
        
        i++;
    }
}

void ExMossDoc2Pdf::ConvertDoc2Pdf(System::String inFileName, System::String outFileName)
{
    // You can load not only DOC here, but any format supported by
    // Aspose.Words: DOC, DOCX, RTF, WordML, HTML, MHTML, ODT etc.
    auto doc = System::MakeObject<Aspose::Words::Document>(inFileName);
    
    doc->Save(outFileName, System::MakeObject<Aspose::Words::Saving::PdfSaveOptions>());
}

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
