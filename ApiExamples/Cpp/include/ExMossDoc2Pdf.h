#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/date_time.h>
#include <system/environment.h>
#include <system/exceptions.h>
#include <system/globalization/culture_info.h>
#include <system/io/stream_writer.h>
#include <system/object.h>
#include <system/scope_guard.h>
#include <system/shared_ptr.h>
#include <system/string.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

/// <summary>
/// DOC2PDF document converter for SharePoint.
/// Uses Aspose.Words to perform the conversion.
/// </summary>
class ExMossDoc2Pdf
{
public:
    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    static void MossDoc2Pdf(ArrayPtr<String> args)
    {
        // Although SharePoint passes "-log <filename>" to us and we are
        // supposed to log there, for the sake of simplicity, we will use
        // our own hard coded path to the log file.
        // Make sure there are permissions to write into this folder.
        // The document converter will be called under the document
        // conversion account (not sure what name), so for testing purposes
        // I would give the Users group write permissions into this folder.
        gLog = MakeObject<System::IO::StreamWriter>(u"C:\\Aspose2Pdf\\log.txt", true);

        {
            auto __finally_guard_0 = ::System::MakeScopeGuard([]() { gLog->Close(); });

            try
            {
                gLog->WriteLine(System::DateTime::get_Now().ToString(System::Globalization::CultureInfo::get_InvariantCulture()) + u" Started");
                gLog->WriteLine(System::Environment::get_CommandLine());

                ParseCommandLine(args);

                // Uncomment the code below when you have purchased a license for Aspose.Words.
                // You need to deploy the license in the same folder as your
                // executable, alternatively you can add the license file as an
                // embedded resource to your project.
                // // Set license for Aspose.Words.
                // Aspose.Words.License wordsLicense = new Aspose.Words.License();
                // wordsLicense.SetLicense("Aspose.Total.lic");

                ConvertDoc2Pdf(gInFileName, gOutFileName);
            }
            catch (System::Exception& e)
            {
                gLog->WriteLine(e->get_Message());
                System::Environment::set_ExitCode(100);
            }
        }
    }

private:
    static void ParseCommandLine(ArrayPtr<String> args)
    {
        int i = 0;
        while (i < args->get_Length())
        {
            String s = args[i];
            const String& switch_value_0 = s.ToLower();
            if (switch_value_0 == u"-in")
            {
                i++;
                gInFileName = args[i];
            }
            else if (switch_value_0 == u"-out")
            {
                i++;
                gOutFileName = args[i];
            }
            else if (switch_value_0 == u"-config")
            {
                i++;
            }
            else if (switch_value_0 == u"-log")
            {
                i++;
            }
            else if (true)
            {
                throw System::Exception(String(u"Unknown command line argument: ") + s);
            }

            i++;
        }
    }

    static void ConvertDoc2Pdf(String inFileName, String outFileName)
    {
        // You can load not only DOC here, but any format supported by
        // Aspose.Words: DOC, DOCX, RTF, WordML, HTML, MHTML, ODT etc.
        auto doc = MakeObject<Document>(inFileName);

        doc->Save(outFileName, MakeObject<PdfSaveOptions>());
    }

    static String gInFileName;
    static String gOutFileName;
    static SharedPtr<System::IO::StreamWriter> gLog;
};

} // namespace ApiExamples
