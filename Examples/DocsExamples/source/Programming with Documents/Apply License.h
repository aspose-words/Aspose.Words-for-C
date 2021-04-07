#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Licensing/License.h>
#include <system/array.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/memory_stream.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents {

class ApplyLicense : public DocsExamplesBase
{
public:
    void ApplyLicenseFromFile()
    {
        //ExStart:ApplyLicenseFromFile
        auto license = MakeObject<License>();

        // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        // You can also use the additional overload to load a license from a stream, this is useful,
        // for instance, when the license is stored as an embedded resource.
        try
        {
            license->SetLicense(u"Aspose.Words.Cpp.lic");

            std::cout << "License set successfully." << std::endl;
        }
        catch (System::Exception& e)
        {
            // We do not ship any license with this example,
            // visit the Aspose site to obtain either a temporary or permanent license.
            std::cout << (String(u"\nThere was an error setting the license: ") + e->get_Message()) << std::endl;
        }

        //ExEnd:ApplyLicenseFromFile
    }

    void ApplyLicenseFromStream()
    {
        //ExStart:ApplyLicenseFromStream
        auto license = MakeObject<License>();

        try
        {
            license->SetLicense(MakeObject<System::IO::MemoryStream>(System::IO::File::ReadAllBytes(u"Aspose.Words.Cpp.lic")));

            std::cout << "License set successfully." << std::endl;
        }
        catch (System::Exception& e)
        {
            // We do not ship any license with this example,
            // visit the Aspose site to obtain either a temporary or permanent license.
            std::cout << (String(u"\nThere was an error setting the license: ") + e->get_Message()) << std::endl;
        }

        //ExEnd:ApplyLicenseFromStream
    }
};

}} // namespace DocsExamples::Programming_with_Documents
