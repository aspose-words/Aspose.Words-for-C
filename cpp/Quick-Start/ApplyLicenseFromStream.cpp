#include "stdafx.h"
#include "examples.h"

#include <Licensing/License.h>

using namespace Aspose::Words;

void ApplyLicenseFromStream()
{
    std::cout << "ApplyLicenseFromStream example started." << std::endl;
    //ExStart:ApplyLicenseFromStream
    System::SharedPtr<License> license = System::MakeObject<License>();

    try
    {
        // Initializes a license from a stream 
        System::SharedPtr<System::IO::MemoryStream> stream = System::MakeObject<System::IO::MemoryStream>(System::IO::File::ReadAllBytes(u"Aspose.Words.Cpp.lic"));
        license->SetLicense(stream);
        std::cout << "License set successfully." << std::endl;
    }
    catch (System::Exception& e)
    {
        // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
        std::cout << "There was an error setting the license: " << e.get_Message().ToUtf8String() << std::endl;
    }

    //ExEnd:ApplyLicenseFromStream
    std::cout << "ApplyLicenseFromStream example finished." << std::endl << std::endl;
}