#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Licensing/License.h>

using namespace Aspose::Words;

void ApplyLicense()
{
    std::cout << "ApplyLicense example started." << std::endl;
    System::SharedPtr<License> license = System::MakeObject<License>();
    // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
    // You can also use the additional overload to load a license from a stream, this is useful for instance when the
    // License is stored as an embedded resource
    try
    {
        license->SetLicense(u"Aspose.Words.lic");
        std::cout << "License set successfully." << std::endl;
    }

    catch (const System::Exception& e)
    {
        // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
        std::cout << "There was an error setting the license:" << e->get_Message().ToUtf8String() << std::endl;
    }
    std::cout << "ApplyLicense example finished." << std::endl << std::endl;
}