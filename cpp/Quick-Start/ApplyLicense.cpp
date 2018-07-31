#include "stdafx.h"
#include "examples.h"

#include "Licensing/License.h"

using namespace Aspose::Words;

void ApplyLicense()
{
    auto license = System::MakeObject<License>();

    // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
    // You can also use the additional overload to load a license from a stream, this is useful for instance when the 
    // License is stored as an embedded resource 
    try
    {
        license->SetLicense(u"D:/Temp/Aspose.Words.CPP.lic");
        std::cout << "License set successfully.\n";
    }

    catch (const System::Exception& e)
    {
        // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
        std::cout << "\nThere was an error setting the license:" << e.get_Message().ToUtf8String() << '\n';
    }
}