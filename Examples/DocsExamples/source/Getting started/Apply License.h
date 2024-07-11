#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/License.h>
#include <system/array.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/memory_stream.h>
#include <windows.h>

#include "DocsExamplesBase.h"
#include <Aspose.Words.Cpp/Document.h>

#define IDR_ASPOSE_WORDS_LIC            101

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
        //GistId:22914ce1f4ad49bfda8f88f5d192fe8e
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
        //GistId:22914ce1f4ad49bfda8f88f5d192fe8e
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
#ifdef _WIN32
    void ApplyLicenseFromEmbeddedResourceWindows()
    {
        //ExStart:ApplyLicenseFromEmbeddedResourceWindows
        //GistId:22914ce1f4ad49bfda8f88f5d192fe8e
        auto hResource = FindResource(nullptr, MAKEINTRESOURCEA(IDR_ASPOSE_WORDS_LIC), RT_RCDATA);
        auto hMemory = LoadResource(nullptr, hResource);

        auto size = SizeofResource(nullptr, hResource);
        auto ptr = LockResource(hMemory);

        auto licResource = System::MakeArray<uint8_t>(size);
        std::copy_n(static_cast<const uint8_t*>(ptr), size, licResource->begin());
        FreeResource(hMemory);

        auto license = System::MakeObject<License>();

        try
        {
            license->SetLicense(MakeObject<System::IO::MemoryStream>(licResource));
            std::cout << "License set successfully." << std::endl;
        }
        catch (System::Exception& e)
        {
            std::cout << (String(u"\nThere was an error setting the license: ") + e->get_Message()) << std::endl;
        }
        //ExEnd:ApplyLicenseFromEmbeddedResourceWindows
    }
#elif __linux__
    void ApplyLicenseFromEmbeddedResourceLinux()
    {
        //ExStart:ApplyLicenseFromEmbeddedResourceLinux
        //GistId:22914ce1f4ad49bfda8f88f5d192fe8e
        // A file named Aspose.Words.lic is 'imported' into an object file 
        // using the following command:
        //
        //      ld -r -b binary -o aspose.words.lic.o Aspose.Words.lic
        //
        // That creates an object file named "aspose.words.lic.o" with the following 
        // symbols:
        //
        //      _binary_aspose_words_lic_start
        //      _binary_aspose_words_lic_end
        //      _binary_aspose_words_lic_size
        //
        // Note that the symbols are addresses 

        extern uint8_t _binary_aspose_words_lic_start[];
        extern uint8_t _binary_aspose_words_lic_end[];
        extern uint8_t _binary_aspose_words_lic_size[];

        std::ptrdiff_t size = _binary_aspose_words_lic_end - _binary_aspose_words_lic_start;

        auto licResource = System::MakeArray<uint8_t>(size);
        std::copy(_binary_aspose_words_lic_start, _binary_aspose_words_lic_end, licResource->begin());

        auto license = MakeObject<License>();

        try
        {
            license->SetLicense(MakeObject<System::IO::MemoryStream>(licResource));
            std::cout << "License set successfully." << std::endl;
        }
        catch (System::Exception& e)
        {
            std::cout << (String(u"\nThere was an error setting the license: ") + e->get_Message()) << std::endl;
        }
        //ExEnd:ApplyLicenseFromEmbeddedResourceLinux
    }
#endif

};

}} // namespace DocsExamples::Programming_with_Documents
