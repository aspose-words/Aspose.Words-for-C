#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fonts/FontSettings.h>
#include <Model/Fonts/FontSourceBase.h>
#include <Model/Fonts/StreamFontSource.h>
#include <Model/Fonts/SystemFontSource.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

namespace
{
    class ResourceSteamFontSourceExample : public StreamFontSource
    {
        typedef ResourceSteamFontSourceExample ThisType;
        typedef Aspose::Words::Fonts::StreamFontSource BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        virtual System::SharedPtr<System::IO::Stream> OpenFontDataStream();
    };

    RTTI_INFO_IMPL_HASH(3895389092u, ResourceSteamFontSourceExample, ThisTypeBaseTypesInfo);

    System::SharedPtr<System::IO::Stream> ResourceSteamFontSourceExample::OpenFontDataStream()
    {
        return System::Reflection::Assembly::GetExecutingAssembly()->GetManifestResourceStream(u"resourceName");
    }
}

void ResourceSteamFontSource()
{
    std::cout << "ResourceSteamFontSource example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetInputDataDir_RenderingAndPrinting();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Rendering.doc");

    // FontSettings.SetFontSources instead.
    FontSettings::get_DefaultInstance()->SetFontsSources(System::MakeArray<System::SharedPtr<FontSourceBase>>({System::MakeObject<SystemFontSource>(), System::MakeObject<ResourceSteamFontSourceExample>()}));
    System::String outputPath = GetOutputDataDir_RenderingAndPrinting() + u"ResourceSteamFontSource.pdf";
    doc->Save(outputPath);
    std::cout << "ResourceSteamFontSource example finished." << std::endl << std::endl;
}