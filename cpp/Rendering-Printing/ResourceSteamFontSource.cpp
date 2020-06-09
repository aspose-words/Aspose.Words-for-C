#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/StreamFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/SystemFontSource.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
// ExStart:ResourceSteamFontSource
namespace
{
    class ResourceSteamFontSourceExample : public StreamFontSource
    {
        typedef ResourceSteamFontSourceExample ThisType;
        typedef Aspose::Words::Fonts::StreamFontSource BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);

    public:
        System::SharedPtr<System::IO::Stream> OpenFontDataStream() override;
    };

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
// ExEnd:ResourceSteamFontSource