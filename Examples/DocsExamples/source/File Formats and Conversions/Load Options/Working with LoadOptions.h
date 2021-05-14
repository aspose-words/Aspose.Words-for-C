#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Loading/ResourceLoadingArgs.h>
#include <Aspose.Words.Cpp/Loading/ResourceType.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningSource.h>
#include <Aspose.Words.Cpp/WarningType.h>
#include <drawing/image.h>
#include <drawing/image_converter.h>
#include <system/array.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/object_ext.h>
#include <system/text/encoding.h>
#include <system/type_info.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Load_Options {

class WorkingWithLoadOptions : public DocsExamplesBase
{
public:
    class DocumentLoadingWarningCallback : public IWarningCallback
    {
    public:
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // Prints warnings and their details as they arise during document loading.
            std::cout << String::Format(u"WARNING: {0}, source: {1}", info->get_WarningType(), info->get_Source()) << std::endl;
            std::cout << "\tDescription: " << info->get_Description() << std::endl;
        }
    };

private:
    class HtmlLinkedResourceLoadingCallback : public IResourceLoadingCallback
    {
    public:
        ResourceLoadingAction ResourceLoading(SharedPtr<ResourceLoadingArgs> args) override
        {
            switch (args->get_ResourceType())
            {
            case ResourceType::CssStyleSheet: {
                std::cout << "External CSS Stylesheet found upon loading: " << args->get_OriginalUri() << std::endl;

                // CSS file will don't used in the document.
                return ResourceLoadingAction::Skip;
            }

            case ResourceType::Image: {
                // Replaces all images with a substitute.
                SharedPtr<System::Drawing::Image> newImage = System::Drawing::Image::FromFile(ImagesDir + u"Logo.jpg");

                auto converter = MakeObject<System::Drawing::ImageConverter>();
                ArrayPtr<uint8_t> imageBytes =
                    System::DynamicCast<System::Array<uint8_t>>(converter->ConvertTo(newImage, System::ObjectExt::GetType<System::Array<uint8_t>>()));

                args->SetData(imageBytes);

                // New images will be used instead of presented in the document.
                return ResourceLoadingAction::UserProvided;
            }

            case ResourceType::Document: {
                std::cout << "External document found upon loading: " << args->get_OriginalUri() << std::endl;

                // Will be used as usual.
                return ResourceLoadingAction::Default;
            }

            default:
                throw System::InvalidOperationException(u"Unexpected ResourceType value.");
            }
        }
    };

public:
    void UpdateDirtyFields()
    {
        //ExStart:UpdateDirtyFields
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_UpdateDirtyFields(true);

        auto doc = MakeObject<Document>(MyDir + u"Dirty field.docx", loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithLoadOptions.UpdateDirtyFields.docx");
        //ExEnd:UpdateDirtyFields
    }

    void LoadEncryptedDocument()
    {
        //ExStart:LoadSaveEncryptedDoc
        //ExStart:OpenEncryptedDocument
        auto doc = MakeObject<Document>(MyDir + u"Encrypted.docx", MakeObject<LoadOptions>(u"docPassword"));
        //ExEnd:OpenEncryptedDocument

        doc->Save(ArtifactsDir + u"WorkingWithLoadOptions.LoadAndSaveEncryptedOdt.odt", MakeObject<OdtSaveOptions>(u"newPassword"));
        //ExEnd:LoadSaveEncryptedDoc
    }

    void ConvertShapeToOfficeMath()
    {
        //ExStart:ConvertShapeToOfficeMath
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_ConvertShapeToOfficeMath(true);

        auto doc = MakeObject<Document>(MyDir + u"Office math.docx", loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithLoadOptions.ConvertShapeToOfficeMath.docx", SaveFormat::Docx);
        //ExEnd:ConvertShapeToOfficeMath
    }

    void SetMsWordVersion()
    {
        //ExStart:SetMSWordVersion
        // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
        // and change the loading version to Microsoft Word 2010.
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_MswVersion(MsWordVersion::Word2010);

        auto doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithLoadOptions.SetMsWordVersion.docx");
        //ExEnd:SetMSWordVersion
    }

    void UseTempFolder()
    {
        //ExStart:UseTempFolder
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_TempFolder(ArtifactsDir);

        auto doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);
        //ExEnd:UseTempFolder
    }

    void WarningCallback()
    {
        //ExStart:WarningCallback
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_WarningCallback(MakeObject<WorkingWithLoadOptions::DocumentLoadingWarningCallback>());

        auto doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);
        //ExEnd:WarningCallback
    }

    void ResourceLoadingCallback()
    {
        //ExStart:ResourceLoadingCallback
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_ResourceLoadingCallback(MakeObject<WorkingWithLoadOptions::HtmlLinkedResourceLoadingCallback>());

        // When we open an Html document, external resources such as references to CSS stylesheet files
        // and external images will be handled customarily by the loading callback as the document is loaded.
        auto doc = MakeObject<Document>(MyDir + u"Images.html", loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithLoadOptions.ResourceLoadingCallback.pdf");
        //ExEnd:ResourceLoadingCallback
    }

    void LoadWithEncoding()
    {
        //ExStart:LoadWithEncoding
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_Encoding(System::Text::Encoding::get_UTF7());

        auto doc = MakeObject<Document>(MyDir + u"Encoded in UTF-7.txt", loadOptions);
        //ExEnd:LoadWithEncoding
    }

    void ConvertMetafilesToPng()
    {
        //ExStart:ConvertMetafilesToPng
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_ConvertMetafilesToPng(true);

        auto doc = MakeObject<Document>(MyDir + u"WMF with image.docx", loadOptions);
        //ExEnd:ConvertMetafilesToPng
    }

    void LoadChm()
    {
        //ExStart:LoadCHM
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_Encoding(System::Text::Encoding::GetEncoding(u"windows-1251"));

        auto doc = MakeObject<Document>(MyDir + u"HTML help.chm", loadOptions);
        //ExEnd:LoadCHM
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Load_Options
