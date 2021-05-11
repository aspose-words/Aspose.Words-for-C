#pragma once

#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/FileFormatInfo.h>
#include <Aspose.Words.Cpp/FileFormatUtil.h>
#include <Aspose.Words.Cpp/LoadFormat.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/convert.h>
#include <system/enumerator_adapter.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/path.h>
#include <system/linq/enumerable.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace File_Formats_and_Conversions {

class WorkingWithFileFormat : public DocsExamplesBase
{
public:
    void DetectFileFormat()
    {
        //ExStart:CheckFormatCompatibility
        String supportedDir = ArtifactsDir + u"Supported";
        String unknownDir = ArtifactsDir + u"Unknown";
        String encryptedDir = ArtifactsDir + u"Encrypted";
        String pre97Dir = ArtifactsDir + u"Pre97";

        // Create the directories if they do not already exist.
        if (System::IO::Directory::Exists(supportedDir) == false)
        {
            System::IO::Directory::CreateDirectory_(supportedDir);
        }
        if (System::IO::Directory::Exists(unknownDir) == false)
        {
            System::IO::Directory::CreateDirectory_(unknownDir);
        }
        if (System::IO::Directory::Exists(encryptedDir) == false)
        {
            System::IO::Directory::CreateDirectory_(encryptedDir);
        }
        if (System::IO::Directory::Exists(pre97Dir) == false)
        {
            System::IO::Directory::CreateDirectory_(pre97Dir);
        }

        //ExStart:GetListOfFilesInFolder
        SharedPtr<System::Collections::Generic::IEnumerable<String>> fileList =
            System::IO::Directory::GetFiles(MyDir)->LINQ_Where([](String name) { return !name.EndsWith(u"Corrupted document.docx"); });
        //ExEnd:GetListOfFilesInFolder
        for (auto fileName : System::IterateOver(fileList))
        {
            String nameOnly = System::IO::Path::GetFileName(fileName);

            std::cout << nameOnly;
            //ExStart:DetectFileFormat
            SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(fileName);

            // Display the document type
            switch (info->get_LoadFormat())
            {
            case LoadFormat::Doc:
                std::cout << "\tMicrosoft Word 97-2003 document." << std::endl;
                break;

            case LoadFormat::Dot:
                std::cout << "\tMicrosoft Word 97-2003 template." << std::endl;
                break;

            case LoadFormat::Docx:
                std::cout << "\tOffice Open XML WordprocessingML Macro-Free Document." << std::endl;
                break;

            case LoadFormat::Docm:
                std::cout << "\tOffice Open XML WordprocessingML Macro-Enabled Document." << std::endl;
                break;

            case LoadFormat::Dotx:
                std::cout << "\tOffice Open XML WordprocessingML Macro-Free Template." << std::endl;
                break;

            case LoadFormat::Dotm:
                std::cout << "\tOffice Open XML WordprocessingML Macro-Enabled Template." << std::endl;
                break;

            case LoadFormat::FlatOpc:
                std::cout << "\tFlat OPC document." << std::endl;
                break;

            case LoadFormat::Rtf:
                std::cout << "\tRTF format." << std::endl;
                break;

            case LoadFormat::WordML:
                std::cout << "\tMicrosoft Word 2003 WordprocessingML format." << std::endl;
                break;

            case LoadFormat::Html:
                std::cout << "\tHTML format." << std::endl;
                break;

            case LoadFormat::Mhtml:
                std::cout << "\tMHTML (Web archive) format." << std::endl;
                break;

            case LoadFormat::Odt:
                std::cout << "\tOpenDocument Text." << std::endl;
                break;

            case LoadFormat::Ott:
                std::cout << "\tOpenDocument Text Template." << std::endl;
                break;

            case LoadFormat::DocPreWord60:
                std::cout << "\tMS Word 6 or Word 95 format." << std::endl;
                break;

            case LoadFormat::Unknown:
                std::cout << "\tUnknown format." << std::endl;
                break;

            default:
                break;
            }
            //ExEnd:DetectFileFormat

            if (info->get_IsEncrypted())
            {
                std::cout << "\tAn encrypted document." << std::endl;
                System::IO::File::Copy(fileName, System::IO::Path::Combine(encryptedDir, nameOnly), true);
            }
            else
            {
                switch (info->get_LoadFormat())
                {
                case LoadFormat::DocPreWord60:
                    System::IO::File::Copy(fileName, System::IO::Path::Combine(pre97Dir, nameOnly), true);
                    break;

                case LoadFormat::Unknown:
                    System::IO::File::Copy(fileName, System::IO::Path::Combine(unknownDir, nameOnly), true);
                    break;

                default:
                    System::IO::File::Copy(fileName, System::IO::Path::Combine(supportedDir, nameOnly), true);
                    break;
                }
            }
        }
        //ExEnd:CheckFormatCompatibility
    }

    void DetectDocumentSignatures()
    {
        //ExStart:DetectDocumentSignatures
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Digitally signed.docx");

        if (info->get_HasDigitalSignature())
        {
            std::cout << (String::Format(u"Document {0} has digital signatures, ", System::IO::Path::GetFileName(MyDir + u"Digitally signed.docx")) +
                          u"they will be lost if you open/save this document with Aspose.Words.")
                      << std::endl;
        }
        //ExEnd:DetectDocumentSignatures
    }

    void VerifyEncryptedDocument()
    {
        //ExStart:VerifyEncryptedDocument
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Encrypted.docx");
        std::cout << System::Convert::ToString(info->get_IsEncrypted()) << std::endl;
        //ExEnd:VerifyEncryptedDocument
    }
};

}} // namespace DocsExamples::File_Formats_and_Conversions
