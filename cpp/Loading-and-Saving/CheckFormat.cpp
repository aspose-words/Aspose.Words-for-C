#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>

using namespace Aspose::Words;

void CheckFormat()
{
    std::cout << "CheckFormat example started." << std::endl;
    // ExStart:CheckFormatCompatibility
    // The path to the documents directory.
    System::String dataDir = GetInputDataDir_LoadingAndSaving();
    
    System::String supportedDir = dataDir + u"OutSupported";
    System::String unknownDir = dataDir + u"OutUnknown";
    System::String encryptedDir = dataDir + u"OutEncrypted";
    System::String pre97Dir = dataDir + u"OutPre97";

    // Create the directories if they do not already exist
    if (!System::IO::Directory::Exists(supportedDir))
    {
        System::IO::Directory::CreateDirectory_(supportedDir);
    }
    if (!System::IO::Directory::Exists(unknownDir))
    {
        System::IO::Directory::CreateDirectory_(unknownDir);
    }
    if (!System::IO::Directory::Exists(encryptedDir))
    {
        System::IO::Directory::CreateDirectory_(encryptedDir);
    }
    if (!System::IO::Directory::Exists(pre97Dir))
    {
        System::IO::Directory::CreateDirectory_(pre97Dir);
    }

    // ExStart:GetListOfFilesInFolder
    System::ArrayPtr<System::String> fileList = System::IO::Directory::GetFiles(dataDir);
    // ExEnd:GetListOfFilesInFolder
    // Loop through all found files.

    for (System::String const &fileName: fileList)
    {
        // Extract and display the file name without the path.
        System::String nameOnly = System::IO::Path::GetFileName(fileName);
        std::cout << nameOnly.ToUtf8String();
        // ExStart:DetectFileFormat
        // Check the file format and move the file to the appropriate folder.
        System::SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(fileName);

        // Display the document type.
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
        default:
            std::cout << "\tUnknown format." << std::endl;
            break;

        }
        // ExEnd:DetectFileFormat

        // Now copy the document into the appropriate folder.
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

    // ExEnd:CheckFormatCompatibility
    std::cout << "Checked the format of all documents successfully." << std::endl;
    std::cout << "CheckFormat example finished." << std::endl << std::endl;
}