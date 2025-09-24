#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/io/stream_writer.h>
#include <system/array.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

/// <summary>
/// DOC2PDF document converter for SharePoint.
/// Uses Aspose.Words to perform the conversion.
/// </summary>
class ExMossDoc2Pdf
{
    typedef ExMossDoc2Pdf ThisType;
    
public:

    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    static void MossDoc2Pdf(System::ArrayPtr<System::String> args);
    
private:

    static System::String& gInFileName();
    static System::String& gOutFileName();
    static System::SharedPtr<System::IO::StreamWriter>& gLog();
    static void ParseCommandLine(System::ArrayPtr<System::String> args);
    static void ConvertDoc2Pdf(System::String inFileName, System::String outFileName);
    
public:
    ExMossDoc2Pdf() = delete;
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


