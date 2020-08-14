#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/list.h>
#include <system/array.h>
#include <drawing/image.h>
#include <cstdint>

#include "ApiExampleBase.h"

namespace ApiExamples { namespace TestData { namespace TestClasses { class SignPersonTestClass; } } }

namespace ApiExamples {

/// <summary>
/// This example demonstrates how to add new signature line to the document and sign it with your personal signature <see cref="SignDocument"></see>.
/// </summary>
class ExSignDocumentCustom : public ApiExampleBase
{
public:

    static void Sign();
    
protected:

    static System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<TestData::TestClasses::SignPersonTestClass>>> gSignPersonList;
    
    /// <summary>
    /// Signs the document obtained at the source location and saves it to the specified destination.
    /// </summary>
    static void SignDocument(System::String srcDocumentPath, System::String dstDocumentPath, System::SharedPtr<TestData::TestClasses::SignPersonTestClass> signPersonInfo, System::String certificatePath, System::String certificatePassword);
    /// <summary>
    /// Converting image file to bytes array.
    /// </summary>
    static System::ArrayPtr<uint8_t> ImageToByteArray(System::SharedPtr<System::Drawing::Image> imageIn);
    /// <summary>
    /// Create test data that contains info about sing persons
    /// </summary>
    static void CreateSignPersonData();
    
};

} // namespace ApiExamples


