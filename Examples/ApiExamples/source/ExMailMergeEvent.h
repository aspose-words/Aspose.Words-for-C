#pragma once

#include <system/shared_ptr.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::MailMerging;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExMailMergeEvent : public ApiExampleBase
{
    typedef ExMailMergeEvent ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// If the mail merge encounters a MERGEFIELD whose name starts with the "html_" prefix,
    /// this callback parses its merge data as HTML content and adds the result to the document location of the MERGEFIELD.
    /// </summary>
    class HandleMergeFieldInsertHtml : public IFieldMergingCallback
    {
        typedef HandleMergeFieldInsertHtml ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    };
    
    /// <summary>
    /// Edits the values that MERGEFIELDs receive during a mail merge.
    /// The name of a MERGEFIELD must have a prefix for this callback to take effect on its value.
    /// </summary>
    class FieldValueMergingCallback : public IFieldMergingCallback
    {
        typedef FieldValueMergingCallback ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> e) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> e) override;
        
    };
    
    /// <summary>
    /// Upon encountering a MERGEFIELD with a specific name, inserts a check box form field instead of merge data text.
    /// </summary>
    class HandleMergeFieldInsertCheckBox : public IFieldMergingCallback
    {
        typedef HandleMergeFieldInsertCheckBox ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        HandleMergeFieldInsertCheckBox();
        
    private:
    
        int32_t mCheckBoxCount;
        
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    };
    
    /// <summary>
    /// Formats table rows as a mail merge takes place to alternate between two colors on odd/even rows.
    /// </summary>
    class HandleMergeFieldAlternatingRows : public IFieldMergingCallback
    {
        typedef HandleMergeFieldAlternatingRows ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        HandleMergeFieldAlternatingRows();
        
    private:
    
        System::SharedPtr<Aspose::Words::DocumentBuilder> mBuilder;
        int32_t mRowIdx;
        
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    };
    
    class HandleMergeImageFieldFromBlob : public IFieldMergingCallback
    {
        typedef HandleMergeImageFieldFromBlob ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        /// <summary>
        /// This is called when a mail merge encounters a MERGEFIELD in the document with an "Image:" tag in its name.
        /// </summary>
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> e) override;
        
    };
    
    
public:

    //ExStart
    //ExFor:DocumentBuilder.InsertHtml(String)
    //ExFor:MailMerge.FieldMergingCallback
    //ExFor:IFieldMergingCallback
    //ExFor:FieldMergingArgs
    //ExFor:FieldMergingArgsBase
    //ExFor:FieldMergingArgsBase.Field
    //ExFor:FieldMergingArgsBase.DocumentFieldName
    //ExFor:FieldMergingArgsBase.Document
    //ExFor:IFieldMergingCallback.FieldMerging
    //ExFor:FieldMergingArgs.Text
    //ExSummary:Shows how to execute a mail merge with a custom callback that handles merge data in the form of HTML documents.
    void MergeHtml();
    //ExEnd
    //ExStart
    //ExFor:FieldMergingArgsBase.FieldValue
    //ExSummary:Shows how to edit values that MERGEFIELDs receive as a mail merge takes place.
    void FieldFormats();
    //ExEnd
    void ImageFromUrl();
    
protected:

    /// <summary>
    /// Function needed for Visual Basic autoporting that returns the parity of the passed number.
    /// </summary>
    static bool IsOdd(int32_t value);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


