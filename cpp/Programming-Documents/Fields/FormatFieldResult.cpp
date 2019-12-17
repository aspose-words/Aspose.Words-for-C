#include "stdafx.h"
#include "examples.h"

#include <Model/Document/DateTimeFormatting/CalendarType.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/IFieldResultFormatter.h>
#include <Model/Fields/Format/GeneralFormat.h>
#include <Model/Fields/Field.h>
#include <Model/Fields/FieldOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace
{
    class FieldResultFormatter : public IFieldResultFormatter
    {
        typedef FieldResultFormatter ThisType;
        typedef IFieldResultFormatter BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        FieldResultFormatter(System::String numberFormat, System::String dateFormat);
        System::String FormatNumeric(double value, System::String format);
        System::String FormatDateTime(System::DateTime value, System::String format, CalendarType calendarType);
        System::String Format(System::String value, GeneralFormat format);
        System::String Format(double value, GeneralFormat format);

    private:
        System::String mNumberFormat;
        System::String mDateFormat;
    };

    RTTI_INFO_IMPL_HASH(2615486347u, FieldResultFormatter, ThisTypeBaseTypesInfo);

    FieldResultFormatter::FieldResultFormatter(System::String numberFormat, System::String dateFormat) : mNumberFormat(numberFormat), mDateFormat(dateFormat)
    {
    }

    System::String FieldResultFormatter::FormatNumeric(double value, System::String format)
    {
        return System::String::IsNullOrEmpty(mNumberFormat) ? nullptr : System::String::Format(mNumberFormat, value);
    }

    System::String FieldResultFormatter::FormatDateTime(System::DateTime value, System::String format, CalendarType calendarType)
    {
        return System::String::IsNullOrEmpty(mDateFormat) ? nullptr : System::String::Format(mDateFormat, value);
    }

    System::String FieldResultFormatter::Format(System::String value, GeneralFormat format)
    {
        throw System::NotImplementedException();
    }

    System::String FieldResultFormatter::Format(double value, GeneralFormat format)
    {
        throw System::NotImplementedException();
    }
}

void FormatFieldResult()
{
    std::cout << "FormatFieldResult example started." << std::endl;
    // ExStart:FormatFieldResult
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();

    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>();
    System::SharedPtr<Document> document = builder->get_Document();

    System::SharedPtr<Field> field = builder->InsertField(u"=-1234567.89 \\# \"### ### ###.000\"", nullptr);
    document->get_FieldOptions()->set_ResultFormatter(System::MakeObject<FieldResultFormatter>(u"[{0:N2}]", nullptr));

    field->Update();

    System::String outputPath = outputDataDir + u"FormatFieldResult.docx";
    builder->get_Document()->Save(outputPath);
    // ExEnd:FormatFieldResult
    std::cout << "Format field result successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "FormatFieldResult example finished." << std::endl << std::endl;
}