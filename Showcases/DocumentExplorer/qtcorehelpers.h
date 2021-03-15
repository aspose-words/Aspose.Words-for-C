#ifndef ASPOSE_QTCOREHELPERS_H
#define ASPOSE_QTCOREHELPERS_H

#include <QByteArray>
#include <QDate>
#include <QDateTime>
#include <QString>
#include <QTime>
#include <QUuid>

#include <system/array.h>
#include <system/date_time.h>
#include <system/date_time_offset.h>
#include <system/exceptions.h>
#include <system/guid.h>
#include <system/string.h>
#include <system/timespan.h>

#include <cstdint>
#include <stdexcept>

namespace Aspose { namespace QtHelpers {

/// Convert System::String to QString.
inline QString Convert(const System::String& str)
{
    if (str.IsNull())
        return QString();

    const QChar* data = reinterpret_cast<const QChar*>(str.begin());
    const int length = str.get_Length();

    return QString(data, length);
}

/// Convert QString to System::String.
inline System::String Convert(const QString& str)
{
    if (str.isNull())
        return System::String();

    const char16_t* data = reinterpret_cast<const char16_t*>(str.data());
    const int length = str.length();

    return System::String(data, length);
}

/// Convert System::Array<uint8_t> to QByteArray.
inline QByteArray Convert(const System::ByteArrayPtr& array)
{
    if (array == nullptr)
        return QByteArray();

    const char* data = reinterpret_cast<const char*>(array->data_ptr());
    const int length = array->get_Length();

    return QByteArray(data, length);
}

/// Convert QByteArray to System::Array<uint8_t>.
inline System::ByteArrayPtr Convert(const QByteArray& array)
{
    if (array.isNull())
        return nullptr;

    const uint8_t* data = reinterpret_cast<const uint8_t*>(array.data());
    const int length = array.length();

    return System::MakeArray<uint8_t>(length, data);
}

/// Convert System::DateTime to QDateTime.
inline QDateTime Convert(const System::DateTime& dateTime)
{
    const auto elapsed = dateTime - System::DateTime::UnixEpoch;
    const auto elapsedMsecs = elapsed.get_Ticks() / System::DateTime::TicksPerMillisecond;

    auto result = QDateTime::fromMSecsSinceEpoch(elapsedMsecs, Qt::TimeSpec::UTC);
    if (dateTime.get_Kind() != System::DateTimeKind::Utc)
        result.setTimeSpec(Qt::TimeSpec::LocalTime);

    return result;
}

/// Convert System::DateTimeOffset to QDateTime.
inline QDateTime Convert(const System::DateTimeOffset& dateTimeOffset)
{
    const auto elapsed = dateTimeOffset.get_DateTime() - System::DateTime::UnixEpoch;
    const auto elapsedMsecs = elapsed.get_Ticks() / System::DateTime::TicksPerMillisecond;
    const auto offsetFromUtc = dateTimeOffset.get_Offset().get_Ticks() / System::TimeSpan::TicksPerSecond;

    return QDateTime::fromMSecsSinceEpoch(elapsedMsecs - offsetFromUtc * 1000, Qt::TimeSpec::OffsetFromUTC, offsetFromUtc);
}

/// Convert QDateTime to System::DateTime.
inline System::DateTime Convert(const QDateTime& dateTime)
{
    if (!dateTime.isValid())
        throw std::invalid_argument("invalid datetime");

    try
    {
        const System::TimeSpan timeSinceUnixEpoch(dateTime.toMSecsSinceEpoch() * System::TimeSpan::TicksPerMillisecond);

        System::TimeSpan offsetFromUtc;
        if (dateTime.timeSpec() == Qt::TimeSpec::LocalTime)
            offsetFromUtc = System::TimeSpan::FromTicks(dateTime.offsetFromUtc() * System::TimeSpan::TicksPerSecond);

        auto result = System::DateTime::UnixEpoch + (timeSinceUnixEpoch + offsetFromUtc);

        if (dateTime.timeSpec() == Qt::TimeSpec::LocalTime)
            result = System::DateTime::SpecifyKind(result, System::DateTimeKind::Local);
        else
            result = System::DateTime::SpecifyKind(result, System::DateTimeKind::Utc);

        return result;
    }
    catch (const System::ArgumentOutOfRangeException& e)
    {
        throw std::out_of_range(e->get_Message().ToUtf8String());
    }
}

/// Convert QDate to System::DateTime.
inline System::DateTime Convert(const QDate& date)
{
    if (!date.isValid())
        throw std::invalid_argument("invalid date");

    int year = 0;
    int month = 0;
    int day = 0;
    date.getDate(&year, &month, &day);

    if (year < 1 || year > 9999)
        throw std::out_of_range("date is out of range (1..9999 year)");

    return System::DateTime(year, month, day);
}

/// Convert QTime to System::TimeSpan.
inline System::TimeSpan Convert(const QTime& time)
{
    if (!time.isValid())
        throw std::invalid_argument("invalid time");

    return System::TimeSpan(0, time.hour(), time.minute(), time.second(), time.msec());
}

/// Convert System::Guid to QUuid.
inline QUuid Convert(const System::Guid& guid)
{
    return QUuid(Convert(guid.ToString()));
}

/// Convert QUuid to System::Guid.
inline System::Guid Convert(const QUuid& guid)
{
    const auto str = guid.toByteArray();
    return System::Guid(System::String::FromAscii(str.data(), str.length()));
}

}} // namespace Aspose::QtHelpers

#endif // ASPOSE_QTCOREHELPERS_H
